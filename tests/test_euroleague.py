"""Official EuroLeague adapter tests (Agent D scope).

All tests are offline: ``fetch`` is a fixture loader serving the committed
``tests/fixtures/euroleague/*.json`` files (captured LIVE from the official
endpoints on 2026-11-06 — LIVE_VERIFIED).  Helpers carry the unique ``el_``
prefix so they cannot collide with other agents' conftest helpers.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ingestion.models import FinalStatus, Game, GameStatus, League, PeriodKind
from ingestion.providers import IdentityError
from ingestion.providers import euroleague as el

FIXTURES = Path(__file__).parent / "fixtures" / "euroleague"

# URL-substring -> committed fixture file. Boxscore needles are checked
# before/instead of the schedule needle; the URL shapes do not overlap.
EL_URLS = {
    "seasons/E2026/games": "schedule.json",
    "gamecode=97": "boxscore_97.json",
    "gamecode=91": "boxscore_91_ot.json",
}


def el_load_fixture(name: str):
    """Load a committed EuroLeague fixture by file name."""
    with open(FIXTURES / name) as f:
        return json.load(f)


def el_fetch(urls=None):
    """Build a FetchFn that serves committed fixtures by URL substring."""
    mapping = urls if urls is not None else EL_URLS

    def _fetch(url: str):
        for needle, name in mapping.items():
            if needle in url:
                return el_load_fixture(name)
        raise AssertionError(f"no fixture mapped for URL: {url}")

    return _fetch


def el_make_game(
    gamecode: int = 91,
    home: str = "Real Madrid",
    away: str = "Panathinaikos AKTOR Athens",
    season: str = "E2026",
) -> Game:
    """Hand-build a canonical Game (seasoncode derived from season: fallback
    path for games that pre-date the euroleague_season provider id)."""
    return Game(
        game_id=f"E2026-{gamecode}",
        league=League.EUROLEAGUE,
        season=season,
        home_team_id=f"H{gamecode}",
        away_team_id=f"A{gamecode}",
        home_name=home,
        away_name=away,
        scheduled_utc="2026-11-04T19:45:00Z",
        provider_ids={"euroleague": str(gamecode)},
    )


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------


def test_schedule_discovers_games_with_utc_dates():
    games = el.discover_games(None, el_fetch())
    assert len(games) == 6
    by_id = {g.game_id: g for g in games}
    # Canonical game ids {seasoncode}-{gameCode}.
    assert set(by_id) == {
        "E2026-379",
        "E2026-380",
        "E2026-97",
        "E2026-96",
        "E2026-91",
        "E2026-66",
    }
    # Future game: FC Barcelona is the HOME side (local club), Partizan the
    # road side — orientation from local/road, never array order.
    future = by_id["E2026-379"]
    assert future.home_name == "FC Barcelona"
    assert future.away_name == "Partizan Mozzart Bet Belgrade"
    assert future.home_team_id == "BAR" and future.away_team_id == "PAR"
    # utcDate is Z-normalised timezone-aware UTC.
    assert future.scheduled_utc == "2027-04-16T18:30:00Z"
    assert future.scheduled_utc.endswith("Z")
    assert future.phase == "Round 38"  # official roundName as Game Type
    assert future.league is League.EUROLEAGUE
    assert future.season == "2026"
    assert future.provider_ids["euroleague"] == "379"
    assert future.provider_ids["euroleague_season"] == "E2026"
    assert future.status is GameStatus.SCHEDULED
    # Played games flip to FINAL.
    for gamecode in ("97", "96", "91", "66"):
        assert by_id[f"E2026-{gamecode}"].status is GameStatus.FINAL
    assert by_id["E2026-66"].phase == "Round 7"
    assert by_id["E2026-91"].home_name == "Real Madrid"
    assert by_id["E2026-91"].away_name == "Panathinaikos AKTOR Athens"
    assert by_id["E2026-91"].scheduled_utc == "2026-11-04T19:45:00Z"


def test_discover_date_filter_is_utc():
    games = el.discover_games("2026-11-06", el_fetch())
    assert [g.game_id for g in games] == ["E2026-97"]
    g = games[0]
    assert g.home_name == "Besiktas Istanbul"
    assert g.away_name == "Zalgiris Kaunas"
    assert g.scheduled_utc == "2026-11-06T17:00:00Z"
    # A UTC date with no games yields an empty list, not an error.
    assert el.discover_games("2026-01-01", el_fetch()) == []


def test_discover_rejects_same_club_on_both_sides():
    """local/road must be different clubs: identical codes raise
    IdentityError instead of guessing orientation."""
    payload = el_load_fixture("schedule.json")
    payload["data"][0]["local"]["club"]["code"] = "PAR"
    with pytest.raises(IdentityError):
        el.discover_games(None, lambda url: payload)


# ---------------------------------------------------------------------------
# Results — regular game 97 (BJK 67-83 ZAL)
# ---------------------------------------------------------------------------


def test_boxscore_byquarter_individual_not_cumulative():
    """CRITICAL: ByQuarter rows are individual quarters; EndOfQuarter rows
    are cumulative and must NEVER be written as individual quarters."""
    fixture = el_load_fixture("boxscore_97.json")
    # The source really does differ: EndOfQuarter Quarter2 is cumulative 31.
    assert fixture["ByQuarter"][0]["Quarter2"] == 14
    assert fixture["EndOfQuarter"][0]["Quarter2"] == 31

    game = el.discover_games("2026-11-06", el_fetch())[0]  # E2026-97
    result = el.fetch_game_result(game, el_fetch())
    assert result.game_id == "E2026-97"
    actual = result.actual
    assert actual is not None
    assert actual.period_kind is PeriodKind.QUARTERS
    # Individual quarters [17,14,13,23] — NOT the cumulative [17,31,44,67].
    assert actual.home_quarters == [17.0, 14.0, 13.0, 23.0]
    assert actual.away_quarters == [20.0, 25.0, 17.0, 21.0]
    assert 31.0 not in actual.home_quarters  # Q2 written 14, not 31
    assert 44.0 not in actual.home_quarters and 67.0 not in actual.home_quarters
    assert actual.home_overtime is None and actual.away_overtime is None
    assert actual.overtime_periods == 0
    # Final from the official cumulative EndOfQuarter last value (67/83) —
    # which the individual quarters sum to.
    assert actual.home_final == 67.0
    assert actual.away_final == 83.0
    assert actual.reconciles() == (True, "")
    assert result.status is GameStatus.FINAL
    assert result.final_status() is FinalStatus.FT


# ---------------------------------------------------------------------------
# Results — OT game 91 (RMB 112-115 PAO, Extra1 11/14)
# ---------------------------------------------------------------------------


def test_boxscore_ot_extra_periods_aggregate():
    games = el.discover_games(None, el_fetch())
    game = next(g for g in games if g.game_id == "E2026-91")
    result = el.fetch_game_result(game, el_fetch())
    actual = result.actual
    assert actual is not None
    # Regulation quarters only — OT lives in home_overtime/away_overtime.
    assert actual.home_quarters == [18.0, 30.0, 25.0, 28.0]
    assert actual.away_quarters == [24.0, 27.0, 23.0, 27.0]
    # Extra1 values aggregate into the single combined OT pair.
    assert actual.home_overtime == 11.0
    assert actual.away_overtime == 14.0
    assert actual.overtime_periods == 1
    # Final 112-115 (cumulative EndOfQuarter Extra1 = 101+11 / 101+14).
    assert actual.home_final == 112.0
    assert actual.away_final == 115.0
    assert actual.reconciles() == (True, "")
    assert result.status is GameStatus.FINAL
    assert result.final_status() is FinalStatus.AOT


def test_ot_schedule_detection():
    """The schedule partials already flag OT via extraPeriods; discovery
    surfaces it per game and via the helper — both OT games found."""
    games = el.discover_games(None, el_fetch())
    ot_games = [g for g in games if el.schedule_ot_periods(g) > 0]
    assert len(ot_games) >= 2
    assert {"E2026-91", "E2026-66"} <= {g.game_id for g in ot_games}
    assert el.schedule_ot_periods(next(g for g in ot_games if g.game_id == "E2026-91")) == 1
    # Future/regular games are not flagged.
    assert el.schedule_ot_periods(next(g for g in games if g.game_id == "E2026-379")) == 0
    assert el.schedule_ot_periods(next(g for g in games if g.game_id == "E2026-97")) == 0
    # The standalone helper agrees with discovery.
    assert set(el.ot_games_from_schedule(el_fetch())) == {"E2026-91", "E2026-66"}


# ---------------------------------------------------------------------------
# Identity matching
# ---------------------------------------------------------------------------


def test_boxscore_team_mismatch_raises_identity():
    """Boxscore teams matching neither schedule team quarantine; array order
    is never assumed, and nothing is half-written on quarantine."""
    payload = {
        "Live": False,
        "ByQuarter": [
            {"Team": "ALBA BERLIN", "Quarter1": 10, "Quarter2": 10, "Quarter3": 10, "Quarter4": 10},
            {"Team": "AS MONACO", "Quarter1": 12, "Quarter2": 12, "Quarter3": 12, "Quarter4": 12},
        ],
    }
    game = el_make_game(gamecode=97, home="Besiktas Istanbul", away="Zalgiris Kaunas")
    with pytest.raises(IdentityError):
        el.fetch_game_result(game, lambda url: payload)
    assert game.actual is None


def test_reversed_row_order_still_maps_by_team_name():
    """Orientation must survive row order: away row listed first."""
    payload = {
        "Live": False,
        "ByQuarter": [
            {
                "Team": "ZALGIRIS KAUNAS",
                "Quarter1": 20,
                "Quarter2": 25,
                "Quarter3": 17,
                "Quarter4": 21,
            },
            {
                "Team": "BESIKTAS ISTANBUL",
                "Quarter1": 17,
                "Quarter2": 14,
                "Quarter3": 13,
                "Quarter4": 23,
            },
        ],
    }
    game = el_make_game(gamecode=97, home="Besiktas Istanbul", away="Zalgiris Kaunas")
    result = el.fetch_game_result(game, lambda url: payload)
    assert result.actual.home_quarters == [17.0, 14.0, 13.0, 23.0]
    assert result.actual.away_quarters == [20.0, 25.0, 17.0, 21.0]
    assert result.actual.home_final == 67.0
    assert result.actual.away_final == 83.0


def test_boxscore_without_byquarter_is_rejected_not_improvised():
    """Without ByQuarter there is no safe individual-quarter source; the
    cumulative EndOfQuarter must never stand in for quarters."""
    payload = {
        "Live": False,
        "ByQuarter": [],
        "EndOfQuarter": [
            {
                "Team": "REAL MADRID",
                "Quarter1": 18,
                "Quarter2": 48,
                "Quarter3": 73,
                "Quarter4": 101,
            },
            {
                "Team": "PANATHINAIKOS AKTOR ATHENS",
                "Quarter1": 24,
                "Quarter2": 51,
                "Quarter3": 74,
                "Quarter4": 101,
            },
        ],
    }
    with pytest.raises(ValueError):
        el.fetch_game_result(el_make_game(), lambda url: payload)


def test_boxscore_seasoncode_from_provider_ids_or_season():
    """Boxscore URL seasoncode: explicit euroleague_season id first, else
    derived from the canonical season field."""
    assert el._seasoncode_of(el_make_game()) == "E2026"  # season "E2026"
    discovered = next(g for g in el.discover_games(None, el_fetch()) if g.game_id == "E2026-91")
    assert el._seasoncode_of(discovered) == "E2026"  # from provider_ids


# ---------------------------------------------------------------------------
# Standings (endpoint not confirmed — defensive, never raises)
# ---------------------------------------------------------------------------


def test_standings_fetch_failure_returns_empty_with_documentation_only_reason():
    report: dict = {}

    def _boom(url):
        raise RuntimeError("network down")

    rows = el.fetch_standings(_boom, report=report)
    assert rows == []
    assert report["status"] == "unavailable"
    assert "network down" in report["reason"]
    assert report["evidence"] == "DOCUMENTATION_ONLY"


def test_standings_malformed_payload_returns_empty_never_raises():
    report: dict = {}
    rows = el.fetch_standings(lambda url: {"unexpected": True}, report=report)
    assert rows == []
    assert report["status"] == "unavailable"


def test_standings_plausible_payload_parses_defensively():
    payload = {
        "data": [
            {
                "position": 1,
                "team": {"name": "Real Madrid", "code": "MAD"},
                "games": 10,
                "wins": 8,
                "losses": 2,
                "pointsFor": 850,
                "pointsAgainst": 800,
            }
        ]
    }
    report: dict = {}
    rows = el.fetch_standings(lambda url: payload, report=report)
    assert len(rows) == 1
    assert rows[0]["team"] == "Real Madrid"
    assert rows[0]["points_for"] == 850
    assert report["status"] == "ok"
    assert report["evidence"] == "DOCUMENTATION_ONLY"  # schema still unconfirmed


# ---------------------------------------------------------------------------
# Smoke
# ---------------------------------------------------------------------------


def test_smoke_on_fixtures_reports_required_shape():
    report = el.smoke(el_fetch())
    # Required shape.
    assert set(report) >= {
        "games_found",
        "results_reconciled",
        "ot_games",
        "verification",
    }
    assert report["games_found"] == 6
    # Only two boxscore fixtures exist, so coverage is an honest 2 of 4; the
    # missing two are recorded in 'failures' instead of raising.
    assert report["results_reconciled"] == {"complete": 2, "total": 4}
    assert {f["game_id"] for f in report["failures"]} == {"E2026-96", "E2026-66"}
    # Both OT games visible in the schedule partials (91 and 66).
    assert report["ot_games"] == ["E2026-91", "E2026-66"]
    # Replaying captures is fixture verification, not a new live call.
    assert report["verification"] == "FIXTURE_VERIFIED"
