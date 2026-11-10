"""ESPN provider tests — fixture-based, no network.

Golden: event 401909938 (Denver Nuggets at Indiana Pacers, BPI observed
2026-11-06T21:51Z with statistics on ``awayTeam``) must map to exact canonical
prediction facts. Orientation tests prove the mapping flips through
team ids when the statistics sit on the other side, and that an
unreconcilable stats-side team id (or identical predictor refs) is
rejected, never guessed. Result tests prove quarter/OT aggregation and
reconciliation; standings tests prove points for/against stay cumulative.
"""

from __future__ import annotations

from typing import Any, Dict

import pytest

from forecasting.types import _period_scores
from ingestion.http import HttpClient
from ingestion.models import (
    FinalStatus,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
)
from ingestion.providers import IdentityError
from ingestion.providers.espn import (
    discover_games,
    discover_games_with_evidence,
    fetch_bpi,
    fetch_game_context,
    fetch_game_result,
    fetch_recent_results,
    fetch_standings,
    smoke,
)

GOLDEN_EVENT = "401909938"
PREDICTOR_KEY = f"competitions/{GOLDEN_EVENT}/predictor"


def _golden_game() -> Game:
    return Game(
        game_id=GOLDEN_EVENT,
        league=League.NBA,
        season="2027",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T00:00:00Z",
        provider_ids={"espn": GOLDEN_EVENT},
    )


def _ncaa_dual_game() -> Game:
    return Game(
        game_id="401909760",
        league=League.NCAA_M,
        season="2027",
        home_team_id="38",
        away_team_id="21",
        home_name="Colorado Buffaloes",
        away_name="San Diego State Aztecs",
        scheduled_utc="2026-11-07T20:30:00Z",
        provider_ids={"espn": "401909760"},
    )


def _summary_game() -> Game:
    return Game(
        game_id="401909936",
        league=League.NBA,
        season="2027",
        home_team_id="21",
        away_team_id="19",
        home_name="Phoenix Suns",
        away_name="Orlando Magic",
        scheduled_utc="2026-11-06T02:00:00Z",
        provider_ids={"espn": "401909936"},
    )


def _ncaaw_game() -> Game:
    return Game(
        game_id="401918841",
        league=League.NCAA_W,
        season="2027",
        home_team_id="251",
        away_team_id="2458",
        home_name="Texas Longhorns",
        away_name="Northern Colorado Bears",
        scheduled_utc="2026-11-07T01:00:00Z",
        provider_ids={"espn": "401918841"},
    )


def _mutated_golden(load_fixture, stat_name: str, value: float) -> Dict[str, Any]:
    """Golden predictor payload with one statistic value tampered."""
    payload: Dict[str, Any] = load_fixture("espn/predictor_401909938.json")
    for stat in payload["awayTeam"]["statistics"]:
        if stat["name"] == stat_name:
            stat["value"] = value
            return payload
    raise AssertionError(f"stat {stat_name!r} not found in golden fixture")  # pragma: no cover


# ---------------------------------------------------------------------------
# BPI golden + orientation
# ---------------------------------------------------------------------------


def test_golden_bpi_401909938_maps_exact_prediction_facts(load_fixture, fixture_fetch):
    # No game passed: orientation falls back to the payload's homeTeam/
    # awayTeam keys (refs 11 vs 7 are present and distinct). Stats sit on
    # awayTeam (Denver), so 76.938 is the AWAY win probability.
    prediction = fetch_bpi(
        GOLDEN_EVENT, fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938.json"})
    )
    assert prediction.provider == "espn"
    assert prediction.model == "espn-bpi"
    assert prediction.game_id == GOLDEN_EVENT
    assert prediction.prob_home == pytest.approx(0.23062)
    assert prediction.prob_away == pytest.approx(0.76938)
    assert prediction.expected_home_score == pytest.approx(110.540)
    assert prediction.expected_away_score == pytest.approx(120.046)
    assert prediction.expected_total == pytest.approx(230.586)
    assert prediction.expected_margin == pytest.approx(-9.506)
    assert prediction.source_modified_utc == "2026-11-06T21:51:00Z"
    assert prediction.source_revision == "2026-11-06T21:51Z"
    assert prediction.pre_tip is True


def test_golden_bpi_with_game_object(load_fixture, fixture_fetch):
    # No event fixture in the mapping: passing the discovered game must be
    # enough — the event id comes from game.provider_ids["espn"] and the
    # stats side (away, id 7) is reconciled against the game's team ids.
    prediction = fetch_bpi(
        GOLDEN_EVENT,
        fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938.json"}),
        game=_golden_game(),
    )
    assert prediction.prob_home == pytest.approx(0.23062)
    assert prediction.prob_away == pytest.approx(0.76938)
    assert prediction.expected_total == pytest.approx(230.586)


def test_bpi_orientation_when_stats_on_home_side(load_fixture, fixture_fetch):
    # Synthetic copy of the golden predictor with the statistics moved to
    # homeTeam (team id 11) and recomputed values: orientation must come
    # from the team refs matched against the game's competitor ids, never
    # from payload side order or array position.
    game = _golden_game()
    prediction = fetch_bpi(
        GOLDEN_EVENT,
        fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938_home_stats.json"}),
        game=game,
    )
    assert prediction.prob_home == pytest.approx(0.55)
    assert prediction.prob_away == pytest.approx(0.45)
    assert prediction.expected_home_score == pytest.approx(118.0)
    assert prediction.expected_away_score == pytest.approx(112.0)
    assert prediction.expected_total == pytest.approx(230.0)
    assert prediction.expected_margin == pytest.approx(6.0)


def test_bpi_unknown_team_reference_raises_identity_error(load_fixture, fixture_fetch):
    # Stats-side team id 99 matches neither competitor (11 home, 7 away):
    # rejected, never guessed.
    with pytest.raises(IdentityError):
        fetch_bpi(
            GOLDEN_EVENT,
            fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938_mismatch.json"}),
            game=_golden_game(),
        )


def test_bpi_identical_team_refs_raises_identity_error(load_fixture, fixture_fetch):
    # Both predictor sides reference team 7: without competitor ids the
    # sides cannot be reconciled, so the fallback path must quarantine.
    fetch = fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938_identical_refs.json"})
    with pytest.raises(IdentityError):
        fetch_bpi(GOLDEN_EVENT, fetch)
    with pytest.raises(IdentityError):
        fetch_bpi(GOLDEN_EVENT, fetch, game=_golden_game())


def test_bpi_missing_expected_scores_keeps_probabilities_and_gaps(load_fixture, fixture_fetch):
    # WNBA/NCAA pattern: probabilities only, no teamExpectedPts/
    # oppExpectedPts — probabilities stay usable, scores stay blank, and
    # smoke must report the expected_scores gap instead of imputing.
    partial_event = "409900001"
    partial_game = Game(
        game_id=partial_event,
        league=League.WNBA,
        season="2026",
        home_team_id="11",
        away_team_id="7",
        home_name="Las Vegas Aces",
        away_name="New York Liberty",
        scheduled_utc="2026-06-10T02:00:00Z",
        provider_ids={"espn": partial_event},
    )
    prediction = fetch_bpi(
        partial_event,
        fixture_fetch(
            {f"competitions/{partial_event}/predictor": "espn/predictor_409900001_partial.json"}
        ),
        game=partial_game,
    )
    assert prediction.prob_away == pytest.approx(0.554)
    assert prediction.prob_home == pytest.approx(0.446)
    assert prediction.expected_home_score is None
    assert prediction.expected_away_score is None
    assert prediction.expected_total is None
    assert prediction.expected_margin is None
    assert prediction.prob_away > prediction.prob_home

    scoreboard = {
        "events": [
            {
                "id": partial_event,
                "date": "2026-06-10T02:00Z",
                "name": "New York Liberty at Las Vegas Aces",
                "season": {"year": 2026},
                "competitions": [
                    {
                        "date": "2026-06-10T02:00Z",
                        "status": {"type": {"state": "pre"}},
                        "competitors": [
                            {
                                "homeAway": "home",
                                "team": {"id": "11", "displayName": "Las Vegas Aces"},
                            },
                            {
                                "homeAway": "away",
                                "team": {"id": "7", "displayName": "New York Liberty"},
                            },
                        ],
                    }
                ],
            }
        ]
    }
    report = smoke(
        "20260610",
        fixture_fetch(
            {
                "dates=20260610": scoreboard,
                f"competitions/{partial_event}/predictor": "espn/predictor_409900001_partial.json",
            }
        ),
    )
    assert report["fixtures_found"] == 1
    assert report["bpi_complete"] == {"complete": 0, "total": 1}
    gap = next(g for g in report["gaps"] if g["required_field"] == "expected_scores")
    assert gap["game_id"] == partial_event
    assert [g["required_field"] for g in report["gaps"]] == ["expected_scores", "total"]


def test_dual_sided_ncaa_bpi_maps_both_team_probabilities(load_fixture, fixture_fetch):
    event_id = "401909760"
    payload = load_fixture("espn/predictor_ncaa_dual_401909760.json")
    prediction = fetch_bpi(
        event_id,
        fixture_fetch({f"competitions/{event_id}/predictor": payload}),
        game=_ncaa_dual_game(),
        now_utc="2026-11-07T20:29:04Z",
    )
    assert prediction.prob_home == pytest.approx(0.5847)
    assert prediction.prob_away == pytest.approx(0.4153)
    assert prediction.source_modified_utc == "2026-11-06T09:09:00Z"
    assert prediction.observed_utc == "2026-11-07T20:29:04Z"
    assert prediction.expected_home_score is None
    assert prediction.expected_away_score is None
    assert prediction.expected_total is None
    assert prediction.expected_margin is None


def test_dual_sided_bpi_cross_validates_expected_scores(load_fixture, fixture_fetch):
    event_id = "401909760"
    payload = load_fixture("espn/predictor_ncaa_dual_401909760.json")
    for side, values in (
        (
            "homeTeam",
            {"teamExpectedPts": 80, "oppExpectedPts": 70, "teamPredPtDiff": 10},
        ),
        (
            "awayTeam",
            {"teamExpectedPts": 70, "oppExpectedPts": 80, "teamPredPtDiff": -10},
        ),
    ):
        payload[side]["statistics"].extend(
            {"name": name, "value": value} for name, value in values.items()
        )
    prediction = fetch_bpi(
        event_id,
        fixture_fetch({f"competitions/{event_id}/predictor": payload}),
        game=_ncaa_dual_game(),
    )
    assert prediction.expected_home_score == 80
    assert prediction.expected_away_score == 70
    assert prediction.expected_total == 150
    assert prediction.expected_margin == 10


def test_dual_sided_bpi_quarantines_identity_and_probability_contradictions(
    load_fixture, fixture_fetch
):
    event_id = "401909760"
    wrong_team = load_fixture("espn/predictor_ncaa_dual_401909760.json")
    wrong_team["homeTeam"]["team"]["$ref"] = wrong_team["homeTeam"]["team"]["$ref"].replace(
        "/teams/38", "/teams/21"
    )
    wrong_team["awayTeam"]["team"]["$ref"] = wrong_team["awayTeam"]["team"]["$ref"].replace(
        "/teams/21", "/teams/38"
    )
    with pytest.raises(IdentityError, match="contradict scoreboard"):
        fetch_bpi(
            event_id,
            fixture_fetch({f"competitions/{event_id}/predictor": wrong_team}),
            game=_ncaa_dual_game(),
        )

    wrong_probability = load_fixture("espn/predictor_ncaa_dual_401909760.json")
    for stat in wrong_probability["awayTeam"]["statistics"]:
        if stat["name"] == "gameProjection":
            stat["value"] = 40.0
    with pytest.raises(ValueError):
        fetch_bpi(
            event_id,
            fixture_fetch({f"competitions/{event_id}/predictor": wrong_probability}),
            game=_ncaa_dual_game(),
        )


def test_bpi_probability_sum_violation_raises(load_fixture, fixture_fetch):
    payload = _mutated_golden(load_fixture, "gameProjection", 65.938)  # 65.938 + 23.062 != 100
    with pytest.raises(ValueError):
        fetch_bpi(
            GOLDEN_EVENT,
            fixture_fetch({PREDICTOR_KEY: payload}),
            game=_golden_game(),
        )


def test_score_margin_contradiction_raises(load_fixture, fixture_fetch):
    payload = _mutated_golden(load_fixture, "teamPredPtDiff", 19.506)  # scores say 9.506
    with pytest.raises(ValueError):
        fetch_bpi(
            GOLDEN_EVENT,
            fixture_fetch({PREDICTOR_KEY: payload}),
            game=_golden_game(),
        )


# ---------------------------------------------------------------------------
# Idempotent mapping
# ---------------------------------------------------------------------------


def test_fetch_bpi_idempotent(load_fixture, fixture_fetch):
    fetch = fixture_fetch({PREDICTOR_KEY: "espn/predictor_401909938.json"})
    first = fetch_bpi(GOLDEN_EVENT, fetch)
    second = fetch_bpi(GOLDEN_EVENT, fetch)

    def payload_fields(prediction):
        return {k: v for k, v in prediction.__dict__.items() if k != "observed_utc"}

    assert payload_fields(first) == payload_fields(second)


def test_discover_games_idempotent(load_fixture, fixture_fetch):
    fetch = fixture_fetch({"dates=20261106": "espn/scoreboard_20261106.json"})
    assert discover_games("20261106", fetch) == discover_games("20261106", fetch)


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------


def test_discover_games_from_scoreboard(load_fixture, fixture_fetch):
    games = discover_games(
        "20261106", fixture_fetch({"dates=20261106": "espn/scoreboard_20261106.json"})
    )
    assert len(games) == 9
    hawks = next(g for g in games if g.game_id == "401909287")
    # WSH @ ATL: home/away orientation from the homeAway markers.
    assert hawks.home_name == "Atlanta Hawks"
    assert hawks.away_name == "Washington Wizards"
    assert hawks.home_team_id == "1"
    assert hawks.away_team_id == "27"
    assert hawks.provider_ids == {"espn": "401909287"}
    # dates=20261106 is the US Eastern slate day; event datetime is UTC.
    assert hawks.scheduled_utc == "2026-11-07T00:00:00Z"
    assert hawks.status is GameStatus.SCHEDULED
    assert hawks.league is League.NBA
    assert hawks.season == "2027"
    assert hawks.phase == "League"
    assert all(g.scheduled_utc.endswith("Z") for g in games)
    assert all(g.provider_ids.get("espn") == g.game_id for g in games)
    names = {g.home_name for g in games} | {g.away_name for g in games}
    assert "Oklahoma City Thunder" in names  # canonicalised via identity aliases
    assert "Philadelphia 76ers" in names


def test_rich_scoreboard_retains_context_and_named_full_game_books(fixture_fetch):
    games, books = discover_games_with_evidence(
        "20261107",
        fixture_fetch({"dates=20261107": "espn/scoreboard_rich_401909938.json"}),
        observed_utc="2026-11-07T18:00:00Z",
    )
    assert len(games) == 1
    game = games[0]
    snapshots = {snapshot.market: snapshot for snapshot in books[game.game_id]}
    assert set(snapshots) == {Market.MONEYLINE, Market.SPREAD, Market.TOTAL}
    moneyline = snapshots[Market.MONEYLINE]
    assert moneyline.provider == "espn"
    assert moneyline.book == "DraftKings"
    assert moneyline.home_price_decimal == pytest.approx(3.61)
    assert moneyline.away_price_decimal == pytest.approx(1 + 100 / 324)
    assert moneyline.pre_match is True
    assert moneyline.observed_utc == "2026-11-07T18:00:00Z"
    assert snapshots[Market.SPREAD].line == pytest.approx(7.5)
    assert snapshots[Market.SPREAD].home_price_decimal == pytest.approx(1 + 100 / 105)
    assert snapshots[Market.TOTAL].line == pytest.approx(229.5)
    assert snapshots[Market.TOTAL].over_price_decimal == pytest.approx(1 + 100 / 107)
    # Open values are context, not three extra falsely timestamped snapshots.
    assert len(books[game.game_id]) == 3

    scoreboard = game.evidence["espn"]["scoreboard"]
    assert scoreboard["neutral_site"] is True
    assert scoreboard["venue"]["fullName"] == "Mexico City Arena"
    assert scoreboard["teams"]["home"]["records"][0]["summary"] == "3-6"
    assert scoreboard["teams"]["away"]["statistics"][0]["displayValue"] == "115.4"
    assert scoreboard["odds"][0]["moneyline"]["home"]["open"]["odds"] == "+199"
    import json

    serialized = json.dumps(game.evidence)
    assert "must-not-persist" not in serialized
    assert "do-not-store" not in serialized
    assert "affiliate" not in serialized


def test_summary_context_retains_form_injuries_and_article_without_links(fixture_fetch):
    fetch = fixture_fetch(
        {
            "dates=20261107": "espn/scoreboard_rich_401909938.json",
            "summary?event=401909938": "espn/summary_401909938_context.json",
        }
    )
    games, _books = discover_games_with_evidence(
        "20261107", fetch, observed_utc="2026-11-07T18:00:00Z"
    )
    game = fetch_game_context(GOLDEN_EVENT, fetch, games[0], observed_utc="2026-11-07T18:01:00Z")
    summary = game.evidence["espn"]["summary"]
    assert summary["observed_utc"] == "2026-11-07T18:01:00Z"
    assert summary["source_modified_utc"] == "2026-11-07T06:02:49Z"
    assert summary["team_boxscore"][0]["statistics"][0]["displayValue"] == "112.3"
    assert summary["last_five_games"][0]["events"][0]["gameResult"] == "W"
    assert summary["injuries"][0]["injuries"][0]["status"] == "Out"
    assert "INJURIES:" in summary["article"]["story"]
    assert "news" not in summary
    import json

    serialized = json.dumps(summary)
    assert "do-not-store" not in serialized
    assert "large taxonomy omitted" not in serialized
    assert "Unrelated league-wide" not in serialized


def test_summary_context_retains_player_boxscore_and_deep_source_coverage(
    load_fixture, fixture_fetch
):
    payload = load_fixture("espn/summary_401909938_context.json")
    payload["boxscore"]["players"] = [
        {
            "team": {"id": "7", "displayName": "Denver Nuggets"},
            "statistics": [
                {
                    "name": "points",
                    "labels": ["MIN", "PTS"],
                    "athletes": [
                        {
                            "athlete": {"id": "3112335", "displayName": "Nikola Jokic"},
                            "stats": ["38:12", "28"],
                            "didNotPlay": False,
                        }
                    ],
                }
            ],
        }
    ]
    payload["boxscore"]["leaders"] = [
        {
            "team": {"id": "7"},
            "leaders": [{"name": "points", "displayValue": "28 N. Jokic"}],
        }
    ]
    payload["plays"] = [{"id": "play-1", "text": "Example scoring play"}]
    payload["winprobability"] = [{"playId": "play-1", "homeWinPercentage": 0.6}]
    payload["shotChart"] = [{"id": "shot-1", "made": True}]
    payload["standings"] = {"groups": [{"id": "6"}]}
    fetch = fixture_fetch({"summary?event=401909938": payload})
    game = fetch_game_context(
        GOLDEN_EVENT,
        fetch,
        _golden_game(),
        observed_utc="2026-11-07T18:01:00Z",
    )
    summary = game.evidence["espn"]["summary"]
    assert summary["player_boxscore"][0]["statistics"][0]["athletes"][0]["stats"] == ["38:12", "28"]
    assert summary["leaders"][0]["leaders"][0]["displayValue"] == "28 N. Jokic"
    assert summary["context_coverage"] == {
        "player_boxscore_teams": 1,
        "player_boxscore_athletes": 1,
        "play_by_play_events_in_raw_payload": 1,
        "win_probability_points_in_raw_payload": 1,
        "shot_chart_events_in_raw_payload": 1,
        "standings_groups_in_raw_payload": 1,
    }
    # Large in-play arrays are represented by coverage, not mislabeled as
    # pre-tip features or silently truncated into canonical evidence.
    assert "plays" not in summary
    assert "winprobability" not in summary
    assert "shotChart" not in summary


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------


def test_finals_reconcile_quarters(load_fixture, fixture_fetch):
    # Scoreboard finals: quarter linescores sum to the final for both
    # teams (117 = 33+30+33+21, 104 = 26+25+22+31), no OT, FT status.
    recent = fetch_recent_results(
        "20261105", fixture_fetch({"dates=20261105": "espn/scoreboard_20261105.json"})
    )
    assert len(recent) == 2
    assert all(g.status is GameStatus.FINAL for g in recent)
    suns = next(g for g in recent if g.game_id == "401909936")
    assert suns.home_name == "Phoenix Suns"
    assert suns.away_name == "Orlando Magic"
    assert suns.actual is not None
    assert suns.actual.period_kind is PeriodKind.QUARTERS
    assert suns.actual.source == "espn"
    assert suns.actual.home_quarters == [33.0, 30.0, 33.0, 21.0]
    assert suns.actual.away_quarters == [26.0, 25.0, 22.0, 31.0]
    assert suns.actual.home_final == 117.0
    assert suns.actual.away_final == 104.0
    assert suns.actual.home_overtime is None
    assert suns.actual.away_overtime is None
    assert suns.actual.overtime_periods == 0
    assert suns.final_status() is FinalStatus.FT
    assert suns.actual.reconciles() == (True, "")
    heat = next(g for g in recent if g.game_id == "401909937")
    assert (heat.actual.home_final, heat.actual.away_final) == (120.0, 100.0)
    assert all(g.actual is not None and g.actual.reconciles()[0] for g in recent)


def test_ot_linescores_aggregate_into_single_ot(load_fixture, fixture_fetch):
    # Synthetic scoreboard with a 5-period linescore: the 5th element is
    # OT, summed into the single combined canonical OT pair, never
    # treated as a 5th quarter.
    games = fetch_recent_results(
        "20261104", fixture_fetch({"dates=20261104": "espn/scoreboard_20261104_ot.json"})
    )
    assert len(games) == 1
    game = games[0]
    assert game.status is GameStatus.FINAL
    assert game.actual is not None
    assert game.actual.home_quarters == [28.0, 26.0, 31.0, 25.0]
    assert game.actual.away_quarters == [29.0, 28.0, 26.0, 28.0]
    assert game.actual.home_overtime == 12.0
    assert game.actual.away_overtime == 7.0
    assert game.actual.overtime_periods == 1
    assert (game.actual.home_final, game.actual.away_final) == (122.0, 118.0)
    assert game.final_status() is FinalStatus.AOT
    assert game.actual.reconciles()[0] is True


def test_game_result_from_summary(load_fixture, fixture_fetch):
    game = fetch_game_result(
        "401909936", fixture_fetch({"summary?event=401909936": "espn/summary_401909936.json"})
    )
    assert game.status is GameStatus.FINAL
    assert game.final_status() is FinalStatus.FT
    assert game.home_name == "Phoenix Suns"
    assert game.away_name == "Orlando Magic"
    assert game.actual is not None
    assert game.actual.period_kind is PeriodKind.QUARTERS
    assert game.actual.source == "espn"
    assert game.actual.home_quarters == [33.0, 30.0, 33.0, 21.0]
    assert game.actual.away_quarters == [26.0, 25.0, 22.0, 31.0]
    assert (game.actual.home_final, game.actual.away_final) == (117.0, 104.0)
    assert game.actual.reconciles() == (True, "")


def test_result_fills_existing_game_without_mutation(load_fixture, fixture_fetch):
    mapping = {"summary?event=401909936": "espn/summary_401909936.json"}
    game = _summary_game()
    result = fetch_game_result("401909936", fixture_fetch(mapping), game=game)
    assert game.actual is None  # input game untouched
    assert result.game_id == "401909936"
    assert result.provider_ids == {"espn": "401909936"}
    assert result.home_team_id == "21"
    assert result.away_team_id == "19"
    assert result.status is GameStatus.FINAL
    assert result.final_status() is FinalStatus.FT


def test_result_ot_aggregation(load_fixture, fixture_fetch):
    mapping = {"summary?event=401909939": "espn/summary_401909939_ot.json"}
    game = fetch_game_result("401909939", fixture_fetch(mapping))
    assert game.actual is not None
    assert game.actual.home_quarters == [33.0, 30.0, 30.0, 27.0]
    assert game.actual.away_quarters == [26.0, 25.0, 32.0, 34.0]
    assert game.actual.home_overtime == 3.0
    assert game.actual.away_overtime == 3.0
    assert game.actual.overtime_periods == 1
    assert game.actual.home_final == 123.0
    assert game.actual.away_final == 120.0
    assert game.status is GameStatus.FINAL
    assert game.final_status() is FinalStatus.AOT
    assert game.actual.reconciles()[0] is True


def test_result_double_ot_aggregation(load_fixture, fixture_fetch):
    mapping = {"summary?event=401909940": "espn/summary_401909940_2ot.json"}
    game = fetch_game_result("401909940", fixture_fetch(mapping))
    assert game.actual is not None
    assert game.actual.home_overtime == 5.0  # OT1 3 + OT2 2
    assert game.actual.away_overtime == 4.0  # OT1 2 + OT2 2
    assert game.actual.overtime_periods == 2
    assert game.actual.home_final == 123.0
    assert game.actual.away_final == 122.0
    assert game.final_status() is FinalStatus.AOT
    assert game.actual.reconciles()[0] is True


def test_ncaaw_four_quarters_are_not_two_halves_plus_fake_overtime(load_fixture, fixture_fetch):
    """Live-shaped Texas/UNCO final: 23+26+23+19 = 91, 6+3+9+11 = 29."""
    game = fetch_game_result(
        "401918841",
        fixture_fetch({"summary": load_fixture("espn/summary_ncaaw_401918841_quarters.json")}),
        game=_ncaaw_game(),
    )
    actual = game.actual
    assert actual is not None
    assert actual.period_kind is PeriodKind.QUARTERS
    assert actual.home_quarters == [23.0, 26.0, 23.0, 19.0]
    assert actual.away_quarters == [6.0, 3.0, 9.0, 11.0]
    assert actual.home_overtime is None
    assert actual.away_overtime is None
    assert actual.overtime_periods == 0
    assert actual.reconciles() == (True, "")
    assert _period_scores(game, Period.H1) == (49.0, 9.0)
    assert _period_scores(game, Period.H2) == (42.0, 20.0)
    assert _period_scores(game, Period.Q3) == (23.0, 9.0)
    assert _period_scores(game, Period.FULL_GAME) == (91.0, 29.0)


def test_ncaaw_explicit_legacy_two_period_payload_remains_halves(load_fixture, fixture_fetch):
    payload = load_fixture("espn/summary_ncaaw_401918841_quarters.json")
    payload["format"] = {"regulation": {"periods": 2}}
    home, away = payload["header"]["competitions"][0]["competitors"]
    home["score"] = "75"
    away["score"] = "55"
    home["linescores"] = [{"value": 40.0}, {"value": 35.0}]
    away["linescores"] = [{"value": 30.0}, {"value": 25.0}]
    game = fetch_game_result("401918841", fixture_fetch({"summary": payload}), game=_ncaaw_game())
    actual = game.actual
    assert actual is not None
    assert actual.period_kind is PeriodKind.HALVES
    assert actual.home_quarters == [40.0, 35.0]
    assert actual.away_quarters == [30.0, 25.0]
    assert actual.overtime_periods == 0
    assert actual.reconciles() == (True, "")
    assert _period_scores(game, Period.H1) == (40.0, 30.0)
    assert _period_scores(game, Period.Q1) is None


def test_result_live_is_quarantined(load_fixture, fixture_fetch):
    mapping = {"summary?event=401909941": "espn/summary_401909941_live.json"}
    with pytest.raises(IdentityError):
        fetch_game_result("401909941", fixture_fetch(mapping))


def test_result_missing_linescores_is_quarantined(load_fixture, fixture_fetch):
    payload = load_fixture("espn/summary_401909936.json")
    for competitor in payload["header"]["competitions"][0]["competitors"]:
        competitor["linescores"] = []
    mapping = {"summary?event=401909936": payload}
    with pytest.raises(IdentityError):
        fetch_game_result("401909936", fixture_fetch(mapping))


# ---------------------------------------------------------------------------
# Standings
# ---------------------------------------------------------------------------


def test_standings_cumulative_pf_pa(load_fixture, fixture_fetch):
    rows = fetch_standings(fixture_fetch({"basketball/nba/standings": "espn/standings.json"}))
    assert len(rows) == 30
    bulls = next(r for r in rows if r.team_name == "Chicago Bulls")
    assert bulls.points_for == 893.0  # cumulative, stored as-is
    assert bulls.points_against == 1037.0  # never divided by games played
    # Cumulative guard: 893 is far above the 3 x 8 plausibility floor.
    assert bulls.points_for > 3 * bulls.games_played
    assert bulls.win_pct == pytest.approx(0.125)
    assert bulls.streak == -6  # ESPN "L6" losing streak encoded negative
    assert bulls.games_played == 8  # wins + losses (payload omits gamesPlayed)
    assert bulls.wins == 1
    assert bulls.losses == 7
    assert bulls.position == 14  # mirrors playoffSeed — NOT a guaranteed rank
    assert bulls.section == "Eastern Conference"  # child name kept verbatim
    assert bulls.source == "espn"
    sections = {r.section for r in rows}
    assert sections == {"Eastern Conference", "Western Conference"}


def test_standings_rejects_non_cumulative_points_for(load_fixture, fixture_fetch):
    # Guard: pointsFor <= 3 x games_played cannot be a cumulative total —
    # someone wrote an average into the cumulative field. Quarantine.
    payload = load_fixture("espn/standings.json")
    for child in payload["children"]:
        for entry in child["standings"]["entries"]:
            if entry["team"]["id"] == "4":
                for stat in entry["stats"]:
                    if stat["name"] == "pointsFor":
                        stat["value"] = 20.0
    with pytest.raises(ValueError):
        fetch_standings(fixture_fetch({"basketball/nba/standings": payload}))


# ---------------------------------------------------------------------------
# Smoke coverage
# ---------------------------------------------------------------------------


def test_smoke_reports_actual_coverage(load_fixture, fixture_fetch):
    mapping = {
        "dates=20261106": "espn/scoreboard_20261106.json",
        "competitions/401909287/predictor": "espn/predictor_401909287.json",
    }
    report = smoke("20261106", fixture_fetch(mapping))
    assert report["fixtures_found"] == 9
    assert report["bpi_complete"] == {"complete": 1, "total": 9}
    # Only 401909287 has a predictor fixture; the other 8 events each miss
    # prob/expected_scores/total.
    assert len(report["gaps"]) == 24
    assert all(gap["detail"].startswith("KeyError") for gap in report["gaps"])
    assert report["results_reconciled"] == {
        "finals": 0,
        "reconciled": 0,
        "flagged": 0,
    }
    assert report["verification"] == "FIXTURE_VERIFIED"


def test_smoke_flags_reconciliation_gap(load_fixture, fixture_fetch):
    # A final whose quarter sums contradict the final score keeps its
    # record but is flagged — no residual is folded into OT.
    payload = load_fixture("espn/scoreboard_20261105.json")
    competitors = payload["events"][0]["competitions"][0]["competitors"]
    home = next(c for c in competitors if c["homeAway"] == "home")
    home["linescores"][0]["value"] = 30.0  # Q1 33 -> 30: sums to 114 != 117
    report = smoke("20261105", fixture_fetch({"dates=20261105": payload}))
    assert report["results_reconciled"] == {"finals": 2, "reconciled": 1, "flagged": 1}
    gap = next(g for g in report["gaps"] if g["required_field"] == "reconciliation")
    assert gap["game_id"] == "401909936"
    assert "114" in gap["detail"]


def test_smoke_verification_live_with_http_client(load_fixture):
    # A real HttpClient (with a fake session) earns the LIVE_VERIFIED
    # label; a bare fixture loader stays FIXTURE_VERIFIED.
    class _FakeResponse:
        status_code = 200
        headers: Dict[str, str] = {}

        def __init__(self, text: str) -> None:
            self.text = text

    class _FakeSession:
        def __init__(self, mapping: Dict[str, Any]) -> None:
            self.mapping = mapping

        def get(self, url: str, headers=None, timeout=None) -> "_FakeResponse":
            for needle, payload in self.mapping.items():
                if needle in url:
                    return _FakeResponse(payload)
            raise KeyError(f"no fixture mapped for URL: {url}")

    import json

    scoreboard_text = json.dumps(load_fixture("espn/scoreboard_20261106.json"))
    predictor_text = json.dumps(load_fixture("espn/predictor_401909287.json"))
    client = HttpClient(
        session=_FakeSession(
            {"dates=20261106": scoreboard_text, "competitions/401909287/predictor": predictor_text}
        ),
        retries=0,
    )
    report = smoke("20261106", client.get_json)
    assert report["verification"] == "LIVE_VERIFIED"
    assert report["bpi_complete"]["complete"] == 1

    from ingestion.acquisition import AcquisitionClient

    wrapped = smoke("20261106", AcquisitionClient(client.get_json, request_budget=100))
    assert wrapped["verification"] == "LIVE_VERIFIED"
    assert wrapped["bpi_complete"]["complete"] == 1
