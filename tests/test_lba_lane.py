"""Italy LBA lane, tested against the captured 2026-11-22 provider bodies.

Two committed fixtures are byte-exact captures of legabasket.it JSON
(``championships`` 867 bytes, focus-day ``calendar`` 11636 bytes;
``tests/test_capture_fidelity.py`` re-hashes them). The play-by-play bodies are
314 KB each and were not transmittable, so the committed period-end actions are
verbatim copies of the captured anchors inside a container this repository
wrote.

What the capture changed: ``game_status`` is a numeric code — ``"2"`` final,
``"1"`` live, ``"0"`` not started — and the focus-day slate mixes all three,
including live rows that already carry real scores. Status is therefore read
from the code, never inferred from the presence of numbers.
"""

from __future__ import annotations

import json
from typing import Any, Callable

from ingestion.http import HttpError
from ingestion.models import Game, GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import IdentityError, lba
from ingestion.store import InMemoryStore

FIXTURES = "tests/fixtures/lba/"


def _fixture_json(name: str) -> Any:
    with open(FIXTURES + name, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _fetch(extra_rows: list | None = None) -> Callable[[str], Any]:
    def fetch(url: str) -> Any:
        if "get-championships?cs_id=1" in url and "s=2026" in url:
            return _fixture_json("championships_2026.json")
        if "get-championships-calendar-by-id?id=602" in url:
            payload = _fixture_json("calendar_602.json")
            if extra_rows:
                payload = {**payload, "matches": extra_rows + payload["matches"]}
            return payload
        if "play-by-play?id=25426" in url:
            return _fixture_json("pbp_25426.json")
        if "play-by-play?id=25422" in url:
            return _fixture_json("pbp_25422_ot.json")
        if "play-by-play?id=999001" in url:
            return _fixture_json("pbp_synthetic_noncontiguous_ot.json")
        if "play-by-play?id=25428" in url:
            # The captured ledger shows this route answering 500 for a match
            # without play-by-play, so the fixture models the real failure.
            raise HttpError(500, url, "play-by-play unavailable")
        raise AssertionError(f"unexpected route {url}")

    return fetch


def _ot_game() -> Game:
    """The captured AET fixture: Cantù 92-106 Tortona, 2026-11-15T16:30Z."""

    return Game(
        game_id="LBA-25422",
        league=League.ITALY_LBA,
        season="2026",
        home_team_id="1765",
        away_team_id="1752",
        home_name="Acqua S.Bernardo Cantù",
        away_name="Bertram Derthona Tortona",
        scheduled_utc="2026-11-15T16:30:00Z",
        status=GameStatus.FINAL,
        phase="Serie A",
        provider_ids={"lba": "25422"},
    )


def test_focus_day_slate_on_the_captured_first_date():
    games = lba.discover_games("2026-11-21", _fetch(), 2026)
    assert {game.game_id for game in games} == {"LBA-25426", "LBA-25428"}
    first = next(g for g in games if g.game_id == "LBA-25426")
    # 2026-11-21T19:00:00.000+01:00 is 18:00Z; the offset is provider-given.
    assert first.scheduled_utc == "2026-11-21T18:00:00Z"
    assert (first.home_name, first.away_name) == (
        "Dolomiti Energia Trentino",
        "Napoli Basketball",
    )
    assert first.status is GameStatus.FINAL
    assert first.season == "2026"
    assert first.provider_ids == {"lba": "25426"}


def test_one_focus_day_spans_a_round_not_a_date():
    # The captured slate runs across two local dates with one round focused
    # (filters.days id 4494), so a date filter yields the matching subset only.
    second_day = lba.discover_games("2026-11-22", _fetch(), 2026)
    assert {game.game_id for game in second_day} == {
        "LBA-25427",
        "LBA-25430",
        "LBA-25429",
        "LBA-25424",
        "LBA-25425",
        "LBA-25431",
    }
    assert lba.discover_games("2026-11-23", _fetch(), 2026) == []
    # The overtime match is not on the focus day and is not fabricated into it.
    assert lba.discover_games("2026-11-15", _fetch(), 2026) == []


def test_status_comes_from_the_code_not_from_the_score():
    games = {g.game_id: g for g in lba.discover_games("2026-11-22", _fetch(), 2026)}
    # Live rows already carry real scores (63-62, 49-48, 84-80, 27-13); a
    # score-shaped row must never be promoted to a settled final.
    assert games["LBA-25429"].status is GameStatus.LIVE
    assert games["LBA-25425"].status is GameStatus.LIVE
    assert games["LBA-25431"].status is GameStatus.LIVE
    # The not-started row publishes 0/0 rather than null.
    assert games["LBA-25430"].status is GameStatus.SCHEDULED
    assert [g.game_id for g in games.values() if g.status is GameStatus.FINAL] == ["LBA-25427"]


def test_an_unknown_status_code_is_never_a_final():
    row = {
        "id": 999999,
        "game_status": "9",
        "h_team_id": 1763,
        "h_team_name": "Dolomiti Energia Trentino",
        "v_team_id": 1772,
        "v_team_name": "Napoli Basketball",
        "match_datetime": "2026-11-21T19:00:00.000+01:00",
        "home_final_score": 100,
        "visitor_final_score": 86,
    }
    games = lba.discover_games(None, _fetch([row]), 2026)
    game = next(g for g in games if g.game_id == "LBA-999999")
    assert game.status is GameStatus.UNKNOWN


def test_ambiguous_championship_resolution_is_refused():
    payload = _fixture_json("championships_2026.json")
    doubled = {**payload, "competitions": [payload["competitions"][0], payload["competitions"][0]]}
    try:
        lba.championship_id(2026, lambda _url: doubled)
    except ValueError as exc:
        assert "ambiguous" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("ambiguous championship list was accepted")


def test_naive_timestamp_becomes_a_gap_not_a_game():
    naive = {
        "id": 999998,
        "game_status": "0",
        "h_team_id": 1763,
        "h_team_name": "Dolomiti Energia Trentino",
        "v_team_id": 1772,
        "v_team_name": "Napoli Basketball",
        "match_datetime": "2026-11-21T19:00:00.000",
        "home_final_score": None,
        "visitor_final_score": None,
    }
    games = lba.discover_games(None, _fetch([naive]), 2026)
    assert all(game.game_id != "LBA-999998" for game in games)
    assert any("no UTC offset" in gap["detail"] for gap in lba.last_gaps())


def test_identical_home_and_away_names_are_refused():
    same = {
        "id": 999997,
        "game_status": "0",
        "h_team_id": 1763,
        "h_team_name": "Dolomiti Energia Trentino",
        "v_team_id": 1772,
        "v_team_name": "Dolomiti Energia Trentino",
        "match_datetime": "2026-11-21T19:00:00.000+01:00",
    }
    games = lba.discover_games(None, _fetch([same]), 2026)
    assert all(game.game_id != "LBA-999997" for game in games)
    assert any("same team on both sides" in gap["detail"] for gap in lba.last_gaps())


def test_regulation_quarters_come_from_the_captured_anchors():
    game = next(g for g in lba.discover_games("2026-11-21", _fetch(), 2026)
                if g.game_id == "LBA-25426")
    result = lba.fetch_game_result(game, _fetch())
    assert result.actual is not None
    assert result.actual.home_quarters == [22.0, 26.0, 24.0, 28.0]
    assert result.actual.away_quarters == [20.0, 30.0, 22.0, 14.0]
    assert (result.actual.home_final, result.actual.away_final) == (100, 86)
    assert result.actual.overtime_periods == 0
    assert result.actual.home_overtime is None
    assert result.actual.reconciles() == (True, "")


def test_captured_overtime_game_reconciles_period_by_period():
    # Cantù 92-106 Tortona after overtime: period ends 23-29, 41-50, 62-69,
    # 87-87 (tied), 92-106 -> quarters 23-29 / 18-21 / 21-19 / 25-18, OT 5-19.
    result = lba.fetch_game_result(_ot_game(), _fetch())
    actual = result.actual
    assert actual is not None
    assert actual.home_quarters == [23.0, 18.0, 21.0, 25.0]
    assert actual.away_quarters == [29.0, 21.0, 19.0, 18.0]
    assert (actual.home_overtime, actual.away_overtime) == (5.0, 19.0)
    assert actual.overtime_periods == 1
    assert (actual.home_final, actual.away_final) == (92, 106)
    assert actual.reconciles() == (True, "")
    assert result.status is GameStatus.FINAL


def test_non_contiguous_overtime_is_refused():
    game = Game(**{**_ot_game().__dict__, "game_id": "LBA-999001",
                   "provider_ids": {"lba": "999001"}})
    try:
        lba.fetch_game_result(game, _fetch())
    except ValueError as exc:
        assert "non-contiguous overtime" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("non-contiguous overtime was accepted")


def test_pipeline_settles_only_rows_with_reconcilable_periods():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="ITALY_LBA", date_yyyymmdd="20261121"),
        store,
        fetch=_fetch(),
        allow_writes=False,
    )
    games = pipeline._discover_games()
    assert {game.game_id for game in games} == {"LBA-25426", "LBA-25428"}
    finals = pipeline._lba_finals(games)
    # One match has captured play-by-play; the other answers 500 and stays a gap.
    assert {game.game_id for game in finals} == {"LBA-25426"}
    assert any(gap.game_id == "LBA-25428" for gap in pipeline.gaps)


def test_smoke_counts_the_captured_status_mix():
    report = lba.smoke(_fetch(), 2026)
    assert report["games_found"] == 8
    assert report["games_final"] == 3
    assert report["games_live"] == 4
    assert report["games_scheduled"] == 1
    failed = lba.smoke(lambda _url: [], 2026)
    assert failed["status"] == "provider_error"


def test_result_refresh_requires_the_provider_binding():
    game = Game(**{**_ot_game().__dict__, "provider_ids": {}})
    try:
        lba.fetch_game_result(game, _fetch())
    except IdentityError as exc:
        assert "no LBA match id" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("a missing provider binding was accepted")


def test_registry_contract_declares_the_fallback_price_lane():
    from ingestion.league_registry import league_spec

    spec = league_spec(League.ITALY_LBA)
    assert spec.schedule_source == "lba"
    assert spec.results_source == "lba"
    assert spec.automated_schedule_supported is True
    # The direct comparison page remains the named-book fallback; the
    # competition's own fixtures page now also feeds aggregate, research-only
    # observations (verified 2026-11-23). Neither is a cleared API lane.
    assert spec.book_providers == ("betexplorer", "betexplorer_direct")


def test_a_round_played_on_other_days_reports_a_receipt_not_silence():
    """LBA's calendar is round-scoped: a full round can answer no fixtures."""

    assert lba.discover_games("2026-11-23", _fetch(), 2026) == []

    assert lba.last_probe()["off_date_dropped"] == 8
