"""Italy LBA lane: referenced route contract, shaped fixtures, fail-closed edges.

The probe that established the field names ran outside this repository on
2026-11-22 (championships list, focus-day calendar with explicit offsets, and
cumulative play-by-play); no first-party body was retained, so every fixture is
shaped from that contract and says so.

Cases: the focus-day slate and its UTC conversion, ambiguous championship
resolution, a naive timestamp, duplicate team identity, quarters derived from
cumulative period ends, non-contiguous overtime, the wired pipeline path, smoke
never raising, and a missing provider binding.
"""

from __future__ import annotations

import json
from typing import Any, Callable

from ingestion.http import HttpError
from ingestion.models import Game, GameStatus
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
        if "play-by-play?id=25428" in url:
            # The second final has no captured play-by-play; a 404 is honest.
            raise HttpError(404, url, "no play-by-play for this match")
        if "play-by-play?id=999001" in url:
            return _fixture_json("pbp_noncontiguous_ot.json")
        raise AssertionError(f"unexpected route {url}")

    return fetch


def test_discover_games_returns_the_focus_day_slate():
    games = lba.discover_games("2026-11-21", _fetch(), 2026)
    assert {game.game_id for game in games} == {"LBA-25426", "LBA-25428"}
    game = next(g for g in games if g.game_id == "LBA-25426")
    # 2026-11-21T19:00:00.000+01:00 is 18:00Z; the offset is provider-given.
    assert game.scheduled_utc == "2026-11-21T18:00:00Z"
    assert game.status is GameStatus.FINAL
    assert game.home_team_id != game.away_team_id
    assert game.provider_ids == {"lba": "25426"}
    assert game.season == "2026"
    # The remaining row falls on 2026-11-22 and is excluded honestly.
    assert lba.discover_games("2026-11-23", _fetch(), 2026) == []


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
        "id": 999999,
        "h_team_id": 1763,
        "h_team_name": "Dolomiti Energia Trentino",
        "v_team_id": 1772,
        "v_team_name": "Napoli Basketball",
        "match_datetime": "2026-11-21T19:00:00.000",
        "home_final_score": None,
        "visitor_final_score": None,
    }
    games = lba.discover_games(None, _fetch([naive]), 2026)
    assert all(game.game_id != "LBA-999999" for game in games)
    assert any("no UTC offset" in gap["detail"] for gap in lba.last_gaps())


def test_identical_home_and_away_names_are_refused():
    same = {
        "id": 999998,
        "h_team_id": 1763,
        "h_team_name": "Dolomiti Energia Trentino",
        "v_team_id": 1772,
        "v_team_name": "Dolomiti Energia Trentino",
        "match_datetime": "2026-11-21T19:00:00.000+01:00",
    }
    games = lba.discover_games(None, _fetch([same]), 2026)
    assert all(game.game_id != "LBA-999998" for game in games)
    assert any("same team on both sides" in gap["detail"] for gap in lba.last_gaps())


def test_quarters_derive_from_cumulative_play_by_play_and_reconcile():
    game = lba.discover_games("2026-11-21", _fetch(), 2026)[0]
    result = lba.fetch_game_result(game, _fetch())
    assert result.actual is not None
    assert result.actual.home_quarters == [22.0, 26.0, 24.0, 28.0]
    assert result.actual.away_quarters == [20.0, 30.0, 22.0, 14.0]
    assert (result.actual.home_final, result.actual.away_final) == (100, 86)
    assert result.actual.overtime_periods == 0
    assert result.actual.reconciles() == (True, "")


def test_non_contiguous_overtime_is_refused():
    game = lba.discover_games("2026-11-21", _fetch(), 2026)[0]
    broken = Game(**{**game.__dict__, "game_id": "LBA-999001", "provider_ids": {"lba": "999001"}})
    try:
        lba.fetch_game_result(broken, _fetch())
    except ValueError as exc:
        assert "non-contiguous overtime" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("non-contiguous overtime was accepted")


def test_pipeline_lane_is_wired_end_to_end():
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
    # Only the match with a captured play-by-play settles; the other stays a gap.
    assert {game.game_id for game in finals} == {"LBA-25426"}
    assert all(game.actual is not None for game in finals)
    assert any(gap.game_id == "LBA-25428" for gap in pipeline.gaps)


def test_smoke_reports_counts_and_never_raises():
    ok = lba.smoke(_fetch(), 2026)
    assert ok["games_found"] == 3
    assert ok["games_final"] == 2
    failed = lba.smoke(lambda _url: [], 2026)
    assert failed["status"] == "provider_error"


def test_result_refresh_requires_the_provider_binding():
    game = lba.discover_games("2026-11-21", _fetch(), 2026)[0]
    stripped = Game(**{**game.__dict__, "provider_ids": {}})
    try:
        lba.fetch_game_result(stripped, _fetch())
    except IdentityError as exc:
        assert "no LBA match id" in str(exc)
    else:  # pragma: no cover
        raise AssertionError("a missing provider binding was accepted")


def test_registry_contract_declares_the_fallback_price_lane():
    from ingestion.league_registry import league_spec
    from ingestion.models import League

    spec = league_spec(League.ITALY_LBA)
    assert spec.schedule_source == "lba"
    assert spec.results_source == "lba"
    assert spec.automated_schedule_supported is True
    assert spec.book_providers == ("betexplorer_direct",)
