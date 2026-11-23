"""Israel Premier League (Winner League) lane: referenced JSON route, fixture-tested.

Route and field contract referenced from the league site and SportsDataverse
``hoopR`` documentation. Ma Golide retains no first-party response, so the
committed fixture is *shaped* from the documented contract rather than being a
recorded body.

The provider serves local wall-clock time with **no UTC offset**, so the
adapter applies ``Asia/Jerusalem`` explicitly and records the assumption on
every game. Cases covered: a clean regulation final, an overtime final whose
extra periods are exposed, an overtime row whose period points are absent (must
refuse), a short quarter string (must refuse), duplicate team identity (must
refuse), and the DST edge flag.
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Callable

from ingestion.models import GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import israel_ibl
from ingestion.store import InMemoryStore

FIXTURES = "tests/fixtures/"


def _fixture_json(relative: str) -> Any:
    with open(FIXTURES + relative, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _board_fetch(payload: Any | None = None) -> Callable[[str], Any]:
    body = payload if payload is not None else _fixture_json("israel_ibl/board_2027.json")

    def fetch(url: str) -> Any:
        if "ws.asmx/Games" not in url:
            raise AssertionError(f"unexpected URL: {url}")
        return body

    return fetch


def test_board_url_carries_the_documented_parameters():
    url = israel_ibl.games_url(2027)
    assert "cYear=2027" in url
    assert "board_id=5" in url
    assert "board_round=0" in url and "team_id=0" in url and "team_uid=0" in url
    assert url.startswith("https://basket.co.il/ws/ws.asmx/Games?")
    for bad in (1999,):
        try:
            israel_ibl.games_url(bad)
        except ValueError:
            continue
        raise AssertionError("season below the floor must be refused")


def test_season_end_year_uses_the_provider_convention():
    assert israel_ibl.season_end_year(datetime(2026, 11, 22)) == 2027
    assert israel_ibl.season_end_year(datetime(2027, 5, 1)) == 2027


def test_wall_clock_is_converted_with_the_documented_assumption():
    iso, ambiguous = israel_ibl._tip_utc("22/11/2026", "20:15")
    assert iso == "2026-11-22T18:15:00Z"  # IST is UTC+2 outside DST
    assert ambiguous is False
    summer, _ = israel_ibl._tip_utc("05/07/2026", "20:15")
    assert summer == "2026-07-05T17:15:00Z"  # IDT is UTC+3


def test_regulation_final_reconciles_and_records_the_assumption():
    games = israel_ibl.discover_games("2026-11-22", _board_fetch(), 2027)
    by_id = {game.game_id: game for game in games}
    assert set(by_id) == {"IBL-700001", "IBL-700002"}
    clean = by_id["IBL-700001"]
    assert clean.league is League.ISRAEL_PREMIER_LEAGUE
    # Names pass through the shared canonicaliser, which folds sponsor aliases.
    from ingestion.identity import canonical_team_name

    assert clean.home_name == canonical_team_name("Maccabi Tel Aviv")
    assert clean.away_name == canonical_team_name("Hapoel Jerusalem")
    assert clean.status is GameStatus.FINAL
    assert clean.provider_ids["israel_assumed_timezone"] == "Asia/Jerusalem"
    assert clean.evidence["tip_timezone_assumed"] == "Asia/Jerusalem"
    assert clean.actual is not None
    assert clean.actual.home_quarters == [22.0, 20.0, 24.0, 22.0]
    ok, reason = clean.actual.reconciles()
    assert ok, reason


def test_overtime_periods_are_summed_and_reconcile():
    games = israel_ibl.discover_games("2026-11-22", _board_fetch(), 2027)
    ot = {game.game_id: game for game in games}["IBL-700002"]
    actual = ot.actual
    assert actual is not None
    assert actual.overtime_periods == 1
    assert (actual.home_overtime, actual.away_overtime) == (10.0, 8.0)
    ok, reason = actual.reconciles()
    assert ok, reason


def test_overtime_without_period_points_is_refused():
    failures: list = []
    results = israel_ibl.season_results(2027, _board_fetch(), failures)
    assert {game.game_id for game in results} == {"IBL-700001", "IBL-700002"}
    reasons = {row["game_id"]: row["error"] for row in failures}
    assert "overtime is unavailable" in reasons["700003"]
    assert "fewer than four quarters" in reasons["700004"]
    assert "same team" in reasons["700005"]


def test_short_quarter_string_never_becomes_a_result():
    payload = _fixture_json("israel_ibl/board_2027.json")
    for row in payload["games"]:
        if row["game_id"] == 700002:
            row["score_by_q_1"] = "20|18"
    games = israel_ibl.discover_games("2026-11-22", _board_fetch(payload), 2027)
    assert {game.game_id for game in games} == {"IBL-700001"}
    assert any("fewer than four quarters" in gap["detail"] for gap in israel_ibl.last_gaps())


def test_empty_board_is_refused_rather_than_a_scoreless_season():
    try:
        israel_ibl.fetch_board(2027, _board_fetch({"games": []}))
    except ValueError as exc:
        assert "no games" in str(exc)
    else:
        raise AssertionError("an empty board must be refused")


def test_fetch_game_result_refreshes_from_the_board_and_checks_identity():
    game = israel_ibl.discover_games("2026-11-22", _board_fetch(), 2027)[0]
    refreshed = israel_ibl.fetch_game_result(game, _board_fetch())
    assert refreshed.game_id == game.game_id
    payload = _fixture_json("israel_ibl/board_2027.json")
    payload["games"][0]["team1_name_eng"] = "Some Other Club"
    try:
        israel_ibl.fetch_game_result(game, _board_fetch(payload))
    except Exception:
        pass
    else:
        raise AssertionError("a renamed team must refuse the refresh")


def test_smoke_reports_counts_without_raising():
    report = israel_ibl.smoke(_board_fetch(), 2027)
    assert report["games_found"] == 5
    assert report["results_reconciled"]["complete"] == 2
    assert report["assumed_timezone"] == "Asia/Jerusalem"

    def broken(_url: str):
        raise RuntimeError("boom")

    failed = israel_ibl.smoke(broken, 2027)
    assert failed["status"] == "provider_error"


def test_pipeline_israel_lane_discovers_and_settles_a_date():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="ISRAEL_PREMIER_LEAGUE", date_yyyymmdd="20261122"),
        store,
        fetch=_board_fetch(),
        allow_writes=False,
    )
    games = pipeline._discover_games()
    assert {game.game_id for game in games} == {"IBL-700001", "IBL-700002"}
    finals = pipeline._israel_ibl_finals(games)
    assert {game.game_id for game in finals} == {"IBL-700001", "IBL-700002"}
    assert all(game.actual is not None for game in finals)


def test_israel_registry_contract_declares_the_fallback_price_lane():
    from ingestion.league_registry import league_spec

    spec = league_spec(League.ISRAEL_PREMIER_LEAGUE)
    assert spec.schedule_source == "israel_ibl"
    assert spec.results_source == "israel_ibl"
    assert spec.automated_schedule_supported is True
    # The direct comparison page remains the named-book fallback; the
    # competition's own fixtures page now also feeds aggregate, research-only
    # observations (verified 2026-11-23). Neither is a cleared API lane.
    assert spec.book_providers == ("betexplorer", "betexplorer_direct")


# --- captured rows (2026-11-22 probe of basket.co.il board 5) ---------------
# Four rows were transmitted verbatim inside a slice file; the committed object
# holds those row objects unchanged and only adds the ``games`` envelope this
# adapter expects. They settle three questions the shaped fixture could not:
# quarters are points-in-period, a double-overtime row is explicit, and tips
# carry no offset.


def test_captured_rows_read_period_points_not_cumulative_totals():
    games = israel_ibl.discover_games(None, _board_fetch(_fixture_json("israel_ibl/board5_captured_rows.json")), 2027)
    by_id = {game.game_id: game for game in games}
    assert set(by_id) == {"IBL-26656", "IBL-26646", "IBL-26647", "IBL-26649"}
    played = israel_ibl.fetch_game_result(by_id["IBL-26647"], _board_fetch(
        _fixture_json("israel_ibl/board5_captured_rows.json")))
    actual = played.actual
    assert actual is not None
    # 15|22|14|18 against 16|27|24|25, finals 69-92: per-period by construction.
    assert actual.home_quarters == [15.0, 22.0, 14.0, 18.0]
    assert actual.away_quarters == [16.0, 27.0, 24.0, 25.0]
    assert (actual.home_final, actual.away_final) == (69, 92)
    assert actual.overtime_periods == 0
    assert actual.reconciles() == (True, "")


def test_captured_double_overtime_row_reconciles():
    board = _board_fetch(_fixture_json("israel_ibl/board5_captured_rows.json"))
    game = next(g for g in israel_ibl.discover_games(None, board, 2027) if g.game_id == "IBL-26656")
    actual = israel_ibl.fetch_game_result(game, board).actual
    assert actual is not None
    assert actual.home_quarters == [21.0, 18.0, 18.0, 18.0]
    assert actual.away_quarters == [20.0, 19.0, 21.0, 15.0]
    # 11|10 and 11|17: two extra periods summed, matching total_ot "2".
    assert (actual.home_overtime, actual.away_overtime) == (21.0, 28.0)
    assert actual.overtime_periods == 2
    assert (actual.home_final, actual.away_final) == (96, 103)
    assert actual.reconciles() == (True, "")


def test_captured_naive_tip_gets_the_recorded_jerusalem_offset():
    board = _board_fetch(_fixture_json("israel_ibl/board5_captured_rows.json"))
    game = next(g for g in israel_ibl.discover_games(None, board, 2027) if g.game_id == "IBL-26646")
    # 10/10/2026 20:15 local (IDT, UTC+3) is 17:15Z; the source carries no offset.
    assert game.scheduled_utc == "2026-10-10T17:15:00Z"
    assert game.provider_ids["israel_ibl"] == "26646"
    assert game.provider_ids["israel_assumed_timezone"] == "Asia/Jerusalem"
    assert game.season == "2027"
