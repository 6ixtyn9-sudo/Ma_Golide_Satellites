"""Deep ESPN history preserves old-score and native-period coverage honestly."""

from __future__ import annotations

import json

from ingestion.duckdb_store import DuckDBStore
from ingestion.models import FinalStatus, GameStatus, League, PeriodKind
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers.espn import fetch_historical_event
from ingestion.store import InMemoryStore


def test_deep_backfill_persists_1947_final_only_and_advances_cursor(tmp_path):
    calls = []

    def fetch(url):
        calls.append(url)
        if "events?dates=1947" in url:
            path = "tests/fixtures/espn/core_events_nba_1947.json"
        elif "summary?event=1770402009" in url:
            path = "tests/fixtures/espn/history_nba_1947_final_only.json"
        else:  # pragma: no cover - defensive route assertion
            raise AssertionError(url)
        with open(path) as handle:
            return json.load(handle)

    store = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=True)
    report = Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd="20261107",
            history_days=0,
            deep_history_event_budget=1,
            deep_history_start_season=1947,
            fixture_now_utc="2026-11-07T18:00:00Z",
        ),
        store,
        fetch=fetch,
        allow_writes=True,
    ).backfill()

    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "acquired"
    assert deep["events_acquired"] == 1
    assert deep["final_scores_written"] == 1
    assert deep["coverage"] == {"final_score_only": 1}
    assert deep["cursor"] == {"season": 1946, "event_index": 0}
    assert report["backfill"]["ledger_dates"] == ["19470402"]
    assert report["backfill"]["learning_or_calibration_claimed"] is False

    game = store.get_game("1770402009")
    assert game.status is GameStatus.FINAL
    assert game.phase == "Playoffs"
    assert game.actual.home_final == 73
    assert game.actual.away_final == 68
    assert game.actual.period_kind is PeriodKind.UNKNOWN
    assert game.actual.home_quarters == []
    assert game.actual.reconciles()[0] is False
    assert game.final_status() is None
    assert game.evidence["espn"]["history"] == {
        "schema_version": "magolide.espn.history-event.v1",
        "index_year": 1947,
        "coverage": "final_score_only",
        "period_note": "final score present without native linescores",
        "provider_event_state": "post",
        "regulation_periods": 4,
    }
    health = store.get_provider_health("espn-history", League.NBA)
    assert health.coverage["next_season"] == 1946
    assert health.coverage["floor_season"] == 1947
    assert any(gap["required_field"] == "native_periods" for gap in report["gaps"])
    store.close()

    reopened = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=False)
    try:
        assert reopened.get_game(game.game_id).evidence == game.evidence
        assert reopened.get_provider_health(
            "espn-history", League.NBA
        ).coverage["next_season"] == 1946
    finally:
        reopened.close()


def test_early_wnba_uses_event_native_halves_not_modern_quarters():
    def fetch(_url):
        with open("tests/fixtures/espn/history_wnba_1997_halves.json") as handle:
            return json.load(handle)

    game = fetch_historical_event(
        "170621006",
        fetch,
        League.WNBA,
        observed_utc="2026-11-07T18:00:00Z",
        index_year=1997,
    )
    assert game.actual.period_kind is PeriodKind.HALVES
    assert game.actual.home_quarters == [30.0, 37.0]
    assert game.actual.away_quarters == [35.0, 36.0]
    assert game.actual.reconciles() == (True, "")
    assert game.final_status() is FinalStatus.FT
    assert game.evidence["espn"]["history"]["coverage"] == "reconciled_periods"
    assert game.evidence["espn"]["history"]["regulation_periods"] == 2
    assert game.evidence["espn"]["summary"]["format"]["regulation"]["periods"] == 2


def test_deep_cursor_cannot_advance_in_a_dry_run():
    def fetch(_url):  # pragma: no cover - must not be called
        raise AssertionError("dry run unexpectedly requested deep history")

    report = Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd="20261107",
            history_days=0,
            deep_history_event_budget=25,
            fixture_now_utc="2026-11-07T18:00:00Z",
        ),
        InMemoryStore(),
        fetch=fetch,
    ).backfill()
    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "disabled_canonical_write_required"
    assert deep["events_acquired"] == 0
    assert report["acquisition"]["network_requests"] == 0
