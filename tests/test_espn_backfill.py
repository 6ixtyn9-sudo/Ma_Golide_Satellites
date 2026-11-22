"""ESPN deep-history backfill: permanent 404s advance, transient errors retain the cursor."""

from __future__ import annotations

import json

from ingestion.duckdb_store import DuckDBStore
from ingestion.http import HttpError
from ingestion.models import League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import espn

SEASON = 2001  # the capped NBA floor season; also exercises the cap directly


def _index_payload(event_ids):
    return {
        "count": len(event_ids),
        "pageCount": 1,
        "items": [
            {
                "$ref": (
                    "https://sports.core.api.espn.com/v2/sports/basketball/"
                    f"leagues/nba/events/{event_id}"
                )
            }
            for event_id in event_ids
        ],
    }


def _summary_fixture():
    with open("tests/fixtures/espn/history_nba_1947_final_only.json") as handle:
        return json.load(handle)


def _config(budget: int, start: int = SEASON) -> PipelineConfig:
    return PipelineConfig(
        league="NBA",
        date_yyyymmdd="20261107",
        history_days=0,
        deep_history_event_budget=budget,
        deep_history_start_season=start,
        fixture_now_utc="2026-11-07T18:00:00Z",
    )


def test_permanent_404_is_recorded_and_cursor_advances(tmp_path):
    calls = []

    def fetch(url):
        calls.append(url)
        if "events?dates=2001" in url:
            return _index_payload(["900000001", "1770402009"])
        if "summary?event=900000001" in url:
            raise HttpError(404, url)
        if "summary?event=1770402009" in url:
            return _summary_fixture()
        raise AssertionError(url)  # pragma: no cover - defensive route assertion

    store = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=True)
    report = Pipeline(_config(2), store, fetch=fetch, allow_writes=True).backfill()
    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "acquired_with_permanent_skips"
    assert deep["events_acquired"] == 1
    assert deep["permanent_skips"] == 1
    assert deep["cursor"] == {"season": 2000, "event_index": 0}
    gaps = [g for g in report["gaps"] if g["required_field"] == "deep_history_event"]
    assert len(gaps) == 1
    assert gaps[0]["reason"] == "missing"
    assert gaps[0]["game_id"] == "900000001"
    assert "404" in gaps[0]["detail"]
    health = store.get_provider_health("espn-history", League.NBA)
    assert health.coverage["next_season"] == 2000
    assert health.coverage["next_event_index"] == 0
    # the run never re-requests the wedged event: one index page + two summaries
    assert sum(1 for url in calls if "summary?event=900000001" in url) == 1
    store.close()


def test_all_permanent_404s_advance_cursor_without_acquisition(tmp_path):
    def fetch(url):
        if "events?dates=2001" in url:
            return _index_payload(["900000001", "900000002"])
        if "summary?event=" in url:
            raise HttpError(404, url)
        raise AssertionError(url)  # pragma: no cover - defensive route assertion

    store = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=True)
    deep = Pipeline(_config(2), store, fetch=fetch, allow_writes=True).backfill()[
        "backfill"
    ]["deep_history"]
    assert deep["status"] == "acquired_with_permanent_skips"
    assert deep["events_acquired"] == 0
    assert deep["permanent_skips"] == 2
    assert deep["cursor"] == {"season": 2000, "event_index": 0}
    # the wedge is gone: the next run reads the advanced cursor, sits below the
    # capped floor, and reports completion without a single request
    def refused_fetch(_url):  # pragma: no cover - must not be called
        raise AssertionError("below-floor re-run unexpectedly requested the index")

    resumed_config = PipelineConfig(
        league="NBA",
        date_yyyymmdd="20261107",
        history_days=0,
        deep_history_event_budget=2,
        fixture_now_utc="2026-11-07T18:00:00Z",
    )
    second = Pipeline(resumed_config, store, fetch=refused_fetch, allow_writes=True).backfill()
    assert second["backfill"]["deep_history"]["status"] == "complete_to_verified_index_floor"
    store.close()


def test_transient_error_retains_cursor(tmp_path):
    def fetch(url):
        if "events?dates=2001" in url:
            return _index_payload(["900000001"])
        if "summary?event=900000001" in url:
            raise HttpError(503, url)
        raise AssertionError(url)  # pragma: no cover - defensive route assertion

    store = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=True)
    deep = Pipeline(_config(2), store, fetch=fetch, allow_writes=True).backfill()[
        "backfill"
    ]["deep_history"]
    assert deep["status"] == "provider_error_cursor_retained"
    assert deep["events_acquired"] == 0
    assert deep["permanent_skips"] == 0
    assert deep["cursor"] == {"season": 2001, "event_index": 0}
    # fail-closed: a retained cursor persists no health record at all
    assert store.get_provider_health("espn-history", League.NBA) is None
    store.close()


def test_nba_history_floor_is_capped_at_2001(tmp_path):
    assert espn.HISTORY_FLOORS[League.NBA] == 2001

    def fetch(_url):  # pragma: no cover - must not be called
        raise AssertionError("below-floor backfill unexpectedly requested ESPN")

    store = DuckDBStore(str(tmp_path / "nba" / "warehouse.duckdb"), persist=True)
    deep = Pipeline(_config(50, start=1994), store, fetch=fetch, allow_writes=True).backfill()[
        "backfill"
    ]["deep_history"]
    assert deep["status"] == "complete_to_verified_index_floor"
    assert deep["floor_season"] == 2001
    assert deep["cursor"] == {"season": 1994, "event_index": 0}
    assert deep["events_acquired"] == 0
    store.close()
