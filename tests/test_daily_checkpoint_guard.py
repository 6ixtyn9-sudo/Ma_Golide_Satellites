from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from ingestion.duckdb_store import DuckDBStore
from scripts.daily import _checkpoint_manifest, assert_checkpoint_not_regressed


def _args(**changes) -> SimpleNamespace:
    values = {
        "league": "NCAA_M",
        "localdata": "localdata",
        "write": True,
        "allow_checkpoint_regression": False,
    }
    values.update(changes)
    return SimpleNamespace(**values)


def _warehouse(path: Path, aliases: int = 1) -> DuckDBStore:
    store = DuckDBStore(str(path))
    for index in range(aliases):
        store.upsert_team_alias(f"canonical-{index}", ["variant"])
    return store


def _state_manifest(root: Path, counts: dict) -> Path:
    path = root / "localdata" / "ncaa_m" / "archive_manifest.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"state_key": "ncaa_m", "canonical_row_counts": counts}))
    return path


@pytest.fixture(autouse=True)
def _isolated_manifest_env(monkeypatch):
    monkeypatch.delenv("MAGOLIDE_CHECKPOINT_MANIFEST", raising=False)


def test_write_run_refuses_a_warehouse_thinner_than_its_checkpoint(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    manifest = _state_manifest(tmp_path, {"games": 9601, "predictions": 234})
    store = _warehouse(tmp_path / "warehouse.duckdb")
    try:
        with pytest.raises(ValueError, match="checkpoint_regression"):
            assert_checkpoint_not_regressed(_args(), store)
    finally:
        store.close()
    assert _checkpoint_manifest(_args()) == Path("localdata/ncaa_m/archive_manifest.json")
    assert manifest.is_file()


def test_operator_can_run_through_the_guard_on_purpose(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    _state_manifest(tmp_path, {"predictions": 234})
    store = _warehouse(tmp_path / "warehouse.duckdb")
    try:
        assert_checkpoint_not_regressed(_args(allow_checkpoint_regression=True), store)
    finally:
        store.close()


def test_lane_without_a_committed_checkpoint_runs(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    store = _warehouse(tmp_path / "warehouse.duckdb")
    try:
        assert_checkpoint_not_regressed(_args(), store)
    finally:
        store.close()


def test_manifest_without_row_counts_cannot_judge(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    path = tmp_path / "localdata" / "ncaa_m" / "archive_manifest.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"state_key": "ncaa_m"}))
    store = _warehouse(tmp_path / "warehouse.duckdb")
    try:
        assert_checkpoint_not_regressed(_args(), store)
    finally:
        store.close()


def test_manifest_next_to_the_warehouse_wins_for_isolated_lanes(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    lane = tmp_path / "localdata" / "eurocup"
    lane.mkdir(parents=True)
    manifest = lane / "archive_manifest.json"
    manifest.write_text(json.dumps({"state_key": "eurocup", "canonical_row_counts": {}}))
    assert _checkpoint_manifest(_args(league="EUROCUP", localdata=str(lane))) == manifest
