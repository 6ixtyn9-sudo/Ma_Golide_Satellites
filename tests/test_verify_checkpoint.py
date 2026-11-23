from __future__ import annotations

import json
from pathlib import Path
import sys

from ingestion.duckdb_store import DuckDBStore
from scripts.verify_checkpoint import main as verify_main


def _warehouse(path: Path, aliases: int = 1) -> Path:
    store = DuckDBStore(str(path))
    for index in range(aliases):
        store.upsert_team_alias(f"canonical-{index}", ["variant"])
    store.close()
    return path


def _manifest(path: Path, counts: dict | None) -> Path:
    payload = {"state_key": "nba"}
    if counts is not None:
        payload["canonical_row_counts"] = counts
    path.write_text(json.dumps(payload))
    return path


def _run(monkeypatch, capsys, warehouse: Path, manifest: Path) -> tuple[int, dict]:
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "verify_checkpoint.py",
            "--warehouse",
            str(warehouse),
            "--manifest",
            str(manifest),
            "--state-key",
            "nba",
        ],
    )
    code = verify_main()
    return code, json.loads(capsys.readouterr().out)


def test_warm_warehouse_matching_the_committed_checkpoint_is_trusted(tmp_path, monkeypatch, capsys):
    warehouse = _warehouse(tmp_path / "warehouse.duckdb")
    manifest = _manifest(tmp_path / "archive_manifest.json", {"team_aliases": 1, "games": 0})

    code, report = _run(monkeypatch, capsys, warehouse, manifest)

    assert code == 0
    assert report["status"] == "ok"
    assert report["counts"]["team_aliases"] == 1


def test_thin_warm_warehouse_is_a_checkpoint_regression(tmp_path, monkeypatch, capsys):
    warehouse = _warehouse(tmp_path / "warehouse.duckdb", aliases=1)
    manifest = _manifest(
        tmp_path / "archive_manifest.json",
        {"games": 9601, "predictions": 234, "book_snapshots": 770, "team_aliases": 1},
    )

    code, report = _run(monkeypatch, capsys, warehouse, manifest)

    assert code == 1
    assert report["status"] == "checkpoint_regression"
    assert report["regressions"] == {
        "games": {"committed": 9601, "new": 0},
        "predictions": {"committed": 234, "new": 0},
        "book_snapshots": {"committed": 770, "new": 0},
    }


def test_legacy_manifest_without_counts_passes_until_the_next_build_records_them(
    tmp_path, monkeypatch, capsys
):
    warehouse = _warehouse(tmp_path / "warehouse.duckdb")
    manifest = _manifest(tmp_path / "archive_manifest.json", None)

    code, report = _run(monkeypatch, capsys, warehouse, manifest)

    assert code == 0
    assert report["status"] == "no_manifest_counts"


def test_missing_warehouse_against_a_committed_checkpoint_is_a_regression(
    tmp_path, monkeypatch, capsys
):
    manifest = _manifest(tmp_path / "archive_manifest.json", {"predictions": 234})

    code, report = _run(monkeypatch, capsys, tmp_path / "absent.duckdb", manifest)

    assert code == 1
    assert report["status"] == "checkpoint_regression"
    assert report["reason"] == "warehouse_missing"


def test_missing_manifest_is_not_a_regression(tmp_path, monkeypatch, capsys):
    warehouse = _warehouse(tmp_path / "warehouse.duckdb")

    code, report = _run(monkeypatch, capsys, warehouse, tmp_path / "absent.json")

    assert code == 0
    assert report["status"] == "no_manifest"
