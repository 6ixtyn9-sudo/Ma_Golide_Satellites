from __future__ import annotations

import json
from pathlib import Path

import pytest

from scripts.publish_localdata import allowed_operator_file, merge_downloaded_localdata


def artifact_file(tmp_path: Path, artifact: str, state: str, name: str, content: bytes):
    path = tmp_path / "artifacts" / artifact / "localdata" / state / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def test_merge_publishes_only_operator_files_and_retains_history(tmp_path):
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_today.json", b"today\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-12.txt", b"picks\n")
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"receipt\n",
    )
    artifact_file(tmp_path, "nba-artifact", "nba", "warehouse.duckdb", b"canonical")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_unapproved.json", b"unexpected")
    artifact_file(tmp_path, "cba-artifact", "cba", "picks_today.json", b"cba\n")
    artifact_file(tmp_path, "wnba-artifact", "wnba", "picks_today.json", b"wnba\n")
    skipped = tmp_path / "artifacts" / "skipped-cba"
    skipped.mkdir(parents=True)
    (skipped / "daily-output.json").write_text("{}")
    repo = tmp_path / "repo"
    historical = repo / "localdata" / "nba" / "picks_2026-11-11.txt"
    historical.parent.mkdir(parents=True)
    historical.write_text("retain me")

    report = merge_downloaded_localdata(tmp_path / "artifacts", repo)

    assert report["copied_count"] == 5
    assert (repo / "localdata" / "nba" / "picks_today.json").read_bytes() == b"today\n"
    assert (repo / "localdata" / "cba" / "picks_today.json").read_bytes() == b"cba\n"
    assert (repo / "localdata" / "wnba" / "picks_today.json").read_bytes() == b"wnba\n"
    assert historical.read_text() == "retain me"
    assert not (repo / "localdata" / "nba" / "warehouse.duckdb").exists()
    assert not (repo / "localdata" / "nba" / "picks_unapproved.json").exists()


def test_immutable_receipt_conflict_aborts_without_overwriting(tmp_path):
    source = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"new-ticket\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba" / source.name
    destination.parent.mkdir(parents=True)
    destination.write_bytes(b"original-frozen-ticket\n")

    with pytest.raises(ValueError, match="immutable ticket receipt differs"):
        merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert destination.read_bytes() == b"original-frozen-ticket\n"


def test_identical_receipt_is_verified_and_not_rewritten(tmp_path):
    source = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"same-receipt\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba" / source.name
    destination.parent.mkdir(parents=True)
    destination.write_bytes(source.read_bytes())

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert report["unchanged_count"] == 1
    assert report["immutable_receipts_verified"] == [
        "localdata/nba/auto_tickets_2026-11-12_decision.json"
    ]


def test_withdrawn_cached_receipt_is_not_republished(tmp_path):
    artifact_file(
        tmp_path,
        "ncaa-m-artifact",
        "ncaa_m",
        "auto_tickets_2026-11-13_decision.json",
        json.dumps({"decision_id": "defective-run"}).encode(),
    )
    localdata = tmp_path / "repo" / "localdata" / "ncaa_m"
    localdata.mkdir(parents=True)
    (localdata / "auto_tickets_state.json").write_text(
        json.dumps(
            {
                "frozen_by_date": {
                    "2026-11-13": {
                        "withdrawn": True,
                        "decision_id": "defective-run",
                    }
                }
            }
        )
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert report["withdrawn_receipts_skipped"] == [
        "localdata/ncaa_m/auto_tickets_2026-11-13_decision.json"
    ]
    assert not (localdata / "auto_tickets_2026-11-13_decision.json").exists()


def test_allowlist_is_narrow_and_missing_artifact_directory_is_a_noop(tmp_path):
    expected = {
        "picks_today.json",
        "picks_2026-11-12.json",
        "picks_2026-11-12.txt",
        "picks_audit_rolling.json",
        "picks_audit_2026-11-12.md",
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
        "auto_tickets_performance.txt",
        "auto_tickets_2026-11-12.txt",
        "auto_tickets_2026-11-12_decision.json",
    }
    assert all(allowed_operator_file(name) for name in expected)
    assert not allowed_operator_file("picks_unapproved.json")
    assert not allowed_operator_file("auto_tickets_2026-11-12.json")
    assert not allowed_operator_file("warehouse.duckdb")
    report = merge_downloaded_localdata(tmp_path / "missing", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert not (tmp_path / "repo" / "localdata").exists()


def test_cli_reports_json(tmp_path, capsys, monkeypatch):
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_today.json", b"[]\n")
    from scripts.publish_localdata import main

    import sys

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "publish_localdata.py",
            "--artifacts",
            str(tmp_path / "artifacts"),
            "--repo",
            str(tmp_path / "repo"),
        ],
    )
    assert main() == 0
    output = json.loads(capsys.readouterr().out)
    assert output["copied_count"] == 1
