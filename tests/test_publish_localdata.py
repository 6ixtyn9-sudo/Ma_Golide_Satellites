from __future__ import annotations

import json
from pathlib import Path

import pytest

from ingestion.league_registry import SPECS
from scripts.publish_localdata import (
    STATE_KEYS,
    allowed_lifecycle_recovery_file,
    allowed_operator_file,
    merge_downloaded_localdata,
)


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
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12.txt",
        b"frozen-ticket-card\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-13.txt",
        b"hourly-candidate-card\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-14_no_bet.json",
        b"no-bet-decision\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-14.txt",
        b"hourly-no-bet-card\n",
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
    assert (repo / "localdata" / "nba" / "picks_2026-11-12.txt").read_bytes() == b"picks\n"
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12_decision.json").read_bytes() == b"receipt\n"
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12.txt").read_bytes() == b"frozen-ticket-card\n"
    # A pre-cutoff card is a mutable published preview, not a receipt or date lock.
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-13.txt").read_bytes() == b"hourly-candidate-card\n"
    assert not (repo / "localdata" / "nba" / "auto_tickets_2026-11-13_decision.json").exists()
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-14_no_bet.json").read_bytes() == b"no-bet-decision\n"
    assert not (repo / "localdata" / "nba" / "auto_tickets_2026-11-14.txt").exists()
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12.txt",
        b"later-settlement-rewrite\n",
    )
    settled = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert settled["copied"] == ["localdata/nba/auto_tickets_2026-11-12.txt"]
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12.txt").read_bytes() == b"later-settlement-rewrite\n"
    for state in ("nba", "cba", "wnba"):
        assert not (repo / "localdata" / state / "picks_today.json").exists()
    assert historical.read_text() == "retain me"
    assert not (repo / "localdata" / "nba" / "warehouse.duckdb").exists()
    assert not (repo / "localdata" / "nba" / "picks_unapproved.json").exists()


def test_publisher_state_keys_follow_the_central_registry_and_publish_expansion_lanes(tmp_path):
    assert STATE_KEYS == {spec.state_key for spec in SPECS}
    artifact_file(
        tmp_path,
        "acb-artifact",
        "spain_acb",
        "picks_2026-11-18.txt",
        b"expansion draft\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["copied"] == ["localdata/spain_acb/picks_2026-11-18.txt"]
    assert report["skipped_states"] == []


def test_publisher_reports_unknown_state_directories_without_publishing_them(tmp_path):
    artifact_file(tmp_path, "unknown-artifact", "unregistered_league", "picks_2026-11-18.txt", b"x\n")

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["status"] == "no_changes"
    assert report["skipped_states"] == [
        (tmp_path / "artifacts" / "unknown-artifact" / "localdata" / "unregistered_league").resolve().as_posix()
    ]
    assert not (tmp_path / "repo" / "localdata" / "unregistered_league").exists()


def test_publisher_keeps_structured_pick_json_artifact_only_globally(tmp_path):
    dated_json_name = "picks_2026-11-18.json"
    lifecycle_name = "picks_2026-11-18_lifecycle.json"
    freeze_name = "pick_freeze_2026-11-18_0123456789abcdef.json"
    for state in ("nba", "ncaa_w"):
        artifact_file(tmp_path, f"{state}-artifact", state, lifecycle_name, b"detailed lifecycle\n")
    artifact_file(tmp_path, "nba-artifact", "nba", dated_json_name, b"draft one\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-18.txt", b"draft one\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "pick_lifecycle_state.json", b"state one\n")
    freeze = artifact_file(tmp_path, "nba-artifact", "nba", freeze_name, b"frozen evidence\n")

    repo = tmp_path / "repo"
    first = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert first["copied_count"] == 3
    assert not allowed_operator_file(dated_json_name)
    assert not allowed_operator_file(lifecycle_name)
    assert not allowed_lifecycle_recovery_file(dated_json_name)
    assert not allowed_lifecycle_recovery_file(lifecycle_name)
    for state in ("nba", "ncaa_w"):
        assert not (repo / "localdata" / state / lifecycle_name).exists()
    destination = repo / "localdata" / "nba"
    assert not (destination / dated_json_name).exists()
    assert (destination / "picks_2026-11-18.txt").read_bytes() == b"draft one\n"
    assert (destination / "pick_lifecycle_state.json").read_bytes() == b"state one\n"

    artifact_file(tmp_path, "nba-artifact", "nba", dated_json_name, b"draft two\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-18.txt", b"draft two\n")
    second = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert second["copied"] == ["localdata/nba/picks_2026-11-18.txt"]
    assert not (destination / dated_json_name).exists()
    assert (destination / "picks_2026-11-18.txt").read_bytes() == b"draft two\n"

    freeze.write_bytes(b"rewritten frozen evidence\n")
    with pytest.raises(ValueError, match="immutable pick freeze differs"):
        merge_downloaded_localdata(tmp_path / "artifacts", repo)


def test_interrupted_lane_artifact_can_publish_only_durable_pick_lifecycle_evidence(tmp_path):
    artifact = "MaGolide-lifecycle-nba-run"
    dated_json_name = "picks_2026-11-18.json"
    freeze_name = "pick_freeze_2026-11-18_0123456789abcdef.json"
    artifact_file(tmp_path, artifact, "nba", dated_json_name, b"raw draft\n")
    artifact_file(tmp_path, artifact, "nba", "picks_2026-11-18.txt", b"draft card\n")
    artifact_file(tmp_path, artifact, "nba", freeze_name, b"freeze\n")
    artifact_file(tmp_path, artifact, "nba", "pick_lifecycle_state.json", b"lifecycle\n")
    artifact_file(
        tmp_path,
        artifact,
        "nba",
        "auto_tickets_2026-11-18_decision.json",
        b"must-not-publish\n",
    )
    artifact_file(
        tmp_path,
        artifact,
        "nba",
        "auto_tickets_2026-11-18.txt",
        b"must-not-publish\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    destination = tmp_path / "repo" / "localdata" / "nba"
    assert not (destination / dated_json_name).exists()
    assert (destination / "picks_2026-11-18.txt").read_bytes() == b"draft card\n"
    assert (destination / freeze_name).read_bytes() == b"freeze\n"
    assert (destination / "pick_lifecycle_state.json").read_bytes() == b"lifecycle\n"
    assert not (destination / "auto_tickets_2026-11-18_decision.json").exists()
    assert not (destination / "auto_tickets_2026-11-18.txt").exists()
    assert {Path(path).name for path in report["recovery_files_skipped"]} == {
        dated_json_name,
        "auto_tickets_2026-11-18_decision.json",
        "auto_tickets_2026-11-18.txt",
    }
    assert allowed_lifecycle_recovery_file(freeze_name)
    assert not allowed_lifecycle_recovery_file("auto_tickets_2026-11-18_decision.json")


def test_publisher_accepts_isolated_euroleague_w_operator_state(tmp_path):
    artifact_file(
        tmp_path,
        "elw-artifact",
        "euroleague_w",
        "auto_tickets_2026-11-18_no_bet.json",
        b"frozen-no-bet\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["copied"] == ["localdata/euroleague_w/auto_tickets_2026-11-18_no_bet.json"]


def test_publisher_strips_hourly_telemetry_before_git_publication(tmp_path):
    state = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_state.json",
        json.dumps(
            {
                "schema_version": "magolide.paper_ticket_state",
                "frozen_by_date": {"2026-11-12": {"final": True, "decision_id": "receipt"}},
                "scans": [{"run_id": "hourly-1", "large": "x" * 100_000}],
                "last_generated_utc": "2026-11-12T18:00:00Z",
            }
        ).encode(),
    )
    performance = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_performance.json",
        json.dumps(
            {
                "schema_version": "magolide.operator",
                "generated_utc": "2026-11-12T18:00:00Z",
                "history": [{"date": "2026-11-12", "positions": []}],
            }
        ).encode(),
    )

    first = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    destination = tmp_path / "repo" / "localdata" / "nba"
    saved_state = json.loads((destination / state.name).read_text())
    saved_performance = json.loads((destination / performance.name).read_text())
    assert first["copied_count"] == 2
    assert saved_state["frozen_by_date"]["2026-11-12"]["decision_id"] == "receipt"
    assert "scans" not in saved_state and "last_generated_utc" not in saved_state
    assert "generated_utc" not in saved_performance

    # A new hourly scan/clock value alone cannot create another Git update.
    state.write_text(
        json.dumps(
            {
                **json.loads(state.read_text()),
                "scans": [{"run_id": "hourly-2", "large": "y" * 100_000}],
                "last_generated_utc": "2026-11-12T19:00:00Z",
            }
        )
    )
    performance.write_text(
        json.dumps(
            {
                **json.loads(performance.read_text()),
                "generated_utc": "2026-11-12T19:00:00Z",
            }
        )
    )
    second = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert second["status"] == "no_changes"


def test_git_no_bet_decision_blocks_a_stale_cached_ticket_receipt(tmp_path):
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"stale-ticket\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba"
    destination.mkdir(parents=True)
    (destination / "auto_tickets_2026-11-12_no_bet.json").write_text(
        json.dumps({"status": "frozen_no_bet", "decision_id": "no-bet"}) + "\n"
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["status"] == "no_changes"
    assert report["no_bet_receipts_skipped"] == [
        "localdata/nba/auto_tickets_2026-11-12_decision.json"
    ]
    assert not (destination / "auto_tickets_2026-11-12_decision.json").exists()


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
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
        "auto_tickets_performance.txt",
        "auto_tickets_2026-11-12.txt",
        "auto_tickets_2026-11-12_decision.json",
        "auto_tickets_2026-11-12_no_bet.json",
        "coverage_gap_2026-11-12.json",
    }
    assert all(allowed_operator_file(name) for name in expected)
    assert all(
        allowed_operator_file(name)
        for name in {
            "picks_2026-11-12.txt",
            "pick_lifecycle_state.json",
            "pick_freeze_2026-11-12_0123456789abcdef.json",
        }
    )
    assert not allowed_operator_file("picks_2026-11-12.json")
    assert not allowed_operator_file("picks_2026-11-12_lifecycle.json")
    assert not allowed_operator_file("picks_today.json")
    assert not allowed_operator_file("picks_audit_rolling.json")
    assert not allowed_operator_file("picks_audit_2026-11-12.md")
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
    assert output["copied_count"] == 0
    assert not (tmp_path / "repo" / "localdata" / "nba" / "picks_today.json").exists()
