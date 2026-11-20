from __future__ import annotations

import json
from pathlib import Path
import tarfile

import pytest

from scripts.release_archive import (
    ARCHIVE_SCHEMA,
    DEFAULT_RELEASE_TAG,
    build_release_archives,
    restore_canonical_checkpoint,
)


def artifact_file(root: Path, artifact: str, state: str, relative: str, raw: bytes) -> Path:
    path = root / artifact / "localdata" / state / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(raw)
    return path


def test_release_archive_contains_only_whitelisted_evidence_and_restores_canonical_csv(tmp_path):
    artifacts = tmp_path / "artifacts"
    games = artifact_file(
        artifacts, "nba-success", "nba", "csv/games.csv.gz", b"canonical-games-gzip-bytes"
    )
    artifact_file(
        artifacts, "nba-success", "nba", "csv/predictions.csv.gz", b"canonical-predictions"
    )
    artifact_file(artifacts, "nba-success", "nba", "reports/markets_NBA_2026-11-18.json", b"{}\n")
    artifact_file(artifacts, "nba-success", "nba", "picks_2026-11-18.json", b"[]\n")
    artifact_file(artifacts, "nba-success", "nba", "warehouse.duckdb", b"must-not-archive")
    artifact_file(artifacts, "nba-success", "nba", "provider_page.html", b"must-not-archive")

    repo = tmp_path / "repo"
    operator = repo / "localdata" / "nba" / "pick_lifecycle_state.json"
    operator.parent.mkdir(parents=True)
    operator.write_text('{"schema_version":"magolide.pick_lifecycle_state"}\n')
    output = tmp_path / "out"

    report = build_release_archives(
        artifacts=artifacts, repo=repo, output=output, release_tag=DEFAULT_RELEASE_TAG
    )

    assert report["schema_version"] == ARCHIVE_SCHEMA
    assert len(report["archives"]) == 1
    archive = output / "magolide-state-nba.tar.gz"
    manifest_path = repo / "localdata" / "nba" / "archive_manifest.json"
    manifest = json.loads(manifest_path.read_text())
    assert manifest["state_key"] == "nba"
    assert manifest["release_tag"] == DEFAULT_RELEASE_TAG
    assert manifest["asset_name"] == archive.name
    assert manifest["canonical_csv_files"] == 2

    with tarfile.open(archive, "r:gz") as bundle:
        names = bundle.getnames()
    assert "archive_content_manifest.json" in names
    assert "canonical/nba/csv/games.csv.gz" in names
    assert "artifact/nba/reports/markets_NBA_2026-11-18.json" in names
    assert "operator/nba/pick_lifecycle_state.json" in names
    assert not any("warehouse.duckdb" in name or "provider_page.html" in name for name in names)

    restored = tmp_path / "restored"
    restore_report = restore_canonical_checkpoint(
        archive=archive,
        state_key="nba",
        destination=restored,
        expected_manifest=manifest_path,
    )
    assert restore_report["restored"] == ["csv/games.csv.gz", "csv/predictions.csv.gz"]
    assert (restored / "csv/games.csv.gz").read_bytes() == games.read_bytes()
    assert not (restored / "pick_lifecycle_state.json").exists()


def test_release_archive_is_deterministic_and_checksum_guarded(tmp_path):
    artifacts = tmp_path / "artifacts"
    artifact_file(artifacts, "aba-success", "aba_league", "csv/games.csv.gz", b"aba-canonical")
    repo = tmp_path / "repo"
    first = tmp_path / "first"
    second = tmp_path / "second"

    build_release_archives(artifacts=artifacts, repo=repo, output=first)
    build_release_archives(artifacts=artifacts, repo=repo, output=second)
    first_archive = first / "magolide-state-aba_league.tar.gz"
    second_archive = second / "magolide-state-aba_league.tar.gz"
    assert first_archive.read_bytes() == second_archive.read_bytes()

    manifest = repo / "localdata" / "aba_league" / "archive_manifest.json"
    corrupted = tmp_path / "corrupt.tar.gz"
    corrupted.write_bytes(first_archive.read_bytes() + b"corruption")
    with pytest.raises(ValueError, match="archive checksum mismatch"):
        restore_canonical_checkpoint(
            archive=corrupted,
            state_key="aba_league",
            destination=tmp_path / "restore",
            expected_manifest=manifest,
        )
