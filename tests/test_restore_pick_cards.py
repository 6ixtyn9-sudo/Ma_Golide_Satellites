from __future__ import annotations

import hashlib
from pathlib import Path
import subprocess

import pytest

from scripts.release_archive import build_release_archives
from scripts.restore_pick_cards import restore_pick_cards


def readable_card(state: str, day: str, body: str = "  (none)\n") -> bytes:
    return (
        f"Ma Golide Picks — {state.upper()} {day}\n"
        "========================================================================\n"
        "Generated at: 2026-11-20 12:00 SAST\n"
        "PAPER ONLY · eligibility labels are evidence states, not profit guarantees.\n\n"
        "ELIGIBLE CLEAN PICKS\n"
        "========================================================================\n"
        f"{body}"
    ).encode()


def marker(state: str, day: str, detailed: bytes) -> bytes:
    return (
        "Ma Golide compact pick lifecycle marker\n"
        f"league_state={state}\n"
        f"target_date={day}\n"
        "full_forecast_board=Actions artifact and GitHub Release archive\n"
        f"detailed_card_sha256={hashlib.sha256(detailed).hexdigest()}\n"
    ).encode()


def write_marker(repo: Path, state: str, day: str, detailed: bytes) -> Path:
    path = repo / "localdata" / state / f"picks_{day}.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(marker(state, day, detailed))
    return path


def artifact_file(root: Path, state: str, relative: str, raw: bytes) -> Path:
    path = root / "successful-lane" / "localdata" / state / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(raw)
    return path


def git(repo: Path, *args: str) -> str:
    return subprocess.check_output(["git", "-C", str(repo), *args], text=True).strip()


def init_git_snapshot(repo: Path, state: str, day: str, raw: bytes) -> str:
    subprocess.run(["git", "init", "-q", str(repo)], check=True)
    git(repo, "config", "user.name", "Test User")
    git(repo, "config", "user.email", "test@example.invalid")
    path = repo / "localdata" / state / f"picks_{day}.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(raw)
    git(repo, "add", "localdata")
    git(repo, "commit", "-qm", "historical readable pick card")
    return git(repo, "rev-parse", "HEAD")


def test_restore_from_git_snapshot_is_dry_run_then_atomic_apply(tmp_path):
    repo = tmp_path / "repo"
    day = "2026-11-20"
    historical = readable_card("nba", day, "  historical row\n")
    snapshot = init_git_snapshot(repo, "nba", day, historical)
    target = write_marker(repo, "nba", day, b"a later artifact revision")

    preview = restore_pick_cards(repo=repo, snapshot_ref=snapshot)

    assert preview["status"] == "complete"
    assert preview["dry_run"] is True
    assert preview["restored_count"] == 1
    assert preview["restored"][0]["source"] == "git_snapshot"
    assert target.read_bytes().startswith(b"Ma Golide compact pick lifecycle marker\n")

    applied = restore_pick_cards(repo=repo, snapshot_ref=snapshot, apply=True)

    assert applied["dry_run"] is False
    assert applied["snapshot_commit"] == snapshot
    assert target.read_bytes() == historical


def test_restore_reports_unavailable_marker_without_overwriting_it(tmp_path):
    repo = tmp_path / "repo"
    day = "2026-11-20"
    target = write_marker(repo, "nba", day, readable_card("nba", day))

    report = restore_pick_cards(repo=repo, apply=True)

    assert report["status"] == "incomplete"
    assert report["restored"] == []
    assert report["unavailable"] == [
        {
            "path": "localdata/nba/picks_2026-11-20.txt",
            "detailed_card_sha256": hashlib.sha256(readable_card("nba", day)).hexdigest(),
        }
    ]
    assert target.read_bytes().startswith(b"Ma Golide compact pick lifecycle marker\n")


def test_restore_prefers_marker_hash_verified_release_archive(tmp_path):
    repo = tmp_path / "repo"
    artifacts = tmp_path / "artifacts"
    day = "2026-11-20"
    detailed = readable_card("nba", day, "  exact archived pick\n")
    target = write_marker(repo, "nba", day, detailed)
    artifact_file(artifacts, "nba", "csv/games.csv.gz", b"canonical")
    artifact_file(artifacts, "nba", f"picks_{day}.txt", detailed)
    output = tmp_path / "archives"
    build_release_archives(artifacts=artifacts, repo=repo, output=output)

    report = restore_pick_cards(
        repo=repo,
        archives=[output / "magolide-state-nba.tar.gz"],
        apply=True,
    )

    assert report["status"] == "complete"
    assert report["restored_count"] == 1
    assert report["restored"][0]["source"] == "release_archive"
    assert report["restored"][0]["sha256"] == hashlib.sha256(detailed).hexdigest()
    assert target.read_bytes() == detailed


def test_restore_never_substitutes_a_nonmatching_release_archive_card(tmp_path):
    repo = tmp_path / "repo"
    artifacts = tmp_path / "artifacts"
    day = "2026-11-20"
    expected = readable_card("nba", day, "  marker revision\n")
    archived = readable_card("nba", day, "  other revision\n")
    target = write_marker(repo, "nba", day, expected)
    artifact_file(artifacts, "nba", "csv/games.csv.gz", b"canonical")
    artifact_file(artifacts, "nba", f"picks_{day}.txt", archived)
    output = tmp_path / "archives"
    build_release_archives(artifacts=artifacts, repo=repo, output=output)

    report = restore_pick_cards(
        repo=repo,
        archives=[output / "magolide-state-nba.tar.gz"],
        apply=True,
    )

    assert report["status"] == "incomplete"
    assert report["restored"] == []
    assert report["unavailable_count"] == 1
    assert target.read_bytes().startswith(b"Ma Golide compact pick lifecycle marker\n")


def test_restore_rejects_archive_when_git_manifest_checksum_does_not_match(tmp_path):
    repo = tmp_path / "repo"
    artifacts = tmp_path / "artifacts"
    day = "2026-11-20"
    detailed = readable_card("nba", day)
    write_marker(repo, "nba", day, detailed)
    artifact_file(artifacts, "nba", "csv/games.csv.gz", b"canonical")
    artifact_file(artifacts, "nba", f"picks_{day}.txt", detailed)
    output = tmp_path / "archives"
    build_release_archives(artifacts=artifacts, repo=repo, output=output)
    archive = output / "magolide-state-nba.tar.gz"
    archive.write_bytes(archive.read_bytes() + b"corruption")

    with pytest.raises(ValueError, match="archive manifest identity or checksum mismatch"):
        restore_pick_cards(repo=repo, archives=[archive])
