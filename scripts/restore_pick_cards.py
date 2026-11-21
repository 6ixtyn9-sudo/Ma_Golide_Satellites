#!/usr/bin/env python3
"""Restore readable legacy pick cards from verified retained evidence.

From 2026-11-20 to 2026-11-21 the publisher wrote compact lifecycle markers
instead of the normal human-readable ``picks_YYYY-MM-DD.txt`` card.  Those
markers identify the detailed card's digest, but do not embed its text.

This migration is deliberately narrow:

* it considers only registered ``localdata/<state>/picks_YYYY-MM-DD.txt``
  marker files;
* it never alters a card that is already readable;
* a release-archive candidate must match both its signed-in-Git archive
  manifest and the marker's detailed-card SHA-256; and
* a selected historical Git snapshot is an explicitly labelled fallback for a
  pre-compaction card.  It is useful where a marker predates an archive, but
  does not pretend to be the marker's later artifact revision.

The command defaults to a dry run.  Use ``--apply`` to atomically replace only
validated markers.  If no retained source is available, the marker remains in
place and is reported as unavailable; ``--require-complete`` turns that state
into a non-zero exit for an operator-controlled migration.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tarfile
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.league_registry import SPECS  # noqa: E402
from scripts.release_archive import (  # noqa: E402
    ARCHIVE_SCHEMA,
    CONTENT_MANIFEST_NAME,
    _safe_tar_member,
    _sha256_path,
)

STATE_KEYS = frozenset(spec.state_key for spec in SPECS)
DATE = r"\d{4}-\d{2}-\d{2}"
PICK_NAME = re.compile(rf"picks_({DATE})\.txt$")
MARKER_HEADER = "Ma Golide compact pick lifecycle marker"
MARKER_ARCHIVE_POINTER = "Actions artifact and GitHub Release archive"
READABLE_CARD_PREFIX = "Ma Golide Picks — "
MAX_CARD_BYTES = 1_000_000
DEFAULT_MAX_TOTAL_BYTES = 5_000_000


@dataclass(frozen=True)
class CompactMarker:
    path: Path
    relative: str
    state_key: str
    target_date: str
    detailed_card_sha256: str


@dataclass(frozen=True)
class CardCandidate:
    raw: bytes
    source: str
    source_id: str


def _atomic_write(path: Path, raw: bytes) -> None:
    temporary = path.with_name(path.name + ".restore-tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, path)


def _load_json(path: Path, *, description: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid {description}: {path}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"invalid {description}: {path}")
    return value


def _parse_marker(path: Path, *, repo: Path) -> CompactMarker | None:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError(f"pick card is not UTF-8: {path}") from exc
    if not text.startswith(MARKER_HEADER + "\n"):
        return None

    lines = text.splitlines()
    if len(lines) != 5 or lines[0] != MARKER_HEADER:
        raise ValueError(f"invalid compact pick lifecycle marker: {path}")
    pairs = [line.split("=", 1) for line in lines[1:]]
    if any(len(pair) != 2 for pair in pairs):
        raise ValueError(f"invalid compact pick lifecycle marker: {path}")
    fields = dict(pairs)
    expected_fields = {
        "league_state",
        "target_date",
        "full_forecast_board",
        "detailed_card_sha256",
    }
    if len(fields) != len(pairs) or set(fields) != expected_fields:
        raise ValueError(f"invalid compact pick lifecycle marker: {path}")

    match = PICK_NAME.fullmatch(path.name)
    state_key = path.parent.name
    relative = path.relative_to(repo).as_posix()
    if (
        match is None
        or state_key not in STATE_KEYS
        or fields["league_state"] != state_key
        or fields["target_date"] != match.group(1)
        or fields["full_forecast_board"] != MARKER_ARCHIVE_POINTER
        or re.fullmatch(r"[0-9a-f]{64}", fields["detailed_card_sha256"]) is None
    ):
        raise ValueError(f"compact pick lifecycle marker identity mismatch: {path}")
    return CompactMarker(
        path=path,
        relative=relative,
        state_key=state_key,
        target_date=fields["target_date"],
        detailed_card_sha256=fields["detailed_card_sha256"],
    )


def find_compact_markers(repo: Path) -> list[CompactMarker]:
    """Find and structurally validate compact marker cards under ``localdata``."""

    localdata = repo / "localdata"
    if not localdata.is_dir():
        return []
    markers: list[CompactMarker] = []
    for path in sorted(localdata.glob("*/picks_*.txt")):
        if path.is_symlink() or not path.is_file() or path.parent.name not in STATE_KEYS:
            continue
        marker = _parse_marker(path, repo=repo)
        if marker is not None:
            markers.append(marker)
    return markers


def _validated_readable_card(raw: bytes, *, description: str, max_card_bytes: int) -> bytes:
    if not raw or len(raw) > max_card_bytes:
        raise ValueError(f"invalid readable pick card size: {description}")
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError(f"readable pick card is not UTF-8: {description}") from exc
    if not text.startswith(READABLE_CARD_PREFIX):
        raise ValueError(f"source is not a readable Ma Golide pick card: {description}")
    return raw


def _git_output(repo: Path, *args: str) -> bytes:
    result = subprocess.run(
        ["git", "-C", str(repo), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise ValueError(f"git evidence lookup failed: {' '.join(args)}: {detail}")
    return result.stdout


def resolve_git_snapshot(repo: Path, source_ref: str) -> str:
    """Resolve a source reference to one immutable commit ID."""

    return _git_output(repo, "rev-parse", "--verify", f"{source_ref}^{{commit}}").decode().strip()


def _snapshot_candidate(
    repo: Path, *, snapshot_commit: str, marker: CompactMarker, max_card_bytes: int
) -> CardCandidate | None:
    probe = subprocess.run(
        ["git", "-C", str(repo), "cat-file", "-e", f"{snapshot_commit}:{marker.relative}"],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if probe.returncode:
        return None
    raw = _git_output(repo, "show", f"{snapshot_commit}:{marker.relative}")
    _validated_readable_card(raw, description=f"{snapshot_commit}:{marker.relative}", max_card_bytes=max_card_bytes)
    return CardCandidate(raw=raw, source="git_snapshot", source_id=snapshot_commit)


def _archive_state_from_name(name: str) -> str | None:
    match = re.fullmatch(r"magolide-state-([a-z0-9_]+)\.tar\.gz", name)
    return match.group(1) if match and match.group(1) in STATE_KEYS else None


def _verified_archive_cards(
    repo: Path, archive: Path, *, max_card_bytes: int
) -> dict[tuple[str, str], CardCandidate]:
    """Read verified detailed artifact cards from one local checkpoint asset.

    The release asset is trusted only after its Git-published archive manifest,
    its content-manifest digest, every member path, and every member digest have
    been checked.  Only ``artifact/<state>/picks_DATE.txt`` entries are exposed;
    operator copies and canonical payloads are never candidates here.
    """

    state_key = _archive_state_from_name(archive.name)
    if state_key is None:
        raise ValueError(f"archive has no registered state-key filename: {archive}")
    expected_path = repo / "localdata" / state_key / "archive_manifest.json"
    expected = _load_json(expected_path, description="release archive manifest")
    if (
        expected.get("schema_version") != ARCHIVE_SCHEMA
        or expected.get("state_key") != state_key
        or expected.get("asset_name") != archive.name
        or expected.get("archive_sha256") != _sha256_path(archive)
        or not isinstance(expected.get("content_manifest_sha256"), str)
    ):
        raise ValueError(f"archive manifest identity or checksum mismatch: {archive}")

    try:
        with tarfile.open(archive, mode="r:gz") as bundle:
            members = bundle.getmembers()
            by_name = {member.name: member for member in members}
            if len(by_name) != len(members) or CONTENT_MANIFEST_NAME not in by_name:
                raise ValueError(f"archive members are incomplete or duplicated: {archive}")
            manifest_handle = bundle.extractfile(by_name[CONTENT_MANIFEST_NAME])
            if manifest_handle is None:
                raise ValueError(f"archive has unreadable content manifest: {archive}")
            manifest_bytes = manifest_handle.read()
            if hashlib.sha256(manifest_bytes).hexdigest() != expected["content_manifest_sha256"]:
                raise ValueError(f"archive content manifest checksum mismatch: {archive}")
            manifest = json.loads(manifest_bytes)
            if (
                not isinstance(manifest, dict)
                or manifest.get("schema_version") != ARCHIVE_SCHEMA
                or manifest.get("state_key") != state_key
                or not isinstance(manifest.get("entries"), list)
            ):
                raise ValueError(f"archive content manifest identity mismatch: {archive}")
            entries = manifest["entries"]
            expected_entries = {
                entry.get("path"): entry
                for entry in entries
                if isinstance(entry, dict) and isinstance(entry.get("path"), str)
            }
            if len(expected_entries) != len(entries) or set(by_name) != {
                CONTENT_MANIFEST_NAME,
                *expected_entries,
            }:
                raise ValueError(f"archive member manifest mismatch: {archive}")

            cards: dict[tuple[str, str], CardCandidate] = {}
            prefix = f"artifact/{state_key}/"
            for name, entry in expected_entries.items():
                member = by_name[name]
                if not _safe_tar_member(name) or not member.isfile():
                    raise ValueError(f"archive contains unsafe member: {name}")
                handle = bundle.extractfile(member)
                if handle is None:
                    raise ValueError(f"archive member cannot be read: {name}")
                raw = handle.read()
                if (
                    len(raw) != entry.get("bytes")
                    or hashlib.sha256(raw).hexdigest() != entry.get("sha256")
                ):
                    raise ValueError(f"archive member checksum mismatch: {name}")
                relative = name.removeprefix(prefix)
                match = PICK_NAME.fullmatch(relative) if name.startswith(prefix) else None
                if match is None:
                    continue
                if entry.get("role") != "run_evidence":
                    raise ValueError(f"archive pick-card role mismatch: {name}")
                _validated_readable_card(raw, description=f"{archive}:{name}", max_card_bytes=max_card_bytes)
                key = (state_key, match.group(1))
                if key in cards:
                    raise ValueError(f"archive has duplicate readable pick card: {name}")
                cards[key] = CardCandidate(raw=raw, source="release_archive", source_id=archive.name)
    except (OSError, tarfile.TarError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid release archive: {archive}") from exc
    return cards


def _archive_candidates(
    repo: Path, archives: Iterable[Path], *, max_card_bytes: int
) -> dict[tuple[str, str], CardCandidate]:
    candidates: dict[tuple[str, str], CardCandidate] = {}
    for archive in archives:
        for key, candidate in _verified_archive_cards(repo, archive, max_card_bytes=max_card_bytes).items():
            previous = candidates.get(key)
            if previous is not None and previous.raw != candidate.raw:
                raise ValueError(f"conflicting verified archive pick cards for {key[0]} {key[1]}")
            candidates[key] = candidate
    return candidates


def restore_pick_cards(
    *,
    repo: Path,
    snapshot_ref: str | None = None,
    archives: Iterable[Path] = (),
    apply: bool = False,
    max_card_bytes: int = MAX_CARD_BYTES,
    max_total_bytes: int = DEFAULT_MAX_TOTAL_BYTES,
) -> dict[str, Any]:
    """Restore markers only when a bounded, validated source is retained."""

    repo = repo.resolve()
    if max_card_bytes <= 0 or max_total_bytes <= 0:
        raise ValueError("card-size limits must be positive")
    markers = find_compact_markers(repo)
    snapshot_commit = resolve_git_snapshot(repo, snapshot_ref) if snapshot_ref else None
    archives = tuple(path.resolve() for path in archives)
    archive_cards = _archive_candidates(repo, archives, max_card_bytes=max_card_bytes)

    restored: list[dict[str, str]] = []
    unavailable: list[dict[str, str]] = []
    total_bytes = 0
    planned: list[tuple[CompactMarker, CardCandidate]] = []
    for marker in markers:
        candidate = archive_cards.get((marker.state_key, marker.target_date))
        if candidate is not None:
            actual_sha = hashlib.sha256(candidate.raw).hexdigest()
            if actual_sha != marker.detailed_card_sha256:
                # An archive is only an exact marker recovery source.  Do not
                # silently substitute a different run's card.
                candidate = None
        if candidate is None and snapshot_commit is not None:
            candidate = _snapshot_candidate(
                repo,
                snapshot_commit=snapshot_commit,
                marker=marker,
                max_card_bytes=max_card_bytes,
            )
        if candidate is None:
            unavailable.append(
                {
                    "path": marker.relative,
                    "detailed_card_sha256": marker.detailed_card_sha256,
                }
            )
            continue
        total_bytes += len(candidate.raw)
        if total_bytes > max_total_bytes:
            raise ValueError(
                f"pick-card restore exceeds {max_total_bytes} byte limit while processing {marker.relative}"
            )
        planned.append((marker, candidate))

    if apply:
        for marker, candidate in planned:
            _atomic_write(marker.path, candidate.raw)
    for marker, candidate in planned:
        restored.append(
            {
                "path": marker.relative,
                "source": candidate.source,
                "source_id": candidate.source_id,
                "sha256": hashlib.sha256(candidate.raw).hexdigest(),
            }
        )
    return {
        "status": "complete" if not unavailable else "incomplete",
        "dry_run": not apply,
        "snapshot_commit": snapshot_commit,
        "archive_assets": [archive.name for archive in archives],
        "marker_count": len(markers),
        "restored": restored,
        "restored_count": len(restored),
        "restored_bytes": total_bytes,
        "unavailable": unavailable,
        "unavailable_count": len(unavailable),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    parser.add_argument(
        "--snapshot-ref",
        help="commit/ref containing a historical readable-card snapshot; this is a labelled fallback",
    )
    parser.add_argument(
        "--archive",
        type=Path,
        action="append",
        default=[],
        help="verified release archive asset, named magolide-state-<state>.tar.gz (repeatable)",
    )
    parser.add_argument("--apply", action="store_true", help="atomically write validated cards")
    parser.add_argument(
        "--require-complete",
        action="store_true",
        help="return a non-zero status if any compact markers have no retained source",
    )
    parser.add_argument("--max-card-bytes", type=int, default=MAX_CARD_BYTES)
    parser.add_argument("--max-total-bytes", type=int, default=DEFAULT_MAX_TOTAL_BYTES)
    args = parser.parse_args()
    report = restore_pick_cards(
        repo=args.repo,
        snapshot_ref=args.snapshot_ref,
        archives=args.archive,
        apply=args.apply,
        max_card_bytes=args.max_card_bytes,
        max_total_bytes=args.max_total_bytes,
    )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 2 if args.require_complete and report["unavailable"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
