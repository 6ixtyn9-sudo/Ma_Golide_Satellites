#!/usr/bin/env python3
"""Merge downloaded league operator reports into Git-tracked ``localdata``.

The workflow's league matrix writes isolated artifact directories. A single
publisher calls this module after all matrix jobs finish, then performs one Git
commit/push. Only derived pick and paper-ticket files are eligible. Canonical
DuckDB/CSV state, acquisition ledgers, caches, and credentials are never copied.

Existing repository files not present in the latest artifact are retained. A
conflicting immutable decision receipt aborts publication instead of rotating
or rewriting a previously frozen ticket.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
from typing import Any, Dict, Iterable

STATE_KEYS = frozenset({"nba", "euroleague", "cba", "ncaa_m", "ncaa_w"})
DATE = r"\d{4}-\d{2}-\d{2}"
ALLOWED_NAMES = tuple(
    re.compile(pattern)
    for pattern in (
        r"picks_today\.json",
        rf"picks_{DATE}\.json",
        rf"picks_{DATE}\.txt",
        r"picks_audit_rolling\.json",
        rf"picks_audit_{DATE}\.md",
        r"auto_tickets_state\.json",
        r"auto_tickets_performance\.json",
        r"auto_tickets_performance\.txt",
        rf"auto_tickets_{DATE}\.txt",
        rf"auto_tickets_{DATE}_decision\.json",
    )
)
IMMUTABLE_NAMES = re.compile(rf"auto_tickets_{DATE}_decision\.json")
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_TOTAL_BYTES = 100 * 1024 * 1024


def allowed_operator_file(name: str) -> bool:
    """Return whether a root-level filename is safe to publish to Git."""
    return any(pattern.fullmatch(name) for pattern in ALLOWED_NAMES)


def _atomic_write(path: Path, raw: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".publish-tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, path)


def _artifact_league_dirs(artifacts: Path) -> Iterable[tuple[str, Path]]:
    """Yield each recovered ``artifact/localdata/<state>`` directory once."""
    if not artifacts.exists():
        return
    for artifact in sorted(artifacts.iterdir()):
        localdata = artifact / "localdata"
        if not localdata.is_dir():
            # CBA's credential-skipped artifact contains only daily-output.json.
            continue
        for state_dir in sorted(localdata.iterdir()):
            if state_dir.name not in STATE_KEYS or not state_dir.is_dir():
                continue
            yield state_dir.name, state_dir


def merge_downloaded_localdata(artifacts: Path, repo: Path) -> Dict[str, Any]:
    """Overlay eligible recovered files without deleting repository history."""
    artifacts = artifacts.resolve()
    repo = repo.resolve()
    copied = []
    unchanged = []
    immutable_verified = []
    total_bytes = 0
    seen_destinations = set()
    for state, source_dir in _artifact_league_dirs(artifacts):
        for source in sorted(source_dir.iterdir()):
            if not allowed_operator_file(source.name):
                continue
            if source.is_symlink() or not source.is_file():
                raise ValueError(f"operator publication requires a regular file: {source}")
            size = source.stat().st_size
            if size > MAX_FILE_BYTES:
                raise ValueError(f"operator file exceeds {MAX_FILE_BYTES} bytes: {source}")
            total_bytes += size
            if total_bytes > MAX_TOTAL_BYTES:
                raise ValueError("operator artifact batch exceeds the 100 MiB publication limit")
            destination = repo / "localdata" / state / source.name
            relative = destination.relative_to(repo).as_posix()
            if relative in seen_destinations:
                raise ValueError(f"duplicate recovered operator destination: {relative}")
            seen_destinations.add(relative)
            raw = source.read_bytes()
            if destination.exists():
                previous = destination.read_bytes()
                if IMMUTABLE_NAMES.fullmatch(source.name) and previous != raw:
                    raise ValueError(
                        f"immutable ticket receipt differs from repository evidence: {relative}"
                    )
                if previous == raw:
                    unchanged.append(relative)
                    if IMMUTABLE_NAMES.fullmatch(source.name):
                        immutable_verified.append(relative)
                    continue
            _atomic_write(destination, raw)
            copied.append(relative)
    return {
        "status": "merged" if copied else "no_changes",
        "copied": copied,
        "unchanged": unchanged,
        "immutable_receipts_verified": immutable_verified,
        "copied_count": len(copied),
        "unchanged_count": len(unchanged),
        "source_bytes_checked": total_bytes,
        "destination": str(repo / "localdata"),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--artifacts", type=Path, required=True)
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    args = parser.parse_args()
    report = merge_downloaded_localdata(args.artifacts, args.repo)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
