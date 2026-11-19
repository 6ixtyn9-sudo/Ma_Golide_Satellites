#!/usr/bin/env python3
"""Merge downloaded league operator reports into Git-tracked ``localdata``.

The workflow's league matrix writes isolated artifact directories. A single
publisher calls this module after all matrix jobs finish, then performs one Git
commit/push. Only derived pick and paper-ticket files are eligible. Canonical
DuckDB/CSV state, acquisition ledgers, caches, and credentials are never copied.

Existing repository files not present in the latest artifact are retained. A
conflicting immutable ticket, no-bet decision, or per-fixture pick freeze aborts
publication instead of rotating/reopening previously frozen evidence. Git-state
withdrawal IDs are rejected when stale caches attempt to return a deliberately
withdrawn receipt.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
from typing import Any, Dict, Iterable

STATE_KEYS = frozenset(
    {"nba", "wnba", "euroleague", "euroleague_w", "cba", "ncaa_m", "ncaa_w"}
)
DATE = r"\d{4}-\d{2}-\d{2}"
ALLOWED_NAMES = tuple(
    re.compile(pattern)
    for pattern in (
        # Dated pick cards are the operator-visible draft/freeze/settlement
        # lifecycle. They are deliberately distinct from volatile picks_today
        # and the expanded audit ledgers, which remain artifact-only.
        rf"picks_{DATE}\.json",
        rf"picks_{DATE}\.txt",
        rf"picks_{DATE}_lifecycle\.json",
        r"pick_lifecycle_state\.json",
        rf"pick_freeze_{DATE}_[0-9a-f]{{16}}\.json",
        r"auto_tickets_state\.json",
        r"auto_tickets_performance\.json",
        r"auto_tickets_performance\.txt",
        rf"auto_tickets_{DATE}\.txt",
        rf"auto_tickets_{DATE}_decision\.json",
        rf"auto_tickets_{DATE}_no_bet\.json",
        rf"auto_tickets_{DATE}_withdrawn\.json",
    )
)
IMMUTABLE_NAMES = re.compile(
    rf"(?:auto_tickets_{DATE}_(?:decision|no_bet)|pick_freeze_{DATE}_[0-9a-f]{{16}})\.json"
)
DECISION_RECEIPT_NAME = re.compile(rf"auto_tickets_({DATE})_decision\.json")
NO_BET_DECISION_NAME = re.compile(rf"auto_tickets_({DATE})_no_bet\.json")
TICKET_CARD_NAME = re.compile(rf"auto_tickets_({DATE})\.txt")
CALIBRATION_EPOCH = "2026-11-10"
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


def _git_durable_bytes(source: Path) -> bytes:
    """Return the compact, durable form of one publishable operator file.

    Dated draft/freeze lifecycle cards are Git durable. The two compact ticket
    summaries below still contain scan clocks that are operational telemetry,
    not evidence needed to reconstruct tickets. Removing only those volatile
    fields prevents no-op runs from producing new Git blobs while receipts and
    date locks remain complete.
    """
    raw = source.read_bytes()
    if source.name not in {
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
    }:
        return raw
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid operator JSON: {source}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"operator JSON must be an object: {source}")
    if source.name == "auto_tickets_state.json":
        data.pop("scans", None)
        data.pop("last_generated_utc", None)
    else:
        data.pop("generated_utc", None)
    return (json.dumps(data, indent=2, sort_keys=True, default=str) + "\n").encode()


def _withdrawn_decision_ids(localdata: Path) -> frozenset[str]:
    """Load Git-authoritative withdrawals before accepting cached receipts."""
    ids: set[str] = set()
    path = localdata / "auto_tickets_state.json"
    if path.is_file():
        try:
            state = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(f"invalid paper-ticket state: {path}") from exc
        locks = state.get("frozen_by_date", {})
        if not isinstance(locks, dict):
            raise ValueError(f"invalid paper-ticket locks: {path}")
        ids.update(
            str(lock["decision_id"])
            for lock in locks.values()
            if isinstance(lock, dict) and lock.get("withdrawn") is True and lock.get("decision_id")
        )
    # Sidecars survive state rollbacks.
    for sidecar in localdata.glob("auto_tickets_*_withdrawn.json"):
        try:
            data = json.loads(sidecar.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        did = data.get("decision_id")
        if did:
            ids.add(str(did))
    return frozenset(ids)


def _no_bet_decision_dates(localdata: Path) -> frozenset[str]:
    """Load Git-authoritative terminal no-bet decisions before cache merge."""
    dates: set[str] = set()
    path = localdata / "auto_tickets_state.json"
    if path.is_file():
        try:
            state = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(f"invalid paper-ticket state: {path}") from exc
        locks = state.get("frozen_by_date", {})
        if not isinstance(locks, dict):
            raise ValueError(f"invalid paper-ticket locks: {path}")
        dates.update(
            str(date)
            for date, lock in locks.items()
            if isinstance(lock, dict) and lock.get("no_bet") is True
        )
    for sidecar in localdata.glob("auto_tickets_*_no_bet.json"):
        match = NO_BET_DECISION_NAME.fullmatch(sidecar.name)
        if match:
            dates.add(match.group(1))
    return frozenset(dates)


def _decision_receipt_dates(paths: Iterable[Path]) -> frozenset[str]:
    """Return dates with an immutable ticket receipt, never a live scan card."""
    return frozenset(
        match.group(1)
        for path in paths
        if (match := DECISION_RECEIPT_NAME.fullmatch(path.name)) is not None
    )


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
    withdrawn_skipped = []
    no_bet_receipts_skipped = []
    total_bytes = 0
    seen_destinations = set()
    for state, source_dir in _artifact_league_dirs(artifacts):
        destination_dir = repo / "localdata" / state
        sources = sorted(source_dir.iterdir())
        withdrawn_decision_ids = _withdrawn_decision_ids(destination_dir)
        no_bet_dates = _no_bet_decision_dates(destination_dir) | _no_bet_decision_dates(
            source_dir
        )
        receipt_dates = _decision_receipt_dates(
            destination_dir.glob("auto_tickets_*_decision.json")
        ) | _decision_receipt_dates(sources)
        for source in sources:
            if not allowed_operator_file(source.name):
                continue
            ticket_card_match = TICKET_CARD_NAME.fullmatch(source.name)
            if ticket_card_match:
                if (
                    ticket_card_match.group(1) not in receipt_dates
                    or ticket_card_match.group(1) in no_bet_dates
                ):
                    # Candidate/no-bet cards contain a latest-scan timestamp
                    # and price movement.  A card tied to an immutable receipt
                    # is instead the mutable, human-visible settlement view:
                    # its selections/prices remain receipt-backed while grades
                    # may advance from pending to settled on a later run.
                    continue
            if source.is_symlink() or not source.is_file():
                raise ValueError(f"operator publication requires a regular file: {source}")
            size = source.stat().st_size
            if size > MAX_FILE_BYTES:
                raise ValueError(f"operator file exceeds {MAX_FILE_BYTES} bytes: {source}")
            total_bytes += size
            if total_bytes > MAX_TOTAL_BYTES:
                raise ValueError("operator artifact batch exceeds the 100 MiB publication limit")
            destination = destination_dir / source.name
            relative = destination.relative_to(repo).as_posix()
            receipt_match = DECISION_RECEIPT_NAME.fullmatch(source.name)
            no_bet_match = NO_BET_DECISION_NAME.fullmatch(source.name)
            if receipt_match and receipt_match.group(1) in no_bet_dates:
                # An older cache must never resurrect a ticket after this date
                # was deliberately closed at its immutable cutoff with no bet.
                no_bet_receipts_skipped.append(relative)
                continue
            if no_bet_match and (destination_dir / f"auto_tickets_{no_bet_match.group(1)}_decision.json").exists():
                raise ValueError(
                    "cannot publish a no-bet decision beside an existing immutable "
                    f"ticket receipt: {relative}"
                )
            # Hard-reset: never re-publish pre-epoch dated files (even from a warm older cache).
            m = re.search(r"(\d{4}-\d{2}-\d{2})", source.name)
            if m and m.group(1) < CALIBRATION_EPOCH:
                withdrawn_skipped.append(relative)
                continue
            if relative in seen_destinations:
                raise ValueError(f"duplicate recovered operator destination: {relative}")
            seen_destinations.add(relative)
            raw = _git_durable_bytes(source)
            if IMMUTABLE_NAMES.fullmatch(source.name) and withdrawn_decision_ids:
                try:
                    candidate = json.loads(raw)
                except json.JSONDecodeError as exc:
                    raise ValueError(f"invalid ticket receipt JSON: {source}") from exc
                if str(candidate.get("decision_id")) in withdrawn_decision_ids:
                    withdrawn_skipped.append(relative)
                    continue
            if destination.exists():
                previous = destination.read_bytes()
                if IMMUTABLE_NAMES.fullmatch(source.name) and previous != raw:
                    kind = (
                        "immutable ticket receipt"
                        if DECISION_RECEIPT_NAME.fullmatch(source.name)
                        else "immutable pick freeze"
                    )
                    raise ValueError(f"{kind} differs from repository evidence: {relative}")
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
        "withdrawn_receipts_skipped": withdrawn_skipped,
        "no_bet_receipts_skipped": no_bet_receipts_skipped,
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
