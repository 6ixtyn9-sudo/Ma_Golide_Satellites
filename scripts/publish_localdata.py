#!/usr/bin/env python3
"""Merge downloaded league operator reports into Git-tracked ``localdata``.

The workflow's league matrix writes isolated artifact directories. A single
publisher calls this module after all matrix jobs finish, then performs one Git
commit/push. Only derived pick, paper-ticket, and daily coverage-gap marker
files are eligible. Canonical DuckDB/CSV state, acquisition ledgers, caches,
and credentials are never copied.

Existing repository files not present in the latest artifact are retained. A
conflicting immutable ticket or inline per-fixture snapshot aborts publication
instead of rotating/reopening previously frozen evidence. Inline no-bet date
locks in durable ticket state reject stale cached ticket receipts. Git-state
withdrawal IDs are rejected when stale caches attempt to return a deliberately
withdrawn receipt.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
import sys
from typing import Any, Dict, Iterable

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.league_registry import SPECS  # noqa: E402

# Publishing must follow the same registry that creates state directories. A
# new supported league cannot become an invisible artifact-only lane because a
# second hand-maintained allow-list forgot its state key.
STATE_KEYS = frozenset(spec.state_key for spec in SPECS)
DATE = r"\d{4}-\d{2}-\d{2}"
ABA_HISTORY_RECEIPTS_FILE = "history_receipts.json"
ABA_HISTORY_RECEIPTS_SCHEMA = "magolide.aba_history_receipts"
_ABA_HISTORY_RECEIPT_FIELDS = frozenset(
    {
        "season_end_year",
        "calendar_final_count",
        "reconciled_final_count",
        "administrative_disposition_count",
        "not_applicable_rows",
        "blocking_gap_count",
        "status",
        "source",
        "evidence",
        "completed_at_utc",
    }
)
_ABA_HISTORY_RECEIPT_IDENTITY_FIELDS = _ABA_HISTORY_RECEIPT_FIELDS - {"completed_at_utc"}
_ABA_LEGACY_HISTORY_RECEIPT_FIELDS = _ABA_HISTORY_RECEIPT_FIELDS - {
    "administrative_disposition_count"
}
LIFECYCLE_RECOVERY_ARTIFACT_PREFIX = "MaGolide-lifecycle-"
LIFECYCLE_RECOVERY_NAMES = tuple(
    re.compile(pattern)
    for pattern in (
        rf"picks_{DATE}\.txt",
        r"pick_lifecycle_state\.json",
        # A time-locked EuroLeague research roster is non-financial but still
        # immutable selection evidence, so an interrupted successful freeze
        # must be recoverable without allowing mutable cards/tickets through.
        rf"euroleague_consensus_paper_{DATE}_roster\.json",
    )
)
ALLOWED_NAMES = tuple(
    re.compile(pattern)
    for pattern in (
        # Dated text cards and lifecycle state (with inline immutable fixture
        # snapshots) are Git-durable operator evidence. Structured dated pick
        # JSON and detailed lifecycle renders remain cache/Actions-artifact-only
        # for every league.
        rf"picks_{DATE}\.txt",
        r"pick_lifecycle_state\.json",
        rf"coverage_gap_{DATE}\.json",
        r"auto_tickets_state\.json",
        r"auto_tickets_performance\.json",
        r"auto_tickets_performance\.txt",
        rf"auto_tickets_{DATE}\.txt",
        rf"auto_tickets_{DATE}_decision\.json",
        rf"auto_tickets_{DATE}_withdrawn\.json",
        # Compact ABA season-count attestations are the only non-ticket
        # evidence summary admitted to Git. Raw source pages remain excluded.
        r"history_receipts\.json",
        # This card is mutable only for official-result grades. The companion
        # roster below is the immutable, no-stake selection snapshot.
        rf"euroleague_consensus_paper_{DATE}\.txt",
        rf"euroleague_consensus_paper_{DATE}_roster\.json",
    )
)
IMMUTABLE_NAMES = re.compile(
    rf"(?:auto_tickets_{DATE}_decision|"
    rf"euroleague_consensus_paper_{DATE}_roster)\.json"
)
DECISION_RECEIPT_NAME = re.compile(rf"auto_tickets_({DATE})_decision\.json")
CALIBRATION_EPOCH = "2026-11-10"
# A later publisher job can be based on an older workflow SHA than a manual
# settlement already merged to main. Never let that stale artifact put a
# rendered pending grade back over Git-durable settled public evidence.
PENDING_GRADE_RE = re.compile(r"(?:grade=|\[\s*)PENDING\b|⏳\s+PENDING\b")
SETTLED_GRADE_RE = re.compile(
    r"(?:grade=|\[\s*)?(?:✅\s+WIN|❌\s+LOSS|➖\s+PUSH|⚪\s+VOID)\b|lifecycle=SETTLED"
)
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_TOTAL_BYTES = 100 * 1024 * 1024


def allowed_operator_file(name: str) -> bool:
    """Return whether a root-level filename is safe to publish to Git."""
    return any(pattern.fullmatch(name) for pattern in ALLOWED_NAMES)


def allowed_lifecycle_recovery_file(name: str) -> bool:
    """Return whether an interrupted-lane artifact may contain ``name``.

    A matrix lane that fails after freezing picks may preserve the durable
    lifecycle needed to replay that fixture. It must never smuggle a mutable
    card, compact ticket summary, or ticket decision into publication.
    """

    return any(pattern.fullmatch(name) for pattern in LIFECYCLE_RECOVERY_NAMES)


def _atomic_write(path: Path, raw: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".publish-tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, path)


def _validated_aba_history_receipt(receipt: object, *, description: str) -> dict[str, Any]:
    # Receipts written before generic administrative dispositions existed had
    # only reconciled played finals. Preserve those immutable attestations as a
    # zero-administrative-count form rather than stranding prior Git evidence.
    if isinstance(receipt, dict) and set(receipt) == _ABA_LEGACY_HISTORY_RECEIPT_FIELDS:
        if receipt.get("status") != "calendar_final_count_reconciled":
            raise ValueError(f"invalid ABA history receipt: {description}")
        receipt = {
            **receipt,
            "administrative_disposition_count": 0,
            "status": "calendar_terminal_count_reconciled",
        }
    if not isinstance(receipt, dict) or set(receipt) != _ABA_HISTORY_RECEIPT_FIELDS:
        raise ValueError(f"invalid ABA history receipt: {description}")
    season = receipt["season_end_year"]
    counts = (
        receipt["calendar_final_count"],
        receipt["reconciled_final_count"],
        receipt["administrative_disposition_count"],
        receipt["not_applicable_rows"],
        receipt["blocking_gap_count"],
    )
    if (
        isinstance(season, bool)
        or not isinstance(season, int)
        or any(isinstance(value, bool) or not isinstance(value, int) or value < 0 for value in counts)
        or receipt["calendar_final_count"]
        != receipt["reconciled_final_count"] + receipt["administrative_disposition_count"]
        or receipt["blocking_gap_count"] != 0
        or receipt["status"] != "calendar_terminal_count_reconciled"
        or receipt["source"] != "aba_liga"
        or not isinstance(receipt["evidence"], str)
        or not receipt["evidence"]
        or not isinstance(receipt["completed_at_utc"], str)
        or not receipt["completed_at_utc"]
    ):
        raise ValueError(f"invalid ABA history receipt: {description}")
    return dict(receipt)


def _load_aba_history_receipts(raw: bytes, *, description: str) -> dict[str, dict[str, Any]]:
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid ABA history receipt summary: {description}") from exc
    if (
        not isinstance(data, dict)
        or data.get("schema_version") != ABA_HISTORY_RECEIPTS_SCHEMA
        or data.get("league") != "ABA_LEAGUE"
        or data.get("provider") != "aba_liga"
        or not isinstance(data.get("receipts"), dict)
    ):
        raise ValueError(f"invalid ABA history receipt summary: {description}")
    receipts: dict[str, dict[str, Any]] = {}
    for season, receipt in data["receipts"].items():
        if not isinstance(season, str) or not season.isdigit():
            raise ValueError(f"invalid ABA history receipt season: {description}")
        validated = _validated_aba_history_receipt(receipt, description=f"{description}:{season}")
        if validated["season_end_year"] != int(season):
            raise ValueError(f"ABA history receipt season mismatch: {description}:{season}")
        receipts[season] = validated
    return receipts


def _merge_aba_history_receipts(destination: Path, incoming_raw: bytes, *, relative: str) -> bytes:
    """Union immutable historical ABA count attestations without raw evidence.

    A repeated receipt with the same count semantics is idempotent. A changed
    historical count must be reviewed rather than letting a stale artifact
    overwrite Git-durable evidence silently.
    """

    incoming = _load_aba_history_receipts(incoming_raw, description=relative)
    existing = (
        _load_aba_history_receipts(destination.read_bytes(), description=str(destination))
        if destination.exists()
        else {}
    )
    merged = dict(existing)
    for season, receipt in incoming.items():
        previous = merged.get(season)
        if previous is None:
            merged[season] = receipt
            continue
        if any(previous[field] != receipt[field] for field in _ABA_HISTORY_RECEIPT_IDENTITY_FIELDS):
            raise ValueError(f"ABA history receipt conflicts with durable evidence: {relative}:{season}")
    payload = {
        "schema_version": ABA_HISTORY_RECEIPTS_SCHEMA,
        "league": "ABA_LEAGUE",
        "provider": "aba_liga",
        "receipts": {key: merged[key] for key in sorted(merged, key=int)},
    }
    return (json.dumps(payload, indent=2, sort_keys=True) + "\n").encode()


def _git_durable_bytes(source: Path) -> bytes:
    """Return the compact, durable form of one publishable operator file.

    Dated text cards and inline immutable fixture snapshots in lifecycle state
    are Git durable. Structured dated pick JSON and detailed lifecycle renders
    remain artifact-only. The two compact ticket summaries
    below still contain scan clocks that are operational telemetry,
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


def _pick_lifecycle_state(raw: bytes, *, description: str) -> dict[str, Any]:
    """Validate the durable inline-fixture state before publishing it."""

    try:
        state = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid pick lifecycle state: {description}") from exc
    if (
        not isinstance(state, dict)
        or state.get("schema_version") != "magolide.pick_lifecycle_state"
        or not isinstance(state.get("league"), str)
        or not isinstance(state.get("pending_frozen_games"), list)
        or not isinstance(state.get("dated_reports"), dict)
    ):
        raise ValueError(f"invalid pick lifecycle state: {description}")
    frozen_by_date = state.setdefault("frozen_by_date", {})
    if not isinstance(frozen_by_date, dict):
        raise ValueError(f"invalid pick lifecycle state locks: {description}")
    for target_date, by_game in frozen_by_date.items():
        if not isinstance(target_date, str) or not re.fullmatch(DATE, target_date):
            raise ValueError(f"invalid pick lifecycle state date lock: {description}")
        if not isinstance(by_game, dict):
            raise ValueError(f"invalid pick lifecycle state fixture locks: {description}")
        for game_id, freeze in by_game.items():
            if (
                not isinstance(freeze, dict)
                or str(game_id) != str(freeze.get("game_id") or "")
                or freeze.get("schema_version") != "magolide.pick_freeze"
                or freeze.get("target_date") != target_date
                or freeze.get("league") != state["league"]
                or not isinstance(freeze.get("rows"), list)
            ):
                raise ValueError(f"invalid inline pick freeze: {description}")
    return state


def _pick_lifecycle_state_bytes(state: dict[str, Any]) -> bytes:
    return (json.dumps(state, indent=2, sort_keys=True, default=str) + "\n").encode()


def _merge_pick_lifecycle_state(
    destination: Path, incoming_raw: bytes, *, relative: str
) -> bytes:
    """Merge immutable inline fixture snapshots without accepting rotations.

    The mutable pending/report fields are taken from the incoming completed
    lane. Snapshots missing from a stale artifact remain Git-authoritative, and
    their existing pending entry is retained until a current lane regrades it.
    """

    incoming = _pick_lifecycle_state(incoming_raw, description=relative)
    if not destination.exists():
        return _pick_lifecycle_state_bytes(incoming)
    existing = _pick_lifecycle_state(destination.read_bytes(), description=str(destination))
    if existing["league"] != incoming["league"]:
        raise ValueError(f"pick lifecycle state league differs from repository evidence: {relative}")

    merged = dict(incoming)
    incoming_locks = incoming["frozen_by_date"]
    merged_locks = {
        date: dict(by_game) for date, by_game in existing["frozen_by_date"].items()
    }
    for target_date, by_game in incoming_locks.items():
        destination_locks = merged_locks.setdefault(target_date, {})
        for game_id, freeze in by_game.items():
            previous = destination_locks.get(game_id)
            if previous is not None and previous != freeze:
                raise ValueError(
                    "immutable inline pick freeze differs from repository evidence: "
                    f"{relative}:{target_date}/{game_id}"
                )
            destination_locks[game_id] = freeze
    merged["frozen_by_date"] = merged_locks

    incoming_game_ids = {
        str(game_id) for by_game in incoming_locks.values() for game_id in by_game
    }
    pending = {
        (str(item[0]), str(item[1]))
        for item in incoming["pending_frozen_games"]
        if isinstance(item, (list, tuple)) and len(item) == 2
    }
    for item in existing["pending_frozen_games"]:
        if not isinstance(item, (list, tuple)) or len(item) != 2:
            raise ValueError(f"invalid pick lifecycle pending fixture: {destination}")
        if str(item[0]) not in incoming_game_ids:
            pending.add((str(item[0]), str(item[1])))
    merged["pending_frozen_games"] = [
        list(item) for item in sorted(pending, key=lambda item: (item[1], item[0]))
    ]
    for target_date, by_game in merged_locks.items():
        report = merged["dated_reports"].get(target_date)
        if not isinstance(report, dict):
            report = {}
            merged["dated_reports"][target_date] = report
        report["frozen_fixtures"] = max(int(report.get("frozen_fixtures") or 0), len(by_game))
    return _pick_lifecycle_state_bytes(merged)


def _would_regress_settled_card(destination: Path, artifact_bytes: bytes) -> bool:
    """Return whether a stale artifact would re-open settled public evidence.

    Text cards are mutable lifecycle views, but an authoritative settled grade
    is never allowed to become ``PENDING`` merely because an older matrix job
    reached its publisher after a newer main commit. An empty/current draft
    still updates normally: this guard applies only when the repository view
    actually contains rendered settled evidence.
    """

    if not destination.name.endswith(".txt"):
        return False
    try:
        current = destination.read_text(encoding="utf-8")
        incoming = artifact_bytes.decode("utf-8")
    except UnicodeDecodeError:
        # Text-card publication must not silently reinterpret non-UTF-8 data;
        # the regular write path will preserve existing validation behavior.
        return False
    return (
        SETTLED_GRADE_RE.search(current) is not None
        and PENDING_GRADE_RE.search(incoming) is not None
    )


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
    return frozenset(dates)


def _artifact_league_dirs(
    artifacts: Path, *, skipped_states: list[str]
) -> Iterable[tuple[str, Path, str]]:
    """Yield recognised ``artifact/localdata/<state>`` directories once.

    Unknown state directories are never published, but they are reported to
    make registry/publisher drift reviewable instead of silently discarding an
    otherwise valid artifact.
    """

    if not artifacts.exists():
        return
    for artifact in sorted(artifacts.iterdir()):
        localdata = artifact / "localdata"
        if not localdata.is_dir():
            # A capability-skipped artifact can contain only daily-output.json.
            continue
        for state_dir in sorted(localdata.iterdir()):
            if not state_dir.is_dir():
                continue
            if state_dir.name not in STATE_KEYS:
                skipped_states.append((artifact / "localdata" / state_dir.name).as_posix())
                continue
            yield state_dir.name, state_dir, artifact.name


def merge_downloaded_localdata(artifacts: Path, repo: Path) -> Dict[str, Any]:
    """Overlay eligible recovered files without deleting repository history."""
    artifacts = artifacts.resolve()
    repo = repo.resolve()
    copied = []
    unchanged = []
    immutable_verified = []
    withdrawn_skipped = []
    no_bet_receipts_skipped = []
    skipped_states: list[str] = []
    recovery_files_skipped = []
    settled_card_regressions_skipped = []
    total_bytes = 0
    seen_destinations = set()
    for state, source_dir, artifact_name in _artifact_league_dirs(
        artifacts, skipped_states=skipped_states
    ):
        destination_dir = repo / "localdata" / state
        lifecycle_recovery = artifact_name.startswith(LIFECYCLE_RECOVERY_ARTIFACT_PREFIX)
        sources = sorted(source_dir.iterdir())
        withdrawn_decision_ids = _withdrawn_decision_ids(destination_dir)
        no_bet_dates = _no_bet_decision_dates(destination_dir) | _no_bet_decision_dates(
            source_dir
        )
        for source in sources:
            if lifecycle_recovery and not allowed_lifecycle_recovery_file(source.name):
                recovery_files_skipped.append((source_dir / source.name).as_posix())
                continue
            if not allowed_operator_file(source.name) or (
                source.name == ABA_HISTORY_RECEIPTS_FILE and state != "aba_league"
            ):
                continue
            # Every dated auto-ticket text card is a mutable public view: it
            # shows the single dated card (including its pre-cutoff draft
            # section), an immutable-ticket settlement view, or the terminal
            # no-bet state. _ticket_receipt alone creates a state lock and
            # immutable decision JSON at the 60-minute pre-tip cutoff. The
            # retired *_preview.txt sidecar is not allowlisted.
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
            if receipt_match and receipt_match.group(1) in no_bet_dates:
                # An older cache must never resurrect a ticket after this date
                # was deliberately closed at its immutable cutoff with no bet.
                no_bet_receipts_skipped.append(relative)
                continue
            # Hard-reset: never re-publish pre-epoch dated files (even from a warm older cache).
            m = re.search(r"(\d{4}-\d{2}-\d{2})", source.name)
            if m and m.group(1) < CALIBRATION_EPOCH:
                withdrawn_skipped.append(relative)
                continue
            if relative in seen_destinations:
                raise ValueError(f"duplicate recovered operator destination: {relative}")
            seen_destinations.add(relative)
            raw = _git_durable_bytes(source)
            if source.name == "pick_lifecycle_state.json":
                raw = _merge_pick_lifecycle_state(destination, raw, relative=relative)
            elif source.name == ABA_HISTORY_RECEIPTS_FILE:
                raw = _merge_aba_history_receipts(destination, raw, relative=relative)
            if DECISION_RECEIPT_NAME.fullmatch(source.name) and withdrawn_decision_ids:
                try:
                    candidate = json.loads(raw)
                except json.JSONDecodeError as exc:
                    raise ValueError(f"invalid ticket receipt JSON: {source}") from exc
                if str(candidate.get("decision_id")) in withdrawn_decision_ids:
                    withdrawn_skipped.append(relative)
                    continue
            if destination.exists():
                previous = destination.read_bytes()
                if _would_regress_settled_card(destination, raw):
                    settled_card_regressions_skipped.append(relative)
                    continue
                if IMMUTABLE_NAMES.fullmatch(source.name) and previous != raw:
                    kind = (
                        "immutable ticket receipt"
                        if DECISION_RECEIPT_NAME.fullmatch(source.name)
                        else "immutable EuroLeague consensus-paper roster"
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
        "skipped_states": sorted(skipped_states),
        "recovery_files_skipped": sorted(recovery_files_skipped),
        "settled_card_regressions_skipped": sorted(settled_card_regressions_skipped),
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
