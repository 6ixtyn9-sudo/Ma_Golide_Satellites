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
from collections import Counter
import hashlib
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
# The temporary 2026-11-20/21 compact marker is opaque presentation debt. A
# stale Actions artifact must never republish it over a readable Git card.
COMPACT_PICK_CARD_MARKER = b"Ma Golide compact pick lifecycle marker\n"
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
    )
)
IMMUTABLE_NAMES = re.compile(
    rf"auto_tickets_{DATE}_decision\.json"
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
_CARD_GRADE_TOKEN_RE = re.compile(
    r"lifecycle=SETTLED(?:_[A-Z_]+)?|"
    r"(?:grade=)?(?:[✅❌➖⚪⏳]\s+)?(?:PENDING|WIN|LOSS|PUSH|VOID)\b"
)
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_TOTAL_BYTES = 100 * 1024 * 1024


def allowed_operator_file(name: str) -> bool:
    """Return whether a root-level filename is safe to publish to Git."""
    return any(pattern.fullmatch(name) for pattern in ALLOWED_NAMES)


def is_compact_pick_card_marker(raw: bytes, *, name: str) -> bool:
    """Return whether an artifact carries the retired opaque pick-card marker."""

    return bool(re.fullmatch(rf"picks_{DATE}\.txt", name)) and raw.startswith(
        COMPACT_PICK_CARD_MARKER
    )


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


COMPACT_FORECAST_EVIDENCE_SCHEMA = "magolide.frozen_forecast_evidence.v1"


def _canonical_hash(value: object) -> str:
    """Return a content hash for an artifact-only detailed evidence record."""

    encoded = json.dumps(value, sort_keys=True, separators=(",", ":"), default=str).encode()
    return hashlib.sha256(encoded).hexdigest()


def _subset(value: object, fields: tuple[str, ...]) -> dict[str, Any]:
    return {field: value[field] for field in fields if isinstance(value, dict) and field in value}


def _compact_roi_band(value: object) -> dict[str, Any] | None:
    if not isinstance(value, dict):
        return None
    return _subset(
        value,
        (
            "market",
            "period",
            "probability_band",
            "individual_odds_band",
            "historical_pick_count",
            "wins",
            "losses",
            "pushes",
            "observed_roi",
            "adjusted_roi",
            "recency_weighted_roi",
            "qualifies",
        ),
    )


def _compact_frozen_forecast(row: object) -> object:
    """Reduce a frozen forecast to settlement and decision-critical evidence.

    The full model explanation, historical features, alternative probabilities,
    and broad candidate-band grids remain in the bounded run artifact and the
    release archive. Git retains the fields needed to prove a frozen selection,
    its price/source status, its decision gates, and later official grading.
    """

    if not isinstance(row, dict):
        return row
    if row.get("evidence_schema") == COMPACT_FORECAST_EVIDENCE_SCHEMA:
        return row
    required = {
        "game_id",
        "league",
        "scheduled_utc",
        "home",
        "away",
        "market",
        "period",
        "selection",
        "raw_probability",
    }
    if not required <= set(row):
        # Keep deliberately minimal legacy/test fixtures unchanged. Production
        # rows have the complete canonical identity above.
        return row

    compact: dict[str, Any] = {
        "evidence_schema": COMPACT_FORECAST_EVIDENCE_SCHEMA,
        "detailed_evidence_sha256": _canonical_hash(row),
        **_subset(
            row,
            (
                "forecast_id",
                "pick_id",
                "game_id",
                "league",
                "date",
                "scheduled_utc",
                "home",
                "away",
                "market",
                "period",
                "selection",
                "line",
                "raw_probability",
                "calibrated_probability",
                "push_probability",
                "model",
                "model_version",
                "calibration_version",
                "role",
                "bucket",
                "price_evidence",
                "paper_only",
                "paper_roi_band_eligible",
                "outcome",
                "grade_status",
            ),
        ),
    }
    compact["book"] = _subset(
        row.get("book"),
        (
            "provider",
            "book",
            "observed_utc",
            "source_modified_utc",
            "pre_match",
            "main_line",
            "frozen",
            "selected_decimal_price",
            "selected_opposite_decimal_price",
            "price_provenance",
            "source_permission",
            "operational_validation",
            "game_id",
            "market",
            "period",
            "line",
        ),
    )
    if not compact["book"]:
        compact.pop("book")
    compact["prediction"] = _subset(
        row.get("prediction"),
        ("provider", "observed_utc", "frozen", "model", "model_version", "produced_by"),
    )
    if not compact["prediction"]:
        compact.pop("prediction")
    compact["calibration"] = _subset(
        row.get("calibration"),
        ("status", "n", "minimum_samples", "sufficient_for_accas"),
    )
    compact["eligibility"] = _subset(
        row.get("eligibility"),
        ("acca_eligible", "reasons", "status"),
    )
    band = _compact_roi_band(row.get("paper_roi_band"))
    if band:
        compact["paper_roi_band"] = band
    return compact


def _compact_pick_filter(value: object, legs: object) -> dict[str, Any] | None:
    if not isinstance(value, dict):
        return None
    compact = _subset(
        value,
        (
            "status",
            "candidate_pick_count",
            "qualifying_pick_count",
            "qualifying_band_count",
            "historical_unique_pick_count",
            "history_source",
            "history_fallback",
        ),
    )
    selected_ids = {
        str(leg.get("forecast_id"))
        for leg in (legs if isinstance(legs, list) else [])
        if isinstance(leg, dict) and leg.get("forecast_id")
    }
    detailed = value.get("candidate_band_details")
    if isinstance(detailed, dict) and selected_ids:
        compact_details = {
            forecast_id: band
            for forecast_id, raw_band in detailed.items()
            if forecast_id in selected_ids and (band := _compact_roi_band(raw_band))
        }
        if compact_details:
            compact["candidate_band_details"] = compact_details
    return compact or None


def _compact_ticket_receipt(data: dict[str, Any]) -> dict[str, Any]:
    """Compact immutable ticket receipts without changing ticket semantics."""

    schema = str(data.get("schema_version") or "")
    if re.sub(r"\.v\d+$", "", schema) != "magolide.paper_ticket_receipt":
        return data
    compact = dict(data)
    for collection in ("accas", "singles"):
        positions = compact.get(collection)
        if not isinstance(positions, list):
            continue
        compact_positions = []
        for position in positions:
            if not isinstance(position, dict):
                compact_positions.append(position)
                continue
            reduced = dict(position)
            legs = position.get("legs")
            if isinstance(legs, list):
                reduced["legs"] = [_compact_frozen_forecast(leg) for leg in legs]
            pick_filter = _compact_pick_filter(position.get("individual_pick_filter"), legs)
            if pick_filter is None:
                reduced.pop("individual_pick_filter", None)
            else:
                reduced["individual_pick_filter"] = pick_filter
            compact_positions.append(reduced)
        compact[collection] = compact_positions
    return compact


def _git_durable_bytes_from_raw(raw: bytes, *, name: str, source: Path) -> bytes:
    """Return the Git-durable form for a publisher-allowlisted file."""

    if name not in {
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
    } and not DECISION_RECEIPT_NAME.fullmatch(name):
        return raw
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        # Preserve existing publisher compatibility for a syntactically
        # opaque immutable receipt; its byte-for-byte conflict guard still
        # applies. Real receipts are structured JSON and are compacted below.
        if DECISION_RECEIPT_NAME.fullmatch(name):
            return raw
        raise ValueError(f"invalid operator JSON: {source}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"operator JSON must be an object: {source}")
    if name == "auto_tickets_state.json":
        data.pop("scans", None)
        data.pop("last_generated_utc", None)
    elif name == "auto_tickets_performance.json":
        data.pop("generated_utc", None)
    else:
        data = _compact_ticket_receipt(data)
    return (json.dumps(data, indent=2, sort_keys=True, default=str) + "\n").encode()


def _git_durable_bytes(source: Path) -> bytes:
    """Return the compact, durable form of one publishable operator file."""

    return _git_durable_bytes_from_raw(source.read_bytes(), name=source.name, source=source)


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


def _compact_pick_lifecycle_state(state: dict[str, Any]) -> dict[str, Any]:
    """Keep inline immutable fixture evidence while removing repeated model bulk."""

    compact = dict(state)
    frozen_by_date = state.get("frozen_by_date") or {}
    compact_locks: dict[str, dict[str, Any]] = {}
    for target_date, by_game in frozen_by_date.items():
        compact_locks[target_date] = {}
        for game_id, freeze in by_game.items():
            reduced = dict(freeze)
            rows = freeze.get("rows") if isinstance(freeze, dict) else None
            if isinstance(rows, list):
                reduced["rows"] = [_compact_frozen_forecast(row) for row in rows]
            compact_locks[target_date][game_id] = reduced
    compact["frozen_by_date"] = compact_locks
    return compact


def _pick_lifecycle_state_bytes(state: dict[str, Any]) -> bytes:
    return (json.dumps(state, indent=2, sort_keys=True, default=str) + "\n").encode()


def _merge_no_bet_releases(destination: Path, incoming_raw: bytes, *, relative: str) -> bytes:
    """Keep a deliberate release from being undone by an artifact cut before it.

    A lane restores Git state when it starts, so a run cut before a release
    publishes a state file that still carries the released no-bet and knows
    nothing about the release. On 2026-11-25 two such runs put the same freeze
    back, twice. Releases union forward; the released decision is dropped from
    the incoming copy only when it is the very decision that was released, so a
    later, legitimately different freeze for that date still publishes.
    """

    try:
        incoming = json.loads(incoming_raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid ticket state JSON: {relative}") from exc
    if not isinstance(incoming, dict):
        raise ValueError(f"ticket state must be an object: {relative}")
    if not destination.exists():
        return incoming_raw
    existing = json.loads(destination.read_bytes())
    releases = existing.get("no_bet_releases") or {}
    if not isinstance(releases, dict) or not releases:
        return incoming_raw
    merged = dict(incoming)
    merged_releases = dict(merged.get("no_bet_releases") or {})
    merged_releases.update(releases)
    frozen = dict(merged.get("frozen_by_date") or {})
    for target_date, release in merged_releases.items():
        released = (release or {}).get("released_lock") or {}
        decision_id = released.get("decision_id")
        lock = frozen.get(target_date)
        if isinstance(lock, dict) and decision_id and lock.get("decision_id") == decision_id:
            # The freeze this run is publishing is the one the operator released,
            # not a new one. It keeps its place in the release record below.
            frozen.pop(target_date, None)
    merged["frozen_by_date"] = frozen
    merged["no_bet_releases"] = merged_releases
    return (json.dumps(merged, indent=2, sort_keys=True, default=str) + "\n").encode()


def _merge_pick_lifecycle_state(
    destination: Path, incoming_raw: bytes, *, relative: str
) -> bytes:
    """Merge immutable inline fixture snapshots without accepting rotations.

    The mutable pending/report fields are taken from the incoming completed
    lane. Snapshots missing from a stale artifact remain Git-authoritative, and
    their existing pending entry is retained until a current lane regrades it.
    """

    incoming = _compact_pick_lifecycle_state(
        _pick_lifecycle_state(incoming_raw, description=relative)
    )
    if not destination.exists():
        return _pick_lifecycle_state_bytes(incoming)
    existing = _compact_pick_lifecycle_state(
        _pick_lifecycle_state(destination.read_bytes(), description=str(destination))
    )
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


DATED_CARD_NAMES = re.compile(rf"(?:picks|auto_tickets)_{DATE}\.txt")


def _card_grade_entries(text: str) -> list[tuple[str, str]]:
    """Return stable line fingerprints for unambiguous rendered grade states.

    A dated card can legitimately contain both settled early games and a late
    pending game.  The publisher must therefore compare the state of each
    rendered line rather than treating *any* pending badge as a whole-card
    rollback.  Summary lines are harmless extra entries; exact line matching
    plus the settled-count fallback below keeps the guard conservative for
    older, less structured card formats.
    """

    entries: list[tuple[str, str]] = []
    for line in text.splitlines():
        is_pending = PENDING_GRADE_RE.search(line) is not None
        is_settled = SETTLED_GRADE_RE.search(line) is not None
        if is_pending == is_settled:
            # Ignore ungraded text and deliberately ambiguous aggregate lines.
            continue
        fingerprint = _CARD_GRADE_TOKEN_RE.sub("", line)
        fingerprint = re.sub(r"[\[\]]", "", fingerprint)
        fingerprint = " ".join(fingerprint.split())
        if fingerprint:
            entries.append(("pending" if is_pending else "settled", fingerprint))
    return entries


def _would_regress_settled_card(destination: Path, artifact_bytes: bytes) -> bool:
    """Return whether an artifact explicitly re-opens settled public evidence.

    A settled line can never become ``PENDING`` because an older matrix job
    reaches the publisher after newer main state.  In contrast, a card with a
    late live game must be able to publish its newer, progressively settled
    lines while that one game remains pending.  A count fallback preserves the
    legacy no-identity guard for simple cards that cannot be line-matched.
    """

    # Only dated cards are progressive settlement evidence. The rolling
    # performance summary legitimately loses settled lines when held legs
    # finish (the HELD block disappears), so the line-count fallback below
    # would otherwise reject every corrected summary forever.
    if not DATED_CARD_NAMES.fullmatch(destination.name):
        return False
    try:
        current = destination.read_text(encoding="utf-8")
        incoming = artifact_bytes.decode("utf-8")
    except UnicodeDecodeError:
        # Text-card publication must not silently reinterpret non-UTF-8 data;
        # the regular write path will preserve existing validation behavior.
        return False

    current_entries = _card_grade_entries(current)
    incoming_entries = _card_grade_entries(incoming)
    settled = Counter(fingerprint for state, fingerprint in current_entries if state == "settled")
    if not settled:
        return False
    incoming_pending = Counter(
        fingerprint for state, fingerprint in incoming_entries if state == "pending"
    )
    # Same rendered subject, explicitly downgraded to pending: always reject.
    if settled.keys() & incoming_pending.keys():
        return True
    incoming_settled = sum(state == "settled" for state, _fingerprint in incoming_entries)
    # Older cards do not always expose a stable per-row subject.  Never accept
    # fewer rendered settled entries than the repository already has.
    return incoming_settled < sum(settled.values())


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


def compact_repository_localdata(repo: Path) -> Dict[str, Any]:
    """Compact mutable state summaries without altering published pick cards.

    The publisher calls this only after the workflow has uploaded the same
    detailed state to a verified release checkpoint. It is idempotent and
    never removes an inline fixture, ticket selection, source-permission flag,
    named price, immutable receipt identity, or readable dated pick card.
    """

    repo = repo.resolve()
    changed: list[str] = []
    saved_bytes = 0
    for state_dir in sorted((repo / "localdata").glob("*")):
        if not state_dir.is_dir() or state_dir.name not in STATE_KEYS:
            continue
        for path in sorted(state_dir.iterdir()):
            if path.is_symlink() or not path.is_file() or not allowed_operator_file(path.name):
                continue
            if path.name == ABA_HISTORY_RECEIPTS_FILE and state_dir.name != "aba_league":
                continue
            previous = path.read_bytes()
            if path.name == "pick_lifecycle_state.json":
                compact = _compact_pick_lifecycle_state(
                    _pick_lifecycle_state(previous, description=str(path))
                )
                raw = _pick_lifecycle_state_bytes(compact)
            else:
                raw = _git_durable_bytes_from_raw(previous, name=path.name, source=path)
            if raw == previous:
                continue
            _atomic_write(path, raw)
            changed.append(path.relative_to(repo).as_posix())
            saved_bytes += len(previous) - len(raw)
    return {"changed": changed, "changed_count": len(changed), "bytes_saved": saved_bytes}


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
    compact_pick_markers_skipped = []
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
            source_raw = source.read_bytes()
            # A marker contains no rows—only a digest/pointer—so it cannot be
            # a legitimate update to a readable report. This guard also blocks
            # an in-flight pre-fix worker from undoing a Git-card restoration.
            if is_compact_pick_card_marker(source_raw, name=source.name):
                compact_pick_markers_skipped.append(relative)
                continue
            raw = _git_durable_bytes_from_raw(source_raw, name=source.name, source=source)
            if source.name == "auto_tickets_state.json":
                raw = _merge_no_bet_releases(destination, raw, relative=relative)
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
                if _would_regress_settled_card(destination, source_raw):
                    settled_card_regressions_skipped.append(relative)
                    continue
                if IMMUTABLE_NAMES.fullmatch(source.name):
                    # The first compact-state publication can replace an older
                    # detailed receipt only when both reduce to identical
                    # selection/price/settlement evidence. Any factual change
                    # remains a hard conflict.
                    durable_previous = _git_durable_bytes_from_raw(
                        previous, name=source.name, source=destination
                    )
                    if durable_previous != raw:
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
        "compact_pick_markers_skipped": sorted(compact_pick_markers_skipped),
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
    parser.add_argument(
        "--compact-existing",
        action="store_true",
        help="compact existing Git-visible state after a verified archive upload",
    )
    args = parser.parse_args()
    report = merge_downloaded_localdata(args.artifacts, args.repo)
    if args.compact_existing:
        report["repository_compaction"] = compact_repository_localdata(args.repo)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
