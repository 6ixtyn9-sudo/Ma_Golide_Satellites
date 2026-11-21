"""Edge-style operator pick cards, paper tickets, settlement, and audits.

This module is a presentation/settlement layer only. It does not select a
winner, alter a probability, or place a bet. Current forecasts come from
``MarketForecastService``; outcomes come only from reconciled official finals.

Operator evidence files live beneath one league's localdata directory:
- ``picks_YYYY-MM-DD.json`` and ``.txt`` are Git-published draft/lifecycle
  cards. They show all forecast picks, but are never evidence that a pick was
  actionable before its own cutoff.
- ``pick_lifecycle_state.json`` embeds immutable per-fixture snapshots by date
  and fixture once the fixture enters its 60-minute pre-tip window.
- ``picks_YYYY-MM-DD_lifecycle.json`` is a mutable rendered view that grades
  only those immutable snapshots after official finals arrive.
- ``picks_today.json`` and rolling/dated audit files retain full active-cache
  forecast evidence and workflow-artifact detail.
- ``auto_tickets_YYYY-MM-DD_decision.json`` is an immutable paper receipt.
- ``auto_tickets_YYYY-MM-DD.txt`` is a mutable, Git-published settlement card
  rendering a frozen decision and its latest available grades. Selections,
  prices, and stakes remain receipt-backed; only settlement/audit presentation
  may change after publication.
- ``auto_tickets_state.json`` stores write-once date locks and bounded scan telemetry.
- ``auto_tickets_performance.json`` / ``.txt`` reconstruct paper performance.
- ``euroleague_consensus_paper_YYYY-MM-DD_roster.json`` is a time-locked,
  no-stake EuroLeague aggregate-consensus research board; its matching ``.txt``
  card renders grades without creating a ticket, receipt, bank, ROI, or P&L.

A known losing leg settles its paper position as a loss even when other legs
are still pending. Push legs are removed from an acca's effective price once
every other leg has resolved; a single settles at its selected decimal price.
Pending tickets never count as wins or losses. Empty scans never cancel or
rotate a frozen receipt. Every position has a flat one-paper-unit stake.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime, time, timedelta, timezone
import hashlib
import itertools
import json
import math
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

from ingestion.identity import parse_utc
from ingestion.models import (
    Game,
    GameStatus,
    League,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
    schema_matches,
)

from .types import (
    BookEvidence,
    Forecast,
    _period_scores,
    current_spelling,
    grade_forecast,
    opposite_selection,
)

OPERATOR_SCHEMA = "magolide.operator"
FADE_AUDIT_SCHEMA = "magolide.paper_fade_audit"
TICKET_SCHEMA = "magolide.paper_ticket_receipt"
NO_BET_DECISION_SCHEMA = "magolide.paper_ticket_no_bet"
STATE_SCHEMA = "magolide.paper_ticket_state"
PICK_FREEZE_SCHEMA = "magolide.pick_freeze"
PICK_LIFECYCLE_SCHEMA = "magolide.pick_lifecycle"
PICK_LIFECYCLE_STATE_SCHEMA = "magolide.pick_lifecycle_state"
UNPRICED_AUDIT_SCHEMA = "magolide.unpriced_forecast_audit"
UNPRICED_SUCCESS_RATE_THRESHOLD = 0.70
UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE = 10
UNPRICED_SUCCESS_RATE_HISTORY_DAYS = 30
UNPRICED_SUCCESS_RATE_RESULT_DELAY_HOURS = 4.0
SHADOW_GROUP_MAX_PICKS = 3
UNPRICED_SHADOW_GROUP_MAX_PICKS = SHADOW_GROUP_MAX_PICKS
PRICED_INELIGIBLE_SHADOW_MIN_DECIMAL_ODDS = 1.30
PRICED_INELIGIBLE_SHADOW_MIN_PROBABILITY = 0.70
PRICED_INELIGIBLE_SHADOW_SCHEMA = "magolide.priced_ineligible_shadow"
UNPRICED_SPREAD_SHADOW_SCHEMA = "magolide.unpriced_spread_shadow"
# EuroLeague's public BetExplorer adapter exposes a representative aggregate
# price, not an executable named-book quote.  It therefore has its own
# deliberately isolated research board: the board preserves the exact observed
# consensus price, but can never feed tickets, receipts, stakes, bank, ROI, or
# P&L.  These quality gates are intentionally stricter than a simple watchlist.
EUROLEAGUE_CONSENSUS_PAPER_SCHEMA = "magolide.euroleague_consensus_paper_roster"
EUROLEAGUE_CONSENSUS_PROVIDER = "betexplorer"
EUROLEAGUE_CONSENSUS_BOOK = "average"
EUROLEAGUE_CONSENSUS_MAX_PICKS = 3
EUROLEAGUE_CONSENSUS_MIN_CALIBRATION_SAMPLES = 10
EUROLEAGUE_CONSENSUS_MIN_MODEL_EDGE = 0.03
EUROLEAGUE_CONSENSUS_ALLOWED_BLOCKERS = frozenset(
    {
        "price_not_named_bookmaker_quote",
        "insufficient_walk_forward_samples",
    }
)
# No-price full-game moneyline picks at or above this model probability are
# re-expressed as the same side against the bookmaker's posted full-game spread.
UNPRICED_SPREAD_SHADOW_MIN_PROBABILITY = 0.70
# A derived spread is allowed into the future-only paper-ticket fallback only
# after its own pre-tip shadow history has cleared the same minimum number of
# settled decisions and cover-rate floor as the no-price audit.  It never
# becomes an after-the-fact ticket for a previously completed fixture.
UNPRICED_SPREAD_TICKET_MIN_DECISIONS_EXCLUSIVE = 10
UNPRICED_SPREAD_TICKET_MIN_COVER_RATE = 0.70
BASE_PAPER_UNITS = 100.0
STAKE_UNITS_PER_ACCA = 1.0
# Keep this compact NBA exception aligned with ingestion.adaptive_gate: NBA
# cards publish at 18:00 SAST so the North-American board can mature; other
# leagues retain the established 09:00 window. An earlier fixture still wins.
TICKET_DECISION_LEAD_TIME = timedelta(hours=1)
TICKET_PUBLICATION_TIMEZONE = "Africa/Johannesburg"
DEFAULT_TICKET_PUBLICATION_LOCAL_TIME = time(hour=9)
NBA_TICKET_PUBLICATION_LOCAL_TIME = time(hour=18)
DEFAULT_TICKET_DECISION_POLICY = "0900_sast_or_first_slate_tip_minus_60_minutes"
NBA_TICKET_DECISION_POLICY = "1800_sast_or_first_slate_tip_minus_60_minutes"


def _ticket_publication_local_time(league: League) -> time:
    return (
        NBA_TICKET_PUBLICATION_LOCAL_TIME
        if league is League.NBA
        else DEFAULT_TICKET_PUBLICATION_LOCAL_TIME
    )


def _ticket_decision_policy(league: League) -> str:
    return NBA_TICKET_DECISION_POLICY if league is League.NBA else DEFAULT_TICKET_DECISION_POLICY
BUCKET_ORDER = [
    "ELIGIBLE_CLEAN",
    "CAUTION",
    "WATCHLIST_NO_ODDS",
    "WATCHLIST_INSUFFICIENT_CONTEXT",
    "SKIPPED_INELIGIBLE",
]
BUCKET_LABELS = {
    "ELIGIBLE_CLEAN": "ELIGIBLE CLEAN PICKS",
    "CAUTION": "CAUTION PICKS",
    "WATCHLIST_NO_ODDS": "WATCHLIST — NO MATCHED ODDS",
    "WATCHLIST_INSUFFICIENT_CONTEXT": "WATCHLIST — INSUFFICIENT CONTEXT",
    "SKIPPED_INELIGIBLE": "SKIPPED — INELIGIBLE",
}
UTC = timezone.utc


def _grade_badge(status: Any) -> str:
    """Use consistent, accessible text markers for every rendered grade."""
    normalized = str(status or "unknown").lower()
    icon, label = {
        "win": ("✅", "WIN"),
        "loss": ("❌", "LOSS"),
        "push": ("➖", "PUSH"),
        "pending": ("⏳", "PENDING"),
        "void": ("⚪", "VOID"),
    }.get(normalized, ("❔", normalized.upper()))
    return f"{icon} {label}"


def _grade_count_summary(
    counts: Dict[str, int],
    *,
    count_first: bool = False,
    suffixes: Optional[Dict[str, str]] = None,
    separator: str = " / ",
) -> str:
    """Render only grade categories with positive counts, in stable grade order."""
    normalized_counts: Counter = Counter()
    for status, count in counts.items():
        normalized_counts[str(status or "unknown").lower()] += int(count)
    order = ["win", "loss", "push", "pending", "void"]
    order.extend(sorted(status for status in normalized_counts if status not in order))
    parts = []
    for status in order:
        count = normalized_counts[status]
        if count <= 0:
            continue
        badge = _grade_badge(status)
        if count_first:
            parts.append(f"{count} {badge}")
        else:
            suffix = (suffixes or {}).get(status, "")
            parts.append(f"{badge} {count}{suffix}")
    return separator.join(parts)


def _atomic_bytes(path: Path, raw: bytes) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_bytes() == raw:
        return str(path)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, path)
    return str(path)


def _write_json(path: Path, value: Any) -> str:
    raw = (json.dumps(value, indent=2, sort_keys=True, default=str) + "\n").encode()
    return _atomic_bytes(path, raw)


def _write_text(path: Path, text: str) -> str:
    return _atomic_bytes(path, text.encode())


def _withdrawn_sidecar_path(localdata: Path, target_date: str) -> Path:
    return localdata / f"auto_tickets_{target_date}_withdrawn.json"


def _read_no_bet_decision(state: Dict[str, Any], target_date: str) -> Optional[Dict[str, Any]]:
    """Return one date's immutable inline no-bet decision from durable state."""

    lock = (state.get("frozen_by_date") or {}).get(target_date)
    if not isinstance(lock, dict) or lock.get("no_bet") is not True:
        return None
    decision = lock.get("no_bet_decision")
    if (
        not isinstance(decision, dict)
        or not schema_matches(decision.get("schema_version"), NO_BET_DECISION_SCHEMA)
        or decision.get("status") != "frozen_no_bet"
        or str(decision.get("target_date") or "") != target_date
        or not decision.get("decision_id")
    ):
        raise ValueError(f"invalid inline frozen no-bet decision for {target_date}")
    if lock.get("decision_id") != decision["decision_id"]:
        raise ValueError(f"no-bet lock decision id mismatch for {target_date}")
    return deepcopy(decision)


def _withdrawn_ids_from_sidecars(localdata: Path) -> set[str]:
    ids: set[str] = set()
    for path in localdata.glob("auto_tickets_*_withdrawn.json"):
        try:
            data = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        did = data.get("decision_id")
        if did:
            ids.add(str(did))
    return ids


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return deepcopy(default)
    try:
        return current_spelling(json.loads(path.read_text()))
    except (OSError, ValueError) as exc:
        # Financial/evidence state must never silently reset to a fresh bank or
        # an empty card. Corruption is an operator-visible failure.
        raise ValueError(f"cannot parse operator state {path}: {exc}") from exc


def _local_time(value: str, timezone_name: str) -> str:
    return parse_utc(value).astimezone(ZoneInfo(timezone_name)).strftime("%Y-%m-%d %H:%M %Z")


def _utc_iso(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _ticket_decision_window(
    games: Dict[str, Game],
    *,
    league: League,
    forecast_rows: List[Dict[str, Any]],
    target_date: str,
    generated_utc: str,
    slate_timezone_name: str,
) -> Dict[str, Any]:
    """Describe the immutable per-league decision point for a dated slate.

    NBA cards publish at 18:00 SAST so their evening board can mature. Other
    leagues retain the established 09:00 SAST window. A genuinely earlier game
    cannot be deferred to the normal publication time: its ordinary
    tip-minus-60-minute safety cutoff takes precedence. Fixture-level pick
    snapshots remain their own tip-minus-60-minute evidence lifecycle.
    """
    try:
        target_day = datetime.strptime(target_date, "%Y-%m-%d").date()
        generated = parse_utc(generated_utc)
        slate_zone = ZoneInfo(slate_timezone_name)
        publication_zone = ZoneInfo(TICKET_PUBLICATION_TIMEZONE)
    except (TypeError, ValueError) as exc:
        raise ValueError("invalid ticket decision date, time, or slate timezone") from exc

    publication_time = _ticket_publication_local_time(league)
    publication_policy = _ticket_decision_policy(league)
    publication_cutoff = datetime.combine(
        target_day, publication_time, tzinfo=publication_zone
    ).astimezone(UTC)
    slate_tips = []
    for game in games.values():
        try:
            tip = parse_utc(game.scheduled_utc)
        except (AttributeError, TypeError, ValueError):
            continue
        if tip.astimezone(slate_zone).date() == target_day:
            slate_tips.append(tip)
    slate_source = "games"
    if not slate_tips:
        # A candidate may survive a transient game-store lookup failure. Its
        # preserved scheduled time is still enough to fail closed at the same
        # cutoff; do not let a missing store row postpone publication forever.
        slate_source = "forecast_rows"
        for row in forecast_rows:
            try:
                tip = parse_utc(str(row["scheduled_utc"]))
            except (KeyError, TypeError, ValueError):
                continue
            if tip.astimezone(slate_zone).date() == target_day:
                slate_tips.append(tip)
    if not slate_tips:
        return {
            "status": "no_slate_game",
            "policy": publication_policy,
            "lead_minutes": int(TICKET_DECISION_LEAD_TIME.total_seconds() // 60),
            "publication_timezone": TICKET_PUBLICATION_TIMEZONE,
            "publication_local_time": publication_time.strftime("%H:%M"),
            "publication_cutoff_utc": _utc_iso(publication_cutoff),
            "slate_timezone": slate_timezone_name,
        }

    earliest_tip = min(slate_tips)
    pre_tip_cutoff = earliest_tip - TICKET_DECISION_LEAD_TIME
    cutoff = min(publication_cutoff, pre_tip_cutoff)
    if generated < cutoff:
        status = "pending_cutoff"
    elif generated < earliest_tip:
        status = "due"
    else:
        # Do not create a fresh position after the earliest slate game has
        # tipped merely because a later scan has a new price/candidate.
        status = "missed_pre_tip_window"
    return {
        "status": status,
        "policy": publication_policy,
        "lead_minutes": int(TICKET_DECISION_LEAD_TIME.total_seconds() // 60),
        "publication_timezone": TICKET_PUBLICATION_TIMEZONE,
        "publication_local_time": publication_time.strftime("%H:%M"),
        "publication_cutoff_utc": _utc_iso(publication_cutoff),
        "pre_tip_safety_cutoff_utc": _utc_iso(pre_tip_cutoff),
        "slate_timezone": slate_timezone_name,
        "slate_time_source": slate_source,
        "earliest_slate_tip_utc": _utc_iso(earliest_tip),
        "planned_cutoff_utc": _utc_iso(cutoff),
        "generated_utc": _utc_iso(generated),
    }


def _ticket_candidate_is_pre_tip(candidate: Dict[str, Any], generated_utc: str) -> bool:
    """Require every proposed leg to remain strictly pre-tip at publication."""
    legs = candidate.get("legs")
    if not isinstance(legs, list) or not legs:
        return False
    generated = parse_utc(generated_utc)
    try:
        return all(parse_utc(str(leg["scheduled_utc"])) > generated for leg in legs)
    except (KeyError, TypeError, ValueError):
        return False


def _finite_number(value: Any) -> Optional[float]:
    """Return a finite numeric value, rejecting bools and malformed evidence."""

    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    try:
        number = float(value)
    except (OverflowError, TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _euroleague_consensus_paper_candidates(
    rows: List[Dict[str, Any]],
    *,
    target_date: str,
    generated_utc: str,
    slate_timezone_name: str,
) -> List[Dict[str, Any]]:
    """Return the small, no-stake EuroLeague aggregate-consensus board.

    This is intentionally *not* a relaxation of ``acca_eligible``.  A row may
    enter only when every ordinary model-quality gate has passed and the only
    strict-lane blockers are the documented aggregate-price evidence axes (and,
    during the developing period, the sample-size gate).  The representative
    BetExplorer average remains explicitly unresolved/probationary throughout.
    """

    now = parse_utc(generated_utc)
    target_zone = ZoneInfo(slate_timezone_name)
    by_game: Dict[str, Dict[str, Any]] = {}
    for original in rows:
        row = deepcopy(original)
        book = row.get("book")
        if not isinstance(book, dict) or not _has_exact_selected_price(row):
            continue
        if (
            str(row.get("league") or "") != League.EUROLEAGUE.value
            or str(row.get("market") or "") != "moneyline"
            or str(row.get("period") or "") != Period.FULL_GAME.value
            or str(row.get("selection") or "") not in {"home", "away"}
            or str(book.get("provider") or "").strip().lower()
            != EUROLEAGUE_CONSENSUS_PROVIDER
            or str(book.get("book") or "").strip().lower() != EUROLEAGUE_CONSENSUS_BOOK
            or book.get("price_provenance") != PriceProvenance.AGGREGATE_CONSENSUS.value
            or book.get("source_permission") != SourcePermission.UNRESOLVED.value
            or book.get("operational_validation") != OperationalValidation.PROBATION.value
            or book.get("ticket_price_eligible") is True
            or book.get("pre_match") is not True
        ):
            continue
        try:
            observed = parse_utc(str(book["observed_utc"]))
            tip = parse_utc(str(row["scheduled_utc"]))
        except (KeyError, TypeError, ValueError):
            continue
        if (
            observed > now
            or observed >= tip
            or now >= tip
            or tip.astimezone(target_zone).date().isoformat() != target_date
        ):
            continue
        eligibility = row.get("eligibility")
        reasons = set(eligibility.get("reasons") or []) if isinstance(eligibility, dict) else set()
        if not reasons or not reasons.issubset(EUROLEAGUE_CONSENSUS_ALLOWED_BLOCKERS):
            continue
        # An aggregate row must remain blocked from the strict execution lane.
        if "price_not_named_bookmaker_quote" not in reasons:
            continue
        calibration = row.get("calibration")
        if not isinstance(calibration, dict):
            continue
        n = calibration.get("n")
        sufficient = calibration.get("sufficient_for_accas") is True
        if sufficient:
            calibration_tier = "fully_calibrated"
        elif (
            calibration.get("status") == "provisional"
            and isinstance(n, int)
            and not isinstance(n, bool)
            and n >= EUROLEAGUE_CONSENSUS_MIN_CALIBRATION_SAMPLES
            and "insufficient_walk_forward_samples" in reasons
        ):
            calibration_tier = "developing_calibration"
        else:
            continue
        probability = _finite_number(row.get("calibrated_probability"))
        price = _finite_number(book.get("selected_decimal_price"))
        if probability is None or price is None or not 0.0 < probability <= 1.0 or price <= 1.0:
            continue
        edge = probability * price - 1.0
        if edge < EUROLEAGUE_CONSENSUS_MIN_MODEL_EDGE:
            continue
        game_id = str(row.get("game_id") or "")
        if not game_id:
            continue
        row["consensus_paper"] = {
            "lane": "EUROLEAGUE_AGGREGATE_CONSENSUS_PRICE_PAPER_LEAGUE",
            "price_source": "BetExplorer representative average",
            "price_provenance": PriceProvenance.AGGREGATE_CONSENSUS.value,
            "model_consensus_edge": round(edge, 6),
            "minimum_model_consensus_edge": EUROLEAGUE_CONSENSUS_MIN_MODEL_EDGE,
            "calibration_tier": calibration_tier,
            "calibration_n": n,
            "maximum_daily_selections": EUROLEAGUE_CONSENSUS_MAX_PICKS,
            "no_stake": True,
            "strict_ticket_eligible": False,
        }
        prior = by_game.get(game_id)
        if prior is None or (
            float(row["consensus_paper"]["model_consensus_edge"]),
            probability,
            str(row.get("forecast_id") or ""),
        ) > (
            float((prior.get("consensus_paper") or {}).get("model_consensus_edge") or -math.inf),
            _finite_number(prior.get("calibrated_probability")) or -math.inf,
            str(prior.get("forecast_id") or ""),
        ):
            by_game[game_id] = row
    return sorted(
        by_game.values(),
        key=lambda row: (
            -float((row.get("consensus_paper") or {})["model_consensus_edge"]),
            -(_finite_number(row.get("calibrated_probability")) or 0.0),
            str(row.get("scheduled_utc") or ""),
            str(row.get("forecast_id") or ""),
        ),
    )[:EUROLEAGUE_CONSENSUS_MAX_PICKS]


def _consensus_paper_roster_path(localdata: Path, target_date: str) -> Path:
    return localdata / f"euroleague_consensus_paper_{target_date}_roster.json"


def _consensus_paper_card_path(localdata: Path, target_date: str) -> Path:
    return localdata / f"euroleague_consensus_paper_{target_date}.txt"


def _read_consensus_paper_roster(path: Path, *, target_date: Optional[str] = None) -> Dict[str, Any]:
    roster = _read_json(path, {})
    if (
        not isinstance(roster, dict)
        or not schema_matches(roster.get("schema_version"), EUROLEAGUE_CONSENSUS_PAPER_SCHEMA)
        or roster.get("league") != League.EUROLEAGUE.value
        or not isinstance(roster.get("selections"), list)
        or not isinstance(roster.get("frozen_at"), str)
    ):
        raise ValueError(f"invalid EuroLeague consensus-paper roster: {path}")
    if target_date is not None and roster.get("target_date") != target_date:
        raise ValueError(f"EuroLeague consensus-paper roster date mismatch: {path}")
    return roster


def _write_consensus_paper_roster(path: Path, roster: Dict[str, Any]) -> str:
    """Write a one-time no-stake roster; a later scan may never rotate it."""

    raw = json.dumps(roster, indent=2, sort_keys=True, default=str) + "\n"
    if path.exists():
        if path.read_text() != raw:
            raise ValueError(f"refusing to replace time-locked consensus-paper roster: {path}")
        return str(path)
    _write_json(path, roster)
    return str(path)


def _consensus_paper_summary(selections: List[Dict[str, Any]]) -> Dict[str, Any]:
    counts = Counter(str(row.get("grade_status") or "pending") for row in selections)
    wins = counts["win"]
    losses = counts["loss"]
    pushes = counts["push"]
    decisions = wins + losses
    scored = [row for row in selections if str(row.get("grade_status")) in {"win", "loss"}]
    brier_terms = []
    for row in scored:
        probability = _finite_number(row.get("calibrated_probability"))
        if probability is not None:
            brier_terms.append((probability - float(row.get("grade_status") == "win")) ** 2)
    return {
        "selections": len(selections),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "pending": counts["pending"],
        "decisions": decisions,
        "hit_rate": round(wins / decisions, 6) if decisions else None,
        "brier_score": round(sum(brier_terms) / len(brier_terms), 6) if brier_terms else None,
        "brier_n": len(brier_terms),
    }


def _render_euroleague_consensus_paper(
    *,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
    ticket_decision: Dict[str, Any],
    candidates: List[Dict[str, Any]],
    roster: Optional[Dict[str, Any]],
    games: Dict[str, Game],
) -> str:
    """Render the separate aggregate-price research board, never a ticket card."""

    frozen = roster is not None
    selections = deepcopy((roster or {}).get("selections") or candidates)
    for row in selections:
        _set_grade(row, games)
    summary = _consensus_paper_summary(selections)
    cutoff = ticket_decision.get("planned_cutoff_utc")
    status = ticket_decision.get("status")
    lines = [
        f"MA GOLIDE EUROLEAGUE CONSENSUS-PRICE PAPER LEAGUE — {target_date}",
        "=" * 72,
        "RESEARCH PAPER BOARD — BetExplorer representative average price; not a named bookmaker quote.",
        "NO STAKE · NO TICKET · NO RECEIPT · NO PAPER BANK · NO ROI/P&L · NO REAL-MONEY INSTRUCTION.",
        "The strict EuroLeague ticket lane remains fail-closed: aggregate/unresolved/probationary prices cannot enter it.",
        "",
        "QUALITY RULES",
        "- Full-game moneyline only; exact pre-tip BetExplorer average observed before the board scan.",
        "- All normal model gates must pass; only aggregate-price evidence blockers and a developing calibration sample may remain.",
        f"- Developing calibration requires n≥{EUROLEAGUE_CONSENSUS_MIN_CALIBRATION_SAMPLES}; fully calibrated rows remain preferred.",
        f"- Model-versus-consensus threshold: calibrated p × displayed average odds − 1 ≥ {EUROLEAGUE_CONSENSUS_MIN_MODEL_EDGE:.0%}.",
        f"- At most {EUROLEAGUE_CONSENSUS_MAX_PICKS} selections; one selection per fixture; ordered by model-versus-consensus edge.",
        "",
    ]
    if frozen:
        frozen_at = str(roster["frozen_at"])
        lines.extend(
            [
                f"TIME-LOCKED BOARD: {_local_time(frozen_at, timezone_name)} · source scan decision: {roster.get('decision_status', 'unknown')}",
                "Selections and their observed consensus prices are fixed in the roster JSON; only official-result grades refresh below.",
            ]
        )
    elif status == "pending_cutoff":
        lines.append(
            "DRAFT BOARD — selections are mutable until "
            f"{_local_time(str(cutoff), timezone_name) if cutoff else 'a confirmed cutoff'}; no roster exists yet."
        )
    elif status == "missed_pre_tip_window":
        lines.append("NO TIME-LOCKED BOARD — the earliest game already tipped before this scan; no late roster is fabricated.")
    elif status == "no_slate_game":
        lines.append("DRAFT BOARD UNAVAILABLE — no confirmed EuroLeague slate fixture is available for this date.")
    else:
        lines.append("DRAFT BOARD — awaiting a confirmed, eligible scan before the time-lock window.")
    lines.extend(
        [
            f"Latest render: {_local_time(generated_utc, timezone_name)}",
            "",
        ]
    )
    if not selections:
        lines.extend(["  (no qualifying consensus-paper selections)", ""])
    else:
        for index, row in enumerate(selections, 1):
            book = row.get("book") or {}
            paper = row.get("consensus_paper") or {}
            probability = _finite_number(row.get("calibrated_probability"))
            price = _finite_number(book.get("selected_decimal_price"))
            edge = _finite_number(paper.get("model_consensus_edge"))
            calibration = row.get("calibration") or {}
            reasons = ", ".join((row.get("eligibility") or {}).get("reasons") or []) or "none"
            lines.extend(
                [
                    f"  [CONSENSUS PAPER PICK #{index}] {row.get('home')} vs {row.get('away')} · "
                    f"TIP-OFF {_local_time(str(row.get('scheduled_utc')), timezone_name)}",
                    f"     {_selection_label(row)} @{'n/a' if price is None else f'{price:.3f}'} average consensus · "
                    f"calibrated p={'n/a' if probability is None else f'{probability:.1%}'} · "
                    f"model-v-consensus edge={'n/a' if edge is None else f'{edge:+.1%}'} · {_grade_badge(row.get('grade_status'))}",
                    f"     calibration={paper.get('calibration_tier', calibration.get('status', 'unknown'))} "
                    f"n={paper.get('calibration_n', calibration.get('n', 'n/a'))} · "
                    "evidence=aggregate_consensus / permission=unresolved / validation=probation",
                    f"     strict-ticket blockers: {reasons}",
                ]
            )
        lines.append("")
    hit_rate = summary["hit_rate"]
    brier = summary["brier_score"]
    lines.extend(
        [
            "OUTCOME TRACKING — selection accuracy/calibration only; displayed prices never generate profit, stake, or ROI accounting.",
            f"Selections {summary['selections']} · {summary['wins']}W-{summary['losses']}L-{summary['pushes']}P · "
            f"{summary['pending']} pending · hit rate {'n/a' if hit_rate is None else f'{hit_rate:.1%}'} · "
            f"Brier {'n/a' if brier is None else f'{brier:.4f}'} (n={summary['brier_n']}).",
            "",
        ]
    )
    return "\n".join(lines)


def _publish_euroleague_consensus_paper(
    *,
    localdata: Path,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
    slate_timezone_name: str,
    ticket_decision: Dict[str, Any],
    current_rows: List[Dict[str, Any]],
    games: Dict[str, Game],
    market_report_available: bool,
) -> Dict[str, Any]:
    """Publish the isolated EuroLeague consensus board and regrade old boards."""

    candidates = _euroleague_consensus_paper_candidates(
        current_rows,
        target_date=target_date,
        generated_utc=generated_utc,
        slate_timezone_name=slate_timezone_name,
    )
    roster_path = _consensus_paper_roster_path(localdata, target_date)
    roster = _read_consensus_paper_roster(roster_path, target_date=target_date) if roster_path.exists() else None
    decision_status = str(ticket_decision.get("status") or "unknown")
    frozen_now = False
    if roster is None and market_report_available and decision_status == "due":
        selections = []
        for candidate in candidates:
            selection = deepcopy(candidate)
            selection.pop("grade_status", None)
            selection.pop("outcome", None)
            selections.append(selection)
        roster = {
            "schema_version": EUROLEAGUE_CONSENSUS_PAPER_SCHEMA,
            "league": League.EUROLEAGUE.value,
            "target_date": target_date,
            "frozen_at": generated_utc,
            "decision_status": decision_status,
            "decision_policy": ticket_decision.get("policy"),
            "planned_cutoff_utc": ticket_decision.get("planned_cutoff_utc"),
            "price_contract": {
                "provider": EUROLEAGUE_CONSENSUS_PROVIDER,
                "book": EUROLEAGUE_CONSENSUS_BOOK,
                "price_provenance": PriceProvenance.AGGREGATE_CONSENSUS.value,
                "source_permission": SourcePermission.UNRESOLVED.value,
                "operational_validation": OperationalValidation.PROBATION.value,
            },
            "selection_policy": {
                "maximum_picks": EUROLEAGUE_CONSENSUS_MAX_PICKS,
                "minimum_developing_calibration_samples": EUROLEAGUE_CONSENSUS_MIN_CALIBRATION_SAMPLES,
                "minimum_model_consensus_edge": EUROLEAGUE_CONSENSUS_MIN_MODEL_EDGE,
                "one_selection_per_fixture": True,
                "full_game_moneyline_only": True,
                "no_stake": True,
                "no_ticket_receipt_bank_roi_or_pnl": True,
            },
            "selections": selections,
        }
        _write_consensus_paper_roster(roster_path, roster)
        frozen_now = True

    card_path = _consensus_paper_card_path(localdata, target_date)
    _write_text(
        card_path,
        _render_euroleague_consensus_paper(
            target_date=target_date,
            generated_utc=generated_utc,
            timezone_name=timezone_name,
            ticket_decision=ticket_decision,
            candidates=candidates,
            roster=roster,
            games=games,
        ),
    )
    refreshed_cards = []
    for old_roster_path in sorted(localdata.glob("euroleague_consensus_paper_????-??-??_roster.json")):
        old_date = old_roster_path.name.removeprefix("euroleague_consensus_paper_").removesuffix("_roster.json")
        if old_date == target_date:
            continue
        old_roster = _read_consensus_paper_roster(old_roster_path, target_date=old_date)
        _write_text(
            _consensus_paper_card_path(localdata, old_date),
            _render_euroleague_consensus_paper(
                target_date=old_date,
                generated_utc=generated_utc,
                timezone_name=timezone_name,
                ticket_decision={"status": "historical_time_locked"},
                candidates=[],
                roster=old_roster,
                games=games,
            ),
        )
        refreshed_cards.append(old_date)
    return {
        "lane": "EUROLEAGUE_AGGREGATE_CONSENSUS_PRICE_PAPER_LEAGUE",
        "paper_only": True,
        "no_stake_ticket_receipt_bank_roi_or_pnl": True,
        "candidate_count": len(candidates),
        "candidate_forecast_ids": [str(row.get("forecast_id")) for row in candidates],
        "roster": str(roster_path) if roster is not None else None,
        "card": str(card_path),
        "frozen_now": frozen_now,
        "decision_status": decision_status,
        "historical_cards_refreshed": refreshed_cards,
    }


def _bucket(forecast: Dict[str, Any]) -> str:
    if _is_derived_spread_ticket_row(forecast):
        return "DERIVED_SPREAD_QUALIFIED"
    book = forecast.get("book") or {}
    price = book.get("selected_decimal_price")
    eligibility = forecast.get("eligibility") or {}
    reasons = set(eligibility.get("reasons") or [])
    if eligibility.get("acca_eligible"):
        return "ELIGIBLE_CLEAN"
    if price is None:
        return "WATCHLIST_NO_ODDS"
    if reasons == {"insufficient_walk_forward_samples"}:
        return "CAUTION"
    if "insufficient_walk_forward_samples" in reasons and not (
        reasons - {"insufficient_walk_forward_samples", "below_calibrated_probability_floor"}
    ):
        return "WATCHLIST_INSUFFICIENT_CONTEXT"
    return "SKIPPED_INELIGIBLE"


def _price_evidence(forecast: Dict[str, Any]) -> str:
    """Render price provenance without upgrading an aggregate quote to a book."""

    book = forecast.get("book") or {}
    provenance = book.get("price_provenance")
    if book.get("selected_decimal_price") is not None:
        if provenance == PriceProvenance.AGGREGATE_CONSENSUS.value:
            return "AGGREGATE_CONSENSUS_PRICE"
        if provenance == PriceProvenance.NO_PRICE.value:
            return "UNPRICED_SOURCE_CONTRADICTION"
        return "NAMED_BOOKMAKER_PRICE"
    if provenance == PriceProvenance.NO_PRICE.value:
        return "UNPRICED_EVIDENCE"
    if book:
        return "MATCHED_BUT_SELECTED_PRICE_MISSING"
    return "UNMATCHED"


def _has_exact_selected_price(row: Dict[str, Any]) -> bool:
    """Whether a row has a usable exact selected-side decimal price."""
    book = row.get("book") or {}
    value = book.get("selected_decimal_price") if isinstance(book, dict) else None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    try:
        price = float(value)
    except OverflowError:
        return False
    return math.isfinite(price) and price > 1.0


def _opposite_selection_for_row(row: Dict[str, Any]) -> Optional[str]:
    book = row.get("book")
    side_prices = book.get("side_prices") if isinstance(book, dict) else None
    has_draw = isinstance(side_prices, dict) and side_prices.get("draw") is not None
    return opposite_selection(
        str(row.get("market") or ""),
        str(row.get("period") or ""),
        str(row.get("selection") or ""),
        has_draw=has_draw,
    )


def _same_line(left: Any, right: Any) -> bool:
    if left is None or right is None:
        return left is right
    if isinstance(left, bool) or isinstance(right, bool):
        return False
    try:
        left_value, right_value = float(left), float(right)
    except (OverflowError, TypeError, ValueError):
        return False
    return (
        math.isfinite(left_value)
        and math.isfinite(right_value)
        and math.isclose(left_value, right_value, rel_tol=0.0, abs_tol=1e-9)
    )


def opposite_quote(row: Dict[str, Any]) -> tuple[Optional[float], Optional[str]]:
    """Return the exact opposite-side quote or a reason it is unavailable.

    The quote must be present in this row's preserved BookSnapshot evidence,
    whose game, market, period, line, bookmaker, and observation are checked.
    No odds or missing market sides are inferred.
    """
    selection = _opposite_selection_for_row(row)
    if selection is None:
        return None, "market_has_no_single_opposite_side"

    book = row.get("book")
    if not isinstance(book, dict):
        return None, "no_book_evidence"
    side_prices = book.get("side_prices")
    if not isinstance(side_prices, dict):
        return None, "opposite_quote_not_preserved"
    if selection not in side_prices or side_prices[selection] is None:
        return None, "opposite_quote_missing_in_payload"

    selected_price = book.get("selected_decimal_price")
    selected_side = str(row.get("selection") or "")
    if selected_price is not None:
        preserved_selected = side_prices.get(selected_side)
        if (
            isinstance(selected_price, bool)
            or not isinstance(selected_price, (int, float))
            or isinstance(preserved_selected, bool)
            or not isinstance(preserved_selected, (int, float))
        ):
            return None, "selected_quote_inconsistent"
        try:
            selected_value = float(selected_price)
            preserved_value = float(preserved_selected)
        except OverflowError:
            return None, "selected_quote_inconsistent"
        if (
            not math.isfinite(selected_value)
            or not math.isfinite(preserved_value)
            or not math.isclose(selected_value, preserved_value, rel_tol=0.0, abs_tol=1e-9)
        ):
            return None, "selected_quote_inconsistent"

    if str(row.get("market") or "") in {"spread", "total"}:
        if row.get("line") is None:
            return None, "market_line_missing"
        if not _same_line(row.get("line"), row.get("line")):
            return None, "invalid_market_line"

    if (
        not row.get("game_id")
        or book.get("game_id") != row.get("game_id")
        or book.get("market") != row.get("market")
        or book.get("period") != row.get("period")
        or not _same_line(book.get("line"), row.get("line"))
    ):
        return None, "opposite_quote_scope_mismatch"
    if book.get("pre_match") is not True:
        return None, "not_pre_match"
    if not str(book.get("provider") or "").strip() or not str(book.get("book") or "").strip():
        return None, "unnamed_book_quote"
    try:
        observed = parse_utc(str(book["observed_utc"]))
        scheduled = parse_utc(str(row["scheduled_utc"]))
    except (KeyError, TypeError, ValueError):
        return None, "missing_quote_timestamp"
    if observed >= scheduled:
        return None, "not_pre_tip"

    raw_price = side_prices[selection]
    if isinstance(raw_price, bool) or not isinstance(raw_price, (int, float)):
        return None, "invalid_opposite_quote"
    try:
        price = float(raw_price)
    except OverflowError:
        return None, "invalid_opposite_quote"
    if not math.isfinite(price) or price <= 1.0:
        return None, "invalid_opposite_quote"
    return price, None


def _confidence_probability(row: Dict[str, Any]) -> tuple[Optional[float], str]:
    value = row.get("calibrated_probability")
    basis = "calibrated_probability"
    if value is None:
        value = row.get("raw_probability")
        basis = "raw_probability"
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None, basis
    try:
        probability = float(value)
    except OverflowError:
        return None, basis
    if not math.isfinite(probability) or not 0.0 <= probability <= 1.0:
        return None, basis
    return probability, basis


def _confidence_band(row: Dict[str, Any]) -> str:
    probability, _ = _confidence_probability(row)
    if probability is None:
        return "unknown"
    if probability < 0.5:
        return "<0.50"
    if probability < 0.6:
        return "0.50-<0.60"
    if probability < 0.7:
        return "0.60-<0.70"
    if probability < 0.8:
        return "0.70-<0.80"
    if probability < 0.9:
        return "0.80-<0.90"
    return "0.90-1.00"


def _fade_metrics(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Describe counterfactual fade results using exact opposite-side quotes."""
    missing_reasons: Counter[str] = Counter()
    wins = losses = pushes = priced_sample = 0
    profit = 0.0
    implied_break_even_sum = 0.0
    decision_quotes = 0
    quoted_prices: List[float] = []

    for row in rows:
        price, reason = opposite_quote(row)
        if price is None:
            missing_reasons[reason or "opposite_quote_unavailable"] += 1
            continue

        fade_grade = {"win": "loss", "loss": "win", "push": "push"}.get(
            str(row.get("grade_status"))
        )
        if fade_grade is None:
            missing_reasons["pick_not_settled"] += 1
            continue

        priced_sample += 1
        quoted_prices.append(price)
        if fade_grade == "win":
            wins += 1
            profit += price - 1.0
        elif fade_grade == "loss":
            losses += 1
            profit -= 1.0
        else:
            pushes += 1
        if fade_grade != "push":
            decision_quotes += 1
            implied_break_even_sum += 1.0 / price

    decisions = wins + losses
    return {
        "candidate_sample": len(rows),
        "priced_sample": priced_sample,
        "unpriced_sample": len(rows) - priced_sample,
        "unpriced_reasons": dict(sorted(missing_reasons.items())),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "decisions": decisions,
        "hit_rate": wins / decisions if decisions else None,
        "profit_units": profit,
        "roi": profit / priced_sample if priced_sample else None,
        "mean_opposite_decimal_price": (
            sum(quoted_prices) / len(quoted_prices) if quoted_prices else None
        ),
        "break_even_hit_rate": (
            implied_break_even_sum / decision_quotes if decision_quotes else None
        ),
        "break_even_sample": decision_quotes,
    }


def _fade_breakdown(rows: List[Dict[str, Any]], key_for_row) -> Dict[str, Any]:
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[str(key_for_row(row))].append(row)
    return {key: _fade_metrics(grouped[key]) for key in sorted(grouped)}


def _pick_row(forecast: Dict[str, Any], target_date: str) -> Dict[str, Any]:
    row = deepcopy(forecast)
    row["pick_id"] = row["forecast_id"]
    row["date"] = target_date
    row["bucket"] = _bucket(row)
    row["price_evidence"] = _price_evidence(row)
    row["grade_status"] = "pending"
    row["outcome"] = None
    row["paper_only"] = True
    return row


def _is_derived_spread_ticket_row(row: Dict[str, Any]) -> bool:
    """Whether this frozen row is a pre-tip moneyline-to-spread derivative."""

    derivation = row.get("derived_spread")
    return isinstance(derivation, dict) and derivation.get("contract") == "unpriced_moneyline_spread"


def _derived_spread_grade(row: Dict[str, Any], game: Game) -> Optional[str]:
    """Grade a derived spread with its selected-side posted line.

    Derived rows retain the source moneyline probability, which must never be
    interpreted as a spread-cover probability.  The grade instead uses only
    the frozen full-game quoted spread line and selected side.
    """

    if game.status is not GameStatus.FINAL or game.actual is None or not game.actual.reconciles()[0]:
        return None
    scores = _period_scores(game, Period.FULL_GAME)
    if scores is None:
        return None
    home_score, away_score = scores
    derivation = row.get("derived_spread")
    if not isinstance(derivation, dict):
        return None
    selected_line = derivation.get("selected_spread_line")
    selection = str(row.get("selection") or "").lower()
    if selection not in {"home", "away"} or isinstance(selected_line, bool):
        return None
    try:
        line = float(selected_line)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(line):
        return None
    margin = home_score - away_score
    result = margin + line if selection == "home" else -margin + line
    return "win" if result > 0 else "loss" if result < 0 else "push"


def _forecast_from_row(row: Dict[str, Any]) -> Forecast:
    book_row = row.get("book")
    book = None
    if book_row:
        stored_prices = book_row.get("side_prices")
        if isinstance(stored_prices, dict):
            prices = {str(key): value for key, value in stored_prices.items()}
            prices.setdefault(
                str(row["selection"]), book_row.get("selected_decimal_price")
            )
        else:
            # Legacy rows retain only the selected quote; never backfill sides.
            prices = {str(row["selection"]): book_row.get("selected_decimal_price")}
        provider = str(book_row.get("provider") or "unknown")
        legacy_aggregate = provider == "betexplorer"
        # A historical receipt missing evidence axes cannot prove permission or
        # provider validation. Keep it gradeable as immutable historical fact,
        # but never let a decode upgrade it into a fresh ticket price.
        book = BookEvidence(
            provider=provider,
            book=str(book_row.get("book") or "unknown"),
            observed_utc=str(book_row["observed_utc"]),
            source_modified_utc=book_row.get("source_modified_utc"),
            pre_match=bool(book_row.get("pre_match", True)),
            main_line=bool(book_row.get("main_line", True)),
            frozen=bool(book_row.get("frozen", False)),
            prices=prices,
            price_provenance=PriceProvenance(
                book_row.get(
                    "price_provenance",
                    "aggregate_consensus" if legacy_aggregate else "named_book",
                )
            ),
            source_permission=SourcePermission(
                book_row.get("source_permission", "unresolved")
            ),
            operational_validation=OperationalValidation(
                book_row.get(
                    "operational_validation",
                    "probation" if legacy_aggregate else "unverified",
                )
            ),
            game_id=book_row.get("game_id"),
            market=book_row.get("market"),
            period=book_row.get("period"),
            line=book_row.get("line"),
        )
    prediction = row.get("prediction") or {}
    explanation = row.get("explanation") or {}
    shadow_explanation = explanation.get("shadow_calibration") or {}
    shadow_calibration = dict(shadow_explanation) if isinstance(shadow_explanation, dict) else {}
    if shadow_calibration:
        candidate_probability = shadow_calibration.pop("candidate_probability", None)
        shadow_calibration.setdefault("probability", candidate_probability)
        if "ess" not in shadow_calibration and "effective_sample_size" in explanation:
            shadow_calibration["ess"] = explanation["effective_sample_size"]
    return Forecast(
        game_id=str(row["game_id"]),
        league=str(row["league"]),
        scheduled_utc=str(row["scheduled_utc"]),
        home_name=str(row["home"]),
        away_name=str(row["away"]),
        market=str(row["market"]),
        period=str(row["period"]),
        selection=str(row["selection"]),
        raw_probability=float(row["raw_probability"]),
        alternative_probabilities=dict(row.get("alternative_probabilities") or {}),
        source=str(row.get("model") or "unknown"),
        line=row.get("line"),
        push_probability=float(row.get("push_probability") or 0.0),
        expected_margin=row.get("expected_margin"),
        expected_total=row.get("expected_total"),
        training_games=int(row.get("training_games") or 0),
        residual_games=int(row.get("residual_games") or 0),
        team_history=(dict(row["team_history"]) if row.get("team_history") is not None else None),
        book=book,
        prediction_provider=prediction.get("provider"),
        prediction_observed_utc=prediction.get("observed_utc"),
        prediction_frozen=bool(prediction.get("frozen", False)),
        prediction_model=prediction.get("model"),
        prediction_model_version=prediction.get("model_version"),
        prediction_produced_by=prediction.get("produced_by"),
        role=row.get("role"),
        outcome=row.get("outcome"),
        calibrated_probability=row.get("calibrated_probability"),
        calibration_version=row.get("calibration_version"),
        shadow_calibration=shadow_calibration,
        calibration=dict(row.get("calibration") or {}),
        eligibility=dict(row.get("eligibility") or {}),
    )


def load_settled_pick_history(localdata: Path) -> Optional[List[Forecast]]:
    """Read the rolling audit's settled single-pick rows for paper ROI bands.

    ``None`` means no audit exists and lets the forecasting service use its
    own prequential settled individual forecasts. A valid but empty/unusable
    audit is returned for the service's explicit cold-start fallback decision.
    A malformed audit fails closed rather than substituting acca outcomes.
    """
    path = Path(localdata) / "picks_audit_rolling.json"
    if not path.exists():
        return None
    report = _read_json(path, {})
    if not isinstance(report, dict) or not isinstance(report.get("settled_ledger"), list):
        raise ValueError(f"settled pick audit has no valid settled_ledger: {path}")

    history = []
    for index, row in enumerate(report["settled_ledger"]):
        if not isinstance(row, dict):
            raise ValueError(f"invalid settled pick row {index} in {path}")
        outcome = row.get("grade_status")
        if outcome not in {"win", "loss", "push"}:
            outcome = row.get("outcome")
        if outcome not in {"win", "loss", "push"}:
            continue
        settled_row = deepcopy(row)
        settled_row["outcome"] = outcome
        try:
            history.append(_forecast_from_row(settled_row))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"invalid settled pick row {index} in {path}: {exc}") from exc
    return history


def _grade_row(row: Dict[str, Any], games: Dict[str, Game]) -> Optional[str]:
    game = games.get(str(row["game_id"]))
    if game is None or game.status is not GameStatus.FINAL:
        return None
    if _is_derived_spread_ticket_row(row):
        return _derived_spread_grade(row, game)
    return grade_forecast(_forecast_from_row(row), game)


def _set_grade(row: Dict[str, Any], games: Dict[str, Game]) -> None:
    outcome = _grade_row(row, games)
    row["outcome"] = outcome
    row["grade_status"] = outcome or "pending"


def _spread_line_for_selection(
    row: Dict[str, Any], selection: Optional[str] = None
) -> Optional[float]:
    """Convert a canonical home-oriented spread line to the selected side."""
    line = row.get("line")
    side = str(
        selection if selection is not None else row.get("selection") or ""
    ).lower()
    if (
        side not in {"home", "away"}
        or isinstance(line, bool)
        or not isinstance(line, (int, float))
    ):
        return None
    value = float(line)
    if not math.isfinite(value):
        return None
    return value if side == "home" else -value


def _selection_label(row: Dict[str, Any]) -> str:
    market = str(row.get("market") or "?").replace("_", " ").upper()
    period = str(row.get("period") or "?")
    selection = str(row.get("selection") or "?").upper()
    line = row.get("line")
    if _is_derived_spread_ticket_row(row):
        derivation = row["derived_spread"]
        selected_line = derivation.get("selected_spread_line")
        if isinstance(selected_line, (int, float)) and not isinstance(selected_line, bool):
            line_text = f" {float(selected_line):+g}"
        else:
            line_text = ""
        return f"DERIVED SPREAD {period}{line_text} -> {selection}"
    if market == "SPREAD":
        selected_line = _spread_line_for_selection(row)
        line_text = "" if selected_line is None else f" {selected_line:+g}"
    else:
        line_text = "" if line is None else f" {float(line):g}"
    return f"{market} {period}{line_text} -> {selection}"


def _canonical_book_for_key(book_name: str | None) -> str | None:
    """Canonicalize book names stored in historical pick rows.

    Older rows used ``Draft Kings`` (space) while newer rows use
    ``DraftKings``. Without normalization the same market would be counted
    as two logical picks, inflating archived/settled totals and ROI.
    This is a presentation-layer fix; canonical snapshots are already
    normalized at ingestion time.
    """
    if not book_name:
        return book_name
    # Local import to avoid circular dependency at module import time.
    try:
        from ingestion.identity import canonical_book_name

        canonical = canonical_book_name(book_name)
        return canonical or book_name
    except Exception:
        return book_name


def _market_pick_key(row: Dict[str, Any]) -> tuple:
    """Return the wager identity without its bookmaker observation.

    A no-price forecast is emitted with ``book=None``.  When a later scan
    finds an exact quote for the same game, market, period, side, line, and
    model, the old no-price evidence must not remain visible as a separate
    current pick merely because its book fields are empty.  The raw evidence
    is still retained in the dated ledger; this key is used only to recognise
    that a valid quote now exists for the same forecasted selection.
    """

    line = row.get("line")
    return (
        row.get("game_id"),
        row.get("market"),
        row.get("period"),
        row.get("selection"),
        float(line) if line is not None else None,
        row.get("model"),
    )


def _logical_pick_key(row: Dict[str, Any]) -> tuple:
    """Identify repeated observations of one displayed pick, not its evidence ID."""

    book = row.get("book") or {}
    return (*_market_pick_key(row), book.get("provider"), _canonical_book_for_key(book.get("book")))


def _evidence_time(row: Dict[str, Any]) -> datetime:
    timestamps = [
        (row.get("book") or {}).get("observed_utc"),
        (row.get("prediction") or {}).get("observed_utc"),
    ]
    observed = [parse_utc(value) for value in timestamps if value]
    return max(observed, default=parse_utc(row["scheduled_utc"]))


def _latest_evidence_utc(row: Dict[str, Any]) -> Optional[str]:
    timestamps = [
        value
        for value in (
            (row.get("book") or {}).get("observed_utc"),
            (row.get("prediction") or {}).get("observed_utc"),
        )
        if value
    ]
    return max(timestamps, key=parse_utc) if timestamps else None


def _latest_logical_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Select displayable current observations without deleting raw evidence.

    Book-attributed observations remain separate logical picks because their
    prices are independently auditable.  A no-price observation is different:
    it is a temporary absence of book evidence, not a bookmaker lane.  Once
    any exact selected-side quote exists for the same modelled selection, omit
    that no-price duplicate from the current/card/audit view.  This prevents a
    later successful odds refresh from still rendering ``no matched book`` for
    a pick that is in fact priced, while preserving every source row in the
    dated JSON ledger.
    """

    groups: Dict[tuple, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[_logical_pick_key(row)].append(row)
    selected = [
        max(group, key=lambda row: (_evidence_time(row), str(row["pick_id"])))
        for group in groups.values()
    ]
    exact_price_market_keys = {
        _market_pick_key(row) for row in selected if _has_exact_selected_price(row)
    }
    displayable = [
        row
        for row in selected
        if _has_exact_selected_price(row)
        or _market_pick_key(row) not in exact_price_market_keys
    ]
    return sorted(
        displayable,
        key=lambda row: (
            row["scheduled_utc"],
            row["game_id"],
            row["market"],
            row["period"],
            row["selection"],
            str(row["pick_id"]),
        ),
    )


def _pick_lifecycle_state_path(localdata: Path) -> Path:
    """Return the sole Git-durable home for per-fixture frozen snapshots."""

    return localdata / "pick_lifecycle_state.json"


def _new_pick_lifecycle_state(league: str) -> Dict[str, Any]:
    return {
        "schema_version": PICK_LIFECYCLE_STATE_SCHEMA,
        "league": league,
        "frozen_by_date": {},
        "pending_frozen_games": [],
        "dated_reports": {},
    }


def _load_pick_lifecycle_state(localdata: Path, *, league: str) -> Dict[str, Any]:
    """Load durable lifecycle state, retaining old empty states compatibly."""

    path = _pick_lifecycle_state_path(localdata)
    if not path.exists():
        return _new_pick_lifecycle_state(league)
    state = _read_json(path, {})
    if (
        not isinstance(state, dict)
        or not schema_matches(state.get("schema_version"), PICK_LIFECYCLE_STATE_SCHEMA)
        or state.get("league") != league
    ):
        raise ValueError(f"invalid pick lifecycle state: {path}")
    frozen_by_date = state.setdefault("frozen_by_date", {})
    if not isinstance(frozen_by_date, dict):
        raise ValueError(f"invalid frozen pick lifecycle state: {path}")
    return state


def _load_pick_freezes(
    lifecycle_state: Dict[str, Any], *, target_date: str, league: str
) -> Dict[str, Dict[str, Any]]:
    """Read immutable fixture snapshots embedded in durable lifecycle state."""

    frozen_by_date = lifecycle_state.get("frozen_by_date")
    if not isinstance(frozen_by_date, dict):
        raise ValueError("invalid frozen pick lifecycle state")
    payloads = frozen_by_date.get(target_date, {})
    if not isinstance(payloads, dict):
        raise ValueError(f"invalid frozen pick date state: {target_date}")
    freezes: Dict[str, Dict[str, Any]] = {}
    for game_id, payload in sorted(payloads.items()):
        if not isinstance(payload, dict):
            raise ValueError(f"invalid immutable inline pick freeze: {target_date}/{game_id}")
        canonical_game_id = str(payload.get("game_id") or "")
        rows = payload.get("rows")
        if (
            not schema_matches(payload.get("schema_version"), PICK_FREEZE_SCHEMA)
            or payload.get("target_date") != target_date
            or payload.get("league") != league
            or not canonical_game_id
            or canonical_game_id != str(game_id)
            or not isinstance(rows, list)
        ):
            raise ValueError(f"invalid immutable inline pick freeze: {target_date}/{game_id}")
        try:
            scheduled_utc = parse_utc(str(payload["scheduled_utc"]))
            cutoff_utc = parse_utc(str(payload["decision_cutoff_utc"]))
            frozen_at = parse_utc(str(payload["frozen_at"]))
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                f"inline pick freeze has invalid lifecycle timestamps: {target_date}/{game_id}"
            ) from exc
        if (
            cutoff_utc != scheduled_utc - TICKET_DECISION_LEAD_TIME
            or not cutoff_utc <= frozen_at < scheduled_utc
        ):
            raise ValueError(f"inline pick freeze violates its fixture cutoff: {target_date}/{game_id}")
        for row in rows:
            if not isinstance(row, dict) or str(row.get("game_id")) != canonical_game_id:
                raise ValueError(
                    f"inline pick freeze contains inconsistent fixture evidence: {target_date}/{game_id}"
                )
            try:
                row_scheduled_utc = parse_utc(str(row.get("scheduled_utc")))
            except ValueError as exc:
                raise ValueError(
                    f"inline pick freeze has invalid row schedule: {target_date}/{game_id}"
                ) from exc
            if row_scheduled_utc != scheduled_utc:
                raise ValueError(
                    f"inline pick freeze contains inconsistent fixture evidence: {target_date}/{game_id}"
                )
        freezes[canonical_game_id] = payload
    return freezes


def _freeze_due_pick_fixtures(
    localdata: Path,
    *,
    lifecycle_state: Dict[str, Any],
    target_date: str,
    league: str,
    rows: List[Dict[str, Any]],
    generated_utc: str,
) -> Dict[str, Dict[str, Any]]:
    """Freeze fixture evidence inline once it enters tip minus 60 minutes.

    A runner can arrive slightly after the exact cutoff, but never after tip.
    The immutable snapshot is nested by date and fixture in the Git-durable
    lifecycle state, so later drafts cannot overwrite it and no sidecar is
    needed for audit or settlement.
    """

    now = parse_utc(generated_utc)
    existing = _load_pick_freezes(lifecycle_state, target_date=target_date, league=league)
    frozen_by_date = lifecycle_state["frozen_by_date"].setdefault(target_date, {})
    froze_any = False
    for game_id, fixture_rows in itertools.groupby(
        _latest_logical_rows(rows), key=lambda row: str(row["game_id"])
    ):
        fixture_rows = list(fixture_rows)
        tips = {str(row["scheduled_utc"]) for row in fixture_rows}
        if len(tips) != 1:
            raise ValueError(f"pick freeze fixture has inconsistent scheduled_utc: {game_id}")
        tip = parse_utc(tips.pop())
        cutoff = tip - TICKET_DECISION_LEAD_TIME
        if game_id in existing or not (cutoff <= now < tip):
            continue
        frozen_rows = deepcopy(fixture_rows)
        for row in frozen_rows:
            # A frozen snapshot is evidence, not a settlement. Keep outcome
            # blank until an official final is read through the lifecycle view.
            row["outcome"] = None
            row["grade_status"] = "pending"
        payload = {
            "schema_version": PICK_FREEZE_SCHEMA,
            "target_date": target_date,
            "league": league,
            "game_id": game_id,
            "scheduled_utc": tip.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "decision_cutoff_utc": cutoff.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "frozen_at": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "rows": frozen_rows,
        }
        frozen_by_date[game_id] = payload
        existing[game_id] = payload
        froze_any = True
    if froze_any:
        # Persist immediately: an interrupted run must not reopen a fixture
        # that has already reached its individual pre-tip cutoff.
        pending = {
            (str(item[0]), str(item[1]))
            for item in lifecycle_state.get("pending_frozen_games", [])
            if isinstance(item, (list, tuple)) and len(item) == 2
        }
        pending.update(
            (game_id, str(payload["scheduled_utc"]))
            for game_id, payload in existing.items()
        )
        lifecycle_state["pending_frozen_games"] = [list(item) for item in sorted(pending)]
        _write_json(_pick_lifecycle_state_path(localdata), lifecycle_state)
    return existing


def _pick_lifecycle_rows(
    raw_rows: List[Dict[str, Any]],
    *,
    freezes: Dict[str, Dict[str, Any]],
    games: Dict[str, Game],
    generated_utc: str,
) -> tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """Render frozen evidence in preference to mutable draft observations."""

    now = parse_utc(generated_utc)
    frozen_game_ids = set(freezes)
    display_rows = [
        deepcopy(row) for row in raw_rows if str(row.get("game_id")) not in frozen_game_ids
    ]
    pending_frozen_games: List[Dict[str, str]] = []
    for game_id, payload in freezes.items():
        fixture_rows = deepcopy(payload["rows"])
        frozen_at = str(payload["frozen_at"])
        fixture_pending = False
        for row in fixture_rows:
            _set_grade(row, games)
            game = games.get(game_id)
            if row["grade_status"] in {"win", "loss", "push", "void"}:
                lifecycle = "SETTLED"
            elif game is not None and game.status is GameStatus.FINAL:
                lifecycle = "RESULT_UNRESOLVABLE"
                fixture_pending = True
            else:
                lifecycle = "FROZEN"
                fixture_pending = True
            row["lifecycle_status"] = lifecycle
            row["frozen_at"] = frozen_at
            row["decision_cutoff_utc"] = payload["decision_cutoff_utc"]
        if fixture_pending:
            pending_frozen_games.append(
                {"game_id": game_id, "scheduled_utc": str(payload["scheduled_utc"])}
            )
        display_rows.extend(fixture_rows)
    for row in display_rows:
        if str(row.get("game_id")) in frozen_game_ids:
            continue
        _set_grade(row, games)
        if row["grade_status"] in {"win", "loss", "push", "void"}:
            row["lifecycle_status"] = "SETTLED_UNFROZEN"
        elif parse_utc(row["scheduled_utc"]) <= now:
            row["lifecycle_status"] = "UNFROZEN_AWAITING_RESULT"
        else:
            row["lifecycle_status"] = "DRAFT"
    return display_rows, pending_frozen_games


def _render_picks(
    rows: List[Dict[str, Any]],
    *,
    league: str,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
) -> str:
    logical_rows = _latest_logical_rows(rows)
    revision_counts = Counter(_logical_pick_key(row) for row in rows)
    generated = parse_utc(generated_utc)
    settled = [row for row in logical_rows if row["grade_status"] in {"win", "loss", "push"}]
    pending = [row for row in logical_rows if row["grade_status"] == "pending"]
    upcoming = [row for row in pending if parse_utc(row["scheduled_utc"]) > generated]
    started = [row for row in pending if parse_utc(row["scheduled_utc"]) <= generated]

    lines = [
        f"Ma Golide Picks — {league} {target_date}",
        "=" * 72,
        f"Generated at: {_local_time(generated_utc, timezone_name)}",
        "PAPER ONLY · eligibility labels are evidence states, not profit guarantees.",
        "Draft rows remain mutable until their fixture enters the 60-minute pre-tip window.",
        "Only FROZEN rows are immutable settlement evidence; all final grades use official results.",
        "Started and settled games are retained separately as archive status.",
        "",
    ]
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in upcoming:
        grouped[str(row.get("bucket") or "UNKNOWN")].append(row)
    order = [*BUCKET_ORDER, *sorted(set(grouped) - set(BUCKET_ORDER))]

    def append_pick(row: Dict[str, Any]) -> None:
        bucket = str(row.get("bucket") or "UNKNOWN")
        probability = row.get("calibrated_probability")
        if probability is None:
            probability = row["raw_probability"]
        price = (row.get("book") or {}).get("selected_decimal_price")
        price_text = f"@{float(price):.3f}" if price is not None else "@n/a"
        book_name = (
            _canonical_book_for_key((row.get("book") or {}).get("book")) or "no matched book"
        )
        lines.append(
            f"  [{row.get('model')}/{row.get('model_version')}] "
            f"{row['home']} vs {row['away']} · "
            f"KO {_local_time(row['scheduled_utc'], timezone_name)}"
        )
        lines.append(
            f"      {_selection_label(row)} · p={probability:.1%} "
            f"(raw {row['raw_probability']:.1%}) {price_text} {book_name}"
        )
        eligibility = row.get("eligibility") or {}
        reasons = ",".join(eligibility.get("reasons") or []) or "none"
        calibration = row.get("calibration") or {}
        lifecycle = str(row.get("lifecycle_status") or "DRAFT")
        lifecycle_detail = f"lifecycle={lifecycle}"
        if row.get("frozen_at"):
            lifecycle_detail += f" · frozen_at={row['frozen_at']}"
        lines.append(
            f"      bucket={bucket} · price_evidence={row['price_evidence']} "
            f"· acca_eligible={bool(eligibility.get('acca_eligible'))} "
            f"· grade={_grade_badge(row['grade_status'])} · {lifecycle_detail}"
        )
        lines.append(
            f"      calibration={calibration.get('status', 'unknown')} "
            f"n={calibration.get('n', 0)} · reasons={reasons} "
            f"· prediction={(row.get('prediction') or {}).get('provider', 'history')}"
        )
        roi_band = row.get("paper_roi_band")
        if isinstance(roi_band, dict):
            weighted_roi = roi_band.get("recency_weighted_roi")
            adjusted_roi = roi_band.get("adjusted_roi")
            weighted_text = (
                f"{float(weighted_roi):+.1%}"
                if isinstance(weighted_roi, (int, float)) and not isinstance(weighted_roi, bool)
                else "n/a"
            )
            adjusted_text = (
                f"{float(adjusted_roi):+.1%}"
                if isinstance(adjusted_roi, (int, float)) and not isinstance(adjusted_roi, bool)
                else "n/a"
            )
            status = "PAPER SINGLE" if row.get("paper_roi_band_eligible") else "ROI-BAND BLOCKED"
            lines.append(
                f"      {status} · {roi_band.get('market', 'unknown')} "
                f"{roi_band.get('period', 'unknown')} · p {roi_band.get('probability_band', 'unknown')} "
                f"· odds {roi_band.get('individual_odds_band', 'unknown')} · "
                f"n={int(roi_band.get('historical_pick_count') or 0)} · "
                f"recency ROI {weighted_text} / adjusted {adjusted_text}"
            )
        elif row.get("paper_roi_band_eligible"):
            lines.append("      PAPER SINGLE · positive settled individual-pick ROI band")
        revisions = revision_counts[_logical_pick_key(row)]
        if revisions > 1:
            latest = _latest_evidence_utc(row) or "unknown"
            lines.append(f"      archived_revisions={revisions} · latest_evidence_utc={latest}")

    def sort_picks(picks: List[Dict[str, Any]]) -> None:
        picks.sort(
            key=lambda row: (
                -(row.get("calibrated_probability") or row["raw_probability"]),
                row["home"],
                row["away"],
                row["pick_id"],
            )
        )

    for bucket in order:
        picks = grouped.get(bucket, [])
        lines.extend(["", BUCKET_LABELS.get(bucket, bucket), "=" * 72])
        if not picks:
            lines.append("  (none)")
            continue
        sort_picks(picks)
        for row in picks:
            append_pick(row)

    for heading, picks in (
        ("IN PROGRESS / AWAITING OFFICIAL RESULT — ARCHIVE ONLY", started),
        ("SETTLED PICKS — ARCHIVE ONLY", settled),
    ):
        lines.extend(["", heading, "=" * 72])
        if not picks:
            lines.append("  (none)")
            continue
        sort_picks(picks)
        for row in picks:
            append_pick(row)

    rendered = len(upcoming) + len(started) + len(settled)
    if rendered != len(logical_rows):
        raise RuntimeError(
            f"daily report dropped logical picks: ledger={len(logical_rows)} rendered={rendered}"
        )
    counts = Counter(row["bucket"] for row in upcoming)
    lines.extend(
        [
            "",
            "Active summary: " + " ".join(f"{bucket}={counts.get(bucket, 0)}" for bucket in order),
            f"ROI-qualified individual paper singles: "
            f"{sum(bool(row.get('paper_roi_band_eligible')) for row in upcoming)}",
            f"Archive status: IN_PROGRESS_OR_AWAITING_RESULT={len(started)} SETTLED={len(settled)}",
            f"Logical picks shown: {len(logical_rows)} · "
            f"superseded observations: {len(rows) - len(logical_rows)}",
            f"Total archived pick evidence rows: {len(rows)}",
            "Flat one-paper-unit tickets only. No real-money stakes or automatic promotion.",
            "Bet only what you can afford to lose.",
            "",
        ]
    )
    return "\n".join(lines)


def _dated_pick_files(localdata: Path) -> List[Path]:
    return sorted(localdata.glob("picks_????-??-??.json"))


def _refresh_dated_pick_reports(
    localdata: Path,
    *,
    games: Dict[str, Game],
    league: str,
    generated_utc: str,
    timezone_name: str,
) -> Dict[str, Any]:
    """Regrade cards and render drafts/frozen evidence as one lifecycle view."""

    lifecycle_state = _load_pick_lifecycle_state(localdata, league=league)
    pending_frozen_games: List[Dict[str, str]] = []
    dated_reports: Dict[str, Dict[str, int]] = {}
    for path in _dated_pick_files(localdata):
        rows = _read_json(path, [])
        if not isinstance(rows, list):
            raise ValueError(f"dated pick ledger must be a list: {path}")
        for row in rows:
            _set_grade(row, games)
        _write_json(path, rows)
        target_date = path.stem.removeprefix("picks_")
        freezes = _load_pick_freezes(
            lifecycle_state, target_date=target_date, league=league
        )
        lifecycle_rows, pending = _pick_lifecycle_rows(
            rows,
            freezes=freezes,
            games=games,
            generated_utc=generated_utc,
        )
        pending_frozen_games.extend(pending)
        lifecycle = {
            "schema_version": PICK_LIFECYCLE_SCHEMA,
            "league": league,
            "target_date": target_date,
            "rows": lifecycle_rows,
        }
        _write_json(localdata / f"picks_{target_date}_lifecycle.json", lifecycle)
        _write_text(
            path.with_suffix(".txt"),
            _render_picks(
                lifecycle_rows,
                league=league,
                target_date=target_date,
                generated_utc=generated_utc,
                timezone_name=timezone_name,
            ),
        )
        dated_reports[target_date] = {
            "frozen_fixtures": len(freezes),
            "logical_picks": len(_latest_logical_rows(lifecycle_rows)),
        }
    # This durable state embeds immutable fixture snapshots and lets the
    # dependency-free gate request a result recheck after a frozen fixture has
    # tipped without publishing raw scan telemetry or one JSON file per game.
    lifecycle_state["schema_version"] = PICK_LIFECYCLE_STATE_SCHEMA
    lifecycle_state["league"] = league
    lifecycle_state["pending_frozen_games"] = [
        list(item)
        for item in sorted(
            {
                (str(item["game_id"]), str(item["scheduled_utc"]))
                for item in pending_frozen_games
            },
            key=lambda item: (item[1], item[0]),
        )
    ]
    lifecycle_state["dated_reports"] = dated_reports
    _write_json(_pick_lifecycle_state_path(localdata), lifecycle_state)
    return lifecycle_state


def _audit_period_label(period: str) -> str:
    return {
        "full_game": "Full-game",
        "H1": "First-half",
        "H2": "Second-half",
        "Q1": "First-quarter",
        "Q2": "Second-quarter",
        "Q3": "Third-quarter",
        "Q4": "Fourth-quarter",
    }.get(period, period.replace("_", " "))


def _spread_side_label(row: Dict[str, Any], selection: str) -> str:
    side = str(selection or "").lower()
    if side == "home":
        team = row.get("home")
    elif side == "away":
        team = row.get("away")
    else:
        team = None
    line = _spread_line_for_selection(row, side)
    parts = [str(team)] if team else []
    if line is not None:
        parts.append(f"{line:+g}")
    if not parts:
        return side or "?"
    label = " ".join(parts)
    return f"{label} ({side})" if side else label


def _audit_market_selection(row: Dict[str, Any]) -> str:
    period = _audit_period_label(str(row["period"]))
    market = str(row["market"])
    selection = str(row["selection"])
    selected_team = (
        row["home"] if selection == "home" else row["away"] if selection == "away" else None
    )
    if market == "moneyline":
        side = f"{selected_team} ({selection})" if selected_team else selection.title()
        return f"{period} moneyline: {side}"
    if market == "total":
        line = row.get("line")
        line_text = f" {float(line):g}" if line is not None else ""
        return f"{period} total{line_text}: {selection.title()}"
    if market == "spread":
        line = _spread_line_for_selection(row)
        line_text = f" {line:+g}" if line is not None else ""
        side = f"{selected_team} ({selection})" if selected_team else selection.title()
        return f"{period} spread{line_text}: {side}"
    if market == "highest_quarter":
        return f"Highest quarter: {selection.upper()}"
    return f"{period} {market.replace('_', ' ')}: {selection.title()}"


def _audit_actual_result(row: Dict[str, Any], games: Dict[str, Game] | None) -> str | None:
    game = (games or {}).get(str(row["game_id"]))
    if game is None or game.actual is None:
        return None
    try:
        scores = _period_scores(game, Period(str(row["period"])))
    except ValueError:
        return None
    if scores is None:
        return None
    home, away = scores
    period = str(row["period"])
    market = str(row["market"])
    if market == "total":
        return f"{_audit_period_label(period)} total: **{home + away:g}**"
    if market == "spread":
        return f"{_audit_period_label(period)} margin: **{home - away:+g}**"
    if period == Period.FULL_GAME.value:
        # The fixture heading already carries the full-game final score.
        return None
    return f"{_audit_period_label(period)} score: **{home:g}–{away:g}**"


def _render_pick_audit_markdown(
    report: Dict[str, Any],
    *,
    games: Dict[str, Game] | None,
    timezone_name: str,
) -> str:
    """Render the rolling machine audit as a readable report."""
    settled = list(report.get("settled_ledger") or [])
    wins = int(report["wins"])
    losses = int(report["losses"])
    decisions = wins + losses
    priced_sample = int(report["priced_sample"])
    grade_summary = _grade_count_summary(
        {"win": wins, "loss": losses, "push": int(report["pushes"])}
    ) or "none"
    lines = [
        f"# Ma Golide — Recent picks audit ({report['window']['start']} to {report['window']['end']})",
        "",
        f"Generated: {_local_time(report['generated_utc'], timezone_name)} · "
        f"{report['window']['days']}-day rolling window · Paper only",
        "",
        "## Overall",
        "",
        f"- archived logical picks: {report['archived_pick_rows']}",
        f"- dated evidence rows: {report['archived_evidence_rows']} "
        f"({report['superseded_observations']} superseded observations)",
        f"- settled picks: {report['settled_picks']} ({grade_summary})",
        (
            f"- hit rate: {report['hit_rate']:.1%} ({wins}/{decisions} win/loss decisions)"
            if decisions
            else "- hit rate: n/a (no win/loss decisions)"
        ),
        f"- pending/unmatched picks: {report['pending_picks']} prior-day · "
        f"{report['pending_picks_total']} total",
        f"- priced picks: {priced_sample}",
    ]
    if priced_sample:
        lines.extend(
            [
                f"- paper P&L: {report['profit_units']:+.3f} units",
                f"- ROI: {report['roi']:+.1%} (P&L / {priced_sample} priced "
                f"{'pick' if priced_sample == 1 else 'picks'})"
                + (" · **provisional — fewer than 30 priced picks**" if priced_sample < 30 else ""),
            ]
        )
    else:
        lines.append("- ROI: n/a (no settled priced picks)")

    lines.extend(
        [
            "",
            "## Settlement policy",
            "",
            "- reconciled official finals settle immediately; pending rows are never losses",
            f"- local-day pending cutoff: {report['same_day_cutoff']}",
            f"- same-day pending rows outside prior-day counters: {report['same_day_pending_picks']}",
            f"- same-day official-final picks included: {report['same_day_settled_picks']}",
            f"- eligible prior-day picks: {report['eligible_prior_picks']}",
            "- frozen fixtures grade only their immutable snapshot; unfrozen legacy/draft rows use the latest observation",
            "",
            "## Bucket performance",
            "",
        ]
    )
    by_bucket = report.get("by_bucket") or {}
    if by_bucket:
        lines.extend(
            [
                "| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for bucket, group in sorted(by_bucket.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if row["bucket"] == bucket
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{bucket}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    # Market breakdown - all leagues, requested for NCAA M/W watchlist analysis
    by_market = report.get("by_market") or {}
    lines.extend(["", "## By market", ""])
    if by_market:
        lines.extend(
            [
                "| Market | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for market, group in sorted(by_market.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if str(row.get("market")) == market
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{market}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    # Over/under + side bias — selection within market, every league
    by_selection = report.get("by_selection") or {}
    lines.extend(["", "## By selection — over/under & side bias", ""])
    if by_selection:
        lines.extend(
            [
                "| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for key, group in sorted(by_selection.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if f"{str(row.get('market'))}:{str(row.get('selection'))}" == key
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{key}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    fade = report.get("fade_audit") or {}
    fade_priced = int(fade.get("priced_sample") or 0)
    fade_candidates = int(fade.get("candidate_sample") or 0)
    fade_decisions = int(fade.get("decisions") or 0)
    fade_hit_rate = fade.get("hit_rate")
    fade_roi = fade.get("roi")
    break_even_hit_rate = fade.get("break_even_hit_rate")
    lines.extend(
        [
            "",
            "## Paper-only opposite-side fade audit",
            "",
            "Counterfactual reporting only — this is not a strategy, pick, or ticket. "
            "A settled result is reversed (win↔loss; push stays push) and priced only "
            "when the exact opposite quote is preserved in the same pre-tip book "
            "snapshot, market, period, line, and observation. No quote is inferred "
            "or historically backfilled; selected-side price is not required.",
            "",
            f"- settled rows with exact opposite quotes: {fade_priced}/{fade_candidates}",
            (
                f"- fade record: {fade.get('wins', 0)}–{fade.get('losses', 0)}–"
                f"{fade.get('pushes', 0)}"
            ),
            (
                f"- fade hit rate: {fade_hit_rate:.1%} "
                f"({fade.get('wins', 0)}/{fade_decisions} decisions)"
                if fade_hit_rate is not None
                else "- fade hit rate: n/a (no priced win/loss decisions)"
            ),
            (
                f"- paper P&L: {fade.get('profit_units', 0.0):+.3f} units · "
                f"ROI: {fade_roi:+.1%} (one unit per priced fade, pushes included)"
                if fade_roi is not None
                else "- fade ROI: n/a (no settled exact opposite-side quotes)"
            ),
            (
                f"- mean market-implied break-even hit rate: {break_even_hit_rate:.1%} "
                f"(n={fade.get('break_even_sample', 0)} non-push quotes; includes vig)"
                if break_even_hit_rate is not None
                else "- mean market-implied break-even hit rate: n/a"
            ),
        ]
    )
    fade_reasons = fade.get("unpriced_reasons") or {}
    if fade_reasons:
        lines.append(
            "- unpriced settled rows: "
            + ", ".join(f"`{reason}`={count}" for reason, count in sorted(fade_reasons.items()))
        )
    fade_by_market = fade.get("by_market") or {}
    if fade_by_market:
        lines.extend(
            [
                "",
                "| Market | Settled candidates | Exact quotes | Fade W–L–P | ROI |",
                "| --- | ---: | ---: | ---: | ---: |",
            ]
        )
        for market, group in sorted(fade_by_market.items()):
            roi = f"{group['roi']:+.1%}" if group.get("roi") is not None else "n/a"
            lines.append(
                f"| `{market}` | {group['candidate_sample']} | {group['priced_sample']} | "
                f"{group['wins']}–{group['losses']}–{group['pushes']} | {roi} |"
            )
    fade_by_band = fade.get("by_confidence_band") or {}
    if fade_by_band:
        lines.extend(
            [
                "",
                "Numeric selected-probability bands (calibrated probability, otherwise raw):",
                "",
                "| Probability band | Settled candidates | Exact quotes | Fade W–L–P | ROI |",
                "| --- | ---: | ---: | ---: | ---: |",
            ]
        )
        for band, group in sorted(fade_by_band.items()):
            roi = f"{group['roi']:+.1%}" if group.get("roi") is not None else "n/a"
            lines.append(
                f"| `{band}` | {group['candidate_sample']} | {group['priced_sample']} | "
                f"{group['wins']}–{group['losses']}–{group['pushes']} | {roi} |"
            )

    lines.extend(
        [
            "",
            "## Settled picks — official-final audit",
            "",
            "Grouped by fixture. Every settled logical pick in the window is included; "
            "unpriced grades do not enter ROI.",
            "",
        ]
    )
    if not settled:
        lines.extend(["No settled picks in this window.", ""])

    grouped: Dict[tuple, List[Dict[str, Any]]] = {}
    for row in settled:
        key = (row["date"], row["game_id"], row["home"], row["away"])
        grouped.setdefault(key, []).append(row)
    for (pick_date, game_id, home, away), rows in grouped.items():
        game = (games or {}).get(str(game_id))
        if (
            game is not None
            and game.actual is not None
            and game.actual.home_final is not None
            and game.actual.away_final is not None
        ):
            final = f"Final: **{game.actual.home_final:g}–{game.actual.away_final:g}**"
        else:
            final = "Official final reconciled"
        lines.extend([f"### {pick_date}: {home} vs {away} ({final})", ""])
        for row in rows:
            grade = str(row["grade_status"])
            lines.append(f"- **{_grade_badge(grade)} — {_audit_market_selection(row)}**")
            probability = row.get("calibrated_probability")
            if probability is None:
                probability = row["raw_probability"]
            probability_text = f"**{probability:.1%}**"
            if abs(probability - row["raw_probability"]) >= 0.0005:
                probability_text += f" calibrated (raw {row['raw_probability']:.1%})"
            facts = []
            actual = _audit_actual_result(row, games)
            if actual:
                facts.append(f"Actual {actual}")
            facts.extend(
                [
                    f"Expected: {probability_text}",
                ]
            )
            price = (row.get("book") or {}).get("selected_decimal_price")
            book_display = _canonical_book_for_key((row.get("book") or {}).get("book"))
            price_text = (
                f"{float(price):.3f} · {book_display or 'book unknown'}"
                if price is not None
                else "No exact bookmaker price"
            )
            facts.append(f"Price: **{price_text}**")
            opposite_price, opposite_reason = opposite_quote(row)
            opposite_side = _opposite_selection_for_row(row)
            if opposite_price is None:
                fade_text = f"not priced ({opposite_reason})"
            else:
                fade_result = _grade_badge(
                    {"win": "loss", "loss": "win", "push": "push"}[grade]
                )
                fade_side = opposite_side
                if row.get("market") == "spread" and opposite_side in {"home", "away"}:
                    fade_side = _spread_side_label(row, opposite_side)
                fade_text = (
                    f"{fade_result} @ {opposite_price:.3f} {fade_side} · "
                    f"{book_display or 'book unknown'}"
                )
            facts.append(f"Counterfactual fade: **{fade_text}**")
            lines.append("  - " + " · ".join(facts))
            calibration = row.get("calibration") or {}
            lines.append(
                f"  - Calibration: **n={calibration.get('n', 0)}** · "
                f"{str(calibration.get('status', 'unknown')).replace('_', ' ')} · "
                f"Bucket: `{row['bucket']}` · "
                f"Model: `{row.get('model') or 'unknown'}`"
            )
        lines.append("")

    lines.extend(
        [
            "## Evidence notes",
            "",
            "- Grades use reconciled official finals only; missing or conflicting results stay pending.",
            "- Frozen fixtures use their immutable inline lifecycle snapshot; legacy rows retain dated forecast evidence.",
            "- Flat one-paper-unit ROI is descriptive history, not proof of future performance.",
            "",
        ]
    )
    return "\n".join(lines)


def _audit_rows_for_date(path: Path) -> List[Dict[str, Any]]:
    """Load one date's audit evidence, preferring its frozen lifecycle view.

    Historical ledgers written before the fixture-freeze lifecycle remain
    readable as legacy evidence. Once a lifecycle file exists, using the raw
    dated draft again would allow a later scan to influence pick settlement.
    """

    payload = _read_json(path, [])
    if not isinstance(payload, list):
        raise ValueError(f"dated pick ledger must be a list: {path}")
    target_date = path.stem.removeprefix("picks_")
    lifecycle_path = path.with_name(f"picks_{target_date}_lifecycle.json")
    if not lifecycle_path.exists():
        return payload
    lifecycle = _read_json(lifecycle_path, {})
    rows = lifecycle.get("rows") if isinstance(lifecycle, dict) else None
    if (
        not isinstance(lifecycle, dict)
        or not schema_matches(lifecycle.get("schema_version"), PICK_LIFECYCLE_SCHEMA)
        or lifecycle.get("target_date") != target_date
        or not isinstance(rows, list)
        or any(not isinstance(row, dict) for row in rows)
    ):
        raise ValueError(f"invalid pick lifecycle evidence: {lifecycle_path}")
    return rows


def _write_pick_audit(
    localdata: Path,
    *,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
    games: Dict[str, Game] | None = None,
) -> Dict[str, Any]:
    end = datetime.strptime(target_date, "%Y-%m-%d").date()
    start = end - timedelta(days=29)
    ledger: List[Dict[str, Any]] = []
    for path in _dated_pick_files(localdata):
        ledger.extend(_audit_rows_for_date(path))
    window = [row for row in ledger if start.isoformat() <= row["date"] <= target_date]
    logical_window = _latest_logical_rows(window)
    local_cutoff = parse_utc(generated_utc).astimezone(ZoneInfo(timezone_name)).date().isoformat()
    # Match Edge's local-day pending cutoff, not the acquisition/slate date.
    # Unlike Edge's source-score default, our grades require reconciled official
    # finals, so a same-day final is already safe to include in settled metrics.
    prior = [row for row in logical_window if row["date"] < local_cutoff]
    same_day_pending = [
        row
        for row in logical_window
        if row["date"] >= local_cutoff and row["grade_status"] == "pending"
    ]
    settled = sorted(
        (row for row in logical_window if row["grade_status"] in {"win", "loss", "push"}),
        key=lambda row: (
            row["date"],
            row["scheduled_utc"],
            row["game_id"],
            str(row["pick_id"]),
        ),
        reverse=True,
    )
    same_day_settled = sum(row["date"] >= local_cutoff for row in settled)
    decisions = [row for row in settled if row["grade_status"] != "push"]
    wins = sum(row["grade_status"] == "win" for row in decisions)
    pushes = sum(row["grade_status"] == "push" for row in settled)
    pending = sum(row["grade_status"] == "pending" for row in prior)
    pending_total = sum(row["grade_status"] == "pending" for row in logical_window)
    priced = [
        row for row in settled if (row.get("book") or {}).get("selected_decimal_price") is not None
    ]
    profit = 0.0
    for row in priced:
        price = float(row["book"]["selected_decimal_price"])
        if row["grade_status"] == "win":
            profit += price - 1.0
        elif row["grade_status"] == "loss":
            profit -= 1.0
    by_bucket: Dict[str, Dict[str, Any]] = {}
    by_market: Dict[str, Dict[str, Any]] = {}
    by_selection: Dict[str, Dict[str, Any]] = {}
    for row in settled:
        group = by_bucket.setdefault(
            row["bucket"],
            {"settled": 0, "wins": 0, "pushes": 0, "priced_profit_units": 0.0},
        )
        group["settled"] += 1
        group["wins"] += int(row["grade_status"] == "win")
        group["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                group["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                group["priced_profit_units"] -= 1.0
        # Market breakdown (all leagues) - requested for NCAA M/W watchlist analysis
        m = str(row.get("market") or "unknown")
        mgroup = by_market.setdefault(
            m, {"settled": 0, "wins": 0, "pushes": 0, "priced_profit_units": 0.0}
        )
        mgroup["settled"] += 1
        mgroup["wins"] += int(row["grade_status"] == "win")
        mgroup["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                mgroup["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                mgroup["priced_profit_units"] -= 1.0
        # Over/under + side bias audit — selection within market (over/under, home/away) — same regime, no guards lowered
        sel = str(row.get("selection") or "unknown")
        key = f"{m}:{sel}"
        sgroup = by_selection.setdefault(
            key,
            {
                "settled": 0,
                "wins": 0,
                "pushes": 0,
                "priced_profit_units": 0.0,
                "market": m,
                "selection": sel,
            },
        )
        sgroup["settled"] += 1
        sgroup["wins"] += int(row["grade_status"] == "win")
        sgroup["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                sgroup["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                sgroup["priced_profit_units"] -= 1.0

    fade_audit = {
        "schema_version": FADE_AUDIT_SCHEMA,
        "paper_only": True,
        "evidence_policy": (
            "use only explicit opposite-side prices preserved in the pick's "
            "same BookSnapshot; never infer or historically backfill quotes"
        ),
        "selected_price_required": False,
        "unit_stake": 1.0,
        "profit_basis": (
            "fade win earns opposite decimal price minus one; fade loss loses one; "
            "push returns stake; ROI divides by one unit per priced fade, including pushes"
        ),
        "break_even_basis": (
            "mean inverse opposite decimal price among non-push fade quotes; "
            "book-implied and includes market margin"
        ),
        "confidence_probability_basis": (
            "numeric calibrated_probability, falling back to raw_probability; "
            "calibration status is not used"
        ),
        **_fade_metrics(settled),
        "by_market": _fade_breakdown(
            settled, lambda row: row.get("market") or "unknown"
        ),
        "by_selection": _fade_breakdown(
            settled,
            lambda row: f"{row.get('market') or 'unknown'}:{row.get('selection') or 'unknown'}",
        ),
        "by_confidence_band": _fade_breakdown(settled, _confidence_band),
        "by_market_period_model_confidence_band": _fade_breakdown(
            settled,
            lambda row: "|".join(
                (
                    str(row.get("market") or "unknown"),
                    str(row.get("period") or "unknown"),
                    str(row.get("model") or "unknown"),
                    _confidence_band(row),
                )
            ),
        ),
    }
    report = {
        "schema_version": OPERATOR_SCHEMA,
        "generated_utc": generated_utc,
        "window": {"start": start.isoformat(), "end": target_date, "days": 30},
        "archived_pick_rows": len(logical_window),
        "archived_evidence_rows": len(window),
        "superseded_observations": len(window) - len(logical_window),
        "same_day_cutoff": local_cutoff,
        "include_same_day_settled": True,
        "same_day_rows_excluded": len(same_day_pending),
        "same_day_pending_picks": len(same_day_pending),
        "same_day_settled_picks": same_day_settled,
        "eligible_prior_picks": len(prior),
        "settled_picks": len(settled),
        "pending_picks": pending,
        "pending_picks_total": pending_total,
        "wins": wins,
        "losses": len(decisions) - wins,
        "pushes": pushes,
        "hit_rate": wins / len(decisions) if decisions else None,
        "priced_sample": len(priced),
        "roi": profit / len(priced) if priced else None,
        "profit_units": profit,
        "stake_policy": "flat_one_paper_unit_per_priced_pick",
        "fade_audit": fade_audit,
        "by_bucket": by_bucket,
        "by_market": by_market,
        "by_selection": by_selection,
        "grading_basis": (
            "reconciled_official_finals_only; frozen_fixture_snapshots_precede_drafts; "
            "official_finals_settle_immediately; pending_is_not_a_loss; "
            "local_day_pending_cutoff; latest_observation_per_unfrozen_logical_pick"
        ),
        "ledger": logical_window,
        "settled_ledger": settled,
    }
    _write_json(localdata / "picks_audit_rolling.json", report)
    _write_text(
        localdata / f"picks_audit_{target_date}.md",
        _render_pick_audit_markdown(report, games=games, timezone_name=timezone_name),
    )
    return report


def _compact_scan_record(record: Dict[str, Any]) -> Dict[str, Any]:
    """Drop row-level duplicates from the bounded, Git-persisted scan history.

    Full candidate evidence remains in the dated/rolling picks reports and the
    run artifact. The ticket state only needs compact counts/policies plus the
    immutable date locks; copying every candidate into every scan made this
    file grow to multiple megabytes per league.
    """
    compact = deepcopy(record)
    for key in ("unpriced_audit", "priced_ineligible_shadow", "unpriced_spread_shadow"):
        summary = compact.get(key)
        if isinstance(summary, dict):
            summary.pop("candidate_records", None)
            summary.pop("forecast_ids", None)
    funnel = compact.get("candidate_funnel")
    if isinstance(funnel, dict):
        pick_filter = funnel.get("paper_pick_filter")
        if isinstance(pick_filter, dict):
            pick_filter.pop("candidate_band_details", None)
            pick_filter.pop("bands", None)
    return compact


def _load_state(localdata: Path) -> Dict[str, Any]:
    state = _read_json(
        localdata / "auto_tickets_state.json",
        {
            "schema_version": STATE_SCHEMA,
            "base_paper_units": BASE_PAPER_UNITS,
            "stake_units_per_acca": STAKE_UNITS_PER_ACCA,
            "frozen_by_date": {},
            "scans": [],
        },
    )
    if not isinstance(state, dict) or not schema_matches(
        state.get("schema_version"), STATE_SCHEMA
    ):
        raise ValueError("unsupported or corrupt auto ticket state; refusing to reset it")
    state["schema_version"] = STATE_SCHEMA  # mutable state adopts the current tag
    state.setdefault("frozen_by_date", {})
    state.setdefault("scans", [])
    if isinstance(state["scans"], list):
        state["scans"] = [
            _compact_scan_record(scan) if isinstance(scan, dict) else scan
            for scan in state["scans"]
        ]
    return state


def _single_ticket_candidate(
    pick: Dict[str, Any], paper_filter: Dict[str, Any] | None
) -> Optional[Dict[str, Any]]:
    """Adapt exactly one ROI-qualified pick to the frozen ticket contract.

    The candidate must already come from ``paper_single_picks``. Rechecking
    its positive ROI-band decision, exact price, calibrated probability, and
    strict-or-narrow-provisional eligibility here prevents this adapter from
    becoming a second path around the service's existing gates.
    """
    if not isinstance(pick, dict):
        return None
    forecast_id = pick.get("forecast_id")
    if not isinstance(forecast_id, str) or not forecast_id.strip():
        return None
    band = pick.get("individual_pick_roi_band")
    if not isinstance(band, dict):
        details = (
            paper_filter.get("candidate_band_details") or {}
            if isinstance(paper_filter, dict)
            else {}
        )
        band = details.get(forecast_id) if isinstance(details, dict) else None
    if not isinstance(band, dict) or band.get("qualifies") is not True:
        return None

    eligibility = pick.get("eligibility") or {}
    if not isinstance(eligibility, dict):
        return None
    provisional = _is_provisional_sample_only_row(pick)
    if eligibility.get("acca_eligible") is not True and not provisional:
        return None

    book = pick.get("book") or {}
    raw_price = book.get("selected_decimal_price") if isinstance(book, dict) else None
    raw_probability = pick.get("calibrated_probability")
    if (
        isinstance(raw_price, bool)
        or not isinstance(raw_price, (int, float))
        or isinstance(raw_probability, bool)
        or not isinstance(raw_probability, (int, float))
    ):
        return None
    try:
        price = float(raw_price)
        probability = float(raw_probability)
    except OverflowError:
        return None
    if not math.isfinite(price) or price <= 1.0:
        return None
    if not math.isfinite(probability) or not 0.0 < probability <= 1.0:
        return None

    ticket_id = hashlib.sha256(f"paper-single|{forecast_id}".encode()).hexdigest()[:20]
    return {
        "ticket_id": ticket_id,
        "acca_id": ticket_id,
        "ticket_type": "single",
        "legs": [pick],
        "leg_count": 1,
        "combined_decimal_odds": price,
        "estimated_win_probability": probability,
        "estimated_joint_win_probability": probability,
        "provisional_sample_gate_bypassed": provisional,
        "provisional_leg_count": int(provisional),
        "ranking_status": (
            "paper_only_provisional_sample_gate_bypassed"
            if provisional
            else "paper_only_roi_qualified_single"
        ),
        "individual_pick_filter": deepcopy(paper_filter)
        if isinstance(paper_filter, dict)
        else None,
    }


def _new_receipt(
    *,
    league: League,
    target_date: str,
    captured_utc: str,
    run_id: str,
    accas: List[Dict[str, Any]],
    ticket_decision: Dict[str, Any],
) -> Dict[str, Any]:
    ticket_rows: Dict[str, List[Dict[str, Any]]] = {"acca": [], "single": []}
    for index, acca in enumerate(accas, 1):
        ticket_type = "single" if acca.get("ticket_type") == "single" else "acca"
        legs = [_pick_row(leg, target_date) for leg in acca["legs"]]
        books = sorted({(leg.get("book") or {}).get("book") for leg in legs} - {None})
        ticket_row = {
            "ticket_id": acca.get("ticket_id") or acca["acca_id"],
            "ticket_type": ticket_type,
            "publication_index": index,
            "leg_count": int(acca.get("leg_count") or len(legs)),
            "combined_decimal_odds": acca["combined_decimal_odds"],
            "stake_units": STAKE_UNITS_PER_ACCA,
            "estimated_joint_win_probability": acca["estimated_joint_win_probability"],
            "paper_only": True,
            "provisional_sample_gate_bypassed": bool(
                acca.get("provisional_sample_gate_bypassed", False)
            ),
            "provisional_leg_count": int(acca.get("provisional_leg_count") or 0),
            "ranking_status": acca.get("ranking_status"),
            "bookmakers": books,
            "legs": legs,
        }
        if ticket_type == "single":
            ticket_row["estimated_win_probability"] = acca.get(
                "estimated_win_probability", acca["estimated_joint_win_probability"]
            )
        pick_filter = acca.get("individual_pick_filter")
        if isinstance(pick_filter, dict):
            ticket_row["individual_pick_filter"] = deepcopy(pick_filter)
        ticket_rows[ticket_type].append(ticket_row)
    return {
        "schema_version": TICKET_SCHEMA,
        "date": target_date,
        "league": league.value,
        "decision_id": run_id,
        "status": "frozen",
        "paper_only": True,
        "captured_at": captured_utc,
        "frozen_at": captured_utc,
        "ticket_decision": deepcopy(ticket_decision),
        "stake_units_per_acca": STAKE_UNITS_PER_ACCA,
        "stake_units_per_position": STAKE_UNITS_PER_ACCA,
        "accas": ticket_rows["acca"],
        "singles": ticket_rows["single"],
    }


def _freeze_no_bet_decision(
    localdata: Path,
    *,
    state: Dict[str, Any],
    league: League,
    target_date: str,
    captured_utc: str,
    run_id: str,
    ticket_decision: Dict[str, Any],
) -> Dict[str, Any]:
    """Store the immutable terminal no-bet decision in the durable state lock."""
    existing = _read_no_bet_decision(state, target_date)
    if existing is not None:
        return existing
    reason = (
        "decision_window_missed_before_publication"
        if ticket_decision.get("status") == "missed_pre_tip_window"
        else "no_qualifying_pre_tip_ticket_at_decision_cutoff"
    )
    decision = {
        "schema_version": NO_BET_DECISION_SCHEMA,
        "status": "frozen_no_bet",
        "league": league.value,
        "target_date": target_date,
        "decision_id": run_id,
        "frozen_at": captured_utc,
        "reason": reason,
        "ticket_decision": deepcopy(ticket_decision),
    }
    state["frozen_by_date"][target_date] = {
        "decision_id": run_id,
        "frozen_at": captured_utc,
        "final": True,
        "no_bet": True,
        "lock_source": "inline_state_no_bet",
        "decision_cutoff_utc": ticket_decision.get("planned_cutoff_utc"),
        "no_bet_decision": decision,
    }
    # Persist the date lock immediately so an interrupted run cannot reopen a
    # reached cutoff before the normal end-of-run state refresh.
    _write_json(localdata / "auto_tickets_state.json", state)
    return decision


def _ticket_receipt(
    localdata: Path,
    *,
    state: Dict[str, Any],
    league: League,
    target_date: str,
    captured_utc: str,
    run_id: str,
    accas: List[Dict[str, Any]],
    ticket_decision: Dict[str, Any],
    allow_new_decision: bool,
) -> Optional[Dict[str, Any]]:
    """Load a prior receipt or make exactly one post-cutoff decision.

    Before the cutoff a scan is only candidate evidence. At/after the cutoff,
    the first scan either freezes its pre-tip ticket candidates or writes a
    durable no-bet decision. Neither decision can later be upgraded by a newly
    arriving bookmaker price.
    """
    path = localdata / f"auto_tickets_{target_date}_decision.json"
    # Sidecar is authoritative even if state cache rolls back.
    if _withdrawn_sidecar_path(localdata, target_date).exists():
        return None
    no_bet = _read_no_bet_decision(state, target_date)
    if no_bet is not None:
        if path.exists():
            raise ValueError(
                f"conflicting frozen ticket receipt and no-bet decision for {target_date}"
            )
        return None

    lock = state["frozen_by_date"].get(target_date)
    if isinstance(lock, dict) and lock.get("withdrawn") is True:
        # Keep immutable receipt bytes for evidence, but never adopt, rotate,
        # settle, or replace a withdrawn publication. This also makes the
        # withdrawal survive restoration of an older Actions cache.
        if not lock.get("decision_id"):
            raise ValueError(f"withdrawn ticket lock lacks decision_id: {target_date}")
        # Persist sidecar so future state rollbacks still see the withdrawal.
        sidecar = _withdrawn_sidecar_path(localdata, target_date)
        if not sidecar.exists():
            _write_json(
                sidecar,
                {
                    "decision_id": lock.get("decision_id"),
                    "withdrawn": True,
                    "target_date": target_date,
                },
            )
        return None
    if isinstance(lock, dict) and lock.get("no_bet") is True:
        # `_read_no_bet_decision` above validates the inline immutable lock.
        return None
    if lock is not None:
        locked_path = localdata / str(lock.get("receipt_path") or path.name)
        if not locked_path.exists():
            raise FileNotFoundError(
                f"frozen ticket receipt is missing: {locked_path}; refusing to rotate the card"
            )
        receipt = _read_json(locked_path, {})
        if not schema_matches(receipt.get("schema_version"), TICKET_SCHEMA):
            raise ValueError(f"invalid frozen ticket receipt: {locked_path}")
        return receipt
    if path.exists():
        # Do not adopt a withdrawn receipt even via immutable-receipt path.
        sidecar_ids = _withdrawn_ids_from_sidecars(localdata)
        if sidecar_ids:
            try:
                cand = json.loads(path.read_text())
                if str(cand.get("decision_id")) in sidecar_ids:
                    return None
            except (OSError, json.JSONDecodeError):
                pass
        # A receipt is itself immutable evidence. Adopt its lock rather than
        # rebuilding a potentially different card after a state-cache loss.
        receipt = _read_json(path, {})
        if not schema_matches(receipt.get("schema_version"), TICKET_SCHEMA):
            raise ValueError(f"invalid frozen ticket receipt: {path}")
        state["frozen_by_date"][target_date] = {
            "frozen_at": receipt["frozen_at"],
            "final": True,
            "lock_source": "immutable_receipt",
            "receipt_path": path.name,
        }
        return receipt
    if not allow_new_decision:
        return None
    if not accas:
        _freeze_no_bet_decision(
            localdata,
            state=state,
            league=league,
            target_date=target_date,
            captured_utc=captured_utc,
            run_id=run_id,
            ticket_decision=ticket_decision,
        )
        return None
    receipt = _new_receipt(
        league=league,
        target_date=target_date,
        captured_utc=captured_utc,
        run_id=run_id,
        accas=accas,
        ticket_decision=ticket_decision,
    )
    _write_json(path, receipt)
    state["frozen_by_date"][target_date] = {
        "frozen_at": captured_utc,
        "final": True,
        "lock_source": "state",
        "receipt_path": path.name,
        "decision_cutoff_utc": ticket_decision.get("planned_cutoff_utc"),
    }
    return receipt


def _receipt_positions(
    receipt: Dict[str, Any],
) -> List[tuple[str, Dict[str, Any]]]:
    """Return positions from current receipts and legacy acca-only receipts."""
    positions: List[tuple[str, Dict[str, Any]]] = []
    for collection, default_type in (("accas", "acca"), ("singles", "single")):
        rows = receipt.get(collection) or []
        if not isinstance(rows, list):
            continue
        for row in rows:
            if not isinstance(row, dict):
                continue
            ticket_type = row.get("ticket_type")
            if ticket_type not in {"acca", "single"}:
                ticket_type = default_type
            positions.append((ticket_type, row))
    return positions


def _pending_result_status(leg: Dict[str, Any], games: Dict[str, Game]) -> str:
    """Explain whether an ungraded leg awaits play or lacks a resolvable result."""
    game = games.get(str(leg["game_id"]))
    if game is not None and game.status in {
        GameStatus.SCHEDULED,
        GameStatus.LIVE,
        GameStatus.POSTPONED,
    }:
        return "awaiting_game"
    # A missing identity, non-playable status, or unreconciled FINAL must not be
    # reported as ordinary wait time; it needs result/identity review.
    return "result_unresolvable"


def _grade_receipt(receipt: Dict[str, Any], games: Dict[str, Game]) -> Dict[str, Any]:
    graded = deepcopy(receipt)
    for ticket_type, position in _receipt_positions(graded):
        position.setdefault("ticket_type", ticket_type)
        outcomes: List[Optional[str]] = []
        for leg in position["legs"]:
            _set_grade(leg, games)
            if leg["outcome"] is None:
                leg["result_resolution_status"] = _pending_result_status(leg, games)
            else:
                leg.pop("result_resolution_status", None)
            outcomes.append(leg["outcome"])
        position["leg_outcomes"] = outcomes
        position["push_legs"] = sum(outcome == "push" for outcome in outcomes)
        position["unresolved_legs"] = sum(outcome is None for outcome in outcomes)
        position["awaiting_game_legs"] = sum(
            leg.get("result_resolution_status") == "awaiting_game"
            for leg in position["legs"]
        )
        position["result_unresolvable_legs"] = sum(
            leg.get("result_resolution_status") == "result_unresolvable"
            for leg in position["legs"]
        )
        stake = float(position["stake_units"])
        effective_odds = math.prod(
            float((leg.get("book") or {})["selected_decimal_price"])
            for leg in position["legs"]
            if leg["outcome"] != "push"
        )
        position["effective_decimal_odds"] = effective_odds
        if any(outcome == "loss" for outcome in outcomes):
            status = "loss"
            profit = -stake
        elif any(outcome is None for outcome in outcomes):
            status = "pending"
            profit = None
        elif outcomes and all(outcome == "push" for outcome in outcomes):
            status = "push"
            profit = 0.0
        else:
            status = "win"
            profit = stake * (effective_odds - 1.0)
        position["grade_status"] = status
        position["profit_units"] = profit
    return graded


def _performance(receipts: List[Dict[str, Any]], games: Dict[str, Game]) -> Dict[str, Any]:
    history = [_grade_receipt(receipt, games) for receipt in receipts]
    positions = [
        position
        for receipt in history
        for _, position in _receipt_positions(receipt)
    ]
    settled = [position for position in positions if position["profit_units"] is not None]
    pending = [position for position in positions if position["profit_units"] is None]
    settled_accas = [position for position in settled if position["ticket_type"] == "acca"]
    settled_singles = [position for position in settled if position["ticket_type"] == "single"]
    pending_accas = [position for position in pending if position["ticket_type"] == "acca"]
    pending_singles = [position for position in pending if position["ticket_type"] == "single"]
    profit = sum(float(position["profit_units"]) for position in settled)
    staked = sum(float(position["stake_units"]) for position in settled)
    committed = sum(float(position["stake_units"]) for position in pending)
    wins = sum(position["grade_status"] == "win" for position in settled)
    losses = sum(position["grade_status"] == "loss" for position in settled)
    pushes = sum(position["grade_status"] == "push" for position in settled)
    unresolved_legs = sum(position["unresolved_legs"] for position in positions)
    brier = []
    seen_legs = set()
    for receipt in history:
        for _, position in _receipt_positions(receipt):
            for leg in position["legs"]:
                identity = (receipt["date"], leg["pick_id"])
                if (
                    identity in seen_legs
                    or leg["outcome"] not in {"win", "loss"}
                    # A source moneyline probability is deliberately not
                    # scored as a spread-cover probability.
                    or _is_derived_spread_ticket_row(leg)
                ):
                    continue
                seen_legs.add(identity)
                probability = leg.get("calibrated_probability")
                if probability is None:
                    probability = leg["raw_probability"]
                brier.append((probability - float(leg["outcome"] == "win")) ** 2)
    days = []
    equity = BASE_PAPER_UNITS
    for receipt in sorted(history, key=lambda row: row["date"]):
        daily = [position for _, position in _receipt_positions(receipt)]
        daily_accas = [position for position in daily if position["ticket_type"] == "acca"]
        daily_singles = [position for position in daily if position["ticket_type"] == "single"]
        daily_settled = [position for position in daily if position["profit_units"] is not None]
        daily_pending = [position for position in daily if position["profit_units"] is None]
        daily_stake = sum(float(position["stake_units"]) for position in daily_settled)
        daily_profit = sum(float(position["profit_units"]) for position in daily_settled)
        daily_pending_stake = sum(float(position["stake_units"]) for position in daily_pending)
        equity += daily_profit
        days.append(
            {
                "date": receipt["date"],
                "league": receipt["league"],
                "decision_id": receipt["decision_id"],
                "frozen_at": receipt["frozen_at"],
                "wins": sum(position["grade_status"] == "win" for position in daily_settled),
                "losses": sum(position["grade_status"] == "loss" for position in daily_settled),
                "pushes": sum(position["grade_status"] == "push" for position in daily_settled),
                "settled_positions": len(daily_settled),
                "pending_positions": len(daily_pending),
                "unresolved_result_legs": sum(
                    position["unresolved_legs"] for position in daily
                ),
                "awaiting_game_legs": sum(
                    position["awaiting_game_legs"] for position in daily
                ),
                "result_unresolvable_legs": sum(
                    position["result_unresolvable_legs"] for position in daily
                ),
                "settled_accas": sum(position["profit_units"] is not None for position in daily_accas),
                "pending_accas": sum(position["profit_units"] is None for position in daily_accas),
                "settled_singles": sum(position["profit_units"] is not None for position in daily_singles),
                "pending_singles": sum(position["profit_units"] is None for position in daily_singles),
                "staked_units": daily_stake,
                "returned_units": daily_stake + daily_profit if daily_settled else None,
                "profit_units": daily_profit if daily_settled else None,
                "pending_units": daily_pending_stake,
                "bank_units": equity,
                "positions": daily,
                "accas": daily_accas,
                "singles": daily_singles,
            }
        )
    return {
        "schema_version": OPERATOR_SCHEMA,
        "paper_only": True,
        "base_paper_units": BASE_PAPER_UNITS,
        "paper_bank_units": BASE_PAPER_UNITS + profit,
        "paper_free_units": BASE_PAPER_UNITS + profit - committed,
        "paper_committed_units": committed,
        "settled_stake_units": staked,
        "settled_positions": len(settled),
        "pending_positions": len(pending),
        "settled_accas": len(settled_accas),
        "settled_singles": len(settled_singles),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "hit_rate": wins / (wins + losses) if wins + losses else None,
        "pending_accas": len(pending_accas),
        "pending_singles": len(pending_singles),
        "unresolved_result_legs": unresolved_legs,
        "awaiting_game_legs": sum(position["awaiting_game_legs"] for position in positions),
        "result_unresolvable_legs": sum(
            position["result_unresolvable_legs"] for position in positions
        ),
        "profit_units": profit,
        "roi": profit / staked if staked else None,
        "selected_leg_brier_score": sum(brier) / len(brier) if brier else None,
        "stake_policy": "flat_one_paper_unit_per_position_no_refill_no_real_money",
        "drawdown_basis": "not_a_settlement_time_equity_curve",
        "history": days,
    }


def _compact_performance_report(report: Dict[str, Any]) -> Dict[str, Any]:
    """Persist settlement summaries without cloning full receipt evidence.

    Receipts are already immutable per-date sources for every full leg,
    prediction, calibration, and bookmaker payload. Repeating those large
    payloads in a rolling Git-tracked performance report made the file grow by
    hundreds of kilobytes for a handful of positions. The renderer keeps the
    rich in-memory report; this serializer keeps only fields needed to inspect
    financial settlement and points readers to the receipt for full evidence.
    """
    compact = {key: value for key, value in report.items() if key != "history"}
    compact["history_detail"] = (
        "compact_settlement_summary; immutable auto_tickets_YYYY-MM-DD_decision.json "
        "receipts retain complete leg evidence"
    )
    compact_history = []
    day_fields = (
        "date",
        "league",
        "decision_id",
        "frozen_at",
        "wins",
        "losses",
        "pushes",
        "settled_positions",
        "pending_positions",
        "unresolved_result_legs",
        "awaiting_game_legs",
        "result_unresolvable_legs",
        "settled_accas",
        "pending_accas",
        "settled_singles",
        "pending_singles",
        "staked_units",
        "returned_units",
        "profit_units",
        "pending_units",
        "bank_units",
    )
    position_fields = (
        "ticket_id",
        "ticket_type",
        "publication_index",
        "leg_count",
        "combined_decimal_odds",
        "effective_decimal_odds",
        "stake_units",
        "estimated_joint_win_probability",
        "estimated_win_probability",
        "paper_only",
        "provisional_sample_gate_bypassed",
        "provisional_leg_count",
        "ranking_status",
        "bookmakers",
        "leg_outcomes",
        "push_legs",
        "unresolved_legs",
        "awaiting_game_legs",
        "result_unresolvable_legs",
        "grade_status",
        "profit_units",
    )
    leg_fields = (
        "game_id",
        "pick_id",
        "forecast_id",
        "home",
        "away",
        "scheduled_utc",
        "market",
        "period",
        "selection",
        "line",
        "grade_status",
        "outcome",
        "result_resolution_status",
    )
    for day in report.get("history") or []:
        compact_day = {field: day.get(field) for field in day_fields}
        compact_day["receipt_path"] = f"auto_tickets_{day['date']}_decision.json"
        compact_positions = []
        for position in day.get("positions") or []:
            compact_position = {
                field: position.get(field)
                for field in position_fields
                if position.get(field) is not None
            }
            compact_legs = []
            for leg in position.get("legs") or []:
                compact_leg = {field: leg.get(field) for field in leg_fields if leg.get(field) is not None}
                book = leg.get("book") or {}
                if isinstance(book, dict):
                    price = book.get("selected_decimal_price")
                    bookmaker = book.get("book")
                    provider = book.get("provider")
                    if price is not None:
                        compact_leg["selected_decimal_price"] = price
                    if bookmaker:
                        compact_leg["book"] = bookmaker
                    if provider:
                        compact_leg["book_provider"] = provider
                compact_legs.append(compact_leg)
            compact_position["legs"] = compact_legs
            compact_positions.append(compact_position)
        compact_day["positions"] = compact_positions
        compact_history.append(compact_day)
    compact["history"] = compact_history
    return compact


def _calibrated_probability_band(row: Dict[str, Any]) -> Optional[str]:
    """Return the historical-comparison band without falling back to raw p."""
    value = row.get("calibrated_probability")
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    try:
        probability = float(value)
    except (OverflowError, TypeError, ValueError):
        return None
    if not math.isfinite(probability) or not 0.0 <= probability <= 1.0:
        return None
    return _confidence_band({"calibrated_probability": probability})


def _unpriced_historical_success_rate(
    row: Dict[str, Any], settled_history: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Measure prior comparable win rate; no odds or current pick result used."""
    probability_band = _calibrated_probability_band(row)
    scope = {
        "league": row.get("league"),
        "market": row.get("market"),
        "period": row.get("period"),
        "model": row.get("model"),
        "calibrated_probability_band": probability_band,
    }
    result = {
        "threshold": UNPRICED_SUCCESS_RATE_THRESHOLD,
        "minimum_decisions_exclusive": UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE,
        "history_window_days": UNPRICED_SUCCESS_RATE_HISTORY_DAYS,
        "result_available_after_hours": UNPRICED_SUCCESS_RATE_RESULT_DELAY_HOURS,
        "comparison_scope": scope,
        "wins": 0,
        "losses": 0,
        "decisions": 0,
        "rate": None,
        "qualifies": False,
        "status": "insufficient_history",
    }
    if probability_band is None:
        result["status"] = "missing_calibrated_probability"
        return result
    if any(
        not value or str(value).strip().lower() == "unknown"
        for key, value in scope.items()
        if key != "calibrated_probability_band"
    ):
        result["status"] = "missing_comparison_scope"
        return result
    prediction = row.get("prediction")
    observed_utc = prediction.get("observed_utc") if isinstance(prediction, dict) else None
    try:
        history_cutoff = (
            parse_utc(str(observed_utc))
            - timedelta(hours=UNPRICED_SUCCESS_RATE_RESULT_DELAY_HOURS)
        )
    except (TypeError, ValueError):
        result["status"] = "missing_forecast_as_of"
        return result
    candidate_game_id = str(row.get("game_id") or "")
    if not candidate_game_id:
        result["status"] = "missing_game_identity"
        return result

    wins = losses = 0
    seen_historical_picks = set()
    for historical in settled_history:
        if not isinstance(historical, dict):
            continue
        outcome = historical.get("grade_status")
        if outcome not in {"win", "loss"}:
            outcome = historical.get("outcome")
        outcome = str(outcome or "").lower()
        if outcome not in {"win", "loss"}:
            # Pushes, voids, pending rows, and unknown outcomes are not hit-rate
            # decisions. No-price historical picks can still count once settled.
            continue
        historical_game_id = str(historical.get("game_id") or "")
        if not historical_game_id or historical_game_id == candidate_game_id:
            continue
        if any(
            str(historical.get(key) or "") != str(scope[key])
            for key in ("league", "market", "period", "model")
        ):
            continue
        if _calibrated_probability_band(historical) != probability_band:
            continue
        try:
            historical_scheduled = parse_utc(str(historical["scheduled_utc"]))
        except (KeyError, TypeError, ValueError):
            continue
        if historical_scheduled > history_cutoff:
            continue
        historical_selection = str(historical.get("selection") or "")
        if not historical_selection:
            continue
        raw_line = historical.get("line")
        if raw_line is None:
            line = None
        elif isinstance(raw_line, bool):
            continue
        else:
            try:
                line = float(raw_line)
            except (OverflowError, TypeError, ValueError):
                continue
            if not math.isfinite(line):
                continue
        wager_key = (
            historical_game_id,
            str(historical.get("market") or ""),
            str(historical.get("period") or ""),
            historical_selection,
            line,
            str(historical.get("model") or ""),
        )
        if wager_key in seen_historical_picks:
            continue
        seen_historical_picks.add(wager_key)
        if outcome == "win":
            wins += 1
        else:
            losses += 1

    decisions = wins + losses
    rate = wins / decisions if decisions else None
    qualifies = (
        decisions > UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE
        and rate is not None
        and rate >= UNPRICED_SUCCESS_RATE_THRESHOLD
    )
    result.update(
        {
            "wins": wins,
            "losses": losses,
            "decisions": decisions,
            "rate": round(rate, 6) if rate is not None else None,
            "qualifies": qualifies,
            "status": (
                "qualified"
                if qualifies
                else "insufficient_history"
                if decisions <= UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE
                else "below_threshold"
            ),
        }
    )
    return result


def _annotate_unpriced_audit_rows(
    rows: List[Dict[str, Any]], settled_history: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Attach a non-persistent gate result for reporting and card rendering."""
    annotated = []
    for row in rows:
        candidate = dict(row)
        candidate["_historical_success_rate"] = _unpriced_historical_success_rate(
            row, settled_history
        )
        annotated.append(candidate)
    return annotated


def _unpriced_candidate_record(
    row: Dict[str, Any], success_rate: Dict[str, Any]
) -> Dict[str, Any]:
    """One audit record for one no-exact-price forecast; never a wager position.

    ``candidate_id`` is the immutable forecast/evidence ID. The logical identity
    is descriptive and mirrors ``_logical_pick_key``; it does not collapse this
    record with another candidate. In particular, markets, periods, selections,
    lines, models, and books remain separately inspectable.
    """
    book = row.get("book") if isinstance(row.get("book"), dict) else {}
    forecast_id = str(row.get("forecast_id") or row.get("pick_id") or "")
    probability = row.get("calibrated_probability")
    if probability is None:
        probability = row.get("raw_probability")

    def finite_number(value):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return None
        try:
            number = float(value)
        except (OverflowError, TypeError, ValueError):
            return None
        return number if math.isfinite(number) else None

    probability = finite_number(probability)
    logical = _logical_pick_key(row)
    eligibility = row.get("eligibility") if isinstance(row.get("eligibility"), dict) else {}
    calibration = row.get("calibration") if isinstance(row.get("calibration"), dict) else {}
    return {
        "record_type": "unpriced_forecast_candidate",
        "candidate_id": forecast_id,
        "forecast_id": forecast_id or None,
        "pick_id": str(row.get("pick_id") or forecast_id) or None,
        "date": str(row.get("date") or "")[:10] or None,
        "game_id": row.get("game_id"),
        "home": row.get("home"),
        "away": row.get("away"),
        "scheduled_utc": row.get("scheduled_utc"),
        "market": row.get("market"),
        "period": row.get("period"),
        "selection": row.get("selection"),
        "line": logical[4],
        "model": row.get("model"),
        "model_version": row.get("model_version"),
        "bucket": row.get("bucket"),
        "acca_eligible": eligibility.get("acca_eligible"),
        "calibration_status": calibration.get("status"),
        "calibration_n": calibration.get("n"),
        "logical_pick_identity": {
            "game_id": logical[0],
            "market": logical[1],
            "period": logical[2],
            "selection": logical[3],
            "line": logical[4],
            "model": logical[5],
            "book_provider": logical[6],
            "book": logical[7],
        },
        "probability": probability,
        "raw_probability": finite_number(row.get("raw_probability")),
        "calibrated_probability": finite_number(row.get("calibrated_probability")),
        "price": None,
        "price_evidence": str(row.get("price_evidence") or _price_evidence(row)),
        "price_state": (
            "unmatched" if not book
            else "selected_side_quote_missing"
            if book.get("selected_decimal_price") is None
            else "selected_price_unusable"
        ),
        "book_provider": book.get("provider"),
        "book": _canonical_book_for_key(book.get("book")),
        "book_observed_utc": book.get("observed_utc"),
        "ticket_id": None,
        "ticket_status": "audit_only_no_exact_selected_price",
        "ticketable": False,
        "audit_only": True,
        "stake_units": 0.0,
        "grade_status": str(row.get("grade_status") or "pending"),
        "outcome": row.get("outcome"),
        "historical_success_rate": success_rate.get("rate"),
        "historical_success_wins": success_rate.get("wins", 0),
        "historical_success_losses": success_rate.get("losses", 0),
        "historical_success_decisions": success_rate.get("decisions", 0),
        "historical_success_threshold": success_rate.get("threshold"),
        "historical_success_minimum_decisions_exclusive": success_rate.get(
            "minimum_decisions_exclusive"
        ),
        "historical_success_qualifies": success_rate.get("qualifies") is True,
        "historical_success_status": success_rate.get("status"),
        "historical_success_comparison_scope": success_rate.get("comparison_scope"),
        "eligibility_reasons": list(eligibility.get("reasons") or []),
    }


def _unpriced_audit_summary(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Summarize no-price audit candidates while retaining one record per pick."""
    grade_counts = Counter(str(row.get("grade_status") or "pending") for row in rows)
    evidence_counts = Counter(
        str(row.get("price_evidence") or _price_evidence(row)) for row in rows
    )
    reason_counts = Counter(
        str(reason)
        for row in rows
        for reason in ((row.get("eligibility") or {}).get("reasons") or [])
    )
    records = []
    for row in rows:
        success_rate = row.get("_historical_success_rate")
        if not isinstance(success_rate, dict):
            success_rate = _unpriced_historical_success_rate(row, [])
        records.append(_unpriced_candidate_record(row, success_rate))
    records.sort(key=lambda item: (
        str(item.get("date") or ""),
        str(item.get("game_id") or ""),
        str(item.get("market") or ""),
        str(item.get("period") or ""),
        str(item.get("selection") or ""),
        str(item.get("line") or ""),
        str(item.get("model") or ""),
        str(item.get("book") or ""),
        str(item.get("candidate_id") or ""),
    ))
    success_status_counts = Counter(
        str(record.get("historical_success_status") or "unknown")
        for record in records
    )
    success_qualified_count = sum(
        record.get("historical_success_qualifies") is True for record in records
    )
    return {
        "schema_version": UNPRICED_AUDIT_SCHEMA,
        "record_type": "unpriced_shadow_ticket_audit",
        "audit_only": True,
        "ticketable": False,
        "stake_units": 0.0,
        "count": len(rows),
        "historical_success_qualified_count": success_qualified_count,
        "historical_success_excluded_count": len(rows) - success_qualified_count,
        "historical_success_status_counts": dict(sorted(success_status_counts.items())),
        "historical_success_policy": {
            "metric": "wins / (wins + losses)",
            "threshold_at_least": UNPRICED_SUCCESS_RATE_THRESHOLD,
            "minimum_decisions_exclusive": UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE,
            "history_window_days": UNPRICED_SUCCESS_RATE_HISTORY_DAYS,
            "result_available_after_hours": UNPRICED_SUCCESS_RATE_RESULT_DELAY_HOURS,
            "comparison_keys": [
                "league", "market", "period", "model", "calibrated_probability_band",
            ],
            "odds_required": False,
            "excluded_outcomes": ["push", "void", "pending", "unknown"],
            "same_game_history_excluded": True,
            "historical_deduplication_key": [
                "game_id", "market", "period", "selection", "line", "model",
            ],
            "history_source": "rolling_pick_audit.settled_ledger",
        },
        "shadow_group_policy": {
            "display_only": True,
            "max_picks_per_group": UNPRICED_SHADOW_GROUP_MAX_PICKS,
            "one_pick_per_game": True,
            "each_qualified_candidate_once": True,
            "joint_success_rate_calculated": False,
            "ticket_id": None,
            "receipt_created": False,
            "stake_units": 0.0,
        },
        "forecast_ids": sorted(
            str(row.get("forecast_id") or row.get("pick_id"))
            for row in rows
            if row.get("forecast_id") or row.get("pick_id")
        ),
        "candidate_records": records,
        "grade_counts": dict(sorted(grade_counts.items())),
        "price_evidence_counts": dict(sorted(evidence_counts.items())),
        "eligibility_reason_counts": dict(sorted(reason_counts.items())),
        "recorded_in": [
            "current_run_artifact", "dated_pick_ledger", "rolling_pick_audit",
        ],
    }


def _group_unpriced_shadow_candidates(
    rows: List[Dict[str, Any]],
) -> List[List[Dict[str, Any]]]:
    """Partition qualifying audit picks into display groups, not ticket objects."""
    qualifying = []
    for row in rows:
        success_rate = row.get("_historical_success_rate")
        if not isinstance(success_rate, dict):
            success_rate = _unpriced_historical_success_rate(row, [])
        if success_rate.get("qualifies") is True:
            candidate = dict(row)
            candidate["_historical_success_rate"] = success_rate
            qualifying.append(candidate)
    qualifying.sort(
        key=lambda row: (
            -float(row["_historical_success_rate"].get("rate") or 0.0),
            str(row.get("scheduled_utc") or ""),
            str(row.get("game_id") or ""),
            str(row.get("market") or ""),
            str(row.get("period") or ""),
            str(row.get("selection") or ""),
            str(row.get("line") or ""),
            str(row.get("forecast_id") or row.get("pick_id") or ""),
        )
    )

    groups: List[Dict[str, Any]] = []
    for candidate in qualifying:
        game_id = str(candidate.get("game_id") or "")
        group = next(
            (
                existing
                for existing in groups
                if len(existing["rows"]) < UNPRICED_SHADOW_GROUP_MAX_PICKS
                and game_id not in existing["game_ids"]
            ),
            None,
        )
        if group is None:
            group = {"rows": [], "game_ids": set()}
            groups.append(group)
        group["rows"].append(candidate)
        group["game_ids"].add(game_id)
    return [group["rows"] for group in groups]


def _finite_float(value: Any) -> Optional[float]:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    try:
        number = float(value)
    except (OverflowError, TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _row_probability(row: Dict[str, Any]) -> Optional[float]:
    probability = _finite_float(row.get("calibrated_probability"))
    if probability is None:
        probability = _finite_float(row.get("raw_probability"))
    return probability


def _is_unpriced_spread_shadow_source(row: Dict[str, Any]) -> bool:
    """A no-price full-game moneyline pick strong enough to re-express ATS."""
    if _has_exact_selected_price(row):
        return False
    if str(row.get("market") or "").lower() != "moneyline":
        return False
    if str(row.get("period") or "").lower() != "full_game":
        return False
    if str(row.get("selection") or "").lower() not in {"home", "away"}:
        return False
    probability = _row_probability(row)
    return probability is not None and probability >= UNPRICED_SPREAD_SHADOW_MIN_PROBABILITY


def _is_posted_spread_row(row: Dict[str, Any]) -> bool:
    """A full-game spread forecast carrying a real pre-match bookmaker line."""
    if str(row.get("market") or "").lower() != "spread":
        return False
    if str(row.get("period") or "").lower() != "full_game":
        return False
    if str(row.get("selection") or "").lower() not in {"home", "away"}:
        return False
    if _finite_float(row.get("line")) is None:
        return False
    book = row.get("book")
    if not isinstance(book, dict) or not book.get("provider"):
        return False
    return book.get("pre_match") is not False


def _translated_spread_grade(spread_row: Dict[str, Any], selection: str) -> str:
    """Grade ``selection`` at the spread row's line from the spread row's grade."""
    grade = str(spread_row.get("grade_status") or "pending").lower()
    if grade not in {"win", "loss", "push"}:
        return "pending"
    if str(spread_row.get("selection") or "").lower() == selection or grade == "push":
        return grade
    return "loss" if grade == "win" else "win"


def _unpriced_spread_shadow_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Re-express strong no-price moneyline picks against the posted spread.

    One record per game: the strongest qualifying moneyline pick picks the
    side; the latest pre-match posted full-game spread on that game supplies
    the line and that side's quoted price. No line is ever model-derived:
    games without a posted spread are returned with ``posted_spread=None`` and
    are not graded. These records are audit-only and never stake units.
    """
    by_game: Dict[str, Dict[str, Any]] = {}
    spreads: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        if not isinstance(row, dict):
            continue
        game_key = f"{row.get('date') or ''}|{row.get('game_id') or ''}"
        if not row.get("game_id"):
            continue
        if _is_posted_spread_row(row):
            spreads[game_key].append(row)
        elif _is_unpriced_spread_shadow_source(row):
            current = by_game.get(game_key)
            if current is None or (
                _row_probability(row),
                str(row.get("forecast_id") or row.get("pick_id") or ""),
            ) > (
                _row_probability(current),
                str(current.get("forecast_id") or current.get("pick_id") or ""),
            ):
                by_game[game_key] = row

    records = []
    for game_key, source in by_game.items():
        selection = str(source["selection"]).lower()
        candidates = spreads.get(game_key) or []
        posted = (
            max(
                candidates,
                key=lambda row: (
                    str((row.get("book") or {}).get("observed_utc") or ""),
                    str(row.get("forecast_id") or row.get("pick_id") or ""),
                ),
            )
            if candidates
            else None
        )
        record: Dict[str, Any] = {
            "record_type": "unpriced_spread_shadow_candidate",
            "date": str(source.get("date") or "")[:10] or None,
            "league": source.get("league"),
            "game_id": source.get("game_id"),
            "home": source.get("home"),
            "away": source.get("away"),
            "scheduled_utc": source.get("scheduled_utc"),
            "moneyline_forecast_id": source.get("forecast_id") or source.get("pick_id"),
            "moneyline_model": source.get("model"),
            "moneyline_probability": _row_probability(source),
            "moneyline_grade_status": str(source.get("grade_status") or "pending"),
            # Keep the actual no-price source snapshot available to any future
            # derived receipt. Its side/probability is evidence only; it does
            # not supply a synthetic spread price or cover probability.
            "moneyline_evidence": deepcopy(source),
            "selection": selection,
            "posted_spread": None,
            "selected_line": None,
            "price": None,
            "book_provider": None,
            "book": None,
            "grade_status": "untracked_no_posted_spread",
            "audit_only": True,
            "ticketable": False,
            "stake_units": 0.0,
        }
        if posted is not None:
            book = posted.get("book") or {}
            side_prices = book.get("side_prices") if isinstance(book.get("side_prices"), dict) else {}
            price = _finite_float(side_prices.get(selection))
            if price is None and str(posted.get("selection") or "").lower() == selection:
                price = _finite_float(book.get("selected_decimal_price"))
            record.update(
                {
                    "posted_spread": {
                        "forecast_id": posted.get("forecast_id") or posted.get("pick_id"),
                        "home_line": float(posted["line"]),
                        "model_selection": str(posted.get("selection") or "").lower(),
                        "model_probability": _row_probability(posted),
                        "observed_utc": book.get("observed_utc"),
                        # Preserve the complete observed quote rather than
                        # synthesizing a cross-market price later.  The audit
                        # renderer uses the compact fields above; the future
                        # paper lane consumes this frozen evidence verbatim.
                        "book_evidence": deepcopy(book),
                    },
                    "selected_line": _spread_line_for_selection(posted, selection),
                    "price": price if price is not None and price > 1.0 else None,
                    "book_provider": book.get("provider"),
                    "book": _canonical_book_for_key(book.get("book")),
                    "grade_status": _translated_spread_grade(posted, selection),
                }
            )
        records.append(record)
    records.sort(
        key=lambda item: (
            str(item.get("scheduled_utc") or ""),
            str(item.get("game_id") or ""),
        )
    )
    return records


def _unpriced_spread_shadow_record(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Win/loss/push and flat-stake units at posted prices for tracked records."""
    tracked = [record for record in records if record.get("posted_spread")]
    wins = sum(record["grade_status"] == "win" for record in tracked)
    losses = sum(record["grade_status"] == "loss" for record in tracked)
    pushes = sum(record["grade_status"] == "push" for record in tracked)
    pending = sum(record["grade_status"] == "pending" for record in tracked)
    priced = [
        record
        for record in tracked
        if record["grade_status"] in {"win", "loss", "push"} and record.get("price")
    ]
    units = sum(
        (float(record["price"]) - 1.0)
        if record["grade_status"] == "win"
        else -1.0
        if record["grade_status"] == "loss"
        else 0.0
        for record in priced
    )
    decisions = wins + losses
    return {
        "tracked": len(tracked),
        "untracked_no_posted_spread": len(records) - len(tracked),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "pending": pending,
        "decisions": decisions,
        "cover_rate": round(wins / decisions, 6) if decisions else None,
        "priced_settled": len(priced),
        "units_at_posted_price": round(units, 6),
        "roi_at_posted_price": round(units / len(priced), 6) if priced else None,
    }


def _unpriced_spread_shadow(
    day_rows: List[Dict[str, Any]], settled_history: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Today's ATS re-expressions plus the rolling settled record."""
    today = _unpriced_spread_shadow_rows(day_rows)
    history = _unpriced_spread_shadow_rows(settled_history)
    return {
        "schema_version": UNPRICED_SPREAD_SHADOW_SCHEMA,
        "record_type": "unpriced_spread_shadow_audit",
        "audit_only": True,
        "ticketable": False,
        "stake_units": 0.0,
        "policy": {
            "source_market": "moneyline",
            "source_period": "full_game",
            "source_requires_no_exact_price": True,
            "source_min_probability": UNPRICED_SPREAD_SHADOW_MIN_PROBABILITY,
            "one_record_per_game": "highest-probability moneyline side",
            "line_source": "latest pre-match posted full-game bookmaker spread",
            "model_derived_lines": False,
            "history_source": "rolling_pick_audit.settled_ledger",
            "history_window_days": UNPRICED_SUCCESS_RATE_HISTORY_DAYS,
        },
        "today": _unpriced_spread_shadow_record(today),
        "rolling": _unpriced_spread_shadow_record(history),
        "candidate_records": today,
    }


def _derived_spread_named_quote(record: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Return a valid selected-side named-book quote preserved by a shadow.

    This intentionally checks the actual posted-spread snapshot, not the
    no-price moneyline source.  A source forecast never supplies or implies a
    price for this derived contract.
    """

    posted = record.get("posted_spread")
    if not isinstance(posted, dict):
        return None
    book = posted.get("book_evidence")
    if not isinstance(book, dict):
        return None
    price = _finite_float(record.get("price"))
    if price is None or price <= 1.0:
        return None
    if book.get("pre_match") is not True:
        return None
    if book.get("price_provenance") != PriceProvenance.NAMED_BOOK.value:
        return None
    if book.get("source_permission") == SourcePermission.PROHIBITED.value:
        return None
    if not str(book.get("provider") or "").strip() or not str(book.get("book") or "").strip():
        return None
    # Exact market identity is part of the price contract.  A similarly named
    # game, a different period, or a different spread line cannot be reused.
    if (
        str(book.get("game_id") or "") != str(record.get("game_id") or "")
        or str(book.get("market") or "") != "spread"
        or str(book.get("period") or "") != "full_game"
        or not _same_line(book.get("line"), posted.get("home_line"))
    ):
        return None
    selection = str(record.get("selection") or "")
    side_prices = book.get("side_prices")
    side_price = _finite_float(side_prices.get(selection)) if isinstance(side_prices, dict) else None
    if side_price is None:
        if str(posted.get("model_selection") or "") != selection:
            return None
        side_price = _finite_float(book.get("selected_decimal_price"))
    if side_price is None or not math.isclose(side_price, price, rel_tol=0.0, abs_tol=1e-9):
        return None
    try:
        observed = parse_utc(str(book["observed_utc"]))
        scheduled = parse_utc(str(record["scheduled_utc"]))
    except (KeyError, TypeError, ValueError):
        return None
    if observed >= scheduled:
        return None
    frozen_book = deepcopy(book)
    frozen_book["selected_decimal_price"] = price
    # Side prices are evidence, but make the selected-side price explicit for
    # the shared receipt/settlement renderer.
    side_prices = frozen_book.get("side_prices")
    if isinstance(side_prices, dict):
        frozen_book["side_prices"] = dict(side_prices)
        frozen_book["side_prices"][str(record["selection"])] = price
    return frozen_book


def _derived_spread_ticket_history(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Evaluate a dedicated historical cover/ROI gate for derived spreads.

    The sample is deliberately not the native spread model's ROI history, nor
    is it calibrated moneyline history.  Every qualifying observation has the
    same cross-market source contract and an exact pre-tip named-book spread
    quote, and is graded against that quoted selected-side line.
    """

    eligible = []
    for record in records:
        book = _derived_spread_named_quote(record)
        if book is None:
            continue
        status = str(record.get("grade_status") or "").lower()
        if status not in {"win", "loss", "push"}:
            continue
        eligible.append((record, book, status))
    wins = sum(status == "win" for _, _, status in eligible)
    losses = sum(status == "loss" for _, _, status in eligible)
    pushes = sum(status == "push" for _, _, status in eligible)
    decisions = wins + losses
    units = sum(
        float(book["selected_decimal_price"]) - 1.0
        if status == "win"
        else -1.0
        if status == "loss"
        else 0.0
        for _, book, status in eligible
    )
    cover_rate = wins / decisions if decisions else None
    roi = units / len(eligible) if eligible else None
    qualifies = (
        decisions > UNPRICED_SPREAD_TICKET_MIN_DECISIONS_EXCLUSIVE
        and cover_rate is not None
        and cover_rate >= UNPRICED_SPREAD_TICKET_MIN_COVER_RATE
        and roi is not None
        and roi > 0.0
    )
    return {
        "contract": "unpriced_moneyline_spread",
        "history_scope": "derived_spread_named_book_pre_tip_quotes",
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "decisions": decisions,
        "priced_settled": len(eligible),
        "cover_rate": round(cover_rate, 6) if cover_rate is not None else None,
        "units": round(units, 6),
        "roi": round(roi, 6) if roi is not None else None,
        "minimum_decisions_exclusive": UNPRICED_SPREAD_TICKET_MIN_DECISIONS_EXCLUSIVE,
        "minimum_cover_rate": UNPRICED_SPREAD_TICKET_MIN_COVER_RATE,
        "positive_roi_required": True,
        "qualifies": qualifies,
        "status": (
            "qualified"
            if qualifies
            else "insufficient_history"
            if decisions <= UNPRICED_SPREAD_TICKET_MIN_DECISIONS_EXCLUSIVE
            else "below_cover_rate"
            if cover_rate is None or cover_rate < UNPRICED_SPREAD_TICKET_MIN_COVER_RATE
            else "non_positive_roi"
        ),
    }


def _derived_spread_ticket_candidates(
    rows: List[Dict[str, Any]],
    settled_history: List[Dict[str, Any]],
    *,
    generated_utc: str,
    target_date: str,
    slate_timezone_name: str,
) -> tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Build future-only derived-spread paper singles from frozen evidence.

    The source selection is a high-confidence, unpriced full-game moneyline
    forecast.  It contributes only its side and provenance; its win chance is
    never converted into a spread-cover chance.  The price/line comes entirely
    from a same-game posted spread snapshot, and this lane has its own rolling
    cover/ROI admission gate.
    """

    history = _derived_spread_ticket_history(_unpriced_spread_shadow_rows(settled_history))
    if not history["qualifies"]:
        return [], history
    now = parse_utc(generated_utc)
    zone = ZoneInfo(slate_timezone_name)
    candidates = []
    for record in _unpriced_spread_shadow_rows(rows):
        book = _derived_spread_named_quote(record)
        if book is None:
            continue
        try:
            scheduled = parse_utc(str(record["scheduled_utc"]))
            observed = parse_utc(str(book["observed_utc"]))
        except (KeyError, TypeError, ValueError):
            continue
        # This is the prospective guard: an already started fixture can never
        # enter the lane, regardless of when audit history becomes available.
        # Nor may an observation timestamp later than the decision scan be
        # used as if it had already been known.
        if (
            now >= scheduled
            or observed > now
            or scheduled.astimezone(zone).date().isoformat() != target_date
        ):
            continue
        source_id = str(record.get("moneyline_forecast_id") or "")
        posted = record.get("posted_spread")
        selected_line = _finite_float(record.get("selected_line"))
        probability = _finite_float(record.get("moneyline_probability"))
        if not source_id or not isinstance(posted, dict) or selected_line is None or probability is None:
            continue
        quote_id = str(posted.get("forecast_id") or "")
        identity = f"derived-spread|{source_id}|{quote_id}|{record.get('selection')}|{selected_line:g}"
        forecast_id = hashlib.sha256(identity.encode()).hexdigest()[:20]
        source_evidence = {
            "forecast_id": source_id,
            "market": "moneyline",
            "period": "full_game",
            "selection": record["selection"],
            "probability": probability,
            "model": record.get("moneyline_model"),
            "native_price_missing": True,
            "frozen_forecast": deepcopy(record.get("moneyline_evidence") or {}),
        }
        posted_evidence = {
            "forecast_id": posted.get("forecast_id"),
            "home_line": posted.get("home_line"),
            "selected_spread_line": selected_line,
            "observed_utc": posted.get("observed_utc"),
            "book": deepcopy(book),
        }
        leg = {
            "forecast_id": forecast_id,
            "pick_id": forecast_id,
            "game_id": record["game_id"],
            "league": record.get("league"),
            "scheduled_utc": record["scheduled_utc"],
            "home": record["home"],
            "away": record["away"],
            "market": "spread",
            "period": "full_game",
            "selection": record["selection"],
            # The canonical home line remains available for generic renderers;
            # grading uses the explicit selected-side frozen line below.
            "line": posted.get("home_line"),
            "raw_probability": probability,
            "calibrated_probability": probability,
            "model": "derived_unpriced_moneyline_spread",
            "book": book,
            "eligibility": {
                "acca_eligible": False,
                "reasons": ["derived_spread_separate_contract"],
            },
            "derived_spread": {
                "contract": "unpriced_moneyline_spread",
                "version": 1,
                "source_moneyline": source_evidence,
                "posted_spread_quote": posted_evidence,
                "selected_spread_line": selected_line,
                "history_gate": deepcopy(history),
                "frozen_before_tip": True,
            },
        }
        ticket_id = hashlib.sha256(f"paper-derived-spread|{forecast_id}".encode()).hexdigest()[:20]
        candidates.append(
            {
                "ticket_id": ticket_id,
                "acca_id": ticket_id,
                "ticket_type": "single",
                "legs": [leg],
                "leg_count": 1,
                "combined_decimal_odds": float(book["selected_decimal_price"]),
                # This remains source-side evidence, not a cover-probability
                # forecast; the receipt marks it accordingly.
                "estimated_win_probability": probability,
                "estimated_joint_win_probability": probability,
                "provisional_sample_gate_bypassed": False,
                "provisional_leg_count": 0,
                "ranking_status": "paper_only_derived_spread_history_qualified",
                "individual_pick_filter": {
                    "status": "qualified_derived_spread_contract",
                    "contract": "unpriced_moneyline_spread",
                    "history_gate": deepcopy(history),
                },
            }
        )
    candidates.sort(key=lambda candidate: (
        str(candidate["legs"][0].get("scheduled_utc") or ""),
        str(candidate["legs"][0].get("forecast_id") or ""),
    ))
    return candidates, history


def _priced_ineligible_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Return exact-priced forecasts whose existing eligibility gate failed."""
    return [
        row
        for row in rows
        if _has_exact_selected_price(row)
        and isinstance(row.get("eligibility"), dict)
        and row["eligibility"].get("acca_eligible") is False
    ]


def _priced_ineligible_shadow_candidates(
    rows: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Filter NCAA_M shadow candidates by selected decimal price and model p.

    Calibrated probability is preferred; raw probability is the same fallback
    used by operator-card rendering when no calibrated value is available.
    This is display-only and never changes the forecast's eligibility state.
    """
    candidates = []
    for row in _priced_ineligible_rows(rows):
        book = row.get("book") or {}
        try:
            price = float(book.get("selected_decimal_price"))
        except (TypeError, ValueError, OverflowError):
            continue
        probability, probability_basis = _confidence_probability(row)
        if (
            price < PRICED_INELIGIBLE_SHADOW_MIN_DECIMAL_ODDS
            or probability is None
            or probability < PRICED_INELIGIBLE_SHADOW_MIN_PROBABILITY
        ):
            continue
        candidate = dict(row)
        candidate["_shadow_probability"] = probability
        candidate["_shadow_probability_basis"] = probability_basis
        candidates.append(candidate)

    candidates.sort(
        key=lambda row: (
            -float(row["_shadow_probability"]),
            str(row.get("scheduled_utc") or ""),
            str(row.get("game_id") or ""),
            str(row.get("market") or ""),
            str(row.get("period") or ""),
            str(row.get("selection") or ""),
            str(row.get("forecast_id") or row.get("pick_id") or ""),
        )
    )
    return candidates


def _group_priced_ineligible_shadow_candidates(
    rows: List[Dict[str, Any]],
) -> List[List[Dict[str, Any]]]:
    """Group priced-ineligible candidates for display, never as tickets."""
    groups: List[Dict[str, Any]] = []
    for candidate in rows:
        game_id = str(candidate.get("game_id") or "")
        group = next(
            (
                existing
                for existing in groups
                if len(existing["rows"]) < SHADOW_GROUP_MAX_PICKS
                and game_id not in existing["game_ids"]
            ),
            None,
        )
        if group is None:
            group = {"rows": [], "game_ids": set()}
            groups.append(group)
        group["rows"].append(candidate)
        group["game_ids"].add(game_id)
    return [group["rows"] for group in groups]


def _priced_ineligible_candidate_record(row: Dict[str, Any]) -> Dict[str, Any]:
    book = row.get("book") if isinstance(row.get("book"), dict) else {}
    probability, probability_basis = _confidence_probability(row)
    price = book.get("selected_decimal_price")
    forecast_id = str(row.get("forecast_id") or row.get("pick_id") or "")
    return {
        "record_type": "priced_ineligible_forecast_candidate",
        "candidate_id": forecast_id,
        "forecast_id": forecast_id or None,
        "pick_id": str(row.get("pick_id") or forecast_id) or None,
        "date": str(row.get("date") or "")[:10] or None,
        "game_id": row.get("game_id"),
        "home": row.get("home"),
        "away": row.get("away"),
        "scheduled_utc": row.get("scheduled_utc"),
        "market": row.get("market"),
        "period": row.get("period"),
        "selection": row.get("selection"),
        "line": row.get("line"),
        "model": row.get("model"),
        "model_version": row.get("model_version"),
        "bucket": row.get("bucket"),
        "acca_eligible": False,
        "eligibility_reasons": list((row.get("eligibility") or {}).get("reasons") or []),
        "probability": probability,
        "probability_basis": probability_basis,
        "selected_decimal_price": float(price),
        "price_evidence": str(row.get("price_evidence") or _price_evidence(row)),
        "book_provider": book.get("provider"),
        "book": _canonical_book_for_key(book.get("book")),
        "book_observed_utc": book.get("observed_utc"),
        "book_pre_match": book.get("pre_match"),
        "grade_status": str(row.get("grade_status") or "pending"),
        "outcome": row.get("outcome"),
        "ticket_id": None,
        "ticketable": False,
        "audit_only": True,
        "stake_units": 0.0,
    }


def _priced_ineligible_shadow_summary(
    audited_rows: List[Dict[str, Any]],
    candidates: List[Dict[str, Any]],
) -> Dict[str, Any]:
    groups = _group_priced_ineligible_shadow_candidates(candidates)
    return {
        "schema_version": PRICED_INELIGIBLE_SHADOW_SCHEMA,
        "record_type": "priced_ineligible_shadow_audit",
        "audit_only": True,
        "ticketable": False,
        "stake_units": 0.0,
        "priced_ineligible_audited_count": len(audited_rows),
        "qualified_count": len(candidates),
        "minimum_selected_decimal_odds": PRICED_INELIGIBLE_SHADOW_MIN_DECIMAL_ODDS,
        "minimum_probability": PRICED_INELIGIBLE_SHADOW_MIN_PROBABILITY,
        "probability_policy": "calibrated_probability_then_raw_probability_fallback",
        "candidate_records": [
            _priced_ineligible_candidate_record(row) for row in candidates
        ],
        "shadow_groups": [
            [str(row.get("forecast_id") or row.get("pick_id") or "") for row in group]
            for group in groups
        ],
        "shadow_group_policy": {
            "display_only": True,
            "max_picks_per_group": SHADOW_GROUP_MAX_PICKS,
            "one_pick_per_game": True,
            "ticket_id": None,
            "receipt_created": False,
            "stake_units": 0.0,
        },
    }


def _render_unpriced_candidate_line(
    row: Dict[str, Any], timezone_name: str
) -> str:
    probability = row.get("calibrated_probability")
    if probability is None:
        probability = row.get("raw_probability")
    try:
        probability_text = f"{float(probability):.1%}"
    except (TypeError, ValueError, OverflowError):
        probability_text = "n/a"
    evidence = str(row.get("price_evidence") or _price_evidence(row))
    book = row.get("book") or {}
    if evidence == "UNMATCHED":
        price_text = "no exact selected-side book match"
    elif evidence == "MATCHED_BUT_SELECTED_PRICE_MISSING":
        price_text = f"{book.get('book') or 'book'} matched; selected-side quote missing"
    else:
        price_text = "no usable exact selected-side price"
    reasons = (row.get("eligibility") or {}).get("reasons") or []
    reason_text = ", ".join(str(reason) for reason in reasons) or "no recorded eligibility reason"
    calibration = row.get("calibration") or {}
    success_rate = row["_historical_success_rate"]
    success_text = (
        f"historical success={success_rate['rate']:.1%} "
        f"({success_rate['wins']}/{success_rate['decisions']})"
    )
    forecast_id = row.get("forecast_id") or row.get("pick_id") or "unknown"
    status = _grade_badge(row.get("grade_status"))
    tip = _local_time(str(row["scheduled_utc"]), timezone_name)
    return (
        f"    [{status}] {row['home']} vs {row['away']} · {tip} · "
        f"{_audit_market_selection(row)} · p={probability_text} · {price_text} · "
        f"{success_text} · WF n={calibration.get('n', 0)} "
        f"({calibration.get('status', 'unknown')}) · "
        f"blockers: {reason_text} · forecast={forecast_id}"
    )


def _render_priced_ineligible_candidate_line(
    row: Dict[str, Any], timezone_name: str
) -> str:
    probability = row.get("_shadow_probability")
    probability_basis = row.get("_shadow_probability_basis") or "probability"
    try:
        probability_text = f"{float(probability):.1%}"
    except (TypeError, ValueError, OverflowError):
        probability_text = "n/a"
    book = row.get("book") or {}
    price = float(book["selected_decimal_price"])
    book_name = _canonical_book_for_key(book.get("book")) or "named book"
    observed = str(book.get("observed_utc") or "time unknown")
    reasons = (row.get("eligibility") or {}).get("reasons") or []
    reason_text = ", ".join(str(reason) for reason in reasons) or "no recorded eligibility reason"
    status = _grade_badge(row.get("grade_status"))
    tip = _local_time(str(row["scheduled_utc"]), timezone_name)
    forecast_id = row.get("forecast_id") or row.get("pick_id") or "unknown"
    return (
        f"    [{status}] {row['home']} vs {row['away']} · {tip} · "
        f"{_audit_market_selection(row)} · p={probability_text} ({probability_basis}) · "
        f"selected decimal odds={price:.3f} {book_name} (observed {observed}) · "
        f"bucket={row.get('bucket') or 'unknown'} · blockers: {reason_text} · "
        f"forecast={forecast_id}"
    )


def _render_priced_ineligible_shadow_section(
    rows: List[Dict[str, Any]],
    *,
    audited_count: int,
    timezone_name: str,
) -> List[str]:
    """Render exact-priced, ineligible NCAA_M candidates without ticketing."""
    if audited_count <= 0:
        return []
    odds_text = f"{PRICED_INELIGIBLE_SHADOW_MIN_DECIMAL_ODDS:.2f}"
    probability_text = f"{PRICED_INELIGIBLE_SHADOW_MIN_PROBABILITY:.0%}"
    groups = _group_priced_ineligible_shadow_candidates(rows)
    if not groups:
        return [
            "",
            "PRICED-INELIGIBLE AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY "
            f"({audited_count} exact-priced ineligible audited; requires decimal odds ≥{odds_text} "
            f"and p ≥{probability_text})",
            "No candidate meets both shadow-display thresholds.",
            "These candidates remain ineligible; none is a paper position or ticket leg.",
        ]

    qualified = [row for group in groups for row in group]
    grade_counts = Counter(str(row.get("grade_status") or "pending") for row in qualified)
    grade_summary = _grade_count_summary(grade_counts, count_first=True, separator=", ")
    group_label = "group" if len(groups) == 1 else "groups"
    lines = [
        "",
        f"PRICED-INELIGIBLE AUDIT-ONLY SHADOW GROUPS (decimal odds ≥{odds_text}; p ≥{probability_text}) — "
        f"{len(qualified)} picks in {len(groups)} {group_label} ({grade_summary}; "
        f"{audited_count} exact-priced ineligible audited)",
        "Display grouping only: selected prices are shown, but eligibility remains false; no ticket ID, receipt, or stake.",
        f"Each pick meets both display thresholds; groups hold up to {SHADOW_GROUP_MAX_PICKS} picks, "
        "with at most one pick per fixture. No combined probability, bank, or ROI impact.",
    ]
    for index, group in enumerate(groups, 1):
        group_grades = Counter(str(row.get("grade_status") or "pending") for row in group)
        group_grade_summary = _grade_count_summary(
            group_grades, count_first=True, separator=", "
        )
        pick_label = "pick" if len(group) == 1 else "picks"
        lines.append(
            f"  SHADOW GROUP {index} — {len(group)} priced-ineligible audit-only {pick_label} "
            f"({group_grade_summary})"
        )
        lines.extend(
            _render_priced_ineligible_candidate_line(row, timezone_name)
            for row in group
        )
    return lines


def _render_unpriced_audit_section(
    rows: List[Dict[str, Any]], timezone_name: str
) -> List[str]:
    """Render hit-rate-qualified no-price picks in non-wagerable display groups."""
    if not rows:
        return []
    threshold_text = f"{UNPRICED_SUCCESS_RATE_THRESHOLD:.0%}"
    minimum_text = f"n>{UNPRICED_SUCCESS_RATE_MIN_DECISIONS_EXCLUSIVE}"
    groups = _group_unpriced_shadow_candidates(rows)
    if not groups:
        return [
            "",
            f"UNPRICED AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY "
            f"({len(rows)} audited; requires ≥{threshold_text} historical success, {minimum_text})",
            "No no-price candidate meets the historical success-rate gate for display.",
            "All candidates remain individually recorded in the audit ledger; none are wager positions.",
        ]

    qualified = [row for group in groups for row in group]
    grade_counts = Counter(str(row.get("grade_status") or "pending") for row in qualified)
    grade_summary = _grade_count_summary(grade_counts, count_first=True, separator=", ")
    group_label = "group" if len(groups) == 1 else "groups"
    lines = [
        "",
        f"UNPRICED AUDIT-ONLY SHADOW GROUPS (≥{threshold_text} HISTORICAL SUCCESS, {minimum_text}) — "
        f"{len(qualified)} picks in {len(groups)} {group_label} ({grade_summary}; {len(rows)} audited)",
        "Display grouping only: no odds, combined success rate, ticket ID, receipt, or stake; no bank/ROI impact.",
        f"Each pick passes the individual historical gate; groups hold up to {UNPRICED_SHADOW_GROUP_MAX_PICKS} picks, "
        "with at most one pick per fixture. Group success is not estimated.",
        "Every no-price forecast remains individually persisted and graded from reconciled official results.",
    ]
    for index, group in enumerate(groups, 1):
        group_grades = Counter(str(row.get("grade_status") or "pending") for row in group)
        group_grade_summary = _grade_count_summary(
            group_grades, count_first=True, separator=", "
        )
        pick_label = "pick" if len(group) == 1 else "picks"
        lines.append(
            f"  SHADOW GROUP {index} — {len(group)} audit-only {pick_label} "
            f"({group_grade_summary})"
        )
        lines.extend(_render_unpriced_candidate_line(row, timezone_name) for row in group)
    return lines


def _render_unpriced_spread_shadow_section(
    shadow: Optional[Dict[str, Any]], timezone_name: str
) -> List[str]:
    """Render no-price moneyline picks re-expressed against posted spreads."""
    if not shadow:
        return []
    records = list(shadow.get("candidate_records") or [])
    rolling = shadow.get("rolling") or {}
    today = shadow.get("today") or {}
    threshold_text = f"{UNPRICED_SPREAD_SHADOW_MIN_PROBABILITY:.0%}"

    def record_text(summary: Dict[str, Any]) -> str:
        text = (
            f"{int(summary.get('wins') or 0)}W-{int(summary.get('losses') or 0)}L-"
            f"{int(summary.get('pushes') or 0)}P"
        )
        rate = summary.get("cover_rate")
        if rate is not None:
            text += f" ({float(rate):.1%} cover)"
        priced = int(summary.get("priced_settled") or 0)
        if priced:
            text += (
                f" · {float(summary.get('units_at_posted_price') or 0.0):+.2f}u at posted "
                f"prices over {priced} priced"
            )
        return text

    lines = [
        "",
        f"UNPRICED → POSTED-SPREAD SHADOW (≥{threshold_text} NO-PRICE MONEYLINE PICKS) — "
        f"{int(today.get('tracked') or 0)} tracked, "
        f"{int(today.get('untracked_no_posted_spread') or 0)} without a posted spread",
        "Audit only: same side re-expressed against the bookmaker's posted full-game spread; "
        "no ticket, receipt, or stake; no bank/ROI impact.",
        "Lines come only from posted pre-match bookmaker spreads; no model-made lines. "
        "Games without a posted spread are not graded.",
        f"Rolling {UNPRICED_SUCCESS_RATE_HISTORY_DAYS}-day settled record: {record_text(rolling)}",
    ]
    tracked = [record for record in records if record.get("posted_spread")]
    for record in tracked:
        price = record.get("price")
        price_text = f"@{float(price):.3f}" if price else "@n/a"
        book = record.get("book") or record.get("book_provider") or "book"
        probability = record.get("moneyline_probability")
        probability_text = f"{float(probability):.1%}" if probability is not None else "n/a"
        selected_line = record.get("selected_line")
        line_text = f" {float(selected_line):+g}" if selected_line is not None else ""
        lines.extend(
            [
                f"   {record.get('home')} vs {record.get('away')} · "
                f"TIP-OFF {_local_time(str(record.get('scheduled_utc')), timezone_name)}",
                f"      SPREAD full_game{line_text} -> {str(record.get('selection')).upper()} "
                f"{price_text} {book} · ML p {probability_text} · "
                f"{_grade_badge(record.get('grade_status'))}",
            ]
        )
    untracked = len(records) - len(tracked)
    if untracked:
        lines.append(
            f"   {untracked} qualifying no-price moneyline pick"
            f"{'' if untracked == 1 else 's'} had no posted spread to track."
        )
    return lines


def _render_ticket_card(
    receipt: Dict[str, Any],
    performance: Dict[str, Any],
    timezone_name: str,
    *,
    unpriced_audit_picks: Optional[List[Dict[str, Any]]] = None,
    priced_ineligible_shadow_picks: Optional[List[Dict[str, Any]]] = None,
    priced_ineligible_audited_count: int = 0,
    unpriced_spread_shadow: Optional[Dict[str, Any]] = None,
) -> str:
    positions = _receipt_positions(receipt)
    acca_count = sum(ticket_type == "acca" for ticket_type, _ in positions)
    single_count = sum(ticket_type == "single" for ticket_type, _ in positions)
    position_labels = []
    if acca_count:
        position_labels.append(f"{acca_count} paper acca" + ("s" if acca_count != 1 else ""))
    if single_count:
        position_labels.append(f"{single_count} paper single" + ("s" if single_count != 1 else ""))
    if not position_labels:
        position_labels.append("0 paper positions")
    lines = [
        f"MA GOLIDE PAPER TICKETS — {receipt['league']} {receipt['date']}",
        "=" * 72,
        f"CURRENT PAPER BANK (all dates): {performance['paper_bank_units']:.3f} units "
        f"· free {performance['paper_free_units']:.3f} / reserved {performance['paper_committed_units']:.3f}",
        f"{', '.join(position_labels)} · 1 ticket publication · times in {timezone_name}",
        "PAPER ONLY — flat one unit per position; no automatic betting writes.",
        "Frozen tickets retain their original decision basis; empty scans do not cancel them.",
        "",
    ]
    if receipt.get("league") == League.EUROLEAGUE.value:
        lines[2:2] = [
            "EUROLEAGUE STRICT NAMED-BOOK TICKET LANE — exact named-book pre-tip quote; calibrated p × odds > 1.00;",
            "at least 10 past settled picks and positive ROI in the matching band; aggregate consensus prices are excluded.",
        ]
    if not positions:
        lines.extend(["NO PAPER BET — no single or multi-leg paper ticket at publication.", ""])
    type_indexes = Counter()
    for publication_index, (ticket_type, position) in enumerate(positions, 1):
        type_indexes[ticket_type] += 1
        type_index = type_indexes[ticket_type]
        books = ", ".join(position.get("bookmakers") or []) or "multiple/no book"
        provisional = bool(position.get("provisional_sample_gate_bypassed"))
        provisional_label = " · PROVISIONAL SAMPLE GATE" if provisional else ""
        leg_count = int(position.get("leg_count") or len(position.get("legs") or []))
        ticket_label = (
            f"PAPER SINGLE #{type_index}"
            if ticket_type == "single"
            else f"PAPER {leg_count}-LEG ACCA #{type_index}"
        )
        derived_spread_position = any(
            _is_derived_spread_ticket_row(leg) for leg in position.get("legs") or []
        )
        probability_label = (
            "source ML p (not cover p)"
            if derived_spread_position
            else "estimated p"
            if ticket_type == "single"
            else "joint p"
        )
        lines.append(
            f"[{ticket_label}]{provisional_label} "
            f"@{position['combined_decimal_odds']:.3f} "
            f"— {position['stake_units']:.6f} paper units — {books} "
            f"· {_grade_badge(position['grade_status'])}"
        )
        lines.append(
            f"   Published {_local_time(receipt['frozen_at'], timezone_name)} "
            f"· reference P1/{publication_index} · {probability_label}="
            f"{position['estimated_joint_win_probability']:.1%}"
        )
        pick_filter = position.get("individual_pick_filter") or {}
        candidate_count = int(pick_filter.get("candidate_pick_count") or 0)
        qualifying_count = int(pick_filter.get("qualifying_pick_count") or 0)
        history_count = int(pick_filter.get("historical_unique_pick_count") or 0)
        if pick_filter:
            lines.append(
                f"   Individual-pick ROI policy: {qualifying_count}/{candidate_count} "
                f"eligible candidates qualify across "
                f"{int(pick_filter.get('qualifying_band_count') or 0)} positive bands; "
                f"{history_count} unique settled picks · status={pick_filter.get('status', 'unknown')}."
            )
        band_details = pick_filter.get("candidate_band_details") or {}
        for leg in position["legs"]:
            details = band_details.get(leg.get("forecast_id"))
            if not isinstance(details, dict):
                details = leg.get("individual_pick_roi_band")
            if isinstance(details, dict):
                observed_roi = details.get("observed_roi")
                adjusted_roi = details.get("adjusted_roi")
                observed_text = (
                    f"{float(observed_roi):+.1%}"
                    if isinstance(observed_roi, (int, float)) and not isinstance(observed_roi, bool)
                    else "n/a"
                )
                adjusted_text = (
                    f"{float(adjusted_roi):+.1%}"
                    if isinstance(adjusted_roi, (int, float)) and not isinstance(adjusted_roi, bool)
                    else "n/a"
                )
                lines.append(
                    f"      ROI band: {details.get('market', leg.get('market'))} "
                    f"{details.get('period', leg.get('period'))} · "
                    f"p {details.get('probability_band', 'unknown')} · "
                    f"odds {details.get('individual_odds_band', 'unknown')} · "
                    f"n={int(details.get('historical_pick_count') or 0)} "
                    f"W-L-P={int(details.get('wins') or 0)}-"
                    f"{int(details.get('losses') or 0)}-{int(details.get('pushes') or 0)} · "
                    f"ROI {observed_text} observed / {adjusted_text} adjusted."
                )
            price = (leg.get("book") or {}).get("selected_decimal_price")
            probability = leg.get("calibrated_probability")
            if probability is None:
                probability = leg["raw_probability"]
            probability_text = (
                f"source ML p {probability:.1%} (not cover p)"
                if _is_derived_spread_ticket_row(leg)
                else f"ML {probability:.1%}"
            )
            lines.extend(
                [
                    f"   {leg['home']} vs {leg['away']} · "
                    f"TIP-OFF {_local_time(leg['scheduled_utc'], timezone_name)}",
                    f"      {_selection_label(leg)} @{float(price):.3f} "
                    f"· {probability_text} · {leg['bucket']} · {_grade_badge(leg['grade_status'])}",
                ]
            )
            if _is_derived_spread_ticket_row(leg):
                derivation = leg["derived_spread"]
                source = derivation.get("source_moneyline") or {}
                quote = derivation.get("posted_spread_quote") or {}
                history_gate = derivation.get("history_gate") or {}
                quote_book = quote.get("book") or {}
                lines.append(
                    "      Derived pre-tip contract: no-price ML source "
                    f"{source.get('forecast_id')} → posted spread quote {quote.get('forecast_id')} "
                    f"at {quote_book.get('book', 'unknown')} observed "
                    f"{_local_time(str(quote.get('observed_utc')), timezone_name)}."
                )
                lines.append(
                    "      Separate derived-spread gate: "
                    f"{history_gate.get('wins', 0)}W-{history_gate.get('losses', 0)}L-"
                    f"{history_gate.get('pushes', 0)}P · cover {history_gate.get('cover_rate')} · "
                    f"ROI {history_gate.get('roi')} (source ML p is not cover p)."
                )
        lines.append("")
    if priced_ineligible_shadow_picks is None:
        lines.extend(
            _render_unpriced_audit_section(unpriced_audit_picks or [], timezone_name)
        )
        lines.extend(
            _render_unpriced_spread_shadow_section(unpriced_spread_shadow, timezone_name)
        )
    else:
        lines.extend(
            _render_priced_ineligible_shadow_section(
                priced_ineligible_shadow_picks,
                audited_count=priced_ineligible_audited_count,
                timezone_name=timezone_name,
            )
        )
    lines.extend(
        [
            f"DECISION FROZEN AT {_local_time(receipt['frozen_at'], timezone_name)[11:16]}",
            "Selections, prices, and stakes stay frozen; displayed grades refresh when reconciled results arrive.",
            "Paper only. No real-money stakes or automatic promotion.",
            "",
        ]
    )
    return "\n".join(lines)


def _is_provisional_sample_only_row(row: Dict[str, Any]) -> bool:
    eligibility = row.get("eligibility") or {}
    if eligibility.get("acca_eligible") is not False:
        return False
    reasons = eligibility.get("reasons")
    if not isinstance(reasons, (list, tuple, set)):
        return False
    if set(reasons) != {"insufficient_walk_forward_samples"}:
        return False
    calibration = row.get("calibration") or {}
    if (
        calibration.get("status") != "provisional"
        or calibration.get("sufficient_for_accas") is not False
    ):
        return False
    n = calibration.get("n")
    if isinstance(n, bool) or not isinstance(n, int) or n < 1:
        return False
    minimum = calibration.get("minimum_samples")
    if minimum is not None and (
        isinstance(minimum, bool)
        or not isinstance(minimum, int)
        or minimum < 1
        or n >= minimum
    ):
        return False
    return True


def _ticket_candidate_funnel(market_report: Dict[str, Any] | None) -> Dict[str, Any]:
    """Summarize eligible picks, ROI-qualified singles, and paper combinations."""
    report = market_report or {}
    rows = list(report.get("forecasts") or [])
    acca_report = report.get("accas") or {}

    def exact_price(row: Dict[str, Any]) -> bool:
        return _has_exact_selected_price(row)

    priced = [row for row in rows if exact_price(row)]
    calibrated = [
        row for row in rows if (row.get("calibration") or {}).get("sufficient_for_accas") is True
    ]
    eligible = [row for row in rows if (row.get("eligibility") or {}).get("acca_eligible")]
    provisional = [
        row for row in rows if exact_price(row) and _is_provisional_sample_only_row(row)
    ]
    eligible_games = {row.get("game_id") for row in eligible}
    provisional_games = {row.get("game_id") for row in provisional}
    paper_results = acca_report.get("paper_results")
    if paper_results is None:  # legacy reports expose only the strict result surface
        paper_results = acca_report.get("results")
    paper_accas = list(paper_results or [])
    paper_single_picks = list(acca_report.get("paper_single_picks") or [])
    effective_paper_legs = acca_report.get("effective_paper_legs")
    if (
        isinstance(effective_paper_legs, bool)
        or not isinstance(effective_paper_legs, int)
        or effective_paper_legs < 2
    ):
        effective_paper_legs = (
            len(paper_accas[0].get("legs") or []) if paper_accas else None
        )
    roi_qualified_games = {
        row.get("game_id") for row in paper_single_picks if row.get("game_id")
    }
    strict_accas = acca_report.get("strict_results")
    if strict_accas is None:  # pre-exploration report compatibility
        strict_accas = list(acca_report.get("results") or [])
    else:
        strict_accas = list(strict_accas or [])
    blockers = Counter(
        reason for row in rows for reason in ((row.get("eligibility") or {}).get("reasons") or [])
    )
    try:
        required_games = int(acca_report.get("requested_legs") or 3)
    except (TypeError, ValueError):
        required_games = 3
    funnel = {
        "scan_rows": len(rows),
        "exact_priced_rows": len(priced),
        "calibration_sufficient_rows": len(calibrated),
        "priced_and_calibrated_rows": sum(
            exact_price(row) and (row.get("calibration") or {}).get("sufficient_for_accas") is True
            for row in rows
        ),
        "eligible_rows": len(eligible),
        "eligible_games": len(eligible_games),
        "provisional_sample_only_rows": len(provisional),
        "provisional_sample_only_games": len(provisional_games),
        "paper_candidate_rows": len(eligible) + len(provisional),
        "paper_candidate_games": len(eligible_games | provisional_games),
        "paper_single_picks": len(paper_single_picks),
        "paper_single_games": len(roi_qualified_games),
        "paper_accas": len(paper_accas),
        "strict_accas": len(strict_accas),
        "required_games": required_games,
        "effective_paper_legs": effective_paper_legs,
        "eligibility_blockers": dict(sorted(blockers.items())),
    }
    pricing_status = acca_report.get("pricing_evidence_status")
    if isinstance(pricing_status, str) and pricing_status:
        funnel["pricing_evidence_status"] = pricing_status
    pick_filter = acca_report.get("paper_pick_filter")
    if isinstance(pick_filter, dict):
        funnel["paper_pick_filter"] = deepcopy(pick_filter)
    return funnel


def _named_price_draft_candidates(
    rows: List[Dict[str, Any]], *, generated_utc: str, maximum: int = 8
) -> List[Dict[str, Any]]:
    """Return transparent Edge-style draft candidates, never ticket legs.

    This is deliberately broader than the strict receipt lane: it shows each
    current exact named-book quote before the decision cutoff, including rows
    that still fail model, calibration, or ROI gates. Source-permission and
    operational-validation labels remain visible audit metadata; aggregate
    consensus prices do not enter this named-book surface.
    """

    now = parse_utc(generated_utc)
    by_game: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        book = row.get("book")
        if not isinstance(book, dict) or not _has_exact_selected_price(row):
            continue
        if book.get("price_provenance") not in {None, PriceProvenance.NAMED_BOOK.value}:
            continue
        if not str(book.get("book") or "").strip() or book.get("pre_match") is not True:
            continue
        try:
            observed = parse_utc(str(book["observed_utc"]))
            tip = parse_utc(str(row["scheduled_utc"]))
        except (KeyError, TypeError, ValueError):
            continue
        if not observed < tip or not now < tip:
            continue
        game_id = str(row.get("game_id") or "")
        if not game_id:
            continue
        probability, _basis = _confidence_probability(row)
        if probability is None:
            continue
        current = by_game.get(game_id)
        # Probability orders candidates.  No price-derived ranking or
        # fabricated EV is used in this draft-only lane.
        if current is None or (
            probability,
            str(row.get("forecast_id") or ""),
        ) > (
            _confidence_probability(current)[0] or -1.0,
            str(current.get("forecast_id") or ""),
        ):
            by_game[game_id] = row
    return sorted(
        by_game.values(),
        key=lambda row: (
            -(_confidence_probability(row)[0] or 0.0),
            str(row.get("scheduled_utc") or ""),
            str(row.get("forecast_id") or ""),
        ),
    )[:maximum]


def _render_named_price_draft_section(
    *,
    generated_utc: str,
    timezone_name: str,
    ticket_decision: Dict[str, Any],
    rows: List[Dict[str, Any]],
    candidate_funnel: Dict[str, Any],
) -> List[str]:
    """Render the pre-cutoff named-price draft section of the dated ticket card.

    The strict ticket receipt and its bank remain unchanged. This section keeps
    current named-price candidates on the single dated card before its
    league-specific SAST publication cutoff (or earlier first-tip safety). It
    never reaches a terminal
    receipt or no-bet card.
    """

    candidates = _named_price_draft_candidates(rows, generated_utc=generated_utc)
    cutoff = ticket_decision.get("planned_cutoff_utc")
    strict_count = int(candidate_funnel.get("paper_accas") or 0)
    strict_singles = int(candidate_funnel.get("paper_single_picks") or 0)
    lines = [
        "PRE-CUTOFF NAMED-PRICE DRAFT SLIPS — DRAFT ONLY / PRE-CUTOFF",
        "- DRAFT ONLY: no paper stake, ticket receipt, settlement grade, bank impact, ROI/P&L impact, or real-money instruction.",
        "- One exact, observed, current pre-tip quote from a named bookmaker per displayed game; no price is inferred or promoted.",
        "- Unresolved source permission and operational-validation labels remain audit metadata; a named-book quote may enter the paper lane only after the model, calibration, and ROI gates pass.",
        "- A strict receipt is created only at the decision cutoff and only from qualifying named-book candidates.",
        (
            f"- Current decision cutoff: {_local_time(str(cutoff), timezone_name)}"
            if cutoff
            else "- Current decision cutoff: awaiting a confirmed slate fixture."
        ),
        f"- Strict paper-lane status this scan: {strict_count} multi-leg candidate(s) and "
        f"{strict_singles} ROI-qualified single(s) passed the strict model/price funnel.",
    ]
    if not candidates:
        lines.append("  (none — no exact, current named-book price was available before tip)")
        return lines

    for index in range(0, len(candidates), 2):
        draft_legs = candidates[index : index + 2]
        prices = [float((row.get("book") or {})["selected_decimal_price"]) for row in draft_legs]
        combined = math.prod(prices)
        lines.append(
            f"  [DRAFT {'2-LEG' if len(draft_legs) == 2 else 'SINGLE'} SLIP #{index // 2 + 1}] "
            f"@{combined:.3f} · no stake / no receipt"
        )
        for row in draft_legs:
            book = row["book"]
            probability, basis = _confidence_probability(row)
            reasons = ",".join((row.get("eligibility") or {}).get("reasons") or []) or "none"
            validation = "NAMED-BOOK PRICE" if book.get("ticket_price_eligible") is True else "RESEARCH-ONLY"
            lines.append(
                f"    [{validation}] {row.get('home')} vs {row.get('away')} · "
                f"{_selection_label(row)} · @{float(book['selected_decimal_price']):.3f} {book.get('book')} · "
                f"p={(probability or 0.0):.1%} ({basis}) · bucket={_bucket(row)} · blockers: {reasons}"
            )
    lines.append("Every displayed price is preserved as observed evidence only; no price is inferred or promoted.")
    return lines


def _render_empty_ticket_scan(
    *,
    league: League,
    target_date: str,
    generated_utc: str,
    performance: Dict[str, Any],
    timezone_name: str,
    candidate_funnel: Dict[str, Any],
    withdrawal: Optional[Dict[str, Any]] = None,
    no_bet_decision: Optional[Dict[str, Any]] = None,
    ticket_decision: Optional[Dict[str, Any]] = None,
    unpriced_audit_picks: Optional[List[Dict[str, Any]]] = None,
    priced_ineligible_shadow_picks: Optional[List[Dict[str, Any]]] = None,
    priced_ineligible_audited_count: int = 0,
    unpriced_spread_shadow: Optional[Dict[str, Any]] = None,
    historical_archive: bool = False,
    named_price_draft_rows: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """Render a mutable current scan or a regraded historical audit archive."""
    paper_single_count = int(candidate_funnel.get("paper_single_picks") or 0)
    paper_single_games = int(candidate_funnel.get("paper_single_games") or 0)
    pick_filter = candidate_funnel.get("paper_pick_filter")
    candidate_pick_count = 0
    qualifying_pick_count = paper_single_count
    if isinstance(pick_filter, dict):
        candidate_pick_count = int(pick_filter.get("candidate_pick_count") or 0)
        raw_qualifying_count = pick_filter.get("qualifying_pick_count")
        if raw_qualifying_count is not None:
            qualifying_pick_count = int(raw_qualifying_count)
    candidate_label = "candidate pick" if candidate_pick_count == 1 else "candidate picks"

    if historical_archive:
        scan_status = (
            "HISTORICAL SCAN ARCHIVE — preserved pick/shadow evidence regraded from "
            "authoritative results; no immutable ticket or no-bet receipt is present."
        )
    elif withdrawal:
        scan_status = (
            "NO PAPER BET — the earlier ticket publication was withdrawn for model-integrity failure."
        )
    elif paper_single_count == 1:
        scan_status = (
            "NO PAPER BET — one ROI-qualified selection was present, but its preserved "
            "eligibility/price evidence could not be frozen as a single; no multi-leg "
            "ticket was forced."
        )
    elif paper_single_count:
        if candidate_pick_count:
            roi_summary = (
                f"{qualifying_pick_count} of {candidate_pick_count} {candidate_label} passed "
                "individual-pick ROI"
            )
        else:
            pick_label = "pick" if paper_single_count == 1 else "picks"
            roi_summary = f"{paper_single_count} ROI-qualified single {pick_label}"
        if paper_single_games == 1:
            game_reason = "only 1 distinct game qualifies; 2 are required"
        elif paper_single_games == 0:
            game_reason = "no distinct games qualify; 2 are required"
        elif paper_single_games < 2:
            game_reason = f"only {paper_single_games} distinct games qualify; 2 are required"
        else:
            game_reason = "no multi-leg acca was produced"
        scan_status = f"NO MULTI-LEG PAPER ACCA — {roi_summary}; {game_reason}."
    else:
        scan_rows = int(candidate_funnel.get("scan_rows") or 0)
        paper_candidate_rows = int(candidate_funnel.get("paper_candidate_rows") or 0)
        if scan_rows == 0:
            scan_status = "NO PAPER BET — no current forecasts were generated."
        elif candidate_pick_count:
            scan_status = (
                f"NO PAPER BET — {qualifying_pick_count} of {candidate_pick_count} {candidate_label} "
                "passed individual-pick ROI."
            )
        elif paper_candidate_rows == 0:
            scan_status = "NO PAPER BET — no forecasts passed the paper eligibility gates."
        else:
            scan_status = (
                "NO PAPER BET — no ROI-qualified paper singles or multi-leg acca in this scan."
            )

    lines = [
        f"MA GOLIDE PAPER TICKETS — {league.value} {target_date}",
        "=" * 72,
        scan_status,
        "",
        f"CURRENT PAPER BANK (all dates): {performance['paper_bank_units']:.3f} units "
        f"· free {performance['paper_free_units']:.3f} / reserved {performance['paper_committed_units']:.3f}",
        f"{'Reconciled at' if historical_archive else 'Latest scan'}: "
        f"{_local_time(generated_utc, timezone_name)}",
        "",
        (
            "This historical view does not create, replace, or retrospectively freeze a ticket/no-bet decision."
            if historical_archive
            else "This date is closed to replacement ticket publication after the withdrawal."
            if withdrawal
            else "This date is closed by an immutable no-bet decision; later prices remain audit evidence only."
            if no_bet_decision
            else "This scan card is not a frozen receipt; a later qualifying scan may publish tickets."
        ),
        (
            "Only the preserved dated pick ledger is regraded; original buckets and revisions remain audit evidence."
            if historical_archive
            else "The withdrawn receipt is excluded from performance and retained unchanged as evidence."
            if withdrawal
            else "No-bet decision: "
            f"{_local_time(str(no_bet_decision.get('frozen_at')), timezone_name)} · "
            f"{no_bet_decision.get('reason', 'unknown')}"
            if no_bet_decision
            else "Empty scans never cancel or resize earlier frozen positions."
        ),
        "Paper only. No real-money stakes or automatic promotion.",
        "",
    ]
    if ticket_decision and not historical_archive and not withdrawal and not no_bet_decision:
        decision_status = ticket_decision.get("status")
        if decision_status == "pending_cutoff":
            cutoff = ticket_decision.get("planned_cutoff_utc")
            earliest = ticket_decision.get("earliest_slate_tip_utc")
            publication = ticket_decision.get("publication_cutoff_utc") or ticket_decision.get(
                "morning_publication_cutoff_utc"
            )
            publication_time = ticket_decision.get("publication_local_time") or "09:00"
            safety = ticket_decision.get("pre_tip_safety_cutoff_utc")
            if cutoff and earliest:
                policy_detail = (
                    f"{publication_time} SAST publication: {_local_time(str(publication), timezone_name)}"
                    if publication and cutoff == publication
                    else f"earlier pre-tip safety cutoff: {_local_time(str(safety or cutoff), timezone_name)}"
                )
                lines[7:7] = [
                    "Ticket decision pending — candidates and exact prices remain mutable until "
                    f"{_local_time(str(cutoff), timezone_name)} ({policy_detail}; first tip: "
                    f"{_local_time(str(earliest), timezone_name)}).",
                    "",
                ]
        elif decision_status == "no_slate_game":
            lines[7:7] = [
                "Ticket decision deferred — no scheduled game for this slate date is available yet.",
                "",
            ]
    # The single dated card is the only pre-cutoff draft surface. A withdrawal,
    # immutable no-bet decision, historical regrade, or passed cutoff suppresses
    # this section so draft candidates cannot reach a terminal card.
    if (
        named_price_draft_rows is not None
        and not historical_archive
        and withdrawal is None
        and no_bet_decision is None
        and ticket_decision
        and ticket_decision.get("status") == "pending_cutoff"
    ):
        lines.extend(
            (
                "",
                *_render_named_price_draft_section(
                    generated_utc=generated_utc,
                    timezone_name=timezone_name,
                    ticket_decision=ticket_decision,
                    rows=named_price_draft_rows,
                    candidate_funnel=candidate_funnel,
                ),
            )
        )
    derived_history = candidate_funnel.get("derived_spread_history")
    derived_count = int(candidate_funnel.get("derived_spread_ticket_candidates") or 0)
    if (
        isinstance(derived_history, dict)
        and not historical_archive
        and withdrawal is None
        and no_bet_decision is None
    ):
        lines.extend(
            [
                "",
                "DERIVED-SPREAD PAPER LANE — FUTURE-ONLY / PRE-TIP",
                "- A qualifying no-price moneyline side may be re-expressed only at an exact, frozen, named-book posted spread quote for that same game.",
                "- The source moneyline probability is not a spread-cover probability; this lane uses its own derived-spread cover-rate and ROI history gate.",
                f"- Historical gate: {derived_history.get('wins', 0)}W-{derived_history.get('losses', 0)}L-"
                f"{derived_history.get('pushes', 0)}P · cover {derived_history.get('cover_rate')} · "
                f"ROI {derived_history.get('roi')} · status={derived_history.get('status')}",
                f"- Current pre-tip derived single candidate(s): {derived_count}. "
                "They create a flat one-unit paper receipt only if still pre-tip at the immutable decision cutoff.",
            ]
        )
    if league is League.EUROLEAGUE:
        lines[3:3] = [
            "EUROLEAGUE STRICT NAMED-BOOK TICKET LANE — exact named-book pre-tip quote; calibrated p × odds > 1.00; "
            "at least 10 past settled picks and positive ROI in the matching band; aggregate consensus prices are excluded.",
            "Separate no-stake aggregate research board: euroleague_consensus_paper_YYYY-MM-DD.txt (never a ticket, receipt, bank, ROI/P&L, or real-money instruction).",
        ]
    elif league is League.EUROLEAGUE_W:
        lines.insert(
            3,
            "EUROLEAGUE WOMEN UNPRICED LANE — no validated named-book source or point-in-time "
            "price evidence; paper counts are not value, EV, or betting recommendations.",
        )
    if priced_ineligible_shadow_picks is None:
        lines.extend(
            _render_unpriced_audit_section(unpriced_audit_picks or [], timezone_name)
        )
        lines.extend(
            _render_unpriced_spread_shadow_section(unpriced_spread_shadow, timezone_name)
        )
    else:
        lines.extend(
            _render_priced_ineligible_shadow_section(
                priced_ineligible_shadow_picks,
                audited_count=priced_ineligible_audited_count,
                timezone_name=timezone_name,
            )
        )
    return "\n".join(lines)


def _render_performance(report: Dict[str, Any], timezone_name: str) -> str:
    settled_positions = int(report.get("settled_positions", report.get("settled_accas", 0)))
    pending_positions = int(report.get("pending_positions", report.get("pending_accas", 0)))
    lines = [
        "MA GOLIDE AUTO-TICKETS (ROLLING) — paper performance",
        "=" * 72,
        f"BANK     {report['paper_bank_units']:.3f} units   "
        f"x{report['paper_bank_units'] / BASE_PAPER_UNITS:.3f} of the {BASE_PAPER_UNITS:.0f}-unit base",
        f"CAPITAL  {report['paper_free_units']:.3f} free / "
        f"{report['paper_committed_units']:.3f} reserved paper units",
    ]
    record_grades = _grade_count_summary(
        {"win": report["wins"], "loss": report["losses"], "push": report["pushes"]},
        suffixes={"win": "W", "loss": "L", "push": "P"},
    )
    if report["wins"] + report["losses"]:
        lines.append(
            f"RECORD   {record_grades} "
            f"({report['hit_rate']:.1%} hit) · ROI {report['roi']:+.1%}"
        )
    elif settled_positions:
        push_summary = (
            f"{_grade_badge('push')} {report['pushes']} pushes · "
            if report["pushes"]
            else ""
        )
        lines.append(
            f"RECORD   no win/loss decisions yet · {push_summary}ROI {report['roi']:+.1%}"
        )
    else:
        lines.append("RECORD   no settled paper positions yet")
    lines.extend(
        [
            f"OPEN     {pending_positions} held paper positions · "
            f"{report['awaiting_game_legs']} awaiting games · "
            f"{report['result_unresolvable_legs']} result unresolvable",
            f"LEGS     Brier {report['selected_leg_brier_score']:.4f}"
            if report["selected_leg_brier_score"] is not None
            else "LEGS     Brier n/a",
            "",
            f"BET-DAYS, oldest first ({len(report['history'])})",
        ]
    )
    for day in report["history"]:
        returned = f"{day['returned_units']:.3f}" if day["returned_units"] is not None else "—"
        profit = f"{day['profit_units']:+.3f}" if day["profit_units"] is not None else "—"
        day_grades = _grade_count_summary(
            {"win": day["wins"], "loss": day["losses"], "push": day["pushes"]},
            suffixes={"win": "W", "loss": "L", "push": "P"},
        )
        grade_prefix = f"{day_grades} " if day_grades else ""
        lines.append(
            f"  {day['date']}  {grade_prefix}"
            f"staked {day['staked_units']:.3f} returned {returned} net {profit} "
            f"pending {day['pending_units']:.3f} bank* {day['bank_units']:.3f}"
        )
    lines.append("")
    lines.append("HELD (not yet settled; stake counted as committed)")
    held = False
    type_indexes = Counter()
    for day in report["history"]:
        daily_positions = day.get("positions")
        if daily_positions is None:
            daily_positions = day.get("accas") or []
        for position in daily_positions:
            if position["grade_status"] != "pending":
                continue
            ticket_type = position.get("ticket_type", "acca")
            ticket_type = "single" if ticket_type == "single" else "acca"
            type_indexes[ticket_type] += 1
            type_label = "SINGLE" if ticket_type == "single" else "ACCA"
            held = True
            lines.append(
                f"  {day['date']} {type_label} #{type_indexes[ticket_type]} "
                f"@{position['combined_decimal_odds']:.3f} "
                f"stake {position['stake_units']:.3f} · receipt {day['decision_id']}"
            )
            for leg in position["legs"]:
                lines.append(
                    f"      {leg['home']} vs {leg['away']} · {_selection_label(leg)} "
                    f"· {_grade_badge(leg['grade_status'])}"
                )
    if not held:
        lines.append("  None.")
    lines.extend(
        [
            "",
            "NOTES",
            "  bank* is reconstructed from latest known results by ticket date; pending is held at cost.",
            "  This is not a settlement-time equity curve or certified betting ROI.",
            "  A known losing leg settles its paper position as a loss even if other legs remain unresolved.",
            "  Push legs are removed from effective price after every other leg resolves; a single uses its exact selected odds.",
            "  Empty scans never cancel tickets. Frozen receipts are never resized or rotated.",
            "  Paper only. No real-money stakes or automatic promotion.",
            "",
        ]
    )
    return "\n".join(lines)


def _ticket_card_audit_kwargs(
    localdata: Path,
    *,
    target_date: str,
    league: League,
    settled_history: List[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    """Build the audit-only card sections from a date's regraded pick ledger."""
    path = Path(localdata) / f"picks_{target_date}.json"
    if not path.exists():
        return None
    rows = _read_json(path, [])
    if not isinstance(rows, list):
        raise ValueError(f"dated pick ledger must be a list: {path}")
    if not rows:
        return None
    logical_rows = _latest_logical_rows(rows)
    unpriced = [row for row in logical_rows if not _has_exact_selected_price(row)]
    if league == League.NCAA_M:
        priced_ineligible = _priced_ineligible_rows(logical_rows)
        return {
            "unpriced_audit_picks": None,
            "priced_ineligible_shadow_picks": _priced_ineligible_shadow_candidates(
                priced_ineligible
            ),
            "priced_ineligible_audited_count": len(priced_ineligible),
        }
    return {
        "unpriced_audit_picks": _annotate_unpriced_audit_rows(
            unpriced, settled_history
        ),
        "priced_ineligible_shadow_picks": None,
        "priced_ineligible_audited_count": 0,
        "unpriced_spread_shadow": _unpriced_spread_shadow(logical_rows, settled_history),
    }


def _has_rendered_pending_grade(path: Path) -> bool:
    """Recognise pending result labels without mistaking explanatory prose for one."""

    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"cannot read ticket card: {path}") from exc
    return any(marker in text for marker in ("⏳ PENDING", "grade=PENDING", "[PENDING]"))


def _historical_pending_ticket_cards(
    localdata: Path, *, target_date: str
) -> List[tuple[str, Path]]:
    """Return prior mutable scan cards whose audit result view needs rebuilding."""

    cards: List[tuple[str, Path]] = []
    for path in sorted(localdata.glob("auto_tickets_????-??-??.txt")):
        card_date = path.stem.removeprefix("auto_tickets_")
        if card_date < target_date and _has_rendered_pending_grade(path):
            cards.append((card_date, path))
    return cards


def _preserve_ticket_audit_section(rendered: str, existing_path: Path) -> str:
    """Keep the original audit-only section if its dated pick artifact is absent."""
    try:
        old_lines = existing_path.read_text().splitlines()
    except OSError:
        return rendered
    start = next(
        (
            index
            for index, line in enumerate(old_lines)
            if line.startswith(
                (
                    "UNPRICED AUDIT-ONLY SHADOW GROUPS",
                    "PRICED-INELIGIBLE AUDIT-ONLY SHADOW GROUPS",
                )
            )
        ),
        None,
    )
    if start is None:
        return rendered
    end = next(
        (
            index
            for index in range(start + 1, len(old_lines))
            if old_lines[index].startswith(("FROZEN AT ", "DECISION FROZEN AT "))
        ),
        len(old_lines),
    )
    audit_lines = old_lines[start:end]
    new_lines = rendered.splitlines()
    footer = next(
        (
            index
            for index, line in enumerate(new_lines)
            if line.startswith("DECISION FROZEN AT ")
        ),
        len(new_lines),
    )
    new_lines[footer:footer] = ["", *audit_lines]
    return "\n".join(new_lines)


def publish_operator_reports(
    store,
    league: League,
    market_report: Optional[Dict[str, Any]],
    localdata: Path,
    *,
    target_date: str,
    generated_utc: str,
    run_id: str,
    timezone_name: str = "Africa/Johannesburg",
    slate_timezone_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Publish current picks and grade every retained pick/ticket in localdata."""
    localdata.mkdir(parents=True, exist_ok=True)
    games = {game.game_id: game for game in store.list_games(league=league)}
    slate_timezone_name = slate_timezone_name or timezone_name
    current = [
        _pick_row(row, target_date)
        for row in deepcopy((market_report or {}).get("forecasts") or [])
    ]
    ticket_decision = _ticket_decision_window(
        games,
        league=league,
        forecast_rows=current,
        target_date=target_date,
        generated_utc=generated_utc,
        slate_timezone_name=slate_timezone_name,
    )
    euroleague_consensus_paper = (
        _publish_euroleague_consensus_paper(
            localdata=localdata,
            target_date=target_date,
            generated_utc=generated_utc,
            timezone_name=timezone_name,
            slate_timezone_name=slate_timezone_name,
            ticket_decision=ticket_decision,
            current_rows=current,
            games=games,
            market_report_available=market_report is not None,
        )
        if league is League.EUROLEAGUE
        else None
    )
    acca_report = (market_report or {}).get("accas") or {}
    candidate_funnel = _ticket_candidate_funnel(market_report)
    paper_filter = acca_report.get("paper_pick_filter") or {}
    band_details = paper_filter.get("candidate_band_details") or {}
    single_rows = list(acca_report.get("paper_single_picks") or [])
    single_details = {
        str(row.get("forecast_id")): row.get("individual_pick_roi_band")
        for row in single_rows
        if row.get("forecast_id")
    }
    single_ids = set(single_details)
    for row in current:
        pick_id = str(row["pick_id"])
        band = band_details.get(pick_id) or single_details.get(pick_id)
        if isinstance(band, dict):
            row["paper_roi_band"] = deepcopy(band)
            row["paper_roi_band_eligible"] = band.get("qualifies") is True
        elif pick_id in single_ids:
            row["paper_roi_band_eligible"] = True
    dated_path = localdata / f"picks_{target_date}.json"
    prior = _read_json(dated_path, [])
    if not isinstance(prior, list):
        raise ValueError(f"dated pick ledger must be a list: {dated_path}")
    # After an immutable fixture snapshot exists, this dated draft ledger is
    # closed for that fixture. It can still receive official grading below for
    # audit compatibility, but no later model/price observation may rewrite or
    # append to the pre-cutoff draft evidence. The lifecycle view uses the
    # snapshot itself as settlement evidence.
    pick_lifecycle_state = _load_pick_lifecycle_state(localdata, league=league.value)
    existing_pick_freezes = _load_pick_freezes(
        pick_lifecycle_state, target_date=target_date, league=league.value
    )
    frozen_game_ids = set(existing_pick_freezes)
    archivable_current = [
        row for row in current if str(row["game_id"]) not in frozen_game_ids
    ]
    merged = {str(row["pick_id"]): row for row in prior}
    merged.update({row["pick_id"]: row for row in archivable_current})
    archived = sorted(
        merged.values(),
        key=lambda row: (row["scheduled_utc"], row["game_id"], row["pick_id"]),
    )
    for row in archived:
        _set_grade(row, games)
    # picks_today is the current scan only; the dated file retains earlier
    # frozen rows even when a later scan produces no fresh forecast for them.
    for row in current:
        if row["pick_id"] in merged:
            row.update(merged[row["pick_id"]])
    current_unpriced_audit = [row for row in current if not _has_exact_selected_price(row)]
    dated_logical_rows = _latest_logical_rows(archived)
    dated_unpriced_audit = [
        row for row in dated_logical_rows if not _has_exact_selected_price(row)
    ]
    current_priced_ineligible: List[Dict[str, Any]] = []
    dated_priced_ineligible: List[Dict[str, Any]] = []
    current_priced_ineligible_shadow: List[Dict[str, Any]] = []
    dated_priced_ineligible_shadow: List[Dict[str, Any]] = []
    if league == League.NCAA_M:
        current_priced_ineligible = _priced_ineligible_rows(current)
        dated_priced_ineligible = _priced_ineligible_rows(dated_logical_rows)
        current_priced_ineligible_shadow = _priced_ineligible_shadow_candidates(
            current_priced_ineligible
        )
        dated_priced_ineligible_shadow = _priced_ineligible_shadow_candidates(
            dated_priced_ineligible
        )
    _write_json(localdata / "picks_today.json", current)
    _write_json(dated_path, archived)
    current_pick_freezes = _freeze_due_pick_fixtures(
        localdata,
        lifecycle_state=pick_lifecycle_state,
        target_date=target_date,
        league=league.value,
        rows=archived,
        generated_utc=generated_utc,
    )
    _refresh_dated_pick_reports(
        localdata,
        games=games,
        league=league.value,
        generated_utc=generated_utc,
        timezone_name=timezone_name,
    )
    audit = _write_pick_audit(
        localdata,
        target_date=target_date,
        generated_utc=generated_utc,
        timezone_name=timezone_name,
        games=games,
    )
    settled_history = list(audit.get("settled_ledger") or [])
    current_unpriced_audit = _annotate_unpriced_audit_rows(
        current_unpriced_audit, settled_history
    )
    dated_unpriced_audit = _annotate_unpriced_audit_rows(
        dated_unpriced_audit, settled_history
    )
    current_spread_shadow: Optional[Dict[str, Any]] = None
    dated_spread_shadow: Optional[Dict[str, Any]] = None
    if league != League.NCAA_M:
        current_spread_shadow = _unpriced_spread_shadow(
            _latest_logical_rows(current), settled_history
        )
        dated_spread_shadow = _unpriced_spread_shadow(dated_logical_rows, settled_history)

    # This is a generic, future-only fallback lane.  It is intentionally built
    # from the current pre-tip scan (not from dated ledgers), so settled audit
    # examples can calibrate the separate contract but can never become new
    # tickets or change prior paper-bank results.
    derived_spread_candidates, derived_spread_history = _derived_spread_ticket_candidates(
        _latest_logical_rows(current),
        settled_history,
        generated_utc=generated_utc,
        target_date=target_date,
        slate_timezone_name=slate_timezone_name,
    )
    # Keep the long-standing native candidate-funnel schema stable on cold
    # starts. Once the derived lane has any quote-eligible history (or an
    # actionable current candidate), expose its separately governed details.
    if derived_spread_history["priced_settled"] or derived_spread_candidates:
        candidate_funnel["derived_spread_ticket_candidates"] = len(derived_spread_candidates)
        candidate_funnel["derived_spread_history"] = deepcopy(derived_spread_history)

    state = _load_state(localdata)
    acca_report = (market_report or {}).get("accas") or {}
    paper_results = acca_report.get("paper_results")
    if paper_results is None:  # legacy reports expose only the strict result surface
        paper_results = acca_report.get("results") or []
    ticket_candidates = deepcopy(paper_results or [])
    if not ticket_candidates and len(single_rows) == 1:
        single_ticket = _single_ticket_candidate(single_rows[0], paper_filter)
        if single_ticket is not None:
            ticket_candidates.append(single_ticket)
    # Native selections retain their existing service-owned funnel.  Derived
    # spreads are separately governed and have already passed their own
    # frozen named-book/cover-rate/ROI contract above.
    ticket_candidates.extend(deepcopy(derived_spread_candidates))
    pre_tip_ticket_candidates = [
        candidate
        for candidate in ticket_candidates
        if isinstance(candidate, dict) and _ticket_candidate_is_pre_tip(candidate, generated_utc)
    ]
    ticket_decision["candidate_tickets_seen"] = len(ticket_candidates)
    ticket_decision["candidate_tickets_pre_tip"] = len(pre_tip_ticket_candidates)
    ticket_decision["derived_spread_candidates_seen"] = len(derived_spread_candidates)
    ticket_decision["derived_spread_history"] = deepcopy(derived_spread_history)
    ticket_decision["market_report_available"] = market_report is not None
    decision_status = ticket_decision["status"]
    allow_new_decision = market_report is not None and decision_status in {
        "due",
        "missed_pre_tip_window",
    }
    if decision_status == "due":
        decision_candidates = pre_tip_ticket_candidates
    elif decision_status == "missed_pre_tip_window":
        # A daily decision missed before the earliest tip must fail closed; it
        # may not turn a late-arriving later-game price into a fresh ticket.
        decision_candidates = []
    else:
        decision_candidates = []
    receipt = _ticket_receipt(
        localdata,
        state=state,
        league=league,
        target_date=target_date,
        captured_utc=generated_utc,
        run_id=run_id,
        accas=decision_candidates,
        ticket_decision=ticket_decision,
        allow_new_decision=allow_new_decision,
    )
    no_bet_decision = _read_no_bet_decision(state, target_date)
    scan_record = {
        "run_id": run_id,
        "captured_at": generated_utc,
        "fresh_accas": sum(row.get("ticket_type") != "single" for row in ticket_candidates),
        "fresh_single_tickets": sum(
            row.get("ticket_type") == "single" for row in ticket_candidates
        ),
        "fresh_paper_singles": candidate_funnel.get("paper_single_picks", 0),
        "fresh_derived_spread_singles": len(derived_spread_candidates),
        "receipt_present": receipt is not None,
        "ticket_decision": deepcopy(ticket_decision),
        "no_bet_decision_present": no_bet_decision is not None,
        "candidate_funnel": deepcopy(candidate_funnel),
        "unpriced_audit": _unpriced_audit_summary(current_unpriced_audit),
    }
    if current_spread_shadow is not None:
        scan_record["unpriced_spread_shadow"] = current_spread_shadow
    if league == League.NCAA_M:
        scan_record["priced_ineligible_shadow"] = _priced_ineligible_shadow_summary(
            current_priced_ineligible,
            current_priced_ineligible_shadow,
        )
    state["scans"].append(_compact_scan_record(scan_record))
    state["scans"] = state["scans"][-100:]
    withdrawal = state["frozen_by_date"].get(target_date)
    if not (isinstance(withdrawal, dict) and withdrawal.get("withdrawn") is True):
        # Sidecar withdrawal also counts for display (survives state rollback)
        if _withdrawn_sidecar_path(localdata, target_date).exists():
            try:
                sc = json.loads(_withdrawn_sidecar_path(localdata, target_date).read_text())
                withdrawal = {
                    "withdrawn": True,
                    "decision_id": sc.get("decision_id"),
                    "sidecar": True,
                }
            except (OSError, json.JSONDecodeError):
                withdrawal = {"withdrawn": True, "sidecar": True}
        else:
            withdrawal = None
    withdrawn_decision_ids = {
        str(lock.get("decision_id"))
        for lock in state["frozen_by_date"].values()
        if isinstance(lock, dict) and lock.get("withdrawn") is True and lock.get("decision_id")
    } | _withdrawn_ids_from_sidecars(localdata)
    receipts = []
    for path in sorted(localdata.glob("auto_tickets_????-??-??_decision.json")):
        candidate = _read_json(path, {})
        if not schema_matches(candidate.get("schema_version"), TICKET_SCHEMA):
            raise ValueError(f"invalid paper ticket receipt: {path}")
        if str(candidate.get("decision_id")) in withdrawn_decision_ids:
            continue
        receipts.append(candidate)
    performance = _performance(receipts, games)
    performance["generated_utc"] = generated_utc
    performance["league"] = league.value
    _write_json(
        localdata / "auto_tickets_performance.json",
        _compact_performance_report(performance),
    )
    _write_text(
        localdata / "auto_tickets_performance.txt",
        _render_performance(performance, timezone_name),
    )
    performance_by_receipt = {
        (str(row["date"]), str(row["decision_id"])): row
        for row in performance["history"]
    }
    # Decision JSON stays immutable; refresh each active receipt's dated card
    # from its graded copy so old PENDING labels do not outlive result scans.
    for frozen_receipt in receipts:
        receipt_key = (
            str(frozen_receipt["date"]),
            str(frozen_receipt["decision_id"]),
        )
        graded_receipt = performance_by_receipt.get(receipt_key)
        if graded_receipt is None:
            raise ValueError(f"graded performance missing frozen receipt {receipt_key}")
        card_path = localdata / f"auto_tickets_{receipt_key[0]}.txt"
        audit_kwargs = _ticket_card_audit_kwargs(
            localdata,
            target_date=receipt_key[0],
            league=league,
            settled_history=settled_history,
        )
        if audit_kwargs is None:
            audit_kwargs = {"unpriced_audit_picks": []}
        rendered_card = _render_ticket_card(
            graded_receipt,
            performance,
            timezone_name,
            **audit_kwargs,
        )
        if "priced_ineligible_shadow_picks" not in audit_kwargs:
            rendered_card = _preserve_ticket_audit_section(rendered_card, card_path)
        _write_text(card_path, rendered_card)

    terminal_ticket_dates = {str(frozen_receipt["date"]) for frozen_receipt in receipts}
    withdrawn_ticket_sidecars: Dict[str, Dict[str, Any]] = {}
    for sidecar_path in sorted(localdata.glob("auto_tickets_????-??-??_withdrawn.json")):
        withdrawn_date = sidecar_path.name.removeprefix("auto_tickets_").removesuffix(
            "_withdrawn.json"
        )
        sidecar = _read_json(sidecar_path, {})
        if (
            not isinstance(sidecar, dict)
            or sidecar.get("withdrawn") is not True
            or sidecar.get("target_date") != withdrawn_date
        ):
            raise ValueError(f"invalid withdrawn ticket sidecar: {sidecar_path}")
        withdrawn_ticket_sidecars[withdrawn_date] = sidecar
    # An inline no-bet lock is equally terminal for ticket creation, while its
    # human-readable card remains a mutable audit view. Rebuild every no-bet
    # card from its dated, regraded ledger so shadow grades do not remain
    # PENDING after official results arrive.
    no_bet_dates = sorted(
        str(date)
        for date, lock in state["frozen_by_date"].items()
        if isinstance(lock, dict) and lock.get("no_bet") is True
    )
    for no_bet_date in no_bet_dates:
        frozen_no_bet = _read_no_bet_decision(state, no_bet_date)
        if frozen_no_bet is None:  # pragma: no cover - guarded by the lock filter
            raise ValueError(f"frozen inline no-bet decision is missing: {no_bet_date}")
        terminal_ticket_dates.add(no_bet_date)
        audit_kwargs = _ticket_card_audit_kwargs(
            localdata,
            target_date=no_bet_date,
            league=league,
            settled_history=settled_history,
        )
        if audit_kwargs is None:
            audit_kwargs = {"unpriced_audit_picks": []}
        _write_text(
            localdata / f"auto_tickets_{no_bet_date}.txt",
            _render_empty_ticket_scan(
                league=league,
                target_date=no_bet_date,
                generated_utc=generated_utc,
                performance=performance,
                timezone_name=timezone_name,
                candidate_funnel={},
                no_bet_decision=frozen_no_bet,
                ticket_decision=frozen_no_bet.get("ticket_decision"),
                **audit_kwargs,
            ),
        )

    # A past scan that never reached a ticket/no-bet cutoff has no immutable
    # sidecar, yet may still expose audit-only shadows. It cannot be turned
    # into a retrospective decision, but recovered dated evidence can safely
    # rebuild its mutable archive section with authoritative grades.
    historical_ticket_audit_cards: List[str] = []
    for historical_date, card_path in _historical_pending_ticket_cards(
        localdata, target_date=target_date
    ):
        if historical_date in terminal_ticket_dates:
            continue
        withdrawal_sidecar = withdrawn_ticket_sidecars.get(historical_date)
        audit_kwargs = _ticket_card_audit_kwargs(
            localdata,
            target_date=historical_date,
            league=league,
            settled_history=settled_history,
        )
        # If retained artifacts cannot supply the original dated evidence, do
        # not manufacture a replacement audit card. Workflow recovery raises
        # before this point for a known pending public grade; this guard also
        # keeps direct/library callers fail-closed.
        if audit_kwargs is None:
            continue
        _write_text(
            card_path,
            _render_empty_ticket_scan(
                league=league,
                target_date=historical_date,
                generated_utc=generated_utc,
                performance=performance,
                timezone_name=timezone_name,
                candidate_funnel={},
                withdrawal=withdrawal_sidecar,
                # A withdrawal is terminal evidence of a prior decision. It
                # renders as withdrawn, never as a newly invented archive/no-bet.
                historical_archive=withdrawal_sidecar is None,
                **audit_kwargs,
            ),
        )
        historical_ticket_audit_cards.append(historical_date)

    ticket_path = str(localdata / f"auto_tickets_{target_date}.txt")
    if receipt is None:
        ticket_path = _write_text(
            localdata / f"auto_tickets_{target_date}.txt",
            _render_empty_ticket_scan(
                league=league,
                target_date=target_date,
                generated_utc=generated_utc,
                performance=performance,
                timezone_name=timezone_name,
                candidate_funnel=candidate_funnel,
                withdrawal=withdrawal,
                no_bet_decision=no_bet_decision,
                ticket_decision=ticket_decision,
                unpriced_audit_picks=dated_unpriced_audit,
                priced_ineligible_shadow_picks=(
                    dated_priced_ineligible_shadow
                    if league == League.NCAA_M
                    else None
                ),
                priced_ineligible_audited_count=(
                    len(dated_priced_ineligible)
                    if league == League.NCAA_M
                    else 0
                ),
                unpriced_spread_shadow=dated_spread_shadow,
                # The current scan feeds the pre-cutoff section of the single
                # dated card; terminal/history renders never receive it.
                named_price_draft_rows=current,
            ),
        )
    state["last_generated_utc"] = generated_utc
    state["paper_bank_units"] = performance["paper_bank_units"]
    _write_json(localdata / "auto_tickets_state.json", state)
    return {
        "paper_only": True,
        "picks_today": str(localdata / "picks_today.json"),
        "dated_picks": str(dated_path),
        "picks_text": str(localdata / f"picks_{target_date}.txt"),
        "pick_lifecycle": str(localdata / f"picks_{target_date}_lifecycle.json"),
        "pick_freezes": len(current_pick_freezes),
        "pick_lifecycle_state": str(localdata / "pick_lifecycle_state.json"),
        "rolling_audit": str(localdata / "picks_audit_rolling.json"),
        "dated_audit": str(localdata / f"picks_audit_{target_date}.md"),
        "ticket_receipt": (
            str(localdata / f"auto_tickets_{target_date}_decision.json")
            if receipt is not None
            else None
        ),
        "ticket_text": ticket_path,
        "historical_ticket_audit_cards_refreshed": historical_ticket_audit_cards,
        "ticket_state": str(localdata / "auto_tickets_state.json"),
        "performance": str(localdata / "auto_tickets_performance.json"),
        "performance_text": str(localdata / "auto_tickets_performance.txt"),
        "candidate_funnel": candidate_funnel,
        "ticket_decision": ticket_decision,
        "euroleague_consensus_paper": euroleague_consensus_paper,
        "no_bet_decision": deepcopy(no_bet_decision),
        "unpriced_audit": _unpriced_audit_summary(dated_unpriced_audit),
        "unpriced_spread_shadow": dated_spread_shadow,
        "derived_spread_ticket_lane": {
            "candidate_count": len(derived_spread_candidates),
            "history_gate": derived_spread_history,
            "future_only": True,
        },
        "priced_ineligible_shadow": (
            _priced_ineligible_shadow_summary(
                dated_priced_ineligible,
                dated_priced_ineligible_shadow,
            )
            if league == League.NCAA_M
            else None
        ),
        "counts": {
            "current_picks": len(current),
            "current_unpriced_picks": len(current_unpriced_audit),
            "dated_unpriced_audit_picks": len(dated_unpriced_audit),
            "archived_picks": len(archived),
            "graded_picks": audit["settled_picks"],
            # Backward-compatible alias; this now includes same-day finals.
            "graded_prior_picks": audit["settled_picks"],
            "published_accas": len(receipt.get("accas") or []) if receipt else 0,
            "published_singles": len(receipt.get("singles") or []) if receipt else 0,
            "settled_positions": performance["settled_positions"],
            "pending_positions": performance["pending_positions"],
            "settled_accas": performance["settled_accas"],
            "pending_accas": performance["pending_accas"],
            "settled_singles": performance["settled_singles"],
            "pending_singles": performance["pending_singles"],
            "paper_single_picks": candidate_funnel.get("paper_single_picks", 0),
            "derived_spread_ticket_candidates": len(derived_spread_candidates),
            "historical_ticket_audit_cards_refreshed": len(historical_ticket_audit_cards),
        },
    }


__all__ = ["load_settled_pick_history", "publish_operator_reports"]
