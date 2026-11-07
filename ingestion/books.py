"""Bookmaker market normalisation and merge rules.

Binding rules (master prompt section 9):

- A full-game total's **line value** maps to upcoming ``FT Score``; prices,
  spreads and over/under direction never do.
- ``totals_q1`` / period ``1q`` lines map to ``Q1``, and likewise Q2-Q4.
- A model expected total is never a book line; nothing here accepts a
  prediction object at all.
- Quarters require pre-match + intended book + intended period + main market.
- A missing market never erases a valid existing manual/book value; the
  merge decision carries the original source/timestamp forward.
- No code path derives a quarter total from a full-game/half total, spread
  or price — this module has no arithmetic that could do so, and tests
  prove it.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .models import BookSnapshot, CoverageGap, CoverageState, Market, Period
from .identity import now_utc_iso

# Providers that must never appear in a quarter-line chain, whatever they
# return. TheRundown Free has no quarter markets by contract.
PROVIDERS_BLOCKED_FROM_QUARTERS = {"therundown"}


@dataclass(frozen=True)
class BookConfig:
    """Per-satellite bookmaker selection policy."""

    provider: str
    book: Optional[str] = None  # None = provider default/first configured
    prefer_main_line: bool = True
    pre_match_only: bool = True


def american_to_decimal(a: float) -> float:
    """Convert American odds to decimal: ``1 + A/100`` if A>0 else ``1 + 100/|A|``."""
    if a is None:
        raise ValueError("american odds required")
    a = float(a)
    if a == 0:
        raise ValueError("american odds cannot be zero")
    return (1 + a / 100.0) if a > 0 else (1 + 100.0 / abs(a))


def is_plausible_decimal(price: Optional[float]) -> bool:
    return price is not None and 1.01 <= price <= 1000.0


def select_moneyline_snapshots(
    snapshots: List[BookSnapshot], config: BookConfig
) -> List[BookSnapshot]:
    """Filter to pre-match, main-line moneyline snapshots for the configured book."""
    out = []
    for s in snapshots:
        if s.market is not Market.MONEYLINE or s.period is not Period.FULL_GAME:
            continue
        if config.pre_match_only and not s.pre_match:
            continue
        if config.prefer_main_line and not s.main_line:
            continue
        if config.book and s.book and s.book.lower() != config.book.lower():
            continue
        out.append(s)
    return out


def select_total_snapshots(
    snapshots: List[BookSnapshot],
    config: BookConfig,
    period: Period = Period.FULL_GAME,
) -> List[BookSnapshot]:
    """Filter to total **line** snapshots for one exact period."""
    out = []
    for s in snapshots:
        if s.market is not Market.TOTAL or s.period is not period:
            continue
        if config.pre_match_only and not s.pre_match:
            continue
        if config.prefer_main_line and not s.main_line:
            continue
        if config.book and s.book and s.book.lower() != config.book.lower():
            continue
        if s.provider in PROVIDERS_BLOCKED_FROM_QUARTERS and period is not Period.FULL_GAME:
            # Hard guard: TheRundown (Free) can never fill a quarter cell.
            continue
        out.append(s)
    return out


def newest_snapshot(snapshots: List[BookSnapshot]) -> Optional[BookSnapshot]:
    """Newest snapshot by observed_utc (ties: first listed)."""
    best: Optional[BookSnapshot] = None
    for s in snapshots:
        if best is None or (s.observed_utc or "") > (best.observed_utc or ""):
            best = s
    return best


# ---------------------------------------------------------------------------
# Merge policy (upcoming Q1..Q4 / FT Score cells)
# ---------------------------------------------------------------------------


@dataclass
class ExistingCellValue:
    """What currently sits in the target cell, with its known provenance."""

    raw: Optional[str] = None
    source: Optional[str] = None  # provider name or "manual"
    observed_utc: Optional[str] = None
    is_manual: bool = False


@dataclass
class MergeDecision:
    action: str  # "keep_existing" | "write" | "blank_with_gap"
    value: Optional[float] = None
    source: Optional[str] = None
    observed_utc: Optional[str] = None
    gap: Optional[CoverageGap] = None


def merge_book_cell(
    period: Period,
    incoming: List[BookSnapshot],
    existing: Optional[ExistingCellValue],
    config: BookConfig,
    now_utc: Optional[str] = None,
    game_id: Optional[str] = None,
    staleness_hours: float = 24.0,
) -> MergeDecision:
    """Decide what belongs in one upcoming book-line cell.

    Policy:
    1. Fresh incoming pre-match main-line snapshot for the configured
       book/period wins and is written with its source/timestamp.
    2. Otherwise an existing non-blank value is PRESERVED (never silently
       erased by an empty refresh) while it stays inside the staleness
       window relative to tip; manual values carry ``source='manual'``.
    3. Otherwise blank with an explicit coverage gap.
    """
    period_label = period.value
    if period is not Period.FULL_GAME:
        candidates = select_total_snapshots(incoming, config, period)
    else:
        candidates = select_total_snapshots(incoming, config, Period.FULL_GAME)
    chosen = newest_snapshot(candidates)
    if chosen is not None and chosen.line is not None:
        return MergeDecision(
            action="write",
            value=float(chosen.line),
            source=chosen.provider,
            observed_utc=chosen.observed_utc,
        )

    if existing is not None and str(existing.raw or "").strip() != "":
        existing_valid = True
        if not existing.is_manual and existing.observed_utc:
            # Book values age out relative to tip only via staleness policy;
            # pre-match lines are dropped when older than the window.
            existing_valid = _within_staleness(existing.observed_utc, now_utc, staleness_hours)
        if existing_valid:
            return MergeDecision(
                action="keep_existing",
                value=_parse_float(existing.raw),
                source=existing.source or ("manual" if existing.is_manual else None),
                observed_utc=existing.observed_utc,
            )

    return MergeDecision(
        action="blank_with_gap",
        value=None,
        gap=CoverageGap(
            game_id=game_id,
            required_field=f"book_{period_label.lower()}_line",
            reason=CoverageState.MISSING if incoming else CoverageState.UNSUPPORTED,
            provider=config.provider,
            detail=f"no fresh pre-match {period_label} total line for configured book",
            observed_utc=now_utc or now_utc_iso(),
        ),
    )


def _within_staleness(observed_utc: str, now_utc: Optional[str], hours: float) -> bool:
    from datetime import timedelta

    from .identity import parse_utc

    base = parse_utc(now_utc) if now_utc else parse_utc(now_utc_iso())
    return (base - parse_utc(observed_utc)) <= timedelta(hours=hours)


def _parse_float(raw: Optional[str]) -> Optional[float]:
    try:
        return float(str(raw).strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Per-quarter coverage accounting
# ---------------------------------------------------------------------------


def quarter_coverage(snapshots: List[BookSnapshot], config: BookConfig) -> Dict[str, int]:
    """Counts of games with each quarter line present, from canonical snapshots.

    Mirrors the writer's selection filters (pre-match, main line, configured
    book, quarter-blocked providers) so coverage counts cannot over-state
    what the sheet would actually receive. Partial quarters are allowed and
    never backfilled.
    """
    per_period = {p.value: set() for p in (Period.Q1, Period.Q2, Period.Q3, Period.Q4)}
    for s in snapshots:
        if s.market is Market.TOTAL and s.period in per_period and s.line is not None:
            if s.provider in PROVIDERS_BLOCKED_FROM_QUARTERS:
                continue
            if config.pre_match_only and not s.pre_match:
                continue
            if config.prefer_main_line and not s.main_line:
                continue
            if config.book and s.book and s.book.lower() != config.book.lower():
                continue
            per_period[s.period.value].add(s.game_id)
    return {p: len(ids) for p, ids in per_period.items()}
