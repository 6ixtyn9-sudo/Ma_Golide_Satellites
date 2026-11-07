"""Bookmaker market filtering and chronological evidence selection.

Predictions and bookmaker observations are separate domain records. A model
expected total is never accepted as a book line, and no function in this module
derives one period's line from another period.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

from .identity import parse_utc
from .models import BookSnapshot, Market, Period

# TheRundown Free has no quarter markets by contract. Keep a second guard here
# even if an adapter is changed incorrectly.
PROVIDERS_BLOCKED_FROM_QUARTERS = {"therundown"}


@dataclass(frozen=True)
class BookConfig:
    """Optional provider/bookmaker filter for observed market evidence."""

    provider: Optional[str] = None
    book: Optional[str] = None
    prefer_main_line: bool = True
    pre_match_only: bool = True

    def matches_source(self, snapshot: BookSnapshot) -> bool:
        provider = (self.provider or "").strip().lower()
        if provider and provider != "none" and snapshot.provider.lower() != provider:
            return False
        if self.book and snapshot.book and snapshot.book.lower() != self.book.lower():
            return False
        return True


def american_to_decimal(value: float) -> float:
    """Convert a nonzero American price to decimal odds."""

    if value is None:
        raise ValueError("american odds required")
    value = float(value)
    if value == 0:
        raise ValueError("american odds cannot be zero")
    return 1 + value / 100.0 if value > 0 else 1 + 100.0 / abs(value)


def is_plausible_decimal(price: Optional[float]) -> bool:
    return price is not None and 1.01 <= price <= 1000.0


def _matches_common(snapshot: BookSnapshot, config: BookConfig) -> bool:
    if config.pre_match_only and not snapshot.pre_match:
        return False
    if config.prefer_main_line and not snapshot.main_line:
        return False
    return config.matches_source(snapshot)


def select_moneyline_snapshots(
    snapshots: List[BookSnapshot],
    config: BookConfig,
    period: Period = Period.FULL_GAME,
) -> List[BookSnapshot]:
    """Return moneyline/1X2 observations for exactly one period."""

    return [
        snapshot
        for snapshot in snapshots
        if snapshot.market is Market.MONEYLINE
        and snapshot.period is period
        and _matches_common(snapshot, config)
    ]


def select_total_snapshots(
    snapshots: List[BookSnapshot],
    config: BookConfig,
    period: Period = Period.FULL_GAME,
) -> List[BookSnapshot]:
    """Return observed total *lines* for exactly one period.

    Snapshots without a line are not candidates. This prevents a newer partial
    payload from hiding an older complete observation merely by timestamp.
    """

    return [
        snapshot
        for snapshot in snapshots
        if snapshot.market is Market.TOTAL
        and snapshot.period is period
        and snapshot.line is not None
        and _matches_common(snapshot, config)
        and not (
            period is not Period.FULL_GAME
            and snapshot.provider in PROVIDERS_BLOCKED_FROM_QUARTERS
        )
    ]


def select_spread_snapshots(
    snapshots: List[BookSnapshot],
    config: BookConfig,
    period: Period,
) -> List[BookSnapshot]:
    """Return observed spreads for one period; ``line`` is home-oriented."""
    return [
        snapshot
        for snapshot in snapshots
        if snapshot.market is Market.SPREAD
        and snapshot.period is period
        and snapshot.line is not None
        and _matches_common(snapshot, config)
        and not (
            period is not Period.FULL_GAME
            and snapshot.provider in PROVIDERS_BLOCKED_FROM_QUARTERS
        )
    ]


def newest_snapshot(snapshots: List[BookSnapshot]) -> Optional[BookSnapshot]:
    """Return the chronologically newest observation, not lexicographically newest."""

    if not snapshots:
        return None
    return max(snapshots, key=lambda snapshot: parse_utc(snapshot.observed_utc))


def quarter_coverage(
    snapshots: List[BookSnapshot], config: BookConfig
) -> Dict[str, int]:
    """Count distinct games with an observed line in each exact quarter."""

    counts: Dict[str, int] = {}
    for period in (Period.Q1, Period.Q2, Period.Q3, Period.Q4):
        counts[period.value] = len(
            {
                snapshot.game_id
                for snapshot in select_total_snapshots(snapshots, config, period)
            }
        )
    return counts
