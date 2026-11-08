"""Competition-specific period semantics shared by acquisition and research."""

from __future__ import annotations

from typing import Tuple

from .models import League, Period, PeriodKind

QUARTER_PERIODS: Tuple[Period, ...] = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)
HALF_PERIODS: Tuple[Period, ...] = (Period.H1, Period.H2)
NCAA_LEAGUES = frozenset({League.NCAA_M, League.NCAA_W})


def native_periods(league: League) -> Tuple[Period, ...]:
    """Return native betting periods, which can differ from stored results.

    NCAA women can have H1/H2 book markets even though modern official results
    are four quarters; grading derives those halves from Q1+Q2 and Q3+Q4.
    """
    return HALF_PERIODS if league in NCAA_LEAGUES else QUARTER_PERIODS


def period_kind(league: League) -> PeriodKind:
    """Default result format when a payload has no explicit regulation format.

    NCAA men use two halves. Modern NCAA women use four quarters; an explicit
    two-period declaration still preserves older NCAAW payloads in the ESPN
    adapter. Other supported basketball leagues default to quarters.
    """
    return PeriodKind.HALVES if league is League.NCAA_M else PeriodKind.QUARTERS


def line_periods(league: League) -> Tuple[Period, ...]:
    """Full game plus the native periods for which line forecasts are defined."""
    return (Period.FULL_GAME, *native_periods(league))


def supports_highest_quarter(league: League) -> bool:
    return league not in NCAA_LEAGUES


__all__ = [
    "HALF_PERIODS",
    "NCAA_LEAGUES",
    "QUARTER_PERIODS",
    "line_periods",
    "native_periods",
    "period_kind",
    "supports_highest_quarter",
]
