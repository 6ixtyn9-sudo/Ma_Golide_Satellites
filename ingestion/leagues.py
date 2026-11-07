"""Competition-specific period semantics shared by acquisition and research."""

from __future__ import annotations

from typing import Tuple

from .models import League, Period, PeriodKind

QUARTER_PERIODS: Tuple[Period, ...] = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)
HALF_PERIODS: Tuple[Period, ...] = (Period.H1, Period.H2)
NCAA_LEAGUES = frozenset({League.NCAA_M, League.NCAA_W})


def native_periods(league: League) -> Tuple[Period, ...]:
    """Return native regulation periods; never turn NCAA halves into quarters."""
    return HALF_PERIODS if league in NCAA_LEAGUES else QUARTER_PERIODS


def period_kind(league: League) -> PeriodKind:
    return PeriodKind.HALVES if league in NCAA_LEAGUES else PeriodKind.QUARTERS


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
