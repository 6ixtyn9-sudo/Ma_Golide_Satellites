"""Competition-specific period semantics shared by acquisition and research."""

from __future__ import annotations

from typing import Tuple

from .models import League, Period, PeriodKind

QUARTER_PERIODS: Tuple[Period, ...] = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)
HALF_PERIODS: Tuple[Period, ...] = (Period.H1, Period.H2)
NCAA_LEAGUES = frozenset({League.NCAA_M, League.NCAA_W})

# Explicit data policy rather than an implicit "all new leagues are quarters"
# fallback.  The researched expansion competitions use four regulation
# quarters; an adapter that encounters another official format must add that
# league here before normalising/settling it.
RESULT_PERIOD_POLICY = {
    League.NCAA_M: PeriodKind.HALVES,
    League.NCAA_W: PeriodKind.QUARTERS,
    League.NBA: PeriodKind.QUARTERS,
    League.WNBA: PeriodKind.QUARTERS,
    League.EUROLEAGUE: PeriodKind.QUARTERS,
    League.EUROLEAGUE_W: PeriodKind.QUARTERS,
    League.CBA: PeriodKind.QUARTERS,
    League.EUROCUP: PeriodKind.QUARTERS,
    League.BASKETBALL_CHAMPIONS_LEAGUE: PeriodKind.QUARTERS,
    League.NBA_G_LEAGUE: PeriodKind.QUARTERS,
    League.BASKETBALL_AFRICA_LEAGUE: PeriodKind.QUARTERS,
    League.AUSTRALIA_NBL: PeriodKind.QUARTERS,
    League.SPAIN_ACB: PeriodKind.QUARTERS,
    League.FRANCE_LNB_PRO_A: PeriodKind.QUARTERS,
    League.GERMANY_BBL: PeriodKind.QUARTERS,
    League.ITALY_LBA: PeriodKind.QUARTERS,
    League.TURKEY_BSL: PeriodKind.QUARTERS,
    League.GREECE_BASKET_LEAGUE: PeriodKind.QUARTERS,
    League.ABA_LEAGUE: PeriodKind.QUARTERS,
    League.ISRAEL_PREMIER_LEAGUE: PeriodKind.QUARTERS,
    League.JAPAN_B_LEAGUE: PeriodKind.QUARTERS,
    League.JAPAN_B_ONE: PeriodKind.QUARTERS,
    League.SOUTH_KOREA_KBL: PeriodKind.QUARTERS,
    League.BRAZIL_NBB: PeriodKind.QUARTERS,
    League.ARGENTINA_LIGA_NACIONAL: PeriodKind.QUARTERS,
}


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
    try:
        return RESULT_PERIOD_POLICY[league]
    except KeyError as exc:
        raise ValueError(f"missing explicit result-period policy for {league.value}") from exc


def line_periods(league: League) -> Tuple[Period, ...]:
    """Full game plus the native periods for which line forecasts are defined."""
    return (Period.FULL_GAME, *native_periods(league))


def supports_highest_quarter(league: League) -> bool:
    return league not in NCAA_LEAGUES


__all__ = [
    "HALF_PERIODS",
    "NCAA_LEAGUES",
    "QUARTER_PERIODS",
    "RESULT_PERIOD_POLICY",
    "line_periods",
    "native_periods",
    "period_kind",
    "supports_highest_quarter",
]
