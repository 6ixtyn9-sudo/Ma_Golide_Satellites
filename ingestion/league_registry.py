"""Single registry for basketball competition operational contracts.

The registry deliberately separates canonical competition identity from source
capability. Adding a league here makes it available to typed storage, CLI,
workflow scheduling, forecasting, and lifecycle state without pretending that
an unverified public endpoint is an official schedule/result feed.

A ``source_unavailable`` entry is therefore an explicit supported *unpriced
forecast lifecycle* lane: existing canonical games/history can be forecast and
settled, but automated acquisition reports a CoverageGap until a compliant
official adapter is implemented. It is not a shadow string or a fabricated
empty schedule.

Price evidence policy also lives here. Provider adapters report observations;
the pipeline stamps the registry contract onto every stored observation. An
adapter cannot promote its own quote to ticket-eligible evidence merely by
omitting evidence fields.
"""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Dict, Iterable, Tuple

from .models import (
    League,
    OperationalValidation,
    PriceProvenance,
    SourcePermission,
)


@dataclass(frozen=True)
class PriceSourcePolicy:
    """Evidence contract for one provider in one league.

    ``acquire`` distinguishes an explicit price-feed adapter from a schedule
    provider that happens to expose embedded odds. Both are still stamped, but
    only the former is called by ``_acquire_book_lines``. Unknown provider
    identities are always downgraded to unresolved/unverified evidence.
    """

    provider: str
    price_provenance: PriceProvenance
    source_permission: SourcePermission
    operational_validation: OperationalValidation
    acquire: bool = True
    detail: str = ""


_UNVERIFIED_NAMED_BOOK = PriceSourcePolicy(
    provider="unknown",
    price_provenance=PriceProvenance.NAMED_BOOK,
    source_permission=SourcePermission.UNRESOLVED,
    operational_validation=OperationalValidation.UNVERIFIED,
    acquire=False,
    detail="No explicit registry evidence policy exists for this provider.",
)


@dataclass(frozen=True)
class LeagueSpec:
    """Stable operational contract for one canonical basketball competition."""

    league: League
    label: str
    state_key: str
    timezone: str
    acquisition_timezone: str
    schedule_source: str
    results_source: str
    gate_schedule_source: str
    price_sources: Tuple[PriceSourcePolicy, ...] = ()
    source_detail: str = ""

    @property
    def book_providers(self) -> Tuple[str, ...]:
        """Configured adapter providers, preserving the historic public API."""

        return tuple(policy.provider for policy in self.price_sources if policy.acquire)

    @property
    def automated_schedule_supported(self) -> bool:
        return self.schedule_source != "source_unavailable"

    @property
    def automated_results_supported(self) -> bool:
        return self.results_source != "source_unavailable"


# Named provider policies are deliberately reusable contracts, not a copy of
# per-adapter defaults. Every quote preserves all three evidence axes for audit;
# the paper-ticket lane admits exact named-book quotes unless the source is
# explicitly prohibited, while aggregate consensus evidence remains research-only.
def _validated_named_book(provider: str) -> PriceSourcePolicy:
    return PriceSourcePolicy(
        provider,
        PriceProvenance.NAMED_BOOK,
        SourcePermission.CLEARED,
        OperationalValidation.VALIDATED,
    )


def _probationary_named_book(provider: str) -> PriceSourcePolicy:
    return PriceSourcePolicy(
        provider,
        PriceProvenance.NAMED_BOOK,
        SourcePermission.CLEARED,
        OperationalValidation.PROBATION,
    )


def _unresolved_named_book(provider: str) -> PriceSourcePolicy:
    return PriceSourcePolicy(
        provider,
        PriceProvenance.NAMED_BOOK,
        SourcePermission.UNRESOLVED,
        OperationalValidation.PROBATION,
    )


def _aggregate_research_only(provider: str, *, detail: str = "") -> PriceSourcePolicy:
    return PriceSourcePolicy(
        provider,
        PriceProvenance.AGGREGATE_CONSENSUS,
        SourcePermission.UNRESOLVED,
        OperationalValidation.PROBATION,
        detail=detail,
    )


_ABA_BETEXPLORER_RESEARCH_ONLY = _aggregate_research_only(
    "betexplorer",
    detail=(
        "ABA League representative average price observation for research only; "
        "never a named-book ticket, receipt, bank, ROI, or execution quote."
    ),
)


def _euroleague_aggregate_consensus_paper_only(provider: str) -> PriceSourcePolicy:
    """Allow a visibly separate no-stake research board, never ticket pricing.

    This does not relax any evidence axis. The operator may preserve the exact
    representative aggregate as a EuroLeague consensus-paper observation, but
    source permission/validation remain unresolved/probationary and strict
    tickets, receipts, bank, ROI, and real-money execution remain unavailable.
    """

    return PriceSourcePolicy(
        provider,
        PriceProvenance.AGGREGATE_CONSENSUS,
        SourcePermission.UNRESOLVED,
        OperationalValidation.PROBATION,
        detail=(
            "EuroLeague aggregate-consensus research-board price only; never a "
            "named-book ticket or execution quote."
        ),
    )


# ESPN scoreboard odds identify a named bookmaker but are a relayed observation,
# not an approved price-feed lane. Preserve their unresolved/unverified labels
# for audit while allowing the ordinary paper-ticket model and ROI gates to
# decide whether an exact current quote is usable.
_ESPN_RELAY_ONLY = PriceSourcePolicy(
    "espn",
    PriceProvenance.NAMED_BOOK,
    SourcePermission.UNRESOLVED,
    OperationalValidation.UNVERIFIED,
    acquire=False,
    detail="Named bookmaker observation relayed by ESPN; unresolved/unverified audit metadata.",
)


# Keep this order user-facing and workflow-stable. The first seven entries are
# pre-existing routes; all later entries are the researched final expansion,
# intentionally added in one canonical registry rather than in 17 bespoke
# pipeline branches.
_SPECS: Tuple[LeagueSpec, ...] = (
    LeagueSpec(
        League.NBA,
        "NBA",
        "nba",
        "America/New_York",
        "America/New_York",
        "espn",
        "espn",
        "espn",
        (_ESPN_RELAY_ONLY, *(_validated_named_book(provider) for provider in ("therundown", "sportsgameodds", "theoddsapi"))),
    ),
    LeagueSpec(
        League.WNBA,
        "WNBA",
        "wnba",
        "America/New_York",
        "America/New_York",
        "espn",
        "espn",
        "espn",
        (_ESPN_RELAY_ONLY, _validated_named_book("theoddsapi")),
    ),
    LeagueSpec(
        League.EUROLEAGUE,
        "EuroLeague",
        "euroleague",
        "Europe/Paris",
        "Europe/Paris",
        "euroleague",
        "euroleague",
        "euroleague",
        (_euroleague_aggregate_consensus_paper_only("betexplorer"),),
    ),
    LeagueSpec(
        League.EUROLEAGUE_W,
        "EuroLeague Women",
        "euroleague_w",
        "Europe/Paris",
        "Europe/Paris",
        "fiba_elw",
        "fiba_elw",
        "fiba_elw",
    ),
    LeagueSpec(
        League.CBA,
        "Chinese CBA",
        "cba",
        "Asia/Shanghai",
        "Asia/Shanghai",
        "cba",
        "cba",
        "cba",
        (_aggregate_research_only("betexplorer"),),
    ),
    LeagueSpec(
        League.NCAA_M,
        "NCAA men",
        "ncaa_m",
        "America/New_York",
        "America/New_York",
        "espn",
        "espn",
        "espn",
        (_ESPN_RELAY_ONLY, _validated_named_book("theoddsapi")),
    ),
    LeagueSpec(
        League.NCAA_W,
        "NCAA women",
        "ncaa_w",
        "America/New_York",
        "America/New_York",
        "espn",
        "espn",
        "espn",
        (_ESPN_RELAY_ONLY, _validated_named_book("theoddsapi")),
    ),
    LeagueSpec(League.EUROCUP, "EuroCup", "eurocup", "Europe/Paris", "Europe/Paris", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.BASKETBALL_CHAMPIONS_LEAGUE, "Basketball Champions League", "basketball_champions_league", "Europe/Paris", "Europe/Paris", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    # ESPN's public NBA G League scoreboard and event-summary route were
    # captured with stable event/team identity, UTC tips, terminal status, and
    # reconciled Q1–Q4 finals. This grants canonical schedule/result handling
    # only; ESPN BPI, deep core-index history, and named-book pricing remain
    # separately unvalidated and fail closed.
    LeagueSpec(League.NBA_G_LEAGUE, "NBA G League", "nba_g_league", "America/New_York", "America/New_York", "espn", "espn", "espn", (_ESPN_RELAY_ONLY,), "ESPN public basketball/nba-development scoreboard and summary are validated for date-bounded schedule/results; forecasts are empirical-history-only until an independent pre-tip prediction source is validated."),
    LeagueSpec(League.BASKETBALL_AFRICA_LEAGUE, "Basketball Africa League", "basketball_africa_league", "Africa/Johannesburg", "Africa/Johannesburg", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    # The Odds API NBL sport key is a documented integration lead. It remains
    # probationary until a credentialed regional/bookmaker smoke proves actual
    # account coverage, and it never supplies canonical schedule/results.
    LeagueSpec(League.AUSTRALIA_NBL, "Australia NBL", "australia_nbl", "Australia/Sydney", "Australia/Sydney", "source_unavailable", "source_unavailable", "source_unavailable", (_probationary_named_book("theoddsapi"),), "No compliant official schedule/result adapter has been validated; The Odds API price lane is probationary."),
    LeagueSpec(League.SPAIN_ACB, "Spain ACB", "spain_acb", "Europe/Madrid", "Europe/Madrid", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.FRANCE_LNB_PRO_A, "France LNB Pro A", "france_lnb_pro_a", "Europe/Paris", "Europe/Paris", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.GERMANY_BBL, "Germany BBL", "germany_bbl", "Europe/Berlin", "Europe/Berlin", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.ITALY_LBA, "Italy LBA", "italy_lba", "Europe/Rome", "Europe/Rome", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.TURKEY_BSL, "Turkey BSL", "turkey_bsl", "Europe/Istanbul", "Europe/Istanbul", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.GREECE_BASKET_LEAGUE, "Greece Basket League", "greece_basket_league", "Europe/Athens", "Europe/Athens", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    # The official ABA League calendar/detail route is the canonical fixture,
    # result, and history source. BetExplorer is deliberately separate: its
    # aggregate observations remain research-only and cannot price tickets or
    # support ROI, receipts, banks, or execution.
    LeagueSpec(League.ABA_LEAGUE, "ABA League", "aba_league", "Europe/Belgrade", "Europe/Belgrade", "aba_liga", "aba_liga", "aba_liga", (_ABA_BETEXPLORER_RESEARCH_ONLY,), source_detail="Official aba-liga.com calendar/detail adapter: season/game scope, official team IDs, Europe/Belgrade tips, and period-reconciled finals are required. Source automation/reuse terms remain unreviewed; BetExplorer averages remain research-only."),
    LeagueSpec(League.ISRAEL_PREMIER_LEAGUE, "Israel Premier League", "israel_premier_league", "Asia/Jerusalem", "Asia/Jerusalem", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.JAPAN_B_LEAGUE, "Japan B.League", "japan_b_league", "Asia/Tokyo", "Asia/Tokyo", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.SOUTH_KOREA_KBL, "South Korea KBL", "south_korea_kbl", "Asia/Seoul", "Asia/Seoul", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.BRAZIL_NBB, "Brazil NBB", "brazil_nbb", "America/Sao_Paulo", "America/Sao_Paulo", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
    LeagueSpec(League.ARGENTINA_LIGA_NACIONAL, "Argentina Liga Nacional", "argentina_liga_nacional", "America/Argentina/Buenos_Aires", "America/Argentina/Buenos_Aires", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated."),
)

SPECS: Tuple[LeagueSpec, ...] = _SPECS
_BY_LEAGUE: Dict[League, LeagueSpec] = {spec.league: spec for spec in SPECS}

# Compatibility spellings are accepted only at the CLI/config boundary. The
# canonical enum value is persisted everywhere else.
_ALIASES: Dict[str, League] = {
    "EUROLEAGUE": League.EUROLEAGUE,
    "EUROLEAGUE_WOMEN": League.EUROLEAGUE_W,
    "BCL": League.BASKETBALL_CHAMPIONS_LEAGUE,
    "BAL": League.BASKETBALL_AFRICA_LEAGUE,
    "NBL": League.AUSTRALIA_NBL,
    "ACB": League.SPAIN_ACB,
    "LNB_PRO_A": League.FRANCE_LNB_PRO_A,
    "BBL": League.GERMANY_BBL,
    "LBA": League.ITALY_LBA,
    "BSL": League.TURKEY_BSL,
    "KBL": League.SOUTH_KOREA_KBL,
    "NBB": League.BRAZIL_NBB,
    "LIGA_NACIONAL": League.ARGENTINA_LIGA_NACIONAL,
}


def league_spec(league: League) -> LeagueSpec:
    """Return the complete contract for a canonical league enum."""

    return _BY_LEAGUE[league]


def price_evidence_policy(league: League, provider: str) -> PriceSourcePolicy:
    """Return the explicit evidence contract for an observed provider quote.

    The fail-closed fallback deliberately does not inherit a league's most
    permissive provider policy: a new or misspelled source remains visible but
    cannot produce a ticket-priced forecast until the registry names it.
    """

    normalized = str(provider or "").strip().lower()
    for policy in league_spec(league).price_sources:
        if policy.provider.lower() == normalized:
            return policy
    return _UNVERIFIED_NAMED_BOOK


def parse_league(value: str | League) -> League:
    """Parse a public CLI/config spelling without admitting unknown strings."""

    if isinstance(value, League):
        return value
    normalized = re.sub(r"[^A-Z0-9]+", "_", str(value or "").strip().upper()).strip("_")
    normalized = _ALIASES.get(normalized, None) or normalized
    if isinstance(normalized, League):
        return normalized
    try:
        return League[normalized]
    except KeyError:
        try:
            return League(normalized)
        except ValueError as exc:
            choices = ", ".join(spec.league.value for spec in SPECS)
            raise ValueError(f"unknown league {value!r}; canonical values: {choices}") from exc


def cli_league_choices(*, include_legacy_display: bool = True) -> Tuple[str, ...]:
    """Return argparse choices, keeping the established ``EuroLeague`` spelling."""

    values = [spec.league.value for spec in SPECS]
    if include_legacy_display:
        values.insert(values.index(League.EUROLEAGUE.value), "EuroLeague")
    return tuple(values)


def source_unavailable_specs() -> Iterable[LeagueSpec]:
    return (spec for spec in SPECS if not spec.automated_schedule_supported)


__all__ = [
    "LeagueSpec",
    "PriceSourcePolicy",
    "SPECS",
    "cli_league_choices",
    "league_spec",
    "parse_league",
    "price_evidence_policy",
    "source_unavailable_specs",
]
