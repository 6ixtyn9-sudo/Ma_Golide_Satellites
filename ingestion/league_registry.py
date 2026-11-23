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
        """Configured price adapters plus the bounded direct BetExplorer fallback.

        The fallback is universal because it discovers only the canonical
        scheduled games passed by the pipeline; unsupported schedule lanes
        still have no candidate game to fetch.  Keeping it here rather than
        inheriting a different league's bookmaker contract makes its distinct
        direct-public-page provenance visible in every league report.
        """

        configured = tuple(policy.provider for policy in self.price_sources if policy.acquire)
        return configured if "betexplorer_direct" in configured else (*configured, "betexplorer_direct")

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


# Operator-approved, bounded public-page fallback.  It reads a named bookmaker
# from BetExplorer's own comparison row and keeps its direct-page acquisition
# posture explicit rather than promoting it to a cleared API integration.
_BETEXPLORER_DIRECT_FALLBACK = _unresolved_named_book("betexplorer_direct")


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


def _euroleague_aggregate_research_only(provider: str) -> PriceSourcePolicy:
    """Retain aggregate observations for ordinary pick reporting, never tickets."""

    return PriceSourcePolicy(
        provider,
        PriceProvenance.AGGREGATE_CONSENSUS,
        SourcePermission.UNRESOLVED,
        OperationalValidation.PROBATION,
        detail=(
            "EuroLeague representative aggregate observation for research and "
            "pick reporting only; never a named-book ticket or execution quote."
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
        (_ESPN_RELAY_ONLY, _validated_named_book("theoddsapi")),
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
        (_euroleague_aggregate_research_only("betexplorer"),),
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
    # EuroCup shares the EuroLeague Basketball host and route shape with the
    # already-validated EuroLeague adapter; only the competition letter differs
    # (U instead of E). Two independent public clients document that contract,
    # but Ma Golide still retains no first-party capture for EuroCup itself, so
    # the lane is labelled as a referenced route contract rather than a live
    # capture. Pricing stays on the universal direct BetExplorer fallback.
    LeagueSpec(
        League.EUROCUP,
        "EuroCup",
        "eurocup",
        "Europe/Paris",
        "Europe/Paris",
        "euroleague",
        "euroleague",
        "euroleague",
        source_detail=(
            "EuroCup schedule/results use the validated EuroLeague Basketball adapter "
            "with competition code U (api-live.euroleague.net v2 competitions/{E|U}); "
            "the route shape is referenced from the same provider host and two "
            "independent public clients, with no Ma Golide first-party capture retained."
        ),
    ),
    # BCL runs on the FIBA GDAP platform the EuroLeague Women lane already reads,
    # so this row reuses that contract module (header name, public fallback key,
    # UTC parser, status vocabulary) instead of restating it. Competition 209123;
    # the schedule index carries totals but no quarter split.
    LeagueSpec(League.BASKETBALL_CHAMPIONS_LEAGUE, "Basketball Champions League", "basketball_champions_league", "Europe/Paris", "Europe/Paris", "bcl", "bcl", "bcl", source_detail="FIBA GDAP getgdapgamesbycompetitionid?gdapCompetitionId=209123 (whole season in one call, date-filtered client-side), gated by the site-shipped Ocp-Apim-Subscription-Key. One verbatim row vendored; settled rows are totals-only (no boxscore route captured).", price_sources=(_aggregate_research_only("betexplorer"),)),
    # ESPN's public NBA G League scoreboard and event-summary route were
    # captured with stable event/team identity, UTC tips, terminal status, and
    # reconciled Q1–Q4 finals. This grants canonical schedule/result handling
    # only; ESPN BPI, deep core-index history, and named-book pricing remain
    # separately unvalidated and fail closed.
    LeagueSpec(League.NBA_G_LEAGUE, "NBA G League", "nba_g_league", "America/New_York", "America/New_York", "espn", "espn", "espn", (_ESPN_RELAY_ONLY,), "ESPN public basketball/nba-development scoreboard and summary are validated for date-bounded schedule/results; forecasts are empirical-history-only until an independent pre-tip prediction source is validated."),
    # BAL is not on the FIBA GDAP service (competition id 48286 answers an empty
    # array), but bal.nba.com server-renders the whole season inside its Next.js
    # __NEXT_DATA__ blob with UTC tips, arena and team ids. One captured game
    # object is vendored under tests/fixtures/bal/captures.
    LeagueSpec(League.BASKETBALL_AFRICA_LEAGUE, "Basketball Africa League", "basketball_africa_league", "Africa/Johannesburg", "Africa/Johannesburg", "bal", "bal", "bal", source_detail="bal.nba.com/schedule renders the whole season in its __NEXT_DATA__ blob (props.pageProps.gameDates[].games[]), keyless, date-filtered client-side on gameDateTimeUTC. GDAP competition 48286 returns an empty array, so the FIBA route is a dead end. Settled rows are totals-only (no quarter split in this body)."),
    # The Odds API NBL sport key is a documented integration lead. It remains
    # probationary until a credentialed regional/bookmaker smoke proves actual
    # account coverage, and it never supplies canonical schedule/results.
    LeagueSpec(League.AUSTRALIA_NBL, "Australia NBL", "australia_nbl", "Australia/Sydney", "Australia/Sydney", "espn", "espn", "espn", (_probationary_named_book("theoddsapi"),), "Schedule/results ride the existing ESPN relay on the league's own competition route (basketball/nbl, league UID 55), verified by an operator probe on 2026-11-23 that returned three NBL games for 2026-11-21; no first-party body is retained, so live behaviour remains monitored rather than assumed. The Odds API price lane stays probationary."),
    # The ACB's own JSON API is a referenced route contract (SportsDataverse
    # ``acb`` module, host api2.acb.com): a season calendar plus a per-game
    # header that carries names, status, final score and quarterScores[].
    # Both routes require the site's public ``X-APIKEY`` header; without it the
    # calendar answered but every match header returned 401/403, which is why
    # this lane recorded zero games for 728 runs. The key is a value shipped in
    # acb.com's own bundle (site-scoped, not an account), overridable through
    # $ACB_API_KEY.
    # Ma Golide retains no first-party capture; pricing stays on the universal
    # direct BetExplorer fallback.
    LeagueSpec(
        League.SPAIN_ACB,
        "Spain ACB",
        "spain_acb",
        "Europe/Madrid",
        "Europe/Madrid",
        "acb",
        "acb",
        "acb",
        source_detail=(
            "Liga Endesa schedule/results use the ACB JSON API "
            "(seasondata/Calendar/season-calendar + matchdata/MatchHeader/match-header); "
            "route contract referenced from SportsDataverse documentation, no Ma Golide "
            "first-party capture retained."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
    ),
    LeagueSpec(League.FRANCE_LNB_PRO_A, "France LNB Pro A", "france_lnb_pro_a", "Europe/Paris", "Europe/Paris", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.GERMANY_BBL, "Germany BBL", "germany_bbl", "Europe/Berlin", "Europe/Berlin", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    # legabasket.it's own JSON API is a referenced route contract (probe run
    # externally on 2026-11-22; no first-party body retained here): a
    # championships list keyed by the season's FIRST year, a focus-day calendar
    # with explicit-offset timestamps and finals, and per-match play-by-play
    # whose cumulative period ends derive reconciled quarters. The calendar
    # serves only the league's current focus day — probed day selectors are
    # ignored server-side — so season-wide enumeration (and therefore history
    # backfill) is refused rather than guessed.
    LeagueSpec(
        League.ITALY_LBA,
        "Italy LBA",
        "italy_lba",
        "Europe/Rome",
        "Europe/Rome",
        "lba",
        "lba",
        "lba",
        source_detail=(
            "legabasket.it JSON contract: championships?cs_id=1&s={season first year}, "
            "get-championships-calendar-by-id focus-day slate with explicit UTC offsets, "
            "get-championships-matches-play-by-play cumulative periods reconciled to "
            "quarters. Referenced probe, no first-party capture; the calendar cannot "
            "enumerate a season, so history backfill is refused."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
    ),
        # The federation's own mini-app API answers one Turkey local date at a
    # time: GET /Match/get-daily-matches?match_date=YYYY-MM-DD, grouping the
    # games of every competition it runs (bsl, tbl, kbsl, tkbl) in one body.
    # No key, no special header. A real response body is vendored for this lane
    # (tests/fixtures/tbf/captures/, attributed in CAPTURE-PROVENANCE.md), so the
    # parser is checked against a body the host actually served.
    # The TBF route is live but cannot serve this league. An external probe on
    # 2026-11-23 (recorded in lane-probe-2026-11-23.md) called
    # /Match/get-daily-matches for 2026-11-21, 11-22 and 11-23 and received
    # byte-identical bodies for all three: the match_date parameter is ignored,
    # and the single row in each body belongs to TKBL, not BSL. Zero BSL rows on
    # any date, so no parser change can produce BSL fixtures from it. Declared
    # unavailable rather than left claiming a source: 732 runs previously
    # produced nothing and looked healthy doing it.
    LeagueSpec(League.TURKEY_BSL, "Turkey BSL", "turkey_bsl", "Europe/Istanbul", "Europe/Istanbul", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="TBF mini-app API /Match/get-daily-matches answers 200 but ignores match_date (identical body for 2026-11-21/22/23) and returns only TKBL rows; no BSL fixture is obtainable from this route. The same host does serve BSL: /Match/get-all-matches-for-filter?ActivityId=22214&Page=1&PageSize=-1 answered 200 with 8 live BSL rows (round 9, played finals) in the same row schema, so this lane is unavailable because the adapter is pointed at the wrong route, not because no source exists. Re-point the adapter before re-enabling. See lane-probe-2026-11-23.md.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.GREECE_BASKET_LEAGUE, "Greece Basket League", "greece_basket_league", "Europe/Athens", "Europe/Athens", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    # The externally referenced aba-liga.com calendar/detail route is the
    # active canonical fixture, result, and history target. Ma Golide retains
    # no first-party capture, so live behavior and source-use terms remain
    # unverified. BetExplorer stays separate: aggregate observations are
    # research-only and cannot price tickets or support ROI, receipts, banks,
    # or execution.
    LeagueSpec(League.ABA_LEAGUE, "ABA League", "aba_league", "Europe/Belgrade", "Europe/Belgrade", "aba_liga", "aba_liga", "aba_liga", (_ABA_BETEXPLORER_RESEARCH_ONLY,), source_detail="Externally referenced aba-liga.com calendar/detail route: season/game scope, official team IDs, Europe/Belgrade tips, and period-reconciled finals are required. Ma Golide retains no first-party capture; live behavior and source automation/reuse terms remain unverified. BetExplorer averages remain research-only."),
    # The league's own JSON board route is a referenced contract (basket.co.il
    # ``ws.asmx/Games``, documented by SportsDataverse hoopR ``ibba_games``).
    # It carries per-quarter scores as pipe-delimited strings, but no UTC
    # offset: the adapter applies Asia/Jerusalem explicitly and records that
    # assumption on every game. No Ma Golide first-party capture is retained.
    LeagueSpec(
        League.ISRAEL_PREMIER_LEAGUE,
        "Israel Premier League",
        "israel_premier_league",
        "Asia/Jerusalem",
        "Asia/Jerusalem",
        "israel_ibl",
        "israel_ibl",
        "israel_ibl",
        source_detail=(
            "Winner League schedule/results use the league JSON board route "
            "(basket.co.il/ws/ws.asmx/Games, board_id=5); route contract referenced from "
            "league-site and SportsDataverse documentation, no Ma Golide first-party capture. "
            "Tip times are local wall clock with no offset, so Asia/Jerusalem is applied as a "
            "recorded assumption rather than a provider fact."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
    ),
    # B.League serves its slate as server-rendered HTML on www.bleague.jp and the
    # structured record on a keyless S3 snapshot (web_json/v2_genius_contexts,
    # not the 403ing web_json/contexts tree). Both were captured by an
    # out-of-sandbox probe on 2026-11-23 and vendored under tests/fixtures/
    # bleague/captures; the adapter records every URL/size/digest it reads.
    LeagueSpec(League.JAPAN_B_LEAGUE, "Japan B.League", "japan_b_league", "Asia/Tokyo", "Asia/Tokyo", "bleague", "bleague", "bleague", source_detail="B1 slate is server-rendered HTML (/schedule/?tab=1&year=&mon=&day=, note: mon not month; day is echoed back but does not scope the content -- four day=22/23/24/25 fetches on 2026-11-23 held the same ten game keys and differed by the two echoed day bytes, so the page serves its current round and rows are dated by their own GameDateTime); club ids, UTC tip and quarter scores come from the keyless S3 snapshot web_json/v2_genius_contexts/{key}/{snapshot}.json via a latestid pointer. No key.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.SOUTH_KOREA_KBL, "South Korea KBL", "south_korea_kbl", "Asia/Seoul", "Asia/Seoul", "kbl", "kbl", "kbl", source_detail="Schedule/results use the KBL data API behind www.kbl.or.kr (GET /match/list with a fromDate/toDate window, plus /common/teamList for English team names; two headers, Channel and TeamCode, no key). Route referenced from the SportsDataverse kbl module (MIT, read 2026-11-23); no Ma Golide first-party body is retained, so each run records the URL, byte count and sha256 it actually saw. Quarter splits are not captured, so a settled row is totals-only and not quarter-reconcilable.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.BRAZIL_NBB, "Brazil NBB", "brazil_nbb", "America/Sao_Paulo", "America/Sao_Paulo", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.ARGENTINA_LIGA_NACIONAL, "Argentina Liga Nacional", "argentina_liga_nacional", "America/Argentina/Buenos_Aires", "America/Argentina/Buenos_Aires", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
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
    if normalized == _BETEXPLORER_DIRECT_FALLBACK.provider:
        return _BETEXPLORER_DIRECT_FALLBACK
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
