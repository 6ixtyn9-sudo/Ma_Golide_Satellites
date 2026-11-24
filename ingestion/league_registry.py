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

from dataclasses import dataclass, replace
from enum import Enum
import re
from typing import Dict, Iterable, Optional, Tuple

from .models import (
    League,
    OperationalValidation,
    PriceProvenance,
    SourcePermission,
)


class LaneStatus(str, Enum):
    """Whether a registered competition is part of the operational run.

    ``RETIRED`` does not delete the contract, the state directory, the stored
    finals, or the adapter: a retired lane keeps its identity and history so
    that its past games can still be settled and so that revival is a single
    deliberate flip.  What it loses is scheduling — the gate will not place it
    in the workflow matrix, and the daily CLI refuses to run it by accident.

    ``OFFSEASON`` is a calendar fact, not a verdict: the lane is real, it keeps
    its state, and it must not be reported as a failure while it is dark.
    """

    ACTIVE = "active"
    OFFSEASON = "offseason"
    RETIRED = "retired"


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
class OperatorAcceptance:
    """A named operator decision that a source may be read without a licence."""

    #: Who decided. A decision with no name is not a decision.
    decided_by: str
    #: The UTC date the decision was given, ISO form.
    decided_on: str
    #: What was accepted: the hosts and reads it covers, and their limits.
    scope: str
    #: Why a decision was needed at all.
    basis: str

    def __post_init__(self) -> None:
        for field_name in ("decided_by", "decided_on", "scope", "basis"):
            if not str(getattr(self, field_name) or "").strip():
                raise ValueError(
                    f"an operator acceptance needs a non-empty {field_name}"
                )

    def as_dict(self) -> Dict[str, str]:
        return {
            "decided_by": self.decided_by,
            "decided_on": self.decided_on,
            "scope": self.scope,
            "basis": self.basis,
        }


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
    status: LaneStatus = LaneStatus.ACTIVE
    #: Why the lane is dark.  Required for every non-ACTIVE spec: an
    #: unexplained dark lane is a silent no-op, which is the failure mode the
    #: status field exists to prevent.
    status_detail: str = ""
    #: The licence gate, when it was opened by a person rather than by a host.
    #: Absence means no decision has been recorded; it does not mean "refused".
    #: A record does **not** switch a lane on by itself: an adapter still has to
    #: exist and be replayed against a served page, which is what keeps this
    #: field from becoming a licence nobody can audit.
    operator_acceptance: Optional[OperatorAcceptance] = None

    @property
    def is_schedulable(self) -> bool:
        """True only for a lane the gate may place in the run matrix."""

        return self.status is LaneStatus.ACTIVE

    @property
    def lane_reason(self) -> str:
        """Machine-readable dark-lane reason, empty for an active lane."""

        if self.status is LaneStatus.ACTIVE:
            return ""
        return f"lane_{self.status.value}: {self.status_detail}"

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


#: Providers whose price is a matched counterparty offer on an exchange, not a
#: bookmaker's quote. Named explicitly because the unknown-provider fallback
#: stamps ``NAMED_BOOK``, which would let an exchange quote reach the
#: paper-ticket lane as if a book had priced it.
_EXCHANGE_ONLY_PROVIDERS = frozenset({"betfair", "betfair_exchange"})


def _exchange_observation(provider: str, *, detail: str = "") -> PriceSourcePolicy:
    """An exchange quote: real, gradeable, never a named-book ticket.

    ``acquire`` stays False until an authorised capture proves the route and its
    terms for this project. The vocabulary exists now so the first such capture
    cannot be forced into ``NAMED_BOOK`` or ``AGGREGATE_CONSENSUS``, neither of
    which describes it.
    """

    return PriceSourcePolicy(
        provider,
        PriceProvenance.EXCHANGE,
        SourcePermission.UNRESOLVED,
        OperationalValidation.UNVERIFIED,
        acquire=False,
        detail=detail,
    )


_EXCHANGE_OBSERVATION = _exchange_observation(
    "betfair",
    detail=(
        "Betfair API-NG exchange price (matched counterparty offer), not a "
        "bookmaker quote. No authorised capture for this project yet; acquire "
        "is disabled and the quote can be graded, never paper-ticketed."
    ),
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
    # Revived 2026-11-24: the league has no official schedule route of its own,
    # but its fixtures table carried 7 both-sided quotes on an in-season page --
    # the only zero-priced-adjacent lane in the probe that was actually pricing.
    # Fixtures and the page average now come from one fetch; results have no
    # route at all, which the lane states instead of pretending to settle.
    LeagueSpec(
        League.GREECE_BASKET_LEAGUE,
        "Greece Basket League",
        "greece_basket_league",
        "Europe/Athens",
        "Europe/Paris",
        "betexplorer_fixtures",
        "source_unavailable",
        "betexplorer_fixtures",
        source_detail=(
            "Fixtures table '/basketball/greece/basket-league/fixtures/' (h1 'Basket "
            "League 2026/2027 Fixtures', served 2026-11-24: 126 rows, 7 priced, 2 of the "
            "8 odds-bearing rows carried no kickoff cell). Deepest calendar window in the "
            "sweep, into 2027-01-03. Page average, not a named book; source terms "
            "unreviewed. No results route, so settlement is unsupported by name."
            " Results route exercised end to end on 2026-11-24 from an egress host: EsakeResults?series=8 answered 200 in 139,358 B (sha256 a721bd89...) and robots.txt is absent (404, an allow state). The reader's row container matches the served markup -- 14 esake-program-game containers for 7 rows -- but it finds no EsakegameView?idgame= identity and no final-score block in any of them -- although the served page does carry them: hrefs of the form EsakegameView?idgame=3C7521D5&mode=2 (8 hex, uppercase; two per game, mode=2 and mode=3) appear in it, and a .esake-program-game-final-score style rule exists. So the anchors precede the row class this reader splits on, which is a split-point bug rather than a missing identity. Captured row markup 2026-11-24 confirms it exactly: the toggle handler in a script attribute writes .closest('.esake-program-game'), which the reader's bare-token split counted as a row start, cutting each game at its own toggle (7 games x 2 = the 14 containers, the first half holding both anchors and no score, the second holding the score and no anchor). The reader now requires the class attribute itself, not the bare token, and reads clubs and total from the final-score block, cutting the hidden per-quarter row before counting names; the captured fragment replays to 1 container, 1 row, identity 3C7521D5. The full-page replay against the saved bytes then read the round: containers 7, rows 7, no_identity 0, names_ambiguous 0, and the totals 84-87, 100-64, 83-87, 99-77 -- but with_final 0, because all seven dates were unreadable. The served date cell breaks inside the month name (newlines and indentation between the letters of the month), so a reader demanding a space between day and month found nothing on a page that plainly shows a date; it now retries with all whitespace removed, and a nested span inside a club cell is read whole rather than trimmed at the inner close tag. What still gates the lane is a dated final from the served bytes, not a route, so the route is served while the parser's identity and score patterns are wrong. That is a parser fix against captured bytes, not a route question: no results are written and the lane stays source_unavailable until the reader is replayed against the saved page and the date+tip join is unambiguous. See docs/RESULTS-ROUTE-ASSESSMENT.md."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
        operator_acceptance=OperatorAcceptance(
            decided_by="operator",
            decided_on="2026-11-24",
            basis=(
                "the esake.gr terms license copying with source credit but scope the "
                "licence to personal, non-commercial use. Operator reading of that "
                "clause, 2026-11-24: accepted for this paper record."
            ),
            scope=(
                "Read esake.gr EsakeResults and EsakegameView pages, at most one "
                "matchday page per series per run, with source and author credit on "
                "every row. No redistribution, no commercial use."
            ),
        ),
    ),
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
    LeagueSpec(
        League.JAPAN_B_LEAGUE,
        "Japan B.League Premier",
        "japan_b_league",
        "Asia/Tokyo",
        "Asia/Tokyo",
        "bleague",
        "bleague",
        "bleague",
        source_detail=(
            "Top division (B.LEAGUE PREMIER, the schedule page's tab=1; the tab "
            "numbering survived the 2026-27 rebrand, the division names did not). "
            "Slate is server-rendered HTML (/schedule/?tab=1&year=&mon=&day=, note: "
            "mon not month; day is echoed back but does not scope the content -- "
            "four day=22/23/24/25 fetches on 2026-11-23 held the same ten game keys "
            "and differed by the two echoed day bytes, so the page serves its "
            "current round and rows are dated by their own GameDateTime); club ids, "
            "UTC tip and quarter scores come from the keyless S3 snapshot "
            "web_json/v2_genius_contexts/{key}/{snapshot}.json via a latestid "
            "pointer. No key. This lane claims no aggregate price source: the only "
            "BetExplorer Japan board page captured (japan/b-league-one) is the "
            "SECOND division and matched the tab=2 slate game-for-game, so pricing "
            "PREMIER fixtures against it guaranteed that no club name could agree "
            "(run 45385819819 priced nothing and claimed full coverage). Its own board page, "
            "basketball/japan/b-league-premier, was read on 2026-11-24 by an external "
            "browser agent, which quoted the served heading 'B.League Premier 2026/2027 "
            "Fixtures' and dated rows from 04.12.; that read is third-party, so the "
            "pipeline's own first fetch confirms the slug rather than assuming it. "
            "Split from JAPAN_B_ONE on 2026-11-24."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
    ),
    LeagueSpec(
        League.JAPAN_B_ONE,
        "Japan B.League One",
        "japan_b_one",
        "Asia/Tokyo",
        "Asia/Tokyo",
        "bleague",
        "bleague",
        "bleague",
        source_detail=(
            "Second division (B.LEAGUE ONE, the schedule page's tab=2), split out of "
            "the Japan lane on 2026-11-24 because one lane cannot be two divisions. "
            "Same contract as the Premier lane -- server-rendered slate at "
            "/schedule/?tab=2&year=&mon=&day= (mon, not month) plus the keyless S3 "
            "snapshot web_json/v2_genius_contexts/{key}/{snapshot}.json via a "
            "latestid pointer, both captured out-of-sandbox and vendored under "
            "tests/fixtures/bleague/captures -- read with tab=2. Its aggregate price "
            "source is the one Japan board page that was captured and matched the "
            "tab=2 slate game-for-game (basketball/japan/b-league-one); the site "
            "average only, never a named book."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
    ),
    LeagueSpec(League.SOUTH_KOREA_KBL, "South Korea KBL", "south_korea_kbl", "Asia/Seoul", "Asia/Seoul", "kbl", "kbl", "kbl", source_detail="Schedule/results use the KBL data API behind www.kbl.or.kr (GET /match/list with a fromDate/toDate window, plus /common/teamList for English team names; two headers, Channel and TeamCode, no key). Route referenced from the SportsDataverse kbl module (MIT, read 2026-11-23); no Ma Golide first-party body is retained, so each run records the URL, byte count and sha256 it actually saw. Quarter splits are not captured, so a settled row is totals-only and not quarter-reconcilable.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.BRAZIL_NBB, "Brazil NBB", "brazil_nbb", "America/Sao_Paulo", "America/Sao_Paulo", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    LeagueSpec(League.ARGENTINA_LIGA_NACIONAL, "Argentina Liga Nacional", "argentina_liga_nacional", "America/Argentina/Buenos_Aires", "America/Argentina/Buenos_Aires", "source_unavailable", "source_unavailable", "source_unavailable", source_detail="No compliant date-bounded official schedule/result adapter has been validated.", price_sources=(_aggregate_research_only("betexplorer"),)),
    # ------------------------------------------------------------------
    # 2026-11-24 additions.  These four competitions have no official schedule
    # route of their own; their public fixtures table is the only one, so the
    # schedule and the price come from the same single fetch.  Acquisition time
    # is Paris because that is the clock the table is served in for every
    # competition; display stays in the league's own timezone.
    # ------------------------------------------------------------------
    LeagueSpec(
        League.CHINA_WCBA_W,
        "WCBA Women",
        "china_wcba_w",
        "Asia/Shanghai",
        "Europe/Paris",
        "betexplorer_fixtures",
        "source_unavailable",
        "betexplorer_fixtures",
        source_detail=(
            "Fixtures table '/basketball/china/wcba-women/fixtures/' (h1 'WCBA Women "
            "2026/2027 Fixtures', served 2026-11-24: 255 rows, 3 priced, 2 rows with no "
            "kickoff cell). Team names carry a ' W' suffix, so identity is the provider "
            "match code and names are display only. The served clock is CET/CEST, not "
            "Asia/Shanghai. Page average, not a named book; source terms unreviewed. "
            "No results route, so settlement is unsupported by name."
            " Results route observed twice, and the two observations disagree. A 2026-11-24 probe read cbamatch/games/matchschedules?competitionId=56&seasonId=2286 and got the whole 2026-27 season (297,504 B, integer gameId, scores, gameStatus, 10/10 coverage). Re-read from an egress host at 2026-11-24T14:02Z the same URL answers HTTP 500 in 179 B: {\"error\":\"Internal Server Error\",\"message\":\"param name: roundId value not exists\"}. So the documented URL is not sufficient as written and a roundId parameter is required -- scripts/probe_cba_results.py enumerates bounded roundIds and records what each returns, because the round space is the one thing the error does not tell us. The second error named the second parameter: with roundId=1&teamId=N the host stops rejecting and answers 200 for every N in 1-30, each time in 42 B with zero games, so the value space is a (round, team) grid and both coordinates are now accepted -- round 1 simply holds no game for those team ids. The grid is bounded (REQUEST_BUDGET 40, refusing rather than trimming a wider sweep) and the next read is the 42-byte body itself, to learn whether it is an empty list or a refusal wearing a 200. It is an empty list: the body is a data array that is empty, a status of 0 and an explicit success message -- not an error wearing a 200 -- so round 1 genuinely holds no game for team 1 and the next move is finding the round that does, not re-reading the refusal. That move was made: teamId=1 walked across roundId 1-30, thirty requests, every one 200 in 42 B with zero games. So no single team id in the range the error implied is a team in this season, and the grid is exhausted rather than merely unexplored: what is missing is a real team id or a real season id, and neither is guessable from the route's own errors. The licence gate is open (operator_acceptance) and the host still publishes no terms document. See docs/RESULTS-ROUTE-ASSESSMENT.md."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
        operator_acceptance=OperatorAcceptance(
            decided_by="operator",
            decided_on="2026-11-24",
            basis=(
                "cba.net.cn publishes no terms document of any kind, so there is no "
                "licence to read, comply with, or breach. The only thing that can "
                "authorise the read is a named operator decision: operator instruction "
                "2026-11-24, activate WCBA now."
            ),
            scope=(
                "Read the documented cba.net.cn data-hub match-schedule endpoint for "
                "this paper-only record, one season-scoped read per run plus the per- "
                "game reads its ids imply. No redistribution, no named-book or ticket "
                "use, attribution on every row."
            ),
        ),
    ),
    LeagueSpec(
        League.LITHUANIA_NKL,
        "Lithuania NKL",
        "lithuania_nkl",
        "Europe/Vilnius",
        "Europe/Paris",
        "betexplorer_fixtures",
        "source_unavailable",
        "betexplorer_fixtures",
        source_detail=(
            "Fixtures table '/basketball/lithuania/nkl/fixtures/' (h1 'NKL 2026/2027 "
            "Fixtures', served 2026-11-24: 167 rows, 7 priced of which 6 still future, "
            "99 rows with no kickoff cell). Clock is the site's CET/CEST wall clock. "
            "Page average, not a named book; source terms unreviewed. No results route."
            " WP-JSON surface exercised 2026-11-24T14:03Z from an egress host: the namespace index answers 200 in 398,405 B and lists nkl-calendar/v1; nkl-calendar/v1/teams?season=2026 answers 200 in 1,390 B as {\"season\":2026,\"teams\":[{\"id\":3386,\"name\":\"Alytaus Patriotai\",\"slug\":\"alytaus-patriotai\"}]}. But the single-game object answers schedule only: keys id, summary, location, start_local, start_utc, end_utc, url, ics -- no score, no status. So the earlier \"endpoint family proven\" reading overstated it: this namespace is a calendar, not a results source, and the scores have to come from another route (the index also lists nkl/v1 and nkl-advanced-statistics/v1, neither yet read). Re-read 2026-11-24T14:2xZ: urbo-nkl/v1/games answers 401 rest_unauthorized ('Trūksta Authorization Bearer antraštės') and nkl/v1/live-scores answers 200 with an empty array, so the score-bearing namespaces observed so far are token-gated or empty rather than merely unread; nkl-amw/v1/games answers 401 rest_forbidden ('Deja, Jūs negalite to daryti') and nkl/v1/live-scores?date=2026-11-24 answers 200 with the same empty array, so every score-bearing namespace observed on this host is auth-gated or empty and the WP-JSON surface has no public results route left to try; the next read is the site's own match pages. The aggregator's own results page for this lane was then captured at 14:45Z -- 200, 192,203 B, sha256 07a382c1774d...16de, 279 match-code links resolving to 140 distinct matches and 141 colon scores -- so the finals this host will not serve over WP-JSON are readable from a path-only, robots-allowed page. robots.txt carries Crawl-delay 10 with no Disallow; the host publishes no terms document. Licence gate open (operator_acceptance). See docs/RESULTS-ROUTE-ASSESSMENT.md."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
        operator_acceptance=OperatorAcceptance(
            decided_by="operator",
            decided_on="2026-11-24",
            basis=(
                "nkl.lt publishes no terms document -- its sitemap lists only cookie "
                "and privacy policies -- so there is no licence to read, comply with, "
                "or breach. Operator instruction 2026-11-24, activate NKL now."
            ),
            scope=(
                "Read the documented nkl.lt WP-JSON namespace (teams by season, game by "
                "id) for this paper-only record, honouring its Crawl-delay: 10 and one "
                "page per game. No redistribution, attribution on every row."
            ),
        ),
    ),
    LeagueSpec(
        League.AUSTRALIA_WNBL_W,
        "WNBL Women",
        "australia_wnbl_w",
        "Australia/Sydney",
        "Europe/Paris",
        "betexplorer_fixtures",
        "source_unavailable",
        "betexplorer_fixtures",
        source_detail=(
            "Fixtures table '/basketball/australia/wnbl-women/fixtures/' (h1 'WNBL Women "
            "2026/2027 Fixtures', served 2026-11-24: 67 rows, 3 priced all still future, "
            "0 rows with no kickoff cell -- the cleanest page in the sweep). Clock is "
            "the site's CET/CEST wall clock. Page average, not a named book; source "
            "terms unreviewed. No results route."
            " Results route re-read 2026-11-24T14:03Z from an egress host: get/wnbl/matches/in/season/2026/regular answers HTTP 401 in 88 B even with Referer: https://www.wnbl.com.au/, where the earlier probe recorded 200 in 493,824 B. The referer alone no longer opens it, so the route now needs either a documented token or the exact headers the site's own client sends, and neither has been observed. wnbl.com.au publishes terms; licence gate open (operator_acceptance). The official route therefore needs a session token this project will not replicate. The next read is the aggregator's own results page for this lane (basketball/australia/wnbl-women/results/, path-only and robots-allowed), captured by scripts/probe_betexplorer_results.py --index. Captured 2026-11-24T14:41Z: 200, 129,186 B, sha256 f4249ca717af428117c9fb4c745a20b3e9f65b1a7d1049c79866b1a04be6115c, has_table true, 67 match-code links, 34 distinct codes, 35 colon scores, plus a 'No matches found' string that belongs to a hidden empty-state template rather than to this table. So the page carries a season of finals and the reader can be written against captured markup. Re-read 14:45Z: 128,973 B, sha256 e47afc235509...15b252, the same 67/34/35 counts. The remaining artifact is the shape of a row, which --index saves as a 2 KB window around the first code (needs the commit that added it) See docs/RESULTS-ROUTE-ASSESSMENT.md."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
        operator_acceptance=OperatorAcceptance(
            decided_by="operator",
            decided_on="2026-11-24",
            basis=(
                "wnbl.com.au publishes terms and prod.rosetta.nbl.com.au answered per- "
                "season match JSON, but this lane's route assessment was never "
                "completed, so nothing was ever enabled by evidence. Operator "
                "instruction 2026-11-24, activate WNBL, permission granted."
            ),
            scope=(
                "Read the documented rosetta.nbl.com.au per-season WNBL match route and "
                "the wnbl.com.au terms it is governed by, for this paper-only record; "
                "one season-scoped read per run, referer-gated as served. No "
                "redistribution, attribution on every row."
            ),
        ),
    ),
    # Retired 2026-11-24 on measurement, like the nine before it: the schedule
    # side works but no permitted results route exists, so the lane can never
    # settle and every run of it is a fixture card that grades nothing. The
    # contract, adapter and state directory stay, so revival is one deliberate
    # flip if a results route ever appears.
    LeagueSpec(
        League.ISRAEL_LIGA_LEUMIT,
        "Israel Liga Leumit",
        "israel_liga_leumit",
        "Asia/Jerusalem",
        "Europe/Paris",
        "betexplorer_fixtures",
        "source_unavailable",
        "betexplorer_fixtures",
        source_detail=(
            "Fixtures table '/basketball/israel/liga-leumit/fixtures/' (h1 'Liga Leumit "
            "2026/2027 Fixtures', served 2026-11-24: 161 rows, 5 priced of which 3 still "
            "future, 83 rows with no kickoff cell). This is the second division, "
            "distinct from the Winner League board lane. Clock is CET/CEST. Page "
            "average, not a named book; source terms unreviewed. No results route: "
            "one.co.il puts the only score-bearing path (/api/) behind a robots "
            "Disallow, safsal.co.il publishes no terms page, the federation host "
            "basketball.org.il timed out on three attempts across two days, and "
            "basket.co.il serves the Premier board only (Leumit board ids answer "
            "empty). See docs/RESULTS-ROUTE-ASSESSMENT.md."
        ),
        price_sources=(_aggregate_research_only("betexplorer"),),
        status=LaneStatus.RETIRED,
        status_detail=(
            "retired 2026-11-24 on the route assessment: the schedule side is served, "
            "but every candidate results route is closed (robots on the only "
            "score-bearing path, no terms page, a federation host that timed out across "
            "two days, and a Premier-only board), so the lane can never settle and its "
            "fixture cards grade nothing. Keeping it in the matrix spent a lane of "
            "wall-clock to produce zero graded rows."
        ),
    ),
)

# --------------------------------------------------------------------------
# Lane status — measured 2026-11-24, single-day probe, ledgered bodies.
#
# The rule this table encodes: retire a lane when its *source* refuses us or
# serves nothing usable, not when its season is simply quiet and not when it
# merely lacks a price route.  Measured zero priced rows on an in-season
# fixtures page, or no compliant schedule route at all, is a fight we stop
# having.  A lane with a working schedule/result adapter keeps its place even
# when no price exists, because prices can arrive later without re-deriving
# the fixtures.
#
# Retirement is reversible and deletes nothing: the contract, state directory,
# stored finals, and adapter all stay.  See ``LaneStatus``.
# --------------------------------------------------------------------------

_RETIRED_ON_2026_11_24: Dict[League, str] = {
    # Every entry below is a lane that ran for ~750 scheduled runs and produced
    # essentially nothing it could ever price.  The run and fixture counts come
    # from each lane's own committed archive manifest; the row counts come from
    # the 2026-11-24 fixtures-page probe.
    League.SPAIN_ACB: (
        "756 runs, 0 games ever acquired and 0 of 225 fixtures rows priced"
    ),
    League.BASKETBALL_CHAMPIONS_LEAGUE: (
        "760 runs, 0 games ever acquired and 0 of 32 fixtures rows priced; season "
        "archives serve whole-season totals only, never a pre-tip slate"
    ),
    League.BASKETBALL_AFRICA_LEAGUE: (
        "760 runs, 0 games ever acquired; odds appear only on already-played rows "
        "and the fixtures table this adapter reads carries none"
    ),
    League.ARGENTINA_LIGA_NACIONAL: (
        "760 runs, 0 games ever acquired and 0 of 218 fixtures rows priced"
    ),
    League.ITALY_LBA: (
        "760 runs for 6 games; the calendar body is byte-identical with and "
        "without the day parameter, so only the focus day can ever be read"
    ),
    League.FRANCE_LNB_PRO_A: (
        "760 runs, 0 games: no compliant date-bounded schedule/result route was validated"
    ),
    League.GERMANY_BBL: (
        "760 runs, 0 games: no compliant date-bounded schedule/result route was validated"
    ),
    League.TURKEY_BSL: (
        "752 runs, 0 games: no compliant date-bounded schedule/result route was validated"
    ),
    League.BRAZIL_NBB: (
        "760 runs, 0 games: no compliant date-bounded schedule/result route was validated"
    ),
}

#: Considered for retirement on 2026-11-24 and **rejected** — recorded so the
#: decision is not silently re-litigated.  ABA League looked like a fighter from
#: the fixtures page (0 of 100 rows priced), but that page only ever serves a
#: site *average*, which is research-only here.  Its ticket route is the direct
#: named-book comparison, and that route demonstrably worked: on 2026-11-22 it
#: produced three priced picks (Stake.com 1.06 / 1.78 / 1.25, named_book,
#: pre-match) out of 30 logical picks.  Those three were CAUTION rather than
#: clean only because calibration held n=13 against a minimum of 30 walk-forward
#: samples.  That is a lane waiting to mature, not a source refusing us.
_KEPT_AFTER_REVIEW: Dict[League, str] = {
    League.ABA_LEAGUE: (
        "4252 games acquired and reconciled finals back to 2004; 3 of 30 frozen "
        "picks carried a named-book price on 2026-11-22 (betexplorer_direct, "
        "Stake.com), CAUTION only for insufficient walk-forward samples"
    ),
}

#: Empty fixtures page is a calendar fact here, not a measured fight.  The
#: lane keeps its state and its place; it must not alarm the heartbeat while
#: dark, and it must not be confused with a retirement.
_OFFSEASON_ON_2026_11_24: Dict[League, str] = {
    League.WNBA: (
        "WNBA season runs April-October; the 2026-11-24 fixture page was "
        "empty because the calendar is, not because the source refused us"
    ),
}


def _with_lane_status(spec: LeagueSpec) -> LeagueSpec:
    if spec.league in _RETIRED_ON_2026_11_24:
        return replace(
            spec,
            status=LaneStatus.RETIRED,
            status_detail=_RETIRED_ON_2026_11_24[spec.league],
        )
    if spec.league in _OFFSEASON_ON_2026_11_24:
        return replace(
            spec,
            status=LaneStatus.OFFSEASON,
            status_detail=_OFFSEASON_ON_2026_11_24[spec.league],
        )
    return spec


SPECS: Tuple[LeagueSpec, ...] = tuple(_with_lane_status(spec) for spec in _SPECS)
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
    if normalized in _EXCHANGE_ONLY_PROVIDERS:
        # An exchange identity is never a named book, even when a league has
        # not registered it: the fallback below would otherwise stamp it
        # NAMED_BOOK and make it ticket-eligible.
        return _EXCHANGE_OBSERVATION
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


def schedulable_specs() -> Tuple[LeagueSpec, ...]:
    """Return the lanes the operational run may schedule."""

    return tuple(spec for spec in SPECS if spec.is_schedulable)


def retirement_guard(league: League, *, allow_retired: bool = False) -> Optional[str]:
    """Return a refusal message for a dark lane, or ``None`` to proceed.

    A retired lane is refused unless the operator explicitly asks for it: a
    stale cron entry or a copied CI command must never resurrect a lane that
    was measured to fight us.  An off-season lane remains runnable on explicit
    invocation, because it is real and keeping its state warm is legitimate;
    the caller only needs to know it will be quiet.
    """

    spec = league_spec(league)
    if spec.status is LaneStatus.RETIRED and not allow_retired:
        return (
            f"{spec.league.value} is retired: {spec.status_detail}. "
            "Re-run with --run-retired only to inspect or revive it; "
            "revival is a one-line status flip in ingestion/league_registry.py."
        )
    return None


__all__ = [
    "LaneStatus",
    "_KEPT_AFTER_REVIEW",
    "LeagueSpec",
    "PriceSourcePolicy",
    "SPECS",
    "cli_league_choices",
    "league_spec",
    "parse_league",
    "price_evidence_policy",
    "retirement_guard",
    "schedulable_specs",
    "source_unavailable_specs",
]
