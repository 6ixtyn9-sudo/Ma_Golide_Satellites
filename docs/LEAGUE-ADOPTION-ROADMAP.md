# League readiness and source-adoption roadmap

**Status:** planning and evidence-governance contract. It does not enable a
provider, make a network request, grant a credential, change a ticket rule, or
rewrite historic operator evidence.

**Reviewed:** 2026-11-23 — **operator policy change.** The project's ambition is
no longer NBA-centric, and scraping publicly served league and media pages is
now permitted for the non-NBA expansion lanes. Sections marked *(2026-11-23)*
state the new position; everything unmarked is unchanged.

## Purpose

Ma Golide must be able to grow beyond today's direct named-book lanes without
confusing an unavailable current source with a permanent league exclusion. This
document describes the route from a supported basketball league to a possible
future paper-ticket path while preserving the repository's evidence, identity,
and no-retroactivity rules.

The The Odds API reservation guard is **provider-scoped quota control**. Its
current routine priority (NCAA women, NBA, NCAA men, WNBA) limits automated use
of that shared key; it is not a project-wide admission policy and does not rule
out a future compliant source or validated path for another league.

This roadmap is derived from current repository contracts. It is intentionally
separate from the source-discovery inventory in
[FREE-API-RESEARCH.md](FREE-API-RESEARCH.md): discovering a source is not
adopting it.

## Scope (2026-11-23): beyond the NBA

The first version of this roadmap treated the NBA as the project. That was a
limit on ambition, not a finding about evidence: the repository already carries
24 typed league lanes, and 18 of them produce no picks because of upstream data
gaps, not model or gate problems. ESPN's own registry was checked on 2026-11-23
and carries 15 basketball leagues — all United States competitions plus NBL,
FIBA and the Olympics — so no American route will ever supply Liga ACB, Lega A,
BSL, KBL or B.League.

The operator's position from 2026-11-23:

- **Ambition.** Every configured league is a candidate for schedule, result,
  forecast, price and paper-ticket coverage. The NBA remains the reference lane
  — the most complete acquisition, the deepest history, the most efficient
  market — but it is no longer the boundary of the project.
- **Access.** Scraping publicly served league and media pages, and consuming
  undocumented public JSON endpoints, is **permitted for non-NBA lanes**,
  subject to the retained invariants in *Explicit non-goals*. This supersedes
  the blanket prohibition stated in earlier revisions of this document.
- **Rationale, stated as a hypothesis rather than a finding.** The NBA is the
  most efficient basketball market in the world: the sharpest books, the most
  arbitrage, the tightest closing lines. Any surviving model edge is therefore
  more likely to sit in thinner domestic markets — Liga ACB, Lega A, BSL, KBL,
  B.League, NBL — where prices are less competed away.

### Testing the market-efficiency hypothesis

The hypothesis is admissible only if it produces evidence:

| question | evidence required |
| --- | --- |
| Is the league priceable at all? | an enumerated key list from every available provider, per lane |
| Are the prices soft? | the per-lane distribution of calibrated EV, and realised ROI by EV band, against the NBA lane as reference |
| Is the softness obtainable? | bookmaker, market, line, period, pre-tip observation time and pre-match state on every emitted price |
| Is it stable? | closing-line movement and quote staleness per lane; a stale quote is not an edge |

**Verified 2026-11-23:** The Odds API's published sports table lists exactly 11
basketball keys, and of the lanes this project runs only `basketball_euroleague`
and `basketball_nbl` (Australia) are among them. Liga ACB, BSL, KBL, B.League,
LNB Pro A, BBL, ESAKE, NBB and Liga Argentina are **absent from that provider's
coverage** — not merely unpriced by us. No amount of adapter work will make those
lanes billable through The Odds API; they need a different book source or stay
research-only. This is the "is the league priceable at all?" row below, filled in
for one provider.

Two cautions must survive any future rewrite. **A soft market with no obtainable
price is not an opportunity** — coverage, not softness, is what bounds 15 of the
24 lanes today. And the settled evidence does not yet support assuming that
inefficiency is easy to harvest: across the 40 settled NCAA-M legs, EV>=1
returned 15W/11L against 12W/2L for EV<1, and the 14 settled positions split
7/7 across that threshold with both halves in profit. Forty legs is not a
result; it is a reason to keep measuring rather than to assume.

## Confirmed current baseline

| Area | Current repository position | Consequence |
| --- | --- | --- |
| League registry | `ingestion/league_registry.py` has a typed contract, independent state key, timezone, schedule/result capability, and price-source policy for every supported league. | An unavailable source produces an explicit coverage gap; it does not invent a fixture, result, probability, or price. |
| Direct named-book routine | The Odds API is configured for NBA, WNBA, NCAA men, and NCAA women. Its routine request is full-game `h2h,totals,spreads` through the named DraftKings group. | These are the only current automatic routine permits for that provider. This says nothing about other providers or future league paths. |
| NCAA women mapping | The documented The Odds API sport key used by the repository is `basketball_wncaab`. | A live coverage result is still required to prove account, region, bookmaker, identity, and timestamp availability for a particular fixture. |
| Provider cost control | The Odds API ledger reserves the documented maximum for the locked request shape before a runner receives the key, then reconciles numeric quota headers after completed work. | A post-request header can tighten future reservations; a mismatch-quarantine policy is a future enhancement. It can never authorise a request that was not pre-reserved. |
| Aggregate/exchange evidence | BetExplorer and any aggregate consensus remain research-only. Prediction/exchange-market observations are not sportsbook prices. | They cannot create a named-book paper ticket, receipt, stake, bank, ROI, or execution quote. |
| Historic evidence | Frozen receipts, ticket stakes, bank, ROI, and settled outcomes are immutable. | Future readiness work is additive and prospective. It must never reclassify or rewrite historic operator evidence. |

The existing registry labels a direct source's reuse/operational policy and
stamps that policy onto observations. A declared provider adapter is not proof
that a specific account, bookmaker, fixture, or timestamp is presently covered.
Live coverage is a separate per-fixture fact.

## Readiness is a projection, not a second registry

The first implementation increment is a **derived readiness report**. It must
read existing structured contracts and retained metadata; it must not introduce
editable status prose that can drift away from the pipeline.

A report row must contain independent axes rather than one `ready` value:

| Axis | Question answered | Example values |
| --- | --- | --- |
| `identity_results` | Can the project identify canonical fixtures and reconcile results for this league? | `unavailable`, `partial`, `validated` |
| `history` | Is retained reconciled history sufficient for the relevant forecast policy? | `none`, `developing`, `sufficient` |
| `forecast` | Can the league emit a valid pre-tip forecast/pick surface? | `unavailable`, `candidate`, `normal_lane_available` |
| `price_path` | Is there a compliant direct named-book path for future validation? | `none_known`, `documented_unverified`, `bounded_validation_available`, `fixture_evidence_observed` |
| `source_contract` | What roles, rights, review status, and operational evidence exist for the source? | see below |

The report must also derive one deterministic **next blocking action**. It must
not show a league-level `ticket_ready` badge. A league can have a future ticket
path while a specific upcoming game has no valid current quote.

## Four distinct evidence layers

### 1. League capability

League capability is the tuple of readiness axes above. It explains what the
project can potentially do for a league, not whether it can issue a ticket for
a particular game.

### 2. Source contract

A future source-contract record must keep these dimensions distinct:

```text
candidate_role
contractually_approved_role
operational_status_for_specific_lane
fixture_evidence_status
review_status
```

Permitted roles are deliberately narrow:

```text
schedule_results
prediction
named_book_price
exchange_research
aggregate_research
historical_calibration
```

A candidate role grants no adapter wiring, secret access, workflow environment,
scheduled request, or ticket authority. Existing approved roles are preserved;
new candidate contracts default to no operational access. Approval for one role
never grants another role.

A future review date can block **new expansion or new validation approval**. It
must not erase historic source provenance or alter a frozen receipt.

### 3. Fixture lifecycle

For future prospective reporting, the project distinguishes:

```text
forecast_candidate
normal_lane_pick
ticket_candidate
frozen_paper_ticket
```

The display/reporting distinction is important:

- A `forecast_candidate` is a retained pre-tip forecast under the applicable
  forecast-evidence policy.
- A `normal_lane_pick` is visible ordinary-lane output. If it has no exact
  permitted price, it must be visibly marked as unpriced and not a paper
  ticket.
- A `ticket_candidate` needs exact fixture-specific named-book evidence:
  bookmaker, market, period, selection, line, opposing price where applicable,
  pre-tip observation time, and pre-match state.
- A `frozen_paper_ticket` additionally satisfies the existing prospective
  decision/freeze policy.

This lifecycle is prospective. Existing frozen receipts are not migrated,
relabelled, downgraded, or rewritten merely because new reporting metadata is
introduced.

### 4. Frozen receipt

A frozen paper-ticket receipt remains the immutable record of the forecast,
price evidence, decision time, and stake at the applicable cutoff. A later
provider outage, source review expiry, market removal, or documentation change
can block future tickets only. It cannot revise settled bank, ROI, stake, or
historic receipt content.

## Pick-first with bounded parallel price discovery

The default priority is to establish useful pick capability before spending a
price-provider budget:

1. canonical fixture/result identity;
2. reconciled history where the forecast policy requires it;
3. an independent prediction or valid empirical forecast surface;
4. visible normal-lane output where policy permits it;
5. bounded direct-price validation for an actual upcoming slate.

This does **not** mean price work is ignored until the end. The following are
safe in parallel because they do not spend a price-provider quota:

- record a documented potential sport key or source role;
- record official terms, rate-limit, regional, and retention/reuse evidence;
- map a prospective source to a specific league and role;
- identify the exact identity/timestamp/bookmaker evidence required for a
  future bounded validation.

A paid or quota-consuming price call is justified only when it has a defined
purpose: an otherwise useful league has an upcoming slate and the direct price
path is the remaining material evidence gate. It must not become blanket price
crawling across every basketball surface.

Unpriced forecasts may contribute to **forecast accuracy/calibration** when
both the retained pre-tip forecast and canonical outcome are valid. They are
excluded from ticket counts, stakes, bank, ticket ROI/P&L, closing-line work,
and any price-based performance reporting.

## Existing paths and next work

| League group | Current project position | Next evidence-focused work |
| --- | --- | --- |
| NBA, WNBA, NCAA men, NCAA women | Schedule/result and forecast routes exist; The Odds API is the configured direct named-book routine path. | Retain bounded live coverage evidence when the existing authorised routine runs observe a real fixture. Do not create a separate broad validation lane merely to repeat the same request shape. |
| EuroLeague | Schedule/result and prediction routes exist; current price reporting is aggregate research only. | Record the direct-price path as documented/unverified only. Any future new The Odds API or other-provider validation must be separately proposed, bounded, non-writing, and evidence-specific. |
| NBA G League, CBA, ABA League, EuroLeague Women | Each has a different combination of schedule/result/history/forecast support and no current validated direct named-book lane. | The readiness report identifies the earliest missing prerequisite. Preserve existing ABA identity and canonical-history contracts. |
| Israel Premier League | Schedule/result lane added on the league JSON board: quarter scores, OT totals, finals, English team names. No UTC offset in the source, so `Asia/Jerusalem` is applied and recorded as an assumption; a DST-gap tip is refused. Four captured rows (2026-11-22) confirmed points-in-period quarters and exposed that a row's season is `year_id`, not `cYear` — the earlier reading left every game's season empty, which made result refresh refuse identity. Pricing stays on the direct BetExplorer fallback. | Watch the first scheduled runs; confirm the applied offset against a captured tip time. |
| Spain ACB | Schedule/result lane added on the ACB JSON API: season calendar plus per-game header for identity, status, final score and quarter/OT reconciliation. Referenced route contract, fixture-tested, no first-party capture; no history backfill yet; pricing stays on the direct BetExplorer fallback. | Watch the first scheduled runs for a real fixture and a reconciled final. A bounded per-game history plan is the next prerequisite before the empirical model can accumulate team history. |
| Italy LBA | Schedule/result lane added on the league's own JSON API: championships list keyed by the season's first year, a focus-day calendar with explicit-offset tips, finals and status codes, and per-match play-by-play differenced into quarters. Two bodies are committed byte-exact from the 2026-11-22 capture and re-hashed by a test; the capture also corrected status handling (live rows carry real scores) and showed one focus day is a round spanning two dates. Probed day selectors are ignored, so no season enumeration and no history backfill; pricing stays on the direct BetExplorer fallback. | Watch the first scheduled runs for a real focus-day slate and a reconciled final. A date-bounded route is the prerequisite before this lane can carry history. |
| EuroCup | Schedule/result lane added by reusing the validated EuroLeague Basketball adapter with competition code `U`; fixture-tested, with no first-party EuroCup capture and no validated price contract. | Watch the first scheduled runs for a real fixture and reconciled result. Price work stays on the direct BetExplorer fallback until a named-book contract is separately evidenced. |
| Other expansion leagues *(2026-11-23)* | Typed registry and isolated state exist; schedule/result acquisition is unavailable through provider APIs. League-owned public JSON is the first choice, media enterprises (Sofascore, LiveScore, Naver, Globo, Mackolik) the fallback. | Establish fixture/result identity from the scraped route, commit one scrape receipt per lane, then forecast. A discovery lead is still not an integration. |

Australia NBL remains a documented/probationary The Odds API mapping only. It
has no automatic provider budget promotion, no automatic schedule/result
assumption, and no ticket authority until its separate evidence requirements
are met.

## Future validation proposal: design contract only

This is a future, code-only mechanism. It does not exist yet and must not make
network calls on its own.

A proposal is required only for a **new provider, new league, new source role,
or new request shape**. It does not add a human approval queue to the existing
approved and permit-gated routine The Odds API lanes.

Before any future external call, a proposal must declare:

```text
source/provider identity
league and role under test
official documentation and terms evidence
operator-region eligibility
ongoing free-tier/rate-limit evidence
expected request shape and maximum cost
one-off or recurring intent
no-write guarantee
expected identity, timestamp, bookmaker, and market evidence
promotion criteria
failure, quota, and rollback behaviour
```

A proposal cannot relax a provider's existing hard quota guard. For The Odds
API, the current routine guard remains a pre-request upper-bound reservation
plus post-request conservative reconciliation. A later mismatch policy may
quarantine new permits if `x-requests-last` exceeds the reserved maximum; it
must never use a post-request header to authorise spending retroactively.

## Phase 1 acceptance criteria

The readiness-report increment is ready to ship only when all of the following
are true:

- It is reproducible from structured repository inputs and contains no manually
  maintained readiness prose.
- It has no scalar `ready`/`not_ready` or league-level `ticket_ready` field.
- It shows the readiness axes and a derived next blocking action for every
  configured league.
- It preserves existing configured source roles; new candidate contracts grant
  no operational access.
- It makes no network call, reads no provider secret, changes no workflow, and
  makes no ticket, bank, ROI, stake, receipt, or settlement write.
- It does not weaken ABA identity validation or reinterpret the `2027` ABA
  season-end-year label as a calendar date.
- Existing `CoverageGap` schema and normal pipeline reports are unchanged.
- Existing frozen receipts and settled performance artifacts remain byte-for-byte
  unchanged.

## Explicit non-goals

This roadmap does not, on its own authority, approve a new provider, source,
price feed, exchange, or prediction market, and it does not treat
aggregate/prediction/exchange prices as named-book ticket prices.

*(2026-11-23)* Earlier revisions also prohibited *scraping, undocumented
sportsbook endpoints, browser automation and proxying*. **The operator has
overridden that prohibition for non-NBA expansion lanes.** What replaces it is a
narrower set of retained invariants, so that "permitted" does not become
"unaccountable":

| Retained invariant | Why it survives the override |
| --- | --- |
| NBA, WNBA and NCAA lanes keep their current acquisition routes | those relays are validated; changing them puts the only lanes that settle at risk |
| The Odds API reservation ledger is untouched | quota control is a cost guard, not an access rule, and scraping may not be used to route around a budget decision |
| No login, credential, paywall or entitlement circumvention | the project uses its own keys or none |
| One request per second floor, hard per-run caps, stop on 403/429 | a blocked source is a CoverageGap, never a retry loop |
| Every scraped body retained as evidence carries URL, retrieval UTC, byte count and sha256 | evidence class must be checkable rather than asserted |
| Scraped prices keep bookmaker, market, line, period, observation time and pre-match state | the named-book/aggregate distinction is what keeps calibration honest |
| Anything averaged, best-of, consensus or exchange-derived stays `research_only` | unchanged — the override is about access, not about what a price proves |
| Unreachable or unparseable pages become CoverageGaps | fail closed, exactly as elsewhere |
| Frozen receipts, stakes, bank, ROI and settled outcomes are never rewritten | no retroactivity, whatever the source policy |

Scraping is therefore permitted, bounded and recorded. It is not a licence to
invent a fixture, a result, a probability or a price. Scraped bodies are
recorded in `tests/fixtures/CAPTURE-PROVENANCE.md` under a distinct class
(`scraped_public_json` / `scraped_public_page`) rather than being described as
first-party captures or provider API captures.

It does not schedule a source-validation call, add a provider key, give a
candidate source ticket authority, or change the current auto-ticket cutoff
policy:

```text
1200_sast_or_first_slate_tip_minus_60_minutes
```

## Implementation order

*(2026-11-23)* Reordered for expansion. Steps 1–2 are now the only
prerequisites for a lane; the reporting increments move to the end.

1. Enumerate price coverage per lane from the available providers' public key
   lists, so that "unpriced" is a measured fact per lane rather than an
   assumption carried over from the NBA.
2. Bring one expansion lane at a time to real fixtures: league-owned public
   JSON first (KBL `api.kbl.or.kr`, TBF `miniappapi.tbf.org.tr`, B.League S3
   contexts, ACB, LBA, ABA, Israel), media-enterprise fallback second, each with
   one committed scrape receipt and hermetic parser tests. Australia NBL is
   already live on the ESPN relay and stays there.
3. Make unpriced forecasts visible per lane so they contribute to calibration
   while remaining unmistakably not tickets, without rewriting historic cards.
4. Bounded price validation for a lane only once its fixtures, history and
   forecast surface exist and its declared source contract is recorded.
5. Promote a lane to ticket pricing only after its live evidence satisfies the
   identity, price-class, timestamp, staleness and provider-budget gates.
6. Keep the derived readiness report as a reporting increment: useful once
   several lanes are moving, not a prerequisite for any of them.

Each later step is a separate reviewable change. No source discovery result,
readiness label, or passing parser fixture automatically promotes a provider to
scheduled acquisition or paper-ticket pricing.
