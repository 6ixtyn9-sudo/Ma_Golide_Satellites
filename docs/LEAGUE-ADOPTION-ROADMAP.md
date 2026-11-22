# League readiness and source-adoption roadmap

**Status:** planning and evidence-governance contract. It does not enable a
provider, make a network request, grant a credential, change a ticket rule, or
rewrite historic operator evidence.

**Reviewed:** 2026-11-18

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
| EuroCup | Schedule/result lane added by reusing the validated EuroLeague Basketball adapter with competition code `U`; fixture-tested, with no first-party EuroCup capture and no validated price contract. | Watch the first scheduled runs for a real fixture and reconciled result. Price work stays on the direct BetExplorer fallback until a named-book contract is separately evidenced. |
| Other expansion leagues | Typed registry and isolated state exist, but schedule/result acquisition is generally unavailable. | Establish compliant fixture/result identity before forecast or price work. A source discovery lead is not an integration. |

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

This roadmap does not approve a new provider, source, price feed, exchange, or
prediction market. It does not authorise scraping, undocumented sportsbook
endpoints, browser automation, proxying, account/key rotation, trial-only
access, or treating aggregate/prediction/exchange prices as named-book ticket
prices.

It does not schedule a source-validation call, add a provider key, give a
candidate source ticket authority, or change the current auto-ticket cutoff
policy:

```text
1200_sast_or_first_slate_tip_minus_60_minutes
```

## Implementation order

1. Implement the derived readiness report and its hermetic tests.
2. Add prospective display metadata so visible unpriced normal-lane picks are
   unmistakably separate from paper tickets, without rewriting historic cards.
3. Add future-source validation proposals, still code-only and unable to make a
   request or expose a secret.
4. Select one high-readiness, evidence-bounded validation only after its source
   contract and cost policy are explicitly approved.
5. Promote a source/lane only after its live evidence satisfies the declared
   identity, rights, region, timestamp, market, and provider-budget gates.

Each later step is a separate reviewable change. No source discovery result,
readiness label, or passing parser fixture automatically promotes a provider to
scheduled acquisition or paper-ticket pricing.
