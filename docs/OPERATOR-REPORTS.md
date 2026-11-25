# Operator picks and paper-ticket grading

`scripts/daily.py` now publishes an Edge-style operator layer directly under
the selected league's `localdata` directory. Keep each league on its own
`LOCALDATA` path.

The workflow has three complementary output tiers, following Edge Factory's
raw-dump storage approach:

- the league Actions cache retains working forecast archives and audit inputs
  across active runs; downloaded Actions artifacts retain the raw picks JSON,
  dated Markdown, operational reports, and recovery evidence for 14 days;
- Git tracks immutable ticket/no-bet decisions, human-readable ticket and pick
  cards, compact date locks (including inline immutable per-fixture pick
  snapshots), and compact performance state;
- structured dated pick JSON, detailed dated lifecycle renders, `picks_today.json`,
  rolling/dated expanded audits, and other scan payloads remain cache/artifact-only
  so volatile structured detail does not become permanent Git history.

A dated pick card begins as a mutable draft. At each fixture's scheduled tip
minus 120 minutes, the operator writes one immutable fixture snapshot inline in
`pick_lifecycle_state.json`. Later settlement reads the snapshot rather than a
refreshed model or price scan, while the detailed lifecycle view records the
frozen/settled status in cache and Actions artifacts. The publisher overlays
Git-durable files without deleting older history and aborts if a recovered
immutable receipt or inline snapshot differs from repository evidence.
Git-tracked operator state takes precedence over stale
cached copies. If the rolling pick audit is absent after a cold-cache recovery,
the forecaster uses its documented past-only prequential fallback rather than
inventing or importing stale picks.
DuckDB, CSV, acquisition ledgers, raw current-scan JSON, and credentials are
never committed by this publisher. Only named external-trigger executions
publish derived state to Git; manual runs retain inspection artifacts only. This
boundary is independent of the canonical `ALLOW_MAGOLIDE_WRITES` gate.

This is a presentation and paper-settlement layer. It does not change forecast
selection, alter calibrated probabilities, place bets, write to a bookmaker, or
add a fade strategy. Its labels describe Ma Golide's actual eligibility and
evidence; they are not Edge Factory labels or a claim of historical
profitability.

## Files

| File | Purpose |
| --- | --- |
| `picks_today.json` | Current scan of every forecast; cache/artifact only, not Git |
| `picks_YYYY-MM-DD.json` | Structured dated draft evidence; cache/Actions-artifact only for every league, never Git-published |
| `pick_lifecycle_state.json` | Git-published lifecycle state with immutable per-fixture evidence nested by date and fixture, captured at the 120-minute cutoff: selections, prices/lines, source/evidence timestamps, scheduled tip, and freeze timestamp; no per-fixture sidecar is published |
| `picks_YYYY-MM-DD_lifecycle.json` | Detailed mutable lifecycle view; grades frozen evidence against reconciled official finals and labels draft/frozen/settled state. Cache/Actions-artifact only for every league, never Git-published. |
| `picks_YYYY-MM-DD.txt` | Git-published human-readable pick card rendered from the lifecycle view |
| `picks_audit_rolling.json` | Thirty-day logical and settled ledgers; cache/artifact only, not Git |
| `picks_audit_YYYY-MM-DD.md` | Expanded rolling summary; cache/artifact only, not Git |
| `auto_tickets_YYYY-MM-DD_decision.json` | Immutable publication receipt for that date's paper singles and accas |
| `auto_tickets_state.json` | Compact Git-durable state and immutable date locks, including inline no-bet decisions; no per-date no-bet sidecar is published |
| `auto_tickets_YYYY-MM-DD.txt` | Mutable, Git-published current card. Before the cutoff its clearly labelled `DRAFT ONLY` / `PRE-CUTOFF` section may change on each scan and creates no position; after an immutable receipt it becomes the settlement card whose selections, prices, and stakes remain frozen while grades refresh; after an inline no-bet lock it displays that terminal no-bet state. |
| `auto_tickets_state.json` | Git-durable write-once date locks; bounded scan telemetry stays cache/artifact-only |
| `auto_tickets_performance.json` | Compact rolling settlement/financial summary; each history day points to its immutable receipt for full leg evidence |
| `auto_tickets_performance.txt` | Bank, settled record, bet-days, held positions, and grading notes |

The daily summary's `operator` object lists the generated paths and counts.

### Legacy compact-marker recovery

Readable dated cards are Git-published directly. The temporary compact-marker
format used during the 2026-11-20/21 archive transition is recoverable with
`scripts/restore_pick_cards.py`; it is an operator migration tool, not a normal
forecast or ticket stage. The tool is dry-run by default and only replaces a
registered `picks_YYYY-MM-DD.txt` marker after validation:

- a release-archive card must pass the Git-published archive checksum, content
  manifest, member digest, state/date identity, readable-card shape, **and**
  the marker's detailed-card SHA-256;
- a deliberately selected pre-compaction Git snapshot may restore its readable
  historical version when the marker's later artifact revision is unavailable;
  the report labels that source as `git_snapshot` rather than claiming an exact
  marker-hash recovery; and
- an unavailable source leaves the marker untouched and is reported. Use
  `--require-complete` when a migration must fail rather than publish a partial
  result.

The normal publisher also rejects an opaque marker from a stale Action artifact;
it cannot overwrite a readable Git card after this migration.

For example, an operator first downloads a retained release asset with its
original `magolide-state-<state>.tar.gz` name, then runs:

```bash
python scripts/restore_pick_cards.py --repo . \
  --snapshot-ref <pre-compaction-commit> \
  --archive /safe/path/magolide-state-nba.tar.gz \
  --apply --require-complete
```

This restores report presentation only. It does not create tickets, alter
frozen selections, relax bookmaker-evidence rules, or modify paper-bank,
ROI, stake, or settlement state.

To keep Git history bounded, the rolling performance JSON contains only compact
financial/settlement fields and minimal leg identifiers, outcome, selected price,
and bookmaker. It does not duplicate prediction, calibration, team-history, or
full book evidence from each immutable receipt. Read the receipt named by a
history row's `receipt_path` for complete provenance. Before Git publication,
the publisher also removes bounded scan telemetry and generation timestamps
from the mutable state/performance JSON. Those fields remain available in the
cache and 14-day artifacts, but an otherwise unchanged hourly run produces no
new Git blob. Before the cutoff, a mutable `auto_tickets_YYYY-MM-DD.txt` card
contains the clearly labelled `DRAFT ONLY` / `PRE-CUTOFF` section and is
deliberately replaced by later scans; it is not a receipt, a date lock, or a
committed paper position. At the cutoff, an immutable receipt or no-bet sidecar
makes the decision. Later updates refresh a receipt-backed card when official
settlement changes a displayed grade and refresh the text card to display a
terminal no-bet decision; the sidecar remains the durable terminal record.

## Pick categories

- **ELIGIBLE CLEAN PICKS**: `Forecast.eligibility.acca_eligible` is true. This
  means the existing exact-book, calibration, probability, and push-risk guards
  passed; it does not certify profit.
- **CAUTION PICKS**: a real selected price exists, but the segment lacks enough
  walk-forward samples. These rows remain strictly ineligible. A row whose only
  blocker is this sample minimum and whose segment has at least one graded
  non-push observation may appear in a provisional paper acca; the frozen ticket
  marks that exception explicitly.
- **PRIOR-HEAVY — THIN TEAM EVIDENCE (PAPER ONLY)**: a real selected price
  exists and the segment is calibrated, but the row's only remaining blocker is
  `insufficient_team_specific_history`: one team's own recency-weighted evidence
  is thinner than the `team_prior_games` strength the model requires, so the
  probability leans on the league-mean prior. These rows stay strictly
  ineligible (no acca leg, no ticket, no ROI-band promotion); they are surfaced
  rather than discarded so the forecast is not silently lost. Each row prints
  both teams' effective games against the required strength and the value at
  the observed price (break-even versus p), because a prior-heavy row can be
  clearly negative at the quote.
- **WATCHLIST — NO MATCHED ODDS**: no exact selected bookmaker price exists.
  These rows remain visible and are not acca candidates.
- **WATCHLIST — INSUFFICIENT CONTEXT**: priced rows whose remaining exclusion
  is limited to evidence/probability context.
- **SKIPPED — INELIGIBLE**: another eligibility guard failed. The exact
  `eligibility.reasons` are printed on the card.

Every row includes fixture, tip-off in Africa/Johannesburg time, native period,
selection and line, raw/calibrated probability, selected price and bookmaker,
price evidence, model/prediction provenance, calibration sample count,
eligibility reasons, and grade.

No-price forecasts remain **audit-only**, never ticket legs. For leagues other
than NCAA_M, the rendered daily auto-ticket card shows only the latest logical
no-price forecasts that pass a historical success-rate gate; qualifying lines
include forecast ID, selection, probability, historical hit rate, price-evidence
state, blockers, and current grade. Those qualifying picks are partitioned into
display-only shadow groups of up to three, with at most one pick from each
fixture per group; each pick appears once. The groups have no odds, ticket ID,
receipt, stake, combined hit rate, or ticket-bank/ROI impact. The gate is wins
divided by wins plus losses among prior settled individual picks in the rolling
30-day audit, matched on league, market, period, model, and calibrated-
probability band (no odds band). Repeated observations or book snapshots of
the same settled wager count once. The gate requires at least 70% success with
more than 10 win/loss decisions; pushes, pending/void/unknown outcomes, the
same game, and historical picks scheduled less than four hours before the
forecast observation are excluded. A missing comparison band or insufficient
history fails closed. Historical priced and unpriced picks both contribute
settled win/loss outcomes; no current price or ROI is inferred.

NCAA_M uses a separate **priced-ineligible** shadow display instead of the
unpriced shadow group. It includes only forecasts with an exact selected-side
decimal price of at least 1.30, existing `acca_eligible: false`, and current
probability of at least 70%. Calibrated probability is used when present, with
raw probability as the display fallback. Candidates are grouped display-only
(up to three, at most one per fixture); the price, bookmaker, eligibility
blockers, and grade remain visible. These thresholds do not change eligibility,
create ticket legs, or affect stake/bank/ROI. NCAA_M unpriced forecasts remain
in the dated and rolling unpriced audit records; they are not substituted into
this priced-ineligible group.

EuroLeague's paper-ticket lane uses a deliberately stricter **priced-value**
filter. It retains the ordinary eligibility (including the narrow provisional
sample-only paper exception), requires a matching exact selected-side decimal
quote observed before tip, and requires positive expected profit at that quote
(calibrated probability × decimal odds − 1 > 0). It also requires at least ten
prior, result-delay-matured settled individual picks in the same league, market,
period, model, calibrated-probability band, and odds band; both recency-weighted
and shrinkage-adjusted ROI for that band must be positive. Only these picks may
enter EuroLeague paper singles/accas. An empty qualifying set fails closed.
This is a stricter monitoring rule, not evidence that EuroLeague picks are
predictively superior, and it does not change forecast probabilities, strict
eligibility, or real-money behavior.

The picks dated ledger and rolling audit remain the grading source of truth;
all audit-only rows have zero stake and do not enter paper-position performance
or ROI. They can never create a ticket receipt or ticket leg; when no priced
candidate qualifies at the scheduled cutoff, the date instead receives the
separate terminal no-bet lock. Full candidate rows and per-pick history
diagnostics remain in the cache-only picks JSON and
run artifacts. To avoid storing the same large row set in every Git-persisted
scan, `auto_tickets_state.json` retains compact counts, blocker/status
histograms, and policy summaries, while dropping duplicate `candidate_records`,
`forecast_ids`, and per-pick ROI-band detail from its bounded scan history.
Immutable date locks and ticket-performance state remain intact. The current
run artifact and picks JSON retain per-candidate details, including for NCAA_M
priced-ineligible shadows; these records remain display-only, with no ticket ID,
ticketability, receipt, or stake. Different markets or selections on one
fixture remain separate wager candidates, not fixture-level aggregates. This
preserves ticket-decision visibility without duplicating large audit rows into
the repository's long-lived state file or relaxing priced-ticket eligibility.

Repeated scans can produce distinct forecast IDs because book and prediction
observation times are part of those IDs. The dated JSON retains every evidence
row. The text card and rolling audit select the latest observation for the same
logical game/market/period/selection/line/model/book pick, and report the number
of superseded observations. A later exact selected-side quote additionally
supersedes an earlier **unpriced** observation of that same
`game/market/period/selection/line/model`, so an old `no matched book` row is
not left active beside its now-priced successor. This is a display/audit
selection only: it never deletes the earlier evidence or invents a quote.
Different selections, lines, models, or named books remain separate logical
picks. Only pending pre-tip logical picks appear in the active bucket sections;
started games awaiting an official final and settled picks appear in separate
archive-only sections. Rolling hit rate, ROI, and profit therefore count one
paper pick per logical selection rather than one paper pick per scan. The
rolling JSON retains the complete logical ledger and a dedicated
`settled_ledger`; the readable audit lists every settled logical pick in its
30-day window, including unpriced grades that do not enter ROI.

The Markdown audit uses proper `Overall`, `Settlement policy`, `Bucket
performance`, and `Settled picks` sections. Settled markets are grouped under
one heading per fixture instead of repeating long single-line rows. Each pick
shows a readable named selection, result, expected probability, exact price (or
its absence), calibration sample, bucket, and model. Production reports include
the reconciled full-game final and relevant period/total/margin context from the
canonical game record; those scores are presentation inputs and are not written
back into frozen pick evidence. The root operator schema remains unchanged; the
rolling JSON has a separately versioned nested `fade_audit` object for the
prospective paper-only counterfactual described below.

## One line per bet, one row per observation

The ledger and the card answer different questions, and they no longer give the
same answer to both.

- The **ledger** keeps one row per observation: every named-book quote, every
  aggregate quote, and every prediction surface, each independently auditable.
  Nothing is deleted, because a settlement dispute turns on those rows.
- The **card** prints one line per modelled selection — fixture, market, period,
  line, and side — and reports the multiplicity on that line instead of
  inflating the number of bets:

```
      MONEYLINE full_game -> HOME · p=64.9% (raw 62.0%) @1.300 BetMGM.us
      collapsed_observations=6 · books_quoted=3: BetMGM.us, Stake.com, bet365.us · surfaces=provider:history_model, team_history_empirical
```

Measured on the frozen 2026-11-24 EuroLeague card: **48 evidence rows, 18
selections, 6 fixtures.** Each selection was quoted by three books and by the
local model twice — as `team_history_empirical` (calibrated, n=27) and as
`provider:history_model`, which carries the *same* raw probability with
calibration `n=0`. Printing both meant one settled loss read as six, and
`Logical picks shown` counted 48.

The surviving line is always one real observation, chosen by
`_display_observation_rank`: named-book quote before aggregate, calibrated before
uncalibrated, freshest quote before higher price, then a stable tie-break on
`pick_id`. A stale 2.45 does not outrank a live 1.90, and nothing is synthesised
by mixing one row's probability with another row's price.

Two sides of one market are *not* merged: opposite selections stay separate
rows, so a source disagreeing with itself about the side remains visible. The
footer reports both counts:

```
Logical picks shown: 18 · superseded observations: 0 · book/model rows collapsed: 30
Total archived pick evidence rows: 48
```

Tickets are unaffected. The acca builder already took at most one leg per
fixture (`accas._best_leg_per_game`), so the duplication never reached a ticket,
a stake, or the paper bank — it was a reporting and counting defect.

## A ticket leg has to be able to pay

Paper tickets are assembled by ranked probability (`accas.build_accas`), and
probability alone will always walk towards the shortest prices on the board. On
2026-11-24 that produced five NCAA_M "3-leg accas" whose headline price was
**@1.003** — one full paper unit risked to return 0.3% — built from legs of
`@1.001` and `@1.012`. The `1.001` quotes are ESPN's `-100000` moneyline on
lopsided games; the lane's own plausibility floor in `ingestion/books.py`
(`1.01`) already called that price implausible, but the relay path never applied
it.

Three rules now bound a paper ticket, and all three live in configuration so a
tuned file can move them and every ticket still carries the `CFG_` hash of the
values that built it:

| key | default | what it does |
|---|---|---|
| `minTicketLegOdds` | `1.20` | no leg cheaper than this — a `@1.001` leg risks a unit for 0.1% |
| `maxTicketLegOdds` | `3.00` | no leg dearer than this — a ceiling keeps the ranking inside quotes the model can speak about |
| `ticketRankingPriceWeight` | `0.15` | how much of the ticket ranking is price rather than probability |

`forecasting/accas.py` reads them as `MIN_TICKET_LEG_DECIMAL`,
`MAX_TICKET_LEG_DECIMAL` and `TICKET_RANKING_PRICE_WEIGHT`, and applies the band
in `_eligible_candidate_picks` — before bands, singles or acca legs exist — so
every downstream count agrees on what is on the table.

**The ranking change is a deliberate reversal.** `build_accas` used to maximise
joint probability, and the code said so out loud: *"price is payout evidence,
never a ranking signal"*. That rule is what walked the NCAA_M card to `@1.003`.
Tickets are now ranked by `_ticket_ranking_score`, the blend
`(1-w)·log p + w·log price`, which is additive so the k-best walk stays exact.
What did **not** change: choosing which row represents a fixture
(`_best_leg_per_game`) is still calibrated-probability-first, so a long-priced
row can never stand in for a short-priced row of the same game.

**Why 0.15 and 3.00.** Measured on the real 2026-11-24 NCAA_M card, 55 usable
legs at one-per-fixture:

| weight | ceiling | best ticket | joint p | EV |
|---|---|---|---|---|
| 0.00 | 3.00 | @2.09 | 48.9% | +2% |
| 0.15 | 3.00 | @2.09 | 48.9% | +2% |
| 0.20 | 3.00 | @2.67 | 46.7% | +25% |
| 0.35 | 3.00 | @10.41 | 28.2% | +194% |
| 0.10 | none | @5.58 | 45.4% | +153% |

The last two rows are the warning, not the target: at that weight the ranking
starts buying the model's own overclaims (Rice was rated p=72% against a 3.50
quote, a claimed +152% edge). A higher weight is a value bet on the model being
right about that; it is not free odds.

The card is unaffected: a forecast outside the ticket band still appears with its
price and its eligibility blockers. It just cannot be sold as a ticket.

### Why a ticket card can say "no acca" while singles qualified

A single and a multi-leg ticket ask different questions. `filter_paper_candidate_picks`
admits a row when its `(market, probability band, odds band)` has positive
settled ROI; `build_accas` additionally requires `acca_eligible` (or the
provisional-sample-only exception) **and** the ticket price band. So a lane can
honestly report nine ROI-qualified picks and still have nothing to combine.

The card used to answer that with "no multi-leg acca was produced", which is
true and useless. The funnel now carries `paper_acca_eligible_games`,
`paper_acca_eligible_rows` and `paper_acca_blockers` computed from
`accas.ticket_leg_pool` — the builder's own pool — and the card names the gate:

```
NO MULTI-LEG PAPER ACCA — 9 of 14 candidate picks passed individual-pick ROI;
no ROI-qualified pick is ticket-eligible (the acca builder applies the price band
and eligibility gates on top of the ROI filter); blockers: no_exact_bookmaker_market.
```

Measured case: NCAA_W, 2026-11-24.

### NBA: what the provisional-sample relaxation actually bought

Operator decision 2026-11-22 relaxed the walk-forward sample gate so a leg whose
only blocker is `insufficient_walk_forward_samples` may still price a provisional
paper ticket. The NBA lane is where that bites, so its settled tickets were split
by whether the provisional gate was used:

| tickets | n | wins | hit | staked | returned | ROI |
|---|---|---|---|---|---|---|
| provisional gate, 3-leg accas | 21 | 3 | **14.3%** | 21.00 | 20.462 | **−2.6%** |
| strict gates, singles | 2 | 2 | 100% | 2.00 | 3.693 | +84.7% |
| provisional gate, single | 1 | 1 | 100% | 1.00 | 1.769 | +76.9% |

Read it carefully rather than optimistically:

- every NBA multi-leg ticket in the archive rode the relaxation — there is no
  strict-gate acca to compare it against;
- a 3-leg acca of ~0.68 legs should land near 31%; 14.3% of 21 is well below
  that, which is what a leg whose probability has not cleared walk-forward
  samples looks like on the card;
- the lane's headline `ROI +8.0%` comes from two 7× wins plus three singles, not
  from the accas, which are a small net loser;
- n=21 is still small, so this is a reason to keep watching, not a verdict.

## The strict lane's declaration is code

`EUROLEAGUE_STRICT_NAMED_BOOK_VALUE_TICKET_LANE` has always printed
`requires_named_book_quote: true` and *"aggregate consensus quotes are excluded"*.
Until 2026-11-25 that was prose. The shared paper policy accepts an aggregate
consensus average so that aggregate-only lanes are not permanently limited to
zero tickets (operator decision 2026-11-22, `BookEvidence.ticket_price_eligible`),
and the strict lane inherited that acceptance as though it were its own.

On 2026-11-25 the run showed what that cost: of 11 eligible EuroLeague rows, 5
were priced on quotes whose own label reads *"AGGREGATE CONSENSUS AVERAGE — NOT A
NAMED-BOOK QUOTE"*, and all four strict acca candidates carried one or two such
legs. Nothing on any card could have said so — the reason vocabulary had no token
for provenance. The tickets did not publish, because the positive-EV gate blocked
every candidate independently (best EV −0.450/unit). That was luck, not a gate.

`filter_paper_candidate_picks(require_named_book_quote=True)` now enforces the
declared contract: `price_provenance` must be `named_book`, and the excluded rows
are counted rather than dropped silently.

| where | key |
|---|---|
| `paper_pick_filter` | `require_named_book_quote`, `excluded_no_named_book_quote_count` |
| `candidate_funnel` | `paper_excluded_no_named_book_quote_rows` |
| `paper_lane` (EuroLeague only) | `named_book_requirement_enforced`, `excluded_no_named_book_quote_candidates` |

The shared policy is untouched: a lane that does not declare the strict contract
still prices consensus averages, and every card line still carries the
provenance label.

## Two funnel numbers that legitimately disagree

`candidate_funnel.eligible_rows` and `priced_and_calibrated_rows` count different
populations. Eligibility also admits a row whose calibration did not reach
`sufficient_for_accas` — a provisional admission, or a probability floor that
passed on a thin sample. So `eligible_rows` may exceed
`priced_and_calibrated_rows`; on 2026-11-25 EuroLeague printed 11 over 10, one
line apart, and the arithmetic looked broken.

It was not. The funnel now names the gap instead of hiding it:

```
"eligible_rows_without_calibration_sufficiency": 1,
"funnel_reconciliation": {
  "eligible_rows_exceed_calibrated_rows": true,
  "explained_by_rows_admitted_without_calibration_sufficiency": 1,
  "unexplained_eligible_rows": 0
}
```

`unexplained_eligible_rows` is the number that matters: it should always be zero.
A non-zero value is a real counting bug.

## Counters that describe a receipt, not a scan

`counts.published_accas` and `counts.published_singles` count tickets in the
**stored receipt** for that date. That receipt may have been frozen by an earlier
run, under earlier rules — on 2026-11-24, NCAA_M's was frozen at 10:07Z holding
five accas built before the leg-price floor existed. Read beside a current
`paper_accas: 0` it looks like a contradiction; it is history.

Both counters are now labelled:

```
"published_counts_scope": "stored_receipt_not_this_run_funnel",
"published_receipt_decision_id": "45366895912_ncaa_m_1_attempt_1",
"published_receipt_frozen_at": "2026-11-24T10:07:12Z"
```

## A missing import is not a timetable

The adaptive gate fails open when a schedule cannot be read, recording
`schedule_uncertain: <detail>`. It used to do that for every exception, so on
2026-11-25 both Japan lanes went dark recording
`schedule_uncertain: No module named 'requests'` — a broken dependency in the
gate's own runner, in a run whose preflight had installed `requests` and passed
1672 tests ten minutes earlier. A `ModuleNotFoundError` now yields
`environment_error: <detail>`: the lane still runs, but the label says CI rot
rather than a league without fixtures.

## One bet, one row

A lane can price the same bet twice. On 2026-11-25 EuroLeague did it for every
priced key: a `provider_prediction` twin at p≈0.484 blocked by the probability
floor, and a `team_history_empirical` twin at p≈0.548 that passed — same fixture,
market, period, side and price, two calibration versions, two verdicts.

Both rows were *evaluated independently*, which made the floor selectable: the
same bet was eligible or ineligible depending on which twin the gate read, and
the pick card showed both verdicts at once. The 11-eligible-rows funnel was
counting rows, not bets.

`harness.resolve_calibration_twins` now enforces one wager-eligible row per
betting key `(league, game_id, market, period, selection, line)`:

- the winner is the row with the most calibration evidence — `sufficient_for_accas`
  first, then `calibration.n`, then probability, then row id (so the answer never
  depends on emission order);
- every other twin stays on the card as audit-only, carrying
  `superseded_by_calibration_twin`, `superseded_by_forecast_id` and
  `superseded_by_calibration_version`;
- when the twins would have disagreed, the loser also carries a
  `calibration_twin_disagreement` advisory, so the disagreement is stated rather
  than resolved in silence;
- `candidate_funnel.calibration_twin_superseded_rows` and
  `calibration_twin_disagreement_rows` count both.

**Not "the strictest verdict wins".** That reads as the cautious choice and is the
same bug inverted: it hands a veto to whichever model is worst, so a lane could
block every bet by also emitting a badly-calibrated pessimistic twin. The floor
has to be a property of the bet, decided by its best evidence — not of the model
someone chose to consult.

## A failed scan says so, and keeps retrying

Coverage markers used to be written only for lanes whose registry entry says
`source_unavailable`. That is a static contract, and on 2026-11-25 it produced a
cliff: Greece wrote markers through 11-24, EuroCup and Israel through 11-22,
Japan Premier through 11-23 — and then nothing, exactly on the day nine lanes
scanned zero rows because providers started failing. The signal appeared for days
and vanished the moment it mattered.

Two marker statuses now exist, and the difference is what the gate does with them:

| status | means | gate |
|---|---|---|
| `source_unavailable` | the registry has no compliant adapter | may skip the lane for the rest of the day |
| `scan_failed` | the lane has a source and it did not answer usefully | **never skips** — reports `coverage_gap_scan_failed: <detail>` and retries |

A `scan_failed` marker carries the provider's own words (`detail`), the failure
`gap_reasons`, `required_fields`, `fixtures_found`, `scan_rows`,
`transport_errors` and `attempted_utc`. Writing it requires evidence of failure —
a gap whose reason is a failure state, or transport errors. **Zero fixtures with
no recorded failure writes nothing**: that is an honest empty slate, and calling
it a gap is exactly the label blur the distinction exists to prevent. Markers
written before the `status` field existed are still read as `source_unavailable`,
which was the only kind the writer produced.

## Publication and freeze

When `accas.paper_results` is present, the operator uses that paper-only
candidate surface; `accas.results` remains the strict surface. Older reports
without `paper_results` retain their prior behavior. For current
`MarketForecastService` reports, each unique paper-eligible wager selection
must match a positive settled individual-pick ROI band before ticket assembly.
The rolling audit's `settled_ledger` is used when it yields at least one usable
settled individual pick; if the audit is absent, empty, or has no usable priced
rows, the service uses its own past-only prequential settled forecasts. A
cold-start fallback is explicitly marked in `paper_pick_filter.history_fallback`.
Pending/future rows, acca outcomes, missing calibrated probabilities, and picks
without an exact pre-tip price are excluded. Bands are split by league, market,
period, model source, calibrated-probability range, and individual-odds range.
Flat one-unit ROI uses win profit `price - 1`, loss profit `-1`, and push profit `0`;
pushes remain in the ROI denominator. A 14-day recency half-life and three-
pick hierarchical shrinkage are applied, but no long minimum-sample waiting
period is imposed: a band can start after one settled pick. The candidate band
must itself have positive recency-weighted and adjusted ROI, so a band with no
history cannot inherit eligibility from a positive parent.

The adaptive decision is recalculated for each scan and may admit or suspend
future picks as outcomes accumulate. Its source, settled-pick counts, per-band
observed/recency-weighted/adjusted ROI, and pass decisions are retained in
`accas.paper_pick_filter`, the structured scan `candidate_funnel`, and each new
ticket receipt. The gate is per individual pick, not combined acca odds or EV, and is
never relaxed to fill a ticket. The existing sample-only provisional allowance
and its other eligibility safeguards remain in force. Positive historical ROI
is an exploratory paper signal, not proof of statistical edge or future profit.

Every positive-band candidate is flagged as a paper single in the current and
dated pick JSON/text, and its outcome is tracked in the individual-pick audit.
When exactly one ROI-qualified selection exists, the operator freezes that pick
as a one-unit paper single. The singleton uses the same exact-price, eligibility,
and positive-ROI-band decision as the candidate surface; it does not lower a gate
or manufacture another leg. If multiple selections qualify but no multi-leg
combination can be built (for example, they are all from one game), the operator
does not issue a separate ticket for each selection. The requested leg count is
an upper bound for paper combinations: the service reduces it to the number of
distinct ROI-qualified games, with a two-leg minimum. Thus the default three-leg
request becomes a double when only two games qualify. No pick-quality or price
gate is relaxed to fill a combination.

Combinations are greedily accepted in descending joint-probability order and do
not reuse an exact wager selection (league, game, market, period, selection, and
line) across separate acca tickets. A different selection from the same game may
appear on another ticket, while same-game legs are still prohibited within an
acca. This is a no-repeated-leg rule, not mathematical outcome exclusivity:
separate tickets may still win simultaneously. Strict `accas.results` rankings
are unchanged and may overlap.

A scan does **not** freeze the first non-empty ticket it sees. It remains
candidate evidence until the dated card's universal Johannesburg publication
time: **12:00** for every league. The first scheduled fixture on the league's
acquisition-timezone slate can
supply an earlier 120-minute pre-tip safety cutoff. The first successful scan at
that cutoff freezes all currently valid candidates whose every leg is
still pre-tip. It records the policy, planned cutoff, source timezone, actual
scan, and candidate counts in the receipt. This includes a provisional paper
position when the existing sample-size-only exception applies. Candidates can
therefore gain a later exact price before the decision point without an early
partial market locking the date.

Before the cutoff, that single dated card also carries a clearly labelled
`DRAFT ONLY` / `PRE-CUTOFF` named-price draft section: it may display one
current exact pre-tip price per game. Unresolved source-permission
and operational-validation labels remain audit metadata; they do not
independently block that quote from the paper-ticket lane. Legs are labelled
**NAMED-BOOK PRICE**, **AGGREGATE-CONSENSUS PRICE — NOT A NAMED-BOOK QUOTE**, or
**NOT TICKET-ELIGIBLE EVIDENCE**, preserve their model/calibration/ROI
blockers, rank only by forecast confidence (not a fabricated EV), and create
no receipt, stake, grade, bank/performance entry, or real-money instruction.

Since 2026-11-22 an aggregate consensus average is a paper-ticket price
(`TICKET_PRICE_PROVENANCE`): it is an observed number with a published
provenance, and refusing it left eleven aggregate-only lanes publishing nothing
but `(none)` for ever. It is never relabelled a bookmaker quote -- every card
line, receipt field, and audit row carries the provenance and its label. A
`no_price` record and a bare exchange match remain outside the lane.
There is no separate preview sidecar artifact; a strict receipt or immutable
no-bet decision renders the terminal card without any draft section.

If there is no qualifying pre-tip candidate at the strict decision point—or if
the scheduler missed the first tip—the operator records an immutable inline
no-bet decision in that date's `auto_tickets_state.json` lock. This fail-closed
decision prevents a later price or an old cache from publishing a replacement
ticket without creating a per-date JSON sidecar. New ticket receipts store
singles separately from accas and preserve the settled-pick ROI-band snapshot
and provisional status/sample counts. Subsequent scans re-render each dated text
card from the immutable ticket receipt (or inline no-bet decision), current
reconciled leg results, and the dated pick ledger when available. Ticket and
audit-shadow grade labels therefore refresh, while selections, prices, stakes,
receipt bytes, and terminal no-bet outcomes remain frozen. A missing locked
receipt or malformed inline no-bet decision is an error; the engine never
silently rebuilds or rotates a published decision. A receipt recovered without
its state lock adopts a lock from the immutable receipt itself.
Existing frozen receipts are never rebalanced or rewritten by selection-policy
changes.

A withdrawal is a deliberate exception recorded in `auto_tickets_state.json`:
its decision ID is excluded from performance, its date cannot receive a
replacement ticket, and the Git publisher rejects that ID if an old cache
returns it. The original receipt is never rewritten; withdrawals preserve it as
evidence.

Each paper single or acca position uses exactly one flat paper unit. A settled
single is credited profit of decimal price minus one on a win, minus one on a
loss, and zero on a push; pending singles reserve one unit. Singles and accas
are both counted in paper-position performance, while their counts remain
separately available. The initial 100-unit figure is only an accounting base,
not money or a recommended bankroll. Empty scans are logged but do not cancel
earlier positions or consume ticket identities.

The human-readable empty latest-scan card is intentionally concise: it leads
with **NO PAPER BET** or **NO MULTI-LEG PAPER ACCA** and a brief reason. Before
the cutoff it also shows the pending decision time; after a terminal no-bet it
states that later price evidence remains audit-only. It omits
the full funnel, per-band ROI breakdown, and blocker list. Those diagnostics
remain structured under `candidate_funnel` in the bounded scan-state entry and
operator result, including current/exact-priced/calibrated rows, eligibility
coverage, provisional sample-only rows, paper/strict acca counts, qualified
singles, requested/effective combination size, ROI-band evidence, and
eligibility blockers. These diagnostics never override an eligibility guard or
synthesize a pick or ticket.

Dated pick ledgers retain earlier rows when later pre-cutoff scans omit them.
Fresh versions of the same forecast id replace dated draft evidence; distinct
forecast ids remain distinct in JSON even when they are repeated observations
of one logical pick. Once an inline fixture snapshot exists, later model/price
observations for that fixture are excluded from the dated draft ledger. `picks_today.json`
never replaces dated history.

## Grading

Grading is separate from publication evidence. Every retained lifecycle view is
refreshed in the cache/artifact layer on each run, including dates before the
current target date. A frozen fixture is graded only from its immutable freeze
snapshot; an unfrozen legacy row is visibly labelled as such and cannot claim
freeze-backed settlement. Only a final game whose actual period scores reconcile
to the official final can settle a forecast. Missing, incomplete, conflicting,
or unreconciled results remain **PENDING**.

Every official-final `WIN`, `LOSS`, or `PUSH` enters the rolling settled ledger,
hit rate, bucket diagnostics, and priced-profit calculation immediately—even
when its pick date is the report's current slate date. This is deliberately more
responsive than Edge Factory's default same-day exclusion because Ma Golide does
not settle from live source scores. Same-day pending rows remain outside the
prior-day pending/unmatched counter and are reported separately as
`same_day_pending_picks`; `pending_picks_total` exposes all pending rows. The
pending cutoff is the generated time's actual league/operator local date, not
the acquisition slate date. Thus a run generated on 8 November for the 7
November slate does not incorrectly exclude all 7 November picks. For schema
compatibility, `same_day_rows_excluded` now counts only those held same-day
pending rows, while `same_day_settled_picks` reports official finals included
immediately.

## Prospective paper-only fade audit

The rolling JSON includes a nested `fade_audit` with schema
`magolide.paper_fade_audit`. It is a counterfactual audit, not a new market,
pick, ticket, or strategy. A settled model win is recorded as a fade loss, a
model loss as a fade win, and a push remains a push. Each hypothetical fade has
a one-unit stake; a win earns the exact opposite decimal quote minus one, a loss
loses one unit, and a push returns the stake. ROI is P&L divided by all priced
one-unit fades, including pushes.

Only an explicit opposite-side price preserved with that pick's own book
snapshot is eligible. The audit verifies the same game, bookmaker, market,
period, line, and pre-tip observation, and requires a finite decimal price
above 1.0. A selected-side price is not required: an available exact opposite
quote can be audited even when the selected price is missing. The serializer
preserves all five side-price fields and the originating snapshot scope; row
rehydration carries them forward without synthesis.

A unique complement exists for home/away spreads and over/under totals at the
same line. Full-game basketball moneylines are complementary home/away outcomes
unless the snapshot includes a draw price. Every non-`FULL_GAME` moneyline is
excluded because a draw is possible, even if a partial payload omitted its draw
quote. Highest-quarter and other markets without one unique opposite outcome
are also excluded. Missing, invalid, mismatched, or old-schema quote evidence
stays unpriced; no opposite price is inferred from the selected price or from a
different snapshot. Historical rows lacking preserved quote scope or
`side_prices` are not backfilled, and their reason counts appear under
`fade_audit.unpriced_reasons`.

`fade_audit.by_confidence_band` uses numeric selected-outcome probability:
`calibrated_probability` when present, otherwise `raw_probability`. It does not
use the calibration status label. The corresponding market/period/model/band
breakdown is available in
`fade_audit.by_market_period_model_confidence_band`. The aggregate
`break_even_hit_rate` is the mean inverse opposite price over non-push quotes;
it is book-implied and includes market margin, not a de-vigged probability.
These descriptive, prospective metrics do not change the selected-side ROI,
picks, grading, calibration epoch or gates, or paper acca performance.

For accas:

- any known losing leg settles the acca as a **LOSS**, even if other legs are
  unresolved;
- otherwise, any unresolved leg keeps the acca **PENDING** and its stake
  reserved;
- when all legs resolve, push legs are removed from the effective decimal
  price;
- all pushes return the paper stake; remaining wins settle at the effective
  price.

The rolling performance report counts settled and pending positions separately.
In `auto_tickets_performance.json`, `awaiting_game_legs` counts ungraded legs
whose game is scheduled, live, or postponed — across **all** positions, so it
also includes legs inside tickets that already settled as a loss on another leg
and will never affect the bank. The rendered `OPEN` line therefore reports
`awaiting_game_legs_open`, which counts only legs inside still-pending
positions — the number that actually explains why a ticket is held.
`result_unresolvable_legs` counts legs with no matching game, a non-playable or
unknown game status, or a final whose result cannot be reconciled/graded. At the overall and history-day levels,
the unresolved total remains `unresolved_result_legs`; each also has the two
category counts. On each position, the total is named `unresolved_legs`, with
`awaiting_game_legs` and `result_unresolvable_legs` alongside it. Pending legs
carry a `result_resolution_status` of `awaiting_game` or `result_unresolvable`
in the graded performance copy. This diagnostic split does not change pending
grades, stakes, or settlement decisions.

A leg settles on the **team** it was priced on, never on a bare `home`/`away`
label. Receipts record both team names, so a provider that re-designates the
sides of a neutral-site game after the ticket was frozen cannot flip a winner
into a loss: when the frozen names match the result only cross-wise, the result
is mirrored back to the frozen orientation before grading (spread lines are
home-oriented, so the line itself is left untouched). Such a leg keeps
`result_side_identity: "mirrored"` in the graded performance copy and renders as
`· sides mirrored` in the held-ticket block. When a recheck fetch reports the
sides the other way round, acquisition also records one `result_side_identity`
coverage gap for that game, so an operator sees the provider move the fixture
even for games that are not yet on a ticket. Names that match neither side fall
back to the previous side-label grading rather than inventing an orientation.

The report gives flat-stake ROI only when settled priced observations exist. It
never invents a return for an unresolved ticket. `bank*` is a latest-known-result
mark by ticket date, not a settlement-time equity curve.

Run the normal daily command again after results acquisition to regrade all
retained picks and receipts without modifying their original predictions,
prices, publication timestamps, or receipts. `--skip-results` leaves existing
positions unchanged until a later reconciled result is available.
