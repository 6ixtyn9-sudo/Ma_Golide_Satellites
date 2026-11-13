# Operator picks and paper-ticket grading

`scripts/daily.py` now publishes an Edge-style operator layer directly under
the selected league's `localdata` directory. Keep each league on its own
`LOCALDATA` path.

The workflow has two complementary outputs:

- artifacts retain operational reports and recovery evidence for 14 days;
- the serialized `publish-localdata` job commits the allowlisted pick, ticket,
  audit, and performance files to `localdata/<league>` on `main`, making them
  visible in the repository and durable across cold caches, as Edge does.

The publisher overlays current files without deleting older history and aborts
if a recovered immutable decision receipt differs from the repository receipt.
After cache restoration, Git-tracked operator state takes precedence over stale
cached copies. DuckDB, CSV, acquisition ledgers, and credentials are never
committed by this publisher. Derived Git publication runs for both manual and
scheduled executions; it is independent of the canonical `ALLOW_MAGOLIDE_WRITES`
gate.

This is a presentation and paper-settlement layer. It does not change forecast
selection, alter calibrated probabilities, place bets, write to a bookmaker, or
add a fade strategy. Its labels describe Ma Golide's actual eligibility and
evidence; they are not Edge Factory labels or a claim of historical
profitability.

## Files

| File | Purpose |
| --- | --- |
| `picks_today.json` | Current scan of every forecast; ROI-qualified singles carry their band decision and diagnostics |
| `picks_YYYY-MM-DD.json` | Cumulative dated pick evidence, keyed by forecast id, with grades and the scan-time ROI-band decision |
| `picks_YYYY-MM-DD.txt` | Human-readable latest picks, marking ROI-qualified paper singles, plus started/settled archive sections |
| `picks_audit_rolling.json` | Thirty-day logical and settled ledgers plus flat-stake diagnostics |
| `picks_audit_YYYY-MM-DD.md` | Readable rolling summary and every settled logical pick in the window |
| `auto_tickets_YYYY-MM-DD_decision.json` | Immutable publication receipt for that date's paper accas |
| `auto_tickets_YYYY-MM-DD.txt` | Frozen receipt with leg grades, or an explicit unfrozen no-bet scan card |
| `auto_tickets_state.json` | Write-once date locks and bounded scan telemetry |
| `auto_tickets_performance.json` | Complete reconstructed paper-performance state and history |
| `auto_tickets_performance.txt` | Bank, settled record, bet-days, held accas, and grading notes |

The daily summary's `operator` object lists the generated paths and counts.

## Pick categories

- **ELIGIBLE CLEAN PICKS**: `Forecast.eligibility.acca_eligible` is true. This
  means the existing exact-book, calibration, probability, and push-risk guards
  passed; it does not certify profit.
- **CAUTION PICKS**: a real selected price exists, but the segment lacks enough
  walk-forward samples. These rows remain strictly ineligible. A row whose only
  blocker is this sample minimum and whose segment has at least one graded
  non-push observation may appear in a provisional paper acca; the frozen ticket
  marks that exception explicitly.
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

Repeated scans can produce distinct forecast IDs because book and prediction
observation times are part of those IDs. The dated JSON retains every evidence
row. The text card and rolling audit select the latest observation for the same
logical game/market/period/selection/line/model/book pick, and report the number
of superseded observations. Only pending pre-tip logical picks appear in the
active bucket sections; started games awaiting an official final and settled
picks appear in separate archive-only sections. Rolling hit rate, ROI, and
profit therefore count one paper pick per logical selection rather than one
paper pick per scan. Different selections, lines, models, or books remain
separate logical picks. The rolling JSON retains the complete logical ledger
and a dedicated `settled_ledger`; the readable audit lists every settled
logical pick in its 30-day window, including unpriced grades that do not enter
ROI.

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

## Publication and freeze

When `accas.paper_results` is present, the operator uses that paper-only
candidate surface; `accas.results` remains the strict surface. Older reports
without `paper_results` retain their prior behavior. For current
`MarketForecastService` reports, each unique paper-eligible wager selection
must match a positive settled individual-pick ROI band before ticket assembly.
The rolling audit's `settled_ledger` is used when available; otherwise the
service uses its own past-only prequential settled forecasts. Pending/future
rows, acca outcomes, missing calibrated probabilities, and picks without an
exact pre-tip price are excluded. Bands are split by league, market, period,
model source, calibrated-probability range, and individual-odds range. Flat
one-unit ROI uses win profit `price - 1`, loss profit `-1`, and push profit `0`;
pushes remain in the ROI denominator. A 14-day recency half-life and three-
pick hierarchical shrinkage are applied, but no long minimum-sample waiting
period is imposed: a band can start after one settled pick. The candidate band
must itself have positive recency-weighted and adjusted ROI, so a band with no
history cannot inherit eligibility from a positive parent.

The adaptive decision is recalculated for each scan and may admit or suspend
future picks as outcomes accumulate. Its source, settled-pick counts, per-band
observed/recency-weighted/adjusted ROI, and pass decisions are reported in
`accas.paper_pick_filter`, the empty-scan candidate funnel, and each new ticket
receipt. The gate is per individual pick, not combined acca odds or EV, and is
never relaxed to fill a ticket. The existing sample-only provisional allowance
and its other eligibility safeguards remain in force. Positive historical ROI
is an exploratory paper signal, not proof of statistical edge or future profit.

Every positive-band candidate is flagged as a paper single in the current and
dated pick JSON/text, and its outcome is tracked in the individual-pick audit.
This remains available even when there are too few distinct games to make a
combination. The requested leg count is an upper bound for paper combinations:
the service reduces it to the number of distinct ROI-qualified games, with a
two-leg minimum. Thus the default three-leg request becomes a double when only
two games qualify, and produces singles only when just one game qualifies. No
pick-quality or price gate is relaxed to fill a combination.

Combinations are greedily accepted in descending joint-probability order and do
not reuse an exact wager selection (league, game, market, period, selection, and
line) across separate acca tickets. A different selection from the same game may
appear on another ticket, while same-game legs are still prohibited within an
acca. This is a no-repeated-leg rule, not mathematical outcome exclusivity:
separate tickets may still win simultaneously. Strict `accas.results` rankings
are unchanged and may overlap. The first non-empty multi-leg paper acca report
for a league/date writes a decision receipt and a write-once state lock. This
includes a provisional paper acca when the sample-size gate alone is bypassed.
The receipt stores the settled-pick ROI-band snapshot and provisional
status/sample counts. Subsequent scans reprint the receipt even if the new
candidate set differs. A missing locked receipt is an error; the engine never
silently rebuilds or rotates a published card. A receipt recovered without its
state lock adopts a lock from the immutable receipt itself. Existing frozen
receipts are never rebalanced or rewritten by selection-policy changes.

A withdrawal is a deliberate exception recorded in `auto_tickets_state.json`:
its decision ID is excluded from performance, its date cannot receive a
replacement ticket, and the Git publisher rejects that ID if an old cache
returns it. The original receipt is never rewritten; withdrawals preserve it as
evidence.

Each acca uses exactly one flat paper unit. The initial 100-unit figure is only
an accounting base, not money or a recommended bankroll. Empty scans are logged
but do not cancel earlier positions or consume ticket identities.

An empty latest-scan card includes a **candidate funnel**: current rows, exact
priced rows, calibration-sufficient rows and their overlap, strict eligible
rows/games, provisional sample-only rows/games, paper-candidate coverage,
settled individual-pick ROI-band counts and observed/adjusted performance,
paper/strict acca counts, ROI-qualified singles, the requested/effective
combination size, and overlapping eligibility-blocker counts. The same
structure is stored under `candidate_funnel` in the bounded scan-state entry
and operator result. These are diagnostics only; they never override an
eligibility guard or synthesize a pick or ticket.

Dated pick ledgers also retain earlier rows when later scans omit them. Fresh
versions of the same forecast id replace the dated evidence row; distinct
forecast ids remain distinct in JSON even when they are repeated observations
of one logical text/audit pick. `picks_today.json` never replaces dated history.

## Grading

Grading is separate from publication evidence. Every retained dated ledger is
refreshed on each run, including dates before the current target date. Only a
final game whose actual period scores reconcile to the official final can settle
a forecast. Missing, incomplete, conflicting, or unreconciled results remain
**PENDING**.

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
`magolide.paper_fade_audit.v1`. It is a counterfactual audit, not a new market,
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
It reports flat-stake ROI only when settled priced observations exist. It never
invents a return for an unresolved ticket. `bank*` is a latest-known-result
mark by ticket date, not a settlement-time equity curve.

Run the normal daily command again after results acquisition to regrade all
retained picks and receipts without modifying their original predictions,
prices, publication timestamps, or receipts. `--skip-results` leaves existing
positions unchanged until a later reconciled result is available.
