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
selection, alter calibrated probabilities, place bets, or write to a bookmaker.
Its labels describe Ma Golide's actual eligibility and evidence; they are not
Edge Factory labels or a claim of historical profitability.

## Files

| File | Purpose |
| --- | --- |
| `picks_today.json` | Current scan only; every current forecast, including excluded rows |
| `picks_YYYY-MM-DD.json` | Cumulative dated pick evidence, keyed by forecast id, with current grades |
| `picks_YYYY-MM-DD.txt` | Human-readable latest logical picks, plus started/settled archive sections |
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
  walk-forward samples.
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
back into frozen pick evidence. The rolling JSON schema and metrics are
unchanged by this rendering.

## Publication and freeze

The first non-empty acca report for a league/date writes a decision receipt and
a write-once state lock. Subsequent scans reprint that receipt even if the new
candidate set differs. A missing locked receipt is an error; the engine never
silently rebuilds or rotates a published card. A receipt recovered without its
state lock adopts a lock from the immutable receipt itself.

A withdrawal is a deliberate exception recorded in `auto_tickets_state.json`:
its decision ID is excluded from performance, its date cannot receive a
replacement ticket, and the Git publisher rejects that ID if an old cache
returns it. The original receipt is never rewritten; withdrawals preserve it as
evidence.

Each acca uses exactly one flat paper unit. The initial 100-unit figure is only
an accounting base, not money or a recommended bankroll. Empty scans are logged
but do not cancel earlier positions or consume ticket identities.

An empty latest-scan card includes a **candidate funnel**: current rows, exact
priced rows, calibration-sufficient rows, their overlap, eligible rows/games,
the required independent-game count, and overlapping eligibility blocker counts.
The same structure is stored under `candidate_funnel` in the bounded scan-state
entry and operator result. These are diagnostics only; they never override a
failed guard or synthesize a ticket.

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
