# Verification and known limitations

Base verification date: **2026-11-09**. The historical cleanup results below
come from local offline commands in the cleaned checkout. No authenticated
provider validation, production migration, notification delivery, or betting
write was performed.

## EuroLeague Women FIBA follow-up (2026-11-18)

This section records the narrower, later FIBA ELW remediation separately from
the 2026-11-09 baseline. `LIVE_CONFIRMED` means a read-only observation on that
date; it is **not** a service-level, licensing, quota, or ongoing-schema
commitment. `FIXTURE_COVERED` means hermetic parser/control-flow coverage only.

| Category | Status | Evidence / boundary |
| --- | --- | --- |
| GDAP competition schedule | `LIVE_CONFIRMED` | Competition `209127` returned 110 ELW rows; 54 future bracket rows had `teamA`/`teamB` null. The adapter skips/report-aggregates those unassignable slots so they cannot abort a valid Paris-day slate or become games. |
| Canonical schedule/final shape | `LIVE_CONFIRMED` | `gameDateTimeUTC` was offset-less UTC; completed rows used `gameResultStatusCode=VALID`, `liveGameStatus=999`, and numeric scores. Future rows also used `liveGameStatus=1`, so no live-state inference relies on that value. |
| Game-page extraction | `LIVE_CONFIRMED` | Read-only RSC probes of three completed regular-season games produced game-bound PBP and reconciled quarter/OT detail; a future page had no usable PBP and failed closed. |
| Team-less slots, terminal statuses, PBP decoys, cursor retention, correction trigger | `FIXTURE_COVERED` | Tests cover a mixed/all-placeholder schedule, CANCELLED/CANCELED/ABANDONED/VOID fail-open gate handling, wrong/missing PBP game binding, all-failed FIBA history detail batch retaining its cursor, and a changed GDAP score forcing re-verification. |
| Cancelled/postponed/abandoned live vocabulary, multi-OT, future playoff slug | `UNRESOLVED` | No real row/page for these cases was observed. Synthetic cases only protect the documented conservative behavior. |
| TeamA home semantics across neutral-site rounds | `UNRESOLVED` | FIBA schedule-side A is the adapter orientation, not independently validated home/host evidence. |
| Automated use, quota, key rotation, storage/redistribution rights | `UNRESOLVED` | A public browser-delivered key and successful read-only calls do not establish these permissions or operational guarantees. |

The full live-vs-fixture distinction and current operational constraints are in
[PROVIDERS.md](PROVIDERS.md). ELW has no dedicated API bookmaker integration,
but its canonical future fixtures are eligible for the bounded direct
BetExplorer named-book fallback. That fallback is full-game moneyline-only and
is explicitly labelled `unresolved` / `probation`; absent a matching displayed
comparison row, the empirical card remains unpriced rather than making an
EV/value claim.

## Clean dependency and test verification

The lockfile was regenerated with `uv lock`. Removed spreadsheet
authentication/client packages are absent from both `pyproject.toml` and
`uv.lock`.

The test virtual environment was recreated and the project installed with its
development dependencies:

```bash
python -m venv /tmp/magolide-test-venv
/tmp/magolide-test-venv/bin/pip install -e '.[dev]'
```

The resulting runtime includes the project, DuckDB, Requests, Pytest, and
Ruff. The complete offline suite and static checks then passed:

```text
$ /tmp/magolide-test-venv/bin/python -m pytest -q
242 passed, 4 deselected in 4.78s

$ /tmp/magolide-test-venv/bin/ruff check .
All checks passed!

$ git diff --check
(no output)
```

The four deselected tests are explicit live smoke tests. They were not run and
are not included in any live-coverage claim.

## Automated behavior covered

### Acquisition and official results

- ESPN and EuroLeague schedule/result fixtures, team orientation, quarter and
  combined-overtime reconciliation, and invalid/partial result quarantine;
- WNBA discovery, BPI, scoreboard books, summaries, native current quarters,
  isolated DuckDB storage, workflow routing, and operator publication;
- timestamped ESPN books independent of predictions, summary context,
  redacted request telemetry, retry/cooldown behavior, request budgets, and
  durable public-cache hits retaining original observation timestamps;
- core season-index discovery, durable deep-history cursors, the 1947 NBA
  final-score-only case, early WNBA native halves, DuckDB cursor reopen, and
  proof that dry runs cannot advance unpersisted history;
- strict timezone-aware timestamps and chronological selection;
- finite/nonnegative prediction expectations and finite book lines/decimal
  prices, rejecting NaN, Infinity, and invalid odds;
- provider-ID joins, conflicting event-map rejection, and post-tip shadow
  evidence;
- strict `observed < tip` boundaries for predictions and books;
- preservation of official finals during metadata-only schedule rediscovery;
- provider fallback without averaging probabilities;
- exact full-game/H1/Q1–Q4 book filtering and draw-price parsing.

### Forecasting and accas

- Full-game moneyline and favorite/underdog price roles;
- independent provider and empirical-history full-game surfaces, including a
  no-provider-archive regression that earns calibration from prior finals and
  admits exactly priced current games without relaxing the sample gate;
- full-game totals and Q1–Q4 totals/spreads against real posted lines;
- first-half home/draw/away forecasts and selected-draw price requirements;
- highest-quarter forecasts and tied-quarter research grading;
- proof that an expected total never becomes a total wager line;
- proof that a Q1 line is not reused for another period;
- chronological residual construction and result-availability cutoffs;
- exclusion of post-cutoff and unfrozen historical evidence;
- walk-forward segment reliability adjustment and prequential metrics;
- exact-tie suppression, push grading, and highest-quarter tie handling;
- real-price/sample/push eligibility checks;
- acca ranking by calibrated joint probability, exact brute-force comparison
  on a small case, and unique game IDs within each acca;
- paper-ticket portfolios do not repeat exact wager selections, while allowing a
  different selection from the same game in another ticket.

The synthetic multi-market service test generates every canonical market
family and returns ranked accas once enough synthetic walk-forward samples and
real test book lines exist. This validates mechanics, not real-world accuracy.

### Persistence and orchestration

- DuckDB canonical round trips and immutable pre/post slots;
- first-half draw column schema migration and backfill from complete row
  payloads;
- six-table CSV.GZ export and recovery;
- dry-run file immutability and guarded fixture isolation;
- canonical ledgers, conflict/no-op behavior, and operational retention;
- daily market-report integration, normalized fixture timestamps, and a
  pinned 18:00 UTC offline forecast cutoff independent of wall-clock time;
- bounded notification deduplication/retention and fixture-send rejection;
- standalone market CLI output with a before/after hash proving the source
  DuckDB file was unchanged.

## End-to-end fixture rehearsal

An isolated dry run used:

```bash
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-final-rehearsal/dry/warehouse.duckdb \
  --localdata /tmp/magolide-final-rehearsal/dry --notify
```

Observed and asserted:

- `FIXTURE_VERIFIED` operational and market reports;
- one current full-game moneyline forecast from the fixture prediction;
- no book quote and no acca;
- notification preview only, with no delivery;
- no DuckDB file and no CSV.GZ canonical exports.

A separate isolated persistence rehearsal used both write guards. It produced:

```text
games: 1
prediction slots: 1
book snapshots: 0
pipeline run records: 3
CSV.GZ tables: 6
ledger stages: pregame, freeze, results
```

The fixture scoreboard has no official final result, so this particular replay
has zero graded forecasts. That is correct missing-evidence behavior.

The standalone CLI then ran read-only:

```bash
python scripts/build_markets.py \
  --warehouse /tmp/magolide-final-rehearsal/persisted/warehouse.duckdb \
  --league NBA --date 2026-11-07 --as-of 2026-11-08T00:00:00Z \
  --output /tmp/magolide-final-rehearsal/markets.json
```

Finally, all six CSV.GZ tables were restored into a new warehouse. Recovery
reported one game, one prediction, zero books/aliases/health rows, and three
pipeline runs. The read-only coverage command found the recovered game and
prediction and correctly reported zero reconciled finals.

## Remaining limitations

- Public ESPN scoreboard, summary, predictor, and historical index probes were
  performed read-only; authenticated CBA/book-provider coverage remains
  unverified. Documentation/fixtures do not prove current quota, availability,
  or schema stability.
- The historical index floors are observed nonempty/empty query boundaries,
  not proof of exhaustive coverage. See `docs/ESPN-ACQUISITION.md`.
- The installed external-trigger GitHub Actions workflow still requires
  post-change run inspection for WNBA cache restoration, deep cursor creation,
  and ledgers.
- Highest-quarter book prices are not supplied by a current adapter, so HQ
  forecasts cannot enter accas.
- EuroLeague bookmaker coverage currently uses BetExplorer's representative
  average full-game moneyline when quotes are posted. Coverage is sparse, source
  reuse terms remain unreviewed, and no spread/total book family is enabled;
  full-game spread observations can be stored but are not a current full-game
  spread forecast family.
- Derived markets require 20 baseline history games plus 20 eligible residual
  simulations by default; strict acca eligibility separately requires 30 graded
  non-push samples. The paper-only exploratory path may bypass only that sample
  gate after one graded non-push sample; early warehouses may still produce no
  accas when any other guard or the independent-game count is unmet.
- Empirical residual probabilities and shrunk reliability adjustments are not
  proof of profitability or optimal calibration. Prospective performance is
  unknown.
- Same-game acca legs are prohibited. Cross-game conditional independence is
  still an assumption and latent dependence is not estimated.
- Book prices do not model execution limits, slippage, void/dead-heat rules,
  or market removal.
- GitHub Actions caches support orchestration but are not an archival backup;
  retain CSV.GZ and immutable ledger artifacts separately as needed.
