# Verification and known limitations

Verification date: **2026-11-07**. All results below come from local offline
commands in the cleaned checkout. No authenticated provider validation,
production migration, notification delivery, or betting write was performed.

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
  on a small case, and unique game IDs across every leg.

The synthetic multi-market service test generates every canonical market
family and returns ranked accas once enough synthetic walk-forward samples and
real test book lines exist. This validates mechanics, not real-world accuracy.

### Persistence and orchestration

- DuckDB canonical round trips and immutable pre/post slots;
- schema-v2 first-half draw column migration and backfill from complete row
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
- The scheduled GitHub Actions workflow still requires post-change run
  inspection for WNBA cache restoration, deep cursor creation, and ledgers.
- Highest-quarter book prices are not supplied by a current adapter, so HQ
  forecasts cannot enter accas.
- EuroLeague book coverage remains unenabled; full-game spread observations
  can be stored but are not a current full-game spread forecast family.
- Derived markets require 20 baseline history games plus 20 eligible residual
  simulations by default; acca segments separately require 30 graded samples.
  Early warehouses will correctly produce few or no eligible accas.
- Empirical residual probabilities and shrunk reliability adjustments are not
  proof of profitability or optimal calibration. Prospective performance is
  unknown.
- Same-game acca legs are prohibited. Cross-game conditional independence is
  still an assumption and latent dependence is not estimated.
- Book prices do not model execution limits, slippage, void/dead-heat rules,
  or market removal.
- GitHub Actions caches support orchestration but are not an archival backup;
  retain CSV.GZ and immutable ledger artifacts separately as needed.
