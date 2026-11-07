# Ma Golide Satellites

Ma Golide is a **headless, DuckDB/CSV.GZ-first sports evidence and forecasting
system**. It acquires schedules, official results, provider predictions, and
bookmaker observations; freezes pre-tip evidence; measures historical market
reliability; and emits current market and accumulator research reports.

There is no web server, spreadsheet store, external database, worksheet
registry, JSON-file datastore, or external betting writer. JSON is used only
for complete row payloads inside DuckDB and for disposable derived reports.

## League and provider coverage

| League | Native regulation periods | Schedule / results | Independent predictions | Books |
| --- | --- | --- | --- | --- |
| NBA | Q1–Q4 | ESPN | ESPN BPI, Hardwood, Predixsport | TheRundown, SportsGameOdds, TheOddsAPI |
| EuroLeague | Q1–Q4 | EuroLeague | Predixsport | No supported EuroLeague book adapter |
| CBA | Q1–Q4 | iSportsAPI | Sportradar Global Basketball | iSportsAPI |
| NCAA_M | H1/H2 | ESPN men's Division-I endpoint | ESPN BPI | TheOddsAPI |
| NCAA_W | H1/H2 | ESPN women's Division-I endpoint | ESPN BPI | TheOddsAPI |

NCAA halves are never relabelled as quarters, and quarter/highest-quarter
products are not generated for NCAA. CBA predictions and books deliberately
come from independent providers; book-implied probabilities are not stored as
model predictions. CBA and NCAA live provider coverage remains
`DOCUMENTATION_ONLY` until authenticated smoke tests verify it. Offline tests
prove parsing, reconciliation, routing, and forecast behavior only.

## Canonical market coverage

| Canonical market | Periods / selections | Evidence rule |
| --- | --- | --- |
| Moneyline | Full game: home or away | Valid provider probabilities; an exact priced bookmaker snapshot is required before acca eligibility |
| Favorite / underdog role | Full-game moneyline | A role on the same moneyline forecast, derived from real home/away prices—not a separate wager model |
| Total | Full game: over or under | The posted bookmaker line is mandatory; a provider expected total is only a model baseline |
| Totals | Q1–Q4 for quarter leagues; H1/H2 for NCAA | Exact native period, exact posted line, and over/under prices |
| Spreads | Q1–Q4 for quarter leagues; H1/H2 for NCAA | Exact native period and home-oriented posted line with both side prices |
| 1X2 | First half: home, draw, or away | Exact first-half book, including a draw price when the draw is selected |
| Highest quarter | Q1, Q2, Q3, or Q4 for quarter leagues only | History-only research forecast; an exact priced highest-quarter book and dead-heat rules are required for acca use |

The forecasting model covers every market family in the former betting
contract without restoring its worksheet-era names or wrappers. Current book
adapters do not yet supply highest-quarter prices, so those forecasts remain
research-only until such evidence exists.

## Quick start

Python 3.11+ is required. Runtime dependencies are DuckDB and Requests; the
development extra adds Pytest and Ruff. `uv` can be installed separately for
locked environment synchronization.

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
ruff check .
```

Run a deterministic, credential-free rehearsal:

```bash
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-check/warehouse.duckdb \
  --localdata /tmp/magolide-check --notify
```

This writes derived reports under `/tmp/magolide-check/reports`, but does not
create a DuckDB file or CSV exports. Its verification label is
`FIXTURE_VERIFIED`; it proves neither live provider coverage nor betting
performance.

A live acquisition attempt remains a canonical dry run unless `--write` is
present:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
python scripts/daily.py --auto-once --date 2026-11-07 --league CBA
python scripts/daily.py --auto-once --date 2026-11-07 --league NCAA_W
```

Live CBA acquisition requires `ISPORTS_API_KEY` for fixtures/results/books and
`SPORTRADAR_API_KEY` for independent probabilities. Set
`SPORTRADAR_ACCESS_LEVEL=production` only for a production Sportradar key.
NCAA book acquisition uses `THE_ODDS_API_KEY`; a missing key disables that
provider rather than inventing book evidence.

Deliberate persistence requires both the CLI flag and environment opt-in:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py \
  --auto-once --date 2026-11-07 --write
```

## Forecast and acca reports

The three-hour orchestrator runs pregame acquisition, current market
generation, immutable pre-tip freezing, and (unless skipped) result
reconciliation.
Generate a report directly from an existing warehouse with:

```bash
python scripts/build_markets.py \
  --league NBA --date 2026-11-07 --as-of 2026-11-07T18:00:00Z \
  --tz Africa/Johannesburg --legs 3 --limit 5 \
  --output /tmp/markets.json
```

For current fixtures, `--as-of` must be before scheduled tip. For historical
backtests, only reconciled finals available at each simulated cutoff are used.
The default result-availability delay is four hours.

The walk-forward harness fits separate reliability profiles by league, market,
period, and model source. Each profile uses only earlier graded outcomes. Its
shrunken observed-minus-claimed adjustment changes later calibrated
probabilities; at least 30 graded samples are required before a segment can
supply acca legs.

An acca is eligible only when every leg has:

- an exact real bookmaker market, period, line where applicable, and selected
  decimal price;
- visible pre-tip book evidence and provider evidence where used, with freeze
  state reported explicitly;
- enough walk-forward samples and a calibrated probability at least 0.50;
- acceptable push risk for totals and spreads.

Accas contain exactly the requested number of legs (two through six) and are
ranked by the product of calibrated leg probabilities. Same-game legs are
prohibited. Cross-game conditional independence is an explicit simplifying
assumption, not a measured guarantee. “Highest” therefore means highest under
that model among eligible combinations—not a guaranteed winner or an
empirically certified global maximum.

Reports have schema `magolide.market_forecasts.v1` and status
`decision_support_only`. They cannot place bets or trigger external betting
writes.

## System of record and read access

- `localdata/warehouse.duckdb`: canonical games, prediction slots, book
  snapshots, aliases, provider health, and operational run records.
- `localdata/csv/*.csv.gz`: six materialized canonical tables for audit and
  recovery.
- `localdata/ledgers/`: immutable, content-addressed stage snapshots.
- `localdata/reports/`: derived operational and market JSON reports.
- `localdata/notification_ledger.json`: bounded notification deduplication
  metadata, not canonical sports evidence.

```bash
python -m ingestion.cli coverage --league NBA
python scripts/mine_assay.py --league NBA
python scripts/build_warehouse.py

python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/warehouse.duckdb')
con.sql('SELECT * FROM latest_pregame_predictions LIMIT 20').show()
con.close()
PY
```

Recover into a **new** path; recovery never overwrites an existing warehouse:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/recover_warehouse.py \
  --warehouse /tmp/recovered.duckdb --csv-dir localdata/csv
```

## Safety boundaries

- Prediction and book evidence are separate; neither is synthesized from the
  other.
- Pre-tip means strictly `observed_utc < scheduled_utc`. Historical training
  requires immutable frozen rows finalized at/after tip; current reports use
  the latest visible pre-tip evidence.
- Dry runs never create or modify canonical DuckDB/CSV state.
- Fixture persistence needs `ALLOW_FIXTURE_WRITES=true`, the normal write
  opt-in, and paths isolated from canonical state.
- Notifications are optional operational heartbeats, not betting advice.
  Delivery has a separate `ALLOW_NOTIFICATION_WRITES=true` gate and fixture
  reports can never send.
- Live-only tests are deselected by default. Offline tests make no network
  calls and no external writes.

## Documentation

- [Architecture](docs/ARCHITECTURE.md)
- [Data and forecasting contract](docs/CONTRACT.md)
- [Operations runbook](docs/OPERATIONS.md)
- [Provider matrix](docs/PROVIDERS.md)
- [Research methodology](docs/RESEARCH.md)
- [Verification and limitations](docs/VERIFICATION.md)
- [Contributor checks](docs/CONTRIBUTING.md)
