# Headless DuckDB architecture — Edge Factory patterns applied to Ma Golide

## Direction

The earlier acquisition implementation treated a dedicated Google Sheets
spreadsheet as the durable store and satellite tabs as operational outputs.
That conflicted with the intended product direction. On 2026-11-07 the
`6ixtyn9-sudo/Edge-Factory` repository was inspected directly. Ma Golide adopts
its local-data patterns:

1. Capture source rows into durable local artifacts.
2. Use CSV + DuckDB for ingest, mining, and backtests.
3. Run one resumable daily orchestrator.
4. Persist `localdata/` between GitHub Actions runs.
5. Keep dated freeze evidence and replayable reports.
6. Notify only after the operational record exists.

Ma Golide deliberately omits external read-model databases and web frameworks.
DuckDB + CSV.GZ is both the operational store and the read/research surface.
Google Sheets remains explicit legacy migration/compatibility only.

## Pattern mapping

| Edge Factory pattern | Ma Golide implementation |
|---|---|
| `localdata/` source state | `localdata/warehouse.duckdb`, `csv/`, `ledgers/`, `reports/` |
| `scripts/daily.py --auto-once` | `scripts/daily.py --auto-once` |
| DuckDB analytics connection | `ingestion.warehouse.connect()` with temporary research views |
| CSV.GZ materialization | `DuckDBStore.export_csv()` / `scripts/build_warehouse.py` |
| CSV → warehouse recovery | `DuckDBStore.restore_csv()` / `scripts/recover_warehouse.py` |
| Immutable dated evidence | SHA-256-addressed pregame/freeze/result JSON ledgers |
| Mining/assay | `ingestion.mining` + `scripts/mine_assay.py` |
| Operational read access | `ingestion.cli coverage` and direct DuckDB SQL |
| Smart notification | `scripts/notify.py`; operational heartbeat with content-hash dedupe |
| Three-hour Actions loop | `.github/workflows/ingestion.yml` with cache restore/save |

## Authority and data flow

```text
provider response
  -> adapter validation / canonical identity
  -> DuckDB natural-key upsert
  -> immutable pregame snapshot
  -> pre-tip freeze
  -> official result + reconciliation
  -> immutable result snapshot
  -> research-only assay
  -> CSV.GZ materialization
  -> SQL / coverage CLI / optional deduplicated WhatsApp heartbeat
```

DuckDB and its CSV materializations are authoritative and directly readable.
There is no derived database to synchronize and no web write path.

## Canonical DuckDB tables

| Table | Natural identity | Purpose |
|---|---|---|
| `games` | `game_id` | Fixture identity, provider IDs, schedule/status, complete actual result JSON |
| `predictions` | `(game_id, slot)` | Whole provider bundle in `provider|pre/post` slots; freeze flag |
| `book_snapshots` | `(game_id, slot)` | Provider/book/market/period and pre/post observation slot |
| `team_aliases` | `canonical` | Canonical team plus known variants |
| `provider_health` | `(provider, league)` | Last attempt/success, state, coverage, quota metadata |
| `pipeline_runs` | `(run_id, command)` | Redacted report for each orchestration stage |

Book `slot` is
`provider|book|market|period|pre/post`. Prediction `slot` is
`provider|pre/post`. Pre-tip frozen predictions and pre-match frozen books
reject changed overwrites and accept identical replays. Post-tip/live shadow
records cannot displace pregame evidence.

Variable complete payloads remain JSON. Queryable fixture, book, timing,
probability, score, and freeze fields are typed columns. The
`prediction_facts`, `book_facts`, `latest_pregame_predictions`, and
`game_results` temporary views expose these values to SQL without granting
research code a write path.

## CSV recovery and retention

All six tables materialize as CSV.GZ through a temporary file and atomic
replace. `DuckDBStore.restore_csv()` deserializes complete payloads back into
domain records, preserving nested results and freeze flags. Recovery requires
an empty destination and never merges into an active warehouse.

Operational reports and `pipeline_runs` telemetry older than 30 days are
pruned. Canonical facts and immutable ledgers are retained. Dated ledgers are
scoped to their league-local slate, preventing every three-hour run from
copying all league history into another artifact.

## Immutability and reruns

Ledgers contain canonical facts, not a mutable pointer to the latest state.
Their SHA-256 is part of the filename:

```text
pregame_NBA_2026-11-07_<hash>.json
freeze_NBA_2026-11-07_<hash>.json
results_NBA_2026-11-07_<hash>.json
```

An identical rerun returns `created: false`. Changed canonical content produces
a new ledger instead of rewriting history. Direct writes to an existing
immutable path with different bytes raise `LedgerConflictError`.

## Assay boundary

The current warehouse miner evaluates only:

- frozen pre-tip predictions;
- explicit, non-tied home/away probabilities;
- official final results whose quarter + overtime sums reconcile;
- non-tied finals.

Post-tip shadows, unfrozen predictions, absent probabilities, tied games, and
unreconciled scores are excluded. Expected totals are **not** transformed into
Over/Under wagers.

Output retains the existing segment, shrinkage, Wilson lower-bound, grade, and
lift calculations, but every edge is `research_only`, `certified: false`, with
`certified_picks: 0`. Walk-forward validation, odds-aware ROI, CLV, and an
explicit activation decision are required before an official pick ledger can be
introduced. This avoids importing Edge Factory's certification *labels* without
its evidence.

## Headless read boundary

Use `python -m ingestion.cli coverage` for operational coverage,
`scripts/mine_assay.py` for research output, and `ingestion.warehouse.connect()`
for SQL. These interfaces read the canonical store directly and do not require
a web framework, browser application, or another database.

## Actions state and gates

The three-hour workflow uses a concurrency group, restores the latest
`magolide-localdata-v1-` cache, runs hermetic tests, executes one orchestrator
iteration, saves state on success, and uploads reports/ledgers. Operational run
telemetry older than 30 days is pruned; immutable ledgers and canonical facts
are retained. Scheduled iterations use fixtures until canonical writes are
enabled, preventing uncommitted provider-quota consumption.

- `ALLOW_MAGOLIDE_WRITES=true` enables canonical persistence.
- `ALLOW_NOTIFICATION_WRITES=true` plus WhatsApp credentials enables delivery.
- `ALLOW_FIXTURE_WRITES=true` is separately required for isolated fixture
  persistence; fixture reports cannot trigger external delivery.
- Manual inputs must agree with the corresponding repository variable.

A missing optional provider key disables only that provider. Coverage remains
explicit; it is never inferred from a successful HTTP response alone.

## Legacy migration rule

Old Sheets code is retained for reading/migrating historical evidence and for
explicit, guarded compatibility exports. Its packages are isolated in the
`legacy` dependency extra. It is not a fallback store. If DuckDB is unavailable,
the correct behavior is a visible error—not silently selecting Sheets.
Likewise, `DATA_STORE_SPREADSHEET_ID` alone no longer changes the CLI backend;
`--data-store-id` is required to enter the legacy path.
