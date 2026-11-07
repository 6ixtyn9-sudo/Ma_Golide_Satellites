# Deliverables — headless DuckDB/CSV architecture

## Primary implementation

- `ingestion/duckdb_store.py` — canonical DuckDB store, natural keys, freeze
  durability, dry-run isolation, CSV.GZ export/recovery, run reports.
- `ingestion/warehouse.py` — independent research connections and SQL views.
- `ingestion/ledger.py` — immutable slate-scoped content-addressed
  pregame/freeze/result ledgers.
- `ingestion/mining.py` — research-only frozen-prediction assay.
- `scripts/daily.py` — single resumable `--auto-once` orchestrator.
- `scripts/build_warehouse.py` — canonical CSV.GZ materialization.
- `scripts/recover_warehouse.py` — guarded CSV.GZ → new DuckDB recovery.
- `scripts/mine_assay.py` — standalone research assay command.
- `scripts/notify.py` — deduplicated operational WhatsApp heartbeat.
- `.github/workflows/ingestion.yml` — three-hour cached Actions service.

## Headless read access

- `python -m ingestion.cli coverage` — actual canonical coverage.
- `ingestion.warehouse.connect()` — direct SQL research views.
- `scripts/mine_assay.py` — derived research-only results.

There is no dashboard/API, web service, external read-model database, or
synchronization dependency. The default runtime needs
only DuckDB and Requests. Legacy Google packages are optional.

## Documentation

- `README.md` — architecture and quick start.
- `docs/DATA_ACQUISITION_BUILD_PROMPT.md` — authoritative future-agent build
  contract, superseding the Sheets-centered prompt.
- `docs/DUCKDB_ARCHITECTURE.md` — Edge Factory pattern mapping and authority
  boundaries.
- `docs/INGESTION_RUNBOOK.md` — dry runs, writes, SQL, recovery, cache,
  notifications, troubleshooting, and legacy migration.
- `docs/ACQUISITION_FINAL_REPORT.md` — implementation and verification report.
- `localdata/README.md` — persisted runtime/cache contract.

## Tests

Added focused coverage in:

- `tests/test_duckdb_store.py`
- `tests/test_daily_duckdb.py`
- `tests/test_mining.py`
- `tests/test_notify.py`

Existing legacy CLI tests were updated to reflect the intentional boundary:
Sheets requires an explicit legacy flag, while canonical DuckDB writes no
longer require a sheet ID.

## Verification

**225 passed, 4 deselected.** Critical Ruff checks passed. The isolated offline
daily fixture cycle and immediate rerun verify persistence and immutable-ledger
idempotency. CSV recovery reproduces warehouse counts. No live external writes
are claimed.
