# Ma Golide — headless DuckDB operations and research

## Overview

Ma Golide is a CSV + DuckDB-first basketball acquisition and research system,
following Edge Factory's local-data operating patterns. It runs through CLI
commands and GitHub Actions. The autonomous orchestrator owns all canonical
writes.

There is no web server, external database, or read-model synchronizer. Google
Sheets is legacy migration/compatibility only.

## Architecture

```text
Official/keyless fixtures + configured prediction/book providers
    ↓
Canonical identity, prediction, book and result adapters
    ↓
localdata/warehouse.duckdb (system of record)
    ↓
CSV.GZ + immutable content-addressed ledgers
    ↓
Research-only frozen-prediction assay
    ↓
DuckDB SQL / coverage CLI / optional deduplicated WhatsApp heartbeat

scripts/daily.py → sole unattended write orchestrator
```

## Key files

- `ingestion/duckdb_store.py` — durable canonical tables and freeze semantics.
- `ingestion/warehouse.py` — research SQL views.
- `ingestion/ledger.py` — immutable evidence.
- `ingestion/mining.py` — research-only assay.
- `scripts/daily.py` — one autonomous entry point.
- `scripts/build_warehouse.py` — canonical CSV.GZ materialization.
- `scripts/recover_warehouse.py` — guarded CSV.GZ recovery.
- `scripts/mine_assay.py` — standalone research assay.
- `scripts/notify.py` — operational heartbeat.

The existing `assayer/` engines remain reusable domain logic. The current assay
uses only frozen pre-tip moneyline predictions and reconciled finals. It does
not certify official picks without walk-forward, odds-aware ROI, and CLV gates.

## Runtime configuration

- `MAGOLIDE_WAREHOUSE` — defaults to `localdata/warehouse.duckdb`.
- `ALLOW_MAGOLIDE_WRITES=true` — canonical write opt-in.
- `ALLOW_FIXTURE_WRITES=true` — separate isolated fixture-write opt-in.
- `WHATSAPP_TOKEN` / `WHATSAPP_PHONE_NUMBER_ID` / `WHATSAPP_RECIPIENT` — optional.
- `ALLOW_NOTIFICATION_WRITES=true` — notification send opt-in.

Optional provider keys are documented in `docs/INGESTION_RUNBOOK.md`. The
default runtime needs only DuckDB and Requests. Legacy Google packages are
isolated in the `legacy` extra.

## Headless read access

```bash
python -m ingestion.cli coverage --league NBA
python scripts/mine_assay.py --league NBA
python scripts/build_warehouse.py
```

For detailed facts, open `localdata/warehouse.duckdb` through
`ingestion.warehouse.connect()` and query `prediction_facts`, `book_facts`,
`latest_pregame_predictions`, or `game_results`.

## Common commands

```bash
# Offline suite
python -m pytest tests -q

# Hermetic daily dry run
python scripts/daily.py --auto-once --offline-fixture

# Persist canonical state deliberately
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py --auto-once --write
```

The Replit run button executes the isolated fixture dry run, not a web service.
See `docs/DUCKDB_ARCHITECTURE.md` and `docs/INGESTION_RUNBOOK.md` for full
operations and migration boundaries.
