# Ma Golide Satellites

Ma Golide is a **headless CSV + DuckDB-first sports data and research system**,
using Edge Factory's local-data operating patterns. Google Sheets is no longer
the system of record or a requirement of acquisition or scheduling. There is no
web service, external database, or synchronization service.

## Architecture

```text
ESPN / EuroLeague / Hardwood / configured book providers
                         │
                         ▼
             canonical provider adapters
                         │
                         ▼
          localdata/warehouse.duckdb  ◀── system of record
                         │
             ┌───────────┴────────────┐
             ▼                        ▼
   localdata/csv/*.csv.gz   content-addressed immutable ledgers
             │                        │
             └───────────┬────────────┘
                         ▼
          frozen-prediction research assay
             (research-only, not auto-certified)
                         │
             ┌───────────┴────────────┐
             ▼                        ▼
     DuckDB SQL / coverage CLI   optional deduplicated WhatsApp heartbeat
```

The scheduled entry point is one orchestrator:

```bash
python scripts/daily.py --auto-once --date 2026-11-07
```

GitHub Actions restores and saves `localdata/` as a cache. Canonical writes and
WhatsApp delivery have separate opt-ins.

## Quick start

Python 3.11+ is required. The default runtime needs only DuckDB and Requests;
the development extra also installs test and legacy-adapter test dependencies.

```bash
python -m pip install -e '.[dev]'
python -m pytest tests -q

# Hermetic end-to-end preview: committed fixtures, no credentials, no DB file.
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-check/warehouse.duckdb \
  --localdata /tmp/magolide-check --notify

# Live capture, still a dry run for canonical state.
python scripts/daily.py --auto-once --date 2026-11-07

# Deliberately persist DuckDB, CSV.GZ and immutable ledgers.
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py \
  --auto-once --date 2026-11-07 --write
```

A dry run never creates or changes the DuckDB file. The fixture path is always
labelled `FIXTURE_VERIFIED`; it is not live provider coverage. Fixture
persistence requires a separate `ALLOW_FIXTURE_WRITES=true` and an isolated
warehouse. Fixture reports cannot trigger external notifications.

## Storage and read access

- `localdata/warehouse.duckdb`: canonical games, prediction slots, book
  snapshots, team aliases, provider health, and per-stage run reports.
- `localdata/csv/*.csv.gz`: materialized canonical tables for SQL research and
  audit. Materialize with `python scripts/build_warehouse.py`; recover a new
  warehouse with `scripts/recover_warehouse.py`.
- `localdata/ledgers/`: immutable content-addressed pregame, freeze, and result
  snapshots. Identical reruns are no-ops; changed content gets a new hash.
- `localdata/reports/`: one operational report per orchestrator invocation.
- `localdata/notification_ledger.json`: hashes of sent operational heartbeats,
  used to suppress duplicates.

Inspect state without a web application:

```bash
python -m ingestion.cli coverage --league NBA
python scripts/mine_assay.py --league NBA

python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/warehouse.duckdb')
con.sql('SELECT * FROM latest_pregame_predictions LIMIT 20').show()
con.close()
PY
```

Runtime state is ignored by Git. See `localdata/README.md` for the cache
boundary and never store credentials there.

## Legacy Google Sheets boundary

The old Sheets adapters and non-destructive writer remain only for explicit
migration/compatibility work. Install their packages only when needed:

```bash
python -m pip install -e '.[legacy]'
```

Legacy storage is never selected from `DATA_STORE_SPREADSHEET_ID` implicitly. A
legacy operation requires an explicit `--data-store-id` or `--sheet-id`, plus
the legacy write guard when applicable. The default runtime has no satellite
registry, worksheet, Apps Script, Google service-account, or web dependency.

## Documentation

- `docs/DATA_ACQUISITION_BUILD_PROMPT.md` — authoritative architecture/build
  contract for future agents.
- `docs/DUCKDB_ARCHITECTURE.md` — Edge Factory pattern mapped to Ma Golide.
- `docs/INGESTION_RUNBOOK.md` — operations, gates, cache recovery, and queries.
- `docs/ACQUISITION_FINAL_REPORT.md` — implemented scope and verification.
- `ingestion/PROVIDER_GUIDE.md` — provider identity and market constraints.

No live production migration, WhatsApp send, or authenticated provider
validation is implied by the offline tests.
