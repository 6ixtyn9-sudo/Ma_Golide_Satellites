# Ma Golide operations runbook — headless DuckDB/CSV

## 1. Prerequisites

- Python 3.11+
- DuckDB and Requests for the default runtime
- Provider credentials only for optional authenticated providers you use
- No web server, external database, Google service account, satellite
  spreadsheet, or Apps Script deployment

Install and verify:

```bash
python -m pip install -e '.[dev]'
python -m pytest tests -q
```

Default runtime paths:

```text
MAGOLIDE_WAREHOUSE=localdata/warehouse.duckdb
localdata/csv/
localdata/ledgers/
localdata/reports/
```

`localdata/` is runtime state, not source. Preserve it between runs. Reports
and `pipeline_runs` telemetry are pruned after 30 days; canonical facts and
immutable ledgers are retained.

## 2. Safe checks

### Hermetic full daily cycle

```bash
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-check/warehouse.duckdb \
  --localdata /tmp/magolide-check \
  --run-id manual-dry-check --notify
```

This runs pregame, freeze, results, assay, and notification rendering. It
creates reports under `/tmp/magolide-check` but does not create a DuckDB file,
canonical CSV/ledger state, or a WhatsApp message. Fixture output
says `FIXTURE_VERIFIED`.

If an isolated fixture warehouse is needed for a recovery rehearsal, fixture
persistence additionally requires `ALLOW_FIXTURE_WRITES=true`. Fixture reports
are always refused for external notification.

### Individual commands

```bash
python -m ingestion.cli smoke   --league NBA --date 20261107 --offline-fixture
python -m ingestion.cli pregame --league NBA --date 20261107 --offline-fixture
python -m ingestion.cli freeze  --league NBA --date 20261107
python -m ingestion.cli results --league NBA --date 20261107
python -m ingestion.cli coverage --league NBA
```

Without `--write`, an existing warehouse is opened read-only. A missing
warehouse becomes an empty in-memory dry-run store and is not created on disk.

## 3. Canonical live acquisition

Dry-run live capture first:

```bash
python scripts/daily.py --auto-once --date 2026-11-07
```

Inspect fixtures found, selected providers, actual book coverage, gaps, and
provider health. Then deliberately persist:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py \
  --auto-once --date 2026-11-07 --write
```

`--write` without `ALLOW_MAGOLIDE_WRITES=true` fails before acquisition. The
write path stores canonical facts without any legacy worksheet.

Expected durable outputs:

```text
localdata/warehouse.duckdb
localdata/csv/{games,predictions,book_snapshots,team_aliases,provider_health,pipeline_runs}.csv.gz
localdata/ledgers/{pregame,freeze,results}_<league>_<date>_<hash>.json
localdata/reports/daily_<run-id>.json
localdata/assay_<league>.json
```

Rerun the same command. Canonical natural keys and identical ledgers should be
no-ops. A changed observation may create a new post/pre slot or a new hashed
ledger; it must not rewrite a frozen pregame value.

## 4. SQL research, coverage and CSV recovery

```bash
python -m ingestion.cli coverage --league NBA
python scripts/build_warehouse.py
python scripts/mine_assay.py --league NBA
```

Recover a **new** warehouse from the six canonical CSV.GZ tables (an existing
warehouse is never overwritten):

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/recover_warehouse.py \
  --csv-dir localdata/csv --warehouse /tmp/recovered-warehouse.duckdb
```

Inspect the recovered counts and freeze state before moving it into the active
runtime path. Recovery refuses a non-empty destination and uses a temporary
file plus atomic rename.

Example SQL:

```bash
python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/warehouse.duckdb')
con.sql('''
  SELECT provider, count(*) AS n,
         count(*) FILTER (WHERE frozen) AS frozen_n
  FROM prediction_facts
  WHERE pre_tip
  GROUP BY provider ORDER BY provider
''').show()
con.sql('''
  SELECT game_id, home_final, away_final, result_source
  FROM game_results ORDER BY scheduled_utc DESC LIMIT 20
''').show()
con.close()
PY
```

The assay artifact is research-only. It does not authorize a bet, stake,
certified status, or pick publication.

## 5. Optional WhatsApp heartbeat

Render without sending:

```bash
python scripts/notify.py --report localdata/reports/daily_<run-id>.json
```

Send deliberately after configuring `WHATSAPP_TOKEN`,
`WHATSAPP_PHONE_NUMBER_ID`, and `WHATSAPP_RECIPIENT`:

```bash
ALLOW_NOTIFICATION_WRITES=true python scripts/notify.py \
  --report localdata/reports/daily_<run-id>.json --send
```

The message reports operational coverage, not a betting recommendation. Its
content hash is stored in `localdata/notification_ledger.json`; an unchanged
heartbeat is not sent twice. A delivery failure is reported but does not undo
committed acquisition state.

## 6. GitHub Actions

The workflow runs every three hours and on `workflow_dispatch`:

1. restore `localdata/` from the newest `magolide-localdata-v1-` cache;
2. run the offline suite;
3. execute `scripts/daily.py --auto-once`;
4. save the cache on success;
5. upload run reports and immutable ledgers.

Canonical writes are disabled until repository variable
`ALLOW_MAGOLIDE_WRITES=true` is set. While disabled, scheduled iterations use
committed fixtures so they cannot consume authenticated provider quota without
persisting state; use a manual dry run for deliberate live coverage checks.
WhatsApp delivery requires its own repository variable and credentials. Manual
runs require both the input and repository variable to agree.

Do not commit the database to Git to work around cache management. Restore the
latest successful cache, rerun in dry-run mode, inspect gaps, and only then
resume writes.

## 7. Troubleshooting

### `warehouse does not exist`

Run a canonical write first, or use an isolated dry-run path. Research commands
intentionally fail instead of treating an absent database as an empty
successful slate.

### `FrozenWriteError`

A changed value attempted to replace frozen pre-tip evidence. Keep the original
record, inspect provider identity/timing, and write a separate post-tip shadow
if appropriate. Never clear the warehouse to bypass this guard.

### New ledger filename after a rerun

Canonical content changed. Compare the two JSON files. Identical content reuses
the same path and returns `created: false`.

### Cache was lost but CSV.GZ files survive

Recover to a new path with `scripts/recover_warehouse.py`, verify counts and
freeze state, and only then replace the active runtime file. Never merge
recovery rows into a non-empty warehouse.

### Missing quarter lines

This is a coverage gap, not permission to synthesize a line. Check configured
book/provider capabilities and provider health. Expected totals, model margins,
and blended bundles are not bookmaker lines.

## 8. Legacy Sheets migration only

Install legacy dependencies only for historical migration/compatibility:

```bash
python -m pip install -e '.[legacy]'
python -m ingestion.cli coverage --data-store-id <legacy-store-id>
python -m ingestion.cli pregame --sheet-id <legacy-satellite-id> ...
```

They require Google credentials and, for writes, both
`ALLOW_MAGOLIDE_WRITES=true` and `ALLOW_SHEET_WRITES=true`. They are not part of
the daily orchestrator or scheduled workflow. Do not enable Apps Script writers
against ranges being migrated.
