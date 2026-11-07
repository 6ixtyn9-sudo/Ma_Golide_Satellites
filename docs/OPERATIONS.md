# Operations runbook

## Install and verify

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
ruff check .
```

The default suite is hermetic. Live tests are deselected unless explicitly
enabled; do not treat offline fixtures as live coverage or profitability
validation.

## Configuration

Runtime state defaults to:

```text
MAGOLIDE_WAREHOUSE=localdata/warehouse.duckdb
localdata/csv/
localdata/reports/
localdata/ledgers/
localdata/notification_ledger.json
```

Optional provider credentials:

```text
PREDIXSPORT_API_KEY
THE_ODDS_API_KEY
SPORTSGAMEODDS_API_KEY
THERUNDOWN_API_KEY
```

Optional notification credentials and gates:

```text
WHATSAPP_TOKEN
WHATSAPP_PHONE_NUMBER_ID
WHATSAPP_RECIPIENT
ALLOW_NOTIFICATION_WRITES=true
```

Canonical persistence separately requires `ALLOW_MAGOLIDE_WRITES=true` and
`--write`. Never commit credentials or runtime state.

## Daily orchestration

One invocation performs one complete iteration and exits:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
```

Useful options:

- `--timezone Africa/Johannesburg` controls local slate dates (the default).
- `--skip-results` performs pregame/freeze/markets only.
- `--skip-markets` skips the derived forecast report.
- `--acca-legs 3` requests exact three-leg combinations (allowed: 2–6).
- `--acca-limit 5` bounds returned combinations.
- `--freshness-hours 24` controls prediction freshness.
- `--book draftkings` and `--book-provider theoddsapi` restrict book evidence.
- `--localdata PATH` redirects reports/ledgers/exports.
- `--warehouse PATH` overrides the canonical DuckDB path.
- `--keep-days 30` bounds operational runs and daily/market reports.

There is no in-repository daemon. The workflow schedule triggers this command
every three hours; local cron or another scheduler can do the same.

## Hermetic rehearsal

```bash
rm -rf /tmp/magolide-check
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-check/warehouse.duckdb \
  --localdata /tmp/magolide-check --notify
```

Expected properties:

- no credentials or network are needed;
- `verification` is `FIXTURE_VERIFIED`;
- operational and market JSON reports are written;
- no canonical DuckDB or CSV.GZ files are created;
- notification remains a preview and cannot send.

The fixture contains a single pregame provider prediction and no bookmaker
snapshot, so its moneyline forecast is not acca eligible. Empty acca output in
this rehearsal is expected, not a provider outage.

To persist fixture state for parser experimentation, all of the following are
mandatory:

```bash
ALLOW_MAGOLIDE_WRITES=true ALLOW_FIXTURE_WRITES=true \
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture --write \
  --warehouse /tmp/magolide-fixture/warehouse.duckdb \
  --localdata /tmp/magolide-fixture
```

The paths must be isolated from both the configured canonical warehouse and
the repository `localdata/` directory. Fixture state must never be promoted to
production evidence.

## Live dry run and deliberate writes

A live run without `--write` may consume provider quota but does not persist
canonical state:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
```

Persist canonical facts, ledgers, CSV.GZ, and reports only with both opt-ins:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py \
  --auto-once --date 2026-11-07 --league NBA --write
```

The write guard fails closed before acquisition if the environment opt-in is
missing. Dry-run `DuckDBStore(persist=False)` uses memory and never opens the
configured file for modification.

## Standalone market reports

```bash
python scripts/build_markets.py \
  --warehouse localdata/warehouse.duckdb \
  --league NBA --date 2026-11-07 \
  --as-of 2026-11-07T18:00:00Z \
  --tz Africa/Johannesburg \
  --legs 3 --limit 5 \
  --output /tmp/markets.json
```

The command reads an existing warehouse and never acquires provider data or
writes canonical state. Supply `--as-of` for a reproducible evidence cutoff;
without it, the command uses the current UTC time. A current game must tip
after this timestamp. Historical results are available only under the
walk-forward delay.

`--output` chooses the derived JSON report. If omitted, the default is
`localdata/markets_<league>_<date-or-upcoming>.json`. The command refuses a
non-JSON suffix or an output path equal to the warehouse. Stdout contains a
short report summary.

## SQL, coverage, and assay

```bash
python -m ingestion.cli coverage --league NBA --warehouse localdata/warehouse.duckdb
python -m ingestion.cli smoke --league NBA --date 2026-11-07
python scripts/mine_assay.py --league NBA --warehouse localdata/warehouse.duckdb
```

Coverage and mining require an existing warehouse; a missing file is an error.
Smoke is read-only and rejects `--write`.

Example read-only SQL:

```bash
python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/warehouse.duckdb')
print(con.execute('''
    select game_id, market, period, line,
           home_price_decimal, away_price_decimal, draw_price_decimal,
           observed_utc, frozen
    from book_facts order by observed_utc desc limit 20
''').fetchall())
con.close()
PY
```

## Export and recovery

Materialize canonical tables:

```bash
python scripts/build_warehouse.py \
  --warehouse localdata/warehouse.duckdb --csv-dir localdata/csv
```

Recover to a new path:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/recover_warehouse.py \
  --warehouse /tmp/recovered.duckdb --csv-dir localdata/csv
python -m ingestion.cli coverage --league NBA --warehouse /tmp/recovered.duckdb
```

Recovery refuses to overwrite an existing file and rejects restoration into a
nonempty warehouse. Keep exports and immutable ledgers in the desired backup
or artifact system; a GitHub Actions cache alone is not an archive.

## Notifications

Notifications are deduplicated operational summaries. They do not contain bet
slips or certified picks.

A deliberate send requires canonical writes, `--write`, `--notify`,
`--send-notifications`, `ALLOW_MAGOLIDE_WRITES=true`,
`ALLOW_NOTIFICATION_WRITES=true`, valid delivery credentials, and non-fixture
verification. The ledger stores message hashes and delivery timestamps, with
bounded count and age retention. Replaying an identical summary returns
`already_sent` without another delivery.

## Retention and failure handling

- Canonical games, predictions, books, aliases, and health are not deleted by
  operational retention.
- `pipeline_runs` older than the retention cutoff are pruned only on writes.
- Old daily and market reports are pruned on both dry and write runs because
  dry runs also create derived reports.
- Unknown or malformed report files are preserved for inspection.
- Frozen-write conflicts, irreconcilable results, missing identities, and
  ledger conflicts are explicit failures or coverage gaps; do not bypass them
  by deleting evidence.
- Check `provider_health`, stage `gaps`, and `uncovered_games` before assuming
  a market is unsupported.
