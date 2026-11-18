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
`--write`. The scheduled workflow defaults the variable to true; manual
workflow dispatch remains a dry run. Never commit credentials or runtime state.

## Daily orchestration

One invocation performs one complete iteration and exits:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
```

Useful options:

- `--timezone Africa/Johannesburg` controls local reporting dates. League
  defaults are America/New_York (NBA/WNBA/NCAA), Europe/Paris (EuroLeague), and
  Asia/Shanghai (CBA). CBA schedule acquisition itself uses iSportsAPI's
  documented UTC calendar window.
- An omitted `--date` is derived in the selected league's acquisition timezone.
- `--skip-results` performs pregame/freeze/markets only.
- `--skip-markets` skips the derived forecast report.
- `--acca-legs 3` requests exact three-leg combinations (allowed: 2–6).
- `--acca-limit 5` bounds returned combinations.
- `--freshness-hours 24` controls prediction freshness.
- `--history-days 21` and `--history-requests 21` bound recent daily backfill.
- `--deep-history-events 50` bounds indexed historical summaries per
  write-enabled league run; dry runs never advance that cursor.
- `--book draftkings` and `--book-provider theoddsapi` restrict book evidence.
- `--localdata PATH` redirects reports/ledgers/exports.
- `--warehouse PATH` overrides the canonical DuckDB path.
- `--keep-days 30` bounds operational runs and daily/market reports.

There is no in-repository daemon. The workflow schedule triggers this command
hourly; local cron or another scheduler should run at least hourly as well. The
cadence gives each slate a pre-tip opportunity to make its 60-minute ticket
decision without letting the first partial-price scan freeze the date.

## Automatic six-league workflow

`workflows/ingestion.yml` is the reviewed workflow source. Install it as
`.github/workflows/ingestion.yml` with a GitHub credential that has the
`workflows` permission; merely editing the local source does not activate it.
It has an hourly schedule (`11 * * * *`, UTC) and a manual dispatch with **no
league/date/mode/write inputs**. A preflight job runs Ruff and the hermetic
suite, then six fail-independent matrix jobs automatically run `NBA`, `WNBA`,
`EuroLeague`, `CBA`, `NCAA_M`, and `NCAA_W` together. One serialized publisher
then commits compact ticket decisions and state to `main`; all pick cards and
expanded audits stay in the cache and bounded workflow artifacts so hourly price
movement does not create permanent Git history.

Each matrix job has isolated state and artifacts:

- `localdata/nba`, `localdata/wnba`, `localdata/euroleague`, `localdata/cba`,
  `localdata/ncaa_m`, and `localdata/ncaa_w`;
- a league-specific DuckDB path, cache prefix, run ID, and artifact name;
- an acquisition date derived automatically in that league's acquisition
  timezone, never a dispatch-selected date.

### Paper-ticket decision cutoff

A current scan is candidate evidence, not a ticket publication. For each
league/date, the operator uses the **acquisition timezone** to find the first
slate tip and waits until 60 minutes before that tip. The first successful scan
in that pre-tip window freezes all currently valid, exact-priced paper
candidates into one immutable receipt. Every leg must still be strictly before
its own tip. The actual receipt records both the planned cutoff and the scan
that made the decision.

If no qualifying pre-tip candidate exists at the cutoff, the operator writes an
immutable `auto_tickets_YYYY-MM-DD_no_bet.json` decision instead. It prevents a
later price, an old Actions cache, or a later-game candidate from reopening the
date. A delayed run that misses the first tip also fails closed to that no-bet
decision. Before the cutoff, the text card explicitly identifies the pending
decision time and later scans may update candidate prices; after either decision,
selections/prices/stakes—or the no-bet result—cannot change.

### Optional odds-provider credentials

The public feeds remain useful without credentials, but they do **not** provide
uniform odds coverage: ESPN's public scoreboard is full-game-only when it
includes a book at all, BetExplorer is full-game moneyline-only, and NCAA_W
moneyline coverage is known to be absent from the ESPN free surface. To enable
the already-implemented supplemental adapters in GitHub Actions, set the
corresponding **Actions secrets** (never repository variables or checked-in
files):

| Secret | Implemented adapter coverage (verify live) | Use it for |
| --- | --- | --- |
| `THE_ODDS_API_KEY` | NBA/WNBA/NCAA_M/NCAA_W full-game markets; NBA quarters and H1; NCAA H1/H2 where the provider offers them | The first supplemental source to validate, especially for NCAA_W |
| `THERUNDOWN_API_KEY` | NBA full-game moneyline and totals | A second named NBA source |
| `SPORTSGAMEODDS_API_KEY` | NBA full-game and Q1–Q4 totals | NBA quarter-total coverage |
| `ISPORTS_API_KEY` | CBA moneyline, totals, spreads, halves, and quarters | CBA market coverage; adapter remains documentation-only until an authenticated smoke succeeds |
| `SPORTRADAR_API_KEY` / `PREDIXSPORT_API_KEY` | Independent CBA / EuroLeague prediction paths | Prediction redundancy, not bookmaker prices |

The workflow injects these only as process environment variables. A missing
secret is safe: its adapter remains disabled and the pregame report records
`book_provider_status: disabled` rather than presenting that state as an empty
market response. The GitHub job summary also includes selected-side price
coverage by market/period and the redacted source status, so a no-price result
can be separated into a provider-capability gap, a disabled credential, an
identity rejection, or a true empty response.

Safety behavior is deliberate:

- Manual dispatch is a live all-league canonical **dry run**. It may consume
  provider quota but never adds `--write` and never sends notifications, even
  when `ALLOW_MAGOLIDE_WRITES` is true.
- Scheduled runs default to `--write`: the workflow falls back to
  `ALLOW_MAGOLIDE_WRITES=true` when the repository variable is unset. Set the
  variable to `false` to disable canonical persistence; runs still acquire
  available providers and report missing credentials.
- Scheduled runs save league state. When canonical writes are disabled, this
  preserves public acquisition caches and receipts but does not persist
  DuckDB/CSV sports state or advance deep-history cursors. Reports and
  immutable ledgers are uploaded on every outcome when present. Git is
  authoritative for operator files after cache recovery.
- Manual runs upload derived `picks_*` and `auto_tickets_*` files only as
  inspection artifacts. **Scheduled runs alone** may publish those files to
  `localdata/<league>` on `main`, so a manual rehearsal cannot create a durable
  ticket or no-bet decision. Git publication never includes warehouses, CSV
  exports, acquisition ledgers, or credentials.
- Notifications additionally require `ALLOW_NOTIFICATION_WRITES=true` and the
  WhatsApp secrets; they are sent only on write-enabled scheduled runs.

Every ESPN league run also attempts seven recent result dates and, when
canonical writes are enabled, 50 deep indexed events. Deep cursors, historical
coverage classes, verified index floors, and cache policy are documented in
[ESPN acquisition and history](ESPN-ACQUISITION.md). Historical rows alone do
not satisfy pre-tip prediction/book requirements or establish calibration.

Configure provider secrets `PREDIXSPORT_API_KEY`, `THERUNDOWN_API_KEY`,
`SPORTSGAMEODDS_API_KEY`, `THE_ODDS_API_KEY`, `ISPORTS_API_KEY`, and
`SPORTRADAR_API_KEY` as needed. `SPORTRADAR_ACCESS_LEVEL` defaults to `trial`
and may be set to `production` as a repository variable for a production key.
Missing credentials disable only the affected provider and remain visible as
coverage gaps; they do not fabricate substitute evidence. CBA/NCAA provider
coverage is `DOCUMENTATION_ONLY` until authenticated smoke runs succeed.

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
short report summary, including ROI-qualified singles and the effective paper
combination leg count. For paper ROI-band eligibility, it reads the settled
single-pick audit from `--localdata` (or `LOCALDATA`; otherwise
`localdata/<league-lowercase>`) when present; it never reads acca performance as
pick history and does not write the audit.

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
