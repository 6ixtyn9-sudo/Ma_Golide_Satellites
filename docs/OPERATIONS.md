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
`--write`. The named external-trigger workflow defaults the variable to true;
manual workflow dispatch remains a dry run. Never commit credentials or runtime
state.

## Daily orchestration

One invocation performs one complete iteration and exits:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
```

Useful options:

- `--timezone Africa/Johannesburg` controls local reporting dates. An omitted
  timezone comes from the selected canonical league registry (for example
  America/New_York for NBA/WNBA/NCAA, Europe/Paris for EuroLeague/FIBA ELW,
  Asia/Shanghai for CBA, and each expansion league's own competition region).
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

There is no in-repository daemon and no GitHub Actions cron fallback. An
external scheduler must invoke the workflow at least every 30 minutes. That
cadence gives each slate a pre-tip opportunity to make its 60-minute ticket
decision without letting the first partial-price scan freeze the date.

## Automatic registry-driven basketball workflow

`workflows/ingestion.yml` is the reviewed workflow source. Install its **entire
file** unchanged as `.github/workflows/ingestion.yml` with a GitHub credential
that has the `workflows` permission; merely editing the local source does not
activate it. The hermetic suite requires the two files to be byte-identical, so
manual installation must replace the full YAML rather than applying a partial
snippet.
It has a durable named external trigger and a manual dispatch with **no
league/date/mode/write inputs**; it deliberately has no GitHub Actions cron
trigger. An adaptive, read-only gate runs first. It reads the latest
Git-authoritative compact operator state and a bounded current-day schedule.
Preflight runs independently, so a gate import/configuration failure remains
visible instead of silently suppressing all matrix work; only matrix rows with
work due run the full league pipeline. The selected rows remain fail-independent across every league in the canonical
registry. A league without a compliant automated official schedule/result
adapter still runs fail-closed: it reports an explicit unsupported-source gap
and may use only already-canonical imported games/history; it never fabricates
fixtures, results, or named-book prices. The gate also wakes a league for its
first dated draft, the 09:00 SAST (or earlier first-fixture safety) dated-ticket
cutoff, each fixture's independent 60-minute pre-tip pick freeze, and
frozen-result reconciliation. One serialized publisher commits compact ticket
decisions/state, dated text cards, and immutable per-fixture pick freezes to
`main`; structured dated pick JSON, detailed
lifecycle renders, current-scan pick JSON, and expanded audits stay in cache and
bounded workflow artifacts.

### Reliable external trigger

GitHub cron is best-effort and may delay or drop scheduled events, so it is
intentionally disabled. Configure an external scheduler with retries every **30
minutes** to invoke GitHub's repository-dispatch endpoint. The scheduler must
keep its GitHub App installation token or fine-grained token outside this
repository and send the exact `magolide-scheduled-tick` event type:

```bash
curl --fail --retry 3 --retry-all-errors --request POST \
  --url https://api.github.com/repos/6ixtyn9-sudo/Ma_Golide_Satellites/dispatches \
  --header "Accept: application/vnd.github+json" \
  --header "Authorization: Bearer $GITHUB_APP_INSTALLATION_TOKEN" \
  --header "X-GitHub-Api-Version: 2022-11-28" \
  --data '{"event_type":"magolide-scheduled-tick","client_payload":{"source":"external-scheduler"}}'
```

That named event is the only write/publish path: it passes `--write` when
`ALLOW_MAGOLIDE_WRITES=true`, saves cache state, and may publish immutable
decisions plus compact settlement data. It is intentionally distinct from
`workflow_dispatch`, which remains a live read-only rehearsal. Configure the
external token with only the repository access needed to create a repository
dispatch; never place it in workflow YAML, repository files, artifacts, or job
logs.

### Adaptive tick gate

The 30-minute external event does **not** automatically run every expensive
pipeline. Before dependency installation, the gate restores `localdata` from
the latest `main` and checks each league's compact state/performance summary
and current acquisition-day schedule. It starts a league job when any of these
conditions holds:

- a ticket decision is due: an unfrozen scheduled/live game is at or inside the
  60-minute decision lead;
- the compact performance report has a pending position, awaiting game leg, or
  unresolved-result leg needing reconciliation or settlement; or
- a receipt-backed settlement card is missing or still displays `PENDING` after
  its compact performance history has no pending/result-reconciliation counts;
  or
- state/receipt validation or schedule retrieval is unavailable, malformed, or
  otherwise uncertain.

A valid frozen ticket or no-bet sidecar is required before the gate can regard a
reached decision window as complete. Missing/contradictory state, a failed
latest-`main` restore, provider errors, and unknown schedule statuses all
**fail open** to the normal full league job; the gate never turns uncertainty
into "no work." NBA, WNBA, NCAA men, NCAA women, and NBA G League use one
ESPN scoreboard request for their current acquisition day; CBA uses one
official Beijing-day schedule request. NBA G League uses the separately
validated `basketball/nba-development` route and cannot inherit NBA BPI or
book-price capability. FIBA EuroLeague Women uses its official GDAP competition
schedule (`209127`) and filters the canonical `gameDateTimeUTC` rows to the
Europe/Paris slate date. GDAP publishes future bracket placeholders with no
assigned teams: the gate ignores those non-fixtures and the full pipeline emits
bounded coverage telemetry for them. The gate and provider share the same FIBA
status contract; cancelled, abandoned, postponed, malformed, and unknown rows
make the gate fail open rather than becoming a `no_work` decision. FIBA APIM
header authentication uses a shared public fallback and retries after one
public-page key refresh; a refresh failure also fails open. Men's EuroLeague
currently has only the verified whole-season schedule endpoint, not a verified
date-bounded alternative, so it deliberately fails open rather than disguising
a season download as a lightweight gate. A `source_unavailable` expansion lane
likewise never fabricates an empty schedule: its first write-enabled run each
acquisition day publishes a durable `coverage_gap_YYYY-MM-DD.json` marker.
With no pending ticket/pick lifecycle state, that marker lets later ticks skip
the otherwise identical coverage-gap pipeline; missing, corrupt, or
identity-mismatched markers fail open.

`workflow_dispatch` remains deliberately different: it forces the existing
all-league read-only rehearsal, even if the adaptive gate would otherwise skip
idle leagues. A skipped external tick has no matrix artifacts and no publisher
job, therefore cannot write or publish anything.

Each matrix job has isolated state and artifacts:

- `localdata/<state_key>`, where `state_key` comes from the canonical league
  registry (including every expansion lane);
- a league-specific DuckDB path, evidence-verified cache prefix, run ID, and
  artifact name. The `evidence-v1` cache epoch is universal: no league restores
  a cache created before the current source-identity/evidence policy;
- an acquisition date derived automatically in that league's acquisition
  timezone, never a dispatch-selected date.

### Paper-ticket decision cutoff

A current scan is candidate evidence, not a ticket publication. The
**acquisition timezone** identifies the fixtures belonging to a league/date,
but the dated-card decision is at **09:00 Africa/Johannesburg** on that named
date. If the first fixture's ordinary 60-minute pre-tip safety cutoff is
earlier, that earlier cutoff takes precedence. The first successful scan at the
applicable pre-tip cutoff freezes all currently valid, exact-priced paper
candidates into one immutable receipt. Every leg must still be strictly before
its own tip. The actual receipt records the policy, planned cutoff, and the
scan that made the decision.

If no qualifying pre-tip candidate exists at the cutoff, the operator writes an
immutable inline no-bet decision into that date's `auto_tickets_state.json` lock
instead. It prevents a later price, an old Actions cache, or a later-game
candidate from reopening the date without creating a per-date JSON sidecar. A
delayed run that misses the first tip also fails closed to that no-bet decision.
Before the cutoff, the single normal `auto_tickets_YYYY-MM-DD.txt`
card carries a clearly labelled `DRAFT ONLY` / `PRE-CUTOFF` section with the
broader exact named-price draft candidates. That section is not a receipt or
instruction: it never creates a stake, grade, bank/performance entry, or
real-money action, and retains source-permission and operational-validation
blockers. After either strict decision, selections/prices/stakes—or the no-bet
result—cannot change.

### EuroLeague aggregate-consensus paper board

Men's EuroLeague has a distinct `EUROLEAGUE_AGGREGATE_CONSENSUS_PRICE_PAPER_LEAGUE`
for the keyless BetExplorer representative average. It is not a route around
the named-book rule: every board price remains `aggregate_consensus`,
`unresolved`, and `probation`, and the strict ticket/receipt/bank lane rejects
it. The board may time-lock at the existing 09:00 SAST-or-earlier-first-tip
cutoff only while all selections remain pre-tip; a late scan never backfills a
board.

The immutable no-stake selection artifact is
`euroleague_consensus_paper_YYYY-MM-DD_roster.json`; the corresponding `.txt`
card can refresh only official-result grades. A roster holds at most three
full-game-moneyline selections, one per fixture, selected by the exact
observed average and calibrated model-versus-consensus edge. The only allowable
strict-lane blockers are the three aggregate evidence blockers and, where
disclosed, a developing calibration segment with at least 10 samples. It has
no stake, ticket, receipt, paper bank, ROI/P&L, or real-money instruction.

The publisher treats the roster as immutable Git-durable evidence and rejects
a cache/artifact that tries to replace it. Its mutable card is also protected
from regressions that would replace a rendered settled grade with `PENDING`.

### EuroLeague Women history, result, and market boundaries

FIBA ELW deep history currently covers **only the GDAP edition exposed by
competition `209127`**; it is not a historical competition catalogue. A FIBA
backfill cursor advances only after at least one reconciled game-page detail is
written. An all-failed detail batch therefore reports
`provider_error_cursor_retained` and retries instead of claiming progress. Each
fresh GDAP pass compares stored schedule-score provenance with the current
index; a changed official score causes a game-page re-verification before the
old final remains canonical.

Result rechecks are bounded by `history_days`. If a pending final ages out, the
results report emits a `pending_final_expired` / `result_unresolvable` coverage
gap rather than silently treating the game as settled or dropping it from
operational visibility. Resolve or explicitly quarantine such a game before
using it as history.

ELW has no validated named bookmaker or point-in-time price source. Market and
operator reports carry `pricing_evidence_status=unpriced_no_validated_named_book_lane`;
any paper singles/acca counts are
unpriced decision-support counts, **not** price, positive-EV, or betting
recommendations. The standalone `scripts/build_markets.py` defaults `--tz` to
the selected league's acquisition timezone (Europe/Paris for ELW), not UTC;
pass `--tz` explicitly only when deliberately inspecting another local slate.

NBA G League has the same no-price boundary: it may publish an empirical
history forecast only after the normal reconciled-history and residual gates.
It has no validated external pre-tip prediction source or named-book lane, so
its reports carry `pricing_evidence_status=unpriced_no_validated_named_book_lane`
and can never turn a relayed ESPN offer into a price, EV, or ticket claim.

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
| `THE_ODDS_API_KEY` | NBA/WNBA/NCAA_M/NCAA_W full-game markets; NBA quarters and H1; NCAA H1/H2 where the provider offers them; probationary Australia NBL sport-key lane | The first supplemental source to validate, especially for NCAA_W; an NBL account/region/bookmaker smoke is required before ticket eligibility |
| `THERUNDOWN_API_KEY` | NBA full-game moneyline and totals | A second named NBA source |
| `SPORTSGAMEODDS_API_KEY` | NBA full-game and Q1–Q4 totals | NBA quarter-total coverage |
| `PREDIXSPORT_API_KEY` | NBA and EuroLeague predictions | Personal/non-commercial free tier only; adapter remains documentation-only until an authenticated smoke succeeds |

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
- Named external-trigger runs default to `--write`: the workflow falls back to
  `ALLOW_MAGOLIDE_WRITES=true` when the repository variable is unset. Set the
  variable to `false` to disable canonical persistence; runs still acquire
  available providers and report missing credentials.
- Named external-trigger runs save league state. When canonical writes are
  disabled, this preserves public acquisition caches and receipts but does not
  persist DuckDB/CSV sports state or advance deep-history cursors. Reports and
  immutable ledgers are uploaded on every outcome when present. Git is
  authoritative for operator files after cache recovery.
- Manual runs upload derived `picks_*` and `auto_tickets_*` files only as
  inspection artifacts. **Named external-trigger runs only** may publish durable
  operator files to `localdata/<league>` on `main`, so a manual rehearsal cannot
  create a durable ticket or no-bet decision. Git publication never includes
  warehouses, CSV exports, acquisition ledgers, or credentials.
- Notifications additionally require `ALLOW_NOTIFICATION_WRITES=true` and the
  WhatsApp secrets; they are sent only on write-enabled named external-trigger
  runs.

Every ESPN league run also attempts seven recent result dates and, when
canonical writes are enabled, 50 deep indexed events. Deep cursors, historical
coverage classes, verified index floors, and cache policy are documented in
[ESPN acquisition and history](ESPN-ACQUISITION.md). Historical rows alone do
not satisfy pre-tip prediction/book requirements or establish calibration.

Configure only permanently-free provider secrets as needed:
`PREDIXSPORT_API_KEY`, `THERUNDOWN_API_KEY`, `SPORTSGAMEODDS_API_KEY`, and
`THE_ODDS_API_KEY`. Missing credentials disable only the affected provider and
remain visible as coverage gaps; they do not fabricate substitute evidence.
Predixsport and NCAA provider coverage is `DOCUMENTATION_ONLY` until
an authenticated smoke succeeds.

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
