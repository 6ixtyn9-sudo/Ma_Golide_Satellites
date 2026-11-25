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
- `--history-days N` and `--history-requests N` bound recent daily backfill.
  Their standard default is 21. NBA G League defaults to a 56-day/56-request
  **date-bounded ESPN scoreboard bootstrap** because its approved route has no
  deep event-index contract; an explicit value overrides that plan, and an
  explicitly lower request budget caps the requested day range. The bootstrap
  does not loosen forecast, calibration, price, or ticket gates.
- `--deep-history-events 50` bounds indexed historical summaries per
  write-enabled league run; dry runs never advance that cursor.
- The **history repair** pass (`PipelineConfig.history_repair_budget`, default
  25 detail reads) runs at the start of a write-enabled `pregame` for lanes in
  `DETAIL_RESULT_REFRESH` — today only KBL, whose list route settles a final
  with no period split. Newest-first, it re-reads the detail route for settled
  games whose stored final cannot reconcile, replacing the totals-only record
  with reconciled quarters. `min_history_games` and `min_residual_games` are
  both 20, so a lane starts with an empty corpus: the first repaired run crosses
  the baseline floor and a second crosses the residual floor before the
  empirical model emits a forecast. A row whose detail is missing, misidentified,
  or still unreconcilable is left untouched and named in a
  `historical_reconciliation` gap; the pass never upgrades a final on faith.
- `--book draftkings` and `--book-provider theoddsapi` restrict book evidence.
- `--localdata PATH` redirects reports/ledgers/exports.
- `--warehouse PATH` overrides the canonical DuckDB path.
- `--keep-days 30` bounds operational runs and daily/market reports.

There is no in-repository daemon and no GitHub Actions cron fallback. An
external scheduler invokes the workflow hourly, at :37. That interval is set
from measurement, not preference: a healthy tick runs ~42 minutes end-to-end
(22-29 minutes of lane work across 19 parallel lanes, then ~13 minutes of
publishing), so the 30-minute cadence it replaced could not even contain one
tick, and because the
workflow serialises on one concurrency group with `cancel-in-progress: false`,
the overlap surfaced as the *pending* tick being cancelled when the next one
arrived — acquisition silently skipped rather than run twice.

Each slate gets its pre-tip opportunity from a 120-minute lead
(`TICKET_DECISION_LEAD_TIME` in `forecasting/operator.py` and
`DECISION_LEAD_MINUTES` in the gate must move together, or the gate rejects the
freeze the writer produced). That lead is what puts two hourly ticks inside
every window:

- Lane cap is 360 minutes (`league` job, raised from 60 at 2026-11-23T19:59Z
  after a healthy tick measured 54 minutes end to end). Overlap is not the
  risk the cap controls: `concurrency` holds one group with
  `cancel-in-progress: false`, so a long run makes the next tick **queue**
  instead of running beside it.
- The cost of the longer cap is therefore a wedge, not a collision: a stuck
  lane holds the group — and every later tick — for up to six hours. If that
  trade is ever wrong, 90 minutes still clears a 54-minute run with the
  publisher's 15 minutes of headroom.
- If both a run and a queued tick are outstanding when a third tick arrives,
  GitHub cancels the queued one. That is the failure mode to watch for in the
  Actions list: a `cancelled` run whose predecessor ran long.

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
first dated draft, the universal midday dated-ticket cutoff (12:00 SAST for
every league, or earlier first-fixture safety), each fixture's independent
120-minute pre-tip snapshot, and
frozen-result reconciliation. One serialized publisher commits compact ticket
and lifecycle state (including immutable inline per-fixture snapshots) and
dated text cards to `main`; structured dated pick JSON, detailed
lifecycle renders, current-scan pick JSON, and expanded audits stay in cache and
bounded workflow artifacts.

### The gate imports no third-party packages

`gate` runs on a bare `setup-python` with nothing installed — deliberately, so a
dependency failure in the pipeline cannot take the scheduler down with it. The
gate still imports provider modules, because it needs their **parsers**; it does
its own page reads with `urllib`.

That contract holds only while a parser stays transport-free. It did not:
`ingestion/http.py` imported `requests` at module scope, so importing the
B.League adapter to reach `parse_slate` raised `ModuleNotFoundError` and both
Japan lanes were gated on 2026-11-25 with

```
schedule_uncertain: No module named 'requests'
```

for a reason no schedule had anything to do with, while preflight in the same run
had installed the package and passed its suite. `requests` is now optional at
import time, and constructing an actual `HttpClient` without it raises a message
that says which dependency is missing. `tests/test_adaptive_gate.py` runs the gate
in a subprocess with `sys.modules["requests"] = None` to keep that true.

Rule for new adapters: a module the gate imports must parse without a transport.
Put network calls behind an injected `fetch`, or lazily inside the function that
needs them.

### What a tick costs

Wall clock is a scheduling input, not an afterthought. A tick used to cost the
same whether a lane had a live slate or not: on 2026-11-25 all 13 lanes ran the
full pipeline (history acquisition, replay, forecast, freeze, results, tuner) to
produce **two** live scans, and ten of them discovered nothing at all. Ticks grew
past an hour, the scheduler kept dispatching, and the concurrency group
(`group: magolide-automatic-available-leagues`, `cancel-in-progress: false`)
meant each new tick replaced the one already queued — so ticks were silently
dropped as well as slow.

Two things changed, and one measurement was added.

**A lane's mode is now part of the matrix.** `adaptive_gate` classifies each
selected lane from the reasons it is due:

| mode | reasons | tick |
|---|---|---|
| `full` | a draft, a freeze, a ticket window, anything unclassified | the whole pipeline |
| `settle_only` | settlement, reconciliation, settlement-view refresh, coverage-gap retry, `schedule_uncertain`, `environment_error`, `state_uncertain` | scan, but `--skip-backfill --skip-tuning` |

The classification fails safe: one unrecognised reason means `full`. A
settlement-only tick still runs the scan — a fixture that appears late must not
be missed — but skips the two most expensive stages, neither of which can change
a settled result. The next full tick replays history normally, and the skipped
backfill is recorded as `status: skipped_settle_only_tick` rather than as an
empty history.

**The matrix has no `max-parallel` cap.** A cap of 8 over 13 lanes ran in two
waves, so a tick cost roughly twice its slowest lane while the second wave had
idle capacity. The matrix itself is the cap.

**Every tick now names where it went.** `summary.stage_seconds` records elapsed
seconds per stage — `backfill`, each command (`pregame`, `freeze`, `results`) —
so the next optimisation is chosen from the run's own numbers instead of from
the log's line counts. Read it before proposing another change to this workflow.

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

### The Odds API free-tier reservation guard

The Odds API key is a shared monthly provider allowance, not a per-runner
credential. The provider documents that current-odds cost is the number of
requested markets times the effective bookmaker-region groups; with a named
bookmaker filter, up to ten books count as one group. The routine price request
therefore asks only for `h2h,totals,spreads` from the one named DraftKings group:
**three credits maximum**. See the provider's [v4 quota
rules](https://the-odds-api.com/liveapi/guides/v4/).

Before an automated matrix run receives `THE_ODDS_API_KEY`, the serialized
`odds-budget` job commits a no-secret reservation to
`localdata/theoddsapi_budget.json`. It is deliberately conservative:

| Guard | Value | Meaning |
| --- | ---: | --- |
| Advertised Starter allowance | 500 credits/month | Provider plan, not a promised local balance |
| Ma Golide reservation cap | 360 credits/month | Never reserve more than this for routine work |
| Protected provider reserve | 140 credits | Stops a new three-credit permit if provider-reported remaining would fall below this value |
| Daily reservation cap | 12 credits | At most four routine full-game scans across **all** leagues per SAST date |
| Routine request | 3 credits maximum | `h2h,totals,spreads` and one effective bookmaker group |

Only NBA, WNBA, NCAA men, and NCAA women can receive a routine permit, ordered
NCAA women, NBA, NCAA men, then WNBA. The permit is issued only for the gate's
`ticket_creation_or_acquisition_due` reason, never for draft rendering, settlement,
results, state uncertainty, a source outage, or a manual rehearsal. A
reservation remains spent if a worker fails: this prevents a failed publisher or
parallel runner from repeatedly probing the provider with the same slate.

This priority is scoped to The Odds API's shared allowance. It is not a global
league-adoption policy: another league may gain a future compliant source path
only through its own evidence contract and bounded validation, as described in
[the league readiness roadmap](LEAGUE-ADOPTION-ROADMAP.md).

Routine acquisition is full-game only. The documented native period markets are
available through a charged per-event endpoint, so the scheduler never fans out
one request per NCAA half or NBA/WNBA quarter. Period parsing remains strict,
but period collection requires a future explicit, higher-cost approval. Every
routine query also fixes `regions=us` and `bookmakers=draftkings`, ignoring
any local override while permit enforcement is enabled, and uses documented
`commenceTimeFrom` and `commenceTimeTo` bounds in the league acquisition
timezone; an off-slate response cannot become a paid catch-all future-events
scan.

`workflow_dispatch` remains a credential-free rehearsal: it receives no Odds
API permit or key. Completed automated pregame reports reconcile only the
provider's numeric `x-requests-remaining`, `x-requests-used`, and
`x-requests-last` headers into the ledger; the lowest observed remaining value
wins. Missing or malformed state, a failed reservation commit, a missing permit,
or a provider balance that would fall below reserve all fail closed for *new*
named-book requests. They produce an explicit unpriced/deferred report, never inferred
odds or a post-cutoff ticket.

### Adaptive tick gate

The hourly external event does **not** automatically run every expensive
pipeline. Before dependency installation, the gate restores `localdata` from
the latest `main` and checks each league's compact state/performance summary
and current acquisition-day schedule. It starts a league job when any of these
conditions holds:

- a ticket decision is due: an unfrozen scheduled/live game is at or inside the
  120-minute decision lead;
- the compact performance report has a pending position, awaiting game leg, or
  unresolved-result leg needing reconciliation or settlement; or
- a receipt-backed settlement card is missing or still displays `PENDING` after
  its compact performance history has no pending/result-reconciliation counts;
  or
- a legacy compact pick-card marker needs its bounded report-only readable-card
  restoration; or
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
while every dated-card decision is at **12:00 Africa/Johannesburg** on that
named date. If the first fixture's ordinary 120-minute pre-tip safety cutoff is
earlier, that earlier cutoff takes precedence. The first successful scan at the
applicable pre-tip
cutoff freezes all currently valid, exact-priced paper
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

ELW has no dedicated named-book API integration, but a future canonical fixture
now receives the bounded `betexplorer_direct` fallback attempt. A response can
supply a full-game named H/A moneyline, O/U total with its explicit line and
sides, or Asian Handicap spread with its explicit home-oriented line and sides.
Basketball source navigation can lack a trustworthy active total/spread line;
exact named rows are still stored and receipted, but the ordinary downstream
main-line preference is deliberately not relaxed. Its direct-page provenance
remains `unresolved` / `probation`, and an empty or unmatched comparison row
remains explicitly unpriced. The standalone `scripts/build_markets.py` defaults
`--tz` to the selected league's acquisition timezone (Europe/Paris for ELW),
not UTC; pass `--tz` explicitly only when deliberately inspecting another local
slate.

NBA G League has the same direct fallback for displayed full-game moneyline,
total, and spread rows. It still has no validated external pre-tip prediction
source, so its empirical history forecast requires the normal reconciled-history
and residual gates; the fallback cannot manufacture prices for half, quarter,
or highest-quarter forecasts.

### Optional odds-provider credentials

The public feeds remain useful without credentials, but they do **not** provide
uniform odds coverage: ESPN's public scoreboard is full-game-only when it
includes a book at all, BetExplorer's bounded direct fallback covers only its
displayed full-game H/A, O/U, and Asian Handicap rows, and NCAA_W moneyline
coverage is known to be absent from the ESPN free surface. To enable the
already-implemented supplemental adapters in GitHub Actions, set the
corresponding **Actions secrets** (never repository variables or checked-in
files):

| Secret | Implemented adapter coverage (verify live) | Use it for |
| --- | --- | --- |
| `THE_ODDS_API_KEY` | Budget-permitted NBA/WNBA/NCAA_M/NCAA_W full-game `h2h,totals,spreads` only; NCAA_W uses the documented `basketball_wncaab` key | The first supplemental source to validate, especially for NCAA_W. The workflow withholds this key from manual/reconciliation/draft work and from unpermitted or probationary lanes |
| `PREDIXSPORT_API_KEY` | NBA and EuroLeague predictions | Personal/non-commercial free tier only; adapter remains documentation-only until an authenticated smoke succeeds |

The workflow injects Predixsport as a process environment variable. It withholds
The Odds API key unless the automated budget job has first committed a matching,
three-credit permit for the current league/date. That permit controls only the
paid API lane: a missing secret or permit still runs the bounded direct
BetExplorer fallback for canonical games lacking a fresh named full-game core
market. The pregame report records both statuses rather than presenting a
no-price result as an empty market response. The GitHub job summary also
includes selected-side price coverage by market/period and the redacted source
status, so a no-price result can be separated into a provider-capability gap, a
disabled credential, an identity rejection, a budget deferral, or a true empty
response.

Safety behavior is deliberate:

- Manual dispatch is a live all-league canonical **dry run**. It never receives
  `THE_ODDS_API_KEY`, never reserves a provider slot, never adds `--write`, and
  never sends notifications, even when `ALLOW_MAGOLIDE_WRITES` is true.
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
`PREDIXSPORT_API_KEY` and `THE_ODDS_API_KEY`. Missing credentials disable only the affected provider and
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

### Refused and unreachable schedule reads

A schedule host that refuses us costs the lane its slate, never the tick. An
`HttpError` on the `betexplorer_fixtures` and EuroLeague schedule branches
becomes a named `official_schedule` `provider_error` gap (a refusal also records
`coverage.http_status`, e.g. `429`) plus a `provider_error` health row, and the
run continues on the canonical stored slate so settlement, grading and the
market report still happen. An unreadable source is named as unreadable; it is
never rendered as an empty slate, and it never unwinds to
`{"status": "failed"}` with nothing recorded.

Run 45376480832 (2026-11-24T10:40Z, `ISRAEL_LIGA_LEUMIT`) is the fixture case:
`HTTP 429` on `/basketball/israel/liga-leumit/fixtures/` after the transport's
two retries exited 1 with no card, no gap and no lifecycle work.

A `Retry-After` header is honoured when the host sends one, clamped to
`ingestion.http.MAX_RETRY_AFTER_SECONDS`; a longer one means "come back later",
which is a degradation, not a reason to sleep through the tick. A page that
arrives *without* its fixtures table (`IdentityError`) still raises: a source we
can read but no longer parse is a code failure, not a host's bad day.
