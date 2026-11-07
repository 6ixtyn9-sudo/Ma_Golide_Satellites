# Ma Golide acquisition and warehouse report

**Architecture:** headless DuckDB + CSV.GZ system of record

**Pattern reference:** `6ixtyn9-sudo/Edge-Factory` local-data operations

**Revision date:** 2026-11-07

**Verification:** 225 passed, 4 deselected; critical Ruff checks passed

## Outcome

Ma Golide no longer uses Google Sheets as the intended system of record. The
validated basketball provider adapters and domain semantics now write to a
canonical DuckDB store. A single orchestrator persists runtime state, creates
immutable evidence, runs a research-only assay, materializes CSV, and can emit
a deduplicated operational WhatsApp heartbeat.

The system is fully headless. The dashboard/API, web server, external read-model
databases, synchronization code, and their dependencies have been removed. Read
access is through DuckDB SQL, the coverage CLI, CSV.GZ, and JSON reports. The default runtime requires only DuckDB and Requests; legacy Google
packages are isolated in an optional extra.

## Implemented components

### Canonical persistence

- `ingestion/duckdb_store.py`
  - Six canonical tables with stable natural keys.
  - Transactional upserts and typed query fields plus complete JSON payloads.
  - Existing prediction/book freeze behavior preserved across database reload.
  - Independent pre/post prediction and book slots.
  - True read-only/no-create dry runs.
  - Atomic CSV.GZ materialization and guarded CSV recovery.
  - Per-stage run reports and 30-day telemetry retention.
- `ingestion/warehouse.py`
  - Independent DuckDB research connections.
  - `prediction_facts`, `book_facts`, `latest_pregame_predictions`, and
    `game_results` temporary views.
- `scripts/build_warehouse.py` and `scripts/recover_warehouse.py`
  - Canonical DuckDB → CSV.GZ materialization.
  - Guarded CSV.GZ → new DuckDB recovery without merging into active state.

### Immutable headless operations

- `ingestion/ledger.py`
  - Deterministic canonical snapshots scoped to the league-local slate.
  - SHA-256-addressed pregame/freeze/result ledgers.
  - Identical rerun no-op and explicit conflict rejection.
- `scripts/daily.py`
  - One `--auto-once` path for capture → freeze → results → ledgers → assay →
    CSV → optional notification.
  - Environment-gated canonical, fixture, and external writes.
  - Thirty-day run telemetry retention; immutable ledgers retained.
- `.github/workflows/ingestion.yml`
  - Three-hour schedule and manual dispatch.
  - One concurrency group, cache restore/save, tests, redacted summary, and
    artifact upload.
  - No Google environment, worksheet steps, web service, or external database.

### Research

- `ingestion/mining.py` and `scripts/mine_assay.py`
  - Frozen pre-tip moneyline outcomes joined only to reconciled official finals.
  - Existing segment/shrinkage/Wilson/grade/lift logic reused.
  - Post-tip, unfrozen, tied, missing, and unreconciled cases excluded.
  - No expected-total-to-O/U conversion.
  - All output remains `research_only`, `certified: false`, with zero certified
    picks pending walk-forward and odds-aware validation.

### Operational output

- `python -m ingestion.cli coverage` reports actual stored coverage.
- `ingestion.warehouse.connect()` exposes direct SQL research reads.
- `scripts/notify.py`
  - Operational coverage heartbeat, not a betting recommendation.
  - Content-hash delivery dedupe and explicit WhatsApp send guard.
  - Fixture reports can be previewed but never sent externally.

## Provider/domain behavior retained

The redesign preserves the existing tested rules for:

- canonical home/away/start identity and ambiguous join quarantine;
- league-local slate dates and UTC storage;
- Hardwood → ESPN → enabled Predixsport NBA selection and enabled Predixsport
  EuroLeague selection;
- coherent single-provider prediction bundles and orientation checks;
- provider-specific book evidence and explicit missing-market gaps;
- pre-tip and pre-match freezing;
- quarter + overtime = final reconciliation;
- no fabricated quarter lines or model-derived bookmaker prices;
- idempotent reruns and redacted coverage reports.

The legacy non-destructive A:R/S:BX worksheet protections remain covered for
explicit compatibility operations, but they are no longer the target data
architecture. Legacy packages are not installed in the minimal default runtime.

## Verification evidence

Executed offline with sockets blocked by the test suite:

```text
225 passed, 4 deselected
```

Critical static checks:

```bash
ruff check ingestion scripts tests --select E9,F63,F7,F82
# All checks passed
```

Additional verification performed:

- DuckDB close/reopen round trip for every canonical table.
- Frozen pre-tip and pre-match overwrite rejection after reload.
- Post-tip predictions not marked as the frozen pre-tip record.
- Dry run leaves both absent and existing DuckDB files unchanged.
- CSV.GZ values and SQL view values agree with canonical records.
- CSV.GZ recovery into a new warehouse preserves nested results and freeze flags.
- Ledger snapshots include only the target league-local slate, and 30-day
  telemetry retention leaves canonical facts and immutable ledgers intact.
- Full isolated fixture daily write produces one game, one prediction, three run
  stages, three immutable ledgers, and six CSV.GZ tables.
- Immediate fixture rerun reuses identical ledger bytes and does not duplicate
  canonical rows, including when the wall clock changes between runs.
- Notification fake dispatcher sends once and suppresses an identical rerun.
- Fixture persistence requires a separate opt-in and paths isolated from
  canonical state in both the CLI and orchestrator; external fixture
  notification is refused.
- A clean editable install with only DuckDB and Requests successfully imports
  the headless runtime and executes the fixture dry run without Google or web
  packages.
- Workflow YAML parses with schedule and workflow-dispatch triggers.
- Final command-line recovery reproduced the exact warehouse counts.

## Explicit non-claims and remaining gates

- No authenticated Predixsport, SportsGameOdds, or TheRundown production schema
  was newly verified.
- No live production enablement, real Google Sheets write, or WhatsApp send was
  performed.
- Fixture results are `FIXTURE_VERIFIED`, not live coverage.
- The assay is not walk-forward certification and produces no official picks.
- Odds-aware ROI, CLV, decay gating, and an official certified pick ledger remain
  prerequisites for betting publication.
- Historical Sheets data still requires an explicit migration if it is to enter
  DuckDB; it was not bulk-imported or assumed trustworthy automatically.
- Pushing workflow changes still depends on the configured GitHub connection
  having workflow-file write permission. No PR is claimed by this report.

## Operator start point

Read `docs/INGESTION_RUNBOOK.md`. The safest first command is the hermetic daily
cycle, followed by a live dry run. Enable canonical and notification writes
independently only after inspecting their previews.
