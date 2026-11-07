# MA GOLIDE — DUCKDB/CSV MASTER BUILD PROMPT
## Authoritative headless acquisition, warehouse, research and operations contract
### Version 2.0 — 2026-11-07

This document supersedes the former Google Sheets-centered acquisition prompt.
The target architecture follows Edge Factory's local-data operating patterns:
**CSV + DuckDB-first, one daily orchestrator, immutable ledgers, headless SQL/
CLI read access, and optional notifications**. There is no web server, external
database, or read-model synchronizer. Google Sheets is
legacy migration or explicit compatibility output only.

A future coding agent must preserve this direction. Do not reintroduce Sheets,
worksheet tabs, satellite registries, Apps Script, or Google service accounts as
the canonical store or as an implicit fallback.

---

# PART A — MISSION AND OPERATING PRINCIPLES

## A1. Mission

Build and operate an unattended basketball acquisition and research system that:

1. Discovers canonical NBA and EuroLeague fixtures from official/keyless sources.
2. Selects coherent, fully attributed provider prediction bundles.
3. Captures only actual bookmaker market observations.
4. Preserves pre-tip predictions and pre-match books through immutable freezes.
5. Reconciles official quarter, overtime, and final results.
6. Persists all canonical state in DuckDB with CSV.GZ materializations.
7. Mines research evidence without auto-certifying unvalidated strategies.
8. Provides direct headless DuckDB SQL and coverage/research CLI read access.
9. Runs through one resumable orchestrator on a three-hour Actions cadence.
10. Emits explicit coverage, gaps, reports, and optional deduplicated operational
    notifications.

## A2. Source of truth

The authority order is:

```text
provider response + capture timestamp
  -> validated canonical adapter record
  -> DuckDB natural-key state
  -> immutable content-addressed ledger / CSV.GZ materialization
  -> derived research, SQL/CLI reads and operational reports
```

`localdata/warehouse.duckdb` is the operational system of record and the direct
read/research surface. Canonical CSV exports must represent the same records.
Notifications, reports, and legacy Sheets exports are derived and disposable.

If DuckDB is unavailable, fail visibly. Never silently fall back to Google
Sheets, an empty successful slate, or an inferred provider record.

## A3. Non-negotiable provider truth

- Never fabricate bookmaker quarter lines, alternate lines, team totals, live
  prices, or half-time lines.
- Never substitute a model expected total, predicted margin, fair price, or
  blend for a bookmaker observation.
- Never mix one provider's probabilities with another provider's scores or
  margins and present the result as a source prediction bundle.
- Never assume an API field exists because another provider exposes it.
- Missing optional provider keys disable that provider only; keyless acquisition
  continues and reports actual coverage.
- An HTTP 200 alone is not market coverage. Count normalized usable records.
- Keep raw source IDs, observed timestamps, source-modified timestamps where
  supplied, model/version metadata, and payload hashes.
- TheRundown Free must not be described as providing quarter/half/alternate/
  team-total/live markets. Predixsport and SportsGameOdds authenticated schemas
  remain unverified until checked against actual authorized responses.
- Forebet and new Forebet acquisition are out of scope. Existing legacy model
  references are not permission to add a Forebet provider.

## A4. Dry-run and least privilege

- Dry run is the default for acquisition and notifications.
- Canonical writes require `--write` and `ALLOW_MAGOLIDE_WRITES=true`.
- WhatsApp sends require the explicit send option, credentials, and
  `ALLOW_NOTIFICATION_WRITES=true`.
- Legacy Sheets writes additionally require `ALLOW_SHEET_WRITES=true`.
- Fixture persistence requires a separate `ALLOW_FIXTURE_WRITES=true` and an
  isolated warehouse. Fixture reports must never trigger external notification.
- No secrets in Git, `localdata/`, reports, exception messages, or URLs unless
  the provider design strictly requires them there.
- Headless SQL/CLI readers need no web framework or external database secret.

---

# PART B — CANONICAL DOMAIN CONTRACTS

## B1. Game identity

A game has:

- canonical `game_id`;
- league and season;
- canonical home/away team IDs and names;
- aware UTC scheduled start;
- scheduled/final/error status as supported by the model;
- game phase;
- provider-to-source-event-ID mapping;
- optional official actual result;
- optional `rescheduled_from` identity.

Join cross-provider records using provider event identity where known. Otherwise
require exact canonical home + away + aware start within the approved tolerance.
A name-only or ID-coincidence join is forbidden. Contradictory or ambiguous
joins are quarantined with an explicit gap.

Use league-local slate dates for discovery and provider date filters. Store
timestamps in UTC. Preserve official rescheduling lineage rather than inventing
a new unrelated game.

## B2. Prediction bundle

A prediction belongs to one provider and contains only observed source fields:

- `game_id`, provider, observed UTC;
- optional source-modified UTC, model, model version, producer, revision;
- optional home/away probabilities;
- optional expected home/away scores;
- optional expected margin and expected total;
- `pre_tip` boolean.

Validation rules:

- probabilities are 0..1;
- a home/away pair must sum to 1 within the approved tolerance;
- expected scores must agree with supplied margin and total within tolerance;
- preserve the entire coherent bundle from one provider;
- reject or quarantine contradictory orientations;
- missing values remain missing.

Provider selection priority remains:

- NBA: Hardwood, ESPN, Predixsport when enabled;
- EuroLeague: Predixsport when enabled.

Selection may prefer a fresher complete bundle. It must not splice fields across
providers. Post-tip observations use `provider|post` slots and cannot displace
`provider|pre` evidence.

## B3. Book snapshot

A book snapshot is one observed book/market/period record:

- canonical game and provider event identity;
- provider and actual book;
- market and exact period;
- observed UTC and optional source-modified UTC;
- optional total/spread line;
- optional decimal home/away or over/under prices;
- pre-match flag, main-line flag, and payload hash;
- source identity metadata used for safe joining.

Do not derive Q2, Q3, Q4, first-half, alternate, or live lines from a full-game
line. Do not infer a total from a price. American prices must be converted to
decimal at the provider boundary. An absent market is a `CoverageGap`.

Book storage slots are:

```text
provider|book|market|period|pre
provider|book|market|period|post
```

Pre-match and post-match records coexist. Every distinct pre-match market/
period snapshot selected at freeze is protected, not merely one arbitrary
snapshot per provider.

## B4. Actual result

Official results retain:

- period kind;
- home and away Q1..Q4 arrays;
- optional combined overtime values and overtime count;
- official home/away final;
- result source.

Before a result enters the reconciled research surface, verify independently
for both teams:

```text
Q1 + Q2 + Q3 + Q4 + OT == official final
```

Missing periods, non-quarter period kinds, contradictory sums, and ambiguous
results are quarantined. Do not guess an overtime allocation or force a final
to reconcile.

## B5. Coverage and health

Coverage must state what actually happened:

- fixtures found;
- selected predictions by provider;
- full-game and quarter book coverage;
- reconciled finals;
- frozen predictions used;
- explicit gaps and provider health;
- quota metadata when supplied by a provider.

Use states already supported by the domain (`missing`, `unauthorized`,
`provider_error`, `partial`, fixture verification, etc.). Never round a partial
provider success up to complete coverage.

---

# PART C — DUCKDB AND CSV SYSTEM OF RECORD

## C1. Store contract

`ingestion.store.Store` remains the shared interface. Implementations have
distinct roles:

- `InMemoryStore`: tests and smoke only;
- `JsonFileStore`: local development/CI scratch only;
- `DuckDBStore`: canonical unattended store;
- `SheetsDataStore`: explicit legacy migration adapter only.

`DuckDBStore(path, persist=False)` is a true dry run. It opens an existing file
read-only, uses memory when the file is absent, and never creates or changes the
file. Persistent mutations mirror the same freeze and natural-key behavior as
the in-memory reference implementation.

## C2. Required tables and keys

| Table | Natural key | Required role |
|---|---|---|
| `games` | `game_id` | Identity, schedule/status, provider IDs, actual JSON |
| `predictions` | `(game_id, slot)` | Coherent bundle, pre/post slot, freeze state |
| `book_snapshots` | `(game_id, slot)` | Typed market observation and payload |
| `team_aliases` | `canonical` | Canonical name and variants |
| `provider_health` | `(provider, league)` | Attempt/success/coverage/quota state |
| `pipeline_runs` | `(run_id, command)` | Redacted stage report |

Use typed SQL columns for query-critical fields and complete JSON payloads for
variable source detail. Keep indexes on league/schedule, game/provider/slot,
and game/market/period. Every mutation is transactional and idempotent.

## C3. Freeze semantics

- A frozen pre-tip prediction rejects a changed overwrite.
- An identical frozen replay is a no-op.
- A post-tip shadow does not reset the provider's pre-tip freeze after reload.
- A frozen pre-match book snapshot rejects a changed overwrite.
- Post-match books remain separate and are not marked as the frozen pre-match
  record.
- Results resolve the frozen whole bundle; they do not use whichever provider
  response happens to arrive last.

These rules must survive closing and reopening DuckDB, not just one process.

## C4. CSV and SQL research

`DuckDBStore.export_csv()` and `scripts/build_warehouse.py` materialize all six
canonical tables as CSV.GZ. Exports use a temporary file + atomic replace.
`DuckDBStore.restore_csv()` and `scripts/recover_warehouse.py` rebuild a new
warehouse from those tables; recovery refuses a non-empty destination and
preserves nested result types and freeze flags. Runtime state is cached, not
committed to Git.

`ingestion.warehouse.connect()` exposes temporary views:

- `prediction_facts`;
- `book_facts`;
- `latest_pregame_predictions`;
- `game_results`.

Views expose observed values only. They must not invent kickoffs, results,
prices, provenance, or official pick selection.

---

# PART D — IMMUTABLE LEDGERS AND RESEARCH

## D1. Ledger rules

Write pregame, freeze, and result ledgers under `localdata/ledgers/`. Serialize
canonical facts deterministically and include a SHA-256 digest in each filename.

- Identical content/path: no-op with `created: false`.
- Changed content: new content-addressed file.
- Same explicit path with different bytes: `LedgerConflictError`.
- Never rewrite an older ledger to make the latest state look tidy.

Per-stage reports belong in `pipeline_runs` and `localdata/reports/`. Reports
may include operational timing; immutable canonical ledgers should remain
stable across identical reruns.

## D2. Current assay scope

The warehouse miner may evaluate only frozen pre-tip home/away probability
picks against official non-tied reconciled finals. It must exclude:

- post-tip shadows;
- unfrozen predictions;
- missing or tied probabilities;
- tied, missing, or unreconciled finals;
- any attempt to convert expected totals into O/U bets.

Reuse the existing segment, shrinkage, Wilson lower-bound, grade, and lift
functions. Preserve provider as a segment dimension.

Until walk-forward validation and odds-aware evaluation exist, every result is:

```text
status: research_only
certified: false
certified_picks: 0
certification_reason: walk_forward_validation_required
```

Do not label a high in-sample hit rate as an official edge or betting pick.

## D3. Future pick certification gate

Before introducing an official pick ledger, implement:

- chronological train/validation splits with no future leakage;
- minimum train and validation sample requirements;
- actual frozen book prices and odds-aware ROI;
- CLV measurement against timestamped closing observations;
- league/book/market coverage requirements;
- decay monitoring and bench/retire transitions;
- deterministic pick IDs and a freeze marker;
- replay tests proving no official pick changes after freeze.

A candidate HQ output or model confidence threshold alone is not certification.

---

# PART E — ONE ORCHESTRATOR AND PERSISTED STATE

## E1. Daily entry point

All unattended paths use:

```bash
python scripts/daily.py --auto-once --date YYYY-MM-DD
```

The orchestrator performs, in order:

1. pregame discovery/prediction/book capture;
2. due pre-tip prediction and pre-match book freezes;
3. official result acquisition and reconciliation;
4. content-addressed stage ledgers when writes are enabled;
5. research-only warehouse assay;
6. canonical CSV.GZ materialization;
7. optional operational notification;
8. atomic daily report creation.

No step needs Google credentials or a satellite worksheet. Unexpected canonical
failures fail the run. A notification failure is visible but does not roll back
committed acquisition.

## E2. Actions cadence and cache

The workflow runs every three hours and manually, with one concurrency group
and no cancellation of an active writer. It must:

1. check out the current branch;
2. install dependencies;
3. restore the latest `magolide-localdata-v1-` cache;
4. run hermetic tests;
5. execute one `daily.py --auto-once` iteration;
6. save `localdata/` on success;
7. upload reports and immutable ledgers;
8. write a redacted operational summary.

Do not commit runtime databases or secrets to Git. Prune parseable daily
reports and `pipeline_runs` telemetry after 30 days without deleting canonical
facts or immutable ledgers. Manual write/send inputs must agree with their
repository variables. Scheduled external writes
remain disabled until deliberately enabled. While canonical writes are
disabled, scheduled iterations use committed fixtures rather than consuming
authenticated provider quota without persisting state; manual dry runs are the
deliberate live-coverage preview path.

## E3. Idempotency acceptance

For identical fixtures and source responses, a second run must demonstrate:

- no duplicate canonical games, predictions, books, or run-stage keys;
- no frozen overwrite;
- unchanged ledger bytes and `created: false`;
- stable assay output;
- stable notification content hash and no duplicate send;
- no new Google calls because Google is not in the path.

---

# PART F — HEADLESS READ ACCESS AND NOTIFICATIONS

## F1. Direct DuckDB/CSV read access

The canonical warehouse is also the application/research read surface. Use
`python -m ingestion.cli coverage` for operational summaries,
`scripts/mine_assay.py` for research output, and
`ingestion.warehouse.connect()` for SQL over `prediction_facts`, `book_facts`,
`latest_pregame_predictions`, and `game_results`.

Do not add a web framework, browser dashboard, external database, or
synchronization layer. A missing warehouse is a visible CLI/SQL error, not an
implicit legacy fallback. Canonical writes remain with the orchestrator.

## F2. Notifications

Notifications are operational heartbeats: league/date, verification mode,
canonical counts, fixtures/finals/gaps, and direct warehouse read availability.
Do not imply a betting recommendation when no certified pick ledger exists.

Deduplicate by message content hash. Keep only hashes in the send ledger. Never
log WhatsApp tokens, recipient details, or authorization headers. Fixture
reports may be previewed but must never trigger external delivery.

---

# PART G — LEGACY MIGRATION BOUNDARY

Google Sheets code may be used only for:

- reading historical workbook evidence;
- one-way migration into canonical records;
- explicit guarded compatibility export during a controlled transition.

It is never the default store and never selected merely because
`DATA_STORE_SPREADSHEET_ID` exists. Legacy access requires `--data-store-id` or
`--sheet-id`.

When the legacy writer is used, retain its existing safety properties:

- owned base columns A:R only;
- protected S:BX formulas/model columns;
- no whole-tab rewrite or `clearContents()`;
- date + home + away row identity;
- no overwrite of formulas or protected cells;
- dry-run diff and idempotent replay;
- one writer per migrated range.

These are compatibility constraints, not a reason to keep Sheets as the target
architecture.

---

# PART H — VERIFICATION AND DEFINITION OF DONE

## H1. Hermetic test expectations

Tests use committed fixtures and block sockets. Cover at minimum:

- canonical identity and ambiguous join quarantine;
- coherent bundle selection and orientation;
- actual book period/market coverage and no fabricated lines;
- pre/post prediction and book slots;
- frozen-write rejection and identical replay;
- DuckDB close/reopen durability and read-only dry runs;
- CSV.GZ content and SQL research views;
- immutable ledger no-op/conflict behavior;
- result reconciliation;
- research-only assay exclusions;
- daily end-to-end idempotency without Google;
- notification dry run and dedupe using a fake dispatcher;
- headless coverage and SQL reads without Google credentials.

No test may require an external database, real WhatsApp send, Google write, or
authenticated provider response.

## H2. Operational acceptance evidence

Before enabling live unattended writes, capture:

1. Offline suite result and critical lint result.
2. Fixture daily report labelled `FIXTURE_VERIFIED`.
3. Live dry-run report with actual provider coverage and gaps.
4. First canonical write report and database/CSV counts.
5. Immediate rerun proving no duplicates and unchanged ledgers.
6. Notification preview; only then an explicitly authorized send.
7. Cache restore/save evidence from the Actions run.

## H3. Completion statement requirements

A truthful completion statement distinguishes:

- implemented vs designed;
- fixture-verified vs live-verified;
- canonical DuckDB facts vs derived research/reports;
- research assay vs certified picks;
- dry-run vs actual external writes;
- legacy compatibility vs target architecture.

Never claim live provider coverage, production enablement, a WhatsApp delivery,
or certified betting performance unless that exact operation was performed and
recorded.
