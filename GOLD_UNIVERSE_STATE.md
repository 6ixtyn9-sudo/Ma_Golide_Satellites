# Gold Universe — current state snapshot

**Updated:** 2026-11-07

**Status:** DuckDB/CSV-first architecture; no Google authentication required

This snapshot supersedes the April 2026 Sheets-orchestrator state. The former
satellite registry, Google setup instructions, and worksheet-writing routes are
legacy history, not the current target.

## Current directory map

```text
ingestion/
  models.py                    Canonical game/prediction/book/result contracts
  store.py                     Store interface + memory/JSON reference stores
  duckdb_store.py              Canonical persistent store and CSV.GZ exports
  warehouse.py                 Temporary SQL research views
  ledger.py                    Immutable content-addressed ledgers
  mining.py                    Research-only frozen-prediction assay
  pipeline.py                  Capture/freeze/results domain orchestration
  providers/                   Provider adapters
  sheets_*.py                  Explicit legacy compatibility only
scripts/
  daily.py                     One --auto-once autonomous orchestrator
  build_warehouse.py           Canonical CSV.GZ materialization
  mine_assay.py                Standalone research assay
  recover_warehouse.py         Guarded CSV.GZ-to-DuckDB recovery
  notify.py                    Deduplicated operational WhatsApp heartbeat
assayer/                       Reusable assay/bet-record/accuracy domain logic
localdata/                     Cached runtime state, ignored by Git
tests/                         Hermetic fixture/unit/integration suite
```

## Authority and operations

- Canonical: `localdata/warehouse.duckdb` and matching CSV.GZ exports.
- Orchestration: `scripts/daily.py --auto-once`.
- State: Actions restores/saves `localdata/` under `magolide-localdata-v1-`.
- Evidence: SHA-256-addressed pregame/freeze/result ledgers.
- Reads: coverage CLI, direct DuckDB SQL, CSV.GZ and assay reports; no web app.
- Notifications: operational heartbeat with content-hash dedupe.
- Legacy: explicit Sheets migration/compatibility commands only.

## Current assay logic

The existing `assayer/assayer_engine.py` still provides:

- Wilson lower bound with z=1.645;
- Bayesian shrinkage with alpha=2, beta=2;
- minimum segment n=10 and minimum absolute lift=0.03;
- BANKER when lower_bound >= 0.60 and win_rate >= 0.72, or lower_bound >=
  0.55 and win_rate >= 0.62;
- ROBBER for n < 10 or win_rate < 0.50;
- NEUTRAL otherwise.

The warehouse miner applies this only to frozen pre-tip moneyline predictions
joined to official reconciled non-tied finals. Outputs are always
`research_only`, `certified: false`. No official pick ledger is activated until
walk-forward validation, actual-price ROI, CLV, and decay gates exist.

## Provider integrity

Canonical identity, coherent provider bundles, actual book observations,
pre/post slots, frozen writes, and score reconciliation remain enforced. No
model-derived quarter lines, fabricated markets, or blended prediction bundles
are allowed. Missing provider keys and markets produce visible coverage states.

## Start and verification

```bash
python -m pytest tests -q
python scripts/daily.py --auto-once --offline-fixture
python -m ingestion.cli coverage --league NBA
```

Canonical persistence requires `ALLOW_MAGOLIDE_WRITES=true` and `--write`.
Fixture persistence requires its separate isolated opt-in. WhatsApp sends have
independent credentials and opt-ins. No live external write is implied by
repository tests.

Read `SESSION_START_INSTRUCTIONS.txt`, `docs/DUCKDB_ARCHITECTURE.md`, and
`docs/INGESTION_RUNBOOK.md` for the complete current contract.
