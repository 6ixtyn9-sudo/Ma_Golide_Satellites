# Headless architecture

## Data flow

```text
Public schedule/result providers       Optional prediction providers
ESPN / EuroLeague                      ESPN BPI / Hardwood / Predixsport
                \                              /
                 \                            /
                  ------ canonical adapters --
                                |
Optional book adapters ----------+------------ identity and UTC validation
TheRundown / SportsGameOdds /    |
TheOddsAPI                       v
                    ingestion.pipeline.Pipeline
                                |
                    +-----------+------------+
                    |                        |
             ephemeral dry run        deliberate persistence
                    |                        |
                    |              localdata/warehouse.duckdb
                    |                        |
                    |              +---------+----------+
                    |              |                    |
                    |       CSV.GZ exports       immutable stage ledgers
                    |              |                    |
                    +------ derived operational reports |
                                   |
                         immutable pre-tip freeze
                                   |
                    +--------------+---------------+
                    |                              |
          walk-forward market service       descriptive mining assay
                    |                              |
          market forecasts + accas          prediction-outcome research
                    |
          optional guarded operational heartbeat
```

No component requires a browser application, spreadsheet, external database,
or JSON-file store. The only long-lived sports state is DuckDB plus CSV.GZ and
immutable audit ledgers.

## Package boundaries

- `ingestion/providers/`: source-specific HTTP parsing and identity
  reconciliation. Adapters return canonical models, never report rows.
- `ingestion/models.py`: strict games, results, predictions, book snapshots,
  periods, markets, coverage gaps, and health records.
- `ingestion/pipeline.py`: pregame acquisition, freeze, result reconciliation,
  coverage, and provider smoke operations.
- `ingestion/store.py`: in-memory behavior, freeze enforcement, and dry-run
  semantics shared by storage implementations.
- `ingestion/duckdb_store.py`: canonical persistence, schema migration,
  CSV.GZ export, and CSV recovery.
- `ingestion/warehouse.py`: independent read-only DuckDB connections and
  temporary research views.
- `forecasting/model.py`: history-only empirical market distributions.
- `forecasting/harness.py`: prequential reliability profiles and eligibility.
- `forecasting/accas.py`: exact best-first ranking under the documented
  cross-game independence assumption.
- `forecasting/service.py`: frozen historical replays, current reports, and
  atomic derived JSON writes.
- `assayer/`: descriptive prediction-outcome statistics; it does not create or
  certify wagers.
- `scripts/daily.py`: the only scheduled orchestrator.

## DuckDB schema

The canonical tables are:

| Table | Purpose |
| --- | --- |
| `games` | Identity, schedule, status, and official result payload |
| `predictions` | One pre/post observation slot per game/provider |
| `book_snapshots` | One pre/post slot per game/provider/book/market/period |
| `team_aliases` | Explicit normalized team-name aliases |
| `provider_health` | Last attempt/success and structured coverage |
| `pipeline_runs` | Per-run/per-stage operational reports |

`book_snapshots` includes typed moneyline home/away/draw prices, total
over/under prices, spread line/prices, period, source timestamps, pre-match and
main-line flags, payload hash, and freeze marker. Schema version 2 adds and
backfills `draw_price_decimal` for first-half 1X2 evidence.

Every row retains its complete `payload_json` (or canonical equivalent) for
lossless export/recovery. Typed columns make SQL convenient; they are not a
second independently edited data source.

### Freeze enforcement

The in-memory implementation defines the behavior and DuckDB delegates to it
before writing. Frozen prediction and book keys are held separately. An exact
replay is a no-op; a changed value for a frozen key raises `FrozenWriteError`.
Post-tip observations occupy separate slots and never overwrite the pre-tip
freeze.

Historical forecasting calls `is_prediction_frozen` and
`is_book_snapshot_frozen`, rather than trusting a boolean payload alone.
Current reports use visible pre-tip observations and expose their freeze state;
the subsequent freeze preserves tip-passed rows for later replay.

## CSV.GZ and recovery

`DuckDBStore.export_csv` writes all six canonical tables as compressed CSV.
Each file includes the complete payload needed to reconstruct typed models.
Export is a materialization, not a parallel mutable store.

`scripts/recover_warehouse.py` restores into a temporary new DuckDB path and
atomically renames it only after all tables load. It refuses an existing
warehouse or temporary destination and requires the normal write opt-in.
Restoring into a nonempty warehouse is rejected, preventing accidental merges
of unrelated evidence.

## Derived research views

`ingestion.warehouse.connect` opens DuckDB read-only by default and creates
temporary views:

- `prediction_facts`
- `book_facts`
- `latest_pregame_predictions`
- `game_results`

A missing warehouse is an error rather than an empty successful slate. These
views expose observed facts and disappear on close. They never mutate
canonical state.

## Forecast service

`MarketForecastService.build_report` performs two distinct passes:

1. **Historical replay:** chronological reconciled finals, frozen pre-tip
   evidence, per-tip availability cutoffs, prequential calibration, grading,
   and segment metrics.
2. **Current slate:** scheduled games whose local date matches the requested
   slate and whose tip is after `as_of`, using only evidence visible at that
   timestamp.

The report contains capabilities, historical and calibrated metrics, segment
profiles, all generated market forecasts, uncovered games, eligibility
reasons, and ranked accas. It uses schema `magolide.market_forecasts.v1` and
explicitly disclaims guaranteed/certified outcomes and external betting
writes.

`write_market_report` writes a temporary sibling and atomically replaces the
derived JSON report. Identical bytes are left untouched. This atomic report
mechanism does not make reports canonical evidence.

## Orchestration order

`scripts/daily.py --auto-once` performs:

1. `pregame` acquisition and canonical stage ledger when writes are enabled;
2. current market/accuracy/acca report generation from visible pre-tip facts;
3. immutable `freeze` of tip-passed evidence and its canonical stage ledger;
4. `results` reconciliation unless `--skip-results` is set;
5. bounded operational-run/report retention;
6. descriptive prediction assay, CSV exports, and final operational report;
7. optional independently guarded notification.

The same `run_id` identifies every stage. Immutable ledgers use content hashes,
so identical replays are no-ops and conflicting content receives a new path or
an explicit conflict rather than overwriting history.

## Deployment state

GitHub Actions has one three-hour schedule. It restores/saves `localdata/` as a
cache and uploads run reports/ledgers as artifacts. A disabled scheduled write
uses the offline fixture path to avoid consuming authenticated provider quota
without persisting state. Manual live dry runs remain available for coverage
checks.

A CI cache is an operational optimization, not an archival guarantee. Export
and retain CSV.GZ/ledger artifacts according to the desired recovery policy.
Runtime state and credentials are excluded from Git.
