# MA GOLIDE — Gold Universe Orchestrator v2

## Overview
Elite unified Python/Flask web orchestrator connecting to Google Sheets satellite spreadsheets.
Runs the Ma Assayer purity engine, the complete accuracy engine (all 7 bet types),
the elite Highest Quarter pipeline, and displays results in a live dashboard.

## Architecture

```
Satellites (Google Sheets)
    ↓ read_client (readonly scopes)
Fetcher → Side, Totals, ResultsClean, UpcomingClean, Bet_Slips
    ↓
Assay Engine (assayer_engine.py)   ← ISOLATED, never modified
    ↓
BetRecord (bet_record.py)          ← Universal contract for ALL pipelines
    ↓
Accuracy Engine (accuracy_engine.py) ← Reads Bet_Slips + ResultsClean ONLY
    ↓
HQ Pipeline (highest_quarter.py)   ← Softmax + Bayesian + Forebet blend
    ↓
Pipeline Writer (pipeline_writer.py) ← Central writer to Bet_Slips
    ↓ write_client (read/write scopes)
Bet_Slips sheet (Google Sheets)
    ↓
Dashboard (dashboard.html)         ← Live UI with all 7 bet type reports
```

## Key Files
```
app.py                          # Flask server — all API routes
auth/google_auth.py             # Read + Write Google Service Account clients
registry/satellite_registry.py # JSON-based satellite registry (CRUD)
registry/registry.json          # Auto-created: satellite data store
fetcher/sheet_fetcher.py        # Rate-limited fetcher + Bet_Slips/ResultsClean readers
assayer/assayer_engine.py       # Ma Assayer purity engine (ISOLATED — do not modify)
assayer/bet_record.py           # Universal BetRecord dataclass + to/from_bet_slip_row()
assayer/accuracy_engine.py      # MA GOLIDE COMPLETE ACCURACY REPORT — all 7 bet types
assayer/highest_quarter.py      # Elite HQ pipeline: softmax + Bayesian + Forebet blend
assayer/pipeline_writer.py      # Central pipeline writer with Coverage % tracking
templates/dashboard.html        # Dashboard UI with Report + Run HQ buttons
```

## BET_TYPES (all 7)
- BANKER — Moneyline high-confidence winner
- ROBBER — Underdog/upset moneyline pick
- SNIPER_OU — Quarter Over/Under totals
- SNIPER_MARGIN — Quarter spread/side bets
- FIRST_HALF_1X2 — First Half winner (1=home, X=draw, 2=away)
- FT_OU — Full-Time Over/Under totals
- HIGHEST_QUARTER — Highest scoring quarter prediction

## ORIGIN_CONFIGS (all 4)
- TIER1 — Config 1 only (deterministic)
- TIER2 — Config 2 only (deterministic)
- BLENDED — Combination of T1+T2
- LEGACY — Pre-v2 or untagged

## HQ Pipeline Thresholds
- STRONG ≥ 58% confidence → accepted
- MEDIUM ≥ 54% confidence → accepted
- Anything < 54% → auto-rejected

## Secrets Required
- `GOOGLE_SERVICE_ACCOUNT_JSON` — Full JSON of a Google Service Account key.
  Share satellites Drive folder with the service account email. Account needs
  spreadsheets write permission for HQ pipeline to write to Bet_Slips.

## API Endpoints
- `GET /` — Dashboard
- `GET /api/status` — Health + satellite count
- `GET /api/satellites` — List registry (filters: date, league, format)
- `POST /api/satellites/add` — Add one satellite
- `POST /api/satellites/bulk-add` — Bulk add from JSON array
- `DELETE /api/satellites/<id>` — Remove from registry
- `GET /api/satellites/<id>` — Get one satellite
- `POST /api/fetch/<id>` — Fetch one satellite's sheet data
- `POST /api/assay/<id>` — Fetch + run Ma Assayer purity engine
- `POST /api/fetch-all` — Background batch fetch all satellites
- `POST /api/assay-all` — Background batch assay all satellites
- `GET /api/job/<job_id>` — Poll background job status
- `POST /api/accuracy-report/<id>?origin_config=TIER1|TIER2|BLENDED|LEGACY` — MA GOLIDE COMPLETE ACCURACY REPORT
- `POST /api/run-hq/<id>` — Run HQ pipeline → write BetRecords to Bet_Slips
- `GET /api/bet-slips-count/<id>` — Count rows in Bet_Slips tab
- `POST /api/reset-auth` — Clear auth cache

## Design Decisions
- Rate limiting: 1.1s between all Google API calls
- Registry stored as registry/registry.json (portable, no DB needed)
- Background threading for bulk operations (non-blocking UI)
- Accuracy engine reads ONLY from Bet_Slips + ResultsClean — zero fallbacks
- Pipeline Writer tracks predictions_generated vs written for Coverage %
- Write client uses full spreadsheets+drive scopes; read client uses readonly scopes
- assayer_engine.py is isolated and must never be modified
