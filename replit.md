# Ma Golide — Gold Universe

## Overview
The Ma Golide Gold Universe is a suite of three interconnected Google Apps Script projects for sports betting prediction, purity testing, and accumulator building. All scripts run inside **Google Sheets** — this is not a traditional web application.

## Architecture
```
Satellites (Ma Golide) → Ma Assayer → Mothership
   league predictions     purity test   builds accas
```

## Projects

### 1. Ma Golide Satellites (`docs/`)
League-level prediction machines. Each satellite spreadsheet runs these 11 modules to process raw sports data and generate predictions.
- v1.1.0 Diagnostic Edition
- Modules: Sheet Setup, Config, Signal Processor, Data Parser, Margin Analyzer, Forecaster, Game Processor, Inventory Manager, Accumulator Builder, Contract Enforcer, Contract Enforcement

### 2. Ma Assayer (`ma_assayer/docs/`)
The purity engine — "separating Gold from Charcoal". Tests prediction quality and produces the official purity contract.
- v4.3.0 — Type-Segmented Totals
- 10 modules: Output Writers, Main Orchestrator, Flagger, Discovery Edge, Config, ColResolver (fuzzy matching), Parser, Stats, Utils, Log
- Outputs: `ASSAYER_EDGES` + `ASSAYER_LEAGUE_PURITY` sheets
- Grades: Platinum (≥85%) → Gold → Silver → Bronze → Rock → Charcoal

### 3. Ma Golide Mothership (`ma_golide_mothership/doc/`)
The central brain. Reads all satellites + Assayer purity contract → builds Accas, Risky Accas, and Bet Slips.
- 11 modules: MIC Intelligence Core, Acca Engine, Assayer Bridge, Genesis, Hive Mind, Menu, Risky Acca Builder, Risky Analyzer, Performance Analyzer, Systems Audit, Config Ledger

## Replit Setup
Since all three are Google Apps Script projects (no web server), a static HTML documentation viewer (`index.html`) is served via Python's HTTP server on port 5000.

- **Workflow:** `Start application` — `python3 -m http.server 5000 --bind 0.0.0.0`
- **Deployment:** Configured as static site serving the root directory

## How to Deploy Scripts to Google Sheets
1. Open a Google Sheet → **Extensions → Apps Script**
2. Copy each `.gs` file from the relevant folder into the editor
3. Run the genesis/setup module first to initialise the sheet structure
4. Or use Clasp: `clasp push` from the project directory

## Project Layout
```
docs/                             # Ma Golide Satellite scripts (11 files)
ma_assayer/docs/                  # Ma Assayer scripts (11 files)
ma_golide_mothership/doc/         # Mothership scripts (11 files)
index.html                        # Documentation viewer (served on port 5000)
```
