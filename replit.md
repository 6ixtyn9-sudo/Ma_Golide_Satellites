# MaGolide Betting System

## Overview
Advanced sports betting prediction and analysis platform built as Google Apps Script (GAS) modules designed to run inside Google Sheets.

## Project Structure
- `docs/` — All Google Apps Script (`.gs`) module files
- `server.py` — Simple Python HTTP server serving a script viewer UI on port 5000

## Running the App
The app runs `python server.py` which starts a web viewer on port 5000.
The viewer lets users browse all GAS modules and copy them to deploy in Google Sheets.

## Modules
| File | Role | Description |
|------|------|-------------|
| Sheet_Setup.gs | Module 1 - The Architect | Sheet infrastructure setup, Single Source of Truth |
| Signal_Processor.gs | Module 3 - The Mouth | Normalizes raw FlashScore/Sofascore data |
| Data_Parser.gs | Module 2 - The Parser | Parses raw incoming data |
| Forecaster.gs | Module 5 - The Brain | Tier 1 predictions with historical context |
| Game_Processor.gs | Module 7 - The Engine | Tier 2 probability models & confidence scoring |
| Accumulator_Builder.gs | Module 6 - The Stacker | Builds parlays and accumulators |
| Inventory_Manager.gs | Module 4 - The Librarian | Historical data and state management |
| Margin_Analyzer.gs | Module 8 - The Analyst | Spreads, Moneylines, Totals analysis |
| Contract_Enforcer.gs | Module 9 - The Validator | Data integrity validation |
| Contract_Enforcement.gs | Module 9b - The Guard | Contract enforcement rules |
| Config_Ledger_Satellite.gs | Config - The Registry | System-wide settings and constants |

## Deployment to Google Sheets
1. Open a Google Spreadsheet → Extensions → Apps Script
2. Create a script file per module (use filename as script name)
3. Paste contents from each module
4. Run `setupAllSheets()` from Sheet_Setup.gs first
5. Follow the pipeline order for subsequent runs

## Tech Stack
- Language: JavaScript (Google Apps Script)
- Platform: Google Workspace (Sheets)
- Web Viewer: Python (stdlib HTTPServer, no dependencies)
- Port: 5000
