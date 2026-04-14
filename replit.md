# MaGolide Betting System

Advanced sports betting prediction system built with Google Apps Script modules.

## Project Overview

This project is a suite of Google Apps Script (GAS) modules designed to run inside Google Sheets. It provides a complete sports betting prediction pipeline — from data ingestion and parsing to probability modeling and bet slip generation.

## Architecture

- **Runtime:** Google Apps Script (JavaScript) — designed to run inside Google Sheets
- **Replit Web Viewer:** A lightweight Node.js HTTP server (`server.js`) serves a code browser at port 5000
- **No build step required** — pure Node.js standard library, no dependencies

## Script Modules (in `docs/`)

| File | Purpose |
|------|---------|
| `Sheet_Setup.gs` | Module 1 — One-click sheet infrastructure setup, accuracy reporting |
| `Config_Ledger_Satellite.gs` | Configuration versioning and audit trail |
| `Contract_Enforcer.gs` | Data validation and integrity checks |
| `Contract_Enforcement.gs` | Structural contract enforcement |
| `Data_Parser.gs` | Parses raw scores/data into structured objects |
| `Game_Processor.gs` | Module 7 — Core probability models and bet picking |
| `Forecaster.gs` | Generates forecasts/predictions from model output |
| `Signal_Processor.gs` | Betting signal generation |
| `Accumulator_Builder.gs` | Aggregates picks into accumulator/parlay slips |
| `Inventory_Manager.gs` | Tracks available games inventory |
| `Margin_Analyzer.gs` | Spread/margin betting analysis |

## How to Deploy to Google Sheets

1. Open a Google Sheet
2. Go to Extensions > Apps Script
3. Paste each `.gs` file as a separate script file
4. Run `setupAllSheets()` to initialize the spreadsheet tabs
5. Load game data into the `Results` sheet
6. Run `runTier2_BothModes()` to generate predictions

## Development

The Replit workflow runs `node server.js` which starts a code viewer at port 5000.
