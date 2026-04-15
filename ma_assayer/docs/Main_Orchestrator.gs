/******************************************************************************
 * MA ASSAYER — Complete Production Module
 * 
 * "Testing the purity of predictions - separating Gold from Charcoal"
 * 
 * COMPLETE ROBUST IMPLEMENTATION
 * 
 * ARCHITECTURE:
 * ├── M1_Output_          Sheet Writers
 * ├── M2_Main_            Controller/Orchestrator
 * ├── M3_Flagger_         Apply Flags to Source
 * ├── M4_Discovery_       Edge Discovery Engine
 * ├── M5_Config_          Configuration & Constants
 * ├── M6_ColResolver_     Fuzzy Column Matching (case-insensitive)
 * ├── M7_Parser_          Data Parsing
 * ├── M8_Stats_           Statistical Calculations
 * ├── M9_Utils_           Utility Functions
 * └── M10_Log_            Logging System
 ******************************************************************************/
 
// ============================================================================
// MODULE: Main_ — Orchestrator
// ============================================================================

const Main_ = {
  
  /**
   * Run the full assay process
   */
  runAssay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Initialize
  Log_.init();
  ConfigLedger_Reader.init(/* optional: "SATELLITE_SPREADSHEET_ID" */);
  Config_.sheets.side = "Side"; // Ensure defaults
  Config_.sheets.totals = "Totals";
  
  try {
    // 2. Parse Data
    const sideData = Parser_.parseSideSheet(ss);
    const totalsData = Parser_.parseTotalsSheet(ss);
    
    const allBets = [...sideData.bets, ...totalsData.bets];
    
    // Apply 48-hour abandonment rule (Phase 5 Safety)
    allBets = applyAbandonmentRule_(allBets);
    
    if (allBets.length === 0) {
      Log_.error("No valid bets found in Side or Totals sheets.");
      Log_.writeToSheet(ss);
      return;
    }
    
    // 3. Statistical Analysis
    Log_.section("Running Statistics");
    const globalStats = Stats_.calcBasic(allBets);
    const sideStats = Stats_.calcBasic(sideData.bets);
    const totalsStats = Stats_.calcBasic(totalsData.bets);
    
    // 4. League Assay
    const leagueAssay = Stats_.assayLeagues(allBets, globalStats);

    // =========================================================
    // PATCH: Team assay + Matchup assay (Side only)
    // =========================================================
    const teamAssay = Stats_.assayTeams(sideData.bets, globalStats);
    const matchupAssay = Stats_.assayMatchups(sideData.bets, globalStats);
    
    // 5. Edge Discovery (Using patched Discovery_)
    const edges = Discovery_.discoverAll(sideData.bets, totalsData.bets);
    
    // 6. Exclusion Analysis
    const exclusionImpact = Stats_.calcExclusionImpact(allBets, globalStats);
    
    // 7. Write Outputs
    Output_.writeVault(ss, globalStats, sideStats, totalsStats, leagueAssay, edges, teamAssay, matchupAssay);
    Output_.writeLeagueAssay(ss, leagueAssay, globalStats);
    Output_.writeDiscoveredEdges(ss, edges, globalStats);
    Output_.writeExclusionImpact(ss, exclusionImpact, globalStats);

    // PATCH (optional): write Team + Matchup assay tabs if present
    if (Output_.writeTeamAssay) Output_.writeTeamAssay(ss, teamAssay);
    if (Output_.writeMatchupAssay) Output_.writeMatchupAssay(ss, matchupAssay);

    // ── MOTHER CONTRACT OUTPUT (additive only) ──
    Output_.writeAssayerEdges(ss, edges);
    Output_.writeAssayerLeaguePurity(ss, leagueAssay);

    // 8. Apply Flags Back to Source (PATCH: pass assays)
    Flagger_.applyFlags(ss, edges, leagueAssay, teamAssay, matchupAssay);
    
    // 9. Generate Summary
    const logSummary = Log_.summary();
    Output_.writeSummary(ss, globalStats, sideStats, totalsStats, leagueAssay, edges, logSummary);
    
    // 10. Finish
    Log_.writeToSheet(ss);
    SpreadsheetApp.getUi().alert("Assay Complete! Check the 'MA_Vault' tab.");
    
  } catch (err) {
    Log_.error(`CRITICAL FAILURE: ${err.message}`, err.stack);
    Log_.writeToSheet(ss);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
},
  
  /**
   * Run just the flagger (useful if manual edits made to edges)
   */
   runFlaggerOnly() {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   Log_.init();
   ConfigLedger_Reader.init(/* optional: "SATELLITE_SPREADSHEET_ID" */);
  
   try {
     const sideData = Parser_.parseSideSheet(ss);
     const totalsData = Parser_.parseTotalsSheet(ss);
     const allBets = [...sideData.bets, ...totalsData.bets];
    
     // Re-run minimal stats needed for flagging
     const globalStats = Stats_.calcBasic(allBets);
     const leagueAssay = Stats_.assayLeagues(allBets, globalStats);

     // PATCH: team + matchup assays for flagger-only mode
     const teamAssay = Stats_.assayTeams(sideData.bets, globalStats);
     const matchupAssay = Stats_.assayMatchups(sideData.bets, globalStats);

     const edges = Discovery_.discoverAll(sideData.bets, totalsData.bets);
    
     // PATCH: pass assays
     Flagger_.applyFlags(ss, edges, leagueAssay, teamAssay, matchupAssay);

     // ── MOTHER CONTRACT OUTPUT (additive only) ──
     Output_.writeAssayerEdges(ss, edges);
     Output_.writeAssayerLeaguePurity(ss, leagueAssay);

     Log_.writeToSheet(ss);
     SpreadsheetApp.getUi().alert("Flags Re-Applied Successfully.");
    
   } catch (err) {
     Log_.error(`FLAGGER FAILED: ${err.message}`);
     Log_.writeToSheet(ss);
   }
 }
};

/**
 * setupAssayerSheets - Create required sheets for Assayer operation
 */
function setupAssayerSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const requiredSheets = [
    // Core betting sheets (23-column contract compliance)
    { name: 'Bet_Slips', headers: ['bet_id', 'league', 'event_date', 'event_time', 'match', 'team', 'opponent', 'side_total', 'line', 'odds', 'implied_prob', 'confidence_pct', 'tier_code', 'tier_display', 'ev', 'kelly_pct', 'status', 'result', 'payout', 'placed_at', 'settled_at', 'config_stamp', 'source', 'gender', 'quarter', 'season', 'created_at'] },
    { name: 'ResultsClean', headers: ['result_id', 'event_date', 'league', 'team', 'opponent', 'side_total', 'line', 'actual_result', 'settled_at', 'status', 'payout', 'config_stamp', 'source', 'season', 'quarter', 'created_at'] },
    
    // Prediction logs (17-column forensic logs)
    { name: 'Tier1_Predictions', headers: ['log_id', 'timestamp', 'league', 'event_id', 'team', 'opponent', 'side_total', 'line', 'prediction', 'confidence', 'tier', 'ev', 'status', 'result', 'config_stamp', 'source', 'notes'] },
    { name: 'Tier2_Log', headers: ['log_id', 'timestamp', 'league', 'event_id', 'team', 'opponent', 'side_total', 'line', 'prediction', 'confidence', 'tier', 'ev', 'status', 'result', 'config_stamp', 'source', 'notes'] },
    { name: 'OU_Log', headers: ['log_id', 'timestamp', 'league', 'event_id', 'team', 'opponent', 'side_total', 'line', 'prediction', 'confidence', 'tier', 'ev', 'status', 'result', 'config_stamp', 'source', 'notes'] },
    
    // Accuracy and reporting sheets
    { name: 'Accuracy_Report', headers: ['Generated', 'Total_Bets_Graded', 'Total_Hits', 'Total_Misses', 'Overall_Hit_Rate'] },
    { name: 'Tier2_Accuracy', headers: ['Metric', 'Value'] },
    { name: 'OU_Accuracy', headers: ['Metric', 'Value'] },
    
    // Configuration sheets (Config_Ledger system)
    { name: 'Config_Ledger', headers: ['config_key', 'config_value', 'description', 'last_updated', 'dominant_stamp', 'stamp_purity'] },
    { name: 'Config_Tier1', headers: ['config_key', 'config_value', 'description', 'last_updated'] },
    { name: 'Config_Tier2', headers: ['config_key', 'config_value', 'description', 'last_updated'] },
    { name: 'Config_Accumulator', headers: ['config_key', 'config_value', 'description', 'last_updated'] },
    
    // Satellite management
    { name: 'Satellite_Identity', headers: ['satellite_id', 'spreadsheet_url', 'satellite_name', 'status', 'last_sync', 'config_version', 'notes'] },
    
    // Analysis and assay sheets
    { name: 'Assayer_Log', headers: ['Timestamp', 'Level', 'Message'] },
    { name: 'League_Assay', headers: ['League', 'Total_Bets', 'Win_Rate', 'Avg_Odds', 'Purity', 'Grade'] },
    { name: 'Team_Assay', headers: ['Team', 'Total_Bets', 'Win_Rate', 'Avg_Odds', 'Purity', 'Grade'] },
    { name: 'Edges', headers: ['Date', 'League', 'Home', 'Away', 'Pick', 'Edge_Type', 'Edge_Value', 'Confidence'] },
    
    // Stats and performance sheets
    { name: 'Stats', headers: ['Metric', 'Value', 'Description'] },
    { name: 'Standings', headers: ['Team', 'League', 'Played', 'Won', 'Lost', 'Points'] },
    { name: 'Sheet_Inventory', headers: ['sheet_name', 'sheet_type', 'row_count', 'last_updated', 'status'] },
    
    // Raw data sheets
    { name: 'ResultsRaw', headers: ['Date', 'League', 'Home', 'Away', 'Score', 'Result'] },
    { name: 'Raw', headers: ['Date', 'League', 'Home', 'Away', 'Pick', 'Type', 'Odds', 'Result'] },
    
    // Tier2 analysis sheets
    { name: 'TeamQuarterStats_Tier2', headers: ['Team', 'League', 'Quarter', 'Games', 'Wins', 'Losses', 'Win_Rate'] },
    { name: 'LeagueQuarterO_U_Stats', headers: ['League', 'Quarter', 'Over_Hits', 'Over_Misses', 'Under_Hits', 'Under_Misses', 'Total_Games'] },
    
    // Legacy sheets (backward compatibility)
    { name: 'Side', headers: ['Date', 'League', 'Home', 'Away', 'Pick', 'Side', 'Odds', 'Result', 'Outcome', 'Notes'] },
    { name: 'Totals', headers: ['Date', 'League', 'Home', 'Away', 'Pick', 'Line', 'Odds', 'Result', 'Outcome', 'Notes'] },
    
    // Upcoming and analysis sheets
    { name: 'UpcomingClean', headers: ['Date', 'League', 'Home', 'Away', 'Pick', 'Type', 'Odds', 'Status'] },
    { name: 'UpcomingRaw', headers: ['Date', 'League', 'Home', 'Away', 'Time', 'Status'] }
  ];
  
  requiredSheets.forEach(({ name, headers }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
            .setFontWeight('bold')
            .setBackground('#f0f0f0');
      Logger.log(`Created sheet: ${name}`);
    } else {
      Logger.log(`Sheet already exists: ${name}`);
    }
  });
  
  SpreadsheetApp.getUi().alert('Assayer sheets setup complete!');
  Logger.log('Assayer sheets setup completed');
}

/**
 * Standard Apps Script Entry Points
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Assayer')
    .addItem('Setup Sheets', 'setupAssayerSheets')
    .addSeparator()
    .addItem('Run Full Assay', 'runAssay')
    .addSeparator()
    .addItem('Refresh Flags Only', 'runFlagger')
    .addItem('Clear Logs', 'clearLogs')
    .addToUi();
}

function runAssay() { Main_.runAssay(); }
function runFlagger() { Main_.runFlaggerOnly(); }
function clearLogs() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(Config_.sheets.logs);
  if (sheet) sheet.clear();
}
