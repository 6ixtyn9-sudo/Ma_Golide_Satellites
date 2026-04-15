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
// MODULE: Config_ — Configuration & Constants (v4.3.0 — Type-Segmented Totals)
// ============================================================================

const Config_ = {
  version: "4.3.0",                    // ◆ PATCH: version bump
  name: "Ma Assayer",
  buildDate: "2025-06-28",             // ◆ PATCH: build date

  // Sheet names
  sheets: {
    side: "Side",
    totals: "Totals",
    vault: "MA_Vault",
    discovery: "MA_Discovery",
    leagueAssay: "MA_LeagueAssay",
    exclusion: "MA_Exclusion",
    config: "MA_Config",
    charts: "MA_Charts",
    logs: "MA_Logs",
    quarterAnalysis: "MA_QuarterAnalysis",
    summary: "MA_Summary",
    teamAssay: "MA_TeamAssay",
    matchupAssay: "MA_MatchupAssay",
    assayerEdges: "ASSAYER_EDGES",
    assayerLeaguePurity: "ASSAYER_LEAGUE_PURITY"
  },

  // Mother Contract — Output Sheets + Schema (additive only)
  // Satellite machine contracts (Ma_Golide_Satellites — Contract_Enforcer)
  satelliteContract: {
    forensicCore17: [
      "Prediction_Record_ID", "Universal_Game_ID", "Config_Version", "Timestamp_UTC",
      "League", "Date", "Home", "Away", "Market", "Period", "Pick_Code", "Pick_Text",
      "Confidence_Pct", "Confidence_Prob", "Tier_Code", "EV", "Edge_Score"
    ],
    betSlips23: [
      "Bet_Record_ID", "Universal_Game_ID", "Source_Prediction_Record_ID",
      "League", "Date", "Home", "Away", "Market", "Period", "Selection_Side", "Selection_Line",
      "Selection_Team", "Selection_Text", "Odds", "Confidence_Pct", "Confidence_Prob", "EV",
      "Tier_Code", "Tier_Display", "Config_Version_T1", "Config_Version_T2", "Config_Version_Acc", "Source_Module"
    ]
  },

  motherContract: {
    EDGE_SHEET_NAME: "ASSAYER_EDGES",
    LEAGUE_PURITY_SHEET: "ASSAYER_LEAGUE_PURITY",

    EDGE_COLUMNS: [
      "edge_id", "source", "pattern", "discovered", "updated_at",
      "quarter", "is_women", "tier", "side", "direction",
      "conf_bucket", "spread_bucket", "line_bucket",
      "type_key",                                                // ◆ PATCH: added
      "filters_json",
      "n", "wins", "losses", "win_rate", "lower_bound", "upper_bound", "lift",
      "grade", "symbol", "reliable", "sample_size"
    ],

    LEAGUE_COLUMNS: [
      "league", "quarter", "source", "gender", "tier", "type_key",
      "n", "win_rate", "grade", "status",
      "dominant_stamp", "stamp_purity",
      "updated_at"
    ]
  },

  // ◆ PATCH: Canonical totals type keys (derived in Discovery_._getTotalsTypeKey)
  // Reference only — the normalization logic lives in Discovery_.
  totalsTypeKeys: [
    "SNIPER_OU",        // Plain Sniper O/U
    "SNIPER_OU_DIR",    // Sniper O/U DIR
    "SNIPER_OU_STAR",   // Sniper O/U STAR
    "OU",               // Generic O/U (no "Sniper" prefix)
    "OU_DIR",           // Generic O/U DIR
    "OU_STAR",          // Generic O/U STAR
    "OTHER",            // Recognized type that doesn't match O/U patterns
    "UNKNOWN"           // Missing / empty type
  ],

  // Column aliases
  sideColumnAliases: {
    league: [
      "league", "lg", "lge", "leag", "leauge", "legue", "comp", "competition",
      "sport league", "sportleague", "conference", "division", "tour"
    ],
    date: [
      "date", "dt", "dte", "game date", "gamedate", "match date", "event date",
      "play date", "playdate", "bet date", "betdate", "event"
    ],
    time: [
      "time", "tm", "start time", "starttime", "game time", "kickoff",
      "tip off", "tipoff", "start", "event time"
    ],
    match: [
      "match", "game", "matchup", "teams", "mch", "mtch", "fixture", "vs",
      "event", "contest", "bout", "meeting", "pairing", "matchup teams"
    ],
    pick: [
      "pick", "selection", "bet", "play", "pck", "wager", "side pick",
      "team pick", "teampick", "chosen", "choice", "prediction", "pred"
    ],
    type: [
      "type", "typ", "bet type", "bettype", "category", "market type",
      "market", "bet market", "wager type", "play type"
    ],
    confidence: [
      "confidence", "conf", "conf%", "confpct", "confidence%", "cnf", "cnf%",
      "confidence_pct", "confidence pct",
      "prob", "probability", "likelihood", "certainty", "edge%", "model conf",
      "model confidence", "predicted prob", "win prob", "win probability"
    ],
    tier: [
      "tier", "tr", "tier level", "strength", "tierlevel", "rating",
      "tier_code", "tier display", "tier_display",
      "grade tier", "quality", "star", "stars", "rank", "level", "class"
    ],
    quarter: [
      "quarter", "qtr", "q", "qrtr", "period", "half", "quater", "quartr",
      "per", "prd", "segment", "section", "part", "phase"
    ],
    actual: [
      "actual", "act", "result score", "score", "actl",
      "actual score", "final score", "actual result", "real score"
    ],
    side: [
      "side", "sd", "h/a", "home away", "homeaway", "team side", "home/away",
      "location", "venue", "home or away", "h or a"
    ],
    outcome: [
      "outcome", "result", "res", "win/loss", "winloss", "w/l", "hit",
      "otcome", "outcom", "status", "graded result", "grade", "graded",
      "final result", "bet result", "wager result", "decision", "verdict"
    ],
    odds: [
      "odds", "price", "line odds", "decimal odds", "american odds",
      "moneyline", "ml", "payout", "juice", "vig"
    ],
    units: [
      "units", "unit", "stake", "bet size", "betsize", "wager size",
      "risk", "amount", "size"
    ],
    ev: [
      "ev", "ev%", "expected value", "expectedvalue", "edge", "value",
      "expected", "roi", "return"
    ],
    notes: [
      "notes", "note", "comments", "comment", "memo", "remarks", "info"
    ],
    home: [
      "home", "hm", "home team", "hometeam", "h team", "team 1", "team1",
      "host", "home side", "homeside"
    ],
    away: [
      "away", "aw", "away team", "awayteam", "a team", "visitor", "team 2",
      "team2", "visiting", "road", "road team", "roadteam", "guest"
    ],
    config_stamp: [
      "config_stamp", "configstamp", "cfg_stamp", "stamp", "stamp_id"
    ]
  },

  totalsColumnAliases: {
    date: [
      "date", "dt", "dte", "game date", "gamedate", "event date",
      "play date", "playdate", "bet date", "match date"
    ],
    league: [
      "league", "lg", "lge", "leag", "leauge", "comp", "competition",
      "sport league", "conference", "division"
    ],
    home: [
      "home", "hm", "home team", "hometeam", "h team", "team 1", "team1",
      "host", "home side", "homeside"
    ],
    away: [
      "away", "aw", "away team", "awayteam", "a team", "visitor", "team 2",
      "team2", "visiting", "road", "road team", "roadteam", "guest"
    ],
    match: [
      "match", "game", "matchup", "teams", "fixture", "vs", "event",
      "contest", "pairing"
    ],
    quarter: [
      "quarter", "qtr", "q", "qrtr", "period", "quater", "quartr",
      "per", "prd", "segment", "half"
    ],
    direction: [
      "direction", "dir", "over/under", "overunder", "o/u", "ou", "bet dir",
      "over under", "over or under", "o or u", "side", "pick direction"
    ],
    line: [
      "line", "ln", "total", "total line", "number", "lne", "points",
      "closing line", "game total", "gametotal", "projected total",
      "total points", "totalpoints", "target", "mark"
    ],
    actual: [
      "actual", "act", "final", "score", "actl", "actual total",
      "final total", "real total", "combined score", "combinedscore",
      "total score", "totalscore", "actual score"
    ],
    result: [
      "result", "res", "outcome", "win/loss", "winloss", "hit", "rslt",
      "status", "graded result", "grade", "graded", "final result",
      "bet result", "decision", "w/l"
    ],
    diff: [
      "diff", "difference", "margin", "dif", "dfference", "delta",
      "variance", "spread", "gap", "deviation"
    ],
    confidence: [
      "confidence", "conf", "conf%", "confpct", "cnf", "cnf%", "prob",
      "probability", "likelihood", "certainty", "model conf", "win prob"
    ],
    ev: [
      "ev", "ev%", "evpct", "expected value", "expectedvalue", "edge",
      "value", "expected", "roi"
    ],
    tier: [
      "tier", "tr", "tier level", "strength", "rating", "grade tier",
      "quality", "star", "rank", "level"
    ],
    type: [
      "type", "typ", "bet type", "bettype", "market type", "market",
      "category", "wager type"
    ],
    odds: [
      "odds", "price", "line odds", "decimal odds", "juice", "vig"
    ],
    units: [
      "units", "unit", "stake", "bet size", "risk", "amount"
    ],
    notes: [
      "notes", "note", "comments", "comment", "memo", "remarks"
    ],
    config_stamp: [
      "config_stamp", "configstamp", "cfg_stamp", "stamp", "stamp_id"
    ]
  },

  // Statistical thresholds
  thresholds: {
    minN: 10,
    minNReliable: 30,
    minNPlatinum: 50,
    minNGold: 25,
    wilsonZ: 1.645,
    wilsonZ90: 1.645,
    wilsonZ95: 1.96,
    liftThreshold: 0.03,
    minEdgeLift: 0.05,
    maxEdgeLift: 0.25,
    minSimilarity: 0.65,
    highSimilarity: 0.85,
    minNTeam: 25,
    minNTeamReliable: 40,
    minNTeamGold: 40,
    minNTeamPlatinum: 60,
    minNMatchup: 5,
    minNMatchupReliable: 30,
    minNMatchupGold: 30,
    minNMatchupPlatinum: 45,
    wilsonLowerBoundGate: 0                 
  },

// Purity grades
  grades: {
    PLATINUM: { min: 0.85, symbol: "⬡",  name: "Platinum", color: "#E5E4E2", bgColor: "#1a1a2e" },
    GOLD:     { min: 0.72, symbol: "Au",  name: "Gold",     color: "#FFD700", bgColor: "#2d2d0d" },
    SILVER:   { min: 0.62, symbol: "Ag",  name: "Silver",   color: "#C0C0C0", bgColor: "#2d2d2d" },
    BRONZE:   { min: 0.55, symbol: "Cu",  name: "Bronze",   color: "#CD7F32", bgColor: "#2d1f0d" },
    ROCK:     { min: 0.50, symbol: "ite", name: "Rock",     color: "#808080", bgColor: "#1a1a1a" },
    CHARCOAL: { min: 0.00, symbol: "🜃",  name: "Charcoal", color: "#363636", bgColor: "#0d0d0d" }
  },

  toxicLeagues: ["UNKNOWN"],
  eliteLeagues: ["UNKNOWN"],
  toxicTeams: [],
  eliteTeams: [],
  toxicMatchups: [],
  eliteMatchups: [],
  teamAliases: {},

  // Spread buckets
  spreadBuckets: [
    { name: "<3",      min: 0,    max: 2.99, label: "Tight (<3)" },
    { name: "3-4",     min: 3,    max: 4,    label: "Close (3-4)" },
    { name: "4.5-5.5", min: 4.5,  max: 5.5,  label: "Medium (4.5-5.5)" },
    { name: "5.5-6",   min: 5.5,  max: 6,    label: "Standard (5.5-6)" },
    { name: "6-7",     min: 6,    max: 7,    label: "Wide (6-7)" },
    { name: ">7",      min: 7.01, max: 100,  label: "Blowout (>7)" }
  ],

  // Line / total buckets
  lineBuckets: [
    { name: "<35",   min: 0,     max: 34.99, label: "Very Low (<35)" },
    { name: "35-40", min: 35,    max: 40,    label: "Low (35-40)" },
    { name: "40-50", min: 40.01, max: 50,    label: "Below Avg (40-50)" },
    { name: "50-60", min: 50.01, max: 60,    label: "Average (50-60)" },
    { name: "60-70", min: 60.01, max: 70,    label: "Above Avg (60-70)" },
    { name: ">70",   min: 70.01, max: 200,   label: "High (>70)" }
  ],

  // Confidence buckets
  confBuckets: [
    { name: "<55%",   min: 0,     max: 0.549, label: "Low (<55%)" },
    { name: "55-60%", min: 0.55,  max: 0.60,  label: "Moderate (55-60%)" },
    { name: "60-65%", min: 0.601, max: 0.65,  label: "Good (60-65%)" },
    { name: "65-70%", min: 0.651, max: 0.70,  label: "Strong (65-70%)" },
    { name: "≥70%",   min: 0.701, max: 1.0,   label: "Elite (≥70%)" }
  ],

  // Tier mappings
  tierMappings: {
    strong: ["strong", "★", "★★★", "high", "s", "3", "elite", "top", "a", "best"],
    medium: ["medium", "●", "★★", "med", "m", "2", "standard", "avg", "b", "mid"],
    weak:   ["weak", "○", "★", "low", "w", "1", "minimal", "c", "bottom", "low"]
  },

  // Outcome mappings
  outcomeMappings: {
    win:  ["✅", "hit", "w", "win", "1", "won", "winner", "y", "yes", "correct", "right", "covered", "cashed"],
    loss: ["❌", "miss", "l", "loss", "0", "lost", "loser", "n", "no", "incorrect", "wrong", "failed"],
    push: ["even", "p", "push", "tie", "draw", "e", "void", "cancel", "cancelled", "refund", "no action", "na"]
  },

  // Colors
  colors: {
    header:     "#1a1a2e",
    headerText: "#FFD700",
    gold:       "#FFD700",
    silver:     "#C0C0C0",
    bronze:     "#CD7F32",
    platinum:   "#E5E4E2",
    success:    "#28a745",
    warning:    "#ffc107",
    danger:     "#dc3545",
    info:       "#17a2b8",
    dark:       "#343a40",
    light:      "#f8f9fa"
  },

  // Report settings
  report: {
    maxEdgesToShow:   20,
    maxLeaguesToShow: 15,
    maxToxicToShow:   10,
    dateFormat:       "yyyy-MM-dd HH:mm:ss"
  }
};

// ============================================================================
// PHASE 3 PATCH 5 + 5B: CONFIG HARDENING - ASSAYER INTEGRATION
// ============================================================================

/**
 * ConfigManager_Assayer - Assayer-specific configuration management
 * Integrates with Satellite Config Managers for state lineage
 */
const ConfigManager_Assayer = {
  
  // --------------------------------------------------------------------------
  // loadAssayerConfig - Load Assayer configuration with Config Ledger integration
  // --------------------------------------------------------------------------
  loadAssayerConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Load base configuration from Config_ object
    const baseConfig = JSON.parse(JSON.stringify(Config_));
    
    // Try to load enhanced config from satellite config sheets
    try {
      const satelliteConfig = this.loadSatelliteConfigs(ss);
      if (satelliteConfig) {
        Object.assign(baseConfig, satelliteConfig);
      }
    } catch (err) {
      Logger.log('[ConfigManager_Assayer] Satellite config load failed: ' + err.message);
    }
    
    // Validate configuration
    if (this.validateAssayerConfig(baseConfig)) {
      return baseConfig;
    } else {
      Logger.log('[ConfigManager_Assayer] Using fallback configuration due to validation failure');
      return Config_;
    }
  },
  
  // --------------------------------------------------------------------------
  // loadSatelliteConfigs - Load configuration from satellite config sheets
  // --------------------------------------------------------------------------
  loadSatelliteConfigs(ss) {
    const enhancedConfig = {};
    
    // Load Tier1 configuration
    const tier1Sheet = ss.getSheetByName("Config_Tier1");
    if (tier1Sheet) {
      const tier1Data = tier1Sheet.getDataRange().getValues();
      const tier1Config = {};
      
      for (let i = 1; i < tier1Data.length; i++) {
        const row = tier1Data[i];
        if (row[0]) { // config_key
          tier1Config[String(row[0]).trim()] = this.parseConfigValue(row[1]);
        }
      }
      
      // Apply to Assayer config
      enhancedConfig.tier1 = tier1Config;
      enhancedConfig.tierThresholds = {
        strong: tier1Config.TIER_STRONG_MIN || 0.65,
        medium: tier1Config.TIER_MEDIUM_MIN || 0.55,
        weak: tier1Config.TIER_WEAK_MIN || 0.45
      };
      enhancedConfig.confidenceThresholds = {
        min: tier1Config.CONF_MIN || 0.60,
        elite: tier1Config.CONF_ELITE || 0.85
      };
    }
    
    // Load Tier2 configuration
    const tier2Sheet = ss.getSheetByName("Config_Tier2");
    if (tier2Sheet) {
      const tier2Data = tier2Sheet.getDataRange().getValues();
      const tier2Config = {};
      
      for (let i = 1; i < tier2Data.length; i++) {
        const row = tier2Data[i];
        if (row[0]) { // config_key
          tier2Config[String(row[0]).trim()] = this.parseConfigValue(row[1]);
        }
      }
      
      // Apply to Assayer config
      enhancedConfig.tier2 = tier2Config;
      enhancedConfig.spreadBuckets = tier2Config.SPREAD_BUCKETS || Config_.spreadBuckets;
      enhancedConfig.lineBuckets = tier2Config.LINE_BUCKETS || Config_.lineBuckets;
      enhancedConfig.confBuckets = tier2Config.CONF_BUCKETS || Config_.confBuckets;
    }
    
    return enhancedConfig;
  },
  
  // --------------------------------------------------------------------------
  // parseConfigValue - Parse configuration value from sheet
  // --------------------------------------------------------------------------
  parseConfigValue(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }
    
    const str = String(value).trim();
    
    // Boolean values
    if (str.toLowerCase() === "true") return true;
    if (str.toLowerCase() === "false") return false;
    
    // JSON values
    if (str.startsWith("[") || str.startsWith("{")) {
      try {
        return JSON.parse(str);
      } catch (e) {
        Logger.log('[ConfigManager_Assayer] Failed to parse JSON value: ' + str);
        return str;
      }
    }
    
    // Numeric values
    const num = parseFloat(str);
    if (!isNaN(num)) {
      return num;
    }
    
    // String values
    return str;
  },
  
  // --------------------------------------------------------------------------
  // validateAssayerConfig - Validate Assayer configuration
  // --------------------------------------------------------------------------
  validateAssayerConfig(config) {
    try {
      // Check required fields
      const required = ['version', 'name', 'sheets'];
      for (const field of required) {
        if (!config[field]) {
          Logger.log('[ConfigManager_Assayer] Missing required field: ' + field);
          return false;
        }
      }
      
      // Validate tier thresholds
      if (config.tierThresholds) {
        const thresholds = config.tierThresholds;
        if (thresholds.strong <= thresholds.medium ||
            thresholds.medium <= thresholds.weak) {
          Logger.log('[ConfigManager_Assayer] Invalid tier thresholds: must be strictly decreasing');
          return false;
        }
      }
      
      // Validate confidence thresholds
      if (config.confidenceThresholds) {
        const conf = config.confidenceThresholds;
        if (conf.min <= 0 || conf.min >= 1 ||
            conf.elite <= 0 || conf.elite >= 1 ||
            conf.elite <= conf.min) {
          Logger.log('[ConfigManager_Assayer] Invalid confidence thresholds');
          return false;
        }
      }
      
      // Validate bucket arrays
      const bucketArrays = ['spreadBuckets', 'lineBuckets', 'confBuckets'];
      for (const bucketType of bucketArrays) {
        if (config[bucketType] && Array.isArray(config[bucketType])) {
          const buckets = config[bucketType];
          if (buckets.length < 2) {
            Logger.log('[ConfigManager_Assayer] ' + bucketType + ' must have at least 2 elements');
            return false;
          }
          
          // Check if sorted
          for (let i = 1; i < buckets.length; i++) {
            if (buckets[i] <= buckets[i-1]) {
              Logger.log('[ConfigManager_Assayer] ' + bucketType + ' must be sorted in ascending order');
              return false;
            }
          }
        }
      }
      
      return true;
    } catch (err) {
      Logger.log('[ConfigManager_Assayer] Config validation failed: ' + err.message);
      return false;
    }
  },
  
  // --------------------------------------------------------------------------
  // getAssayerConfigWithFallback - Get config with tolerant fallback
  // --------------------------------------------------------------------------
  getAssayerConfigWithFallback() {
    try {
      return this.loadAssayerConfig();
    } catch (err) {
      Logger.log('[ConfigManager_Assayer] Using fallback config due to error: ' + err.message);
      return Config_;
    }
  },
  
  // --------------------------------------------------------------------------
  // updateAssayerFromSatellite - Update Assayer config from satellite changes
  // --------------------------------------------------------------------------
  updateAssayerFromSatellite() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const newConfig = this.loadAssayerConfig();
    
    // Update global Config_ object
    Object.assign(Config_, newConfig);
    
    Logger.log('[ConfigManager_Assayer] Updated Assayer config from satellite sheets');
    return newConfig;
  }
};

/**
 * tolerantAssayerParser_ - Assayer parser with tolerant matching for legacy data
 * Integrates with tolerant matching from Phase 3
 */
function tolerantAssayerParser_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0];
  const rows = [];
  
  // Use ContractEnforcer header mapping with tolerant fallback
  let headerMap;
  if (typeof createCanonicalHeaderMap_ !== 'undefined') {
    // Determine contract type based on sheet name
    let contract;
    if (sheetName.includes('Tier1') || sheetName.includes('Tier2') || sheetName.includes('OU_Log')) {
      contract = FORENSIC_LOGS_CONTRACT;
    } else if (sheetName === 'Bet_Slips') {
      contract = BET_SLIPS_CONTRACT;
    } else if (sheetName === 'ResultsClean') {
      contract = RESULTSCLEAN_CONTRACT;
    } else {
      // Use tolerant matching for unknown sheets
      headerMap = tolerantHeaderMatch_(headers, headers);
    }
    
    if (contract) {
      headerMap = createCanonicalHeaderMap_(contract, headers);
    }
  } else {
    headerMap = tolerantHeaderMatch_(headers, headers);
  }
  
  Logger.log('[tolerantAssayerParser_] Sheet: ' + sheetName + ', Headers mapped: ' + Object.keys(headerMap).length);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;
    
    const parsedRow = {};
    
    // Map columns using header map
    Object.keys(headerMap).forEach(function(canonical) {
      const colIdx = headerMap[canonical];
      if (colIdx >= 0 && colIdx < row.length) {
        parsedRow[canonical] = row[colIdx];
      }
    });
    
    // Apply tolerant data cleaning
    const cleanedRow = tolerantDataCleaning_(parsedRow);
    
    // Validate row has essential fields
    if (cleanedRow.league && cleanedRow.team) {
      rows.push(cleanedRow);
    }
  }
  
  Logger.log('[tolerantAssayerParser_] Parsed ' + rows.length + ' rows from ' + sheetName);
  return rows;
}

/**
 * validateConfigState_ - Assayer wrapper for config validation
 * @param {Object} config - Configuration object
 * @param {string} tier - Configuration tier
 * @returns {boolean} True if valid
 */
function validateConfigState_(config, tier) {
  if (typeof ConfigManager_Assayer !== 'undefined') {
    return ConfigManager_Assayer.validateAssayerConfig(config);
  }
  
  // Fallback validation
  return config && config.version && config.name;
}
