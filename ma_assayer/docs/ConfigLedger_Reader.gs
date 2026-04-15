/******************************************************************************
 * CONFIG LEDGER READER — Assayer Module
 * Repo: Ma_Assayer
 *
 * Paste BEFORE MODULE 02 (or immediately after MODULE 05 Config) so Parser
 * and Stats can call ConfigLedger_Reader.
 ******************************************************************************/

var ConfigLedger_Reader = {

  _cache: null,
  _satelliteId: null,
  _log: null,

  init: function (satelliteSpreadsheetId) {
    this._satelliteId = satelliteSpreadsheetId || null;
    this._cache = null;
    this._log = (typeof Log_ !== "undefined") ? Log_.module("CFG_LEDGER") : {
      info: function (m) { console.log(m); },
      warn: function (m) { console.warn(m); },
      error: function (m) { console.error(m); }
    };
    this._log.info(
      "ConfigLedger_Reader initialised. Source: " + (this._satelliteId || "same sheet")
    );
  },

  resolveStamp: function (bet) {
    if (!bet) {
      return bet;
    }
    var rawStamp = bet.config_stamp || bet.configStamp || null;
    if (!rawStamp) {
      bet.stampId = null;
      bet.configVersion = null;
      bet.configBuiltAt = null;
      bet.configMeta = null;
      return bet;
    }
    var meta = this._lookup(rawStamp);
    bet.stampId = rawStamp;
    bet.configVersion = meta ? meta.version : null;
    bet.configBuiltAt = meta ? meta.built_at : null;
    bet.configMeta = meta || null;
    return bet;
  },

  tagSlice: function (sliceStats, bets) {
    if (!sliceStats || !Array.isArray(bets) || bets.length === 0) {
      return sliceStats;
    }
    var counts = {};
    var i;
    for (i = 0; i < bets.length; i++) {
      var b = bets[i];
      var sid = b.stampId || "__UNSTAMPED__";
      counts[sid] = (counts[sid] || 0) + 1;
    }
    var dominant = null;
    var maxCount = 0;
    var sid2;
    for (sid2 in counts) {
      if (Object.prototype.hasOwnProperty.call(counts, sid2)) {
        var cnt = counts[sid2];
        if (cnt > maxCount) {
          maxCount = cnt;
          dominant = sid2;
        }
      }
    }
    var meta = dominant && dominant !== "__UNSTAMPED__" ? this._lookup(dominant) : null;
    var total = bets.length;
    sliceStats.dominantStampId = dominant !== "__UNSTAMPED__" ? dominant : null;
    sliceStats.dominantVersion = meta ? meta.version : null;
    sliceStats.stampMix = counts;
    sliceStats.stampPurity = dominant ? (maxCount / total) : 0;
    return sliceStats;
  },

  getAll: function () {
    this._ensureLoaded();
    return this._cache ? Array.from(this._cache.values()) : [];
  },

  getStampMeta: function (stampId) {
    return this._lookup(stampId);
  },

  summariseStamps: function (allBets) {
    if (!Array.isArray(allBets)) {
      return {};
    }
    var counts = {};
    var j;
    for (j = 0; j < allBets.length; j++) {
      var sid0 = allBets[j].stampId || "__UNSTAMPED__";
      counts[sid0] = (counts[sid0] || 0) + 1;
    }
    var total = allBets.length;
    var rows = [];
    var sid3;
    for (sid3 in counts) {
      if (Object.prototype.hasOwnProperty.call(counts, sid3)) {
        var cnt0 = counts[sid3];
        var meta0 = (sid3 !== "__UNSTAMPED__") ? (this._lookup(sid3) || {}) : {};
        rows.push({
          stampId: sid3,
          version: meta0.version || "unknown",
          builtAt: meta0.built_at || "unknown",
          count: cnt0,
          pct: total > 0 ? (cnt0 / total) : 0
        });
      }
    }
    rows.sort(function (a, b) {
      return b.count - a.count;
    });
    return { total: total, rows: rows, uniqueStamps: rows.length };
  },

  _lookup: function (stampId) {
    this._ensureLoaded();
    return (this._cache && this._cache.has(stampId)) ? this._cache.get(stampId) : null;
  },

  _ensureLoaded: function () {
    if (this._cache !== null) {
      return;
    }
    this._cache = new Map();
    try {
      var ss = this._satelliteId
        ? SpreadsheetApp.openById(this._satelliteId)
        : SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Config_Ledger");
      if (!sheet) {
        if (this._log) {
          this._log.warn("Config_Ledger sheet not found — stamps will be null");
        }
        return;
      }
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        return;
      }
      var headers = data[0].map(function (h) {
        return String(h).trim().toLowerCase().replace(/\s+/g, "_");
      });
      var r;
      for (r = 1; r < data.length; r++) {
        var row = data[r];
        var obj = {};
        var c;
        for (c = 0; c < headers.length; c++) {
          obj[headers[c]] = row[c];
        }
        var stampId = String(obj.stamp_id || "").trim();
        if (stampId) {
          this._cache.set(stampId, obj);
        }
      }
      if (this._log) {
        this._log.info("Config_Ledger loaded: " + this._cache.size + " stamp(s)");
      }
    } catch (err) {
      if (this._log) {
        this._log.warn("Config_Ledger load failed: " + err.message);
      }
    }
  }
};

// =====================================================================
// FIX: Prevent duplicate ColResolver_ declaration
// Apps Script loads files alphabetically, so we need conditional check
// =====================================================================
if (typeof ColResolver_ === 'undefined') {
  // Only define if not already present from Assayer modules
  const ColResolver_ = {
    // (the full ColResolver_ object from Assayer is already loaded)
    // We just need the existence check to prevent duplicate declaration
  };
}

/**
 * DYNAMIC TIMEOUT HANDLER - Adjusts timeout based on satellite game count
 * Prevents timeouts when accessing hundreds of satellites with different league sizes
 */
ConfigLedger_Reader.getSatelliteTimeout = function(satelliteId) {
  // Default timeouts for different game counts
  const timeouts = {
    '3': 15000,    // Small leagues (3 games)
    '7': 25000,    // Medium leagues (7 games) 
    '9': 35000,    // Large leagues (9 games)
    '14': 45000,   // Very large leagues (14 games)
    'default': 30000  // Default for unknown sizes
  };
  
  try {
    if (!satelliteId) return timeouts.default;
    
    // Try to detect game count by accessing satellite briefly
    const ss = SpreadsheetApp.openById(satelliteId);
    const resultsSheet = ss.getSheetByName('ResultsClean') || ss.getSheetByName('Results');
    if (resultsSheet) {
      const gameCount = resultsSheet.getLastRow() - 1;
      if (gameCount <= 3) return timeouts['3'];
      if (gameCount <= 7) return timeouts['7'];
      if (gameCount <= 9) return timeouts['9'];
      if (gameCount <= 14) return timeouts['14'];
    }
    
    return timeouts.default;
  } catch (e) {
    // If we can't detect, use conservative default
    return timeouts.default;
  }
};
