# Ma Golide — Recent picks audit (2026-10-12 to 2026-11-10)

Generated: 2026-11-10 17:44 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 916
- dated evidence rows: 2525 (1609 superseded observations)
- settled picks: 546 (370 wins / 176 losses / 0 pushes)
- hit rate: 67.8% (370/546 win/loss decisions)
- pending/unmatched picks: 0 prior-day · 370 total
- priced picks: 327
- paper P&L: -6.834 units
- ROI: -2.1% (P&L / 327 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-10
- same-day pending rows outside prior-day counters: 370
- same-day official-final picks included: 0
- eligible prior-day picks: 546
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `CAUTION` | 146 | 100–46–0 | 68.5% | 146 | -2.8% |
| `SKIPPED_INELIGIBLE` | 181 | 100–81–0 | 55.2% | 181 | -1.5% |
| `WATCHLIST_NO_ODDS` | 219 | 170–49–0 | 77.6% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 372 | 287–85–0 | 77.2% | 153 | +5.5% |
| `total` | 174 | 83–91–0 | 47.7% | 174 | -8.8% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 27 | 17–10–0 | 63.0% | 19 | -19.4% |
| `moneyline:home` | 345 | 270–75–0 | 78.3% | 134 | +9.1% |
| `total:over` | 113 | 51–62–0 | 45.1% | 113 | -13.6% |
| `total:under` | 61 | 32–29–0 | 52.5% | 61 | +0.1% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-09: Hawai'i Rainbow Warriors vs Portland Pilots (Final: **66–73**)

- 🔴 **LOST — Full-game moneyline: Hawai'i Rainbow Warriors (home)**
  - Expected: **60.0%** · Price: **1.278 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Hawai'i Rainbow Warriors (home)**
  - Actual First-half score: **36–40** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Hawai'i Rainbow Warriors (home)**
  - Expected: **78.9%** · Price: **1.278 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **139** · Expected: **54.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **139** · Expected: **56.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Washington Huskies vs Mississippi Valley State Delta Devils (Final: **77–64**)

- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **141** · Expected: **56.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Washington Huskies (home)**
  - Expected: **98.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Washington Huskies (home)**
  - Expected: **63.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Over**
  - Actual Full-game total: **141** · Expected: **51.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Over**
  - Actual Full-game total: **141** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Washington Huskies (home)**
  - Actual First-half score: **43–33** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **141** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Stanford Cardinal vs Rice Owls (Final: **91–73**)

- 🟢 **WON — First-half moneyline: Stanford Cardinal (home)**
  - Actual First-half score: **45–29** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **164** · Expected: **65.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Stanford Cardinal (home)**
  - Expected: **63.0%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **164** · Expected: **66.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Stanford Cardinal (home)**
  - Expected: **91.4%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: Bethesda Flames vs UT Arlington Mavericks (Final: **65–116**)

- 🟢 **WON — Full-game moneyline: UT Arlington Mavericks (away)**
  - Expected: **99.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Bethesda Flames (home)**
  - Actual First-half score: **32–57** · Expected: **52.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UT Arlington Mavericks (away)**
  - Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UT Arlington Mavericks (away)**
  - Actual First-half score: **32–57** · Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Oregon State Beavers vs Eastern Washington Eagles (Final: **94–84**)

- 🟢 **WON — First-half moneyline: Oregon State Beavers (home)**
  - Actual First-half score: **49–37** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oregon State Beavers (home)**
  - Expected: **80.7%** · Price: **1.042 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **178** · Expected: **68.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oregon State Beavers (home)**
  - Expected: **68.0%** · Price: **1.042 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 162.5: Under**
  - Actual Full-game total: **178** · Expected: **69.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Charlotte 49ers vs Seton Hall Pirates (Final: **62–68**)

- 🔴 **LOST — First-half moneyline: Charlotte 49ers (home)**
  - Actual First-half score: **34–36** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Seton Hall Pirates (away)**
  - Expected: **73.2%** · Price: **1.505 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 130.5: Over**
  - Actual Full-game total: **130** · Expected: **89.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Charlotte 49ers (home)**
  - Expected: **60.0%** · Price: **2.640 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **130** · Expected: **87.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Auburn Tigers vs App State Mountaineers (Final: **75–53**)

- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **128** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **60.0%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **128** · Expected: **72.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Auburn Tigers (home)**
  - Actual First-half score: **38–36** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **128** · Expected: **68.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **95.4%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: DePaul Blue Demons vs Alabama A&M Bulldogs (Final: **84–47**)

- 🟢 **WON — Full-game moneyline: DePaul Blue Demons (home)**
  - Expected: **94.7%** · Price: **1.011 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **131** · Expected: **75.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 141.5: Over**
  - Actual Full-game total: **131** · Expected: **75.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: DePaul Blue Demons (home)**
  - Expected: **52.0%** · Price: **1.011 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: DePaul Blue Demons (home)**
  - Actual First-half score: **37–27** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **131** · Expected: **79.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **131** · Expected: **76.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Purdue Boilermakers vs Illinois State Redbirds (Final: **76–59**)

- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **85.7%** · Price: **1.056 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **135** · Expected: **73.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **52.0%** · Price: **1.056 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **135** · Expected: **70.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Purdue Boilermakers (home)**
  - Actual First-half score: **36–28** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Louisiana Ragin' Cajuns vs Southern New Orleans Knights (Final: **56–39**)

- 🟢 **WON — First-half moneyline: Louisiana Ragin' Cajuns (home)**
  - Actual First-half score: **26–15** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Louisiana Ragin' Cajuns (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Louisiana Ragin' Cajuns (home)**
  - Expected: **58.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Missouri Tigers vs SIU Edwardsville Cougars (Final: **96–58**)

- 🟢 **WON — Full-game moneyline: Missouri Tigers (home)**
  - Expected: **96.8%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Missouri Tigers (home)**
  - Actual First-half score: **48–24** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 149.5: Over**
  - Actual Full-game total: **154** · Expected: **64.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Missouri Tigers (home)**
  - Expected: **59.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **154** · Expected: **65.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: LSU Tigers vs Louisiana Tech Bulldogs (Final: **90–50**)

- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **140** · Expected: **56.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: LSU Tigers (home)**
  - Actual First-half score: **47–20** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Under**
  - Actual Full-game total: **140** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: LSU Tigers (home)**
  - Expected: **58.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **140** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: LSU Tigers (home)**
  - Expected: **88.4%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: Tennessee State Tigers vs Fisk Bulldogs (Final: **123–73**)

- 🟢 **WON — Full-game moneyline: Tennessee State Tigers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Tennessee State Tigers (home)**
  - Actual First-half score: **54–36** · Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee State Tigers (home)**
  - Expected: **69.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Northwestern Wildcats vs Boston University Terriers (Final: **72–55**)

- 🟢 **WON — Full-game moneyline: Northwestern Wildcats (home)**
  - Expected: **92.8%** · Price: **1.048 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **127** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **127** · Expected: **67.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **127** · Expected: **68.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northwestern Wildcats (home)**
  - Actual First-half score: **39–24** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northwestern Wildcats (home)**
  - Expected: **60.0%** · Price: **1.048 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Texas A&M Aggies vs UL Monroe Warhawks (Final: **108–67**)

- 🔴 **LOST — Full-game total 168.5: Under**
  - Actual Full-game total: **175** · Expected: **79.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 167.5: Under**
  - Actual Full-game total: **175** · Expected: **79.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas A&M Aggies (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Texas A&M Aggies (home)**
  - Actual First-half score: **54–27** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas A&M Aggies (home)**
  - Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Oklahoma Sooners vs North Alabama Lions (Final: **81–72**)

- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **153** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Oklahoma Sooners (home)**
  - Actual First-half score: **37–36** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **62.0%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **153** · Expected: **65.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **98.2%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: McNeese Cowboys vs College of Biblical Studies Ambassadors (Final: **135–87**)

- 🟢 **WON — Full-game moneyline: McNeese Cowboys (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: McNeese Cowboys (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: McNeese Cowboys (home)**
  - Actual First-half score: **63–44** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: UAB Blazers vs Florida A&M Rattlers (Final: **92–64**)

- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **156** · Expected: **63.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UAB Blazers (home)**
  - Expected: **95.1%** · Price: **1.036 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **156** · Expected: **58.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UAB Blazers (home)**
  - Expected: **62.0%** · Price: **1.036 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UAB Blazers (home)**
  - Actual First-half score: **47–27** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Mississippi State Bulldogs vs Jackson State Tigers (Final: **98–57**)

- 🟢 **WON — Full-game moneyline: Mississippi State Bulldogs (home)**
  - Expected: **98.0%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Mississippi State Bulldogs (home)**
  - Actual First-half score: **51–29** · Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **155** · Expected: **69.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Mississippi State Bulldogs (home)**
  - Expected: **67.0%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **155** · Expected: **57.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **155** · Expected: **64.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Alabama Crimson Tide vs Oakland Golden Grizzlies (Final: **109–68**)

- 🔴 **LOST — Full-game total 171.5: Under**
  - Actual Full-game total: **177** · Expected: **82.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 172.5: Under**
  - Actual Full-game total: **177** · Expected: **86.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Alabama Crimson Tide (home)**
  - Actual First-half score: **56–29** · Expected: **69.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 170.5: Under**
  - Actual Full-game total: **177** · Expected: **81.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama Crimson Tide (home)**
  - Expected: **97.4%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Alabama Crimson Tide (home)**
  - Expected: **67.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Western Illinois Leathernecks vs Simpson (IA) Storm (Final: **90–48**)

- 🟢 **WON — First-half moneyline: Western Illinois Leathernecks (home)**
  - Actual First-half score: **40–22** · Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Illinois Leathernecks (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Western Illinois Leathernecks (home)**
  - Expected: **58.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Akron Zips vs Muskingum Fighting Muskies (Final: **112–64**)

- 🟢 **WON — Full-game moneyline: Akron Zips (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Akron Zips (home)**
  - Actual First-half score: **62–27** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Akron Zips (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: UIC Flames vs UNC Greensboro Spartans (Final: **82–72**)

- 🟢 **WON — Full-game total 149.5: Over**
  - Actual Full-game total: **154** · Expected: **57.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UIC Flames (home)**
  - Expected: **87.6%** · Price: **1.028 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: UIC Flames (home)**
  - Actual First-half score: **33–32** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UIC Flames (home)**
  - Expected: **67.0%** · Price: **1.028 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **154** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Notre Dame Fighting Irish vs Eastern Illinois Panthers (Final: **94–70**)

- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **98.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Notre Dame Fighting Irish (home)**
  - Actual First-half score: **46–33** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **67.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **164** · Expected: **73.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Coastal Carolina Chanticleers vs South Carolina Upstate Spartans (Final: **65–56**)

- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **121** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **73.9%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **121** · Expected: **62.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **52.0%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Coastal Carolina Chanticleers (home)**
  - Actual First-half score: **29–25** · Expected: **54.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Miami Hurricanes vs North Florida Ospreys (Final: **89–51**)

- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **99.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **69.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **140** · Expected: **70.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Miami Hurricanes (home)**
  - Actual First-half score: **44–25** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **140** · Expected: **72.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Loyola Maryland Greyhounds vs Washington Maryland Shoremen (Final: **106–64**)

- 🟢 **WON — Full-game moneyline: Loyola Maryland Greyhounds (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Loyola Maryland Greyhounds (home)**
  - Expected: **58.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Loyola Maryland Greyhounds (home)**
  - Actual First-half score: **48–25** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Chattanooga Mocs vs Covenant Fighting Scots (Final: **97–67**)

- 🟢 **WON — Full-game moneyline: Chattanooga Mocs (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Chattanooga Mocs (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Chattanooga Mocs (home)**
  - Actual First-half score: **52–22** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Jacksonville Dolphins vs Kansas City Roos (Final: **79–61**)

- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **140** · Expected: **68.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **140** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Jacksonville Dolphins (home)**
  - Actual First-half score: **36–20** · Expected: **55.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Jacksonville Dolphins (home)**
  - Expected: **60.4%** · Price: **1.521 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Jacksonville Dolphins (home)**
  - Expected: **56.0%** · Price: **1.521 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **140** · Expected: **70.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: UConn Huskies vs Wagner Seahawks (Final: **83–41**)

- 🟢 **WON — Full-game moneyline: UConn Huskies (home)**
  - Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **124** · Expected: **74.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UConn Huskies (home)**
  - Actual First-half score: **44–18** · Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UConn Huskies (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-09: Columbia Lions vs Stonehill Skyhawks (Final: **75–55**)

- 🟢 **WON — Full-game moneyline: Columbia Lions (home)**
  - Expected: **63.0%** · Price: **1.100 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Columbia Lions (home)**
  - Expected: **88.3%** · Price: **1.100 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Columbia Lions (home)**
  - Actual First-half score: **34–36** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **130** · Expected: **68.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **130** · Expected: **68.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Towson Tigers vs Long Island University Sharks (Final: **66–46**)

- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **112** · Expected: **74.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **112** · Expected: **75.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 141.5: Over**
  - Actual Full-game total: **112** · Expected: **74.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Towson Tigers (home)**
  - Expected: **59.0%** · Price: **1.360 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Towson Tigers (home)**
  - Actual First-half score: **38–28** · Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Towson Tigers (home)**
  - Expected: **80.1%** · Price: **1.360 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: Le Moyne Dolphins vs Elmira Soaring Eagles (Final: **109–39**)

- 🟢 **WON — First-half moneyline: Le Moyne Dolphins (home)**
  - Actual First-half score: **53–20** · Expected: **55.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Le Moyne Dolphins (home)**
  - Expected: **57.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Le Moyne Dolphins (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-09: Chicago State Cougars vs South Carolina State Bulldogs (Final: **67–82**)

- 🔴 **LOST — Full-game moneyline: Chicago State Cougars (home)**
  - Expected: **58.0%** · Price: **1.741 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **149** · Expected: **63.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Chicago State Cougars (home)**
  - Expected: **59.0%** · Price: **1.741 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **149** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Chicago State Cougars (home)**
  - Actual First-half score: **33–53** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **149** · Expected: **65.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Clemson Tigers vs North Dakota Fighting Hawks (Final: **92–77**)

- 🟢 **WON — Full-game moneyline: Clemson Tigers (home)**
  - Expected: **97.2%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Clemson Tigers (home)**
  - Actual First-half score: **38–40** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Over**
  - Actual Full-game total: **169** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Clemson Tigers (home)**
  - Expected: **60.0%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Cincinnati Bearcats vs Minnesota Golden Gophers (Final: **61–63**)

- 🔴 **LOST — First-half moneyline: Cincinnati Bearcats (home)**
  - Actual First-half score: **30–30** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Cincinnati Bearcats (home)**
  - Expected: **58.0%** · Price: **1.323 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **124** · Expected: **68.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Cincinnati Bearcats (home)**
  - Expected: **83.9%** · Price: **1.323 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: Bowling Green Falcons vs Cleveland State Vikings (Final: **95–79**)

- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **174** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Over**
  - Actual Full-game total: **174** · Expected: **51.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **58.0%** · Price: **1.105 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **174** · Expected: **52.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Bowling Green Falcons (home)**
  - Actual First-half score: **44–37** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **87.7%** · Price: **1.105 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-09: Ball State Cardinals vs Brescia Bearcats (Final: **89–64**)

- 🟢 **WON — Full-game moneyline: Ball State Cardinals (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Ball State Cardinals (home)**
  - Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Ball State Cardinals (home)**
  - Actual First-half score: **45–26** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Texas Longhorns vs Grambling Tigers (Final: **93–62**)

- 🟢 **WON — First-half moneyline: Texas Longhorns (home)**
  - Actual First-half score: **55–27** · Expected: **70.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Over**
  - Actual Full-game total: **155** · Expected: **52.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **155** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Tennessee Volunteers vs Radford Highlanders (Final: **70–47**)

- 🔴 **LOST — Full-game total 152.5: Over**
  - Actual Full-game total: **117** · Expected: **56.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **117** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Volunteers (home)**
  - Expected: **98.8%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Tennessee Volunteers (home)**
  - Actual First-half score: **33–24** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Over**
  - Actual Full-game total: **117** · Expected: **52.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Volunteers (home)**
  - Expected: **60.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Georgia Bulldogs vs Florida Gulf Coast Eagles (Final: **80–72**)

- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **152** · Expected: **67.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia Bulldogs (home)**
  - Expected: **58.0%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Georgia Bulldogs (home)**
  - Actual First-half score: **44–41** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **152** · Expected: **70.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **152** · Expected: **69.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia Bulldogs (home)**
  - Expected: **95.9%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **152** · Expected: **65.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **152** · Expected: **73.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Pennsylvania Quakers vs Howard Bison (Final: **78–65**)

- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **143** · Expected: **68.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **143** · Expected: **52.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **143** · Expected: **62.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pennsylvania Quakers (home)**
  - Expected: **71.0%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **143** · Expected: **63.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Pennsylvania Quakers (home)**
  - Actual First-half score: **25–31** · Expected: **58.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pennsylvania Quakers (home)**
  - Expected: **55.0%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **143** · Expected: **58.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Buffalo Bulls vs Penn State Schuylkill Nittany Lions (Final: **107–65**)

- 🟢 **WON — Full-game moneyline: Buffalo Bulls (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Buffalo Bulls (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Buffalo Bulls (home)**
  - Actual First-half score: **58–31** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Indiana Hoosiers vs Syracuse Orange (Final: **83–80**)

- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **163** · Expected: **65.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Indiana Hoosiers (home)**
  - Expected: **68.1%** · Price: **1.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Indiana Hoosiers (home)**
  - Actual First-half score: **34–38** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **163** · Expected: **65.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Indiana Hoosiers (home)**
  - Expected: **60.0%** · Price: **1.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-09: Alabama State Hornets vs Georgia Gwinnett Grizzlies (Final: **77–79**)

- 🔴 **LOST — Full-game moneyline: Alabama State Hornets (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Alabama State Hornets (home)**
  - Expected: **99.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Alabama State Hornets (home)**
  - Actual First-half score: **34–32** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: West Florida Argonauts vs Baptist (Fl) Eagles (Final: **110–59**)

- 🟢 **WON — Full-game moneyline: West Florida Argonauts (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: West Florida Argonauts (home)**
  - Actual First-half score: **61–29** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: West Florida Argonauts (home)**
  - Expected: **58.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: UMBC Retrievers vs McDaniel Green Terror (Final: **91–61**)

- 🟢 **WON — First-half moneyline: UMBC Retrievers (home)**
  - Actual First-half score: **51–25** · Expected: **64.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UMBC Retrievers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UMBC Retrievers (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Canisius Golden Griffins vs Hiram Terriers (Final: **117–60**)

- 🟢 **WON — Full-game moneyline: Canisius Golden Griffins (home)**
  - Expected: **56.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Canisius Golden Griffins (home)**
  - Expected: **98.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Canisius Golden Griffins (home)**
  - Actual First-half score: **57–19** · Expected: **57.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Ohio State Buckeyes vs Marshall Thundering Herd (Final: **84–74**)

- 🟢 **WON — Full-game moneyline: Ohio State Buckeyes (home)**
  - Expected: **95.5%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **158** · Expected: **71.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ohio State Buckeyes (home)**
  - Expected: **66.0%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Ohio State Buckeyes (home)**
  - Actual First-half score: **45–27** · Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-09: Coppin State Eagles vs Cheyney (PA) Wolves (Final: **90–53**)

- 🟢 **WON — Full-game moneyline: Coppin State Eagles (home)**
  - Expected: **98.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Coppin State Eagles (home)**
  - Actual First-half score: **47–23** · Expected: **59.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Coppin State Eagles (home)**
  - Expected: **60.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Hawai'i Rainbow Warriors vs UT Arlington Mavericks (Final: **71–62**)

- 🔴 **LOST — First-half moneyline: Hawai'i Rainbow Warriors (home)**
  - Actual First-half score: **27–27** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hawai'i Rainbow Warriors (home)**
  - Expected: **84.0%** · Price: **1.235 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **133** · Expected: **63.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **133** · Expected: **61.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **133** · Expected: **64.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hawai'i Rainbow Warriors (home)**
  - Expected: **78.0%** · Price: **1.235 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Wichita State Shockers vs USC Trojans (Final: **93–86**)

- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **179** · Expected: **55.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: USC Trojans (away)**
  - Expected: **72.6%** · Price: **1.370 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Wichita State Shockers (home)**
  - Actual First-half score: **43–45** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **179** · Expected: **53.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wichita State Shockers (home)**
  - Expected: **78.0%** · Price: **3.200 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: California Golden Bears vs Houston Cougars (Final: **86–68**)

- 🔴 **LOST — Full-game moneyline: Houston Cougars (away)**
  - Expected: **93.4%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **154** · Expected: **74.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **154** · Expected: **69.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: California Golden Bears (home)**
  - Actual First-half score: **43–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **154** · Expected: **73.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: California Golden Bears (home)**
  - Expected: **78.0%** · Price: **15.000 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Penn State Nittany Lions vs Pittsburgh Panthers (Final: **76–82**)

- 🟢 **WON — Full-game total 153.5: Over**
  - Actual Full-game total: **158** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Penn State Nittany Lions (home)**
  - Expected: **78.0%** · Price: **2.200 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Penn State Nittany Lions (home)**
  - Actual First-half score: **38–42** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pittsburgh Panthers (away)**
  - Expected: **57.6%** · Price: **1.704 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-08: Bethesda Flames vs Portland Pilots (Final: **56–130**)

- 🔴 **LOST — First-half moneyline: Bethesda Flames (home)**
  - Actual First-half score: **30–60** · Expected: **62.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Bethesda Flames (home)**
  - Expected: **61.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Portland Pilots (away)**
  - Expected: **99.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: UTEP Miners vs Incarnate Word Cardinals (Final: **79–60**)

- 🟢 **WON — First-half moneyline: UTEP Miners (home)**
  - Actual First-half score: **46–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **78.0%** · Price: **1.118 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 141.5: Over**
  - Actual Full-game total: **139** · Expected: **75.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **83.3%** · Price: **1.118 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **83.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: William & Mary Tribe vs Norfolk State Spartans (Final: **104–73**)

- 🟢 **WON — Full-game moneyline: William & Mary Tribe (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: William & Mary Tribe (home)**
  - Expected: **78.0%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: William & Mary Tribe (home)**
  - Expected: **72.2%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: William & Mary Tribe (home)**
  - Actual First-half score: **54–35** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: William & Mary Tribe (home)**
  - Expected: **72.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **177** · Expected: **60.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Cal State Bakersfield Roadrunners vs North Dakota State Bison (Final: **91–80**)

- 🟢 **WON — Full-game total 153.5: Over**
  - Actual Full-game total: **171** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: North Dakota State Bison (away)**
  - Expected: **54.5%** · Price: **1.704 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Cal State Bakersfield Roadrunners (home)**
  - Expected: **78.0%** · Price: **2.200 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **171** · Expected: **57.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **171** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Cal State Bakersfield Roadrunners (home)**
  - Actual First-half score: **38–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Idaho Vandals vs Lincoln (CA) Oaklanders (Final: **108–69**)

- 🟢 **WON — First-half moneyline: Idaho Vandals (home)**
  - Actual First-half score: **59–33** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Idaho Vandals (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Idaho Vandals (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: Temple Owls vs Richmond Spiders (Final: **53–79**)

- 🔴 **LOST — Full-game total 152.5: Over**
  - Actual Full-game total: **132** · Expected: **57.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Richmond Spiders (away)**
  - Expected: **58.3%** · Price: **1.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **132** · Expected: **61.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **132** · Expected: **58.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Temple Owls (home)**
  - Actual First-half score: **23–37** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Temple Owls (home)**
  - Expected: **78.0%** · Price: **2.050 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Nevada Wolf Pack vs Saint Mary's Gaels (Final: **96–95**)

- 🟢 **WON — First-half moneyline: Nevada Wolf Pack (home)**
  - Actual First-half score: **42–35** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nevada Wolf Pack (home)**
  - Expected: **78.0%** · Price: **1.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nevada Wolf Pack (home)**
  - Expected: **53.2%** · Price: **1.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **191** · Expected: **64.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: NJIT Highlanders vs Saint Peter's Peacocks (Final: **54–52**)

- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **106** · Expected: **90.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: NJIT Highlanders (home)**
  - Actual First-half score: **25–18** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NJIT Highlanders (home)**
  - Expected: **78.0%** · Price: **2.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Saint Peter's Peacocks (away)**
  - Expected: **52.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Saint Peter's Peacocks (away)**
  - Expected: **52.0%** · Price: **1.820 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: NJIT Highlanders (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Binghamton Bearcats vs Rider Broncs (Final: **75–63**)

- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (home)**
  - Expected: **75.3%** · Price: **1.278 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (home)**
  - Expected: **78.0%** · Price: **1.278 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (home)**
  - Expected: **75.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game total 137.5: Over**
  - Actual Full-game total: **138** · Expected: **84.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **138** · Expected: **82.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Binghamton Bearcats (home)**
  - Actual First-half score: **42–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Drake Bulldogs vs Southern Jaguars (Final: **75–68**)

- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **78.0%** · Price: **1.095 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **92.3%** · Price: **1.095 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **143** · Expected: **58.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Drake Bulldogs (home)**
  - Actual First-half score: **33–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **92.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: Tennessee Tech Golden Eagles vs Western Kentucky Hilltoppers (Final: **97–94**)

- 🔴 **LOST — Full-game moneyline: Western Kentucky Hilltoppers (away)**
  - Expected: **68.6%** · Price: **1.645 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **191** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Tech Golden Eagles (home)**
  - Expected: **78.0%** · Price: **2.300 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 164.5: Under**
  - Actual Full-game total: **191** · Expected: **71.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 162.5: Under**
  - Actual Full-game total: **191** · Expected: **68.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Tennessee Tech Golden Eagles (home)**
  - Actual First-half score: **50–41** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Air Force Falcons vs Idaho State Bengals (Final: **89–82**)

- 🟢 **WON — Full-game moneyline: Air Force Falcons (home)**
  - Expected: **78.0%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **171** · Expected: **66.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Air Force Falcons (home)**
  - Actual First-half score: **41–42** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **171** · Expected: **65.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Air Force Falcons (home)**
  - Expected: **60.5%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-08: Wofford Terriers vs Allen Yellow Jackets (Final: **100–67**)

- 🟢 **WON — First-half moneyline: Wofford Terriers (home)**
  - Actual First-half score: **55–22** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Cornell Big Red vs Keuka Storm (Final: **99–59**)

- 🟢 **WON — Full-game moneyline: Cornell Big Red (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Cornell Big Red (home)**
  - Actual First-half score: **46–31** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Cornell Big Red (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: UAlbany Great Danes vs Merrimack Warriors (Final: **60–73**)

- 🔴 **LOST — Full-game moneyline: UAlbany Great Danes (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UAlbany Great Danes (home)**
  - Actual First-half score: **26–30** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Merrimack Warriors (away)**
  - Expected: **52.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 134.5: Over**
  - Actual Full-game total: **133** · Expected: **89.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UAlbany Great Danes (home)**
  - Expected: **78.0%** · Price: **2.140 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Merrimack Warriors (away)**
  - Expected: **52.0%** · Price: **1.741 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-08: St. Thomas Tommies vs Green Bay Phoenix (Final: **67–76**)

- 🔴 **LOST — First-half moneyline: St. Thomas Tommies (home)**
  - Actual First-half score: **33–33** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **78.0%** · Price: **1.290 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **143** · Expected: **62.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **75.2%** · Price: **1.290 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-08: Western Michigan Broncos vs Campbell Fighting Camels (Final: **87–71**)

- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **158** · Expected: **58.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Michigan Broncos (home)**
  - Expected: **77.8%** · Price: **1.488 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **158** · Expected: **63.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Western Michigan Broncos (home)**
  - Actual First-half score: **41–35** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Michigan Broncos (home)**
  - Expected: **78.0%** · Price: **1.488 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Vermont Catamounts vs Quinnipiac Bobcats (Final: **76–79**)

- 🔴 **LOST — Full-game moneyline: Vermont Catamounts (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Vermont Catamounts (home)**
  - Expected: **78.2%** · Price: **1.263 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **155** · Expected: **81.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Vermont Catamounts (home)**
  - Expected: **78.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Vermont Catamounts (home)**
  - Actual First-half score: **40–37** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Vermont Catamounts (home)**
  - Expected: **78.0%** · Price: **1.263 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Dartmouth Big Green vs Central Connecticut Blue Devils (Final: **71–78**)

- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **149** · Expected: **78.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **80.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Dartmouth Big Green (home)**
  - Actual First-half score: **30–33** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **80.3%** · Price: **1.267 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **78.0%** · Price: **1.267 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **149** · Expected: **74.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Lehigh Mountain Hawks vs Brown Bears (Final: **55–65**)

- 🔴 **LOST — First-half moneyline: Lehigh Mountain Hawks (home)**
  - Actual First-half score: **25–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **120** · Expected: **81.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **120** · Expected: **82.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **72.9%** calibrated (raw 78.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=230** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **66.2%** · Price: **1.444 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **78.0%** · Price: **1.444 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **66.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: Boise State Broncos vs Butler Bulldogs (Final: **100–81**)

- 🟢 **WON — First-half moneyline: Boise State Broncos (home)**
  - Actual First-half score: **51–31** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **78.0%** · Price: **1.676 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **181** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **66.7%** · Price: **1.676 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **181** · Expected: **60.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **181** · Expected: **63.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Milwaukee Panthers vs Ripon Red Hawks (Final: **99–48**)

- 🟢 **WON — First-half moneyline: Milwaukee Panthers (home)**
  - Actual First-half score: **47–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Milwaukee Panthers (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Milwaukee Panthers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: Siena Saints vs Princeton Tigers (Final: **82–68**)

- 🟢 **WON — First-half moneyline: Siena Saints (home)**
  - Actual First-half score: **39–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Siena Saints (home)**
  - Expected: **78.0%** · Price: **2.100 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Princeton Tigers (away)**
  - Expected: **67.3%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-08: South Carolina Gamecocks vs Southeast Missouri State Redhawks (Final: **78–68**)

- 🟢 **WON — Full-game moneyline: South Carolina Gamecocks (home)**
  - Expected: **78.0%** · Price: **1.048 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **146** · Expected: **69.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Carolina Gamecocks (home)**
  - Expected: **93.9%** · Price: **1.048 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **146** · Expected: **69.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: South Carolina Gamecocks (home)**
  - Actual First-half score: **33–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: UMass Lowell River Hawks vs Marist Red Foxes (Final: **63–61**)

- 🟢 **WON — Full-game moneyline: UMass Lowell River Hawks (home)**
  - Expected: **73.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: UMass Lowell River Hawks (home)**
  - Actual First-half score: **30–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **124** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UMass Lowell River Hawks (home)**
  - Expected: **73.7%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UMass Lowell River Hawks (home)**
  - Expected: **78.0%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 142.5: Over**
  - Actual Full-game total: **124** · Expected: **74.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: New Hampshire Wildcats vs Sacred Heart Pioneers (Final: **76–81**)

- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **157** · Expected: **63.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New Hampshire Wildcats (home)**
  - Expected: **78.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New Hampshire Wildcats (home)**
  - Expected: **60.6%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: New Hampshire Wildcats (home)**
  - Expected: **60.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: New Hampshire Wildcats (home)**
  - Actual First-half score: **34–33** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **157** · Expected: **69.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Maine Black Bears vs Manhattan Jaspers (Final: **62–68**)

- 🔴 **LOST — Full-game moneyline: Maine Black Bears (home)**
  - Expected: **71.7%** · Price: **1.704 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **130** · Expected: **86.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Maine Black Bears (home)**
  - Actual First-half score: **29–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Maine Black Bears (home)**
  - Expected: **71.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **130** · Expected: **81.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Maine Black Bears (home)**
  - Expected: **78.0%** · Price: **1.704 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Bryant Bulldogs vs Fairfield Stags (Final: **59–68**)

- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **127** · Expected: **69.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fairfield Stags (away)**
  - Expected: **56.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Bryant Bulldogs (home)**
  - Expected: **78.0%** · Price: **2.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fairfield Stags (away)**
  - Expected: **56.4%** · Price: **1.820 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **127** · Expected: **64.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **127** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Bryant Bulldogs (home)**
  - Actual First-half score: **30–26** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Navy Midshipmen vs Central Penn Knights (Final: **102–46**)

- 🟢 **WON — First-half moneyline: Navy Midshipmen (home)**
  - Actual First-half score: **52–14** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Navy Midshipmen (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Navy Midshipmen (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-08: Virginia Cavaliers vs Arkansas Razorbacks (Final: **91–72**)

- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **78.0%** · Price: **1.571 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Arkansas Razorbacks (away)**
  - Expected: **58.2%** · Price: **2.450 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Virginia Cavaliers (home)**
  - Actual First-half score: **46–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 162.5: Under**
  - Actual Full-game total: **163** · Expected: **68.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-08: Chicago Maroons vs Harvard Crimson (Final: **47–63**)

- 🔴 **LOST — First-half moneyline: Chicago Maroons (home)**
  - Actual First-half score: **20–35** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Chicago Maroons (home)**
  - Expected: **78.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Harvard Crimson (away)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-08: St. John's Red Storm vs Iona Gaels (Final: **104–65**)

- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **78.0%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: St. John's Red Storm (home)**
  - Actual First-half score: **54–23** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **169** · Expected: **58.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **169** · Expected: **61.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **95.3%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: UC San Diego Tritons vs Fresno State Bulldogs (Final: **81–65**)

- 🟢 **WON — Full-game moneyline: UC San Diego Tritons (home)**
  - Expected: **88.7%** · Price: **1.154 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: UC San Diego Tritons (home)**
  - Actual First-half score: **32–36** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **146** · Expected: **51.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Santa Clara Broncos vs Sacramento State Hornets (Final: **83–77**)

- 🟢 **WON — Full-game total 170.5: Under**
  - Actual Full-game total: **160** · Expected: **82.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Santa Clara Broncos (home)**
  - Expected: **91.7%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Santa Clara Broncos (home)**
  - Actual First-half score: **45–36** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Seattle U Redhawks vs Portland State Vikings (Final: **79–60**)

- 🟢 **WON — Full-game moneyline: Seattle U Redhawks (home)**
  - Expected: **81.8%** · Price: **1.160 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **139** · Expected: **59.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Seattle U Redhawks (home)**
  - Actual First-half score: **28–28** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: UC Irvine Anteaters vs San José State Spartans (Final: **91–79**)

- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **170** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UC Irvine Anteaters (home)**
  - Actual First-half score: **51–31** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Irvine Anteaters (home)**
  - Expected: **76.3%** · Price: **1.253 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Pepperdine Waves vs Bethesda Flames (Final: **116–43**)

- 🟢 **WON — First-half moneyline: Pepperdine Waves (home)**
  - Actual First-half score: **48–18** · Expected: **86.1%** calibrated (raw 88.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pepperdine Waves (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-07: Weber State Wildcats vs Utah Valley Wolverines (Final: **93–101**)

- 🟢 **WON — Full-game moneyline: Utah Valley Wolverines (away)**
  - Expected: **66.9%** · Price: **1.294 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Weber State Wildcats (home)**
  - Actual First-half score: **45–52** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **194** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Southern Utah Thunderbirds vs Bethesda Flames (Final: **143–49**)

- 🟢 **WON — First-half moneyline: Southern Utah Thunderbirds (home)**
  - Actual First-half score: **73–24** · Expected: **86.1%** calibrated (raw 88.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Southern Utah Thunderbirds (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-07: Oral Roberts Golden Eagles vs Northwestern (OK) Rangers (Final: **94–39**)

- 🟢 **WON — Full-game moneyline: Oral Roberts Golden Eagles (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Oral Roberts Golden Eagles (home)**
  - Actual First-half score: **48–17** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Murray State Racers vs Northwestern State Demons (Final: **93–71**)

- 🟢 **WON — First-half moneyline: Murray State Racers (home)**
  - Actual First-half score: **42–39** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **164** · Expected: **56.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **93.0%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Evansville Purple Aces vs Southern Indiana Screaming Eagles (Final: **78–76**)

- 🔴 **LOST — First-half moneyline: Evansville Purple Aces (home)**
  - Actual First-half score: **30–35** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **154** · Expected: **63.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Evansville Purple Aces (home)**
  - Expected: **54.3%** · Price: **1.294 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Little Rock Trojans vs SE Louisiana Lions (Final: **79–83**)

- 🔴 **LOST — Full-game moneyline: Little Rock Trojans (home)**
  - Expected: **77.3%** · Price: **1.206 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **162** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Little Rock Trojans (home)**
  - Actual First-half score: **39–46** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Bradley Braves vs Pacific Tigers (Final: **52–63**)

- 🔴 **LOST — Full-game moneyline: Bradley Braves (home)**
  - Expected: **70.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **115** · Expected: **74.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Bradley Braves (home)**
  - Actual First-half score: **19–29** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Grand Canyon Lopes vs Bethune-Cookman Wildcats (Final: **99–81**)

- 🟢 **WON — First-half moneyline: Grand Canyon Lopes (home)**
  - Actual First-half score: **56–43** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **180** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Grand Canyon Lopes (home)**
  - Expected: **96.3%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: UC Riverside Highlanders vs Denver Pioneers (Final: **64–60**)

- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **124** · Expected: **58.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Riverside Highlanders (home)**
  - Expected: **63.7%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: UC Riverside Highlanders (home)**
  - Actual First-half score: **22–29** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Duquesne Dukes vs Central State Marauders (Final: **68–58**)

- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Duquesne Dukes (home)**
  - Actual First-half score: **31–30** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Georgetown Hoyas vs Maryland Terrapins (Final: **86–74**)

- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **160** · Expected: **56.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Georgetown Hoyas (home)**
  - Actual First-half score: **44–28** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Maryland Terrapins (away)**
  - Expected: **60.1%** · Price: **1.617 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Purdue Fort Wayne Mastodons vs Albion Britons (Final: **103–65**)

- 🟢 **WON — First-half moneyline: Purdue Fort Wayne Mastodons (home)**
  - Actual First-half score: **60–26** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Purdue Fort Wayne Mastodons (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-07: Davidson Wildcats vs Mercer Bears (Final: **88–76**)

- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **164** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **88.1%** · Price: **1.133 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Davidson Wildcats (home)**
  - Actual First-half score: **34–33** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Western Illinois Leathernecks vs Houston Christian Huskies (Final: **67–74**)

- 🟢 **WON — Full-game moneyline: Houston Christian Huskies (away)**
  - Expected: **65.4%** · Price: **1.213 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 136.5: Over**
  - Actual Full-game total: **141** · Expected: **79.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Western Illinois Leathernecks (home)**
  - Actual First-half score: **31–37** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Canisius Golden Griffins vs St. Bonaventure Bonnies (Final: **49–88**)

- 🟢 **WON — Full-game moneyline: St. Bonaventure Bonnies (away)**
  - Expected: **68.5%** · Price: **1.235 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Canisius Golden Griffins (home)**
  - Actual First-half score: **14–43** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **137** · Expected: **52.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Dayton Flyers vs Northern Illinois Huskies (Final: **76–81**)

- 🔴 **LOST — Full-game moneyline: Dayton Flyers (home)**
  - Expected: **96.4%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **157** · Expected: **55.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Dayton Flyers (home)**
  - Actual First-half score: **29–40** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Nebraska Cornhuskers vs Providence Friars (Final: **67–58**)

- 🟢 **WON — Full-game moneyline: Nebraska Cornhuskers (home)**
  - Expected: **70.0%** · Price: **1.645 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **125** · Expected: **59.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Nebraska Cornhuskers (home)**
  - Actual First-half score: **30–25** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Niagara Purple Eagles vs Morgan State Bears (Final: **67–73**)

- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **140** · Expected: **54.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Niagara Purple Eagles (home)**
  - Expected: **63.3%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Niagara Purple Eagles (home)**
  - Actual First-half score: **34–42** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: George Washington Revolutionaries vs Hampton Pirates (Final: **75–61**)

- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **136** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **91.9%** · Price: **1.036 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: George Washington Revolutionaries (home)**
  - Actual First-half score: **41–27** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: New Mexico Lobos vs Northern Arizona Lumberjacks (Final: **94–65**)

- 🟢 **WON — First-half moneyline: New Mexico Lobos (home)**
  - Actual First-half score: **46–30** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Mexico Lobos (home)**
  - Expected: **96.2%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **159** · Expected: **59.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: UT Rio Grande Valley Vaqueros vs Prairie View A&M Panthers (Final: **68–63**)

- 🟢 **WON — Full-game moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Expected: **84.4%** · Price: **1.235 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **131** · Expected: **52.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Actual First-half score: **30–29** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: UNC Asheville Bulldogs vs Western Carolina Catamounts (Final: **75–83**)

- 🔴 **LOST — First-half moneyline: UNC Asheville Bulldogs (home)**
  - Actual First-half score: **40–44** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Carolina Catamounts (away)**
  - Expected: **56.2%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **158** · Expected: **51.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Gonzaga Bulldogs vs New Mexico State Aggies (Final: **109–68**)

- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **177** · Expected: **51.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Gonzaga Bulldogs (home)**
  - Expected: **97.8%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Gonzaga Bulldogs (home)**
  - Actual First-half score: **47–42** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: UTSA Roadrunners vs VMI Keydets (Final: **70–68**)

- 🟢 **WON — Full-game moneyline: UTSA Roadrunners (home)**
  - Expected: **76.1%** · Price: **1.459 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **138** · Expected: **72.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UTSA Roadrunners (home)**
  - Actual First-half score: **31–35** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Furman Paladins vs Northern Iowa Panthers (Final: **98–83**)

- 🟢 **WON — Full-game total 149.5: Over**
  - Actual Full-game total: **181** · Expected: **52.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Furman Paladins (home)**
  - Actual First-half score: **44–31** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Furman Paladins (home)**
  - Expected: **76.3%** · Price: **1.521 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Bucknell Bison vs Old Dominion Monarchs (Final: **62–82**)

- 🟢 **WON — Full-game moneyline: Old Dominion Monarchs (away)**
  - Expected: **60.9%** · Price: **1.408 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Bucknell Bison (home)**
  - Actual First-half score: **25–26** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **144** · Expected: **69.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: IU Indianapolis Jaguars vs UT Martin Skyhawks (Final: **75–74**)

- 🟢 **WON — Full-game total 171.5: Under**
  - Actual Full-game total: **149** · Expected: **83.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UT Martin Skyhawks (away)**
  - Expected: **58.3%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: IU Indianapolis Jaguars (home)**
  - Actual First-half score: **36–34** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 168.5: Under**
  - Actual Full-game total: **149** · Expected: **77.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Lipscomb Bisons vs SIU Edwardsville Cougars (Final: **78–77**)

- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **155** · Expected: **59.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Lipscomb Bisons (home)**
  - Actual First-half score: **46–35** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lipscomb Bisons (home)**
  - Expected: **77.6%** · Price: **1.294 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: High Point Panthers vs Washington State Cougars (Final: **83–71**)

- 🟢 **WON — Full-game moneyline: High Point Panthers (home)**
  - Expected: **91.7%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **154** · Expected: **58.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: High Point Panthers (home)**
  - Actual First-half score: **32–31** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Miami (OH) RedHawks vs Bellarmine Knights (Final: **80–73**)

- 🟢 **WON — Full-game moneyline: Miami (OH) RedHawks (home)**
  - Expected: **92.2%** · Price: **1.028 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Miami (OH) RedHawks (home)**
  - Actual First-half score: **36–31** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **153** · Expected: **60.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: Lindenwood Lions vs East West Phantoms (Final: **115–44**)

- 🟢 **WON — Full-game moneyline: Lindenwood Lions (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Lindenwood Lions (home)**
  - Actual First-half score: **63–24** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Maryland Eastern Shore Hawks vs Cheyney (PA) Wolves (Final: **81–55**)

- 🟢 **WON — First-half moneyline: Maryland Eastern Shore Hawks (home)**
  - Actual First-half score: **37–24** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Maryland Eastern Shore Hawks (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-07: Utah State Aggies vs Omaha Mavericks (Final: **90–56**)

- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **96.6%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Utah State Aggies (home)**
  - Actual First-half score: **44–21** · Expected: **84.1%** calibrated (raw 86.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **146** · Expected: **54.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-07: UC Santa Barbara Gauchos vs Oakland Golden Grizzlies (Final: **90–67**)

- 🟢 **WON — Full-game total 149.5: Over**
  - Actual Full-game total: **157** · Expected: **52.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UC Santa Barbara Gauchos (home)**
  - Actual First-half score: **38–27** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Santa Barbara Gauchos (home)**
  - Expected: **75.3%** · Price: **1.095 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: UNLV Rebels vs Samford Bulldogs (Final: **88–79**)

- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **167** · Expected: **59.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UNLV Rebels (home)**
  - Actual First-half score: **45–40** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNLV Rebels (home)**
  - Expected: **85.9%** · Price: **1.125 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Saint Joseph's Hawks vs Lafayette Leopards (Final: **67–57**)

- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **124** · Expected: **55.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Joseph's Hawks (home)**
  - Expected: **91.7%** · Price: **1.056 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Saint Joseph's Hawks (home)**
  - Actual First-half score: **31–24** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Rhode Island Rams vs Monmouth Hawks (Final: **84–66**)

- 🟢 **WON — First-half moneyline: Rhode Island Rams (home)**
  - Actual First-half score: **36–22** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **150** · Expected: **59.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **66.8%** · Price: **1.377 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn`

### 2026-11-07: Mercyhurst Lakers vs Penn State Shenango Nittany Lions (Final: **89–50**)

- 🟢 **WON — First-half moneyline: Mercyhurst Lakers (home)**
  - Actual First-half score: **44–25** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Mercyhurst Lakers (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-07: Sam Houston Bearkats vs Howard Payne Yellow Jackets (Final: **102–59**)

- 🟢 **WON — First-half moneyline: Sam Houston Bearkats (home)**
  - Actual First-half score: **53–23** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: La Salle Explorers vs Long Island University Sharks (Final: **88–80**)

- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **168** · Expected: **52.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: La Salle Explorers (home)**
  - Actual First-half score: **39–37** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-07: Colorado Buffaloes vs San Diego State Aztecs (Final: **66–77**)

- 🔴 **LOST — First-half moneyline: Colorado Buffaloes (home)**
  - Actual First-half score: **17–35** · Expected: **83.1%** calibrated (raw 85.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=169** · calibrated · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **143** · Expected: **51.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
