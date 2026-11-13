# Ma Golide — Recent picks audit (2026-10-13 to 2026-11-11)

Generated: 2026-11-12 04:14 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 403
- dated evidence rows: 795 (392 superseded observations)
- settled picks: 373 (287 wins / 86 losses / 0 pushes)
- hit rate: 76.9% (287/373 win/loss decisions)
- pending/unmatched picks: 30 prior-day · 30 total
- priced picks: 37
- paper P&L: -5.461 units
- ROI: -14.8% (P&L / 37 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-12
- same-day pending rows outside prior-day counters: 0
- same-day official-final picks included: 0
- eligible prior-day picks: 403
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `SKIPPED_INELIGIBLE` | 37 | 20–17–0 | 54.1% | 37 | -14.8% |
| `WATCHLIST_NO_ODDS` | 336 | 267–69–0 | 79.5% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 344 | 275–69–0 | 79.9% | 8 | +7.4% |
| `spread` | 16 | 8–8–0 | 50.0% | 16 | -4.8% |
| `total` | 13 | 4–9–0 | 30.8% | 13 | -40.7% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 23 | 17–6–0 | 73.9% | 0 | n/a |
| `moneyline:home` | 321 | 258–63–0 | 80.4% | 8 | +7.4% |
| `spread:away` | 14 | 6–8–0 | 42.9% | 14 | -18.2% |
| `spread:home` | 2 | 2–0–0 | 100.0% | 2 | +88.9% |
| `total:over` | 6 | 3–3–0 | 50.0% | 6 | -4.5% |
| `total:under` | 7 | 1–6–0 | 14.3% | 7 | -71.7% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-11: Illinois State Redbirds vs Green Bay Phoenix (Final: **75–66**)

- 🔴 **LOST — Full-game moneyline: Green Bay Phoenix (away)**
  - Expected: **69.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Illinois State Redbirds (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Illinois State Redbirds (home)**
  - Actual First-half score: **27–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Wisconsin Badgers vs Delaware Blue Hens (Final: **90–48**)

- 🟢 **WON — First-half moneyline: Wisconsin Badgers (home)**
  - Actual First-half score: **54–21** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **97.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Mississippi State Bulldogs vs Troy Trojans (Final: **79–57**)

- 🟢 **WON — Full-game moneyline: Mississippi State Bulldogs (home)**
  - Expected: **96.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Mississippi State Bulldogs (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Mississippi State Bulldogs (home)**
  - Actual First-half score: **45–14** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Vanderbilt Commodores vs Western Kentucky Lady Toppers (Final: **101–59**)

- 🟢 **WON — Full-game spread +45.5: Western Kentucky Lady Toppers (away)**
  - Actual Full-game margin: **+42** · Expected: **79.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +46.5: Western Kentucky Lady Toppers (away)**
  - Actual Full-game margin: **+42** · Expected: **82.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Vanderbilt Commodores (home)**
  - Actual First-half score: **57–25** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 140.5: Under**
  - Actual Full-game total: **160** · Expected: **58.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Under**
  - Actual Full-game total: **160** · Expected: **57.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UT Rio Grande Valley Vaqueros vs Abilene Christian Wildcats (Final: **59–76**)

- 🔴 **LOST — Full-game moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Actual First-half score: **21–41** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Abilene Christian Wildcats (away)**
  - Expected: **90.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Western Illinois Leathernecks vs Cornell (IA) Rams (Final: **74–41**)

- 🟢 **WON — Full-game moneyline: Western Illinois Leathernecks (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Illinois Leathernecks (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Western Illinois Leathernecks (home)**
  - Actual First-half score: **33–19** · Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Elon Phoenix vs East Tennessee State Buccaneers (Final: **60–65**)

- 🔴 **LOST — Full-game moneyline: Elon Phoenix (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Elon Phoenix (home)**
  - Expected: **67.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Elon Phoenix (home)**
  - Actual First-half score: **31–23** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Mercer Bears vs Charleston Southern Buccaneers (Final: **54–45**)

- 🔴 **LOST — First-half moneyline: Mercer Bears (home)**
  - Actual First-half score: **30–31** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Mercer Bears (home)**
  - Expected: **86.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Mercer Bears (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Florida Gulf Coast Eagles vs Ave Maria Gyrenes (Final: **99–31**)

- 🟢 **WON — Full-game moneyline: Florida Gulf Coast Eagles (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Florida Gulf Coast Eagles (home)**
  - Actual First-half score: **46–12** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Gulf Coast Eagles (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: UAB Blazers vs Stillman Tigers (Final: **87–47**)

- 🟢 **WON — First-half moneyline: UAB Blazers (home)**
  - Actual First-half score: **41–15** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UAB Blazers (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UAB Blazers (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Alcorn State Lady Braves vs Mississippi U for Women Owls (Final: **91–44**)

- 🟢 **WON — Full-game moneyline: Alcorn State Lady Braves (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Alcorn State Lady Braves (home)**
  - Actual First-half score: **48–21** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alcorn State Lady Braves (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: DePaul Blue Demons vs Northern Illinois Huskies (Final: **85–65**)

- 🟢 **WON — Full-game moneyline: DePaul Blue Demons (home)**
  - Expected: **90.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: DePaul Blue Demons (home)**
  - Actual First-half score: **33–32** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: DePaul Blue Demons (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Texas Longhorns vs South Dakota State Jackrabbits (Final: **93–63**)

- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **89.0%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +27.5: South Dakota State Jackrabbits (away)**
  - Actual Full-game margin: **+30** · Expected: **53.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Texas Longhorns (home)**
  - Actual First-half score: **45–31** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **98.3%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 137.5: Under**
  - Actual Full-game total: **156** · Expected: **57.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Longhorns (home)**
  - Expected: **98.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Davidson Wildcats vs App State Mountaineers (Final: **67–71**)

- 🔴 **LOST — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **74.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Davidson Wildcats (home)**
  - Actual First-half score: **35–34** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Sacred Heart Pioneers vs Northeastern Huskies (Final: **46–61**)

- 🔴 **LOST — Full-game moneyline: Sacred Heart Pioneers (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Sacred Heart Pioneers (home)**
  - Actual First-half score: **21–32** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Sacred Heart Pioneers (home)**
  - Expected: **57.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: SMU Mustangs vs Long Island University Sharks (Final: **69–50**)

- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: SMU Mustangs (home)**
  - Actual First-half score: **34–26** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **97.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Harvard Crimson vs Massachusetts Minutewomen (Final: **78–65**)

- 🟢 **WON — Full-game moneyline: Harvard Crimson (home)**
  - Expected: **86.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Harvard Crimson (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Harvard Crimson (home)**
  - Actual First-half score: **43–28** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: East Carolina Pirates vs High Point Panthers (Final: **54–67**)

- 🔴 **LOST — Full-game moneyline: East Carolina Pirates (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: East Carolina Pirates (home)**
  - Expected: **54.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: East Carolina Pirates (home)**
  - Actual First-half score: **23–32** · Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Ole Miss Rebels vs UAlbany Great Danes (Final: **95–41**)

- 🟢 **WON — Full-game total 128.5: Over**
  - Actual Full-game total: **136** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **84.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **99.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +30.5: UAlbany Great Danes (away)**
  - Actual Full-game margin: **+54** · Expected: **65.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Ole Miss Rebels (home)**
  - Actual First-half score: **45–17** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Providence Friars vs Bryant Bulldogs (Final: **47–53**)

- 🔴 **LOST — Full-game moneyline: Providence Friars (home)**
  - Expected: **62.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Providence Friars (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Providence Friars (home)**
  - Actual First-half score: **23–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Rutgers Scarlet Knights vs Monmouth Hawks (Final: **64–62**)

- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **78.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Rutgers Scarlet Knights (home)**
  - Actual First-half score: **26–38** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: NC State Wolfpack vs Liberty Flames (Final: **71–47**)

- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **93.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: NC State Wolfpack (home)**
  - Actual First-half score: **28–25** · Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Creighton Bluejays vs South Dakota Coyotes (Final: **54–49**)

- 🟢 **WON — Full-game moneyline: Creighton Bluejays (home)**
  - Expected: **81.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Creighton Bluejays (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Creighton Bluejays (home)**
  - Actual First-half score: **20–19** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: North Carolina Tar Heels vs North Carolina Central Eagles (Final: **109–40**)

- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +46.5: North Carolina Central Eagles (away)**
  - Actual Full-game margin: **+69** · Expected: **86.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: North Carolina Tar Heels (home)**
  - Actual First-half score: **54–15** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Under**
  - Actual Full-game total: **149** · Expected: **57.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 137.5: Under**
  - Actual Full-game total: **149** · Expected: **58.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Michigan State Spartans vs Oakland Golden Grizzlies (Final: **91–50**)

- 🟢 **WON — Full-game moneyline: Michigan State Spartans (home)**
  - Expected: **98.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Michigan State Spartans (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Michigan State Spartans (home)**
  - Actual First-half score: **51–26** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: La Salle Explorers vs Temple Owls (Final: **67–89**)

- 🟢 **WON — Full-game moneyline: Temple Owls (away)**
  - Expected: **70.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: La Salle Explorers (home)**
  - Actual First-half score: **29–39** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: La Salle Explorers (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Kentucky Wildcats vs Evansville Purple Aces (Final: **98–30**)

- 🟢 **WON — Full-game total 137.5: Under**
  - Actual Full-game total: **128** · Expected: **53.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +50.5: Evansville Purple Aces (away)**
  - Actual Full-game margin: **+68** · Expected: **86.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Kentucky Wildcats (home)**
  - Actual First-half score: **46–11** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: UCF Knights vs Austin Peay Governors (Final: **64–59**)

- 🔴 **LOST — Full-game moneyline: Austin Peay Governors (away)**
  - Expected: **50.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: UCF Knights (home)**
  - Actual First-half score: **39–24** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Cincinnati Bearcats vs Saint Joseph's Hawks (Final: **56–67**)

- 🟢 **WON — Full-game moneyline: Saint Joseph's Hawks (away)**
  - Expected: **73.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Cincinnati Bearcats (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Cincinnati Bearcats (home)**
  - Actual First-half score: **26–41** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Dartmouth Big Green vs New Hampshire Wildcats (Final: **51–46**)

- 🟢 **WON — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Dartmouth Big Green (home)**
  - Actual First-half score: **25–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Dartmouth Big Green (home)**
  - Expected: **63.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Georgia Lady Bulldogs vs Kennesaw State Owls (Final: **84–42**)

- 🟢 **WON — Full-game moneyline: Georgia Lady Bulldogs (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Georgia Lady Bulldogs (home)**
  - Actual First-half score: **41–22** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia Lady Bulldogs (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Georgia State Panthers vs Toccoa Falls Eagles (Final: **84–41**)

- 🟢 **WON — Full-game moneyline: Georgia State Panthers (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia State Panthers (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Georgia State Panthers (home)**
  - Actual First-half score: **45–22** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Robert Morris Colonials vs Eastern Michigan Eagles (Final: **69–51**)

- 🟢 **WON — First-half moneyline: Robert Morris Colonials (home)**
  - Actual First-half score: **29–28** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Robert Morris Colonials (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Robert Morris Colonials (home)**
  - Expected: **56.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Utah State Aggies vs Idaho State Bengals (Final: **41–54**)

- 🔴 **LOST — First-half moneyline: Utah State Aggies (home)**
  - Actual First-half score: **9–32** · Expected: **79.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Idaho State Bengals (away)**
  - Expected: **87.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Colgate Raiders vs Binghamton Bearcats (Final: **36–59**)

- 🔴 **LOST — First-half moneyline: Colgate Raiders (home)**
  - Actual First-half score: **15–33** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (away)**
  - Expected: **81.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Colgate Raiders (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Army Black Knights vs George Mason Patriots (Final: **80–75**)

- 🟢 **WON — Full-game moneyline: Army Black Knights (home)**
  - Expected: **59.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Army Black Knights (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Army Black Knights (home)**
  - Actual First-half score: **33–22** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Hofstra Pride vs Coppin State Eagles (Final: **58–54**)

- 🟢 **WON — Full-game moneyline: Hofstra Pride (home)**
  - Expected: **77.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Hofstra Pride (home)**
  - Expected: **85.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Hofstra Pride (home)**
  - Actual First-half score: **31–28** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Coastal Carolina Chanticleers vs Jacksonville Dolphins (Final: **89–78**)

- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Coastal Carolina Chanticleers (home)**
  - Actual First-half score: **48–31** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **67.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: James Madison Dukes vs Navy Midshipmen (Final: **77–61**)

- 🟢 **WON — First-half moneyline: James Madison Dukes (home)**
  - Actual First-half score: **38–34** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: James Madison Dukes (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: James Madison Dukes (home)**
  - Expected: **84.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Syracuse Orange vs Siena Saints (Final: **82–58**)

- 🟢 **WON — Full-game moneyline: Syracuse Orange (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Syracuse Orange (home)**
  - Actual First-half score: **45–33** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Syracuse Orange (home)**
  - Expected: **97.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Air Force Falcons vs San Francisco Dons (Final: **70–77**)

- 🔴 **LOST — Full-game moneyline: Air Force Falcons (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Francisco Dons (away)**
  - Expected: **60.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Air Force Falcons (home)**
  - Actual First-half score: **34–36** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Niagara Purple Eagles vs Stony Brook Seawolves (Final: **40–71**)

- 🔴 **LOST — First-half moneyline: Niagara Purple Eagles (home)**
  - Actual First-half score: **18–39** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Stony Brook Seawolves (away)**
  - Expected: **96.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Niagara Purple Eagles (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Lehigh Mountain Hawks vs Mount St. Mary's Mountaineers (Final: **79–44**)

- 🟢 **WON — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **96.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Lehigh Mountain Hawks (home)**
  - Actual First-half score: **35–22** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Virginia Cavaliers vs Drexel Dragons (Final: **81–39**)

- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **95.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Virginia Cavaliers (home)**
  - Actual First-half score: **49–22** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Duquesne Dukes vs Marist Red Foxes (Final: **72–67**)

- 🟢 **WON — First-half moneyline: Duquesne Dukes (home)**
  - Actual First-half score: **33–28** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **80.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Lafayette Leopards vs NJIT Highlanders (Final: **60–64**)

- 🟢 **WON — First-half moneyline: Lafayette Leopards (home)**
  - Actual First-half score: **36–31** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NJIT Highlanders (away)**
  - Expected: **61.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Florida State Seminoles vs West Florida Argonauts (Final: **86–58**)

- 🟢 **WON — First-half moneyline: Florida State Seminoles (home)**
  - Actual First-half score: **45–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: George Washington Revolutionaries vs Georgetown Hoyas (Final: **78–74**)

- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **56.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: George Washington Revolutionaries (home)**
  - Actual First-half score: **43–36** · Expected: **79.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Bucknell Bison vs Fordham Rams (Final: **71–78**)

- 🔴 **LOST — Full-game moneyline: Bucknell Bison (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Bucknell Bison (home)**
  - Expected: **58.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Bucknell Bison (home)**
  - Actual First-half score: **38–41** · Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Campbell Fighting Camels vs Radford Highlanders (Final: **56–83**)

- 🔴 **LOST — First-half moneyline: Campbell Fighting Camels (home)**
  - Actual First-half score: **29–42** · Expected: **79.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Campbell Fighting Camels (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Radford Highlanders (away)**
  - Expected: **82.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Gardner-Webb Runnin' Bulldogs vs UNC Greensboro Spartans (Final: **54–59**)

- 🔴 **LOST — Full-game moneyline: Gardner-Webb Runnin' Bulldogs (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Gardner-Webb Runnin' Bulldogs (home)**
  - Actual First-half score: **14–26** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNC Greensboro Spartans (away)**
  - Expected: **54.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Boston College Eagles vs UMass Lowell River Hawks (Final: **86–57**)

- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boston College Eagles (home)**
  - Actual First-half score: **41–30** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **91.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Saint Louis Billikens vs Eastern Illinois Panthers (Final: **64–58**)

- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Saint Louis Billikens (home)**
  - Actual First-half score: **33–28** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **89.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: St. Bonaventure Bonnies vs Mercyhurst Lakers (Final: **73–64**)

- 🟢 **WON — Full-game moneyline: St. Bonaventure Bonnies (home)**
  - Expected: **57.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: St. Bonaventure Bonnies (home)**
  - Actual First-half score: **29–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Bonaventure Bonnies (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Vermont Catamounts vs Oregon Ducks (Final: **56–81**)

- 🔴 **LOST — First-half moneyline: Vermont Catamounts (home)**
  - Actual First-half score: **35–38** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oregon Ducks (away)**
  - Expected: **94.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Vermont Catamounts (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: St. John's Red Storm vs Quinnipiac Bobcats (Final: **73–52**)

- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: St. John's Red Storm (home)**
  - Actual First-half score: **32–24** · Expected: **82.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **91.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Florida Atlantic Owls vs South Alabama Jaguars (Final: **85–57**)

- 🟢 **WON — Full-game total 131.5: Over**
  - Actual Full-game total: **142** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Florida Atlantic Owls (home)**
  - Actual First-half score: **43–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Florida Atlantic Owls (home)**
  - Actual Full-game margin: **+28** · Expected: **78.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **89.0%** · Price: **1.182 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **86.9%** · Price: **1.182 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **86.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: BYU Cougars vs Bethune-Cookman Wildcats (Final: **82–56**)

- 🟢 **WON — Full-game total 129.5: Over**
  - Actual Full-game total: **138** · Expected: **60.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +39.5: Bethune-Cookman Wildcats (away)**
  - Actual Full-game margin: **+26** · Expected: **68.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: BYU Cougars (home)**
  - Actual First-half score: **43–30** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: North Texas Mean Green vs Houston Christian Huskies (Final: **66–57**)

- 🟢 **WON — Full-game moneyline: North Texas Mean Green (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Texas Mean Green (home)**
  - Expected: **96.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: North Texas Mean Green (home)**
  - Actual First-half score: **30–34** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: UTSA Roadrunners vs Incarnate Word Cardinals (Final: **73–51**)

- 🟢 **WON — Full-game moneyline: UTSA Roadrunners (home)**
  - Expected: **84.0%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTSA Roadrunners (home)**
  - Expected: **90.2%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 125.5: Over**
  - Actual Full-game total: **124** · Expected: **68.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -14.5: UTSA Roadrunners (home)**
  - Actual Full-game margin: **+22** · Expected: **67.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTSA Roadrunners (home)**
  - Expected: **90.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UTSA Roadrunners (home)**
  - Expected: **84.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UTSA Roadrunners (home)**
  - Actual First-half score: **35–29** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Miami (OH) RedHawks vs Mount Vernon Nazarene MT VERNON (OH) (Final: **106–45**)

- 🟢 **WON — Full-game moneyline: Miami (OH) RedHawks (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Miami (OH) RedHawks (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Miami (OH) RedHawks (home)**
  - Actual First-half score: **55–25** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: USC Trojans vs Charleston Cougars (Final: **105–44**)

- 🟢 **WON — Full-game moneyline: USC Trojans (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 146.5: Under**
  - Actual Full-game total: **149** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: USC Trojans (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +52.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **87.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +42.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **76.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +50.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **86.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: USC Trojans (home)**
  - Actual First-half score: **58–23** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +46.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **80.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Saint Mary's Gaels vs Chico State Wildcats (Final: **81–44**)

- 🟢 **WON — First-half moneyline: Saint Mary's Gaels (home)**
  - Actual First-half score: **36–25** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Mary's Gaels (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Saint Mary's Gaels (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: UNLV Lady Rebels vs Arizona State Sun Devils (Final: **58–65**)

- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (away)**
  - Expected: **76.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: UNLV Lady Rebels (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UNLV Lady Rebels (home)**
  - Actual First-half score: **24–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: UTEP Miners vs New Mexico State Aggies (Final: **50–69**)

- 🔴 **LOST — Full-game moneyline: UTEP Miners (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UTEP Miners (home)**
  - Actual First-half score: **21–28** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UTEP Miners (home)**
  - Expected: **74.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Pepperdine Waves vs Cal State Dominguez Hills Toros (Final: **83–65**)

- 🟢 **WON — First-half moneyline: Pepperdine Waves (home)**
  - Actual First-half score: **37–28** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pepperdine Waves (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Pepperdine Waves (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Nevada Wolf Pack vs Jessup Warriors (Final: **84–51**)

- 🟢 **WON — Full-game moneyline: Nevada Wolf Pack (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Nevada Wolf Pack (home)**
  - Actual First-half score: **41–24** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nevada Wolf Pack (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: San Diego Toreros vs Portland State Vikings (Final: **70–67**)

- 🔴 **LOST — First-half moneyline: San Diego Toreros (home)**
  - Actual First-half score: **35–44** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Diego Toreros (home)**
  - Expected: **88.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: San Diego Toreros (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Boise State Broncos vs Cal State Northridge Matadors (Final: **82–59**)

- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boise State Broncos (home)**
  - Actual First-half score: **40–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **81.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Wyoming Cowgirls vs UT Arlington Mavericks (Final: **75–64**)

- 🔴 **LOST — Full-game moneyline: UT Arlington Mavericks (away)**
  - Expected: **60.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Wyoming Cowgirls (home)**
  - Actual First-half score: **32–26** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wyoming Cowgirls (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Houston Cougars vs Stephen F. Austin Ladyjacks (Final: **89–59**)

- 🔴 **LOST — Full-game moneyline: Stephen F. Austin Ladyjacks (away)**
  - Expected: **63.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Houston Cougars (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Houston Cougars (home)**
  - Actual First-half score: **48–33** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Denver Pioneers vs Colorado State Rams (Final: **50–86**)

- 🟢 **WON — Full-game moneyline: Colorado State Rams (away)**
  - Expected: **91.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Denver Pioneers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Denver Pioneers (home)**
  - Actual First-half score: **25–49** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Northern Arizona Lumberjacks vs UC Irvine Anteaters (Final: **74–76**)

- 🔴 **LOST — First-half moneyline: Northern Arizona Lumberjacks (home)**
  - Actual First-half score: **32–37** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Northern Arizona Lumberjacks (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Irvine Anteaters (away)**
  - Expected: **73.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Nebraska Cornhuskers vs North Carolina A&T Aggies (Final: **76–35**)

- 🟢 **WON — Full-game moneyline: Nebraska Cornhuskers (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Nebraska Cornhuskers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Nebraska Cornhuskers (home)**
  - Actual First-half score: **34–15** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Minnesota Golden Gophers vs Southern Jaguars (Final: **86–47**)

- 🟢 **WON — Full-game moneyline: Minnesota Golden Gophers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Minnesota Golden Gophers (home)**
  - Actual First-half score: **40–24** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Minnesota Golden Gophers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: St. Thomas Tommies vs St. Catherine Wildcats (Final: **76–29**)

- 🟢 **WON — First-half moneyline: St. Thomas Tommies (home)**
  - Actual First-half score: **40–17** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: LSU Tigers vs McNeese Cowgirls (Final: **100–43**)

- 🟢 **WON — Full-game moneyline: LSU Tigers (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: LSU Tigers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: LSU Tigers (home)**
  - Actual First-half score: **48–26** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Sam Houston Bearkats vs Jarvis Christian Bulldogs (Final: **114–31**)

- 🟢 **WON — Full-game moneyline: Sam Houston Bearkats (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Sam Houston Bearkats (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Sam Houston Bearkats (home)**
  - Actual First-half score: **61–20** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Kansas Jayhawks vs Oral Roberts Golden Eagles (Final: **99–51**)

- 🟢 **WON — First-half moneyline: Kansas Jayhawks (home)**
  - Actual First-half score: **46–21** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas Jayhawks (home)**
  - Expected: **98.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Kansas Jayhawks (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Western Michigan Broncos vs Purdue Fort Wayne Mastodons (Final: **56–61**)

- 🟢 **WON — First-half moneyline: Western Michigan Broncos (home)**
  - Actual First-half score: **32–29** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Western Michigan Broncos (home)**
  - Expected: **52.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Western Michigan Broncos (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Oklahoma Sooners vs San José State Spartans (Final: **98–44**)

- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Oklahoma Sooners (home)**
  - Actual First-half score: **50–23** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Florida International Panthers vs Stetson Hatters (Final: **63–57**)

- 🟢 **WON — First-half moneyline: Florida International Panthers (home)**
  - Actual First-half score: **33–24** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida International Panthers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida International Panthers (home)**
  - Expected: **85.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: South Carolina State Lady Bulldogs vs South Carolina Upstate Spartans (Final: **50–54**)

- 🟢 **WON — Full-game moneyline: South Carolina Upstate Spartans (away)**
  - Expected: **78.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: South Carolina State Lady Bulldogs (home)**
  - Actual First-half score: **29–25** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: South Carolina State Lady Bulldogs (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Notre Dame Fighting Irish vs Detroit Mercy Titans (Final: **108–60**)

- 🟢 **WON — First-half moneyline: Notre Dame Fighting Irish (home)**
  - Actual First-half score: **56–21** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Purdue Boilermakers vs Southeast Missouri State Redhawks (Final: **73–50**)

- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Purdue Boilermakers (home)**
  - Actual First-half score: **33–21** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **98.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Indiana State Sycamores vs IU Indianapolis Jaguars (Final: **86–74**)

- 🟢 **WON — Full-game moneyline: Indiana State Sycamores (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: IU Indianapolis Jaguars (away)**
  - Expected: **63.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Indiana State Sycamores (home)**
  - Actual First-half score: **35–45** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: West Georgia Wolves vs Fort Valley State Wildcats (Final: **89–50**)

- 🟢 **WON — First-half moneyline: West Georgia Wolves (home)**
  - Actual First-half score: **45–24** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: West Georgia Wolves (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: West Georgia Wolves (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Loyola Maryland Greyhounds vs Yale Bulldogs (Final: **67–71**)

- 🟢 **WON — Full-game moneyline: Yale Bulldogs (away)**
  - Expected: **74.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Loyola Maryland Greyhounds (home)**
  - Actual First-half score: **24–34** · Expected: **69.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Loyola Maryland Greyhounds (home)**
  - Expected: **71.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Princeton Tigers vs Villanova Wildcats (Final: **81–74**)

- 🟢 **WON — Full-game moneyline: Princeton Tigers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Villanova Wildcats (away)**
  - Expected: **50.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Princeton Tigers (home)**
  - Actual First-half score: **40–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Pittsburgh Panthers vs Delaware State Hornets (Final: **129–51**)

- 🟢 **WON — First-half moneyline: Pittsburgh Panthers (home)**
  - Actual First-half score: **62–19** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pittsburgh Panthers (home)**
  - Expected: **76.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Pittsburgh Panthers (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Louisville Cardinals vs Lindenwood Lions (Final: **71–35**)

- 🟢 **WON — Full-game moneyline: Louisville Cardinals (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Louisville Cardinals (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Louisville Cardinals (home)**
  - Actual First-half score: **47–18** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Tennessee Lady Volunteers vs New Haven Chargers (Final: **83–40**)

- 🟢 **WON — Full-game moneyline: Tennessee Lady Volunteers (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Tennessee Lady Volunteers (home)**
  - Actual First-half score: **37–24** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Lady Volunteers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Le Moyne Dolphins vs SUNY Geneseo Knights (Final: **67–52**)

- 🔴 **LOST — First-half moneyline: Le Moyne Dolphins (home)**
  - Actual First-half score: **32–35** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Le Moyne Dolphins (home)**
  - Expected: **98.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Le Moyne Dolphins (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Boston University Terriers vs UMass Boston UMASS BOSTON (Final: **88–47**)

- 🟢 **WON — First-half moneyline: Boston University Terriers (home)**
  - Actual First-half score: **50–28** · Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston University Terriers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston University Terriers (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Marshall Thundering Herd vs Howard Bison (Final: **76–66**)

- 🟢 **WON — First-half moneyline: Marshall Thundering Herd (home)**
  - Actual First-half score: **37–23** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Marshall Thundering Herd (home)**
  - Expected: **70.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Marshall Thundering Herd (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Alabama State Lady Hornets vs Talladega Tornadoes (Final: **69–63**)

- 🟢 **WON — Full-game moneyline: Alabama State Lady Hornets (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Alabama State Lady Hornets (home)**
  - Actual First-half score: **28–32** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama State Lady Hornets (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Eastern Kentucky Colonels vs Midway Eagles (Final: **94–48**)

- 🟢 **WON — Full-game moneyline: Eastern Kentucky Colonels (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Eastern Kentucky Colonels (home)**
  - Actual First-half score: **52–27** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Eastern Kentucky Colonels (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Winthrop Eagles vs Johnson C. Smith Golden Bulls (Final: **84–57**)

- 🟢 **WON — First-half moneyline: Winthrop Eagles (home)**
  - Actual First-half score: **45–29** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Winthrop Eagles (home)**
  - Expected: **99.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Winthrop Eagles (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: TCU Horned Frogs vs Iona Gaels (Final: **78–53**)

- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **131** · Expected: **64.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +43.5: Iona Gaels (away)**
  - Actual Full-game margin: **+25** · Expected: **77.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +46.5: Iona Gaels (away)**
  - Actual Full-game margin: **+25** · Expected: **80.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: TCU Horned Frogs (home)**
  - Actual First-half score: **40–30** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 132.5: Over**
  - Actual Full-game total: **131** · Expected: **63.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +44.5: Iona Gaels (away)**
  - Actual Full-game margin: **+25** · Expected: **78.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Loyola Marymount Lions vs La Sierra Golden Eagles (Final: **118–37**)

- 🟢 **WON — Full-game moneyline: Loyola Marymount Lions (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Loyola Marymount Lions (home)**
  - Actual First-half score: **61–20** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Loyola Marymount Lions (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Auburn Tigers vs Furman Paladins (Final: **74–56**)

- 🟢 **WON — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Auburn Tigers (home)**
  - Actual First-half score: **26–21** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **94.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Nicholls Colonels vs LSU Alexandria Generals (Final: **83–40**)

- 🟢 **WON — Full-game moneyline: Nicholls Colonels (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Nicholls Colonels (home)**
  - Actual First-half score: **34–12** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nicholls Colonels (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Valparaiso Beacons vs Culver Stockton Wildcats (Final: **89–47**)

- 🟢 **WON — Full-game moneyline: Valparaiso Beacons (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Valparaiso Beacons (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Valparaiso Beacons (home)**
  - Actual First-half score: **42–19** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Chicago State Cougars vs Huston Tillotson Rams (Final: **75–60**)

- 🟢 **WON — Full-game moneyline: Chicago State Cougars (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Chicago State Cougars (home)**
  - Expected: **98.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Chicago State Cougars (home)**
  - Actual First-half score: **38–28** · Expected: **71.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Drake Bulldogs vs Cleveland State Vikings (Final: **76–71**)

- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **73.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Drake Bulldogs (home)**
  - Actual First-half score: **40–35** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Marquette Golden Eagles vs Presbyterian Blue Hose (Final: **87–54**)

- 🟢 **WON — Full-game moneyline: Marquette Golden Eagles (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Marquette Golden Eagles (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Marquette Golden Eagles (home)**
  - Actual First-half score: **47–30** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Wichita State Shockers vs Missouri State Lady Bears (Final: **48–76**)

- 🔴 **LOST — Full-game moneyline: Wichita State Shockers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Missouri State Lady Bears (away)**
  - Expected: **94.1%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Wichita State Shockers (home)**
  - Actual First-half score: **20–34** · Expected: **72.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Kansas State Wildcats vs Montana Lady Griz (Final: **67–62**)

- 🟢 **WON — Full-game moneyline: Kansas State Wildcats (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Kansas State Wildcats (home)**
  - Actual First-half score: **34–25** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas State Wildcats (home)**
  - Expected: **96.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Wofford Terriers vs North Greenville Trailblazers (Final: **64–42**)

- 🟢 **WON — First-half moneyline: Wofford Terriers (home)**
  - Actual First-half score: **31–22** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Toledo Rockets vs Youngstown State Penguins (Final: **64–60**)

- 🟢 **WON — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **73.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Toledo Rockets (home)**
  - Actual First-half score: **35–26** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: North Florida Ospreys vs Edward Waters Tigers (Final: **70–43**)

- 🟢 **WON — Full-game moneyline: North Florida Ospreys (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: North Florida Ospreys (home)**
  - Actual First-half score: **26–22** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Florida Ospreys (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Bowling Green Falcons vs Bellarmine Knights (Final: **91–66**)

- 🟢 **WON — First-half moneyline: Bowling Green Falcons (home)**
  - Actual First-half score: **40–24** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **97.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
