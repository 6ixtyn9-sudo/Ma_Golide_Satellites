# Ma Golide — Recent picks audit (2026-10-13 to 2026-11-11)

Generated: 2026-11-11 12:46 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 708
- dated evidence rows: 1653 (945 superseded observations)
- settled picks: 413 (260 wins / 153 losses / 0 pushes)
- hit rate: 63.0% (260/413 win/loss decisions)
- pending/unmatched picks: 0 prior-day · 295 total
- priced picks: 308
- paper P&L: -27.329 units
- ROI: -8.9% (P&L / 308 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-11
- same-day pending rows outside prior-day counters: 295
- same-day official-final picks included: 0
- eligible prior-day picks: 413
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `SKIPPED_INELIGIBLE` | 308 | 176–132–0 | 57.1% | 308 | -8.9% |
| `WATCHLIST_NO_ODDS` | 105 | 84–21–0 | 80.0% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 207 | 164–43–0 | 79.2% | 102 | -4.0% |
| `spread` | 80 | 34–46–0 | 42.5% | 80 | -19.6% |
| `total` | 126 | 62–64–0 | 49.2% | 126 | -6.0% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 8 | 3–5–0 | 37.5% | 8 | -48.8% |
| `moneyline:home` | 199 | 161–38–0 | 80.9% | 94 | -0.2% |
| `spread:away` | 39 | 17–22–0 | 43.6% | 39 | -17.6% |
| `spread:home` | 41 | 17–24–0 | 41.5% | 41 | -21.5% |
| `total:over` | 67 | 31–36–0 | 46.3% | 67 | -11.5% |
| `total:under` | 59 | 31–28–0 | 52.5% | 59 | +0.3% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-10: San Francisco Dons vs Montana Grizzlies (Final: **63–58**)

- 🔴 **LOST — Full-game spread -13.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **54.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Francisco Dons (home)**
  - Expected: **69.0%** · Price: **1.125 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **121** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -11.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **58.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: San Francisco Dons (home)**
  - Actual First-half score: **26–28** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **121** · Expected: **58.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Francisco Dons (home)**
  - Expected: **77.0%** · Price: **1.125 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -12.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **57.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Kentucky Wildcats vs Kansas Jayhawks (Final: **69–89**)

- 🔴 **LOST — First-half moneyline: Kentucky Wildcats (home)**
  - Actual First-half score: **36–46** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **158** · Expected: **57.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **158** · Expected: **64.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **158** · Expected: **58.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **73.0%** · Price: **2.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Kentucky Wildcats (home)**
  - Actual Full-game margin: **-20** · Expected: **80.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas Jayhawks (away)**
  - Expected: **64.1%** · Price: **1.459 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Fresno State Bulldogs vs Georgia State Panthers (Final: **95–92**)

- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **187** · Expected: **61.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fresno State Bulldogs (home)**
  - Expected: **55.4%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **187** · Expected: **58.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **187** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Fresno State Bulldogs (home)**
  - Actual First-half score: **39–38** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Fresno State Bulldogs (home)**
  - Actual Full-game margin: **+3** · Expected: **62.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: Fresno State Bulldogs (home)**
  - Actual Full-game margin: **+3** · Expected: **64.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fresno State Bulldogs (home)**
  - Expected: **70.0%** · Price: **1.388 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Utah State Aggies vs Utah Tech Trailblazers (Final: **118–86**)

- 🟢 **WON — First-half moneyline: Utah State Aggies (home)**
  - Actual First-half score: **53–35** · Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -25.5: Utah Tech Trailblazers (away)**
  - Actual Full-game margin: **+32** · Expected: **66.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Under**
  - Actual Full-game total: **204** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **204** · Expected: **63.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **204** · Expected: **57.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -27.5: Utah Tech Trailblazers (away)**
  - Actual Full-game margin: **+32** · Expected: **67.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **94.9%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **204** · Expected: **59.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **80.0%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: New Mexico Lobos vs Little Rock Trojans (Final: **94–69**)

- 🟢 **WON — First-half moneyline: New Mexico Lobos (home)**
  - Actual First-half score: **47–34** · Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **163** · Expected: **63.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -23.5: Little Rock Trojans (away)**
  - Actual Full-game margin: **+25** · Expected: **64.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -22.5: Little Rock Trojans (away)**
  - Actual Full-game margin: **+25** · Expected: **62.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Mexico Lobos (home)**
  - Expected: **94.6%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **163** · Expected: **64.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Mexico Lobos (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **163** · Expected: **58.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: BYU Cougars vs Queens University Royals (Final: **98–82**)

- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **83.0%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 181.5: Under**
  - Actual Full-game total: **180** · Expected: **90.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -28.5: Queens University Royals (away)**
  - Actual Full-game margin: **+16** · Expected: **66.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **97.6%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: BYU Cougars (home)**
  - Actual First-half score: **54–42** · Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -29.5: Queens University Royals (away)**
  - Actual Full-game margin: **+16** · Expected: **67.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 182.5: Under**
  - Actual Full-game total: **180** · Expected: **92.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arizona State Sun Devils vs Hofstra Pride (Final: **85–63**)

- 🔴 **LOST — Full-game spread -18.5: Hofstra Pride (away)**
  - Actual Full-game margin: **+22** · Expected: **51.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arizona State Sun Devils (home)**
  - Actual First-half score: **41–32** · Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **86.8%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 133.5: Over**
  - Actual Full-game total: **148** · Expected: **77.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **80.0%** · Price: **1.031 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Utah Utes vs Weber State Wildcats (Final: **85–64**)

- 🟢 **WON — Full-game moneyline: Utah Utes (home)**
  - Expected: **77.0%** · Price: **1.028 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Utah Utes (home)**
  - Actual First-half score: **30–37** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Over**
  - Actual Full-game total: **149** · Expected: **53.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -18.5: Weber State Wildcats (away)**
  - Actual Full-game margin: **+21** · Expected: **55.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah Utes (home)**
  - Expected: **93.8%** · Price: **1.028 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 155.5: Over**
  - Actual Full-game total: **149** · Expected: **51.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Virginia Cavaliers vs Bucknell Bison (Final: **95–61**)

- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **156** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Virginia Cavaliers (home)**
  - Actual First-half score: **53–31** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **99.2%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -38.5: Bucknell Bison (away)**
  - Actual Full-game margin: **+34** · Expected: **76.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **156** · Expected: **55.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **156** · Expected: **57.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: TCU Horned Frogs vs UL Monroe Warhawks (Final: **100–92**)

- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -33.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **68.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -35.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **192** · Expected: **65.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **192** · Expected: **65.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: TCU Horned Frogs (home)**
  - Actual First-half score: **41–32** · Expected: **83.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -34.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **69.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Colorado Buffaloes vs San Diego Toreros (Final: **83–75**)

- 🟢 **WON — Full-game moneyline: Colorado Buffaloes (home)**
  - Expected: **93.2%** · Price: **1.095 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **158** · Expected: **55.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Colorado Buffaloes (home)**
  - Actual First-half score: **34–37** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Colorado Buffaloes (home)**
  - Expected: **74.0%** · Price: **1.095 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **158** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -12.5: Colorado Buffaloes (home)**
  - Actual Full-game margin: **+8** · Expected: **53.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arizona Wildcats vs Northern Arizona Lumberjacks (Final: **90–59**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **149** · Expected: **61.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona Wildcats (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -35.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+31** · Expected: **74.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -36.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+31** · Expected: **74.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **149** · Expected: **58.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arizona Wildcats (home)**
  - Actual First-half score: **46–22** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona Wildcats (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Kansas State Wildcats vs Creighton Bluejays (Final: **89–82**)

- 🟢 **WON — Full-game spread +3.5: Kansas State Wildcats (home)**
  - Actual Full-game margin: **+7** · Expected: **81.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas State Wildcats (home)**
  - Expected: **79.0%** · Price: **2.400 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **171** · Expected: **64.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **171** · Expected: **62.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Kansas State Wildcats (home)**
  - Actual First-half score: **40–40** · Expected: **79.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Creighton Bluejays (away)**
  - Expected: **69.5%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +2.5: Kansas State Wildcats (home)**
  - Actual Full-game margin: **+7** · Expected: **81.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Wisconsin Badgers vs Fairfield Stags (Final: **89–42**)

- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **94.5%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **77.0%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Wisconsin Badgers (home)**
  - Actual First-half score: **52–20** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -25.5: Fairfield Stags (away)**
  - Actual Full-game margin: **+47** · Expected: **67.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **131** · Expected: **51.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Ole Miss Rebels vs SE Louisiana Lions (Final: **91–51**)

- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **142** · Expected: **60.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -25.5: SE Louisiana Lions (away)**
  - Actual Full-game margin: **+40** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **142** · Expected: **61.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **142** · Expected: **58.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Ole Miss Rebels (home)**
  - Actual First-half score: **49–20** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **97.7%** · Price: **1.011 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -26.5: SE Louisiana Lions (away)**
  - Actual Full-game margin: **+40** · Expected: **67.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **80.0%** · Price: **1.011 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Omaha Mavericks vs York (NE) Panthers (Final: **78–66**)

- 🔴 **LOST — First-half moneyline: Omaha Mavericks (home)**
  - Actual First-half score: **28–30** · Expected: **66.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Omaha Mavericks (home)**
  - Expected: **70.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Omaha Mavericks (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Tarleton State Texans vs UT Dallas Comets (Final: **76–57**)

- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Tarleton State Texans (home)**
  - Actual First-half score: **37–22** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Missouri State Bears vs Newman Jets (Final: **89–62**)

- 🟢 **WON — Full-game moneyline: Missouri State Bears (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Missouri State Bears (home)**
  - Actual First-half score: **41–29** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Missouri State Bears (home)**
  - Expected: **79.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Northern Iowa Panthers vs Green Bay Phoenix (Final: **89–76**)

- 🟢 **WON — First-half moneyline: Northern Iowa Panthers (home)**
  - Actual First-half score: **46–30** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Iowa Panthers (home)**
  - Expected: **64.8%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **165** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **165** · Expected: **55.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Iowa Panthers (home)**
  - Expected: **74.0%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -3.5: Northern Iowa Panthers (home)**
  - Actual Full-game margin: **+13** · Expected: **69.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arkansas State Red Wolves vs UT Martin Skyhawks (Final: **65–82**)

- 🔴 **LOST — Full-game total 153.5: Over**
  - Actual Full-game total: **147** · Expected: **53.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Arkansas State Red Wolves (home)**
  - Expected: **74.0%** · Price: **1.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Arkansas State Red Wolves (home)**
  - Actual First-half score: **34–54** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Arkansas State Red Wolves (home)**
  - Actual Full-game margin: **-17** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: Arkansas State Red Wolves (home)**
  - Actual Full-game margin: **-17** · Expected: **68.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Arkansas State Red Wolves (home)**
  - Expected: **77.4%** · Price: **1.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Northern Colorado Bears vs Northern New Mexico Eagles (Final: **102–52**)

- 🟢 **WON — Full-game moneyline: Northern Colorado Bears (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Northern Colorado Bears (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northern Colorado Bears (home)**
  - Actual First-half score: **48–17** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Murray State Racers vs Alcorn State Braves (Final: **94–57**)

- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **151** · Expected: **72.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **96.2%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **80.0%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -30.5: Alcorn State Braves (away)**
  - Actual Full-game margin: **+37** · Expected: **69.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Murray State Racers (home)**
  - Actual First-half score: **51–30** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **151** · Expected: **70.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **151** · Expected: **72.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Texas A&M-Corpus Christi Islanders vs Dallas Christian Crusaders (Final: **102–58**)

- 🟢 **WON — Full-game moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Actual First-half score: **46–28** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: South Dakota Coyotes vs Houston Christian Huskies (Final: **76–87**)

- 🔴 **LOST — Full-game moneyline: South Dakota Coyotes (home)**
  - Expected: **57.9%** · Price: **1.617 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: South Dakota Coyotes (home)**
  - Expected: **74.0%** · Price: **1.617 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **163** · Expected: **57.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: South Dakota Coyotes (home)**
  - Actual Full-game margin: **-11** · Expected: **70.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: South Dakota Coyotes (home)**
  - Actual First-half score: **34–44** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **163** · Expected: **55.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Lipscomb Bisons vs Belmont Bruins (Final: **88–75**)

- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **163** · Expected: **62.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lipscomb Bisons (home)**
  - Expected: **55.3%** · Price: **2.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Lipscomb Bisons (home)**
  - Actual First-half score: **40–32** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **163** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Lipscomb Bisons (home)**
  - Actual Full-game margin: **+13** · Expected: **81.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +4.5: Lipscomb Bisons (home)**
  - Actual Full-game margin: **+13** · Expected: **81.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **163** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lipscomb Bisons (home)**
  - Expected: **77.0%** · Price: **2.420 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: SMU Mustangs vs Prairie View A&M Panthers (Final: **98–63**)

- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **161** · Expected: **69.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **97.6%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **77.0%** · Price: **1.005 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: SMU Mustangs (home)**
  - Actual First-half score: **52–36** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -26.5: Prairie View A&M Panthers (away)**
  - Actual Full-game margin: **+35** · Expected: **67.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Austin Peay Governors vs Livingstone Blue Bears (Final: **82–61**)

- 🟢 **WON — Full-game moneyline: Austin Peay Governors (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Austin Peay Governors (home)**
  - Actual First-half score: **46–31** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Austin Peay Governors (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Texas Tech Red Raiders vs Illinois Fighting Illini (Final: **84–78**)

- 🟢 **WON — First-half moneyline: Texas Tech Red Raiders (home)**
  - Actual First-half score: **49–29** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Tech Red Raiders (home)**
  - Expected: **77.0%** · Price: **2.200 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 167.5: Under**
  - Actual Full-game total: **162** · Expected: **74.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Illinois Fighting Illini (away)**
  - Expected: **55.8%** · Price: **1.704 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **162** · Expected: **73.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +2.5: Texas Tech Red Raiders (home)**
  - Actual Full-game margin: **+6** · Expected: **80.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Saint Louis Billikens vs Northern Kentucky Norse (Final: **125–63**)

- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **94.5%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 169.5: Under**
  - Actual Full-game total: **188** · Expected: **75.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -23.5: Northern Kentucky Norse (away)**
  - Actual Full-game margin: **+62** · Expected: **64.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **80.0%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Saint Louis Billikens (home)**
  - Actual First-half score: **56–24** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 168.5: Under**
  - Actual Full-game total: **188** · Expected: **74.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Iowa State Cyclones vs Southern Jaguars (Final: **84–47**)

- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **131** · Expected: **59.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -35.5: Southern Jaguars (away)**
  - Actual Full-game margin: **+37** · Expected: **70.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Iowa State Cyclones (home)**
  - Actual First-half score: **41–18** · Expected: **82.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Iowa State Cyclones (home)**
  - Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Iowa State Cyclones (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **131** · Expected: **57.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Baylor Bears vs Yale Bulldogs (Final: **90–61**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **151** · Expected: **59.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Baylor Bears (home)**
  - Actual First-half score: **41–30** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Baylor Bears (home)**
  - Expected: **78.0%** · Price: **1.025 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Baylor Bears (home)**
  - Expected: **91.3%** · Price: **1.025 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -19.5: Yale Bulldogs (away)**
  - Actual Full-game margin: **+29** · Expected: **60.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **151** · Expected: **57.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Middle Tennessee Blue Raiders vs Evansville Purple Aces (Final: **73–68**)

- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **72.0%** · Price: **1.227 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **141** · Expected: **61.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **82.8%** · Price: **1.227 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -9.5: Middle Tennessee Blue Raiders (home)**
  - Actual Full-game margin: **+5** · Expected: **57.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Middle Tennessee Blue Raiders (home)**
  - Actual First-half score: **27–20** · Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: UT Rio Grande Valley Vaqueros vs Texas State Bobcats (Final: **65–72**)

- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **137** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Actual First-half score: **28–30** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **137** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **137** · Expected: **67.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas State Bobcats (away)**
  - Expected: **72.3%** · Price: **1.435 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +5.5: UT Rio Grande Valley Vaqueros (home)**
  - Actual Full-game margin: **-7** · Expected: **84.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Expected: **77.0%** · Price: **2.900 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: UT Rio Grande Valley Vaqueros (home)**
  - Actual Full-game margin: **-7** · Expected: **83.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Northwestern State Demons vs Millsaps Majors (Final: **77–49**)

- 🟢 **WON — Full-game moneyline: Northwestern State Demons (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northwestern State Demons (home)**
  - Actual First-half score: **34–21** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northwestern State Demons (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Tulane Green Wave vs Nicholls Colonels (Final: **77–63**)

- 🟢 **WON — First-half moneyline: Tulane Green Wave (home)**
  - Actual First-half score: **35–24** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulane Green Wave (home)**
  - Expected: **77.0%** · Price: **1.227 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulane Green Wave (home)**
  - Expected: **67.9%** · Price: **1.227 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **140** · Expected: **52.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Tulane Green Wave (home)**
  - Actual Full-game margin: **+14** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **140** · Expected: **53.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **140** · Expected: **51.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Central Arkansas Bears vs Ecclesia Royals (Final: **127–54**)

- 🟢 **WON — First-half moneyline: Central Arkansas Bears (home)**
  - Actual First-half score: **71–23** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Central Arkansas Bears (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Central Arkansas Bears (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Vanderbilt Commodores vs Memphis Tigers (Final: **105–81**)

- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **85.6%** · Price: **1.062 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **77.0%** · Price: **1.062 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **186** · Expected: **65.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -13.5: Vanderbilt Commodores (home)**
  - Actual Full-game margin: **+24** · Expected: **53.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Vanderbilt Commodores (home)**
  - Actual First-half score: **47–42** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Texas Southern Tigers vs College of Biblical Studies Ambassadors (Final: **98–78**)

- 🟢 **WON — First-half moneyline: Texas Southern Tigers (home)**
  - Actual First-half score: **42–37** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Southern Tigers (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Texas Southern Tigers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Toledo Rockets vs Wright State Raiders (Final: **68–76**)

- 🔴 **LOST — First-half moneyline: Toledo Rockets (home)**
  - Actual First-half score: **31–36** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **144** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **54.8%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +1.5: Toledo Rockets (home)**
  - Actual Full-game margin: **-8** · Expected: **77.0%** · Price: **1.820 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **74.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Hampton Pirates vs Apprentice School Builders (Final: **67–56**)

- 🟢 **WON — Full-game moneyline: Hampton Pirates (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Hampton Pirates (home)**
  - Actual First-half score: **33–25** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hampton Pirates (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: The Citadel Bulldogs vs Virginia Lynchburg Dragons (Final: **131–50**)

- 🟢 **WON — First-half moneyline: The Citadel Bulldogs (home)**
  - Actual First-half score: **66–20** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: The Citadel Bulldogs (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: The Citadel Bulldogs (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Duquesne Dukes vs Washington State Cougars (Final: **72–71**)

- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **80.0%** · Price: **1.617 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **143** · Expected: **58.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: Duquesne Dukes (home)**
  - Actual Full-game margin: **+1** · Expected: **73.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **143** · Expected: **55.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Duquesne Dukes (home)**
  - Actual First-half score: **30–39** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **59.7%** · Price: **1.617 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **143** · Expected: **61.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Davidson Wildcats vs Bradley Braves (Final: **62–45**)

- 🟢 **WON — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **81.8%** · Price: **1.160 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **80.0%** · Price: **1.160 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -10.5: Davidson Wildcats (home)**
  - Actual Full-game margin: **+17** · Expected: **64.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **107** · Expected: **68.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **107** · Expected: **70.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Davidson Wildcats (home)**
  - Actual First-half score: **30–20** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **107** · Expected: **72.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Kennesaw State Owls vs Winthrop Eagles (Final: **88–79**)

- 🟢 **WON — Full-game total 168.5: Under**
  - Actual Full-game total: **167** · Expected: **73.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **81.0%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 166.5: Under**
  - Actual Full-game total: **167** · Expected: **72.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -10.5: Kennesaw State Owls (home)**
  - Actual Full-game margin: **+9** · Expected: **60.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 170.5: Under**
  - Actual Full-game total: **167** · Expected: **76.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **77.0%** · Price: **1.174 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Kennesaw State Owls (home)**
  - Actual First-half score: **32–33** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Lafayette Leopards vs Penn State Lehigh Valley Nittany Lions (Final: **89–46**)

- 🟢 **WON — First-half moneyline: Lafayette Leopards (home)**
  - Actual First-half score: **54–16** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Longwood Lancers vs St. Mary's (MD) Seahawks (Final: **101–45**)

- 🟢 **WON — First-half moneyline: Longwood Lancers (home)**
  - Actual First-half score: **48–23** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Longwood Lancers (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Longwood Lancers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Villanova Wildcats vs La Salle Explorers (Final: **80–59**)

- 🟢 **WON — Full-game moneyline: Villanova Wildcats (home)**
  - Expected: **95.7%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Villanova Wildcats (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **139** · Expected: **63.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -20.5: La Salle Explorers (away)**
  - Actual Full-game margin: **+21** · Expected: **61.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Villanova Wildcats (home)**
  - Actual First-half score: **36–31** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -23.5: La Salle Explorers (away)**
  - Actual Full-game margin: **+21** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: UCF Knights vs South Alabama Jaguars (Final: **76–69**)

- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **78.0%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UCF Knights (home)**
  - Actual First-half score: **32–34** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **145** · Expected: **55.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **89.8%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **145** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: UCF Knights (home)**
  - Actual Full-game margin: **+7** · Expected: **54.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgia Southern Eagles vs Mercer Bears (Final: **91–97**)

- 🔴 **LOST — Full-game total 165.5: Under**
  - Actual Full-game total: **188** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 164.5: Under**
  - Actual Full-game total: **188** · Expected: **68.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Georgia Southern Eagles (home)**
  - Expected: **59.9%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Georgia Southern Eagles (home)**
  - Expected: **77.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Georgia Southern Eagles (home)**
  - Actual First-half score: **46–39** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Georgia Southern Eagles (home)**
  - Actual Full-game margin: **-6** · Expected: **73.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Iowa Hawkeyes vs Virginia Tech Hokies (Final: **88–93**)

- 🔴 **LOST — Full-game spread -4.5: Iowa Hawkeyes (home)**
  - Actual Full-game margin: **-5** · Expected: **73.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **181** · Expected: **67.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Iowa Hawkeyes (home)**
  - Expected: **80.0%** · Price: **1.505 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **181** · Expected: **66.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Iowa Hawkeyes (home)**
  - Actual First-half score: **40–41** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Iowa Hawkeyes (home)**
  - Expected: **67.3%** · Price: **1.505 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **181** · Expected: **64.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Northeastern Huskies vs Harvard Crimson (Final: **74–70**)

- 🟢 **WON — First-half moneyline: Northeastern Huskies (home)**
  - Actual First-half score: **29–19** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **144** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **144** · Expected: **56.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: Northeastern Huskies (home)**
  - Actual Full-game margin: **+4** · Expected: **77.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Harvard Crimson (away)**
  - Expected: **52.2%** · Price: **1.769 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Northeastern Huskies (home)**
  - Expected: **74.0%** · Price: **2.100 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **144** · Expected: **61.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgia Tech Yellow Jackets vs Eastern Kentucky Colonels (Final: **85–71**)

- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **156** · Expected: **66.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia Tech Yellow Jackets (home)**
  - Expected: **85.5%** · Price: **1.056 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Georgia Tech Yellow Jackets (home)**
  - Expected: **73.0%** · Price: **1.056 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **156** · Expected: **66.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **156** · Expected: **69.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -16.5: Eastern Kentucky Colonels (away)**
  - Actual Full-game margin: **+14** · Expected: **55.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Georgia Tech Yellow Jackets (home)**
  - Actual First-half score: **41–42** · Expected: **73.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -15.5: Eastern Kentucky Colonels (away)**
  - Actual Full-game margin: **+14** · Expected: **51.0%** · Price: **1.833 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Rhode Island Rams vs Maine Black Bears (Final: **71–65**)

- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **80.0%** · Price: **1.022 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -20.5: Maine Black Bears (away)**
  - Actual Full-game margin: **+6** · Expected: **61.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **136** · Expected: **66.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **93.6%** · Price: **1.022 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Rhode Island Rams (home)**
  - Actual First-half score: **23–28** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: NC State Wolfpack vs Presbyterian Blue Hose (Final: **94–60**)

- 🟢 **WON — First-half moneyline: NC State Wolfpack (home)**
  - Actual First-half score: **46–19** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **154** · Expected: **57.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -19.5: Presbyterian Blue Hose (away)**
  - Actual Full-game margin: **+34** · Expected: **61.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **154** · Expected: **55.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **77.0%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **96.7%** · Price: **1.015 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Troy Trojans vs Loyola Marymount Lions (Final: **74–69**)

- 🟢 **WON — Full-game moneyline: Troy Trojans (home)**
  - Expected: **71.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Troy Trojans (home)**
  - Expected: **51.6%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +1.5: Troy Trojans (home)**
  - Actual Full-game margin: **+5** · Expected: **73.0%** · Price: **1.833 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Troy Trojans (home)**
  - Actual First-half score: **29–37** · Expected: **67.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **143** · Expected: **60.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **143** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: North Carolina Tar Heels vs Wyoming Cowboys (Final: **76–63**)

- 🟢 **WON — Full-game spread -12.5: North Carolina Tar Heels (home)**
  - Actual Full-game margin: **+13** · Expected: **53.0%** · Price: **1.833 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **73.0%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: North Carolina Tar Heels (home)**
  - Actual Full-game margin: **+13** · Expected: **51.0%** · Price: **2.000 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **89.0%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **139** · Expected: **55.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: North Carolina Tar Heels (home)**
  - Actual First-half score: **31–29** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: New Haven Chargers vs Army Black Knights (Final: **63–55**)

- 🟢 **WON — Full-game spread +1.5: New Haven Chargers (home)**
  - Actual Full-game margin: **+8** · Expected: **77.0%** · Price: **1.847 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Haven Chargers (home)**
  - Expected: **73.0%** · Price: **2.000 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 133.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Army Black Knights (away)**
  - Expected: **56.2%** · Price: **1.833 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 132.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: New Haven Chargers (home)**
  - Actual First-half score: **24–19** · Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: George Washington Revolutionaries vs Elon Phoenix (Final: **69–50**)

- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **84.3%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: George Washington Revolutionaries (home)**
  - Actual First-half score: **23–22** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **119** · Expected: **55.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **119** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -13.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+19** · Expected: **53.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -14.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+19** · Expected: **53.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **77.0%** · Price: **1.083 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Fordham Rams vs Iona Gaels (Final: **78–72**)

- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **150** · Expected: **69.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Fordham Rams (home)**
  - Actual First-half score: **43–37** · Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **150** · Expected: **71.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **80.0%** · Price: **1.125 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **79.7%** · Price: **1.125 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -11.5: Fordham Rams (home)**
  - Actual Full-game margin: **+6** · Expected: **64.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Dayton Flyers vs Siena Saints (Final: **89–99**)

- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **188** · Expected: **65.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: Dayton Flyers (home)**
  - Actual Full-game margin: **-10** · Expected: **51.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **188** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Dayton Flyers (home)**
  - Actual First-half score: **45–46** · Expected: **68.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dayton Flyers (home)**
  - Expected: **92.0%** · Price: **1.091 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Dayton Flyers (home)**
  - Expected: **74.0%** · Price: **1.091 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Central Michigan Chippewas vs Youngstown State Penguins (Final: **91–78**)

- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **169** · Expected: **52.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Central Michigan Chippewas (home)**
  - Actual First-half score: **50–34** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **169** · Expected: **54.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Youngstown State Penguins (away)**
  - Expected: **52.1%** · Price: **2.400 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -3.5: Central Michigan Chippewas (home)**
  - Actual Full-game margin: **+13** · Expected: **71.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Central Michigan Chippewas (home)**
  - Expected: **77.0%** · Price: **1.602 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Wake Forest Demon Deacons vs Charleston Southern Buccaneers (Final: **82–64**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **146** · Expected: **63.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **146** · Expected: **53.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wake Forest Demon Deacons (home)**
  - Expected: **77.0%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **146** · Expected: **58.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Wake Forest Demon Deacons (home)**
  - Actual First-half score: **38–22** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wake Forest Demon Deacons (home)**
  - Expected: **93.5%** · Price: **1.020 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -18.5: Charleston Southern Buccaneers (away)**
  - Actual Full-game margin: **+18** · Expected: **58.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **146** · Expected: **56.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Duke Blue Devils vs Michigan State Spartans (Final: **73–79**)

- 🔴 **LOST — First-half moneyline: Duke Blue Devils (home)**
  - Actual First-half score: **25–34** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **152** · Expected: **59.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **77.0%** · Price: **1.294 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **152** · Expected: **61.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **73.4%** · Price: **1.294 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -7.5: Duke Blue Devils (home)**
  - Actual Full-game margin: **-6** · Expected: **68.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Boston College Eagles vs Maryland Eastern Shore Hawks (Final: **82–54**)

- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **73.0%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 130.5: Over**
  - Actual Full-game total: **136** · Expected: **80.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 129.5: Over**
  - Actual Full-game total: **136** · Expected: **82.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **96.9%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 131.5: Over**
  - Actual Full-game total: **136** · Expected: **77.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 128.5: Over**
  - Actual Full-game total: **136** · Expected: **84.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boston College Eagles (home)**
  - Actual First-half score: **43–25** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -27.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+28** · Expected: **68.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -26.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+28** · Expected: **68.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Xavier Musketeers vs Detroit Mercy Titans (Final: **100–55**)

- 🔴 **LOST — Full-game spread -34.5: Detroit Mercy Titans (away)**
  - Actual Full-game margin: **+45** · Expected: **70.0%** · Price: **1.980 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Xavier Musketeers (home)**
  - Expected: **81.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Xavier Musketeers (home)**
  - Expected: **98.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **155** · Expected: **65.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Xavier Musketeers (home)**
  - Actual First-half score: **43–20** · Expected: **77.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **155** · Expected: **65.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -35.5: Detroit Mercy Titans (away)**
  - Actual Full-game margin: **+45** · Expected: **73.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Providence Friars vs Bryant Bulldogs (Final: **78–55**)

- 🟢 **WON — Full-game moneyline: Providence Friars (home)**
  - Expected: **77.0%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -30.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+23** · Expected: **69.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **133** · Expected: **55.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -31.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+23** · Expected: **73.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Providence Friars (home)**
  - Expected: **97.7%** · Price: **1.002 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **133** · Expected: **52.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Providence Friars (home)**
  - Actual First-half score: **38–25** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Rutgers Scarlet Knights vs NJIT Highlanders (Final: **101–64**)

- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **80.0%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -25.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+37** · Expected: **66.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **165** · Expected: **60.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **97.5%** · Price: **1.012 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **165** · Expected: **58.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Rutgers Scarlet Knights (home)**
  - Actual First-half score: **46–26** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -24.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+37** · Expected: **65.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgetown Hoyas vs Delaware State Hornets (Final: **83–78**)

- 🟢 **WON — Full-game spread -29.5: Delaware State Hornets (away)**
  - Actual Full-game margin: **+5** · Expected: **68.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Georgetown Hoyas (home)**
  - Actual First-half score: **33–43** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -30.5: Delaware State Hornets (away)**
  - Actual Full-game margin: **+5** · Expected: **69.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgetown Hoyas (home)**
  - Expected: **78.0%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **161** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgetown Hoyas (home)**
  - Expected: **97.2%** · Price: **1.001 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Florida State Seminoles vs Florida Gators (Final: **72–81**)

- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **153** · Expected: **72.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +8.5: Florida State Seminoles (home)**
  - Actual Full-game margin: **-9** · Expected: **87.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **153** · Expected: **73.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **77.0%** · Price: **4.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Florida State Seminoles (home)**
  - Actual First-half score: **29–31** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +9.5: Florida State Seminoles (home)**
  - Actual Full-game margin: **-9** · Expected: **89.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Gators (away)**
  - Expected: **86.2%** · Price: **1.200 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
