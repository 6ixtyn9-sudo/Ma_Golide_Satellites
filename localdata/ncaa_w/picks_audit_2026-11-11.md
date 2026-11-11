# Ma Golide — Recent picks audit (2026-10-13 to 2026-11-11)

Generated: 2026-11-11 18:55 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 400
- dated evidence rows: 716 (316 superseded observations)
- settled picks: 166 (135 wins / 31 losses / 0 pushes)
- hit rate: 81.3% (135/166 win/loss decisions)
- pending/unmatched picks: 0 prior-day · 234 total
- priced picks: 10
- paper P&L: -4.273 units
- ROI: -42.7% (P&L / 10 priced picks) · **provisional — fewer than 30 priced picks**

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-11
- same-day pending rows outside prior-day counters: 234
- same-day official-final picks included: 3
- eligible prior-day picks: 163
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `SKIPPED_INELIGIBLE` | 10 | 3–7–0 | 30.0% | 10 | -42.7% |
| `WATCHLIST_NO_ODDS` | 156 | 132–24–0 | 84.6% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 156 | 132–24–0 | 84.6% | 0 | n/a |
| `spread` | 7 | 3–4–0 | 42.9% | 7 | -18.2% |
| `total` | 3 | 0–3–0 | 0.0% | 3 | -100.0% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 10 | 6–4–0 | 60.0% | 0 | n/a |
| `moneyline:home` | 146 | 126–20–0 | 86.3% | 0 | n/a |
| `spread:away` | 7 | 3–4–0 | 42.9% | 7 | -18.2% |
| `total:over` | 2 | 0–2–0 | 0.0% | 2 | -100.0% |
| `total:under` | 1 | 0–1–0 | 0.0% | 1 | -100.0% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

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
- 🔴 **LOST — Full-game spread -52.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **87.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -42.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **76.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -50.5: Charleston Cougars (away)**
  - Actual Full-game margin: **+61** · Expected: **86.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: USC Trojans (home)**
  - Actual First-half score: **58–23** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -46.5: Charleston Cougars (away)**
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
- 🟢 **WON — Full-game spread -43.5: Iona Gaels (away)**
  - Actual Full-game margin: **+25** · Expected: **77.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -46.5: Iona Gaels (away)**
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
- 🟢 **WON — Full-game spread -44.5: Iona Gaels (away)**
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
