# Ma Golide — Recent picks audit (2026-10-12 to 2026-11-10)

Generated: 2026-11-11 03:37 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 413
- dated evidence rows: 1358 (945 superseded observations)
- settled picks: 45 (32 wins / 13 losses / 0 pushes)
- hit rate: 71.1% (32/45 win/loss decisions)
- pending/unmatched picks: 368 prior-day · 368 total
- priced picks: 31
- paper P&L: +2.975 units
- ROI: +9.6% (P&L / 31 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-11
- same-day pending rows outside prior-day counters: 0
- same-day official-final picks included: 0
- eligible prior-day picks: 413
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `SKIPPED_INELIGIBLE` | 31 | 21–10–0 | 67.7% | 31 | +9.6% |
| `WATCHLIST_NO_ODDS` | 14 | 11–3–0 | 78.6% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 24 | 18–6–0 | 75.0% | 10 | -26.1% |
| `spread` | 11 | 4–7–0 | 36.4% | 11 | -31.9% |
| `total` | 10 | 10–0–0 | 100.0% | 10 | +90.9% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 1 | 1–0–0 | 100.0% | 1 | +20.0% |
| `moneyline:home` | 23 | 17–6–0 | 73.9% | 9 | -31.2% |
| `spread:away` | 6 | 2–4–0 | 33.3% | 6 | -36.9% |
| `spread:home` | 5 | 2–3–0 | 40.0% | 5 | -25.9% |
| `total:over` | 4 | 4–0–0 | 100.0% | 4 | +91.0% |
| `total:under` | 6 | 6–0–0 | 100.0% | 6 | +90.9% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

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

### 2026-11-10: Lafayette Leopards vs Penn State Lehigh Valley Nittany Lions (Final: **89–46**)

- 🟢 **WON — First-half moneyline: Lafayette Leopards (home)**
  - Actual First-half score: **52–14** · Expected: **75.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

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
