# Ma Golide — Recent picks audit (2026-10-13 to 2026-11-11)

Generated: 2026-11-12 04:14 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 123
- dated evidence rows: 290 (167 superseded observations)
- settled picks: 56 (30 wins / 26 losses / 0 pushes)
- hit rate: 53.6% (30/56 win/loss decisions)
- pending/unmatched picks: 67 prior-day · 67 total
- priced picks: 43
- paper P&L: -3.472 units
- ROI: -8.1% (P&L / 43 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-12
- same-day pending rows outside prior-day counters: 0
- same-day official-final picks included: 0
- eligible prior-day picks: 123
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `CAUTION` | 15 | 8–7–0 | 53.3% | 15 | -3.6% |
| `SKIPPED_INELIGIBLE` | 28 | 15–13–0 | 53.6% | 28 | -10.5% |
| `WATCHLIST_NO_ODDS` | 13 | 7–6–0 | 53.8% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter` | 7 | 3–4–0 | 42.9% | 0 | n/a |
| `moneyline` | 20 | 14–6–0 | 70.0% | 14 | +3.6% |
| `spread` | 11 | 5–6–0 | 45.5% | 11 | -12.5% |
| `total` | 18 | 8–10–0 | 44.4% | 18 | -14.5% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter:Q2` | 7 | 3–4–0 | 42.9% | 0 | n/a |
| `moneyline:away` | 13 | 8–5–0 | 61.5% | 8 | -8.3% |
| `moneyline:home` | 7 | 6–1–0 | 85.7% | 6 | +19.5% |
| `spread:away` | 7 | 2–5–0 | 28.6% | 7 | -44.9% |
| `spread:home` | 4 | 3–1–0 | 75.0% | 4 | +44.4% |
| `total:over` | 14 | 6–8–0 | 42.9% | 14 | -17.6% |
| `total:under` | 4 | 2–2–0 | 50.0% | 4 | -3.5% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-11: Indiana Pacers vs Washington Wizards (Final: **114–101**)

- 🔴 **LOST — Full-game spread -2.5: Washington Wizards (away)**
  - Actual Full-game margin: **+13** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: Washington Wizards (away)**
  - Actual Full-game margin: **+13** · Expected: **56.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 225.5: Over**
  - Actual Full-game total: **215** · Expected: **57.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 227.5: Over**
  - Actual Full-game total: **215** · Expected: **53.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.3%** calibrated (raw 35.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Indiana Pacers (home)**
  - Expected: **52.2%** · Price: **1.699 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Indiana Pacers (home)**
  - Expected: **55.7%** calibrated (raw 52.0%) · Price: **1.699 · DraftKings**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-10: Portland Trail Blazers vs Oklahoma City Thunder (Final: **123–103**)

- 🔴 **LOST — Full-game total 225.5: Under**
  - Actual Full-game total: **226** · Expected: **62.0%** · Price: **1.885 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 224.5: Under**
  - Actual Full-game total: **226** · Expected: **61.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+20** · Expected: **54.0%** · Price: **1.901 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 226.5: Under**
  - Actual Full-game total: **226** · Expected: **64.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **1.800 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +2.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+20** · Expected: **52.0%** · Price: **1.962 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Oklahoma City Thunder (away)**
  - Actual First-half score: **59–59** · Expected: **48.6%** calibrated (raw 52.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **60.0%** · Price: **1.800 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Sacramento Kings vs Orlando Magic (Final: **100–107**)

- 🟢 **WON — First-half moneyline: Orlando Magic (away)**
  - Actual First-half score: **52–57** · Expected: **47.6%** calibrated (raw 51.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **55.6%** calibrated (raw 48.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **1.461 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **71.6%** · Price: **1.461 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +6.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **56.0%** · Price: **1.962 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 223.5: Over**
  - Actual Full-game total: **207** · Expected: **58.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 222.5: Over**
  - Actual Full-game total: **207** · Expected: **60.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +7.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **53.0%** · Price: **1.971 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread +5.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **55.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Utah Jazz vs Miami Heat (Final: **109–115**)

- 🟢 **WON — Full-game spread +8.5: Utah Jazz (home)**
  - Actual Full-game margin: **-6** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Heat (away)**
  - Expected: **61.2%** · Price: **1.279 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +9.5: Utah Jazz (home)**
  - Actual Full-game margin: **-6** · Expected: **64.0%** · Price: **1.870 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Miami Heat (away)**
  - Actual First-half score: **54–64** · Expected: **48.6%** calibrated (raw 52.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 227.5: Over**
  - Actual Full-game total: **224** · Expected: **71.0%** · Price: **1.917 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Heat (away)**
  - Expected: **54.4%** calibrated (raw 52.0%) · Price: **1.279 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 228.5: Over**
  - Actual Full-game total: **224** · Expected: **67.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Phoenix Suns vs Los Angeles Clippers (Final: **133–110**)

- 🔴 **LOST — Full-game moneyline: Los Angeles Clippers (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **2.450 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.6%** calibrated (raw 37.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Phoenix Suns (home)**
  - Expected: **65.5%** · Price: **1.578 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 224.5: Over**
  - Actual Full-game total: **243** · Expected: **56.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Los Angeles Clippers (away)**
  - Actual First-half score: **64–47** · Expected: **50.6%** calibrated (raw 54.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 223.5: Over**
  - Actual Full-game total: **243** · Expected: **57.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 225.5: Over**
  - Actual Full-game total: **243** · Expected: **52.0%** · Price: **1.909 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Philadelphia 76ers vs Boston Celtics (Final: **101–123**)

- 🔴 **LOST — Full-game moneyline: Philadelphia 76ers (home)**
  - Expected: **52.0%** · Price: **1.990 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +1.5: Philadelphia 76ers (home)**
  - Actual Full-game margin: **-22** · Expected: **51.0%** · Price: **1.935 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston Celtics (away)**
  - Expected: **54.4%** calibrated (raw 52.0%) · Price: **1.855 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 218.5: Over**
  - Actual Full-game total: **224** · Expected: **56.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 217.5: Over**
  - Actual Full-game total: **224** · Expected: **57.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **48.6%** calibrated (raw 41.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 216.5: Over**
  - Actual Full-game total: **224** · Expected: **58.0%** · Price: **1.901 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boston Celtics (away)**
  - Actual First-half score: **51–62** · Expected: **52.6%** calibrated (raw 56.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: New York Knicks vs Milwaukee Bucks (Final: **117–85**)

- 🔴 **LOST — Full-game total 219.5: Over**
  - Actual Full-game total: **202** · Expected: **56.0%** · Price: **1.893 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: New York Knicks (home)**
  - Actual First-half score: **58–34** · Expected: **56.6%** calibrated (raw 60.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -15.5: New York Knicks (home)**
  - Actual Full-game margin: **+32** · Expected: **51.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New York Knicks (home)**
  - Expected: **90.6%** · Price: **1.096 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 220.5: Over**
  - Actual Full-game total: **202** · Expected: **55.0%** · Price: **1.926 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New York Knicks (home)**
  - Expected: **67.4%** calibrated (raw 65.0%) · Price: **1.096 · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 221.5: Under**
  - Actual Full-game total: **202** · Expected: **51.0%** · Price: **1.952 · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
