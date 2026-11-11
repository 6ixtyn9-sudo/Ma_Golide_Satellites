# Ma Golide — Recent picks audit (2026-10-12 to 2026-11-10)

Generated: 2026-11-11 06:21 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 49
- dated evidence rows: 150 (101 superseded observations)
- settled picks: 16 (11 wins / 5 losses / 0 pushes)
- hit rate: 68.8% (11/16 win/loss decisions)
- pending/unmatched picks: 33 prior-day · 33 total
- priced picks: 12
- paper P&L: +1.757 units
- ROI: +14.6% (P&L / 12 priced picks) · **provisional — fewer than 30 priced picks**

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-11
- same-day pending rows outside prior-day counters: 0
- same-day official-final picks included: 0
- eligible prior-day picks: 49
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `CAUTION` | 5 | 4–1–0 | 80.0% | 5 | +53.2% |
| `SKIPPED_INELIGIBLE` | 7 | 4–3–0 | 57.1% | 7 | -12.9% |
| `WATCHLIST_NO_ODDS` | 4 | 3–1–0 | 75.0% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter` | 2 | 1–1–0 | 50.0% | 0 | n/a |
| `moneyline` | 6 | 5–1–0 | 83.3% | 4 | +1.2% |
| `spread` | 2 | 1–1–0 | 50.0% | 2 | -2.4% |
| `total` | 6 | 4–2–0 | 66.7% | 6 | +29.3% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter:Q2` | 2 | 1–1–0 | 50.0% | 0 | n/a |
| `moneyline:away` | 2 | 2–0–0 | 100.0% | 1 | +85.5% |
| `moneyline:home` | 4 | 3–1–0 | 75.0% | 3 | -26.9% |
| `spread:home` | 2 | 1–1–0 | 50.0% | 2 | -2.4% |
| `total:over` | 5 | 3–2–0 | 60.0% | 5 | +16.1% |
| `total:under` | 1 | 1–0–0 | 100.0% | 1 | +95.2% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

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
