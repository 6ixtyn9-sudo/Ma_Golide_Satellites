# Ma Golide — Recent picks audit (2026-10-17 to 2026-11-15)

Generated: 2026-11-15 18:18 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 350
- dated evidence rows: 1044 (694 superseded observations)
- settled picks: 278 (150 wins / 128 losses / 0 pushes)
- hit rate: 54.0% (150/278 win/loss decisions)
- pending/unmatched picks: 0 prior-day · 72 total
- priced picks: 211
- paper P&L: +2.577 units
- ROI: +1.2% (P&L / 211 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-15
- same-day pending rows outside prior-day counters: 72
- same-day official-final picks included: 0
- eligible prior-day picks: 278
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `CAUTION` | 107 | 66–41–0 | 61.7% | 107 | +12.0% |
| `SKIPPED_INELIGIBLE` | 104 | 57–47–0 | 54.8% | 104 | -9.8% |
| `WATCHLIST_NO_ODDS` | 67 | 27–40–0 | 40.3% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter` | 36 | 10–26–0 | 27.8% | 0 | n/a |
| `moneyline` | 98 | 63–35–0 | 64.3% | 69 | +1.4% |
| `spread` | 67 | 36–31–0 | 53.7% | 66 | +1.5% |
| `total` | 77 | 41–36–0 | 53.2% | 76 | +0.8% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `highest_quarter:Q1` | 1 | 0–1–0 | 0.0% | 0 | n/a |
| `highest_quarter:Q2` | 31 | 8–23–0 | 25.8% | 0 | n/a |
| `highest_quarter:Q3` | 4 | 2–2–0 | 50.0% | 0 | n/a |
| `moneyline:away` | 53 | 30–23–0 | 56.6% | 35 | -14.9% |
| `moneyline:home` | 45 | 33–12–0 | 73.3% | 34 | +18.2% |
| `spread:away` | 43 | 17–26–0 | 39.5% | 43 | -24.5% |
| `spread:home` | 24 | 19–5–0 | 79.2% | 23 | +50.3% |
| `total:over` | 55 | 32–23–0 | 58.2% | 55 | +11.4% |
| `total:under` | 22 | 9–13–0 | 40.9% | 21 | -27.1% |

## Paper-only opposite-side fade audit

Counterfactual reporting only — this is not a strategy, pick, or ticket. A settled result is reversed (win↔loss; push stays push) and priced only when the exact opposite quote is preserved in the same pre-tip book snapshot, market, period, line, and observation. No quote is inferred or historically backfilled; selected-side price is not required.

- settled rows with exact opposite quotes: 113/278
- fade record: 44–69–0
- fade hit rate: 38.9% (44/113 decisions)
- paper P&L: -6.095 units · ROI: -5.4% (one unit per priced fade, pushes included)
- mean market-implied break-even hit rate: 45.6% (n=113 non-push quotes; includes vig)
- unpriced settled rows: `market_has_no_single_opposite_side`=65, `opposite_quote_not_preserved`=100

| Market | Settled candidates | Exact quotes | Fade W–L–P | ROI |
| --- | ---: | ---: | ---: | ---: |
| `highest_quarter` | 36 | 0 | 0–0–0 | n/a |
| `moneyline` | 98 | 37 | 11–26–0 | +17.8% |
| `spread` | 67 | 38 | 15–23–0 | -24.3% |
| `total` | 77 | 38 | 18–20–0 | -9.1% |

Numeric selected-probability bands (calibrated probability, otherwise raw):

| Probability band | Settled candidates | Exact quotes | Fade W–L–P | ROI |
| --- | ---: | ---: | ---: | ---: |
| `0.50-<0.60` | 169 | 73 | 29–44–0 | -18.1% |
| `0.60-<0.70` | 48 | 24 | 9–15–0 | -20.9% |
| `0.70-<0.80` | 15 | 11 | 5–6–0 | +83.8% |
| `0.80-<0.90` | 6 | 4 | 0–4–0 | -100.0% |
| `0.90-1.00` | 2 | 1 | 1–0–0 | +690.0% |
| `<0.50` | 38 | 0 | 0–0–0 | n/a |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-14: Minnesota Timberwolves vs Chicago Bulls (Final: **106–117**)

- 🔴 **LOST — Full-game moneyline: Minnesota Timberwolves (home)**
  - Expected: **90.7%** · Price: **1.093 · DraftKings** · Counterfactual fade: **WON @ 7.900 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -15.5: Minnesota Timberwolves (home)**
  - Actual Full-game margin: **-11** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Chicago Bulls +15.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 234.5: Over**
  - Actual Full-game total: **223** · Expected: **54.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -14.5: Minnesota Timberwolves (home)**
  - Actual Full-game margin: **-11** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Chicago Bulls +14.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q1**
  - Expected: **27.2%** calibrated (raw 29.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Minnesota Timberwolves (home)**
  - Expected: **71.6%** calibrated (raw 69.0%) · Price: **1.093 · DraftKings** · Counterfactual fade: **WON @ 7.900 away · DraftKings**
  - Calibration: **n=2** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Minnesota Timberwolves (home)**
  - Actual First-half score: **45–58** · Expected: **58.6%** calibrated (raw 59.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 236.5: Under**
  - Actual Full-game total: **223** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-14: Memphis Grizzlies vs Brooklyn Nets (Final: **129–110**)

- 🟢 **WON — Full-game total 227.5: Over**
  - Actual Full-game total: **239** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -4.5: Memphis Grizzlies (home)**
  - Actual Full-game margin: **+19** · Expected: **52.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Brooklyn Nets +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +5.5: Brooklyn Nets (away)**
  - Actual Full-game margin: **+19** · Expected: **55.0%** · Price: **1.855 · DraftKings** · Counterfactual fade: **WON @ 1.962 Memphis Grizzlies -5.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -1.5: Memphis Grizzlies (home)**
  - Actual Full-game margin: **+19** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Brooklyn Nets +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Memphis Grizzlies (home)**
  - Expected: **66.0%** · Price: **1.769 · DraftKings** · Counterfactual fade: **LOST @ 2.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 226.5: Over**
  - Actual Full-game total: **239** · Expected: **56.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **LOST @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Memphis Grizzlies (home)**
  - Actual First-half score: **71–54** · Expected: **54.6%** calibrated (raw 55.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 228.5: Under**
  - Actual Full-game total: **239** · Expected: **52.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -2.5: Memphis Grizzlies (home)**
  - Actual Full-game margin: **+19** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Brooklyn Nets +2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Memphis Grizzlies (home)**
  - Expected: **61.6%** calibrated (raw 59.0%) · Price: **1.769 · DraftKings** · Counterfactual fade: **LOST @ 2.100 away · DraftKings**
  - Calibration: **n=2** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **35.2%** calibrated (raw 37.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 225.5: Over**
  - Actual Full-game total: **239** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-14: Miami Heat vs Detroit Pistons (Final: **137–117**)

- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **30.1%** calibrated (raw 31.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=1** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Detroit Pistons (away)**
  - Expected: **63.7%** · Price: **1.578 · DraftKings** · Counterfactual fade: **WON @ 2.460 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Miami Heat (home)**
  - Expected: **55.6%** calibrated (raw 53.0%) · Price: **2.460 · DraftKings** · Counterfactual fade: **LOST @ 1.578 away · DraftKings**
  - Calibration: **n=2** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Detroit Pistons (away)**
  - Actual First-half score: **68–60** · Expected: **53.3%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=1** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: Detroit Pistons (away)**
  - Actual Full-game margin: **+20** · Expected: **53.0%** · Price: **1.855 · DraftKings** · Counterfactual fade: **WON @ 1.962 Miami Heat +3.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 228.5: Under**
  - Actual Full-game total: **254** · Expected: **51.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread +4.5: Miami Heat (home)**
  - Actual Full-game margin: **+20** · Expected: **56.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **LOST @ 1.855 Detroit Pistons -4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q3**
  - Expected: **27.2%** calibrated (raw 29.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 229.5: Under**
  - Actual Full-game total: **254** · Expected: **54.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 228.5: Over**
  - Actual Full-game total: **254** · Expected: **53.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **LOST @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-14: Cleveland Cavaliers vs Dallas Mavericks (Final: **127–117**)

- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **29.1%** calibrated (raw 30.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=1** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Cleveland Cavaliers (home)**
  - Expected: **65.6%** · Price: **2.200 · DraftKings** · Counterfactual fade: **LOST @ 1.704 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 231.5: Under**
  - Actual Full-game total: **244** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Cleveland Cavaliers (home)**
  - Expected: **58.6%** calibrated (raw 56.0%) · Price: **2.840 · DraftKings** · Counterfactual fade: **LOST @ 1.450 away · DraftKings**
  - Calibration: **n=2** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dallas Mavericks (away)**
  - Expected: **51.0%** · Price: **1.450 · DraftKings** · Counterfactual fade: **WON @ 2.840 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -2.5: Cleveland Cavaliers (home)**
  - Actual Full-game margin: **+10** · Expected: **56.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Dallas Mavericks +2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Dallas Mavericks (away)**
  - Actual First-half score: **71–53** · Expected: **53.3%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=1** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -3.5: Cleveland Cavaliers (home)**
  - Actual Full-game margin: **+10** · Expected: **54.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **LOST @ 1.952 Dallas Mavericks +3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread +5.5: Cleveland Cavaliers (home)**
  - Actual Full-game margin: **+10** · Expected: **59.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Dallas Mavericks -5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q3**
  - Expected: **29.2%** calibrated (raw 31.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 228.5: Over**
  - Actual Full-game total: **244** · Expected: **53.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **LOST @ 1.885 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -1.5: Cleveland Cavaliers (home)**
  - Actual Full-game margin: **+10** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Dallas Mavericks +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread +2.5: Cleveland Cavaliers (home)**
  - Actual Full-game margin: **+10** · Expected: **64.0%** · Price: **1.901 · DraftKings** · Counterfactual fade: **LOST @ 1.917 Dallas Mavericks -2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-14: New Orleans Pelicans vs Denver Nuggets (Final: **108–116**)

- 🟢 **WON — Full-game spread +12.5: New Orleans Pelicans (home)**
  - Actual Full-game margin: **-8** · Expected: **53.0%** · Price: **1.901 · DraftKings** · Counterfactual fade: **LOST @ 1.917 Denver Nuggets -12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 235.5: Over**
  - Actual Full-game total: **224** · Expected: **59.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q3**
  - Expected: **28.2%** calibrated (raw 30.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +11.5: New Orleans Pelicans (home)**
  - Actual Full-game margin: **-8** · Expected: **52.0%** · Price: **1.962 · DraftKings** · Counterfactual fade: **LOST @ 1.855 Denver Nuggets -11.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Denver Nuggets (away)**
  - Expected: **57.6%** calibrated (raw 55.0%) · Price: **1.154 · DraftKings** · Counterfactual fade: **LOST @ 5.700 home · DraftKings**
  - Calibration: **n=2** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Denver Nuggets (away)**
  - Actual First-half score: **51–56** · Expected: **50.6%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=2** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 234.5: Over**
  - Actual Full-game total: **224** · Expected: **61.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Denver Nuggets (away)**
  - Expected: **80.7%** · Price: **1.154 · DraftKings** · Counterfactual fade: **LOST @ 5.700 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **29.1%** calibrated (raw 30.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=1** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: Los Angeles Clippers vs Oklahoma City Thunder (Final: **113–112**)

- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **58.1%** calibrated (raw 53.0%) · Price: **1.353 · DraftKings** · Counterfactual fade: **WON @ 3.310 home · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **30.7%** calibrated (raw 31.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Oklahoma City Thunder (away)**
  - Actual First-half score: **57–44** · Expected: **58.1%** calibrated (raw 53.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+1** · Expected: **56.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 Los Angeles Clippers +6.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+1** · Expected: **55.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **WON @ 1.893 Los Angeles Clippers +7.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 222.5: Under**
  - Actual Full-game total: **225** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **76.3%** · Price: **1.353 · DraftKings** · Counterfactual fade: **WON @ 3.310 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Sacramento Kings vs Portland Trail Blazers (Final: **107–120**)

- 🟢 **WON — Full-game moneyline: Portland Trail Blazers (away)**
  - Expected: **82.5%** · Price: **1.214 · DraftKings** · Counterfactual fade: **LOST @ 4.610 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 229.5: Under**
  - Actual Full-game total: **227** · Expected: **54.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Portland Trail Blazers (away)**
  - Actual First-half score: **53–64** · Expected: **59.1%** calibrated (raw 54.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **41.7%** calibrated (raw 42.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Portland Trail Blazers (away)**
  - Expected: **60.1%** calibrated (raw 55.0%) · Price: **1.214 · DraftKings** · Counterfactual fade: **LOST @ 4.610 home · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 228.5: Over**
  - Actual Full-game total: **227** · Expected: **51.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -10.5: Portland Trail Blazers (away)**
  - Actual Full-game margin: **-13** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Sacramento Kings +10.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: San Antonio Spurs vs Golden State Warriors (Final: **123–85**)

- 🔴 **LOST — Full-game total 217.5: Over**
  - Actual Full-game total: **208** · Expected: **67.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 219.5: Over**
  - Actual Full-game total: **208** · Expected: **63.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 215.5: Over**
  - Actual Full-game total: **208** · Expected: **74.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -11.5: San Antonio Spurs (home)**
  - Actual Full-game margin: **+38** · Expected: **54.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **LOST @ 1.935 Golden State Warriors +11.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Antonio Spurs (home)**
  - Expected: **60.1%** calibrated (raw 55.0%) · Price: **1.154 · DraftKings** · Counterfactual fade: **LOST @ 5.730 away · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **29.7%** calibrated (raw 30.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 218.5: Over**
  - Actual Full-game total: **208** · Expected: **65.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.862 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Antonio Spurs (home)**
  - Expected: **86.7%** · Price: **1.154 · DraftKings** · Counterfactual fade: **LOST @ 5.730 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 220.5: Over**
  - Actual Full-game total: **208** · Expected: **63.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **WON @ 1.885 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -12.5: San Antonio Spurs (home)**
  - Actual Full-game margin: **+38** · Expected: **52.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **LOST @ 1.847 Golden State Warriors +12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: Houston Rockets vs Utah Jazz (Final: **125–118**)

- 🟢 **WON — Full-game spread +17.5: Utah Jazz (away)**
  - Actual Full-game margin: **+7** · Expected: **55.0%** · Price: **1.855 · DraftKings** · Counterfactual fade: **LOST @ 1.962 Houston Rockets -17.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 219.5: Over**
  - Actual Full-game total: **243** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 221.5: Over**
  - Actual Full-game total: **243** · Expected: **59.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **LOST @ 1.885 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 220.5: Over**
  - Actual Full-game total: **243** · Expected: **60.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Houston Rockets (home)**
  - Expected: **88.5%** · Price: **1.075 · DraftKings** · Counterfactual fade: **LOST @ 9.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -14.5: Houston Rockets (home)**
  - Actual Full-game margin: **+7** · Expected: **52.0%** · Price: **1.855 · DraftKings** · Counterfactual fade: **WON @ 1.971 Utah Jazz +14.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 218.5: Over**
  - Actual Full-game total: **243** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Houston Rockets (home)**
  - Actual First-half score: **63–53** · Expected: **65.1%** calibrated (raw 60.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +15.5: Utah Jazz (away)**
  - Actual Full-game margin: **+7** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Houston Rockets -15.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Houston Rockets (home)**
  - Expected: **71.1%** calibrated (raw 66.0%) · Price: **1.075 · DraftKings** · Counterfactual fade: **LOST @ 9.100 away · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +16.5: Utah Jazz (away)**
  - Actual Full-game margin: **+7** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Houston Rockets -16.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: New Orleans Pelicans vs Memphis Grizzlies (Final: **111–142**)

- 🟢 **WON — Full-game moneyline: Memphis Grizzlies (away)**
  - Expected: **57.7%** · Price: **1.617 · DraftKings** · Counterfactual fade: **LOST @ 2.360 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Memphis Grizzlies (away)**
  - Actual First-half score: **58–70** · Expected: **58.1%** calibrated (raw 53.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 235.5: Over**
  - Actual Full-game total: **253** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -3.5: Memphis Grizzlies (away)**
  - Actual Full-game margin: **-31** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 New Orleans Pelicans +3.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Memphis Grizzlies (away)**
  - Expected: **60.1%** calibrated (raw 55.0%) · Price: **1.617 · DraftKings** · Counterfactual fade: **LOST @ 2.360 home · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -2.5: Memphis Grizzlies (away)**
  - Actual Full-game margin: **-31** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 New Orleans Pelicans +2.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **31.7%** calibrated (raw 32.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: Chicago Bulls vs Atlanta Hawks (Final: **125–117**)

- 🔴 **LOST — First-half moneyline: Atlanta Hawks (away)**
  - Actual First-half score: **60–53** · Expected: **58.1%** calibrated (raw 53.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 245.5: Under**
  - Actual Full-game total: **242** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.901 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **35.7%** calibrated (raw 36.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 246.5: Under**
  - Actual Full-game total: **242** · Expected: **60.0%** · Price: **1.877 · DraftKings** · Counterfactual fade: **LOST @ 1.943 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Atlanta Hawks (away)**
  - Expected: **60.1%** calibrated (raw 55.0%) · Price: **1.345 · DraftKings** · Counterfactual fade: **WON @ 3.360 home · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Atlanta Hawks (away)**
  - Expected: **75.0%** · Price: **1.345 · DraftKings** · Counterfactual fade: **WON @ 3.360 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -7.5: Atlanta Hawks (away)**
  - Actual Full-game margin: **+8** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Chicago Bulls +7.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Atlanta Hawks (away)**
  - Actual Full-game margin: **+8** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Chicago Bulls +6.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: Miami Heat vs Cleveland Cavaliers (Final: **133–121**)

- 🔴 **LOST — Full-game moneyline: Cleveland Cavaliers (away)**
  - Expected: **64.3%** · Price: **2.140 · DraftKings** · Counterfactual fade: **WON @ 1.741 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Miami Heat (home)**
  - Expected: **57.1%** calibrated (raw 52.0%) · Price: **1.741 · DraftKings** · Counterfactual fade: **LOST @ 2.140 away · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **34.7%** calibrated (raw 35.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 235.5: Over**
  - Actual Full-game total: **254** · Expected: **59.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +2.5: Cleveland Cavaliers (away)**
  - Actual Full-game margin: **+12** · Expected: **69.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Miami Heat -2.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 236.5: Over**
  - Actual Full-game total: **254** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: Orlando Magic vs Milwaukee Bucks (Final: **107–99**)

- 🟢 **WON — First-half moneyline: Orlando Magic (home)**
  - Actual First-half score: **52–37** · Expected: **64.1%** calibrated (raw 59.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **29.7%** calibrated (raw 30.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +9.5: Milwaukee Bucks (away)**
  - Actual Full-game margin: **+8** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Orlando Magic -9.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 223.5: Over**
  - Actual Full-game total: **206** · Expected: **54.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **WON @ 1.935 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (home)**
  - Expected: **77.5%** · Price: **1.253 · DraftKings** · Counterfactual fade: **LOST @ 4.120 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Orlando Magic (home)**
  - Expected: **65.1%** calibrated (raw 60.0%) · Price: **1.253 · DraftKings** · Counterfactual fade: **LOST @ 4.120 away · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 222.5: Over**
  - Actual Full-game total: **206** · Expected: **59.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-13: Indiana Pacers vs New York Knicks (Final: **106–104**)

- 🔴 **LOST — Full-game spread -7.5: New York Knicks (away)**
  - Actual Full-game margin: **+2** · Expected: **56.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **WON @ 1.847 Indiana Pacers +7.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 218.5: Under**
  - Actual Full-game total: **210** · Expected: **59.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **LOST @ 1.885 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: New York Knicks (away)**
  - Actual First-half score: **58–49** · Expected: **60.1%** calibrated (raw 55.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -8.5: New York Knicks (away)**
  - Actual Full-game margin: **+2** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Indiana Pacers +8.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 219.5: Under**
  - Actual Full-game total: **210** · Expected: **62.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **LOST @ 1.935 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **28.7%** calibrated (raw 29.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=6** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New York Knicks (away)**
  - Expected: **58.1%** calibrated (raw 53.0%) · Price: **1.297 · DraftKings** · Counterfactual fade: **WON @ 3.710 home · DraftKings**
  - Calibration: **n=6** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: New York Knicks (away)**
  - Actual Full-game margin: **+2** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Indiana Pacers +6.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New York Knicks (away)**
  - Expected: **77.4%** · Price: **1.297 · DraftKings** · Counterfactual fade: **WON @ 3.710 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-12: Phoenix Suns vs Charlotte Hornets (Final: **101–96**)

- 🟢 **WON — Full-game spread +6.5: Charlotte Hornets (away)**
  - Actual Full-game margin: **+5** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Phoenix Suns -6.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 214.5: Over**
  - Actual Full-game total: **197** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Phoenix Suns (home)**
  - Actual Full-game margin: **+5** · Expected: **51.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Charlotte Hornets +6.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Phoenix Suns (home)**
  - Expected: **60.6%** calibrated (raw 58.0%) · Price: **1.437 · DraftKings** · Counterfactual fade: **LOST @ 2.900 away · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Phoenix Suns (home)**
  - Actual First-half score: **55–50** · Expected: **56.5%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Phoenix Suns (home)**
  - Expected: **70.2%** · Price: **1.437 · DraftKings** · Counterfactual fade: **LOST @ 2.900 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **31.8%** calibrated (raw 33.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: Minnesota Timberwolves vs Dallas Mavericks (Final: **133–94**)

- 🟢 **WON — First-half moneyline: Minnesota Timberwolves (home)**
  - Actual First-half score: **67–52** · Expected: **56.5%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **31.8%** calibrated (raw 33.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 227.5: Under**
  - Actual Full-game total: **227** · Expected: **55.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Minnesota Timberwolves (home)**
  - Expected: **71.5%** · Price: **1.448 · DraftKings** · Counterfactual fade: **LOST @ 2.850 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Minnesota Timberwolves (home)**
  - Expected: **54.6%** calibrated (raw 52.0%) · Price: **1.448 · DraftKings** · Counterfactual fade: **LOST @ 2.850 away · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 225.5: Over**
  - Actual Full-game total: **227** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -5.5: Minnesota Timberwolves (home)**
  - Actual Full-game margin: **+39** · Expected: **51.0%** · Price: **1.877 · DraftKings** · Counterfactual fade: **LOST @ 1.943 Dallas Mavericks +5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-12: Washington Wizards vs Detroit Pistons (Final: **140–144**)

- 🟢 **WON — Full-game spread +10.5: Washington Wizards (home)**
  - Actual Full-game margin: **-4** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Detroit Pistons (away)**
  - Expected: **57.6%** calibrated (raw 55.0%) · Price: **1.228 · DraftKings** · Counterfactual fade: **LOST @ 4.410 home · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **29.8%** calibrated (raw 31.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 220.5: Over**
  - Actual Full-game total: **284** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Detroit Pistons (away)**
  - Actual First-half score: **61–64** · Expected: **58.5%** calibrated (raw 53.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +9.5: Washington Wizards (home)**
  - Actual Full-game margin: **-4** · Expected: **52.0%** · Price: **1.962 · DraftKings** · Counterfactual fade: **LOST @ 1.862 Detroit Pistons -9.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 221.5: Over**
  - Actual Full-game total: **284** · Expected: **51.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **LOST @ 1.885 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Detroit Pistons (away)**
  - Expected: **72.2%** · Price: **1.228 · DraftKings** · Counterfactual fade: **LOST @ 4.410 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-12: Philadelphia 76ers vs Toronto Raptors (Final: **123–112**)

- 🟢 **WON — Full-game moneyline: Philadelphia 76ers (home)**
  - Expected: **73.1%** · Price: **1.372 · DraftKings** · Counterfactual fade: **LOST @ 3.210 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **34.8%** calibrated (raw 36.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +8.5: Toronto Raptors (away)**
  - Actual Full-game margin: **+11** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 222.5: Under**
  - Actual Full-game total: **235** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Philadelphia 76ers (home)**
  - Expected: **54.6%** calibrated (raw 52.0%) · Price: **1.372 · DraftKings** · Counterfactual fade: **LOST @ 3.210 away · DraftKings**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 223.5: Under**
  - Actual Full-game total: **235** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Philadelphia 76ers (home)**
  - Actual First-half score: **60–60** · Expected: **56.5%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +9.5: Toronto Raptors (away)**
  - Actual Full-game margin: **+11** · Expected: **55.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +7.5: Toronto Raptors (away)**
  - Actual Full-game margin: **+11** · Expected: **53.0%** · Price: **1.862 · DraftKings** · Counterfactual fade: **WON @ 1.962 Philadelphia 76ers -7.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: Portland Trail Blazers vs Oklahoma City Thunder (Final: **118–108**)

- 🔴 **LOST — Full-game spread -1.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+10** · Expected: **52.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Portland Trail Blazers (home)**
  - Expected: **54.0%** · Price: **1.680 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Oklahoma City Thunder (away)**
  - Actual First-half score: **56–50** · Expected: **49.1%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 224.5: Under**
  - Actual Full-game total: **226** · Expected: **66.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +2.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+10** · Expected: **54.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **46.3%** calibrated (raw 37.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **55.2%** · Price: **1.885 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Portland Trail Blazers (home)**
  - Expected: **55.7%** calibrated (raw 52.0%) · Price: **1.680 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 225.5: Under**
  - Actual Full-game total: **226** · Expected: **66.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: Houston Rockets vs Golden State Warriors (Final: **115–108**)

- 🟢 **WON — Full-game moneyline: Houston Rockets (home)**
  - Expected: **66.9%** · Price: **1.402 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +6.5: Golden State Warriors (away)**
  - Actual Full-game margin: **+7** · Expected: **53.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 219.5: Over**
  - Actual Full-game total: **223** · Expected: **61.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q3**
  - Expected: **42.3%** calibrated (raw 33.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 220.5: Over**
  - Actual Full-game total: **223** · Expected: **58.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 222.5: Over**
  - Actual Full-game total: **223** · Expected: **56.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Houston Rockets (home)**
  - Actual First-half score: **56–51** · Expected: **50.1%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Utah Jazz vs Orlando Magic (Final: **108–123**)

- 🟢 **WON — Full-game total 214.5: Over**
  - Actual Full-game total: **231** · Expected: **75.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **49.3%** calibrated (raw 40.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 217.5: Over**
  - Actual Full-game total: **231** · Expected: **67.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **61.3%** · Price: **1.559 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -3.5: Orlando Magic (away)**
  - Actual Full-game margin: **-15** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -4.5: Orlando Magic (away)**
  - Actual Full-game margin: **-15** · Expected: **52.0%** · Price: **1.962 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **54.7%** calibrated (raw 51.0%) · Price: **1.559 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-11: New Orleans Pelicans vs Brooklyn Nets (Final: **111–145**)

- 🟢 **WON — Full-game total 229.5: Over**
  - Actual Full-game total: **256** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New Orleans Pelicans (home)**
  - Expected: **56.7%** calibrated (raw 53.0%) · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +4.5: Brooklyn Nets (away)**
  - Actual Full-game margin: **-34** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: New Orleans Pelicans (home)**
  - Actual First-half score: **66–83** · Expected: **50.1%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Brooklyn Nets (away)**
  - Actual Full-game margin: **-34** · Expected: **53.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 228.5: Over**
  - Actual Full-game total: **256** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **44.3%** calibrated (raw 35.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New Orleans Pelicans (home)**
  - Expected: **58.3%** · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Milwaukee Bucks vs Los Angeles Lakers (Final: **112–123**)

- 🔴 **LOST — Full-game spread -12.5: Los Angeles Lakers (away)**
  - Actual Full-game margin: **-11** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Los Angeles Lakers (away)**
  - Expected: **85.8%** · Price: **1.127 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **45.3%** calibrated (raw 36.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -14.5: Los Angeles Lakers (away)**
  - Actual Full-game margin: **-11** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Los Angeles Lakers (away)**
  - Expected: **61.7%** calibrated (raw 58.0%) · Price: **1.127 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Los Angeles Lakers (away)**
  - Actual First-half score: **50–61** · Expected: **59.1%** calibrated (raw 61.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 228.5: Over**
  - Actual Full-game total: **235** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: Memphis Grizzlies vs Denver Nuggets (Final: **118–123**)

- 🟢 **WON — Full-game spread -3.5: Denver Nuggets (away)**
  - Actual Full-game margin: **-5** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Memphis Grizzlies (home)**
  - Expected: **56.7%** calibrated (raw 53.0%) · Price: **2.390 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Denver Nuggets (away)**
  - Expected: **59.0%** · Price: **1.602 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.3%** calibrated (raw 35.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 230.5: Over**
  - Actual Full-game total: **241** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 229.5: Over**
  - Actual Full-game total: **241** · Expected: **57.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: Chicago Bulls vs San Antonio Spurs (Final: **119–124**)

- 🔴 **LOST — First-half moneyline: San Antonio Spurs (away)**
  - Actual First-half score: **70–56** · Expected: **60.1%** calibrated (raw 62.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 240.5: Under**
  - Actual Full-game total: **243** · Expected: **62.0%** · Price: **1.901 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -15.5: San Antonio Spurs (away)**
  - Actual Full-game margin: **-5** · Expected: **53.0%** · Price: **1.917 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **50.3%** calibrated (raw 41.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Antonio Spurs (away)**
  - Expected: **88.5%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 238.5: Under**
  - Actual Full-game total: **243** · Expected: **52.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Antonio Spurs (away)**
  - Expected: **65.7%** calibrated (raw 62.0%) · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Toronto Raptors vs Boston Celtics (Final: **93–98**)

- 🔴 **LOST — Full-game spread -5.5: Boston Celtics (away)**
  - Actual Full-game margin: **-5** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 215.5: Over**
  - Actual Full-game total: **191** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston Celtics (away)**
  - Expected: **78.1%** · Price: **1.450 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.3%** calibrated (raw 35.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: Boston Celtics (away)**
  - Actual Full-game margin: **-5** · Expected: **56.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston Celtics (away)**
  - Expected: **56.7%** calibrated (raw 53.0%) · Price: **1.450 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Boston Celtics (away)**
  - Actual First-half score: **61–54** · Expected: **54.1%** calibrated (raw 56.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 216.5: Over**
  - Actual Full-game total: **191** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: New York Knicks vs Cleveland Cavaliers (Final: **107–115**)

- 🟢 **WON — Full-game spread +4.5: Cleveland Cavaliers (away)**
  - Actual Full-game margin: **-8** · Expected: **52.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **41.3%** calibrated (raw 32.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New York Knicks (home)**
  - Expected: **65.5%** · Price: **1.559 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 222.5: Over**
  - Actual Full-game total: **222** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: New York Knicks (home)**
  - Expected: **57.7%** calibrated (raw 54.0%) · Price: **1.559 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: New York Knicks (home)**
  - Actual First-half score: **53–59** · Expected: **56.1%** calibrated (raw 58.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 221.5: Over**
  - Actual Full-game total: **222** · Expected: **56.0%** · Price: **1.943 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Cleveland Cavaliers (away)**
  - Actual Full-game margin: **-8** · Expected: **51.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-11: Indiana Pacers vs Washington Wizards (Final: **114–101**)

- 🔴 **LOST — Full-game spread +2.5: Washington Wizards (away)**
  - Actual Full-game margin: **+13** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +3.5: Washington Wizards (away)**
  - Actual Full-game margin: **+13** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 225.5: Over**
  - Actual Full-game total: **215** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 227.5: Over**
  - Actual Full-game total: **215** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.3%** calibrated (raw 35.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=5** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Indiana Pacers (home)**
  - Expected: **52.2%** · Price: **1.699 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Indiana Pacers (home)**
  - Expected: **55.7%** calibrated (raw 52.0%) · Price: **1.699 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=5** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`

### 2026-11-10: Portland Trail Blazers vs Oklahoma City Thunder (Final: **123–103**)

- 🔴 **LOST — Full-game total 225.5: Under**
  - Actual Full-game total: **226** · Expected: **62.0%** · Price: **1.885 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 224.5: Under**
  - Actual Full-game total: **226** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+20** · Expected: **54.0%** · Price: **1.901 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 226.5: Under**
  - Actual Full-game total: **226** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **1.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -2.5: Oklahoma City Thunder (away)**
  - Actual Full-game margin: **+20** · Expected: **52.0%** · Price: **1.962 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Oklahoma City Thunder (away)**
  - Actual First-half score: **59–59** · Expected: **48.6%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Oklahoma City Thunder (away)**
  - Expected: **60.0%** · Price: **1.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Sacramento Kings vs Orlando Magic (Final: **100–107**)

- 🟢 **WON — First-half moneyline: Orlando Magic (away)**
  - Actual First-half score: **52–57** · Expected: **47.6%** calibrated (raw 51.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **55.6%** calibrated (raw 48.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **1.461 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Orlando Magic (away)**
  - Expected: **71.6%** · Price: **1.461 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -6.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **56.0%** · Price: **1.962 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 223.5: Over**
  - Actual Full-game total: **207** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game total 222.5: Over**
  - Actual Full-game total: **207** · Expected: **60.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **53.0%** · Price: **1.971 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game spread -5.5: Orlando Magic (away)**
  - Actual Full-game margin: **-7** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Utah Jazz vs Miami Heat (Final: **109–115**)

- 🟢 **WON — Full-game spread +8.5: Utah Jazz (home)**
  - Actual Full-game margin: **-6** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Heat (away)**
  - Expected: **61.2%** · Price: **1.279 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +9.5: Utah Jazz (home)**
  - Actual Full-game margin: **-6** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Miami Heat (away)**
  - Actual First-half score: **54–64** · Expected: **48.6%** calibrated (raw 52.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 227.5: Over**
  - Actual Full-game total: **224** · Expected: **71.0%** · Price: **1.917 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Heat (away)**
  - Expected: **54.4%** calibrated (raw 52.0%) · Price: **1.279 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 228.5: Over**
  - Actual Full-game total: **224** · Expected: **67.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Phoenix Suns vs Los Angeles Clippers (Final: **133–110**)

- 🔴 **LOST — Full-game moneyline: Los Angeles Clippers (away)**
  - Expected: **53.4%** calibrated (raw 51.0%) · Price: **2.450 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **44.6%** calibrated (raw 37.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Phoenix Suns (home)**
  - Expected: **65.5%** · Price: **1.578 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 224.5: Over**
  - Actual Full-game total: **243** · Expected: **56.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Los Angeles Clippers (away)**
  - Actual First-half score: **64–47** · Expected: **50.6%** calibrated (raw 54.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 223.5: Over**
  - Actual Full-game total: **243** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 225.5: Over**
  - Actual Full-game total: **243** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

### 2026-11-10: Philadelphia 76ers vs Boston Celtics (Final: **101–123**)

- 🔴 **LOST — Full-game moneyline: Philadelphia 76ers (home)**
  - Expected: **52.0%** · Price: **1.990 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +1.5: Philadelphia 76ers (home)**
  - Actual Full-game margin: **-22** · Expected: **51.0%** · Price: **1.935 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston Celtics (away)**
  - Expected: **54.4%** calibrated (raw 52.0%) · Price: **1.855 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `CAUTION` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 218.5: Over**
  - Actual Full-game total: **224** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game total 217.5: Over**
  - Actual Full-game total: **224** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Highest quarter: Q2**
  - Expected: **48.6%** calibrated (raw 41.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 216.5: Over**
  - Actual Full-game total: **224** · Expected: **58.0%** · Price: **1.901 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `CAUTION` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boston Celtics (away)**
  - Actual First-half score: **51–62** · Expected: **52.6%** calibrated (raw 56.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: New York Knicks vs Milwaukee Bucks (Final: **117–85**)

- 🔴 **LOST — Full-game total 219.5: Over**
  - Actual Full-game total: **202** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: New York Knicks (home)**
  - Actual First-half score: **58–34** · Expected: **56.6%** calibrated (raw 60.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -15.5: New York Knicks (home)**
  - Actual Full-game margin: **+32** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New York Knicks (home)**
  - Expected: **90.6%** · Price: **1.096 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 220.5: Over**
  - Actual Full-game total: **202** · Expected: **55.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New York Knicks (home)**
  - Expected: **67.4%** calibrated (raw 65.0%) · Price: **1.096 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=4** · provisional · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Highest quarter: Q2**
  - Expected: **45.6%** calibrated (raw 38.0%) · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=4** · provisional · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 221.5: Under**
  - Actual Full-game total: **202** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn+team_history_empirical_v1`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
