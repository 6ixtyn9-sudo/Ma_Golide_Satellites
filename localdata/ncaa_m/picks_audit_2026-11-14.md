# Ma Golide — Recent picks audit (2026-10-16 to 2026-11-14)

Generated: 2026-11-14 08:17 SAST · 30-day rolling window · Paper only

## Overall

- archived logical picks: 1720
- dated evidence rows: 4446 (2726 superseded observations)
- settled picks: 1404 (906 wins / 498 losses / 0 pushes)
- hit rate: 64.5% (906/1404 win/loss decisions)
- pending/unmatched picks: 0 prior-day · 316 total
- priced picks: 1055
- paper P&L: -24.338 units
- ROI: -2.3% (P&L / 1055 priced picks)

## Settlement policy

- reconciled official finals settle immediately; pending rows are never losses
- local-day pending cutoff: 2026-11-14
- same-day pending rows outside prior-day counters: 316
- same-day official-final picks included: 0
- eligible prior-day picks: 1404
- latest observation per logical pick; superseded dated evidence remains in JSON

## Bucket performance

| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `SKIPPED_INELIGIBLE` | 1055 | 622–433–0 | 59.0% | 1055 | -2.3% |
| `WATCHLIST_NO_ODDS` | 349 | 284–65–0 | 81.4% | 0 | n/a |

## By market

| Market | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline` | 683 | 525–158–0 | 76.9% | 335 | -8.3% |
| `spread` | 311 | 165–146–0 | 53.1% | 311 | +0.9% |
| `total` | 410 | 216–194–0 | 52.7% | 409 | +0.2% |

## By selection — over/under & side bias

| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |
| --- | ---: | ---: | ---: | ---: | ---: |
| `moneyline:away` | 34 | 17–17–0 | 50.0% | 33 | -31.4% |
| `moneyline:home` | 649 | 508–141–0 | 78.3% | 302 | -5.8% |
| `spread:away` | 133 | 81–52–0 | 60.9% | 133 | +16.0% |
| `spread:home` | 178 | 84–94–0 | 47.2% | 178 | -10.5% |
| `total:over` | 224 | 104–120–0 | 46.4% | 223 | -12.0% |
| `total:under` | 186 | 112–74–0 | 60.2% | 186 | +14.8% |

## Paper-only opposite-side fade audit

Counterfactual reporting only — this is not a strategy, pick, or ticket. A settled result is reversed (win↔loss; push stays push) and priced only when the exact opposite quote is preserved in the same pre-tip book snapshot, market, period, line, and observation. No quote is inferred or historically backfilled; selected-side price is not required.

- settled rows with exact opposite quotes: 434/1404
- fade record: 162–272–0
- fade hit rate: 37.3% (162/434 decisions)
- paper P&L: -97.953 units · ROI: -22.6% (one unit per priced fade, pushes included)
- mean market-implied break-even hit rate: 43.3% (n=434 non-push quotes; includes vig)
- unpriced settled rows: `market_has_no_single_opposite_side`=226, `no_book_evidence`=122, `opposite_quote_not_preserved`=622

| Market | Settled candidates | Exact quotes | Fade W–L–P | ROI |
| --- | ---: | ---: | ---: | ---: |
| `moneyline` | 683 | 144 | 41–103–0 | -27.1% |
| `spread` | 311 | 131 | 48–83–0 | -30.2% |
| `total` | 410 | 159 | 73–86–0 | -12.2% |

Numeric selected-probability bands (calibrated probability, otherwise raw):

| Probability band | Settled candidates | Exact quotes | Fade W–L–P | ROI |
| --- | ---: | ---: | ---: | ---: |
| `0.50-<0.60` | 254 | 105 | 42–63–0 | -23.3% |
| `0.60-<0.70` | 284 | 99 | 54–45–0 | +8.3% |
| `0.70-<0.80` | 487 | 141 | 46–95–0 | -27.9% |
| `0.80-<0.90` | 246 | 65 | 20–45–0 | -28.2% |
| `0.90-1.00` | 133 | 24 | 0–24–0 | -100.0% |

## Settled picks — official-final audit

Grouped by fixture. Every settled logical pick in the window is included; unpriced grades do not enter ROI.

### 2026-11-13: California Baptist Lancers vs Arkansas-Pine Bluff Golden Lions (Final: **74–60**)

- 🟢 **WON — Full-game spread +16.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+14** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 California Baptist Lancers -16.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +17.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+14** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 California Baptist Lancers -17.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **134** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **134** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **134** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +15.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+14** · Expected: **51.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 California Baptist Lancers -15.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: California Baptist Lancers (home)**
  - Actual First-half score: **33–27** · Expected: **76.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: California Baptist Lancers (home)**
  - Expected: **85.2%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: California Baptist Lancers (home)**
  - Expected: **76.0%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Utah State Aggies vs UC San Diego Tritons (Final: **120–82**)

- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **84.8%** · Price: **1.100 · DraftKings** · Counterfactual fade: **LOST @ 7.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **202** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **202** · Expected: **60.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Utah State Aggies (home)**
  - Actual First-half score: **73–34** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -12.5: Utah State Aggies (home)**
  - Actual Full-game margin: **+38** · Expected: **55.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 UC San Diego Tritons +12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **75.0%** · Price: **1.100 · DraftKings** · Counterfactual fade: **LOST @ 7.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **202** · Expected: **65.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **202** · Expected: **68.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -13.5: Utah State Aggies (home)**
  - Actual Full-game margin: **+38** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 UC San Diego Tritons +13.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Portland Pilots vs East Texas A&M Lions (Final: **88–94**)

- 🔴 **LOST — Full-game spread -7.5: Portland Pilots (home)**
  - Actual Full-game margin: **-6** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 East Texas A&M Lions +7.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Portland Pilots (home)**
  - Expected: **74.2%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Portland Pilots (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Portland Pilots (home)**
  - Actual First-half score: **36–31** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **182** · Expected: **52.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **182** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Portland Pilots (home)**
  - Expected: **75.0%** · Price: **1.323 · DraftKings** · Counterfactual fade: **WON @ 3.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Portland Pilots (home)**
  - Expected: **74.2%** · Price: **1.323 · DraftKings** · Counterfactual fade: **WON @ 3.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 151.5: Under**
  - Actual Full-game total: **182** · Expected: **52.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Saint Mary's Gaels vs UIC Flames (Final: **84–61**)

- 🟢 **WON — Full-game moneyline: Saint Mary's Gaels (home)**
  - Expected: **76.6%** · Price: **1.360 · DraftKings** · Counterfactual fade: **LOST @ 3.250 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -5.5: Saint Mary's Gaels (home)**
  - Actual Full-game margin: **+23** · Expected: **65.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 UIC Flames +5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Saint Mary's Gaels (home)**
  - Actual First-half score: **42–24** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Mary's Gaels (home)**
  - Expected: **74.0%** · Price: **1.360 · DraftKings** · Counterfactual fade: **LOST @ 3.250 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **145** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **145** · Expected: **52.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -6.5: Saint Mary's Gaels (home)**
  - Actual Full-game margin: **+23** · Expected: **62.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 UIC Flames +6.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Oregon State Beavers vs Mississippi Valley State Delta Devils (Final: **81–65**)

- 🟢 **WON — Full-game moneyline: Oregon State Beavers (home)**
  - Expected: **98.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Oregon State Beavers (home)**
  - Expected: **76.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **146** · Expected: **54.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +33.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+16** · Expected: **78.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Oregon State Beavers -33.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Oregon State Beavers (home)**
  - Actual First-half score: **45–31** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +34.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+16** · Expected: **78.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Oregon State Beavers -34.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **146** · Expected: **66.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: San Diego Toreros vs Chicago State Cougars (Final: **87–67**)

- 🟢 **WON — Full-game moneyline: San Diego Toreros (home)**
  - Expected: **79.5%** · Price: **1.048 · DraftKings** · Counterfactual fade: **LOST @ 12.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: San Diego Toreros (home)**
  - Actual First-half score: **46–29** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +13.5: Chicago State Cougars (away)**
  - Actual Full-game margin: **+20** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 San Diego Toreros -13.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Diego Toreros (home)**
  - Expected: **71.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **LOST @ 12.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **154** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +14.5: Chicago State Cougars (away)**
  - Actual Full-game margin: **+20** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 San Diego Toreros -14.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **154** · Expected: **72.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **LOST @ 1.847 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **154** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **154** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +16.5: Chicago State Cougars (away)**
  - Actual Full-game margin: **+20** · Expected: **63.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 San Diego Toreros -16.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: UNLV Rebels vs Stanford Cardinal (Final: **76–68**)

- 🟢 **WON — First-half moneyline: UNLV Rebels (home)**
  - Actual First-half score: **39–33** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **144** · Expected: **51.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNLV Rebels (home)**
  - Expected: **57.6%** · Price: **2.000 · DraftKings** · Counterfactual fade: **LOST @ 1.833 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UNLV Rebels (home)**
  - Expected: **74.0%** · Price: **2.000 · DraftKings** · Counterfactual fade: **LOST @ 1.833 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: UNLV Rebels (home)**
  - Actual Full-game margin: **+8** · Expected: **75.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Stanford Cardinal -1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **144** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Eastern Washington Eagles vs Corban Warriors (Final: **102–68**)

- 🟢 **WON — First-half moneyline: Eastern Washington Eagles (home)**
  - Actual First-half score: **58–24** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Eastern Washington Eagles (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Eastern Washington Eagles (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: Idaho Vandals vs Cal Poly Mustangs (Final: **74–83**)

- 🟢 **WON — Full-game total 167.5: Under**
  - Actual Full-game total: **157** · Expected: **79.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Idaho Vandals (home)**
  - Expected: **60.6%** · Price: **1.833 · DraftKings** · Counterfactual fade: **WON @ 2.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Idaho Vandals (home)**
  - Actual First-half score: **32–38** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Idaho Vandals (home)**
  - Actual Full-game margin: **-9** · Expected: **74.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Cal Poly Mustangs +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Idaho Vandals (home)**
  - Expected: **75.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **WON @ 2.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 168.5: Under**
  - Actual Full-game total: **157** · Expected: **80.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Arizona State Sun Devils vs UL Monroe Warhawks (Final: **92–62**)

- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **97.8%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +30.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+30** · Expected: **71.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Arizona State Sun Devils -30.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +32.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+30** · Expected: **74.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Arizona State Sun Devils -32.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **77.0%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **154** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **154** · Expected: **63.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arizona State Sun Devils (home)**
  - Actual First-half score: **47–22** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +29.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+30** · Expected: **68.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Arizona State Sun Devils -29.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **154** · Expected: **67.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: SMU Mustangs vs Ball State Cardinals (Final: **90–65**)

- 🟢 **WON — First-half moneyline: SMU Mustangs (home)**
  - Actual First-half score: **51–17** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **75.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **155** · Expected: **52.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +23.5: Ball State Cardinals (away)**
  - Actual Full-game margin: **+25** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 SMU Mustangs -23.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Under**
  - Actual Full-game total: **155** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **93.4%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **155** · Expected: **53.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +22.5: Ball State Cardinals (away)**
  - Actual Full-game margin: **+25** · Expected: **63.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 SMU Mustangs -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Auburn Tigers vs Arizona Wildcats (Final: **69–80**)

- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **149** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **56.6%** · Price: **2.240 · DraftKings** · Counterfactual fade: **WON @ 1.676 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +2.5: Auburn Tigers (home)**
  - Actual Full-game margin: **-11** · Expected: **75.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 Arizona Wildcats -2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **149** · Expected: **70.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +3.5: Auburn Tigers (home)**
  - Actual Full-game margin: **-11** · Expected: **76.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **WON @ 1.980 Arizona Wildcats -3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Auburn Tigers (home)**
  - Actual First-half score: **32–36** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Auburn Tigers (home)**
  - Expected: **74.0%** · Price: **2.240 · DraftKings** · Counterfactual fade: **WON @ 1.676 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Northwestern Wildcats vs DePaul Blue Demons (Final: **47–67**)

- 🔴 **LOST — Full-game spread -1.5: Northwestern Wildcats (home)**
  - Actual Full-game margin: **-20** · Expected: **70.0%** · Price: **2.100 · DraftKings** · Counterfactual fade: **WON @ 1.769 DePaul Blue Demons +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Northwestern Wildcats (home)**
  - Actual First-half score: **23–27** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: Northwestern Wildcats (home)**
  - Actual Full-game margin: **-20** · Expected: **75.0%** · Price: **1.800 · DraftKings** · Counterfactual fade: **WON @ 2.050 DePaul Blue Demons -1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **114** · Expected: **73.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Northwestern Wildcats (home)**
  - Expected: **74.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **114** · Expected: **72.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **114** · Expected: **68.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Northwestern Wildcats (home)**
  - Expected: **67.6%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Mississippi State Bulldogs vs Xavier Musketeers (Final: **63–81**)

- 🔴 **LOST — Full-game moneyline: Mississippi State Bulldogs (home)**
  - Expected: **74.0%** · Price: **2.450 · DraftKings** · Counterfactual fade: **WON @ 1.571 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +3.5: Mississippi State Bulldogs (home)**
  - Actual Full-game margin: **-18** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Xavier Musketeers -3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Xavier Musketeers (away)**
  - Expected: **55.6%** · Price: **1.571 · DraftKings** · Counterfactual fade: **LOST @ 2.450 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **144** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Mississippi State Bulldogs (home)**
  - Actual First-half score: **31–41** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **144** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **144** · Expected: **77.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **144** · Expected: **78.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Mississippi State Bulldogs (home)**
  - Actual Full-game margin: **-18** · Expected: **79.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 Xavier Musketeers -4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Alabama Crimson Tide vs Seton Hall Pirates (Final: **82–76**)

- 🟢 **WON — First-half moneyline: Alabama Crimson Tide (home)**
  - Actual First-half score: **38–28** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama Crimson Tide (home)**
  - Expected: **75.0%** · Price: **1.036 · DraftKings** · Counterfactual fade: **LOST @ 14.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama Crimson Tide (home)**
  - Expected: **87.9%** · Price: **1.036 · DraftKings** · Counterfactual fade: **LOST @ 14.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **158** · Expected: **68.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +17.5: Seton Hall Pirates (away)**
  - Actual Full-game margin: **+6** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Alabama Crimson Tide -17.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +18.5: Seton Hall Pirates (away)**
  - Actual Full-game margin: **+6** · Expected: **61.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Alabama Crimson Tide -18.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Omaha Mavericks vs Lindenwood Lions (Final: **67–56**)

- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **123** · Expected: **57.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Omaha Mavericks (home)**
  - Actual First-half score: **30–31** · Expected: **76.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **123** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -4.5: Omaha Mavericks (home)**
  - Actual Full-game margin: **+11** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Lindenwood Lions +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Omaha Mavericks (home)**
  - Expected: **75.0%** · Price: **1.505 · DraftKings** · Counterfactual fade: **LOST @ 2.640 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Lindenwood Lions (away)**
  - Expected: **60.2%** · Price: **2.640 · DraftKings** · Counterfactual fade: **WON @ 1.505 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **123** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Texas State Bobcats vs Southern Miss Golden Eagles (Final: **85–65**)

- 🟢 **WON — Full-game moneyline: Texas State Bobcats (home)**
  - Expected: **74.0%** · Price: **1.167 · DraftKings** · Counterfactual fade: **LOST @ 5.400 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -10.5: Texas State Bobcats (home)**
  - Actual Full-game margin: **+20** · Expected: **55.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **LOST @ 2.000 Southern Miss Golden Eagles +10.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Texas State Bobcats (home)**
  - Actual First-half score: **44–37** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **150** · Expected: **68.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Texas State Bobcats (home)**
  - Actual Full-game margin: **+20** · Expected: **56.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **LOST @ 2.000 Southern Miss Golden Eagles +9.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **150** · Expected: **73.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas State Bobcats (home)**
  - Expected: **85.6%** · Price: **1.167 · DraftKings** · Counterfactual fade: **LOST @ 5.400 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **150** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -11.5: Texas State Bobcats (home)**
  - Actual Full-game margin: **+20** · Expected: **54.0%** · Price: **2.000 · DraftKings** · Counterfactual fade: **LOST @ 1.833 Southern Miss Golden Eagles +11.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Murray State Racers vs Texas Southern Tigers (Final: **99–78**)

- 🟢 **WON — Full-game spread +27.5: Texas Southern Tigers (away)**
  - Actual Full-game margin: **+21** · Expected: **71.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Murray State Racers -27.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 163.5: Under**
  - Actual Full-game total: **177** · Expected: **73.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **95.9%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **177** · Expected: **71.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Murray State Racers (home)**
  - Actual First-half score: **50–38** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +29.5: Texas Southern Tigers (away)**
  - Actual Full-game margin: **+21** · Expected: **74.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Murray State Racers -29.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **75.0%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: North Dakota State Bison vs Illinois State Redbirds (Final: **82–81**)

- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **163** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **163** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: North Dakota State Bison (home)**
  - Actual First-half score: **31–31** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +2.5: North Dakota State Bison (home)**
  - Actual Full-game margin: **+1** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Illinois State Redbirds -2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: North Dakota State Bison (home)**
  - Actual Full-game margin: **+1** · Expected: **79.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Illinois State Redbirds -3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **163** · Expected: **60.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Illinois State Redbirds (away)**
  - Expected: **71.3%** · Price: **1.617 · DraftKings** · Counterfactual fade: **WON @ 2.360 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: North Dakota State Bison (home)**
  - Expected: **75.0%** · Price: **2.360 · DraftKings** · Counterfactual fade: **LOST @ 1.617 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Samford Bulldogs vs Wright State Raiders (Final: **66–90**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **156** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Samford Bulldogs (home)**
  - Expected: **74.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **WON @ 1.800 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Samford Bulldogs (home)**
  - Actual First-half score: **34–40** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wright State Raiders (away)**
  - Expected: **59.6%** · Price: **1.800 · DraftKings** · Counterfactual fade: **LOST @ 2.050 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **156** · Expected: **65.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **156** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: Samford Bulldogs (home)**
  - Actual Full-game margin: **-24** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Wright State Raiders -1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Tarleton State Texans vs Alcorn State Braves (Final: **73–69**)

- 🟢 **WON — Full-game total 151.5: Under**
  - Actual Full-game total: **142** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Tarleton State Texans (home)**
  - Actual First-half score: **29–33** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -14.5: Tarleton State Texans (home)**
  - Actual Full-game margin: **+4** · Expected: **54.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Alcorn State Braves +14.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +16.5: Alcorn State Braves (away)**
  - Actual Full-game margin: **+4** · Expected: **54.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Tarleton State Texans -16.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **85.8%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **76.0%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Houston Christian Huskies vs College of Biblical Studies Ambassadors (Final: **108–73**)

- 🟢 **WON — First-half moneyline: Houston Christian Huskies (home)**
  - Actual First-half score: **52–33** · Expected: **76.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Houston Christian Huskies (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Houston Christian Huskies (home)**
  - Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: UCF Knights vs Georgia Southern Eagles (Final: **84–70**)

- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **91.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +22.5: Georgia Southern Eagles (away)**
  - Actual Full-game margin: **+14** · Expected: **64.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 UCF Knights -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UCF Knights (home)**
  - Actual First-half score: **36–27** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **154** · Expected: **73.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +23.5: Georgia Southern Eagles (away)**
  - Actual Full-game margin: **+14** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 UCF Knights -23.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **154** · Expected: **78.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **75.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: TCU Horned Frogs vs Milwaukee Panthers (Final: **76–64**)

- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **140** · Expected: **66.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **140** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +23.5: Milwaukee Panthers (away)**
  - Actual Full-game margin: **+12** · Expected: **67.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 TCU Horned Frogs -23.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: TCU Horned Frogs (home)**
  - Actual First-half score: **36–33** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **140** · Expected: **59.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **96.1%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +22.5: Milwaukee Panthers (away)**
  - Actual Full-game margin: **+12** · Expected: **65.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 TCU Horned Frogs -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **140** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **74.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Arkansas Razorbacks vs Central Arkansas Bears (Final: **79–66**)

- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **145** · Expected: **73.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **145** · Expected: **71.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arkansas Razorbacks (home)**
  - Expected: **94.6%** · Price: **1.011 · DraftKings** · Counterfactual fade: **LOST @ 26.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Arkansas Razorbacks (home)**
  - Expected: **70.0%** · Price: **1.011 · DraftKings** · Counterfactual fade: **LOST @ 26.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arkansas Razorbacks (home)**
  - Actual First-half score: **43–27** · Expected: **67.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +25.5: Central Arkansas Bears (away)**
  - Actual Full-game margin: **+13** · Expected: **74.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Arkansas Razorbacks -25.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **145** · Expected: **73.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +24.5: Central Arkansas Bears (away)**
  - Actual Full-game margin: **+13** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Arkansas Razorbacks -24.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +23.5: Central Arkansas Bears (away)**
  - Actual Full-game margin: **+13** · Expected: **69.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Arkansas Razorbacks -23.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Maryland Terrapins vs Virginia Cavaliers (Final: **83–87**)

- 🟢 **WON — First-half moneyline: Maryland Terrapins (home)**
  - Actual First-half score: **47–36** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (away)**
  - Expected: **59.5%** · Price: **1.400 · DraftKings** · Counterfactual fade: **LOST @ 3.050 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **170** · Expected: **52.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **170** · Expected: **51.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +6.5: Maryland Terrapins (home)**
  - Actual Full-game margin: **-4** · Expected: **79.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Virginia Cavaliers -6.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +5.5: Maryland Terrapins (home)**
  - Actual Full-game margin: **-4** · Expected: **77.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Virginia Cavaliers -5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Under**
  - Actual Full-game total: **170** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Maryland Terrapins (home)**
  - Expected: **70.0%** · Price: **3.050 · DraftKings** · Counterfactual fade: **WON @ 1.400 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Duke Blue Devils vs SIU Edwardsville Cougars (Final: **90–47**)

- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **137** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +42.5: SIU Edwardsville Cougars (away)**
  - Actual Full-game margin: **+43** · Expected: **87.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Duke Blue Devils -42.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Duke Blue Devils (home)**
  - Actual First-half score: **43–20** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +44.5: SIU Edwardsville Cougars (away)**
  - Actual Full-game margin: **+43** · Expected: **87.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Duke Blue Devils -44.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Oklahoma Sooners vs Fairleigh Dickinson Knights (Final: **74–58**)

- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **132** · Expected: **60.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Oklahoma Sooners (home)**
  - Actual First-half score: **40–28** · Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +31.5: Fairleigh Dickinson Knights (away)**
  - Actual Full-game margin: **+16** · Expected: **78.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Oklahoma Sooners -31.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oklahoma Sooners (home)**
  - Expected: **98.5%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **132** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Drake Bulldogs vs South Dakota Coyotes (Final: **81–72**)

- 🟢 **WON — First-half moneyline: Drake Bulldogs (home)**
  - Actual First-half score: **30–29** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **75.0%** · Price: **1.091 · DraftKings** · Counterfactual fade: **LOST @ 8.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **153** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +13.5: South Dakota Coyotes (away)**
  - Actual Full-game margin: **+9** · Expected: **51.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Drake Bulldogs -13.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Drake Bulldogs (home)**
  - Expected: **91.9%** · Price: **1.091 · DraftKings** · Counterfactual fade: **LOST @ 8.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **153** · Expected: **69.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Liberty Flames vs North Carolina Central Eagles (Final: **74–66**)

- 🟢 **WON — Full-game moneyline: Liberty Flames (home)**
  - Expected: **91.9%** · Price: **1.048 · DraftKings** · Counterfactual fade: **LOST @ 12.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +16.5: North Carolina Central Eagles (away)**
  - Actual Full-game margin: **+8** · Expected: **59.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Liberty Flames -16.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Liberty Flames (home)**
  - Actual First-half score: **40–34** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +18.5: North Carolina Central Eagles (away)**
  - Actual Full-game margin: **+8** · Expected: **63.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Liberty Flames -18.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 137.5: Over**
  - Actual Full-game total: **140** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 136.5: Over**
  - Actual Full-game total: **140** · Expected: **73.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +17.5: North Carolina Central Eagles (away)**
  - Actual Full-game margin: **+8** · Expected: **61.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Liberty Flames -17.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Liberty Flames (home)**
  - Expected: **75.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **LOST @ 12.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Middle Tennessee Blue Raiders vs Southeast Missouri State Redhawks (Final: **88–75**)

- 🟢 **WON — First-half moneyline: Middle Tennessee Blue Raiders (home)**
  - Actual First-half score: **50–32** · Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -5.5: Middle Tennessee Blue Raiders (home)**
  - Actual Full-game margin: **+13** · Expected: **66.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Southeast Missouri State Redhawks +5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **71.7%** · Price: **1.400 · DraftKings** · Counterfactual fade: **LOST @ 3.050 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **163** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **163** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **75.0%** · Price: **1.400 · DraftKings** · Counterfactual fade: **LOST @ 3.050 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: UConn Huskies vs Ohio State Buckeyes (Final: **84–68**)

- 🟢 **WON — Full-game spread -12.5: UConn Huskies (home)**
  - Actual Full-game margin: **+16** · Expected: **55.0%** · Price: **2.000 · DraftKings** · Counterfactual fade: **LOST @ 1.833 Ohio State Buckeyes +12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **152** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **152** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -11.5: UConn Huskies (home)**
  - Actual Full-game margin: **+16** · Expected: **59.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Ohio State Buckeyes +11.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UConn Huskies (home)**
  - Expected: **85.7%** · Price: **1.125 · DraftKings** · Counterfactual fade: **LOST @ 6.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: UConn Huskies (home)**
  - Actual First-half score: **42–40** · Expected: **76.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UConn Huskies (home)**
  - Expected: **75.0%** · Price: **1.125 · DraftKings** · Counterfactual fade: **LOST @ 6.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Eastern Kentucky Colonels vs Kentucky State Thorobreds (Final: **106–59**)

- 🟢 **WON — Full-game moneyline: Eastern Kentucky Colonels (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Eastern Kentucky Colonels (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Eastern Kentucky Colonels (home)**
  - Actual First-half score: **39–23** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: UT Martin Skyhawks vs Champion Christian Tigers (Final: **86–57**)

- 🟢 **WON — Full-game moneyline: UT Martin Skyhawks (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UT Martin Skyhawks (home)**
  - Actual First-half score: **36–39** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UT Martin Skyhawks (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-13: VMI Keydets vs Mary Baldwin Fighting Squirrels (Final: **71–53**)

- 🟢 **WON — Full-game moneyline: VMI Keydets (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: VMI Keydets (home)**
  - Actual First-half score: **38–18** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: VMI Keydets (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-13: Purdue Boilermakers vs Ohio Bobcats (Final: **80–62**)

- 🟢 **WON — Full-game spread +19.5: Ohio Bobcats (away)**
  - Actual Full-game margin: **+18** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Purdue Boilermakers -19.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **142** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **75.0%** · Price: **1.028 · DraftKings** · Counterfactual fade: **LOST @ 16.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **142** · Expected: **60.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +18.5: Ohio Bobcats (away)**
  - Actual Full-game margin: **+18** · Expected: **63.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Purdue Boilermakers -18.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Purdue Boilermakers (home)**
  - Actual First-half score: **37–28** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Purdue Boilermakers (home)**
  - Expected: **92.8%** · Price: **1.028 · DraftKings** · Counterfactual fade: **LOST @ 16.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Charlotte 49ers vs App State Mountaineers (Final: **60–66**)

- 🔴 **LOST — Full-game spread -8.5: Charlotte 49ers (home)**
  - Actual Full-game margin: **-6** · Expected: **60.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **WON @ 1.847 App State Mountaineers +8.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: Charlotte 49ers (home)**
  - Actual Full-game margin: **-6** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 App State Mountaineers +7.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Charlotte 49ers (home)**
  - Expected: **75.0%** · Price: **1.282 · DraftKings** · Counterfactual fade: **WON @ 3.800 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **126** · Expected: **80.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 129.5: Over**
  - Actual Full-game total: **126** · Expected: **82.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Charlotte 49ers (home)**
  - Actual First-half score: **25–28** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Charlotte 49ers (home)**
  - Expected: **66.4%** · Price: **1.282 · DraftKings** · Counterfactual fade: **WON @ 3.800 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 130.5: Over**
  - Actual Full-game total: **126** · Expected: **81.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Michigan State Spartans vs Northern Iowa Panthers (Final: **106–72**)

- 🟢 **WON — Full-game moneyline: Michigan State Spartans (home)**
  - Expected: **97.9%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **178** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Under**
  - Actual Full-game total: **178** · Expected: **56.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +28.5: Northern Iowa Panthers (away)**
  - Actual Full-game margin: **+34** · Expected: **74.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Michigan State Spartans -28.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Michigan State Spartans (home)**
  - Expected: **74.0%** · Price: **1.002 · DraftKings** · Counterfactual fade: **LOST @ 41.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +29.5: Northern Iowa Panthers (away)**
  - Actual Full-game margin: **+34** · Expected: **77.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Michigan State Spartans -29.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +27.5: Northern Iowa Panthers (away)**
  - Actual Full-game margin: **+34** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Michigan State Spartans -27.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Michigan State Spartans (home)**
  - Actual First-half score: **59–44** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: Lehigh Mountain Hawks vs Marist Red Foxes (Final: **65–60**)

- 🟢 **WON — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **74.0%** · Price: **1.741 · DraftKings** · Counterfactual fade: **LOST @ 2.140 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 132.5: Over**
  - Actual Full-game total: **125** · Expected: **82.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Lehigh Mountain Hawks (home)**
  - Actual First-half score: **30–23** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lehigh Mountain Hawks (home)**
  - Expected: **58.7%** · Price: **1.741 · DraftKings** · Counterfactual fade: **LOST @ 2.140 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -1.5: Lehigh Mountain Hawks (home)**
  - Actual Full-game margin: **+5** · Expected: **71.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Marist Red Foxes +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Bucknell Bison vs Columbia Lions (Final: **59–77**)

- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **136** · Expected: **60.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Bucknell Bison (home)**
  - Actual Full-game margin: **-18** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Columbia Lions -4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Bucknell Bison (home)**
  - Actual First-half score: **29–36** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +5.5: Bucknell Bison (home)**
  - Actual Full-game margin: **-18** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Columbia Lions -5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **136** · Expected: **59.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Columbia Lions (away)**
  - Expected: **73.9%** · Price: **1.408 · DraftKings** · Counterfactual fade: **LOST @ 3.000 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Bucknell Bison (home)**
  - Expected: **70.0%** · Price: **3.000 · DraftKings** · Counterfactual fade: **WON @ 1.408 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Florida Gulf Coast Eagles vs Florida International Panthers (Final: **80–89**)

- 🔴 **LOST — Full-game moneyline: Florida Gulf Coast Eagles (home)**
  - Expected: **74.0%** · Price: **1.769 · DraftKings** · Counterfactual fade: **WON @ 2.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Florida Gulf Coast Eagles (home)**
  - Actual First-half score: **41–50** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **169** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **169** · Expected: **72.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Florida Gulf Coast Eagles (home)**
  - Actual Full-game margin: **-9** · Expected: **71.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **WON @ 1.980 Florida International Panthers +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **169** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Florida Gulf Coast Eagles (home)**
  - Expected: **56.3%** · Price: **1.769 · DraftKings** · Counterfactual fade: **WON @ 2.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -3.5: Florida Gulf Coast Eagles (home)**
  - Actual Full-game margin: **-9** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Florida International Panthers +3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -2.5: Florida Gulf Coast Eagles (home)**
  - Actual Full-game margin: **-9** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Florida International Panthers +2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Vanderbilt Commodores vs Prairie View A&M Panthers (Final: **89–62**)

- 🟢 **WON — First-half moneyline: Vanderbilt Commodores (home)**
  - Actual First-half score: **46–17** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +39.5: Prairie View A&M Panthers (away)**
  - Actual Full-game margin: **+27** · Expected: **83.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Vanderbilt Commodores -39.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 152.5: Over**
  - Actual Full-game total: **151** · Expected: **53.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **151** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +40.5: Prairie View A&M Panthers (away)**
  - Actual Full-game margin: **+27** · Expected: **83.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Vanderbilt Commodores -40.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Kentucky Wildcats vs Northern Arizona Lumberjacks (Final: **78–77**)

- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **155** · Expected: **52.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **98.2%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +35.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+1** · Expected: **80.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Kentucky Wildcats -35.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +34.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+1** · Expected: **80.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Kentucky Wildcats -34.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Kentucky Wildcats (home)**
  - Actual First-half score: **45–37** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **75.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **LOST @ 51.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **98.2%** · Price: **1.001 · DraftKings** · Counterfactual fade: **LOST @ 51.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 153.5: Under**
  - Actual Full-game total: **155** · Expected: **56.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +30.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+1** · Expected: **76.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Kentucky Wildcats -30.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **155** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Hofstra Pride vs Iona Gaels (Final: **62–59**)

- 🔴 **LOST — Full-game total 134.5: Over**
  - Actual Full-game total: **121** · Expected: **78.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: Hofstra Pride (home)**
  - Actual Full-game margin: **+3** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Iona Gaels +7.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Hofstra Pride (home)**
  - Actual First-half score: **22–20** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hofstra Pride (home)**
  - Expected: **74.0%** · Price: **1.328 · DraftKings** · Counterfactual fade: **LOST @ 3.450 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hofstra Pride (home)**
  - Expected: **74.2%** · Price: **1.328 · DraftKings** · Counterfactual fade: **LOST @ 3.450 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 133.5: Over**
  - Actual Full-game total: **121** · Expected: **81.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Elon Phoenix vs Chattanooga Mocs (Final: **70–77**)

- 🔴 **LOST — Full-game spread -2.5: Elon Phoenix (home)**
  - Actual Full-game margin: **-7** · Expected: **66.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **WON @ 1.847 Chattanooga Mocs +2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Elon Phoenix (home)**
  - Expected: **70.0%** · Price: **1.769 · DraftKings** · Counterfactual fade: **WON @ 2.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Elon Phoenix (home)**
  - Actual First-half score: **32–41** · Expected: **67.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Under**
  - Actual Full-game total: **147** · Expected: **52.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Elon Phoenix (home)**
  - Expected: **64.1%** · Price: **1.769 · DraftKings** · Counterfactual fade: **WON @ 2.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Northeastern Huskies vs Yale Bulldogs (Final: **84–73**)

- 🟢 **WON — Full-game moneyline: Northeastern Huskies (home)**
  - Expected: **75.0%** · Price: **2.020 · DraftKings** · Counterfactual fade: **LOST @ 1.820 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Under**
  - Actual Full-game total: **157** · Expected: **62.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: Northeastern Huskies (home)**
  - Actual Full-game margin: **+11** · Expected: **77.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Yale Bulldogs -1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northeastern Huskies (home)**
  - Actual First-half score: **43–34** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Yale Bulldogs (away)**
  - Expected: **66.1%** · Price: **1.820 · DraftKings** · Counterfactual fade: **WON @ 2.020 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +2.5: Northeastern Huskies (home)**
  - Actual Full-game margin: **+11** · Expected: **79.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Yale Bulldogs -2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Virginia Tech Hokies vs Mount St. Mary's Mountaineers (Final: **82–62**)

- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **144** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +28.5: Mount St. Mary's Mountaineers (away)**
  - Actual Full-game margin: **+20** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Virginia Tech Hokies -28.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Tech Hokies (home)**
  - Expected: **75.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **LOST @ 36.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **144** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Virginia Tech Hokies (home)**
  - Actual First-half score: **42–30** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **144** · Expected: **62.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Tech Hokies (home)**
  - Expected: **97.7%** · Price: **1.005 · DraftKings** · Counterfactual fade: **LOST @ 36.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +27.5: Mount St. Mary's Mountaineers (away)**
  - Actual Full-game margin: **+20** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Virginia Tech Hokies -27.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: NC State Wolfpack vs Troy Trojans (Final: **81–62**)

- 🟢 **WON — First-half moneyline: NC State Wolfpack (home)**
  - Actual First-half score: **36–27** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **143** · Expected: **54.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **75.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **143** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **95.9%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +22.5: Troy Trojans (away)**
  - Actual Full-game margin: **+19** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 NC State Wolfpack -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Illinois Fighting Illini vs Sam Houston Bearkats (Final: **108–86**)

- 🔴 **LOST — Full-game total 173.5: Under**
  - Actual Full-game total: **194** · Expected: **86.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +36.5: Sam Houston Bearkats (away)**
  - Actual Full-game margin: **+22** · Expected: **83.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Illinois Fighting Illini -36.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Illinois Fighting Illini (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 172.5: Under**
  - Actual Full-game total: **194** · Expected: **85.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Illinois Fighting Illini (home)**
  - Expected: **98.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Illinois Fighting Illini (home)**
  - Actual First-half score: **45–40** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 171.5: Under**
  - Actual Full-game total: **194** · Expected: **84.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: George Washington Revolutionaries vs Harvard Crimson (Final: **85–57**)

- 🟢 **WON — First-half moneyline: George Washington Revolutionaries (home)**
  - Actual First-half score: **40–33** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 133.5: Over**
  - Actual Full-game total: **142** · Expected: **78.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **75.0%** · Price: **1.133 · DraftKings** · Counterfactual fade: **LOST @ 6.250 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -11.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+28** · Expected: **55.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **LOST @ 1.847 Harvard Crimson +11.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **84.3%** · Price: **1.133 · DraftKings** · Counterfactual fade: **LOST @ 6.250 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -12.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+28** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Harvard Crimson +12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -10.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+28** · Expected: **59.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Harvard Crimson +10.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Fordham Rams vs Long Island University Sharks (Final: **80–58**)

- 🟢 **WON — First-half moneyline: Fordham Rams (home)**
  - Actual First-half score: **36–25** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **138** · Expected: **60.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **138** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -12.5: Fordham Rams (home)**
  - Actual Full-game margin: **+22** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Long Island University Sharks +12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **75.0%** · Price: **1.100 · DraftKings** · Counterfactual fade: **LOST @ 7.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **87.9%** · Price: **1.100 · DraftKings** · Counterfactual fade: **LOST @ 7.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Tennessee Volunteers vs Tennessee State Tigers (Final: **83–78**)

- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **161** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +27.5: Tennessee State Tigers (away)**
  - Actual Full-game margin: **+5** · Expected: **76.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Tennessee Volunteers -27.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Volunteers (home)**
  - Expected: **70.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **LOST @ 36.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +29.5: Tennessee State Tigers (away)**
  - Actual Full-game margin: **+5** · Expected: **79.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Tennessee Volunteers -29.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **161** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tennessee Volunteers (home)**
  - Expected: **98.8%** · Price: **1.005 · DraftKings** · Counterfactual fade: **LOST @ 36.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Tennessee Volunteers (home)**
  - Actual First-half score: **42–37** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +28.5: Tennessee State Tigers (away)**
  - Actual Full-game margin: **+5** · Expected: **78.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 Tennessee Volunteers -28.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: North Carolina Tar Heels vs Georgia Bulldogs (Final: **70–78**)

- 🔴 **LOST — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **75.0%** · Price: **1.556 · DraftKings** · Counterfactual fade: **WON @ 2.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: North Carolina Tar Heels (home)**
  - Actual Full-game margin: **-8** · Expected: **68.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Georgia Bulldogs +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: North Carolina Tar Heels (home)**
  - Actual First-half score: **31–38** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **75.5%** · Price: **1.556 · DraftKings** · Counterfactual fade: **WON @ 2.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **148** · Expected: **78.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **148** · Expected: **78.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Florida Gators vs Merrimack Warriors (Final: **83–54**)

- 🟢 **WON — Full-game moneyline: Florida Gators (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Florida Gators (home)**
  - Actual First-half score: **43–26** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Gators (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +37.5: Merrimack Warriors (away)**
  - Actual Full-game margin: **+29** · Expected: **83.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Florida Gators -37.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **137** · Expected: **67.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +36.5: Merrimack Warriors (away)**
  - Actual Full-game margin: **+29** · Expected: **82.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Florida Gators -36.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 141.5: Over**
  - Actual Full-game total: **137** · Expected: **72.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Butler Bulldogs vs Pittsburgh Panthers (Final: **64–76**)

- 🟢 **WON — First-half moneyline: Butler Bulldogs (home)**
  - Actual First-half score: **42–34** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **140** · Expected: **72.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **140** · Expected: **70.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Butler Bulldogs (home)**
  - Expected: **74.0%** · Price: **1.617 · DraftKings** · Counterfactual fade: **WON @ 2.360 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Butler Bulldogs (home)**
  - Expected: **53.1%** · Price: **1.617 · DraftKings** · Counterfactual fade: **WON @ 2.360 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -3.5: Butler Bulldogs (home)**
  - Actual Full-game margin: **-12** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 Pittsburgh Panthers +3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Rutgers Scarlet Knights vs Bryant Bulldogs (Final: **72–51**)

- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **77.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +22.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+21** · Expected: **63.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Rutgers Scarlet Knights -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +21.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+21** · Expected: **63.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 Rutgers Scarlet Knights -21.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **96.3%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **123** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Rutgers Scarlet Knights (home)**
  - Actual First-half score: **35–31** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **123** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **123** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Clemson Tigers vs Georgetown Hoyas (Final: **78–85**)

- 🔴 **LOST — Full-game moneyline: Clemson Tigers (home)**
  - Expected: **81.4%** · Price: **1.505 · DraftKings** · Counterfactual fade: **WON @ 2.640 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **163** · Expected: **60.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **163** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Clemson Tigers (home)**
  - Actual First-half score: **32–47** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **163** · Expected: **59.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Clemson Tigers (home)**
  - Actual Full-game margin: **-7** · Expected: **66.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Georgetown Hoyas +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Clemson Tigers (home)**
  - Expected: **74.0%** · Price: **1.505 · DraftKings** · Counterfactual fade: **WON @ 2.640 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Miami Hurricanes vs North Alabama Lions (Final: **94–52**)

- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Miami Hurricanes (home)**
  - Actual First-half score: **54–26** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +34.5: North Alabama Lions (away)**
  - Actual Full-game margin: **+42** · Expected: **78.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 Miami Hurricanes -34.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Under**
  - Actual Full-game total: **146** · Expected: **52.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +32.5: North Alabama Lions (away)**
  - Actual Full-game margin: **+42** · Expected: **76.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Miami Hurricanes -32.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **146** · Expected: **56.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Miami Hurricanes (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-13: Florida State Seminoles vs Texas A&M-Corpus Christi Islanders (Final: **84–60**)

- 🟢 **WON — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **89.3%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Florida State Seminoles (home)**
  - Actual First-half score: **40–28** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +24.5: Texas A&M-Corpus Christi Islanders (away)**
  - Actual Full-game margin: **+24** · Expected: **72.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Florida State Seminoles -24.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **70.0%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +22.5: Texas A&M-Corpus Christi Islanders (away)**
  - Actual Full-game margin: **+24** · Expected: **67.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 Florida State Seminoles -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **144** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **144** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-13: Cleveland State Vikings vs West Florida Argonauts (Final: **69–67**)

- 🟢 **WON — Full-game moneyline: Cleveland State Vikings (home)**
  - Expected: **70.0%** · Price: **1.800 · DraftKings** · Counterfactual fade: **LOST @ 2.050 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: Cleveland State Vikings (home)**
  - Actual Full-game margin: **+2** · Expected: **65.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 West Florida Argonauts +3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **136** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Cleveland State Vikings (home)**
  - Actual First-half score: **31–36** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **136** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **136** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Cleveland State Vikings (home)**
  - Actual Full-game margin: **+2** · Expected: **63.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 West Florida Argonauts +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -2.5: Cleveland State Vikings (home)**
  - Actual Full-game margin: **+2** · Expected: **66.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 West Florida Argonauts +2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **136** · Expected: **57.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Over**
  - Actual Full-game total: **136** · Expected: **52.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -1.5: Cleveland State Vikings (home)**
  - Actual Full-game margin: **+2** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 West Florida Argonauts +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: West Florida Argonauts (away)**
  - Expected: **58.7%** · Price: **2.050 · DraftKings** · Counterfactual fade: **WON @ 1.800 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-13: Green Bay Phoenix vs Viterbo V-Hawks (Final: **66–46**)

- 🟢 **WON — Full-game moneyline: Green Bay Phoenix (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Green Bay Phoenix (home)**
  - Actual First-half score: **24–23** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Green Bay Phoenix (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-12: Cal State Northridge Matadors vs Grace (CA) Talons (Final: **108–69**)

- 🟢 **WON — Full-game moneyline: Cal State Northridge Matadors (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Cal State Northridge Matadors (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Cal State Northridge Matadors (home)**
  - Actual First-half score: **57–37** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: San José State Spartans vs UC Riverside Highlanders (Final: **75–80**)

- 🟢 **WON — First-half moneyline: San José State Spartans (home)**
  - Actual First-half score: **37–36** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -8.5: San José State Spartans (home)**
  - Actual Full-game margin: **-5** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **WON @ 1.926 UC Riverside Highlanders +8.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: San José State Spartans (home)**
  - Expected: **67.6%** · Price: **1.253 · DraftKings** · Counterfactual fade: **WON @ 4.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **155** · Expected: **54.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: San José State Spartans (home)**
  - Expected: **80.0%** · Price: **1.253 · DraftKings** · Counterfactual fade: **WON @ 4.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Under**
  - Actual Full-game total: **155** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: BYU Cougars vs Montana State Bobcats (Final: **86–71**)

- 🟢 **WON — Full-game spread +22.5: Montana State Bobcats (away)**
  - Actual Full-game margin: **+15** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 BYU Cougars -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +23.5: Montana State Bobcats (away)**
  - Actual Full-game margin: **+15** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **157** · Expected: **70.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: BYU Cougars (home)**
  - Actual First-half score: **43–29** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **157** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **93.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **157** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +21.5: Montana State Bobcats (away)**
  - Actual Full-game margin: **+15** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 BYU Cougars -21.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Cal State Fullerton Titans vs Lincoln (CA) Oaklanders (Final: **101–71**)

- 🟢 **WON — Full-game moneyline: Cal State Fullerton Titans (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Cal State Fullerton Titans (home)**
  - Expected: **85.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Cal State Fullerton Titans (home)**
  - Actual First-half score: **48–41** · Expected: **78.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: UTEP Miners vs New Mexico State Aggies (Final: **60–58**)

- 🔴 **LOST — First-half moneyline: UTEP Miners (home)**
  - Actual First-half score: **28–28** · Expected: **72.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **118** · Expected: **64.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **80.0%** · Price: **1.459 · DraftKings** · Counterfactual fade: **LOST @ 2.800 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 142.5: Over**
  - Actual Full-game total: **118** · Expected: **68.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: UTEP Miners (home)**
  - Actual Full-game margin: **+2** · Expected: **76.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 New Mexico State Aggies +4.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: UTEP Miners (home)**
  - Actual Full-game margin: **+2** · Expected: **76.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 New Mexico State Aggies +5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **118** · Expected: **61.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UTEP Miners (home)**
  - Expected: **66.9%** · Price: **1.459 · DraftKings** · Counterfactual fade: **LOST @ 2.800 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-12: Alabama A&M Bulldogs vs West Florida Argonauts (Final: **85–83**)

- 🔴 **LOST — Full-game spread -7.5: Alabama A&M Bulldogs (home)**
  - Actual Full-game margin: **+2** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Alabama A&M Bulldogs (home)**
  - Actual Full-game margin: **+2** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama A&M Bulldogs (home)**
  - Expected: **59.0%** · Price: **1.388 · DraftKings** · Counterfactual fade: **LOST @ 3.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **168** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **168** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama A&M Bulldogs (home)**
  - Expected: **74.0%** · Price: **1.388 · DraftKings** · Counterfactual fade: **LOST @ 3.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **168** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: Alabama A&M Bulldogs (home)**
  - Actual Full-game margin: **+2** · Expected: **62.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 West Florida Argonauts +5.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Alabama A&M Bulldogs (home)**
  - Actual First-half score: **44–32** · Expected: **66.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: Texas A&M Aggies vs Oklahoma State Cowboys (Final: **88–69**)

- 🟢 **WON — Full-game moneyline: Texas A&M Aggies (home)**
  - Expected: **57.4%** · Price: **1.263 · DraftKings** · Counterfactual fade: **LOST @ 4.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Texas A&M Aggies (home)**
  - Actual First-half score: **50–37** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 179.5: Under**
  - Actual Full-game total: **157** · Expected: **89.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -8.5: Texas A&M Aggies (home)**
  - Actual Full-game margin: **+19** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas A&M Aggies (home)**
  - Expected: **81.0%** · Price: **1.263 · DraftKings** · Counterfactual fade: **LOST @ 4.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -7.5: Texas A&M Aggies (home)**
  - Actual Full-game margin: **+19** · Expected: **73.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **LOST @ 1.980 Oklahoma State Cowboys +7.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 178.5: Under**
  - Actual Full-game total: **157** · Expected: **88.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: South Alabama Jaguars vs Loyola Marymount Lions (Final: **79–72**)

- 🟢 **WON — Full-game spread +1.5: South Alabama Jaguars (home)**
  - Actual Full-game margin: **+7** · Expected: **81.0%** · Price: **1.820 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Loyola Marymount Lions (away)**
  - Expected: **68.4%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **151** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **151** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: South Alabama Jaguars (home)**
  - Actual First-half score: **39–34** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Alabama Jaguars (home)**
  - Expected: **80.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -1.5: South Alabama Jaguars (home)**
  - Actual Full-game margin: **+7** · Expected: **80.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **LOST @ 1.800 Loyola Marymount Lions +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Minnesota Golden Gophers vs Southern Jaguars (Final: **91–61**)

- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **152** · Expected: **67.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **152** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **152** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Minnesota Golden Gophers (home)**
  - Actual First-half score: **46–26** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +25.5: Southern Jaguars (away)**
  - Actual Full-game margin: **+30** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Minnesota Golden Gophers -25.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Minnesota Golden Gophers (home)**
  - Expected: **96.5%** · Price: **1.010 · DraftKings** · Counterfactual fade: **LOST @ 31.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Minnesota Golden Gophers (home)**
  - Expected: **81.0%** · Price: **1.010 · DraftKings** · Counterfactual fade: **LOST @ 31.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Southern Illinois Salukis vs Richmond Spiders (Final: **80–70**)

- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **150** · Expected: **61.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **150** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Southern Illinois Salukis (home)**
  - Actual Full-game margin: **+10** · Expected: **81.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Richmond Spiders (away)**
  - Expected: **63.1%** · Price: **1.676 · DraftKings** · Counterfactual fade: **WON @ 2.240 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 150.5: Under**
  - Actual Full-game total: **150** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Southern Illinois Salukis (home)**
  - Actual First-half score: **43–32** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +2.5: Southern Illinois Salukis (home)**
  - Actual Full-game margin: **+10** · Expected: **80.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 Richmond Spiders -2.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Southern Illinois Salukis (home)**
  - Expected: **80.0%** · Price: **2.240 · DraftKings** · Counterfactual fade: **LOST @ 1.676 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Coastal Carolina Chanticleers vs Virginia Lynchburg Dragons (Final: **103–54**)

- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Coastal Carolina Chanticleers (home)**
  - Actual First-half score: **55–32** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Coastal Carolina Chanticleers (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: Purdue Fort Wayne Mastodons vs Western Michigan Broncos (Final: **59–81**)

- 🔴 **LOST — Full-game spread +3.5: Purdue Fort Wayne Mastodons (home)**
  - Actual Full-game margin: **-22** · Expected: **82.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **WON @ 1.980 Western Michigan Broncos -3.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Purdue Fort Wayne Mastodons (home)**
  - Actual First-half score: **28–35** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Under**
  - Actual Full-game total: **140** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **140** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Michigan Broncos (away)**
  - Expected: **50.9%** · Price: **1.676 · DraftKings** · Counterfactual fade: **LOST @ 2.240 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Purdue Fort Wayne Mastodons (home)**
  - Expected: **80.0%** · Price: **2.240 · DraftKings** · Counterfactual fade: **WON @ 1.676 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Kennesaw State Owls vs LaGrange Panthers (Final: **107–54**)

- 🟢 **WON — First-half moneyline: Kennesaw State Owls (home)**
  - Actual First-half score: **53–28** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-12: Sacred Heart Pioneers vs Central Connecticut Blue Devils (Final: **75–79**)

- 🔴 **LOST — Full-game moneyline: Sacred Heart Pioneers (home)**
  - Expected: **80.0%** · Price: **1.388 · DraftKings** · Counterfactual fade: **WON @ 3.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **154** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **154** · Expected: **61.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Sacred Heart Pioneers (home)**
  - Actual Full-game margin: **-4** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 Central Connecticut Blue Devils +6.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **154** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Sacred Heart Pioneers (home)**
  - Expected: **74.3%** · Price: **1.388 · DraftKings** · Counterfactual fade: **WON @ 3.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Sacred Heart Pioneers (home)**
  - Actual First-half score: **43–31** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Sacred Heart Pioneers (home)**
  - Actual Full-game margin: **-4** · Expected: **75.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Quinnipiac Bobcats vs Monmouth Hawks (Final: **70–74**)

- 🔴 **LOST — Full-game moneyline: Quinnipiac Bobcats (home)**
  - Expected: **80.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Quinnipiac Bobcats (home)**
  - Actual Full-game margin: **-4** · Expected: **80.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **WON @ 1.870 Monmouth Hawks +1.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Quinnipiac Bobcats (home)**
  - Actual First-half score: **37–36** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Quinnipiac Bobcats (home)**
  - Expected: **58.1%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **144** · Expected: **65.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Howard Bison vs Stetson Hatters (Final: **75–62**)

- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **137** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Under**
  - Actual Full-game total: **137** · Expected: **51.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Howard Bison (home)**
  - Expected: **80.0%** · Price: **1.182 · DraftKings** · Counterfactual fade: **LOST @ 5.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **137** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **WON @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Howard Bison (home)**
  - Expected: **80.8%** · Price: **1.182 · DraftKings** · Counterfactual fade: **LOST @ 5.100 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -10.5: Howard Bison (home)**
  - Actual Full-game margin: **+13** · Expected: **59.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Howard Bison (home)**
  - Actual Full-game margin: **+13** · Expected: **62.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Stetson Hatters +9.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Howard Bison (home)**
  - Actual First-half score: **40–37** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: West Virginia Mountaineers vs Mercer Bears (Final: **75–61**)

- 🟢 **WON — Full-game moneyline: West Virginia Mountaineers (home)**
  - Expected: **80.0%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: West Virginia Mountaineers (home)**
  - Expected: **97.2%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +26.5: Mercer Bears (away)**
  - Actual Full-game margin: **+14** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +24.5: Mercer Bears (away)**
  - Actual Full-game margin: **+14** · Expected: **63.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 West Virginia Mountaineers -24.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +27.5: Mercer Bears (away)**
  - Actual Full-game margin: **+14** · Expected: **65.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: West Virginia Mountaineers (home)**
  - Actual First-half score: **31–28** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **136** · Expected: **54.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Towson Tigers vs Loyola Maryland Greyhounds (Final: **90–60**)

- 🟢 **WON — Full-game moneyline: Towson Tigers (home)**
  - Expected: **80.0%** · Price: **1.091 · DraftKings** · Counterfactual fade: **LOST @ 8.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Towson Tigers (home)**
  - Actual First-half score: **40–24** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 133.5: Over**
  - Actual Full-game total: **150** · Expected: **82.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **LOST @ 1.926 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Towson Tigers (home)**
  - Expected: **87.8%** · Price: **1.091 · DraftKings** · Counterfactual fade: **LOST @ 8.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 134.5: Over**
  - Actual Full-game total: **150** · Expected: **82.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: High Point Panthers vs Southern Utah Thunderbirds (Final: **88–59**)

- 🔴 **LOST — Full-game spread +23.5: Southern Utah Thunderbirds (away)**
  - Actual Full-game margin: **+29** · Expected: **60.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 High Point Panthers -23.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: High Point Panthers (home)**
  - Expected: **90.9%** · Price: **1.010 · DraftKings** · Counterfactual fade: **LOST @ 31.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **147** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +22.5: Southern Utah Thunderbirds (away)**
  - Actual Full-game margin: **+29** · Expected: **59.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: High Point Panthers (home)**
  - Expected: **80.0%** · Price: **1.010 · DraftKings** · Counterfactual fade: **LOST @ 31.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: High Point Panthers (home)**
  - Actual First-half score: **38–31** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **147** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **147** · Expected: **67.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Western Illinois Leathernecks vs Cal State Bakersfield Roadrunners (Final: **83–91**)

- 🟢 **WON — Full-game spread +9.5: Western Illinois Leathernecks (home)**
  - Actual Full-game margin: **-8** · Expected: **92.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Cal State Bakersfield Roadrunners (away)**
  - Expected: **56.9%** · Price: **1.227 · DraftKings** · Counterfactual fade: **LOST @ 4.400 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **174** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **174** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Western Illinois Leathernecks (home)**
  - Actual First-half score: **38–32** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Western Illinois Leathernecks (home)**
  - Expected: **81.0%** · Price: **4.400 · DraftKings** · Counterfactual fade: **WON @ 1.227 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +8.5: Western Illinois Leathernecks (home)**
  - Actual Full-game margin: **-8** · Expected: **91.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Cal State Bakersfield Roadrunners -8.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **174** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Rhode Island Rams vs Maryland Eastern Shore Hawks (Final: **67–47**)

- 🟢 **WON — Full-game spread +25.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+20** · Expected: **60.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **114** · Expected: **64.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **81.0%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +22.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+20** · Expected: **58.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **LOST @ 1.847 Rhode Island Rams -22.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +24.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+20** · Expected: **59.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **114** · Expected: **75.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **WON @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Rhode Island Rams (home)**
  - Actual First-half score: **32–28** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **114** · Expected: **74.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **95.4%** · Price: **1.012 · DraftKings** · Counterfactual fade: **LOST @ 23.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-12: Brown Bears vs Canisius Golden Griffins (Final: **81–86**)

- 🔴 **LOST — Full-game moneyline: Brown Bears (home)**
  - Expected: **88.3%** · Price: **1.095 · DraftKings** · Counterfactual fade: **WON @ 7.750 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 138.5: Over**
  - Actual Full-game total: **167** · Expected: **78.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +12.5: Canisius Golden Griffins (away)**
  - Actual Full-game margin: **-5** · Expected: **52.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **LOST @ 1.847 Brown Bears -12.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 137.5: Over**
  - Actual Full-game total: **167** · Expected: **80.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Brown Bears (home)**
  - Actual First-half score: **30–44** · Expected: **66.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Brown Bears (home)**
  - Expected: **76.0%** · Price: **1.095 · DraftKings** · Counterfactual fade: **WON @ 7.750 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Bowling Green Falcons vs Niagara Purple Eagles (Final: **88–76**)

- 🔴 **LOST — Full-game spread -14.5: Bowling Green Falcons (home)**
  - Actual Full-game margin: **+12** · Expected: **54.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 Niagara Purple Eagles +14.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Bowling Green Falcons (home)**
  - Actual First-half score: **44–31** · Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **89.7%** · Price: **1.061 · DraftKings** · Counterfactual fade: **LOST @ 10.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **164** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **LOST @ 1.893 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bowling Green Falcons (home)**
  - Expected: **80.0%** · Price: **1.061 · DraftKings** · Counterfactual fade: **LOST @ 10.500 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Penn State Nittany Lions vs Mercyhurst Lakers (Final: **72–57**)

- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **129** · Expected: **54.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **129** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **129** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Penn State Nittany Lions (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Penn State Nittany Lions (home)**
  - Actual First-half score: **33–29** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +21.5: Mercyhurst Lakers (away)**
  - Actual Full-game margin: **+15** · Expected: **59.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +20.5: Mercyhurst Lakers (away)**
  - Actual Full-game margin: **+15** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Penn State Nittany Lions -20.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Penn State Nittany Lions (home)**
  - Expected: **95.6%** · Price: **1.020 · DraftKings** · Counterfactual fade: **LOST @ 19.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-12: Indiana Hoosiers vs Tennessee Tech Golden Eagles (Final: **90–66**)

- 🟢 **WON — Full-game moneyline: Indiana Hoosiers (home)**
  - Expected: **97.5%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Indiana Hoosiers (home)**
  - Actual First-half score: **41–37** · Expected: **69.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Indiana Hoosiers (home)**
  - Expected: **80.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **LOST @ 21.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 169.5: Under**
  - Actual Full-game total: **156** · Expected: **83.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 167.5: Under**
  - Actual Full-game total: **156** · Expected: **81.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +25.5: Tennessee Tech Golden Eagles (away)**
  - Actual Full-game margin: **+24** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Indiana Hoosiers -25.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Coppin State Eagles vs Bradley Braves (Final: **54–56**)

- 🟢 **WON — First-half moneyline: Coppin State Eagles (home)**
  - Actual First-half score: **34–32** · Expected: **78.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 134.5: Over**
  - Actual Full-game total: **110** · Expected: **76.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bradley Braves (away)**
  - Expected: **87.3%** · Price: **1.125 · DraftKings** · Counterfactual fade: **LOST @ 6.500 home · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Coppin State Eagles (home)**
  - Expected: **82.0%** · Price: **6.500 · DraftKings** · Counterfactual fade: **WON @ 1.125 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +10.5: Coppin State Eagles (home)**
  - Actual Full-game margin: **-2** · Expected: **94.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +12.5: Coppin State Eagles (home)**
  - Actual Full-game margin: **-2** · Expected: **94.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **LOST @ 1.952 Bradley Braves -12.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 135.5: Over**
  - Actual Full-game total: **110** · Expected: **75.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 132.5: Over**
  - Actual Full-game total: **110** · Expected: **82.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **WON @ 1.952 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: Boston University Terriers vs Johnson & Wales Griffins (Final: **100–74**)

- 🟢 **WON — First-half moneyline: Boston University Terriers (home)**
  - Actual First-half score: **51–40** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston University Terriers (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Boston University Terriers (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-12: UMBC Retrievers vs Morgan State Bears (Final: **85–70**)

- 🟢 **WON — First-half moneyline: UMBC Retrievers (home)**
  - Actual First-half score: **42–39** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -5.5: UMBC Retrievers (home)**
  - Actual Full-game margin: **+15** · Expected: **76.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UMBC Retrievers (home)**
  - Expected: **73.4%** · Price: **1.377 · DraftKings** · Counterfactual fade: **LOST @ 3.150 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -6.5: UMBC Retrievers (home)**
  - Actual Full-game margin: **+15** · Expected: **74.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **LOST @ 1.870 Morgan State Bears +6.5 (away) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **155** · Expected: **71.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **155** · Expected: **80.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UMBC Retrievers (home)**
  - Expected: **81.0%** · Price: **1.377 · DraftKings** · Counterfactual fade: **LOST @ 3.150 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 162.5: Under**
  - Actual Full-game total: **155** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **LOST @ 1.893 over · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-12: George Mason Patriots vs Kansas City Roos (Final: **76–68**)

- 🟢 **WON — Full-game moneyline: George Mason Patriots (home)**
  - Expected: **82.5%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **144** · Expected: **73.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Mason Patriots (home)**
  - Expected: **80.0%** · Price: **1.062 · DraftKings** · Counterfactual fade: **LOST @ 10.000 away · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +15.5: Kansas City Roos (away)**
  - Actual Full-game margin: **+8** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 George Mason Patriots -15.5 (home) · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **144** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **LOST @ 1.909 under · DraftKings**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: George Mason Patriots (home)**
  - Actual First-half score: **37–33** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: California Golden Bears vs USC Trojans (Final: **65–77**)

- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **142** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: USC Trojans (away)**
  - Expected: **74.3%** · Price: **1.278 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: California Golden Bears (home)**
  - Expected: **86.0%** · Price: **3.850 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 160.5: Under**
  - Actual Full-game total: **142** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +8.5: California Golden Bears (home)**
  - Actual Full-game margin: **-12** · Expected: **93.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: California Golden Bears (home)**
  - Actual First-half score: **30–39** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: UC Irvine Anteaters vs Lincoln (CA) Oaklanders (Final: **90–40**)

- 🟢 **WON — First-half moneyline: UC Irvine Anteaters (home)**
  - Actual First-half score: **49–21** · Expected: **88.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Irvine Anteaters (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: UC Irvine Anteaters (home)**
  - Expected: **89.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Oregon Ducks vs Mississippi Valley State Delta Devils (Final: **100–66**)

- 🟢 **WON — First-half moneyline: Oregon Ducks (home)**
  - Actual First-half score: **50–28** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oregon Ducks (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +36.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+34** · Expected: **74.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **166** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +35.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+34** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **166** · Expected: **58.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +37.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+34** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **166** · Expected: **53.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +33.5: Mississippi Valley State Delta Devils (away)**
  - Actual Full-game margin: **+34** · Expected: **70.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Oregon Ducks (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: California Baptist Lancers vs Chicago State Cougars (Final: **62–81**)

- 🔴 **LOST — Full-game spread -11.5: California Baptist Lancers (home)**
  - Actual Full-game margin: **-19** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -10.5: California Baptist Lancers (home)**
  - Actual Full-game margin: **-19** · Expected: **71.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **143** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **143** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: California Baptist Lancers (home)**
  - Actual First-half score: **26–39** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **143** · Expected: **71.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: California Baptist Lancers (home)**
  - Expected: **80.8%** · Price: **1.118 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: California Baptist Lancers (home)**
  - Expected: **87.0%** · Price: **1.118 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 137.5: Over**
  - Actual Full-game total: **143** · Expected: **81.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Long Beach State Beach vs Pacific Tigers (Final: **62–59**)

- 🔴 **LOST — First-half moneyline: Long Beach State Beach (home)**
  - Actual First-half score: **33–34** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Long Beach State Beach (home)**
  - Expected: **50.9%** · Price: **3.100 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **121** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +6.5: Long Beach State Beach (home)**
  - Actual Full-game margin: **+3** · Expected: **91.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Long Beach State Beach (home)**
  - Expected: **86.0%** · Price: **3.100 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **121** · Expected: **78.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +5.5: Long Beach State Beach (home)**
  - Actual Full-game margin: **+3** · Expected: **90.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **121** · Expected: **71.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Seattle U Redhawks vs UTSA Roadrunners (Final: **59–86**)

- 🔴 **LOST — First-half moneyline: Seattle U Redhawks (home)**
  - Actual First-half score: **22–41** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -16.5: Seattle U Redhawks (home)**
  - Actual Full-game margin: **-27** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -17.5: Seattle U Redhawks (home)**
  - Actual Full-game margin: **-27** · Expected: **57.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **145** · Expected: **56.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Seattle U Redhawks (home)**
  - Expected: **89.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Seattle U Redhawks (home)**
  - Expected: **87.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **145** · Expected: **54.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UCLA Bruins vs Arkansas-Pine Bluff Golden Lions (Final: **106–37**)

- 🟢 **WON — Full-game moneyline: UCLA Bruins (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UCLA Bruins (home)**
  - Actual First-half score: **51–9** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +46.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+69** · Expected: **88.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UCLA Bruins (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +45.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+69** · Expected: **85.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **143** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +44.5: Arkansas-Pine Bluff Golden Lions (away)**
  - Actual Full-game margin: **+69** · Expected: **82.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Boise State Broncos vs Oakland Golden Grizzlies (Final: **111–79**)

- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **190** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **190** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **93.1%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Boise State Broncos (home)**
  - Expected: **90.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -16.5: Boise State Broncos (home)**
  - Actual Full-game margin: **+32** · Expected: **70.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boise State Broncos (home)**
  - Actual First-half score: **52–42** · Expected: **89.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Nebraska Cornhuskers vs Lindenwood Lions (Final: **94–60**)

- 🟢 **WON — Full-game spread +34.5: Lindenwood Lions (away)**
  - Actual Full-game margin: **+34** · Expected: **77.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nebraska Cornhuskers (home)**
  - Expected: **85.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Nebraska Cornhuskers (home)**
  - Expected: **98.3%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **154** · Expected: **51.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Nebraska Cornhuskers (home)**
  - Actual First-half score: **61–32** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Colorado State Rams vs Portland State Vikings (Final: **87–71**)

- 🔴 **LOST — Full-game spread -16.5: Colorado State Rams (home)**
  - Actual Full-game margin: **+16** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Colorado State Rams (home)**
  - Actual First-half score: **42–26** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Colorado State Rams (home)**
  - Expected: **91.9%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **158** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -15.5: Colorado State Rams (home)**
  - Actual Full-game margin: **+16** · Expected: **64.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Colorado State Rams (home)**
  - Expected: **87.0%** · Price: **1.048 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UC Santa Barbara Gauchos vs Nevada Wolf Pack (Final: **76–63**)

- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **139** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: UC Santa Barbara Gauchos (home)**
  - Actual Full-game margin: **+13** · Expected: **89.0%** · Price: **1.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Santa Barbara Gauchos (home)**
  - Expected: **88.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UC Santa Barbara Gauchos (home)**
  - Actual First-half score: **40–24** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **139** · Expected: **68.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Nevada Wolf Pack (away)**
  - Expected: **67.7%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Grand Canyon Lopes vs Cal Poly Mustangs (Final: **99–71**)

- 🟢 **WON — Full-game total 182.5: Under**
  - Actual Full-game total: **170** · Expected: **88.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Grand Canyon Lopes (home)**
  - Expected: **87.0%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Grand Canyon Lopes (home)**
  - Actual First-half score: **37–29** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -18.5: Grand Canyon Lopes (home)**
  - Actual Full-game margin: **+28** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Grand Canyon Lopes (home)**
  - Expected: **95.6%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 180.5: Under**
  - Actual Full-game total: **170** · Expected: **87.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 181.5: Under**
  - Actual Full-game total: **170** · Expected: **87.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: South Dakota State Jackrabbits vs Valparaiso Beacons (Final: **81–69**)

- 🟢 **WON — First-half moneyline: South Dakota State Jackrabbits (home)**
  - Actual First-half score: **49–32** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: South Dakota State Jackrabbits (home)**
  - Actual Full-game margin: **+12** · Expected: **88.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **150** · Expected: **68.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Dakota State Jackrabbits (home)**
  - Expected: **87.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Valparaiso Beacons (away)**
  - Expected: **51.3%** · Price: **1.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Arkansas Razorbacks vs Robert Morris Colonials (Final: **76–53**)

- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **129** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +31.5: Robert Morris Colonials (away)**
  - Actual Full-game margin: **+23** · Expected: **69.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arkansas Razorbacks (home)**
  - Actual First-half score: **39–30** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **129** · Expected: **62.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **129** · Expected: **62.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arkansas Razorbacks (home)**
  - Expected: **86.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +30.5: Robert Morris Colonials (away)**
  - Actual Full-game margin: **+23** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arkansas Razorbacks (home)**
  - Expected: **96.8%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Eastern Illinois Panthers vs Blackburn Beavers (Final: **102–44**)

- 🟢 **WON — Full-game moneyline: Eastern Illinois Panthers (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Eastern Illinois Panthers (home)**
  - Actual First-half score: **53–24** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Eastern Illinois Panthers (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Tulsa Golden Hurricane vs Oral Roberts Golden Eagles (Final: **80–71**)

- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **151** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulsa Golden Hurricane (home)**
  - Expected: **85.0%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **151** · Expected: **62.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Tulsa Golden Hurricane (home)**
  - Actual First-half score: **37–39** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulsa Golden Hurricane (home)**
  - Expected: **88.2%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Houston Cougars vs Jacksonville State Gamecocks (Final: **92–46**)

- 🟢 **WON — Full-game moneyline: Houston Cougars (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 135.5: Over**
  - Actual Full-game total: **138** · Expected: **84.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Houston Cougars (home)**
  - Actual First-half score: **45–23** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Houston Cougars (home)**
  - Expected: **98.2%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +36.5: Jacksonville State Gamecocks (away)**
  - Actual Full-game margin: **+46** · Expected: **74.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +35.5: Jacksonville State Gamecocks (away)**
  - Actual Full-game margin: **+46** · Expected: **71.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +34.5: Jacksonville State Gamecocks (away)**
  - Actual Full-game margin: **+46** · Expected: **71.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Southern Miss Golden Eagles vs Jackson State Tigers (Final: **84–80**)

- 🟢 **WON — Full-game moneyline: Southern Miss Golden Eagles (home)**
  - Expected: **84.7%** · Price: **1.220 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Southern Miss Golden Eagles (home)**
  - Actual First-half score: **32–40** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **164** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **164** · Expected: **51.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Southern Miss Golden Eagles (home)**
  - Expected: **88.0%** · Price: **1.220 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -9.5: Southern Miss Golden Eagles (home)**
  - Actual Full-game margin: **+4** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -8.5: Southern Miss Golden Eagles (home)**
  - Actual Full-game margin: **+4** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: St. Thomas Tommies vs Macalester Scots (Final: **119–82**)

- 🟢 **WON — First-half moneyline: St. Thomas Tommies (home)**
  - Actual First-half score: **54–42** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: St. Thomas Tommies (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Abilene Christian Wildcats vs Stephen F. Austin Lumberjacks (Final: **71–76**)

- 🔴 **LOST — Full-game moneyline: Abilene Christian Wildcats (home)**
  - Expected: **87.0%** · Price: **3.150 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **147** · Expected: **68.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **147** · Expected: **66.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Abilene Christian Wildcats (home)**
  - Actual Full-game margin: **-5** · Expected: **90.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Stephen F. Austin Lumberjacks (away)**
  - Expected: **72.2%** · Price: **1.377 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +6.5: Abilene Christian Wildcats (home)**
  - Actual Full-game margin: **-5** · Expected: **92.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Abilene Christian Wildcats (home)**
  - Actual First-half score: **29–35** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: North Texas Mean Green vs Santa Clara Broncos (Final: **71–76**)

- 🟢 **WON — First-half moneyline: North Texas Mean Green (home)**
  - Actual First-half score: **41–34** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +2.5: North Texas Mean Green (home)**
  - Actual Full-game margin: **-5** · Expected: **88.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: North Texas Mean Green (home)**
  - Actual Full-game margin: **-5** · Expected: **87.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: North Texas Mean Green (home)**
  - Expected: **87.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **147** · Expected: **54.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **147** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: North Texas Mean Green (home)**
  - Expected: **63.3%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Gardner-Webb Runnin' Bulldogs vs Erskine Flying Fleet (Final: **88–66**)

- 🟢 **WON — First-half moneyline: Gardner-Webb Runnin' Bulldogs (home)**
  - Actual First-half score: **33–29** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Gardner-Webb Runnin' Bulldogs (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Gardner-Webb Runnin' Bulldogs (home)**
  - Expected: **99.1%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Western Kentucky Hilltoppers vs Campbellsville Tigers (Final: **78–51**)

- 🟢 **WON — First-half moneyline: Western Kentucky Hilltoppers (home)**
  - Actual First-half score: **35–22** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Western Kentucky Hilltoppers (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Western Kentucky Hilltoppers (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Sam Houston Bearkats vs Pepperdine Waves (Final: **85–77**)

- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **162** · Expected: **64.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **162** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Sam Houston Bearkats (home)**
  - Actual First-half score: **41–35** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Sam Houston Bearkats (home)**
  - Expected: **86.0%** · Price: **2.360 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Sam Houston Bearkats (home)**
  - Actual Full-game margin: **+8** · Expected: **88.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Sam Houston Bearkats (home)**
  - Expected: **62.8%** · Price: **2.360 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +2.5: Sam Houston Bearkats (home)**
  - Actual Full-game margin: **+8** · Expected: **87.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Wichita State Shockers vs Friends Falcons (Final: **113–60**)

- 🟢 **WON — Full-game moneyline: Wichita State Shockers (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wichita State Shockers (home)**
  - Expected: **100.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Wichita State Shockers (home)**
  - Actual First-half score: **61–28** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Misericordia Cougars vs Binghamton Bearcats (Final: **62–87**)

- 🟢 **WON — Full-game moneyline: Binghamton Bearcats (away)**
  - Expected: **98.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Misericordia Cougars (home)**
  - Actual First-half score: **34–36** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Misericordia Cougars (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Indiana State Sycamores vs Milwaukee Panthers (Final: **73–82**)

- 🔴 **LOST — Full-game total 154.5: Under**
  - Actual Full-game total: **155** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Indiana State Sycamores (home)**
  - Expected: **72.4%** · Price: **1.541 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Indiana State Sycamores (home)**
  - Actual First-half score: **37–44** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Indiana State Sycamores (home)**
  - Actual Full-game margin: **-9** · Expected: **75.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Indiana State Sycamores (home)**
  - Expected: **85.0%** · Price: **1.541 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 153.5: Under**
  - Actual Full-game total: **155** · Expected: **52.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UNC Wilmington Seahawks vs Kent State Golden Flashes (Final: **91–73**)

- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **164** · Expected: **62.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNC Wilmington Seahawks (home)**
  - Expected: **87.0%** · Price: **1.294 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNC Wilmington Seahawks (home)**
  - Expected: **61.2%** · Price: **1.294 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -7.5: UNC Wilmington Seahawks (home)**
  - Actual Full-game margin: **+18** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **164** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: UNC Wilmington Seahawks (home)**
  - Actual First-half score: **49–30** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -8.5: UNC Wilmington Seahawks (home)**
  - Actual Full-game margin: **+18** · Expected: **73.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **164** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UNC Greensboro Spartans vs Lees McRae Bobcats (Final: **91–84**)

- 🟢 **WON — Full-game moneyline: UNC Greensboro Spartans (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UNC Greensboro Spartans (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: UNC Greensboro Spartans (home)**
  - Actual First-half score: **46–48** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Wagner Seahawks vs Bryant & Stratton-Albany Bobcats (Final: **117–51**)

- 🟢 **WON — First-half moneyline: Wagner Seahawks (home)**
  - Actual First-half score: **66–18** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wagner Seahawks (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wagner Seahawks (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Norfolk State Spartans vs Old Dominion Monarchs (Final: **60–71**)

- 🟢 **WON — Full-game moneyline: Old Dominion Monarchs (away)**
  - Expected: **72.4%** · Price: **1.370 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Norfolk State Spartans (home)**
  - Actual First-half score: **24–36** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Norfolk State Spartans (home)**
  - Actual Full-game margin: **-11** · Expected: **88.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **131** · Expected: **84.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Norfolk State Spartans (home)**
  - Expected: **85.0%** · Price: **3.200 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 135.5: Over**
  - Actual Full-game total: **131** · Expected: **85.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +6.5: Norfolk State Spartans (home)**
  - Actual Full-game margin: **-11** · Expected: **89.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **131** · Expected: **78.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Notre Dame Fighting Irish vs South Carolina State Bulldogs (Final: **99–59**)

- 🔴 **LOST — Full-game spread +25.5: South Carolina State Bulldogs (away)**
  - Actual Full-game margin: **+40** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Notre Dame Fighting Irish (home)**
  - Actual First-half score: **50–28** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +26.5: South Carolina State Bulldogs (away)**
  - Actual Full-game margin: **+40** · Expected: **61.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **158** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **87.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Notre Dame Fighting Irish (home)**
  - Expected: **95.6%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Alabama State Hornets vs Oklahoma Panhandle PANHANDLE ST (Final: **66–60**)

- 🟢 **WON — Full-game moneyline: Alabama State Hornets (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Alabama State Hornets (home)**
  - Expected: **98.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Alabama State Hornets (home)**
  - Actual First-half score: **32–28** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Lamar Cardinals vs Centenary (LA) Ladies (Final: **122–57**)

- 🟢 **WON — First-half moneyline: Lamar Cardinals (home)**
  - Actual First-half score: **65–26** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lamar Cardinals (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lamar Cardinals (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Rider Broncs vs William Paterson Pioneers (Final: **83–53**)

- 🟢 **WON — First-half moneyline: Rider Broncs (home)**
  - Actual First-half score: **46–24** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rider Broncs (home)**
  - Expected: **86.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rider Broncs (home)**
  - Expected: **98.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Princeton Tigers vs Stony Brook Seawolves (Final: **69–59**)

- 🟢 **WON — Full-game spread -4.5: Princeton Tigers (home)**
  - Actual Full-game margin: **+10** · Expected: **82.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Princeton Tigers (home)**
  - Actual First-half score: **39–21** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Princeton Tigers (home)**
  - Expected: **86.0%** · Price: **1.459 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **128** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -5.5: Princeton Tigers (home)**
  - Actual Full-game margin: **+10** · Expected: **77.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **128** · Expected: **81.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Princeton Tigers (home)**
  - Expected: **75.2%** · Price: **1.459 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: South Carolina Gamecocks vs Saint Peter's Peacocks (Final: **77–41**)

- 🔴 **LOST — Full-game spread +22.5: Saint Peter's Peacocks (away)**
  - Actual Full-game margin: **+36** · Expected: **54.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Carolina Gamecocks (home)**
  - Expected: **87.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -21.5: South Carolina Gamecocks (home)**
  - Actual Full-game margin: **+36** · Expected: **51.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: South Carolina Gamecocks (home)**
  - Actual First-half score: **37–17** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 126.5: Over**
  - Actual Full-game total: **118** · Expected: **90.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Carolina Gamecocks (home)**
  - Expected: **94.3%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Holy Cross Crusaders vs Furman Paladins (Final: **67–80**)

- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **147** · Expected: **54.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Furman Paladins (away)**
  - Expected: **88.7%** · Price: **1.118 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +11.5: Holy Cross Crusaders (home)**
  - Actual Full-game margin: **-13** · Expected: **97.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Holy Cross Crusaders (home)**
  - Expected: **86.0%** · Price: **6.750 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Holy Cross Crusaders (home)**
  - Actual First-half score: **28–42** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **147** · Expected: **56.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +12.5: Holy Cross Crusaders (home)**
  - Actual Full-game margin: **-13** · Expected: **97.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Bellarmine Knights vs Franklin Grizzlies (Final: **99–82**)

- 🟢 **WON — First-half moneyline: Bellarmine Knights (home)**
  - Actual First-half score: **55–25** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bellarmine Knights (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Bellarmine Knights (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Northern Illinois Huskies vs Clarke Pride (Final: **82–43**)

- 🟢 **WON — First-half moneyline: Northern Illinois Huskies (home)**
  - Actual First-half score: **36–24** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Illinois Huskies (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Illinois Huskies (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Delaware Blue Hens vs James Madison Dukes (Final: **61–71**)

- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **132** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Delaware Blue Hens (home)**
  - Expected: **87.0%** · Price: **2.400 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Delaware Blue Hens (home)**
  - Actual First-half score: **26–35** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Delaware Blue Hens (home)**
  - Expected: **57.1%** · Price: **2.400 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +3.5: Delaware Blue Hens (home)**
  - Actual Full-game margin: **-10** · Expected: **89.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: East Tennessee State Buccaneers vs Campbell Fighting Camels (Final: **68–53**)

- 🟢 **WON — Full-game spread -5.5: East Tennessee State Buccaneers (home)**
  - Actual Full-game margin: **+15** · Expected: **84.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **121** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: East Tennessee State Buccaneers (home)**
  - Expected: **87.0%** · Price: **1.444 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 145.5: Over**
  - Actual Full-game total: **121** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **121** · Expected: **66.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: East Tennessee State Buccaneers (home)**
  - Actual First-half score: **25–25** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -4.5: East Tennessee State Buccaneers (home)**
  - Actual Full-game margin: **+15** · Expected: **85.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: East Tennessee State Buccaneers (home)**
  - Expected: **85.0%** · Price: **1.444 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: St. John's Red Storm vs Colgate Raiders (Final: **90–59**)

- 🔴 **LOST — Full-game spread +29.5: Colgate Raiders (away)**
  - Actual Full-game margin: **+31** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +28.5: Colgate Raiders (away)**
  - Actual Full-game margin: **+31** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: St. John's Red Storm (home)**
  - Actual First-half score: **49–32** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **149** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **149** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **88.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. John's Red Storm (home)**
  - Expected: **95.3%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: VCU Rams vs NJIT Highlanders (Final: **87–49**)

- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **136** · Expected: **53.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: VCU Rams (home)**
  - Actual First-half score: **43–28** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +30.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+38** · Expected: **63.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +34.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+38** · Expected: **70.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: VCU Rams (home)**
  - Expected: **88.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: VCU Rams (home)**
  - Expected: **99.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: Florida Atlantic Owls vs William & Mary Tribe (Final: **95–90**)

- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **81.8%** · Price: **1.459 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Florida Atlantic Owls (home)**
  - Actual First-half score: **39–42** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Atlantic Owls (home)**
  - Expected: **86.0%** · Price: **1.459 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 165.5: Under**
  - Actual Full-game total: **185** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 166.5: Under**
  - Actual Full-game total: **185** · Expected: **75.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -4.5: Florida Atlantic Owls (home)**
  - Actual Full-game margin: **+5** · Expected: **77.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: West Georgia Wolves vs UNC Asheville Bulldogs (Final: **70–73**)

- 🔴 **LOST — Full-game spread -1.5: West Georgia Wolves (home)**
  - Actual Full-game margin: **-3** · Expected: **86.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: West Georgia Wolves (home)**
  - Expected: **87.0%** · Price: **1.769 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: West Georgia Wolves (home)**
  - Expected: **52.6%** · Price: **1.769 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: West Georgia Wolves (home)**
  - Actual First-half score: **32–37** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **143** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **143** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Drexel Dragons vs American University Eagles (Final: **64–66**)

- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **130** · Expected: **75.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Drexel Dragons (home)**
  - Expected: **53.2%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 141.5: Over**
  - Actual Full-game total: **130** · Expected: **69.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Drexel Dragons (home)**
  - Actual Full-game margin: **-2** · Expected: **86.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **130** · Expected: **79.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Drexel Dragons (home)**
  - Actual First-half score: **30–31** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Drexel Dragons (home)**
  - Expected: **87.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: St. Bonaventure Bonnies vs Cornell Big Red (Final: **88–64**)

- 🟢 **WON — First-half moneyline: St. Bonaventure Bonnies (home)**
  - Actual First-half score: **48–31** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **152** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Bonaventure Bonnies (home)**
  - Expected: **87.0%** · Price: **1.244 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: St. Bonaventure Bonnies (home)**
  - Expected: **67.3%** · Price: **1.244 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -8.5: St. Bonaventure Bonnies (home)**
  - Actual Full-game margin: **+24** · Expected: **73.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Massachusetts Minutemen vs UMass Lowell River Hawks (Final: **72–62**)

- 🟢 **WON — Full-game moneyline: Massachusetts Minutemen (home)**
  - Expected: **87.0%** · Price: **1.253 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -7.5: Massachusetts Minutemen (home)**
  - Actual Full-game margin: **+10** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **134** · Expected: **55.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Massachusetts Minutemen (home)**
  - Actual First-half score: **31–37** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -11.5: Massachusetts Minutemen (home)**
  - Actual Full-game margin: **+10** · Expected: **70.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Massachusetts Minutemen (home)**
  - Actual Full-game margin: **+10** · Expected: **71.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Massachusetts Minutemen (home)**
  - Expected: **70.3%** · Price: **1.253 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **134** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **134** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Michigan Wolverines vs Marquette Golden Eagles (Final: **84–70**)

- 🟢 **WON — First-half moneyline: Michigan Wolverines (home)**
  - Actual First-half score: **35–27** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Michigan Wolverines (home)**
  - Expected: **50.2%** · Price: **1.588 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Michigan Wolverines (home)**
  - Expected: **87.0%** · Price: **1.588 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **154** · Expected: **62.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -3.5: Michigan Wolverines (home)**
  - Actual Full-game margin: **+14** · Expected: **85.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **154** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Temple Owls vs Charleston Cougars (Final: **81–91**)

- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **172** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Temple Owls (home)**
  - Expected: **50.8%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -1.5: Temple Owls (home)**
  - Actual Full-game margin: **-10** · Expected: **85.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: Temple Owls (home)**
  - Actual Full-game margin: **-10** · Expected: **87.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **172** · Expected: **64.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Temple Owls (home)**
  - Expected: **86.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Temple Owls (home)**
  - Actual First-half score: **28–46** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: South Carolina Upstate Spartans vs Bob Jones Bruins (Final: **104–49**)

- 🟢 **WON — First-half moneyline: South Carolina Upstate Spartans (home)**
  - Actual First-half score: **62–28** · Expected: **78.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: South Carolina Upstate Spartans (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: South Carolina Upstate Spartans (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: IU Indianapolis Jaguars vs Southern Indiana Screaming Eagles (Final: **74–70**)

- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **144** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: IU Indianapolis Jaguars (home)**
  - Actual First-half score: **30–31** · Expected: **78.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: IU Indianapolis Jaguars (home)**
  - Expected: **87.0%** · Price: **1.282 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Southern Indiana Screaming Eagles (away)**
  - Expected: **50.6%** · Price: **3.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **144** · Expected: **70.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -8.5: IU Indianapolis Jaguars (home)**
  - Actual Full-game margin: **+4** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -7.5: IU Indianapolis Jaguars (home)**
  - Actual Full-game margin: **+4** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Eastern Michigan Eagles vs Great Lakes Christian Crusaders (Final: **108–44**)

- 🟢 **WON — Full-game moneyline: Eastern Michigan Eagles (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Eastern Michigan Eagles (home)**
  - Actual First-half score: **56–19** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Eastern Michigan Eagles (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Wofford Terriers vs Voorhees Tigers (Final: **91–47**)

- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Wofford Terriers (home)**
  - Actual First-half score: **40–16** · Expected: **84.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wofford Terriers (home)**
  - Expected: **88.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: Western Carolina Catamounts vs Virginia Lynchburg Dragons (Final: **136–68**)

- 🟢 **WON — Full-game moneyline: Western Carolina Catamounts (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Western Carolina Catamounts (home)**
  - Expected: **90.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Western Carolina Catamounts (home)**
  - Actual First-half score: **75–27** · Expected: **88.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: VMI Keydets vs North Carolina Central Eagles (Final: **65–51**)

- 🟢 **WON — Full-game spread -8.5: VMI Keydets (home)**
  - Actual Full-game margin: **+14** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: VMI Keydets (home)**
  - Actual First-half score: **32–22** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: VMI Keydets (home)**
  - Expected: **87.0%** · Price: **1.267 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -7.5: VMI Keydets (home)**
  - Actual Full-game margin: **+14** · Expected: **75.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: VMI Keydets (home)**
  - Expected: **60.0%** · Price: **1.267 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **116** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **116** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: New Hampshire Wildcats vs Thomas College Terriers (Final: **114–55**)

- 🟢 **WON — First-half moneyline: New Hampshire Wildcats (home)**
  - Actual First-half score: **51–26** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Hampshire Wildcats (home)**
  - Expected: **87.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Hampshire Wildcats (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-11: Navy Midshipmen vs Pennsylvania Quakers (Final: **81–76**)

- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **157** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Navy Midshipmen (home)**
  - Expected: **88.0%** · Price: **3.600 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +7.5: Navy Midshipmen (home)**
  - Actual Full-game margin: **+5** · Expected: **97.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Navy Midshipmen (home)**
  - Actual First-half score: **36–34** · Expected: **88.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +9.5: Navy Midshipmen (home)**
  - Actual Full-game margin: **+5** · Expected: **97.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +8.5: Navy Midshipmen (home)**
  - Actual Full-game margin: **+5** · Expected: **97.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **157** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Pennsylvania Quakers (away)**
  - Expected: **64.8%** · Price: **1.308 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: San Diego State Aztecs vs Bethune-Cookman Wildcats (Final: **88–75**)

- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **163** · Expected: **53.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Diego State Aztecs (home)**
  - Expected: **87.0%** · Price: **1.028 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Diego State Aztecs (home)**
  - Expected: **91.4%** · Price: **1.028 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: San Diego State Aztecs (home)**
  - Actual First-half score: **36–35** · Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -18.5: San Diego State Aztecs (home)**
  - Actual Full-game margin: **+13** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: East Carolina Pirates vs North Carolina A&T Aggies (Final: **72–67**)

- 🟢 **WON — First-half moneyline: East Carolina Pirates (home)**
  - Actual First-half score: **35–28** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -11.5: East Carolina Pirates (home)**
  - Actual Full-game margin: **+5** · Expected: **70.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **139** · Expected: **51.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: East Carolina Pirates (home)**
  - Expected: **77.7%** · Price: **1.148 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: East Carolina Pirates (home)**
  - Expected: **87.0%** · Price: **1.148 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -10.5: East Carolina Pirates (home)**
  - Actual Full-game margin: **+5** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Saint Joseph's Hawks vs South Florida Bulls (Final: **83–80**)

- 🟢 **WON — Full-game moneyline: Saint Joseph's Hawks (home)**
  - Expected: **87.0%** · Price: **1.336 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 152.5: Under**
  - Actual Full-game total: **163** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: South Florida Bulls (away)**
  - Expected: **68.6%** · Price: **3.400 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -6.5: Saint Joseph's Hawks (home)**
  - Actual Full-game margin: **+3** · Expected: **78.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Saint Joseph's Hawks (home)**
  - Actual First-half score: **39–34** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-11: South Florida Bulls vs Saint Joseph's Hawks (Final: **83–80**)

- 🟢 **WON — Full-game moneyline: South Florida Bulls (home)**
  - Expected: **68.6%** · Price: **1.308 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -7.5: South Florida Bulls (home)**
  - Actual Full-game margin: **+3** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **163** · Expected: **54.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Air Force Falcons vs Loyola Chicago Ramblers (Final: **107–93**)

- 🟢 **WON — Full-game moneyline: Air Force Falcons (home)**
  - Expected: **87.0%** · Price: **5.100 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **200** · Expected: **64.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **200** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Loyola Chicago Ramblers (away)**
  - Expected: **80.3%** · Price: **1.182 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Air Force Falcons (home)**
  - Actual First-half score: **50–35** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +10.5: Air Force Falcons (home)**
  - Actual Full-game margin: **+14** · Expected: **97.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: UC Davis Aggies vs Incarnate Word Cardinals (Final: **81–69**)

- 🟢 **WON — First-half moneyline: UC Davis Aggies (home)**
  - Actual First-half score: **49–33** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -11.5: UC Davis Aggies (home)**
  - Actual Full-game margin: **+12** · Expected: **70.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Davis Aggies (home)**
  - Expected: **87.0%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **150** · Expected: **62.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UC Davis Aggies (home)**
  - Expected: **80.3%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-11: UAlbany Great Danes vs Dartmouth Big Green (Final: **76–85**)

- 🔴 **LOST — Full-game spread -1.5: UAlbany Great Danes (home)**
  - Actual Full-game margin: **-9** · Expected: **86.0%** · Price: **2.050 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 138.5: Over**
  - Actual Full-game total: **161** · Expected: **75.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UAlbany Great Danes (home)**
  - Actual First-half score: **36–49** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +1.5: UAlbany Great Danes (home)**
  - Actual Full-game margin: **-9** · Expected: **87.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Dartmouth Big Green (away)**
  - Expected: **56.3%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **161** · Expected: **71.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UAlbany Great Danes (home)**
  - Expected: **87.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: Florida A&M Rattlers vs Kansas City Roos (Final: **65–87**)

- 🔴 **LOST — Full-game spread +8.5: Florida A&M Rattlers (home)**
  - Actual Full-game margin: **-22** · Expected: **94.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Florida A&M Rattlers (home)**
  - Actual First-half score: **21–51** · Expected: **78.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **152** · Expected: **63.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +6.5: Florida A&M Rattlers (home)**
  - Actual Full-game margin: **-22** · Expected: **92.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas City Roos (away)**
  - Expected: **84.5%** · Price: **1.290 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **152** · Expected: **66.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Florida A&M Rattlers (home)**
  - Expected: **86.0%** · Price: **3.750 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-11: East Texas A&M Lions vs Fairleigh Dickinson Knights (Final: **76–80**)

- 🔴 **LOST — Full-game moneyline: East Texas A&M Lions (home)**
  - Expected: **87.0%** · Price: **1.192 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -10.5: East Texas A&M Lions (home)**
  - Actual Full-game margin: **-4** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: East Texas A&M Lions (home)**
  - Expected: **87.5%** · Price: **1.192 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: East Texas A&M Lions (home)**
  - Actual First-half score: **37–31** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 137.5: Over**
  - Actual Full-game total: **156** · Expected: **81.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: San Francisco Dons vs Montana Grizzlies (Final: **63–58**)

- 🔴 **LOST — Full-game spread -13.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **54.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Francisco Dons (home)**
  - Expected: **69.0%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **121** · Expected: **56.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -11.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **58.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: San Francisco Dons (home)**
  - Actual First-half score: **26–28** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **121** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: San Francisco Dons (home)**
  - Expected: **77.0%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -12.5: San Francisco Dons (home)**
  - Actual Full-game margin: **+5** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Kentucky Wildcats vs Kansas Jayhawks (Final: **69–89**)

- 🔴 **LOST — First-half moneyline: Kentucky Wildcats (home)**
  - Actual First-half score: **36–46** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **158** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **158** · Expected: **64.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **158** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Kentucky Wildcats (home)**
  - Expected: **73.0%** · Price: **2.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: Kentucky Wildcats (home)**
  - Actual Full-game margin: **-20** · Expected: **80.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas Jayhawks (away)**
  - Expected: **64.1%** · Price: **1.459 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Fresno State Bulldogs vs Georgia State Panthers (Final: **95–92**)

- 🟢 **WON — Full-game total 146.5: Over**
  - Actual Full-game total: **187** · Expected: **61.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fresno State Bulldogs (home)**
  - Expected: **55.4%** · Price: **1.388 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **187** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **187** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Fresno State Bulldogs (home)**
  - Actual First-half score: **39–38** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -6.5: Fresno State Bulldogs (home)**
  - Actual Full-game margin: **+3** · Expected: **62.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: Fresno State Bulldogs (home)**
  - Actual Full-game margin: **+3** · Expected: **64.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fresno State Bulldogs (home)**
  - Expected: **70.0%** · Price: **1.388 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Utah State Aggies vs Utah Tech Trailblazers (Final: **118–86**)

- 🟢 **WON — First-half moneyline: Utah State Aggies (home)**
  - Actual First-half score: **53–35** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +25.5: Utah Tech Trailblazers (away)**
  - Actual Full-game margin: **+32** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Under**
  - Actual Full-game total: **204** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **204** · Expected: **63.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **204** · Expected: **57.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +27.5: Utah Tech Trailblazers (away)**
  - Actual Full-game margin: **+32** · Expected: **67.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **94.9%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **204** · Expected: **59.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah State Aggies (home)**
  - Expected: **80.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: New Mexico Lobos vs Little Rock Trojans (Final: **94–69**)

- 🟢 **WON — First-half moneyline: New Mexico Lobos (home)**
  - Actual First-half score: **47–34** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **163** · Expected: **63.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +23.5: Little Rock Trojans (away)**
  - Actual Full-game margin: **+25** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +22.5: Little Rock Trojans (away)**
  - Actual Full-game margin: **+25** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Mexico Lobos (home)**
  - Expected: **94.6%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **163** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Mexico Lobos (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **163** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: BYU Cougars vs Queens University Royals (Final: **98–82**)

- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **83.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 181.5: Under**
  - Actual Full-game total: **180** · Expected: **90.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +28.5: Queens University Royals (away)**
  - Actual Full-game margin: **+16** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: BYU Cougars (home)**
  - Expected: **97.6%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: BYU Cougars (home)**
  - Actual First-half score: **54–42** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +29.5: Queens University Royals (away)**
  - Actual Full-game margin: **+16** · Expected: **67.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 182.5: Under**
  - Actual Full-game total: **180** · Expected: **92.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arizona State Sun Devils vs Hofstra Pride (Final: **85–63**)

- 🔴 **LOST — Full-game spread +18.5: Hofstra Pride (away)**
  - Actual Full-game margin: **+22** · Expected: **51.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arizona State Sun Devils (home)**
  - Actual First-half score: **41–32** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **86.8%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 133.5: Over**
  - Actual Full-game total: **148** · Expected: **77.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona State Sun Devils (home)**
  - Expected: **80.0%** · Price: **1.031 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Utah Utes vs Weber State Wildcats (Final: **85–64**)

- 🟢 **WON — Full-game moneyline: Utah Utes (home)**
  - Expected: **77.0%** · Price: **1.028 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Utah Utes (home)**
  - Actual First-half score: **30–37** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 154.5: Over**
  - Actual Full-game total: **149** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +18.5: Weber State Wildcats (away)**
  - Actual Full-game margin: **+21** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Utah Utes (home)**
  - Expected: **93.8%** · Price: **1.028 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 155.5: Over**
  - Actual Full-game total: **149** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Virginia Cavaliers vs Bucknell Bison (Final: **95–61**)

- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **156** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Virginia Cavaliers (home)**
  - Actual First-half score: **53–31** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Virginia Cavaliers (home)**
  - Expected: **99.2%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +38.5: Bucknell Bison (away)**
  - Actual Full-game margin: **+34** · Expected: **76.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **156** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **156** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: TCU Horned Frogs vs UL Monroe Warhawks (Final: **100–92**)

- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +33.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **68.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +35.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 161.5: Under**
  - Actual Full-game total: **192** · Expected: **65.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **192** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: TCU Horned Frogs (home)**
  - Actual First-half score: **41–32** · Expected: **83.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +34.5: UL Monroe Warhawks (away)**
  - Actual Full-game margin: **+8** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: TCU Horned Frogs (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Colorado Buffaloes vs San Diego Toreros (Final: **83–75**)

- 🟢 **WON — Full-game moneyline: Colorado Buffaloes (home)**
  - Expected: **93.2%** · Price: **1.095 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **158** · Expected: **55.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Colorado Buffaloes (home)**
  - Actual First-half score: **34–37** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Colorado Buffaloes (home)**
  - Expected: **74.0%** · Price: **1.095 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **158** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -12.5: Colorado Buffaloes (home)**
  - Actual Full-game margin: **+8** · Expected: **53.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arizona Wildcats vs Northern Arizona Lumberjacks (Final: **90–59**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **149** · Expected: **61.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona Wildcats (home)**
  - Expected: **99.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +35.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+31** · Expected: **74.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +36.5: Northern Arizona Lumberjacks (away)**
  - Actual Full-game margin: **+31** · Expected: **74.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **149** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Arizona Wildcats (home)**
  - Actual First-half score: **46–22** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Arizona Wildcats (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Kansas State Wildcats vs Creighton Bluejays (Final: **89–82**)

- 🟢 **WON — Full-game spread +3.5: Kansas State Wildcats (home)**
  - Actual Full-game margin: **+7** · Expected: **81.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kansas State Wildcats (home)**
  - Expected: **79.0%** · Price: **2.400 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **171** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **171** · Expected: **62.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Kansas State Wildcats (home)**
  - Actual First-half score: **40–40** · Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Creighton Bluejays (away)**
  - Expected: **69.5%** · Price: **1.602 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +2.5: Kansas State Wildcats (home)**
  - Actual Full-game margin: **+7** · Expected: **81.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Wisconsin Badgers vs Fairfield Stags (Final: **89–42**)

- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **94.5%** · Price: **1.012 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Wisconsin Badgers (home)**
  - Expected: **77.0%** · Price: **1.012 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Wisconsin Badgers (home)**
  - Actual First-half score: **52–20** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +25.5: Fairfield Stags (away)**
  - Actual Full-game margin: **+47** · Expected: **67.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **131** · Expected: **51.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Ole Miss Rebels vs SE Louisiana Lions (Final: **91–51**)

- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **142** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +25.5: SE Louisiana Lions (away)**
  - Actual Full-game margin: **+40** · Expected: **66.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **142** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **142** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Ole Miss Rebels (home)**
  - Actual First-half score: **49–20** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **97.7%** · Price: **1.011 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +26.5: SE Louisiana Lions (away)**
  - Actual Full-game margin: **+40** · Expected: **67.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Ole Miss Rebels (home)**
  - Expected: **80.0%** · Price: **1.011 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Omaha Mavericks vs York (NE) Panthers (Final: **78–66**)

- 🔴 **LOST — First-half moneyline: Omaha Mavericks (home)**
  - Actual First-half score: **28–30** · Expected: **66.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Omaha Mavericks (home)**
  - Expected: **70.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Omaha Mavericks (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Tarleton State Texans vs UT Dallas Comets (Final: **76–57**)

- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Tarleton State Texans (home)**
  - Actual First-half score: **37–22** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tarleton State Texans (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Missouri State Bears vs Newman Jets (Final: **89–62**)

- 🟢 **WON — Full-game moneyline: Missouri State Bears (home)**
  - Expected: **99.6%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Missouri State Bears (home)**
  - Actual First-half score: **41–29** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Missouri State Bears (home)**
  - Expected: **79.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Northern Iowa Panthers vs Green Bay Phoenix (Final: **89–76**)

- 🟢 **WON — First-half moneyline: Northern Iowa Panthers (home)**
  - Actual First-half score: **46–30** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Iowa Panthers (home)**
  - Expected: **64.8%** · Price: **1.602 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **165** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **165** · Expected: **55.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northern Iowa Panthers (home)**
  - Expected: **74.0%** · Price: **1.602 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -3.5: Northern Iowa Panthers (home)**
  - Actual Full-game margin: **+13** · Expected: **69.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Arkansas State Red Wolves vs UT Martin Skyhawks (Final: **65–82**)

- 🔴 **LOST — Full-game total 153.5: Over**
  - Actual Full-game total: **147** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Arkansas State Red Wolves (home)**
  - Expected: **74.0%** · Price: **1.420 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Arkansas State Red Wolves (home)**
  - Actual First-half score: **34–54** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -4.5: Arkansas State Red Wolves (home)**
  - Actual Full-game margin: **-17** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -5.5: Arkansas State Red Wolves (home)**
  - Actual Full-game margin: **-17** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Arkansas State Red Wolves (home)**
  - Expected: **77.4%** · Price: **1.420 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Northern Colorado Bears vs Northern New Mexico Eagles (Final: **102–52**)

- 🟢 **WON — Full-game moneyline: Northern Colorado Bears (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Northern Colorado Bears (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northern Colorado Bears (home)**
  - Actual First-half score: **48–17** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Murray State Racers vs Alcorn State Braves (Final: **94–57**)

- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **151** · Expected: **72.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **96.2%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Murray State Racers (home)**
  - Expected: **80.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +30.5: Alcorn State Braves (away)**
  - Actual Full-game margin: **+37** · Expected: **69.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Murray State Racers (home)**
  - Actual First-half score: **51–30** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **151** · Expected: **70.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **151** · Expected: **72.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Texas A&M-Corpus Christi Islanders vs Dallas Christian Crusaders (Final: **102–58**)

- 🟢 **WON — Full-game moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Actual First-half score: **46–28** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas A&M-Corpus Christi Islanders (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: South Dakota Coyotes vs Houston Christian Huskies (Final: **76–87**)

- 🔴 **LOST — Full-game moneyline: South Dakota Coyotes (home)**
  - Expected: **57.9%** · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: South Dakota Coyotes (home)**
  - Expected: **74.0%** · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 156.5: Under**
  - Actual Full-game total: **163** · Expected: **57.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: South Dakota Coyotes (home)**
  - Actual Full-game margin: **-11** · Expected: **70.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: South Dakota Coyotes (home)**
  - Actual First-half score: **34–44** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **163** · Expected: **55.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Lipscomb Bisons vs Belmont Bruins (Final: **88–75**)

- 🔴 **LOST — Full-game total 157.5: Under**
  - Actual Full-game total: **163** · Expected: **62.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lipscomb Bisons (home)**
  - Expected: **55.3%** · Price: **2.420 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: Lipscomb Bisons (home)**
  - Actual First-half score: **40–32** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 155.5: Under**
  - Actual Full-game total: **163** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +3.5: Lipscomb Bisons (home)**
  - Actual Full-game margin: **+13** · Expected: **81.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +4.5: Lipscomb Bisons (home)**
  - Actual Full-game margin: **+13** · Expected: **81.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **163** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lipscomb Bisons (home)**
  - Expected: **77.0%** · Price: **2.420 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: SMU Mustangs vs Prairie View A&M Panthers (Final: **98–63**)

- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 161.5: Under**
  - Actual Full-game total: **161** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **97.6%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: SMU Mustangs (home)**
  - Expected: **77.0%** · Price: **1.005 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: SMU Mustangs (home)**
  - Actual First-half score: **52–36** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +26.5: Prairie View A&M Panthers (away)**
  - Actual Full-game margin: **+35** · Expected: **67.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Austin Peay Governors vs Livingstone Blue Bears (Final: **82–61**)

- 🟢 **WON — Full-game moneyline: Austin Peay Governors (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Austin Peay Governors (home)**
  - Actual First-half score: **46–31** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Austin Peay Governors (home)**
  - Expected: **99.9%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Texas Tech Red Raiders vs Illinois Fighting Illini (Final: **84–78**)

- 🟢 **WON — First-half moneyline: Texas Tech Red Raiders (home)**
  - Actual First-half score: **49–29** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Tech Red Raiders (home)**
  - Expected: **77.0%** · Price: **2.200 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 167.5: Under**
  - Actual Full-game total: **162** · Expected: **74.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Illinois Fighting Illini (away)**
  - Expected: **55.8%** · Price: **1.704 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **162** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +2.5: Texas Tech Red Raiders (home)**
  - Actual Full-game margin: **+6** · Expected: **80.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Saint Louis Billikens vs Northern Kentucky Norse (Final: **125–63**)

- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **94.5%** · Price: **1.015 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 169.5: Under**
  - Actual Full-game total: **188** · Expected: **75.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +23.5: Northern Kentucky Norse (away)**
  - Actual Full-game margin: **+62** · Expected: **64.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Saint Louis Billikens (home)**
  - Expected: **80.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Saint Louis Billikens (home)**
  - Actual First-half score: **56–24** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 168.5: Under**
  - Actual Full-game total: **188** · Expected: **74.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Iowa State Cyclones vs Southern Jaguars (Final: **84–47**)

- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **131** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +35.5: Southern Jaguars (away)**
  - Actual Full-game margin: **+37** · Expected: **70.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Iowa State Cyclones (home)**
  - Actual First-half score: **41–18** · Expected: **82.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Iowa State Cyclones (home)**
  - Expected: **81.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Iowa State Cyclones (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **131** · Expected: **57.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Baylor Bears vs Yale Bulldogs (Final: **90–61**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **151** · Expected: **59.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Baylor Bears (home)**
  - Actual First-half score: **41–30** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Baylor Bears (home)**
  - Expected: **78.0%** · Price: **1.025 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Baylor Bears (home)**
  - Expected: **91.3%** · Price: **1.025 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +19.5: Yale Bulldogs (away)**
  - Actual Full-game margin: **+29** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **151** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Middle Tennessee Blue Raiders vs Evansville Purple Aces (Final: **73–68**)

- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **72.0%** · Price: **1.227 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **141** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Middle Tennessee Blue Raiders (home)**
  - Expected: **82.8%** · Price: **1.227 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -9.5: Middle Tennessee Blue Raiders (home)**
  - Actual Full-game margin: **+5** · Expected: **57.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Middle Tennessee Blue Raiders (home)**
  - Actual First-half score: **27–20** · Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: UT Rio Grande Valley Vaqueros vs Texas State Bobcats (Final: **65–72**)

- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **137** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Actual First-half score: **28–30** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 138.5: Over**
  - Actual Full-game total: **137** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **137** · Expected: **67.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas State Bobcats (away)**
  - Expected: **72.3%** · Price: **1.435 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +5.5: UT Rio Grande Valley Vaqueros (home)**
  - Actual Full-game margin: **-7** · Expected: **84.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: UT Rio Grande Valley Vaqueros (home)**
  - Expected: **77.0%** · Price: **2.900 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +4.5: UT Rio Grande Valley Vaqueros (home)**
  - Actual Full-game margin: **-7** · Expected: **83.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Northwestern State Demons vs Millsaps Majors (Final: **77–49**)

- 🟢 **WON — Full-game moneyline: Northwestern State Demons (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Northwestern State Demons (home)**
  - Actual First-half score: **34–21** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Northwestern State Demons (home)**
  - Expected: **99.5%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Tulane Green Wave vs Nicholls Colonels (Final: **77–63**)

- 🟢 **WON — First-half moneyline: Tulane Green Wave (home)**
  - Actual First-half score: **35–24** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulane Green Wave (home)**
  - Expected: **77.0%** · Price: **1.227 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Tulane Green Wave (home)**
  - Expected: **67.9%** · Price: **1.227 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **140** · Expected: **52.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -9.5: Tulane Green Wave (home)**
  - Actual Full-game margin: **+14** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 153.5: Under**
  - Actual Full-game total: **140** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 152.5: Under**
  - Actual Full-game total: **140** · Expected: **51.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Central Arkansas Bears vs Ecclesia Royals (Final: **127–54**)

- 🟢 **WON — First-half moneyline: Central Arkansas Bears (home)**
  - Actual First-half score: **71–23** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Central Arkansas Bears (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Central Arkansas Bears (home)**
  - Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Vanderbilt Commodores vs Memphis Tigers (Final: **105–81**)

- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **85.6%** · Price: **1.062 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Vanderbilt Commodores (home)**
  - Expected: **77.0%** · Price: **1.062 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 160.5: Under**
  - Actual Full-game total: **186** · Expected: **65.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -13.5: Vanderbilt Commodores (home)**
  - Actual Full-game margin: **+24** · Expected: **53.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Vanderbilt Commodores (home)**
  - Actual First-half score: **47–42** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Texas Southern Tigers vs College of Biblical Studies Ambassadors (Final: **98–78**)

- 🟢 **WON — First-half moneyline: Texas Southern Tigers (home)**
  - Actual First-half score: **42–37** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Texas Southern Tigers (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Texas Southern Tigers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Toledo Rockets vs Wright State Raiders (Final: **68–76**)

- 🔴 **LOST — First-half moneyline: Toledo Rockets (home)**
  - Actual First-half score: **31–36** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **144** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **54.8%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread +1.5: Toledo Rockets (home)**
  - Actual Full-game margin: **-8** · Expected: **77.0%** · Price: **1.820 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Toledo Rockets (home)**
  - Expected: **74.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Hampton Pirates vs Apprentice School Builders (Final: **67–56**)

- 🟢 **WON — Full-game moneyline: Hampton Pirates (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Hampton Pirates (home)**
  - Actual First-half score: **33–25** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Hampton Pirates (home)**
  - Expected: **99.7%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: The Citadel Bulldogs vs Virginia Lynchburg Dragons (Final: **131–50**)

- 🟢 **WON — First-half moneyline: The Citadel Bulldogs (home)**
  - Actual First-half score: **66–20** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: The Citadel Bulldogs (home)**
  - Expected: **99.3%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: The Citadel Bulldogs (home)**
  - Expected: **80.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Duquesne Dukes vs Washington State Cougars (Final: **72–71**)

- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **80.0%** · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **143** · Expected: **58.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -3.5: Duquesne Dukes (home)**
  - Actual Full-game margin: **+1** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **143** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Duquesne Dukes (home)**
  - Actual First-half score: **30–39** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Duquesne Dukes (home)**
  - Expected: **59.7%** · Price: **1.617 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 144.5: Over**
  - Actual Full-game total: **143** · Expected: **61.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Davidson Wildcats vs Bradley Braves (Final: **62–45**)

- 🟢 **WON — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **81.8%** · Price: **1.160 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Davidson Wildcats (home)**
  - Expected: **80.0%** · Price: **1.160 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -10.5: Davidson Wildcats (home)**
  - Actual Full-game margin: **+17** · Expected: **64.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 139.5: Over**
  - Actual Full-game total: **107** · Expected: **68.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 137.5: Over**
  - Actual Full-game total: **107** · Expected: **70.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Davidson Wildcats (home)**
  - Actual First-half score: **30–20** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 136.5: Over**
  - Actual Full-game total: **107** · Expected: **72.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Kennesaw State Owls vs Winthrop Eagles (Final: **88–79**)

- 🟢 **WON — Full-game total 168.5: Under**
  - Actual Full-game total: **167** · Expected: **73.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **81.0%** · Price: **1.174 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 166.5: Under**
  - Actual Full-game total: **167** · Expected: **72.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -10.5: Kennesaw State Owls (home)**
  - Actual Full-game margin: **+9** · Expected: **60.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 170.5: Under**
  - Actual Full-game total: **167** · Expected: **76.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Kennesaw State Owls (home)**
  - Expected: **77.0%** · Price: **1.174 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Kennesaw State Owls (home)**
  - Actual First-half score: **32–33** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Lafayette Leopards vs Penn State Lehigh Valley Nittany Lions (Final: **89–46**)

- 🟢 **WON — First-half moneyline: Lafayette Leopards (home)**
  - Actual First-half score: **54–16** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Lafayette Leopards (home)**
  - Expected: **99.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`

### 2026-11-10: Longwood Lancers vs St. Mary's (MD) Seahawks (Final: **101–45**)

- 🟢 **WON — First-half moneyline: Longwood Lancers (home)**
  - Actual First-half score: **48–23** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Longwood Lancers (home)**
  - Expected: **99.8%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Longwood Lancers (home)**
  - Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Villanova Wildcats vs La Salle Explorers (Final: **80–59**)

- 🟢 **WON — Full-game moneyline: Villanova Wildcats (home)**
  - Expected: **95.7%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Villanova Wildcats (home)**
  - Expected: **80.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 159.5: Under**
  - Actual Full-game total: **139** · Expected: **63.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +20.5: La Salle Explorers (away)**
  - Actual Full-game margin: **+21** · Expected: **61.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Villanova Wildcats (home)**
  - Actual First-half score: **36–31** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +23.5: La Salle Explorers (away)**
  - Actual Full-game margin: **+21** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: UCF Knights vs South Alabama Jaguars (Final: **76–69**)

- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **78.0%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: UCF Knights (home)**
  - Actual First-half score: **32–34** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 150.5: Over**
  - Actual Full-game total: **145** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: UCF Knights (home)**
  - Expected: **89.8%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 151.5: Over**
  - Actual Full-game total: **145** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: UCF Knights (home)**
  - Actual Full-game margin: **+7** · Expected: **54.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgia Southern Eagles vs Mercer Bears (Final: **91–97**)

- 🔴 **LOST — Full-game total 165.5: Under**
  - Actual Full-game total: **188** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 164.5: Under**
  - Actual Full-game total: **188** · Expected: **68.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Georgia Southern Eagles (home)**
  - Expected: **59.9%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Georgia Southern Eagles (home)**
  - Expected: **77.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Georgia Southern Eagles (home)**
  - Actual First-half score: **46–39** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -1.5: Georgia Southern Eagles (home)**
  - Actual Full-game margin: **-6** · Expected: **73.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Iowa Hawkeyes vs Virginia Tech Hokies (Final: **88–93**)

- 🔴 **LOST — Full-game spread -4.5: Iowa Hawkeyes (home)**
  - Actual Full-game margin: **-5** · Expected: **73.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **181** · Expected: **67.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Iowa Hawkeyes (home)**
  - Expected: **80.0%** · Price: **1.505 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 141.5: Over**
  - Actual Full-game total: **181** · Expected: **66.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Iowa Hawkeyes (home)**
  - Actual First-half score: **40–41** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Iowa Hawkeyes (home)**
  - Expected: **67.3%** · Price: **1.505 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **181** · Expected: **64.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Northeastern Huskies vs Harvard Crimson (Final: **74–70**)

- 🟢 **WON — First-half moneyline: Northeastern Huskies (home)**
  - Actual First-half score: **29–19** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **144** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **144** · Expected: **56.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +1.5: Northeastern Huskies (home)**
  - Actual Full-game margin: **+4** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Harvard Crimson (away)**
  - Expected: **52.2%** · Price: **1.769 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Northeastern Huskies (home)**
  - Expected: **74.0%** · Price: **2.100 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **144** · Expected: **61.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgia Tech Yellow Jackets vs Eastern Kentucky Colonels (Final: **85–71**)

- 🟢 **WON — Full-game total 163.5: Under**
  - Actual Full-game total: **156** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgia Tech Yellow Jackets (home)**
  - Expected: **85.5%** · Price: **1.056 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game moneyline: Georgia Tech Yellow Jackets (home)**
  - Expected: **73.0%** · Price: **1.056 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **156** · Expected: **66.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 165.5: Under**
  - Actual Full-game total: **156** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +16.5: Eastern Kentucky Colonels (away)**
  - Actual Full-game margin: **+14** · Expected: **55.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Georgia Tech Yellow Jackets (home)**
  - Actual First-half score: **41–42** · Expected: **73.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +15.5: Eastern Kentucky Colonels (away)**
  - Actual Full-game margin: **+14** · Expected: **51.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Rhode Island Rams vs Maine Black Bears (Final: **71–65**)

- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **80.0%** · Price: **1.022 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +20.5: Maine Black Bears (away)**
  - Actual Full-game margin: **+6** · Expected: **61.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 140.5: Over**
  - Actual Full-game total: **136** · Expected: **66.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rhode Island Rams (home)**
  - Expected: **93.6%** · Price: **1.022 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — First-half moneyline: Rhode Island Rams (home)**
  - Actual First-half score: **23–28** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: NC State Wolfpack vs Presbyterian Blue Hose (Final: **94–60**)

- 🟢 **WON — First-half moneyline: NC State Wolfpack (home)**
  - Actual First-half score: **46–19** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 148.5: Over**
  - Actual Full-game total: **154** · Expected: **57.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +19.5: Presbyterian Blue Hose (away)**
  - Actual Full-game margin: **+34** · Expected: **61.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 150.5: Over**
  - Actual Full-game total: **154** · Expected: **55.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **77.0%** · Price: **1.015 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: NC State Wolfpack (home)**
  - Expected: **96.7%** · Price: **1.015 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Troy Trojans vs Loyola Marymount Lions (Final: **74–69**)

- 🟢 **WON — Full-game moneyline: Troy Trojans (home)**
  - Expected: **71.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Troy Trojans (home)**
  - Expected: **51.6%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +1.5: Troy Trojans (home)**
  - Actual Full-game margin: **+5** · Expected: **73.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Troy Trojans (home)**
  - Actual First-half score: **29–37** · Expected: **67.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 146.5: Over**
  - Actual Full-game total: **143** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 143.5: Over**
  - Actual Full-game total: **143** · Expected: **64.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: North Carolina Tar Heels vs Wyoming Cowboys (Final: **76–63**)

- 🟢 **WON — Full-game spread -12.5: North Carolina Tar Heels (home)**
  - Actual Full-game margin: **+13** · Expected: **53.0%** · Price: **1.833 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **73.0%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: North Carolina Tar Heels (home)**
  - Actual Full-game margin: **+13** · Expected: **51.0%** · Price: **2.000 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: North Carolina Tar Heels (home)**
  - Expected: **89.0%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **139** · Expected: **55.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: North Carolina Tar Heels (home)**
  - Actual First-half score: **31–29** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: New Haven Chargers vs Army Black Knights (Final: **63–55**)

- 🟢 **WON — Full-game spread +1.5: New Haven Chargers (home)**
  - Actual Full-game margin: **+8** · Expected: **77.0%** · Price: **1.847 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: New Haven Chargers (home)**
  - Expected: **73.0%** · Price: **2.000 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 133.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Army Black Knights (away)**
  - Expected: **56.2%** · Price: **1.833 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 132.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: New Haven Chargers (home)**
  - Actual First-half score: **24–19** · Expected: **74.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 131.5: Over**
  - Actual Full-game total: **118** · Expected: **77.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: George Washington Revolutionaries vs Elon Phoenix (Final: **69–50**)

- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **84.3%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — First-half moneyline: George Washington Revolutionaries (home)**
  - Actual First-half score: **23–22** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **119** · Expected: **55.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 148.5: Over**
  - Actual Full-game total: **119** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -13.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+19** · Expected: **53.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread -14.5: George Washington Revolutionaries (home)**
  - Actual Full-game margin: **+19** · Expected: **53.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: George Washington Revolutionaries (home)**
  - Expected: **77.0%** · Price: **1.083 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Fordham Rams vs Iona Gaels (Final: **78–72**)

- 🟢 **WON — Full-game total 140.5: Over**
  - Actual Full-game total: **150** · Expected: **69.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Fordham Rams (home)**
  - Actual First-half score: **43–37** · Expected: **81.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 139.5: Over**
  - Actual Full-game total: **150** · Expected: **71.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **80.0%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Fordham Rams (home)**
  - Expected: **79.7%** · Price: **1.125 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -11.5: Fordham Rams (home)**
  - Actual Full-game margin: **+6** · Expected: **64.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Dayton Flyers vs Siena Saints (Final: **89–99**)

- 🔴 **LOST — Full-game total 159.5: Under**
  - Actual Full-game total: **188** · Expected: **65.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread -13.5: Dayton Flyers (home)**
  - Actual Full-game margin: **-10** · Expected: **51.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 158.5: Under**
  - Actual Full-game total: **188** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Dayton Flyers (home)**
  - Actual First-half score: **45–46** · Expected: **68.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Dayton Flyers (home)**
  - Expected: **92.0%** · Price: **1.091 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game moneyline: Dayton Flyers (home)**
  - Expected: **74.0%** · Price: **1.091 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Central Michigan Chippewas vs Youngstown State Penguins (Final: **91–78**)

- 🟢 **WON — Full-game total 152.5: Over**
  - Actual Full-game total: **169** · Expected: **52.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Central Michigan Chippewas (home)**
  - Actual First-half score: **50–34** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 151.5: Over**
  - Actual Full-game total: **169** · Expected: **54.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Youngstown State Penguins (away)**
  - Expected: **52.1%** · Price: **2.400 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread -3.5: Central Michigan Chippewas (home)**
  - Actual Full-game margin: **+13** · Expected: **71.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Central Michigan Chippewas (home)**
  - Expected: **77.0%** · Price: **1.602 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Wake Forest Demon Deacons vs Charleston Southern Buccaneers (Final: **82–64**)

- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **146** · Expected: **63.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 154.5: Under**
  - Actual Full-game total: **146** · Expected: **53.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wake Forest Demon Deacons (home)**
  - Expected: **77.0%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 156.5: Under**
  - Actual Full-game total: **146** · Expected: **58.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Wake Forest Demon Deacons (home)**
  - Actual First-half score: **38–22** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Wake Forest Demon Deacons (home)**
  - Expected: **93.5%** · Price: **1.020 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game spread +18.5: Charleston Southern Buccaneers (away)**
  - Actual Full-game margin: **+18** · Expected: **58.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 155.5: Under**
  - Actual Full-game total: **146** · Expected: **56.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Duke Blue Devils vs Michigan State Spartans (Final: **73–79**)

- 🔴 **LOST — First-half moneyline: Duke Blue Devils (home)**
  - Actual First-half score: **25–34** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 147.5: Over**
  - Actual Full-game total: **152** · Expected: **59.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **77.0%** · Price: **1.294 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 145.5: Over**
  - Actual Full-game total: **152** · Expected: **61.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Duke Blue Devils (home)**
  - Expected: **73.4%** · Price: **1.294 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game spread -7.5: Duke Blue Devils (home)**
  - Actual Full-game margin: **-6** · Expected: **68.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Boston College Eagles vs Maryland Eastern Shore Hawks (Final: **82–54**)

- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **73.0%** · Price: **1.002 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 130.5: Over**
  - Actual Full-game total: **136** · Expected: **80.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 129.5: Over**
  - Actual Full-game total: **136** · Expected: **82.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Boston College Eagles (home)**
  - Expected: **96.9%** · Price: **1.002 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 131.5: Over**
  - Actual Full-game total: **136** · Expected: **77.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 128.5: Over**
  - Actual Full-game total: **136** · Expected: **84.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Boston College Eagles (home)**
  - Actual First-half score: **43–25** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +27.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+28** · Expected: **68.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +26.5: Maryland Eastern Shore Hawks (away)**
  - Actual Full-game margin: **+28** · Expected: **68.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Xavier Musketeers vs Detroit Mercy Titans (Final: **100–55**)

- 🔴 **LOST — Full-game spread +34.5: Detroit Mercy Titans (away)**
  - Actual Full-game margin: **+45** · Expected: **70.0%** · Price: **1.980 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Xavier Musketeers (home)**
  - Expected: **81.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Xavier Musketeers (home)**
  - Expected: **98.4%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (no_book_evidence)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `provider:espn`
- 🟢 **WON — Full-game total 158.5: Under**
  - Actual Full-game total: **155** · Expected: **65.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Xavier Musketeers (home)**
  - Actual First-half score: **43–20** · Expected: **77.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 157.5: Under**
  - Actual Full-game total: **155** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +35.5: Detroit Mercy Titans (away)**
  - Actual Full-game margin: **+45** · Expected: **73.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Providence Friars vs Bryant Bulldogs (Final: **78–55**)

- 🟢 **WON — Full-game moneyline: Providence Friars (home)**
  - Expected: **77.0%** · Price: **1.002 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +30.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+23** · Expected: **69.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game total 147.5: Over**
  - Actual Full-game total: **133** · Expected: **55.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +31.5: Bryant Bulldogs (away)**
  - Actual Full-game margin: **+23** · Expected: **73.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Providence Friars (home)**
  - Expected: **97.7%** · Price: **1.002 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🔴 **LOST — Full-game total 149.5: Over**
  - Actual Full-game total: **133** · Expected: **52.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Providence Friars (home)**
  - Actual First-half score: **38–25** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`

### 2026-11-10: Rutgers Scarlet Knights vs NJIT Highlanders (Final: **101–64**)

- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **80.0%** · Price: **1.012 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +25.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+37** · Expected: **66.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **165** · Expected: **60.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Rutgers Scarlet Knights (home)**
  - Expected: **97.5%** · Price: **1.012 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`
- 🟢 **WON — Full-game total 144.5: Over**
  - Actual Full-game total: **165** · Expected: **58.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — First-half moneyline: Rutgers Scarlet Knights (home)**
  - Actual First-half score: **46–26** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +24.5: NJIT Highlanders (away)**
  - Actual Full-game margin: **+37** · Expected: **65.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`

### 2026-11-10: Georgetown Hoyas vs Delaware State Hornets (Final: **83–78**)

- 🟢 **WON — Full-game spread +29.5: Delaware State Hornets (away)**
  - Actual Full-game margin: **+5** · Expected: **68.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Georgetown Hoyas (home)**
  - Actual First-half score: **33–43** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 142.5: Over**
  - Actual Full-game total: **161** · Expected: **66.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +30.5: Delaware State Hornets (away)**
  - Actual Full-game margin: **+5** · Expected: **69.0%** · Price: **1.893 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgetown Hoyas (home)**
  - Expected: **78.0%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 143.5: Over**
  - Actual Full-game total: **161** · Expected: **64.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Georgetown Hoyas (home)**
  - Expected: **97.2%** · Price: **1.001 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

### 2026-11-10: Florida State Seminoles vs Florida Gators (Final: **72–81**)

- 🟢 **WON — Full-game total 164.5: Under**
  - Actual Full-game total: **153** · Expected: **72.0%** · Price: **1.926 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game spread +8.5: Florida State Seminoles (home)**
  - Actual Full-game margin: **-9** · Expected: **87.0%** · Price: **1.952 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game total 166.5: Under**
  - Actual Full-game total: **153** · Expected: **73.0%** · Price: **1.909 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — Full-game moneyline: Florida State Seminoles (home)**
  - Expected: **77.0%** · Price: **4.800 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🔴 **LOST — First-half moneyline: Florida State Seminoles (home)**
  - Actual First-half score: **29–31** · Expected: **75.0%** · Price: **No exact bookmaker price** · Counterfactual fade: **not priced (market_has_no_single_opposite_side)**
  - Calibration: **n=0** · no history · Bucket: `WATCHLIST_NO_ODDS` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game spread +9.5: Florida State Seminoles (home)**
  - Actual Full-game margin: **-9** · Expected: **89.0%** · Price: **1.870 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `team_history_empirical_v1`
- 🟢 **WON — Full-game moneyline: Florida Gators (away)**
  - Expected: **86.2%** · Price: **1.200 · DraftKings** · Counterfactual fade: **not priced (opposite_quote_not_preserved)**
  - Calibration: **n=0** · no history · Bucket: `SKIPPED_INELIGIBLE` · Model: `provider:espn`

## Evidence notes

- Grades use reconciled official finals only; missing or conflicting results stay pending.
- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.
- Flat one-paper-unit ROI is descriptive history, not proof of future performance.
