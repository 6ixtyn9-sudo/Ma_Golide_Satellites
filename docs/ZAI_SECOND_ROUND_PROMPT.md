# Z.ai Second Round — Why still NO AUTO TICKETS after 4 free fixes

**Pre-read:** First Z.ai audit (2026-11-09) is in conversation history. Do NOT re-prove ESPN 404 history, DK timing 1/6 vs 8/8, or BPI pre-tip. Assume those.

## Current state on main 9f03282 (after PR #20 merges)

- **History window:** 21/21/50 (was 7/7/25 → 10/10/50 → 21/21/50)
- **Book free fixes just shipped:**
  - `ingestion/providers/espn.py`: open→close fallback for moneyline/spread/total when close not posted
  - `pipeline.py: BOOK_STALENESS 24h→72h`, `+ fix missing Market import`
  - `_espn_odds_catchup()` now triggers on any missing core market (moneyline/total/spread) within 36h, not just zero snapshots
- **Calibration:** `forcing/harness.py` keeps eligible 4-tuple profiles, adds diagnostics-only 5-tuple `team_history_insufficient` pool (visible in `walk_forward.segments`, never applied)
- **Tests:** 260 passed

**Latest live funnel (run 42618504921, before 9f03282, 2026-11-09 14:37 SAST):**

NBA (38 rows):
```
38 scan rows · 22 exact priced · 0 calibration-sufficient
0 priced + calibrated · 0 eligible games
Blockers: 38 insufficient_walk_forward, 30 insufficient_team_specific_history,
          16 no_exact_bookmaker_market, 8 below_calibrated_probability_floor
```
NCAA_M (183 rows):
```
183 scan rows · 91 exact priced · 0 calibration-sufficient
Blockers: 183 insufficient_walk_forward, 133 insufficient_team_specific_history, 92 no_exact
```

**Sample forecast that should have been eligible but isn't:**
```json
{
  "team_history": {"home_effective_games": 1.88, "away_effective_games": 2.51, "required": 8.0, "sufficient": false},
  "eligibility": ["no_exact_bookmaker_market","insufficient_walk_forward_samples","insufficient_team_specific_history","below_calibrated_probability_floor"],
  "calibration": {"n":0, "status":"no_history", "sufficient": false},
  "training_games": 146, "residual_games": 54
}
```

Effective games mean 1.82 across NBA pick_today, despite `training_games=146` (league-wide) and `residual_games=54`. `team_prior_games=8.0` is also `required_effective`.

## Ask Z.ai to dig EVEN DEEPER (stay free — no paid odds keys)

1. **effective_games math deep dive:**
   - `forecasting/model.py: _team_rates()` weight = 0.5**(age/20) sum → effective = weight_sum. Verify against early-season reality: 21 days ≈ 7-10 raw team games → effective 6.3-8.6. Why is live effective only 1.1-2.5? Is it per-team filtering vs league-wide history? Is `training_games` misleading? Should `recent_half_life_games=20` or `team_prior_games=8.0` be tuned, or history window 21→45 days, or should we use league prior longer?
   - Calculate how many raw team games needed to reach 8.0 under current decay. How many schedule cron cycles (every 3h with cursor advance only on schedule) to grow from 2→8?

2. **walk_forward coupling — zero is structural:**
   - `harness.py: build_profiles()` excludes `sufficient_for_accas is False` from eligible pool, so `walk_forward.segments` n=0 forever until team_history crosses 8.0 first. Diagnostics 5-tuple pool exists but never applied. Is this intent vs catch-22? Should we seed calibration with 15-20 samples provisional, or keep 30 hard? Show p50 trajectory.

3. **Over/under bias — why unders don't fire:**
   - Historic totals: 52 forecasts `over 28 / under 24` balanced, but settled `over 14 / under 6`. Sample total row: `expected_total 234.98 vs line 235.5 → under 53%`, but over is picked 68% when baseline 221 vs 214.5. Is residual mean >0 (baseline undershoots book)? Should we center residuals on `baseline.total - league avg`? Should `expected_total` blend BPI vs empirical? Provide residual distribution stats.

4. **NCAA player stats — as much info as possible:**
   - Inventory what ESPN summary actually returns for NCAA_M/W vs NBA: `statistics`, `leaders`, `injuries`, `records`, `againstTheSpread`, `venue`, `broadcasts`, `recent form`, `coverage counts` (play-by-play/win-prob/shot-chart/standings). Which fields are pre-tip safe vs post-game leak? Which are already retained in `game.evidence.espn.summary` vs dropped? Is there a free player boxscore endpoint (per-team/per-player points/minutes/efficiency) usable pre-tip (last 5 games, not current game) for NCAA? Prior audit listed 6 usable fish (BPI matchupQuality, boxscore avgPoints 121.9 L10 8-2, etc.) — expand for NCAA specifically.

5. **Every potential market — free exhaustive coverage:**
   - User wants "every potential market we can bet on" but stay free. Map: ESPN free gives moneyline/spread/total full-game only. What other free markets exist (ESPN Q1-Q4/H1, hardwood artifact probabilities only, etc.)? What stays paid-only (player props, alt lines, team totals)? Conclude: is free exhaustive impossible, or can we at least add Q/H1 from ESPN where available?

**Deliverable:** Concrete payload evidence (which ESPN endpoint fields exist, which pipeline retains, which model could use), plus 2-3 ranked free-only fixes that would move funnel from `0 eligible games` → `≥3 games` without lowering 8.0/30/exact-price paper guards, or state why guards must be tuned.

**Context file refs:**
- `forecasting/model.py:120-210` team_rates/baseline/effective_games
- `forecasting/harness.py:17-220` CalibrationConfig 30/0.50/8.0
- `ingestion/providers/espn.py:340-520` open→close, book_snapshots_from_event
- `ingestion/pipeline.py:420-530` discover/catchup, BOOK_STALENESS 72h
- `localdata/nba/picks_today.json` 38 rows, 146 training, 1.88 effective
- `localdata/nba/auto_tickets_2026-11-09.txt` funnel 38/22/0
- `docs/ESPN-ACQUISITION.md`, `docs/PROVIDERS.md`, `docs/RESEARCH.md`, `docs/OPERATIONS.md`
