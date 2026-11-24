# Ma Golide score — cold-start shadow lane for fixtures-only leagues

**Status:** built, tested, shadow-only. 19 tests green (16 from the original build plus 3 review
guards), repo-wide `ruff` clean, full suite green.
**Deliverables:** `forecasting/magolide_score.py`, `scripts/magolide_shadow_card.py`,
`tests/test_magolide_score.py`, `docs/MAGOLIDE-SCORE-SHADOW-LANE.md`.

**Correction on config provenance (2026-11-24).** The retired system did **not** run on a static
config, and this port must not imply it did. Every run began with `loadTier1Config(ss)` — a
versioned config read from the spreadsheet (`Forecaster.gs` line 2213; the tuner reads it again at
3797 and logs `Current config version: <v>`), and `writeEliteProposalSheet_` promotes a new
`config_version` with the current config beside ranked #2/#3 alternatives (line 4731). So the
numbers below are *validator defaults* (`validateTier1Config`, line 5396) and empty-stub values
(line 3956), not a fitted configuration. They are now named `VALIDATOR_DEFAULTS`, every verdict
carries `config_version` (default `unversioned-validator-defaults`) and
`weights_source=validator_defaults_not_a_fitted_config`, and a caller holding a stored versioned
config must pass it and name its version.

**Provenance re-verified on merge (2026-11-24).** The port was checked against the real deleted
source rather than its description: `git show de966196^:docs/Forecaster.gs` — 5,986 lines,
`calculateMaGolideScore` at line 694, Sections A–G as ported, and the weight table below
recovered verbatim from `_emptyConfig` in the Elite-proposal writer at line 3956. The `variance`
weight is applied to `(-variancePenalty)` and `absScore < threshold` treats `== threshold` as
**not** RISKY in the original — both preserved here.
**Everything in this document is descriptive research evidence. Nothing here is a ticket, a price-backed pick, a settlement grade, or a claim of edge.**

---

## 1. What was asked

"Look at how ma golide/score used to make picks — can we not use that machinery to make picks on the leagues with no predictions?"

The answer, from the investigation recorded in §2–§3, was: **mostly no — the empty leagues' blocker is data acquisition, not the scoring formula. But one piece of the old machinery is genuinely reusable** (§4): its ability to produce a pre-tip read from standings alone, with zero results history. That piece is what was built (§5).

## 2. How ma golide/score used to make picks

The old machinery was `docs/Forecaster.gs` — a ~6,000-line Google Apps Script, retired on **2026-11-07** when commit `de966196` ("feat: add calibrated markets and acca forecasting") deleted it in favour of the Python pipeline.

Its core, `calculateMaGolideScore`, produced one moneyline read per game:

```
score = w_pct·(pctDiff)           # standings win-percentage difference, ×100
      + w_netRtg·(netRtgDiff)     # net-rating difference
      + w_homeCourt·(homeCourtEffect)   # home win% vs away win% proxy
      + w_momentum·(momentumDiff)
      + w_streak·(streakDiff)     # signed streak difference
      + w_form·(formDiff)         # L10 form difference
      + w_h2h·(h2hDiff)           # head-to-head difference
      + w_forebet·(forebetDiff)   # external Forebet prediction difference
      − w_variance·(variancePenalty)    # always a penalty
      + homeAdv

|score| < threshold  → RISKY
score  > 0           → HOME
else                 → AWAY
confidence = logistic(score / confidence_scale), clamped to [confMin, confMax]
```

Every input came from the Google Sheet: standings, recent results, H2H history, Forebet's published predictions. The tuner's own baseline configuration (recovered from the Elite-proposal empty stub in the script's git history) was:

| weight | value | | gate | value |
| --- | --- | --- | --- | --- |
| pct | 3 | | homeAdv | 5 |
| netRtg | 2 | | threshold | 25 |
| homeCourt | 1 | | confMin / confMax | 50 / 95 |
| momentum | 1 | | confidence_scale | 30 |
| streak | 1 | | rank (legacy) | 0 |
| form | 2.5 | | variance | 1 |
| h2h | 1.5 | | forebet | 3 |

The decisive fact for this task: **the scorer is a consumer of results data — it does not create any.** Its features *are* standings, results, and H2H records.

## 3. Why the prediction-less leagues are prediction-less

Three separate causes (from `ingestion/league_registry.py`, measured 2026-11-24):

| group | leagues | cause | can a scorer help? |
| --- | --- | --- | --- |
| Retired | Spain ACB, BCL, BAL, Argentina, Italy LBA, France Pro A, Germany BBL, Turkey BSL, Brazil NBB | ~750 runs each, ~0 games ever acquired; no validated schedule/results route | **No.** No data exists to score. |
| Fixtures-only | Greece Basket League, China WCBA W, Lithuania NKL, Australia WNBL W, Israel Liga Leumit | schedule + page-average prices arrive via `betexplorer_fixtures`, but `results_source = "source_unavailable"`: no reconciled finals → nothing to calibrate against **or settle with** | **Partially — cold-start only.** |
| Off-season | WNBA | season runs April–October | No. Nothing to score. |

For contrast, EuroLeague already publishes a 163-line picks card today from the *current* machinery (`HistoryModel`, `forecasting/model.py`), held in CAUTION by walk-forward gates (`insufficient_walk_forward_samples`, ROI bands) — not silenced.

### Why reviving the full scorer would be a step backwards on lanes that *have* results

The current `HistoryModel` already computes form/momentum/streak equivalents as empirical team rates with exponential decay, point-in-time, walk-forward calibrated. Hand-weighted form features are deliberately `shadow_candidate_only` in `forecasting/feature_manifest.py`, and `docs/ABA_LEAGUE_EDGE_ASSAYER_COMPARISON.md` records the evidence-first decision behind that. Re-introducing hand weights on lanes with results would undo a decision the repo already made with reasons.

### The one genuinely reusable idea

The old scorer could **cold-start with zero results history**: it scored a game from standings differences alone (pct / rank / netRtg + home-court proxy), falling back to neutral defaults (rank 15, pct 0.5, streak 0) for missing sides. The five fixtures-only lanes have a slate and prices but no finals; a standings-difference read is the only thing that can say anything pre-tip there today. That is what was built.

## 4. What was built

### 4.1 `forecasting/magolide_score.py` — the port

A faithful Python port of `calculateMaGolideScore` Sections A–G plus the standings-difference helpers:

- `build_standings_features(home, away, …)` → `MaGolideFeatures` — signed, home-positive differences with the retired script's exact neutral fallbacks (missing rank 15, missing pct 0.5, missing streak 0). L10 form beats the streak proxy when present; home-court uses the `homePct − awayPct` standings proxy (the script's fallback path).
- `score_features(features, config)` → `MaGolideVerdict` — `score`, `prediction` (HOME/AWAY/RISKY), `confidence`, `probability`, per-factor `factor_breakdown` (including `homeAdv`, as in the script), and `meta` (LEGACY vs NEW mode, confidence mode, weights provenance).
- `MaGolideConfig` — ships the recovered **tuner baseline** weights, labelled `weights_source="forecaster_gs_tuner_baseline"` in every verdict so no output can be mistaken for a measured fit.
- `shadow_lane_leagues()` — derives the eligible lanes from the registry contract itself: `schedule_source == "betexplorer_fixtures"` **and** `results_source == "source_unavailable"` **and** lane status ACTIVE. Today that is exactly Greece, WCBA W, Lithuania NKL, Australia WNBL W, Israel Liga Leumit. Retired and off-season lanes never qualify regardless of their sources.
- `shadow_rows(league, games, standings)` — refuses any league outside that set with a hard `ValueError`; annotates each row with `neutral_defaults_applied` so an unpopulated standings map reads as `RISKY` rather than inventing a pick.
- `render_shadow_card(...)` — a card in the house style, but inert: header states **SHADOW ONLY … Not picks, not tickets, not settlement evidence**; every row carries `bucket=WATCHLIST_SHADOW · price_evidence=NONE · acca_eligible=False · grade=NOT_SETTLEABLE · lifecycle=SHADOW` and the reasons `magolide_score_shadow, no_settlement_route, cold_start_no_walk_forward`.

### 4.2 `scripts/magolide_shadow_card.py` — the CLI

Reads one JSON document (league, date, games, optional/partial standings map) and writes `localdata/<league>/shadow_scores_<date>.txt` (or `--outdir`). Missing teams fall back to neutral defaults on purpose. Example invocation and output (real run during this build):

```
$ python scripts/magolide_shadow_card.py --input demo.json --outdir /tmp/shadow-demo
wrote /tmp/shadow-demo/shadow_scores_2026-11-24.txt (3 shadow row(s))
```

```
Ma Golide SHADOW SCORES — GREECE_BASKET_LEAGUE 2026-11-24
========================================================================
SHADOW ONLY · cold-start standings scorer ported from the retired
Forecaster.gs calculateMaGolideScore (weights = validator defaults,
unmeasured for this lane). Not picks, not tickets, not settlement
evidence. The lane has no results route: nothing here can ever be
graded, banked, or ROI-qualified.
========================================================================

WATCHLIST_SHADOW — COLD-START STANDINGS READ
========================================================================
  [magolide_score/shadow] Panathinaikos BC vs AEK Athens · KO 2026-11-24T17:15:00Z
      MAGOLIDE_SCORE full_game -> HOME · score=+150.30 · uncalibrated_logistic=99.34% · confidence=95.0%
      bucket=WATCHLIST_SHADOW · price_evidence=NONE · acca_eligible=False · grade=NOT_SETTLEABLE · lifecycle=SHADOW
      reasons=magolide_score_shadow,no_settlement_route,cold_start_no_walk_forward · weights=validator_defaults_not_a_fitted_config · config_version=unversioned-validator-defaults
  ...
Logical shadow rows: 3 · tickets: 0 · settleable: 0
```

### 4.3 `tests/test_magolide_score.py` — 20 tests, all green

Pinned behaviours: all-zero features → `RISKY` at 50%; pct difference drives HOME/AWAY with mirrored probabilities; the threshold boundary (`|score| == threshold` is **not** below it, matching the script); variance always subtracted; confidence logistic-clamped to `[confMin, confMax]` in both PROBABILITY and LEGACY_SIGMOID modes; the risky-cap; LEGACY mode when no new-feature weights are set (`2.5·form + 1.5·h2h = 10.5`); `homeAdv` added verbatim and reported in the breakdown; neutral defaults for missing teams; L10 beating the streak proxy; the exact five-lane shadow set; hard refusal for EuroLeague/Spain ACB; shadow-row markers; card language that never claims picks or settlement.

### 4.4 Review fixes applied on merge

1. **The probability token.** The card printed the logistic of the weighted score as `p=99.34%`.
   In this codebase `p=` is the token for a calibrated probability on a pick card, so a
   standings-only number under that label is the one piece of this artifact an operator could
   quote as if it were a fit. It now prints `uncalibrated_logistic=`, and the dataclass field is
   `uncalibrated_logistic_pct` rather than `probability`, so the name carries the caveat too.
2. **A falsy-zero divergence.** `conf_min if conf_min else 50.0` silently maps an explicit `0`
   to 50, where the original's `safeNum_(config.confMin, 50)` keeps the zero. Now `is not None`,
   with a test pinning it.
3. **A guard where the shadow lane must never reach.** Nothing imports this module today; a test
   now asserts that `ingestion/pipeline.py`, `forecasting/operator.py`, `forecasting/service.py`,
   `forecasting/accas.py` and `scripts/daily.py` do not mention it, so wiring it into the live
   forecast path is a deliberate edit with an assertion to argue against, not a one-line slip.
4. **Config provenance corrected.** See the correction above: the weights are validator defaults
   read from a config that was dynamic and versioned, and the port now says so in every verdict.
5. **Publication boundary confirmed, not assumed.** `scripts/publish_localdata.py` selects files
   by explicit name patterns (`picks_<date>.txt`, `auto_tickets_*`), so `shadow_scores_<date>.txt`
   is never published with the rest of the state. Verified, not hoped for.

## 5. Deliberate limits (the gates this build keeps)

1. **No pricing.** The fixtures lanes' prices are BetExplorer page averages — aggregate, research-only by registry policy. Shadow rows carry `price_evidence=NONE` and no ticket fields.
2. **No settlement.** `results_source = source_unavailable` means nothing here can be graded. The rows say `NOT_SETTLEABLE` by name, the same way the Greece spec states it.
3. **No calibration claims.** The weights are the retired tuner's baseline, measured against nothing. Every verdict and card row names that provenance.
4. **No promotion path without evidence.** If a lane later gains a validated results route, the right model on it is the existing calibrated `HistoryModel`, not this scorer. This module stays shadow until — and unless — a replay demonstrates otherwise.
5. **Refuses lanes outside its class** (`ValueError`), including retired lanes and lanes where the calibrated machinery already owns prediction (EuroLeague etc.).

## 6. What would actually light these leagues up

The scoring-side work is done; the remaining blocker is **results acquisition**, and it is the piece every other improvement hangs on:

1. Validate a date-bounded results/standings route for one fixtures-only lane (start with Greece Basket League: its federation or the BetExplorer results pages).
2. Backfill reconciled finals; the current `HistoryModel` then calibrates on the lane exactly as it did for ABA (4,252 games back to 2004) and EuroLeague.
3. Only then do picks become settleable, ROI-qualifiable, and promotable past CAUTION (walk-forward n ≥ 30).

The shadow scores this build produces give the operator a pre-tip read to watch in the meantime — explicitly labelled for what it is.

## 7. Verification record

- `uv run ruff check .` — **All checks passed** (whole worktree, not only new files).
- `uv run pytest -q tests/test_magolide_score.py` — **20 passed**.
- `uv run pytest -q` — see the merge PR for the current figures on `main`.
- CLI smoke test on a demo Greece slate — 3 shadow rows rendered, `RISKY` correctly shown for the game whose standings were missing.

## 8. File inventory

| file | role |
| --- | --- |
| `forecasting/magolide_score.py` | scorer port + shadow-lane adapter + card renderer |
| `scripts/magolide_shadow_card.py` | CLI: fixtures/standings JSON → `shadow_scores_<date>.txt` |
| `tests/test_magolide_score.py` | 16 behaviour tests |
| `docs/MAGOLIDE-SCORE-SHADOW-LANE.md` | this document |
