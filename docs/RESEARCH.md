# Research methodology

Ma Golide has two related but distinct analysis paths:

1. `ingestion.mining` + `assayer`: descriptive provider-prediction outcome
   statistics;
2. `forecasting`: chronological multi-market simulation, calibration, current
   forecasts, and acca ranking.

Neither path certifies a betting edge, guarantees a winner, or places a bet.

## Provider prediction assay

```bash
python scripts/mine_assay.py \
  --warehouse localdata/warehouse.duckdb --league NBA \
  --output /tmp/prediction-assay.json
```

The assay joins official home/away winners to visible pre-tip provider
probabilities. Independent exclusion reasons distinguish missing predictions,
post-tip shadow observations, incoherent probability pairs, missing results,
and ties. It does not convert an expected total into an over/under bet.

Segments can describe provider, confidence band, margin band, expected-total
band, and early/late season behavior. Output includes sample sizes, observed
accuracy, shrunk rates, and Wilson lower bounds where applicable. Small
segments are descriptive and should not be promoted into production rules.

## Multi-market walk-forward harness

```bash
python scripts/build_markets.py \
  --warehouse localdata/warehouse.duckdb --league NBA \
  --as-of 2026-11-07T18:00:00Z --output /tmp/markets.json
```

For every historical final, the harness reconstructs what was knowable at that
game's tip. It uses only frozen pre-tip predictions/books and earlier results
that satisfy the availability delay. Forecasts are graded before being added
to the history for subsequent games.

The resulting `walk_forward` report section contains:

- `historical_finals`: reconciled finals visible to the top-level report;
- `graded_forecasts`: historical forecasts with valid outcomes;
- `raw_metrics`: unadjusted probability diagnostics;
- `prequential_calibrated_metrics`: diagnostics after each round's past-only
  champion adjustment;
- `segments`: versioned per-league/market/period/source calibration profiles;
- `shadow_challenger`: prequential raw/champion/recency-decayed metrics,
  coverage, effective sample sizes, and the fixed challenger configuration;
- `guards`: the leakage and grading exclusions applied.

Brier score and log loss evaluate probability quality; accuracy alone does
not. One-sided Wilson lower bounds describe uncertainty but are not edge
certificates. Pushes are reported separately and excluded from win/loss
calibration.

History-model calibration segments include a stable fingerprint of the model
revision and prediction-affecting `ForecastConfig`. Provider-prediction segments
fingerprint provider/model identity, `model_version`, and `produced_by`; ESPN
BPI and Predixsport receive explicit adapter-owned versions because their feeds
do not publish stable model releases. These adapter tags require a manual bump
when a provider change is known and cannot detect a silent upstream change.
Per-game artifact revisions are excluded. Hybrid segments fingerprint both
sources, and legacy records remain separate. The service runs the recency-decayed
challenger on the same chronological rows and result-delay boundary. Current
forecasts expose its candidate probability and effective sample size as
separate shadow diagnostics only. The acca engine continues to rank on the
champion probability, and any promotion requires untouched prospective data.

## Rich ESPN context and leakage boundary

Summary acquisition can retain team/player boxscores, leaders, injuries,
recent form, predictor context, and coverage counts for play-by-play,
win-probability, shot-chart, and standings resources. These are evidence
coverage, not automatically model inputs. A post-game summary can contain the
game's own player statistics and in-play win probability; using either as a
pre-tip feature would leak the result.

A future context feature therefore needs an identity-checked observation that
was available before the simulated tip, an immutable/frozen acquisition time,
and a walk-forward comparison against the existing baseline. Current rich
context does not relax the exact-book/period rules, the 30-sample calibration
floor, same-game ACCA exclusion, or paper-only operation. Untimestamped core
odds open/close objects and post-game predictor updates do not create pre-tip
history.

## Probability adjustment

Each segment compares observed win frequency with its mean raw forecast
probability. The difference is shrunk toward zero using
`n / (n + prior_strength)`, with prior strength 30 by default. The adjustment
is added to later raw probabilities and clamped away from 0 and 1.

This is deliberately interpretable reliability adjustment, not a claim of
optimal calibration. It avoids choosing markets using future outcomes and
prevents a small sample from receiving the same correction as a large one.
Acca use additionally requires at least 30 graded samples in the exact
segment by default.

## Dependence policy

Same-game legs can be strongly dependent—for example, a side and a low total.
The acca builder therefore prohibits them completely rather than multiplying
their probabilities or pretending to know their correlation.

Across games, joint probability is modeled as the product of calibrated leg
probabilities. This conditional-independence assumption may still be wrong
because of shared league conditions, scheduling, model error, or other latent
factors. The limitation is attached to every returned acca.

The best-first combination search is exact for the product objective after
selecting the strongest eligible forecast per game. It does not search over
same-game parlays and does not establish an empirically certified maximum.

## Price and settlement limitations

- A model probability is not a bookmaker quote.
- Expected scores/totals are baselines, never substitute total lines.
- Highest-quarter research grading counts ties as wins; real dead-heat rules
  may settle differently, so unpriced HQ forecasts are ineligible.
- Totals and spreads with material push probability are excluded unless the
  configured policy and book rules support them.
- Historical book prices describe observed quotes, not guaranteed available
  execution, limits, closing value, or slippage.
- Estimated joint probability and EV arithmetic do not include stake limits,
  void rules, account restrictions, or market removal.

## Appropriate interpretation

Use the reports to inspect data coverage, compare historical probability
quality, identify poorly calibrated segments, and rank explicitly eligible
combinations under stated assumptions. Do not describe output as a lock,
certainty, guaranteed return, certified maximum, or production-validated
betting system without separate prospective evidence and execution controls.
