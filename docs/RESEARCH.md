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
fingerprint provider/model identity and `model_version`; payload-derived
`produced_by` labels remain provenance but arbitrary values are excluded
because they may vary row by row. A closed normalization preserves only the
recognized Predixsport `earlyseason` and `mature` model phases. ESPN BPI and
Predixsport receive explicit adapter-owned versions because their feeds do not
publish stable model releases. These adapter tags
require a manual bump when a provider change is known and cannot detect a silent
upstream change.
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

`MarketForecastService` now emits `espn_player_context_candidates`: rolling
player minutes/points/rebounds/assists from prior ESPN boxscores that were
observed by the forecast cutoff, timestamp-valid last-five results filtered to
pre-target event dates, and timestamped against-the-spread summaries. It also
includes current-game injury entries only when an injury summary or league
snapshot was observed before both the cutoff and tip.
The NBA league injury endpoint is captured once per eligible pregame run; it is
not treated as league-wide coverage unless both teams are present. NCAA summary
injury coverage remains provider-specific and absent entries are not converted
to healthy status. The output preserves observation time, source hash, and
coverage state, and explicitly labels these values candidate-only.

Forecast reports also contain `feature_use_manifest`, schema
`magolide.feature_use_manifest.v1`. It separates active inputs from diagnostic
and shadow-only candidates, provides per-league support distinctions, and
counts point-in-time-visible prediction/book fields on upcoming games. It also
counts stored ESPN summary families and historical results available after the
configured delay. Its raw evidence-resource inventory is descriptive only and
is explicitly not a claim of temporal admissibility or predictive value. NBA
and EuroLeague `smoke` reports now sample their implemented standings adapters
with fetch-time observation and payload hashes; standings remain coverage-only
and are not persisted or used by the model. The current standings adapter path
does not cover WNBA, NCAA, or CBA. ESPN BPI finite numeric
non-core fields are now retained as diagnostic auxiliary metrics (including
NBA matchup quality and dual-sided NCAA metrics), but are not applied to
probabilities, calibration, eligibility, or stakes.

For offline historical research, `scripts/mine_espn_player_minutes.py` reads
already-stored ESPN summaries from a local DuckDB warehouse and writes a
provenance-bearing CSV without network requests or canonical writes. It only
includes reconciled finals and identity-matched teams, preserves raw stat
labels/values and the summary payload hash, and gives a conservative
`available_no_earlier_than_utc` bound. This is descriptive history, not a
same-game feature; EuroLeague and CBA player-level history are not included
because the inspected repository evidence does not currently retain those
sources' athlete boxscores.

`scripts/report_book_movement.py` compares repeated, timestamped, pre-match
`BookSnapshot` rows in immutable slate ledgers for a requested league/date. It
deduplicates identical ledger observations and excludes post-match snapshots
and untimestamped scoreboard open/close context. The result is observed price
movement only—not CLV, a feature, or a value claim. It will report no movement
when fewer than two distinct observations were archived.

These changes instrument coverage and capture leakage-safe candidate evidence;
they do not establish predictive value or improve the model. Missing injury
entries remain unknown, not proof of a healthy squad.

Any future feature promotion still requires an identity-checked observation
available before the simulated tip, an immutable/frozen acquisition time, and
a walk-forward comparison against the existing baseline. Current rich context
does not relax the exact-book/period rules, the strict 30-sample calibration
gate, same-game acca exclusion, or paper-only operation. The exploratory paper
lane may bypass only the sample gate once a segment has at least one graded
non-push sample; it never promotes a row to strict eligibility. Untimestamped
core odds open/close objects and post-game predictor updates do not create
pre-tip history.

## Probability adjustment

Each segment compares observed win frequency with its mean raw forecast
probability. The difference is shrunk toward zero using
`n / (n + prior_strength)`, with prior strength 30 by default. The adjustment
is added to later raw probabilities and clamped away from 0 and 1.

This is deliberately interpretable reliability adjustment, not a claim of
optimal calibration. It avoids choosing markets using future outcomes and
prevents a small sample from receiving the same correction as a large one.
Strict acca eligibility additionally requires at least 30 graded non-push
samples in the exact segment by default. The separate exploratory paper-only
surface can admit a provisional segment from its first graded non-push sample
when sample size is the only blocker. Ticket grading remains separate from the
calibration fit: it measures frozen-paper-ticket performance but does not
bootstrap or double-count the prequential forecast evidence used to adjust
probabilities.

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
