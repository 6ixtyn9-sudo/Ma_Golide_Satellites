# Data and forecasting contract

This document defines the implemented headless contract. DuckDB and its
CSV.GZ materializations are the system of record. Reports, assays, and
notifications are derived outputs and may be regenerated.

## 1. Canonical evidence

The warehouse stores six tables:

1. `games` — schedule identity and reconciled official results;
2. `predictions` — provider/model observations in immutable pre/post slots;
3. `book_snapshots` — attributed bookmaker markets, lines, and decimal prices;
4. `team_aliases` — explicit identity aliases;
5. `provider_health` — acquisition attempts, successes, coverage, and gaps;
6. `pipeline_runs` — operational stage reports with bounded retention.

Each sports row also retains a complete JSON payload column. This is a DuckDB
column for lossless round trips, not a JSON-file datastore. Typed columns and
SQL views expose observed facts without deriving missing provider values.

Prediction evidence and bookmaker evidence have different primary keys and
freeze sets. A predicted winner or expected total never becomes a bookmaker
price or line.

## 2. Time and identity

All timestamps are parsed as timezone-aware UTC values and normalized to ISO
8601. A trailing `Z` alone is not accepted as proof of validity.

For any forecast at `as_of`:

- evidence must have `observed_utc <= as_of`;
- pre-tip evidence must have `observed_utc < scheduled_utc`;
- observations exactly at tip are excluded;
- historical prediction/book rows must belong to the immutable pre-tip freeze;
- only `FINAL` games with a reconciled official result are graded;
- a result is available to later training only four hours after scheduled tip
  by default.

Game joins prefer provider IDs. Team/time matching is a guarded fallback, not
permission to merge contradictory IDs, names, dates, or scores. Existing
provider-event mappings cannot be replaced by a different event through a
name-only match. Schedule rediscovery cannot erase an already reconciled final.

## 3. Market semantics

### Full-game moneyline and price role

The base probabilities are the selected provider's validated home/away
probabilities. Providers are selected in configured priority order; their
probabilities are never averaged. Exact ties produce no selection.

If an exact full-game moneyline book has both prices, the forecast also reports
whether the selected side is the `favorite` or `underdog` by decimal price.
These are roles on one canonical moneyline forecast, not separate prediction
models. A missing book leaves a research forecast but no bettable selection.

### Full-game and quarter totals

Full-game totals use the provider's expected score/total as a baseline when
available, adjusted with strictly historical empirical residuals. Quarter
baselines use recency-weighted, shrunk team scoring rates and quarter-specific
residuals.

The over/under decision is evaluated against the **posted bookmaker line**.
The provider expected total is never copied into `Forecast.line`. Without an
exact total snapshot for the requested period, no total wager is emitted.

### Quarter spreads

Q1–Q4 spreads use the same history-only period margin distribution. Lines are
stored and interpreted as home-oriented handicaps: home cover value is
`home_score + line - away_score`. The exact period and both side prices are
required for acca eligibility.

### First-half 1X2

The first-half margin distribution produces `home`, `draw`, and `away`
probabilities. The selected outcome must have a matching first-half h2h book.
A selected draw specifically requires `draw_price_decimal`; home/away prices
alone are insufficient.

### Highest quarter

A highest-quarter forecast compares four empirical quarter-total simulations
and selects Q1, Q2, Q3, or Q4. Historical grading treats “highest or tied for
highest” as a research win. Bookmaker dead-heat settlement may differ, so an
unpriced highest-quarter forecast is never acca eligible. No current book
adapter fabricates this market.

## 4. History-only model

`HistoryModel` builds baselines from games strictly before the target and
available under the result delay. Team rates are recency weighted with a
half-life and shrunk toward the league mean. Margin and total residual samples
are generated in chronological order: every residual's own baseline can use
only games available before that residual's target.

The distribution is empirical. Probabilities are counts over residual
scenarios, not claims from a fitted sportsbook or a certified probability
model. Exact probability ties are not broken arbitrarily. Totals and spreads
report push probability separately.

Full-game moneyline remains a provider-probability forecast. The other market
families are historical/residual forecasts; full-game totals may combine a
provider baseline with historical residuals. Every forecast identifies its
model source and evidence timestamps.

## 5. Walk-forward reliability harness

For each reconciled final, the service simulates a forecast at that game's
tip:

1. remove later and unavailable finals;
2. select only frozen pre-tip prediction/book evidence;
3. generate all market forecasts possible at that cutoff;
4. fit reliability profiles from earlier graded records only;
5. calibrate the simulated forecast;
6. grade it against the official result and append it for later rounds.

Books are deduplicated to one observation per game/model/market/source in the
training rows, preventing a game with many books from dominating a segment.
Unreconciled finals and ungradeable rows are excluded. Pushes do not count as
wins or losses in the reliability adjustment.

Profiles are segmented by `(league, market, period, source)`. For a segment
with `n` non-push outcomes:

```text
observed adjustment = observed win rate - mean raw predicted probability
shrink weight        = n / (n + 30)
applied adjustment   = shrink weight * observed adjustment
calibrated p         = clamp(raw p + applied adjustment, 0.01, 0.99)
```

The report includes segment sample size, observed win rate, mean raw
probability, adjustment, one-sided Wilson lower bound, Brier score, and log
loss. Raw and prequential-calibrated overall metrics are reported separately.
These are diagnostics, not certification. The default minimum for acca use is
30 graded non-push samples in the exact segment.

## 6. Acca eligibility and ranking

A current forecast can supply an acca leg only if:

- its exact visible pre-tip bookmaker market and selected decimal price exist;
- any provider prediction used by the forecast is visible and strictly pre-tip;
- its segment has at least the configured walk-forward sample minimum;
- calibrated probability is at least 0.50 by default;
- estimated push probability is at most 0.005 by default.

The highest-calibrated-probability eligible forecast is selected per game.
Same-game legs are prohibited rather than silently multiplied as independent.
For the requested leg count, a best-first search ranks combinations by:

```text
estimated joint win probability = product(calibrated leg probabilities)
```

The search is exact for that objective under the stated cross-game conditional
independence assumption. Combined decimal odds and estimated EV may be shown
as secondary arithmetic, but neither determines ranking.

Latent cross-game dependence is not estimated. The output therefore means
“highest estimated probability among eligible combinations under these model
assumptions.” It does **not** mean guaranteed, certain, empirically certified,
or the global optimum of a validated dependence model.

## 7. Freezing, reports, and writes

Pregame acquisition stores observations without replacing an existing frozen
value. The freeze stage fixes the latest eligible pre-tip prediction per
provider and the latest exact pre-match book market slots. Attempts to change
a frozen value raise `FrozenWriteError`; idempotent replay of identical bytes
is allowed.

The daily current-market report is generated from visible pre-tip evidence
before the freeze stage. Tip-passed prediction/book rows are then frozen for
future historical replay. Current forecasts report their freeze state; being
not yet frozen does not block a current acca, but historical training always
requires the later immutable freeze. Unpriced forecasts remain visible for
research and cannot enter accas.

Canonical writes require `--write` and `ALLOW_MAGOLIDE_WRITES=true`. Fixture
persistence additionally requires `ALLOW_FIXTURE_WRITES=true` and isolated
paths. Notification delivery is independently guarded. No code path places a
bet, writes a bet slip, or certifies a winner.
