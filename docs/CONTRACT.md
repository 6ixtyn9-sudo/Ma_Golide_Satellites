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

Two independent model surfaces can emit full-game moneyline forecasts:

- the selected provider's validated home/away probabilities, using configured
  provider priority without averaging providers;
- the empirical history model's full-game margin distribution, which requires
  no provider prediction and can therefore be reconstructed at every historical
  tip for walk-forward calibration.

Their sources and calibration profiles remain separate. Neither surface inherits
the other's samples or probabilities, and exact ties within a surface produce no
selection from that surface. Full-game selections are two-way (`home` or `away`);
an empirical simulated draw is not mapped to a two-way book selection.

If an exact full-game moneyline book has both prices, the provider forecast also
reports whether the selected side is the `favorite` or `underdog` by decimal
price. This is a role on the provider forecast, not a separate prediction model.
A missing book leaves research forecasts but no bettable selections.

### Full-game and quarter totals

Full-game totals use the provider's expected score/total as a baseline when
available, adjusted with strictly historical empirical residuals. Quarter
baselines use recency-weighted, shrunk team scoring rates and quarter-specific
residuals.

The over/under decision is evaluated against the **posted bookmaker line**.
The provider expected total is never copied into `Forecast.line`. Without an
exact total snapshot for the requested period, no total wager is emitted.

### Spread lines

All spread lines (full game and sub-periods) are stored home-oriented. The
home-side handicap is `line`; the away-side handicap is `-line`. For example,
a stored line of `-27.5` means the home team is laying 27.5 and the away team
is receiving `+27.5`. Grading uses the home cover value
`home_score + line - away_score`. Human-readable reports must display the
selected side's handicap, not the canonical home-oriented line beside an away
selection. The exact period and both side prices are required for acca
eligibility.

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

Full-game moneyline has both the provider-probability surface described above
and a historical/residual surface. The other market families are
historical/residual forecasts; full-game totals may combine a provider baseline
with historical residuals. Every forecast identifies its model source and
evidence timestamps. Because the empirical full-game surface needs only prior
finals, lack of a frozen provider-prediction archive does not prevent its
prequential calibration.

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
History-based rows that fail the team-specific evidence guard are also excluded
from their segment's reliability fit, so fallback rows cannot manufacture an
ACCA-eligible sample count. Unreconciled finals and ungradeable rows are
excluded. Pushes do not count as wins or losses in the reliability adjustment.

Profiles are segmented by `(league, market, period, source,
calibration_version)`. History-model forecasts stamp a deterministic fingerprint
of `MODEL_VERSION` and all prediction-affecting `ForecastConfig` fields.
Provider-prediction forecasts fingerprint provider/model identity and the
`model_version`. Payload-derived `produced_by` labels remain provenance for
display but are excluded from segmentation because they may vary per row. A
closed Predixsport normalization recognizes only its exact `earlyseason` and
``mature` model-phase suffixes, preserving that known distinction without
fingerprinting arbitrary text. Hardwood carries an artifact version; ESPN BPI
and Predixsport use explicit adapter-owned versions because their feeds do not
expose a stable model release. Bump those adapter versions when a known
provider-model or normalization change occurs. They cannot reveal a silent
upstream change on
their own. Per-game source revisions and observation times are deliberately
excluded. Hybrid history/provider forecasts include both configurations.
Previously stored rows without a calibration version retain the explicit
`legacy` segment tag and never merge with stamped history.
A version change starts a separate calibration sample rather than borrowing
another version's evidence. For a segment with `n` non-push outcomes:

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

The same `walk_forward` object includes `shadow_challenger`: raw, champion, and
recency-decayed challenger Brier/log-loss metrics, coverage, effective sample
size, the challenger fingerprint, and the result-availability delay. Historical
rows are evaluated prequentially with the service's exact calibration epoch,
champion settings, and result delay. Current forecast explanations may include
a separate shadow candidate probability, status, sample count, and effective
sample size. Evaluation or annotation failure marks the shadow result
`unavailable` while preserving the normal champion report. Shadow fields remain
diagnostic only: `build_accas` never consumes challenger probabilities. Its
strict mode uses champion `calibrated_probability` and the existing eligibility
flag; the service's paper-only mode also checks champion calibration status and
eligibility reasons solely to recognize the sample-only exception.

## 6. Acca eligibility and ranking

A current forecast is **strictly eligible** to supply an acca leg only if:

- its exact visible pre-tip bookmaker market and selected decimal price exist;
- any provider prediction used by the forecast is visible and strictly pre-tip;
- its segment has at least the configured walk-forward sample minimum;
- every history-based forecast has at least as much effective prior evidence for
  each team as the configured league-prior strength (eight effective games by
  default), so a league-average fallback cannot masquerade as team evidence;
- calibrated probability is at least 0.50 by default;
- estimated push probability is at most 0.005 by default.

A separate **paper-only exploratory** report surface may bypass only the sample
minimum: a forecast must have at least one graded non-push sample, a provisional
calibration profile, and `insufficient_walk_forward_samples` as its sole
eligibility reason. Exact price, team-history, push-risk, and probability-floor
gates remain enforced. `Forecast.eligibility.acca_eligible` stays false for
these rows. `MarketForecastService` keeps `accas.results` (and its explicit
`accas.strict_results` alias) strict, and emits the expanded exploratory set as
`accas.paper_results` for the paper-ticket publisher only. The default
`build_accas()` behavior remains strict.

Before paper-ticket assembly, the current paper-eligible candidate pool is
coalesced to one forecast row per exact wager selection (league, game, market,
period, selection, and line), keeping the highest-calibrated-probability row for
duplicate model forecasts. Each pick is then checked against settled
individual-pick ROI for its league, market, period, model source, calibrated-
probability band, and individual-decimal-odds band. Probability bands are
`<0.50`, `0.50-<0.60`, `0.60-<0.70`, `0.70-<0.80`, `0.80-<0.90`, and
`0.90-1.00`; odds bands are `1.00-<1.20`, `1.20-<1.50`, `1.50-<2.00`, and
`2.00+`.

The historical source is the rolling audit's `settled_ledger` when available,
otherwise past-only prequential settled individual forecasts. Pending rows,
future results, missing calibrated probabilities, and rows without an exact
pre-tip price do not count. Flat one-unit profit is `price - 1` for a win, `-1`
for a loss, and `0` for a push; pushes remain in the one-unit ROI denominator.
History is weighted with a 14-day half-life and the band estimate is shrunk
toward model, market, then league-level ROI using three effective prior picks.
There is no long sample-size qualification delay: one settled pick can seed a
band, but the candidate's own band must have positive recency-weighted ROI and
positive adjusted ROI. A band with no settled history cannot inherit eligibility
from a profitable parent. This decision is recomputed on each scan as outcomes
settle; acca-ledger outcomes are never used as individual-pick evidence. A
positive band is an exploratory paper-only signal, not statistical validation
or proof of future profit.

If too few independent qualifying picks remain, no paper acca is emitted. This
performance gate does not relax or add a calibration-sample gate: the existing
sample-only provisional path and all other strict eligibility safeguards remain
in force.

The paper-ticket surface then greedily accepts combinations in descending joint
probability order and does not reuse an exact wager selection across tickets.
A different selection from the same game may be used on another ticket, while
same-game legs remain prohibited within any one acca. This is a no-repeated-leg
rule, not mathematical outcome exclusivity: separate tickets may still win
simultaneously. The strict `results`/`strict_results` analytical ranking is
unchanged and can contain overlapping combinations. Fewer paper tickets are
returned when the requested limit exceeds the available non-repeated leg sets.

Frozen paper receipts label every combination that uses a provisional leg.
Paper ticket outcomes are collected in the separate paper-performance ledger;
they do not directly count as calibration samples or feed the profile fit.
Calibration continues to use prequential, past-only forecast evidence.

The ranking stage selects the highest-calibrated-probability candidate per
game from the chosen strict or paper-only surface. On paper tickets, the
adaptive individual-pick ROI-band gate is applied before this selection;
bookmaker price never breaks a probability tie. Deterministic fixture, market,
source, and forecast identifiers are used instead. Same-game legs are
prohibited rather than silently multiplied as independent.
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
