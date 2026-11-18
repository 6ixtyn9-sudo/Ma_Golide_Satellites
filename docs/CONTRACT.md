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

Full-game moneyline forecasts are emitted independently from every valid
provider prediction visible at the forecast cutoff, plus the empirical history
model's full-game margin distribution. Provider priority controls stable output
ordering only; one provider no longer suppresses other valid providers. Each
provider/model retains its own source identity and calibration profile; the
service does not average provider probabilities. Providers with only score,
total, or margin expectations can contribute to the corresponding hybrid line
forecasts, but not to moneyline forecasts without a coherent home/away pair.
The empirical surface requires no provider prediction and can be reconstructed
at each historical tip for walk-forward calibration.

Sources and calibration profiles remain separate. No surface inherits another's
samples or probabilities, and exact ties within a surface produce no selection
from that surface. Full-game selections are two-way (`home` or `away`); an
empirical simulated draw is not mapped to a two-way book selection.

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

`HistoryModel` builds baselines from every reconciled final strictly before
the target and available under the result delay. There is no default league-wide
200-game training cap; a finite `training_window` is available only when an
operator explicitly configures one. Prefix indexes make league and team
lookups efficient across the full chronological history. Team rates are
recency weighted with a half-life and shrunk toward the league mean. Margin and
total residual samples use the most recent configured residual scenarios, but
each residual's own baseline can use all earlier available games.

The distribution is empirical. Probabilities are counts over residual
scenarios, not claims from a fitted sportsbook or a certified probability
model. Exact probability ties are not broken arbitrarily. Totals and spreads
report push probability separately.

### Availability-adjusted empirical surface (context_availability)

ESPN injury and prior-minutes evidence drives one additional independent
surface on ESPN-covered leagues. For each team, the window-average scoring
contribution of every player is estimated from prior ESPN boxscores as
appearance rate times per-appearance points. Players a point-in-time injury
snapshot reports as Out / Out For Season / Inactive are confirmed out, so
their window-average contribution is the estimated scoring loss; expected
points are scaled by ``1 - lost_share``, bounded by
``ContextAdjustmentConfig.max_scale``. The scale applies only to full-game
moneyline, total, and spread forecasts reconstructed from the empirical
history surface, under the source
``context_availability+team_history_empirical`` with its own calibration
version. It is never blended into the base or provider surfaces, and provider
expectations — which may already internalize availability — are never scaled
again.

The surface emits forecasts only when a point-in-time injury snapshot exists,
both teams carry sufficient window evidence (asymmetric evidence is rejected,
never half-applied), and at least one scale is below 1.0; a zero adjustment
emits nothing rather than duplicating the base surface. Missing injury entries
remain unknown and are never read as healthy. The full candidate record stays
published in the report for audit. Recent form/ATS context, BPI auxiliary
metrics, standings, and book movement remain non-inputs: the first two would
double-count information already reflected in reconciled history and posted
lines, and the last two have no pre-tip snapshot semantics.

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

`walk_forward.replay_diagnostics` reports the replay funnel separately from the
settled-pick audit: eligible final inventory, generated and ungradeable replay
forecasts, forecasts withheld by the result-availability delay, and calibration
population counts after the epoch and team-history filters. Its segment rows
use the exact `(league, market, period, source, calibration_version)` key and
show the current forecast count, replayed/available counts, epoch exclusions,
team-history diagnostic-pool samples, pushes, resulting `profile_n`, and a
`zero_n_reason` when the exact segment has no non-push sample. Pushes are shown
but never counted in `n`; team-history fallback samples are diagnostic-only.
These fields explain profile construction only and do not relax the sample
minimum, exact-price, eligibility, or ROI policies. The rolling settled-pick
audit remains a separate input to the individual-pick performance gate, not a
source of calibration samples.

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
- every history-based forecast has at least as much effective prior evidence
  for each team as the configured league-prior strength (eight effective games
  by default), so a league-average fallback cannot masquerade as team evidence;
  a provider-only prediction does not depend on internal team-history data and
  is not subject to this guard; history-plus-provider hybrids are subject to it;
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

The historical source is the rolling audit's `settled_ledger` when it contains
at least one usable settled individual pick; otherwise the service falls back
to its past-only prequential settled individual forecasts. The fallback also
covers a valid-but-empty or unusable audit after a cold-cache start and is
explicitly identified in `paper_pick_filter.history_fallback`. It does not
relax current forecast eligibility or calibration requirements. Pending rows,
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

Every passing individual pick is exposed in `accas.paper_single_picks` and
flagged as a `paper_roi_band_eligible` single in the operator pick JSON/text, so
a qualifying pick remains available even when there are not enough games for a
multi-leg ticket. This performance gate does not relax or add a calibration-
sample gate: the existing sample-only provisional path and all other strict
eligibility safeguards remain in force.

For combinations, the requested leg count is an upper bound. The effective
paper leg count is the smaller of that request and the number of distinct games
with ROI-qualified picks. A combination is emitted only when at least two such
games exist. For the default three-leg request, two qualifying games produce a
double. When exactly one ROI-qualified selection exists, the operator instead
freezes that pick as a one-unit paper single; it does not force an acca. This
single-ticket rule does not turn a pool of multiple selections into separate
singleton bets, even if they all come from one game. The paper combination
surface greedily accepts tickets in descending joint-probability order; it does
not reuse an exact wager selection across separate acca tickets. A different
selection from the same game may be used on another ticket, while same-game legs
remain prohibited
within any one acca. This is a no-repeated-leg rule, not mathematical outcome
exclusivity: separate tickets may still win simultaneously. The strict
`results`/`strict_results` analytical ranking is unchanged and can contain
overlapping combinations. Fewer paper combinations are returned when the
requested limit exceeds the available non-repeated leg sets.

Frozen paper receipts label every combination or single that uses the existing
provisional sample-size exception. Singles and combinations are graded as paper
positions in the separate paper-performance ledger; those outcomes do not
become calibration samples or feed the profile fit. Calibration continues to
use prequential, past-only forecast evidence.

The ranking stage selects the highest-calibrated-probability candidate per
game from the chosen strict or paper-only surface. On paper tickets, the
adaptive individual-pick ROI-band gate is applied before this selection;
bookmaker price never breaks a probability tie. Deterministic fixture, market,
source, and forecast identifiers are used instead. Same-game legs are
prohibited rather than silently multiplied as independent.
For the effective paper leg count (or the unchanged requested count on the strict
analytical surface), a best-first search ranks combinations by:

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
