# Ma Golide Satellites

Ma Golide is a **headless, DuckDB/CSV.GZ-first sports evidence and forecasting
system**. It acquires schedules, official results, provider predictions, and
bookmaker observations; freezes pre-tip evidence; measures historical market
reliability; and emits current market and accumulator research reports.

There is no web server, spreadsheet store, external database, worksheet
registry, canonical JSON-file datastore, or external betting writer. JSON is used for
complete row payloads inside DuckDB and for derived local operator reports,
receipts, and audit state; DuckDB remains the canonical sports system of record.

## League and provider coverage

| League | Native regulation periods | Schedule / results | Independent predictions | Books |
| --- | --- | --- | --- | --- |
| NBA | Q1–Q4 | ESPN | ESPN BPI, Hardwood, Predixsport | ESPN scoreboard, TheRundown, SportsGameOdds, TheOddsAPI* |
| WNBA | Event-declared H1/H2 or Q1–Q4 | ESPN | ESPN BPI | ESPN scoreboard* |
| EuroLeague | Q1–Q4 | Euroleague Basketball | Predixsport | BetExplorer aggregate-consensus research board only; no validated named-book ticket lane |
| EuroLeague Women (`EUROLEAGUE_W`) | Q1–Q4 + official OT | FIBA GDAP schedule + FIBA game-page detail | No independently validated pre-tip provider yet | No validated named-book adapter; intentionally unpriced |
| CBA | Q1–Q4 | CBA public data service | Empirical team history after reconciled-history thresholds | BetExplorer aggregate research only; no validated named-book lane |
| NCAA_M | H1/H2 | ESPN men's Division-I endpoint | ESPN BPI | ESPN scoreboard, TheOddsAPI* |
| NCAA_W | Q1–Q4; event-declared legacy H1/H2 | ESPN women's Division-I endpoint | ESPN BPI | ESPN scoreboard, TheOddsAPI* |
| NBA G League (`NBA_G_LEAGUE`) | Q1–Q4 | ESPN `basketball/nba-development` scoreboard + summary | No external pre-tip provider validated; empirical-history forecasts only after reconciled-history thresholds | No validated named-book lane; intentionally unpriced |

* TheOddsAPI requires `THE_ODDS_API_KEY`; without it WNBA/NCAA books are ESPN-scoreboard only (ESPN carries no NCAA_W moneyline — live 153/153 ML rows unpriced on free tier).

The canonical registry also includes EuroCup, Basketball Champions League,
Basketball Africa League, Australia NBL, Spain ACB, France LNB Pro A, Germany
BBL, Italy LBA, Turkey BSL, Greece Basket League, ABA League, Israel Premier
League, Japan B.League, South Korea KBL, Brazil NBB, and Argentina Liga
Nacional. They have typed four-quarter contracts, independent local state, and
normal forecast/freeze/settlement handling for already-canonical imported games.
Until a compliant schedule/result adapter is validated for a lane, its automated
run emits an explicit unsupported-source `CoverageGap`; it does not invent a
fixture, final, probability, or price. NBA G League is the current exception:
its captured ESPN source contract enables date-bounded canonical fixtures and
reconciled results, but not ESPN BPI, deep ESPN-index history, or a named-book
execution lane. Australia NBL has a probationary The Odds API sport-key mapping
only: credentialed region/bookmaker coverage must be observed before it can be
treated as a validated named-book lane.

NCAA men's halves are never relabelled as quarters. Modern NCAA women's
results are four quarters, while an explicitly declared older two-half event
remains halves. Supported NCAAW H1/H2 **book markets** remain half markets:
grading derives H1/H2 from Q1+Q2/Q3+Q4, and does not reuse a full-game line.
Quarter/highest-quarter products remain disabled for both NCAA leagues. Early
WNBA two-half games likewise remain halves; current event-declared four-quarter
games remain quarters. Every ESPN-backed league—NBA, WNBA, NCAA_M, NCAA_W, and
NBA G League—requires a matching event-level ESPN league UID before discovery
or result reconciliation; missing or contradictory identity is quarantined.
CBA's public schedule/result source is kept separate from its empirical
prediction surface; aggregate book observations are never stored as model
predictions. EuroLeague Women is a separate FIBA competition with isolated
`EUROLEAGUE_W` IDs and `localdata/euroleague_w` state: it never reuses the
men's EuroLeague schedule, receipts, mappings, price lane, or model claims.
FIBA's canonical tip field is `gameDateTimeUTC`; visible game-page local/display
times are never used to derive it. CBA and NCAA live provider coverage remains
`DOCUMENTATION_ONLY` until authenticated smoke tests verify it. Offline tests
prove parsing, reconciliation, routing, and forecast behavior only.

## Canonical market coverage

| Canonical market | Periods / selections | Evidence rule |
| --- | --- | --- |
| Moneyline | Full game: home or away | Independent provider-probability and empirical-history forecasts; neither shares the other's calibration, and an exact priced bookmaker snapshot is required before acca eligibility |
| Favorite / underdog role | Provider full-game moneyline | A role on the provider forecast, derived from real home/away prices—not a separate wager model |
| Total | Full game: over or under | The posted bookmaker line is mandatory; a provider expected total is only a model baseline |
| Totals | Exact offered Q1–Q4 or H1/H2 periods | Exact native book period, exact posted line, and over/under prices |
| Spreads | Exact offered Q1–Q4 or H1/H2 periods | Exact native book period and home-oriented posted line with both side prices |
| 1X2 | First half: home, draw, or away | Exact first-half book, including a draw price when the draw is selected |
| Highest quarter | Q1, Q2, Q3, or Q4 for non-NCAA quarter leagues only | History-only research forecast; an exact priced highest-quarter book and dead-heat rules are required for acca use |

The forecasting model covers every market family in the former betting
contract without restoring its worksheet-era names or wrappers. Current book
adapters do not yet supply highest-quarter prices, so those forecasts remain
research-only until such evidence exists.

## Quick start

Python 3.11+ is required. Runtime dependencies are DuckDB and Requests; the
development extra adds Pytest and Ruff. `uv` can be installed separately for
locked environment synchronization.

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
ruff check .
```

Run a deterministic, credential-free rehearsal:

```bash
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-check/warehouse.duckdb \
  --localdata /tmp/magolide-check --notify
```

This writes derived operational reports under `/tmp/magolide-check/reports`
and Edge-style pick cards, paper-ticket state, and grading reports directly
under `/tmp/magolide-check`, but does not create a DuckDB file or CSV exports.
Its verification label is `FIXTURE_VERIFIED`; it proves neither live provider
coverage nor betting performance. See [Operator picks and paper-ticket
grading](docs/OPERATOR-REPORTS.md) for the files and settlement rules.

A live acquisition attempt remains a canonical dry run unless `--write` is
present:

```bash
python scripts/daily.py --auto-once --date 2026-11-07 --league NBA
python scripts/daily.py --auto-once --date 2026-11-07 --league WNBA
python scripts/daily.py --auto-once --date 2026-11-07 --league EUROLEAGUE_W
python scripts/daily.py --auto-once --date 2026-11-07 --league CBA
python scripts/daily.py --auto-once --date 2026-11-07 --league NCAA_W
```

Write-enabled runs also ingest a bounded deep-history batch from ESPN's
season/event indexes. Durable per-league cursors currently begin at verified
nonempty index years (NBA 1947, WNBA 1997, NCAA 2001), while preserving old
final-only scores and event-declared halves without claiming exhaustive
coverage. A versioned one-time NCAAW reparse repairs rows previously stored as
two halves plus fake overtime. See
[ESPN acquisition and history](docs/ESPN-ACQUISITION.md).

CBA schedule/result acquisition uses the public CBA data service; its aggregate
book observations remain research-only and cannot create a ticket. NCAA book
acquisition uses the permanently-free tier of `THE_ODDS_API_KEY`; a missing key
disables that provider rather than inventing book evidence. In GitHub Actions,
optional permanently-free provider credentials are supplied as Actions secrets
(`THE_ODDS_API_KEY`, `THERUNDOWN_API_KEY`, `SPORTSGAMEODDS_API_KEY`, and
`PREDIXSPORT_API_KEY`); without them the report explicitly marks the adapter as
disabled. See [Operations: optional odds-provider credentials](docs/OPERATIONS.md#optional-odds-provider-credentials)
for league/market coverage and safe rollout steps.

Direct CLI persistence requires both the CLI flag and environment opt-in:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/daily.py \
  --auto-once --date 2026-11-07 --write
```

The named external-trigger GitHub Actions workflow defaults
`ALLOW_MAGOLIDE_WRITES` to `true`; set that repository variable to `false` to
opt out of canonical writes. Manual workflow dispatch remains a dry run.

## Forecast and acca reports

When the external adaptive gate finds work due, the orchestrator runs pregame
acquisition, current market generation, immutable pre-tip freezing, and (unless
skipped) result reconciliation. Idle external ticks perform only a read-only
schedule/state check; unavailable state or schedule data fails open to the full
league run. Paper positions are not selected on the first usable scan: each
dated card is decided at **09:00 Africa/Johannesburg** on its named date, unless
the first fixture's 60-minute pre-tip safety cutoff is earlier. At that point,
either a strictly evidenced pre-tip ticket or an immutable no-bet decision is
recorded. A separate `auto_tickets_YYYY-MM-DD_preview.txt` can show current
named-price draft candidates before then, but it is explicitly non-execution
research: it creates no receipt, stake, grade, bank impact, or real-money
instruction. Generate a report directly from an existing warehouse with:

```bash
python scripts/build_markets.py \
  --league NBA --date 2026-11-07 --as-of 2026-11-07T18:00:00Z \
  --tz Africa/Johannesburg --legs 3 --limit 5 \
  --output /tmp/markets.json
```

For current fixtures, `--as-of` must be before scheduled tip. For historical
backtests, only reconciled finals available at each simulated cutoff are used.
The default result-availability delay is four hours. The report reads a
settled single-pick audit from `localdata/<league-lowercase>` (or the path passed
with `--localdata`) when it contains usable settled individual picks; this is
read-only. If the audit is absent, empty, or has no usable priced rows, it falls
back to past-only prequential settled individual forecasts and reports that
cold-start fallback explicitly.

### EuroLeague consensus-price paper league

EuroLeague also has a separate, deliberately non-execution **consensus-price
paper league**. It uses only the exact pre-tip representative average captured
by BetExplorer, stored as `aggregate_consensus` with `unresolved` source
permission and `probation` operational validation. It is **not** a bookmaker
quote and is never admitted to the strict named-book ticket lane.

At the same 09:00 SAST-or-earlier-safety cutoff,
`localdata/euroleague/euroleague_consensus_paper_YYYY-MM-DD_roster.json` can
time-lock up to three full-game-moneyline research selections. A row must retain
all ordinary model-quality gates, have either a full calibration or a disclosed
developing calibration sample of at least 10, and clear a 3%
model-versus-displayed-consensus threshold. The matching `.txt` card reports
the exact average price, calibrated probability, quality tier, strict-lane
blockers, and later official-result accuracy/Brier tracking.

This board is isolated from tickets: it creates **no stake, ticket, receipt,
paper-bank movement, ROI/P&L, or real-money instruction**. The time-locked
roster preserves its selection/evidence data unchanged; only text-card grades
refresh after official reconciled results.

### Point-in-time model-versus-market evaluation

Run the separate research-only moneyline comparison against the existing
read-only warehouse:

```bash
python scripts/market_eval.py --league NBA --book Pinnacle \
  --as-of 2026-11-17T00:00:00Z --output /tmp/market-eval.json
```

The lane requires a frozen pre-tip prediction, keeps each provider/model
version and named-book lane separate, accepts only reconciled official finals,
and uses a quote recorded no later than that forecast's decision timestamp. Missing,
stale, ambiguous, and post-decision evidence is reported as an explicit
exclusion. Paper return/Kelly fields are diagnostics only: observed prices do
not guarantee execution and the report makes no profitability claim. Men's EuroLeague
can be selected as `EUROLEAGUE` (or `EuroLeague`); FIBA EuroLeague Women is the
separate `EUROLEAGUE_W` value.

The walk-forward harness fits separate reliability profiles by league, market,
period, model source, and calibration-version fingerprint. Each profile uses
only earlier graded outcomes. Its shrunken observed-minus-claimed adjustment
changes later calibrated probabilities; strict ACCA eligibility requires at
least 30 graded non-push samples in the exact segment by default. The market
JSON's `walk_forward.replay_diagnostics` shows replay counts by stage and exact
segment (including result-delay, calibration-epoch, and team-history exclusions)
when a profile has no samples. These diagnostics do not change eligibility,
minimum-sample, paper-ticket, or ROI policies; the settled-pick audit remains
separate from calibration evidence.

Strict ACCA eligibility requires every leg to have:

- an exact real bookmaker market, period, line where applicable, and selected
  decimal price;
- visible pre-tip book evidence and provider evidence where used, with freeze
  state reported explicitly;
- a calibrated probability at least 0.50;
- for history-based forecasts, at least eight effective prior games per team by
  default—the same weight as the league prior—so fallback league averages cannot
  supply ACCA legs;
- acceptable push risk for totals and spreads.

Accas contain exactly the requested number of legs (two through six) and are
ranked by the product of calibrated leg probabilities. Bookmaker prices are
reported as payout evidence but never break probability ties. Same-game legs are
prohibited. Cross-game conditional independence is an explicit simplifying
assumption, not a measured guarantee. “Highest” therefore means highest under
that model among eligible combinations—not a guaranteed winner or an
empirically certified global maximum.

The strict `Forecast.eligibility` gate is unchanged. For compatibility,
`accas.results` (also exposed as `accas.strict_results`) remains the strict
candidate list. A separate `accas.paper_results` surface is used only by the
auto-ticket publisher and may include a provisional segment when at least one
graded non-push sample exists and insufficient sample size is the segment's
only blocker. Exact priced-market, team-history, push-risk, and probability-floor
checks still apply. Such legs remain marked ineligible in the forecast
and the frozen paper receipt labels the sample-size exception. Before paper
tickets are assembled, each unique eligible wager selection must match a
positive historical individual-pick performance band. Every passing selection
is also published as a single in `accas.paper_single_picks` and flagged in the
picks JSON/text. The rolling `picks_audit_rolling.json` settled ledger is used
when it contains usable settled picks; otherwise the service uses its own
past-only prequential settled forecasts and marks the cold-start fallback in
`paper_pick_filter.history_fallback`. Only settled single picks with calibrated
probabilities and exact pre-tip prices count. Bands are separated by league,
market, period, model source, calibrated-probability range, and individual-odds
range (`<50%`, 50–60%, 60–70%, 70–80%, 80–90%, 90%+; odds below 1.20,
1.20–1.50, 1.50–2.00,
2.00+).

Band ROI uses flat one-unit stakes (win: decimal price minus one; loss: minus
one; push: zero), a 14-day recency half-life, and shrinkage toward model,
market, then league-level ROI. A band can qualify after one settled pick—there
is no long fixed sample-size wait—but the band itself must have positive
recency-weighted ROI and positive shrinkage-adjusted ROI. An unseen band cannot
inherit eligibility from a profitable parent band. The policy is recalculated
from settled individual picks each scan, so newer wins and losses can change
future paper eligibility. Acca outcomes are kept separate and never enter this
single-pick ROI calculation. A positive historical band is an exploratory
paper-only signal, not statistical validation or proof of future profit.

These are per-pick performance conditions, not a combined-acca odds or
expected-value gate. The requested leg count is an upper bound for combinations:
when fewer distinct ROI-qualified games are available, the paper combo shrinks to
that count if at least two games qualify. When exactly one ROI-qualified
selection exists, the operator freezes it as a one-unit paper single instead of
forcing an acca. This does not issue separate singles for a pool containing
multiple selections, even if they all come from one game. Thus a default
three-leg request can produce a triple, a double, or a single position without
relaxing the pick checks. Singles and accas are both graded and counted as paper
positions; their results remain separate from calibration samples, which
continue to use prequential forecast evidence. Multi-leg combinations are ranked
by joint probability and do not reuse the same wager selection across separate
acca tickets. Different selections from one game can appear on separate tickets,
but no acca includes two legs from one game. Existing frozen receipts are not
rewritten by later policy changes.

Reports have schema `magolide.market_forecasts` and status
`decision_support_only`. They cannot place bets or trigger external betting
writes.

## System of record and read access

- `localdata/warehouse.duckdb`: canonical games, prediction slots, book
  snapshots, aliases, provider health, and operational run records.
- `localdata/csv/*.csv.gz`: six materialized canonical tables for audit and
  recovery.
- `localdata/ledgers/`: immutable, content-addressed stage snapshots.
- `localdata/reports/`: derived operational and market JSON reports.
- `localdata/<league>/picks_YYYY-MM-DD.txt`: Git-published compact dated
  human card. Individual fixture draft evidence stops changing at its own
  60-minute pre-tip cutoff.
- `localdata/<league>/picks_YYYY-MM-DD.json` and
  `picks_YYYY-MM-DD_lifecycle.json`: structured mutable draft/lifecycle views
  retained only in the workflow cache and bounded Actions artifacts, never
  Git-published.
- `localdata/<league>/pick_freeze_YYYY-MM-DD_<fixture-hash>.json`: immutable
  per-fixture pick snapshots, including selections, prices/lines, evidence
  timestamps, scheduled tip, and freeze timestamp.
- `localdata/<league>/picks_today.json` and `picks_audit_*`: volatile scan and
  expanded audit detail retained in the workflow cache and 14-day artifacts.
- `localdata/<league>/auto_tickets_*`: immutable paper receipts/no-bet
  decisions, compact durable locks, and compact grading/performance summaries
  are allowlisted for Git publication; see `docs/OPERATOR-REPORTS.md`.
- `localdata/notification_ledger.json`: bounded notification deduplication
  metadata, not canonical sports evidence.

After each named external-trigger GitHub run, one serialized publisher commits
paper-ticket decisions, durable locks, settlement summaries, dated text cards,
and immutable per-fixture pick freezes to `main`; it never overwrites
conflicting immutable evidence. Structured dated pick JSON, detailed lifecycle
renders, `picks_today.json`, expanded audits, canonical DuckDB/CSV, acquisition
ledgers, and caches remain outside Git.

```bash
python -m ingestion.cli coverage --league NBA
python scripts/mine_assay.py --league NBA
python scripts/build_warehouse.py

python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/warehouse.duckdb')
con.sql('SELECT * FROM latest_pregame_predictions LIMIT 20').show()
con.close()
PY
```

Recover into a **new** path; recovery never overwrites an existing warehouse:

```bash
ALLOW_MAGOLIDE_WRITES=true python scripts/recover_warehouse.py \
  --warehouse /tmp/recovered.duckdb --csv-dir localdata/csv
```

## Safety boundaries

- Prediction and book evidence are separate; neither is synthesized from the
  other.
- Pre-tip means strictly `observed_utc < scheduled_utc`. Historical training
  requires immutable frozen rows finalized at/after tip; current reports use
  the latest visible pre-tip evidence.
- Dry runs never create or modify canonical DuckDB/CSV state. Their derived
  operator reports may still be committed to Git by the explicit publisher.
- Fixture persistence needs `ALLOW_FIXTURE_WRITES=true`, the normal write
  opt-in, and paths isolated from canonical state.
- Notifications are optional operational heartbeats, not betting advice.
  Delivery has a separate `ALLOW_NOTIFICATION_WRITES=true` gate and fixture
  reports can never send.
- Live-only tests are deselected by default. Offline tests make no network
  calls and no external writes.

## Documentation

- [Architecture](docs/ARCHITECTURE.md)
- [Data and forecasting contract](docs/CONTRACT.md)
- [Operations runbook](docs/OPERATIONS.md)
- [Provider matrix](docs/PROVIDERS.md)
- [ESPN acquisition and history](docs/ESPN-ACQUISITION.md)
- [Research methodology](docs/RESEARCH.md)
- [Verification and limitations](docs/VERIFICATION.md)
- [Contributor checks](docs/CONTRIBUTING.md)
