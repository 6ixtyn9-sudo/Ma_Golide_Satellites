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
| NBA | Q1–Q4 | ESPN | ESPN BPI, Hardwood, Predixsport | ESPN scoreboard (relay only), The Odds API*, BetExplorer direct fallback† |
| WNBA | Event-declared H1/H2 or Q1–Q4 | ESPN | ESPN BPI | ESPN scoreboard (relay only), The Odds API*, BetExplorer direct fallback† |
| EuroLeague | Q1–Q4 | Euroleague Basketball | Predixsport | BetExplorer aggregate research lane plus direct named-book fallback† |
| EuroLeague Women (`EUROLEAGUE_W`) | Q1–Q4 + official OT | FIBA GDAP schedule + FIBA game-page detail | No independently validated pre-tip provider yet | BetExplorer direct named-book fallback† |
| CBA | Q1–Q4 | CBA public data service | Empirical team history after reconciled-history thresholds | BetExplorer aggregate research lane plus direct named-book fallback† |
| ABA League (`ABA_LEAGUE`) | Q1–Q4 + official OT | Externally referenced aba-liga.com calendar + reconciled match detail; no Ma Golide first-party capture | Empirical team history only after reconciled-history thresholds | BetExplorer aggregate research lane plus direct named-book fallback† |
| NCAA_M | H1/H2 | ESPN men's Division-I endpoint | ESPN BPI | ESPN scoreboard (relay only), The Odds API*, BetExplorer direct fallback† |
| NCAA_W | Q1–Q4; event-declared legacy H1/H2 | ESPN women's Division-I endpoint | ESPN BPI | ESPN scoreboard (relay only), The Odds API*, BetExplorer direct fallback† |
| NBA G League (`NBA_G_LEAGUE`) | Q1–Q4 | ESPN `basketball/nba-development` scoreboard + summary; normal CLI uses a 56-day date-bounded history bootstrap | No external pre-tip provider validated; empirical-history forecasts only after reconciled-history thresholds | BetExplorer direct named-book fallback† |

* The Odds API remains the scheduled, permit-gated direct API lane for NBA, WNBA, NCAA_M, and NCAA_W. A scheduled run receives `THE_ODDS_API_KEY` only after a durable three-credit permit; it requests date-bounded full-game `h2h,totals,spreads` from the named DraftKings group. ESPN observations remain relayed and unresolved/unverified audit evidence.

† The BetExplorer fallback is an operator-authorised direct public-page capture for this single-user project, not an API integration. It reads a dated basketball board and then only matching full-game comparison responses: H/A (`ha`) for the two-way moneyline, O/U (`ou`) for explicit total lines and Over/Under prices, and Asian Handicap (`ah`) for explicit home-oriented spread lines and home/away prices. It retains the response-supplied bookmaker name, source-line selection, pre-tip observation time, source URL/digest receipt, and strict ordered-team identity. It has a three-second request interval, hard cap, and run-scoped 403/429 stop; it uses no browser automation, relay, proxy, login, or aggregate quote. Its direct-page permission is deliberately labelled `unresolved` / `probation`, while the displayed bookmaker row remains a named-book price for the paper lane. It does not supply halves, quarters, or highest-quarter prices.

The current The Odds API routine priority is a shared-provider quota policy, not
an exclusion of other leagues from future compliant adoption. The
[league readiness and source-adoption roadmap](docs/LEAGUE-ADOPTION-ROADMAP.md)
defines the evidence-first path from a league contract to visible unpriced
normal-lane output, bounded price validation, and only then a prospective
paper-ticket path.

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
execution lane. To make the approved empirical lane useful after a cold start,
the normal CLI defaults it to a bounded 56-day scoreboard-history bootstrap;
this is still only date-bounded ESPN acquisition, not a deep-history claim or a
model/calibration override. The report's `empirical_lane_readiness` object shows
whether reconciled history, team evidence, and prequential calibration have
actually reached their existing gates. Australia NBL has a probationary The Odds
API sport-key mapping only: credentialed region/bookmaker coverage must be
observed before it can be treated as a validated named-book lane.

NCAA men's halves are never relabelled as quarters. Modern NCAA women's
results are four quarters, while an explicitly declared older two-half event
remains halves. The routine The Odds API path deliberately captures full-game
markets only: native H1/H2 and quarter payload parsing is retained for supplied
payloads, but scheduled collection would require charged per-event calls and is
not authorised. Settlement still derives H1/H2 from Q1+Q2/Q3+Q4 where the
canonical result declares quarters; a full-game line is never substituted for
either half. Quarter/highest-quarter products remain disabled for both NCAA
leagues. Early WNBA two-half games likewise remain halves; current
event-declared four-quarter games remain quarters. Every ESPN-backed league—NBA, WNBA, NCAA_M, NCAA_W, and
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

Write-enabled NBA, WNBA, NCAA_M, and NCAA_W runs also ingest a bounded
deep-history batch from ESPN's season/event indexes. Durable per-league cursors
currently begin at verified nonempty index years (NBA 1947, WNBA 1997, NCAA
2001), while preserving old final-only scores and event-declared halves without
claiming exhaustive coverage. NBA G League is deliberately excluded from that
index route; its bounded 56-day normal-run bootstrap uses only the separately
validated date scoreboard/results path. A versioned one-time NCAAW reparse
repairs rows previously stored as two halves plus fake overtime. See
[ESPN acquisition and history](docs/ESPN-ACQUISITION.md).

CBA schedule/result acquisition uses the public CBA data service; its aggregate
book observations remain research-only and cannot create a ticket. NCAA book
acquisition uses the permanently-free tier of `THE_ODDS_API_KEY` only through
the scheduled reservation guard: a matching permit is committed before the key
reaches one runner, while manual rehearsal and unpermitted work make no
named-book request. In GitHub Actions, the permanently-free provider secrets
are `THE_ODDS_API_KEY` and `PREDIXSPORT_API_KEY`; missing credentials or permits
are recorded as disabled/deferred source states rather than invented book
evidence. See [Operations: optional odds-provider credentials](docs/OPERATIONS.md#optional-odds-provider-credentials)
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
All dated cards are decided at **12:00 Africa/Johannesburg** on their named
date. An earlier first fixture's 60-minute pre-tip safety cutoff always takes
precedence. At that point, either a strictly evidenced pre-tip ticket or an
immutable no-bet decision is
recorded. Before then, the single `auto_tickets_YYYY-MM-DD.txt` card carries a
clearly labelled `DRAFT ONLY` / `PRE-CUTOFF` section showing current named-price
draft candidates, but that section is explicitly non-execution research: it
creates no receipt, stake, grade, bank impact, or real-money instruction.
Generate a report directly from an existing warehouse with:

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
- `localdata/<league>/picks_YYYY-MM-DD.txt`: Git-published readable dated
  human card. Individual fixture draft evidence stops changing at its own
  60-minute pre-tip cutoff.
- `localdata/<league>/picks_YYYY-MM-DD.json` and
  `picks_YYYY-MM-DD_lifecycle.json`: structured mutable draft/lifecycle views
  retained only in the workflow cache and bounded Actions artifacts, never
  Git-published.
- `localdata/<league>/pick_lifecycle_state.json`: Git-published lifecycle
  state with immutable per-fixture snapshots nested by date and fixture,
  including selections, prices/lines, evidence timestamps, scheduled tip, and
  freeze timestamp. No per-fixture JSON sidecar is produced.
- `localdata/<league>/picks_today.json` and `picks_audit_*`: volatile scan and
  expanded audit detail retained in the workflow cache and 14-day artifacts.
- `localdata/<league>/auto_tickets_*`: immutable paper receipts, compact
  durable locks (including inline no-bet decisions), and compact
  grading/performance summaries are allowlisted for Git publication; see
  `docs/OPERATOR-REPORTS.md`.
- `localdata/notification_ledger.json`: bounded notification deduplication
  metadata, not canonical sports evidence.

After each named external-trigger GitHub run, one serialized publisher commits
paper-ticket decisions, durable locks (including inline immutable per-fixture
pick snapshots), settlement summaries, and dated text cards to `main`; it never
overwrites conflicting immutable evidence. Structured dated pick JSON, detailed lifecycle
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
