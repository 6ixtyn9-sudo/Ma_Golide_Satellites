# Provider matrix and evidence status

Provider adapters normalize external payloads into canonical facts. Fixture
coverage proves parser behavior only. `DOCUMENTATION_ONLY` means an endpoint or
quota contract has not been authenticated and verified live in this cleanup.
No live-provider claim should be inferred from a passing offline suite.

## Schedules and official results

| Provider | League | Capabilities | Authentication | Current evidence |
| --- | --- | --- | --- | --- |
| ESPN | NBA | Scoreboard discovery, BPI predictions, summaries, recent finals, standings | Public endpoints | Committed live-captured fixtures plus offline regressions |
| EuroLeague | EuroLeague | Schedule discovery, game results, quarter/OT reconciliation, standings | Public endpoints | Fixture-based parser regressions; live availability must be checked at run time |

Only final, reconciled results enter grading. In-progress scores, inconsistent
quarter sums, invalid overtime totals, and irreconcilable team identities are
quarantined. Missing periods remain missing.

## Prediction providers

Priority is `espn`, `hardwood`, then `predixsport` by default. One coherent
provider bundle is selected; probabilities and expected scores from different
providers are never averaged.

| Provider | League | Prediction fields | Key / freshness |
| --- | --- | --- | --- |
| ESPN BPI | NBA | Home/away probabilities; expected scores, total, and margin when present | Public; per-prediction source modification time |
| Hardwood | NBA | One artifact's probabilities, expected scores, total, margin, model provenance | Public artifact; generated-at freshness and team/time/ID join |
| Predixsport | NBA and EuroLeague | Probabilities; detail endpoint can add expected total/margin and derived scores | `PREDIXSPORT_API_KEY`; disabled without a key; documentation-only endpoint schema |

A probabilities-only prediction remains partial: it may support moneyline, but
missing expected scores are not invented. Predixsport score derivation uses
`home = (total + margin) / 2` and `away = (total - margin) / 2`, with margin
defined as home minus away. Provider model disagreement is preserved rather
than “repaired” by forcing margin and probability favorite to agree.

## Bookmaker providers

All prices are normalized to decimal. The market, exact period, book, source
event, observation time, line, and pre-match status are retained. American
prices are converted where the source contract requires it. Missing lines or
prices remain missing.

| Provider | Implemented markets | Periods | League / key |
| --- | --- | --- | --- |
| TheOddsAPI | h2h, totals, spreads; first-half h2h includes draw prices | Full-game h2h/total/spread; Q1–Q4 total/spread; H1 h2h | NBA; `THE_ODDS_API_KEY` |
| SportsGameOdds | Totals only | Full game and Q1–Q4 | NBA; `SPORTSGAMEODDS_API_KEY` |
| TheRundown | Moneyline and totals | Full game only | NBA; `THERUNDOWN_API_KEY` |

Important restrictions:

- TheRundown has no code path that can emit quarter or half markets. Its
  requested spread reference data is deliberately not mapped to canonical
  bets.
- SportsGameOdds does not infer quarter coverage from an oddID pattern alone;
  only literal documented total IDs in a response are parsed.
- TheOddsAPI quarter/half markets come from the per-event endpoint. H1 is not
  Q1+Q2 inference and Q1 is never relabelled as H1.
- EuroLeague book coverage is not enabled by these adapters. An empty or
  unsupported book response is not converted into an invented market.
- No current book adapter supplies highest-quarter prices. Highest-quarter
  forecasts therefore remain research-only.
- Full-game spread observations may be stored from TheOddsAPI, but the
  replacement forecast contract currently predicts quarter spreads, not a
  full-game spread pick.

Book coverage and quota details are source-contract claims until an
authenticated smoke run verifies them. Secrets travel through the configured
HTTP authentication mechanism and are redacted from surfaced errors. Do not
print request URLs containing credentials.

## Identity and chronological selection

Book joins use a known provider event ID first. A guarded fallback requires
compatible home/away identities and scheduled time. A new provider event ID
cannot overwrite a conflicting existing mapping merely because names match.
Unknown IDs without sufficient identity evidence are not trusted.

Within a canonical slot, observations are ordered by parsed UTC timestamps,
not ISO string suffixes. Pregame selection requires freshness and strict
pre-tip observation time. Post-tip values are stored in separate evidence
slots and cannot replace frozen pre-tip facts.

## Smoke behavior

```bash
python -m ingestion.cli smoke --league NBA --date 2026-11-07
python -m ingestion.cli smoke --league EuroLeague --date 2026-11-07
```

Smoke is read-only; `--write` is rejected. Provider health distinguishes
missing credentials, stale evidence, provider errors, missing identity, and
successful coverage. A provider returning no rows is not automatically proof
that the league or market is unsupported.
