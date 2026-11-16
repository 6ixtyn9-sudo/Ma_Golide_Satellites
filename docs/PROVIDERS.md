# Provider matrix and evidence status

Provider adapters normalize external payloads into canonical facts. Fixture
coverage proves parser behavior only. `DOCUMENTATION_ONLY` means an endpoint or
quota contract has not been authenticated and verified live in this cleanup.
No live-provider claim should be inferred from a passing offline suite.

## Schedules and official results

| Provider | League | Capabilities | Authentication | Current evidence |
| --- | --- | --- | --- | --- |
| ESPN | NBA | Scoreboard discovery/books, BPI predictions, summaries, recent and indexed deep history | Public endpoints | Live read-only probes plus committed fixtures and offline regressions |
| ESPN | WNBA | Scoreboard discovery/books, BPI predictions, summaries, event-native halves/quarters, indexed deep history | Public endpoints | Live read-only probes plus hermetic end-to-end league regressions |
| ESPN | NCAA_M | Division-I men's scoreboard/books, BPI predictions, summaries, two-half/OT results, recent and indexed deep history | Public endpoints | Historical probes and hermetic endpoint/two-half/OT regressions; scheduled smoke still required |
| ESPN | NCAA_W | Division-I women's scoreboard/books, BPI predictions, summaries, modern quarters and declared legacy halves, recent and indexed deep history | Public endpoints | Live four-quarter final probe plus hermetic quarter/legacy-half and one-time reparse regressions |
| EuroLeague | EuroLeague | Schedule discovery, game results, quarter/OT reconciliation, standings | Public endpoints | Schedule/boxscore live captures 2026-11-06; season schedule scores/partials, empty-season trap, and v3 per-round standings live-verified 2026-11-16; upstream reuse terms not reviewed |
| CBA | CBA | Official league data API: season/day discovery, Beijing-local UTC times, quarter/OT reconciliation via observed period-score POST endpoint | Public endpoints; AES-128-ECB bodies decrypted with the key constant shipped in the site's own bundle | Live-verified 2026-11-16 from a datacenter IP (decryption, statuses, quarter sums, OT shape); unofficial/undocumented API and terms not reviewed — see caveats |
| iSportsAPI | CBA | Exact Chinese league discovery, UTC schedule/results, four-quarter/OT reconciliation | `ISPORTS_API_KEY` | `DOCUMENTATION_ONLY`; fixture-tested adapter, paid access and live league ID not verified |

Only final, reconciled results enter grading. In-progress scores, inconsistent
native-period sums, invalid overtime totals, and irreconcilable team identities
are quarantined. Missing periods remain missing. NBA, EuroLeague, CBA, and
modern NCAA women use Q1–Q4; NCAA men retain H1/H2. NCAAW and WNBA events that
explicitly declare two regulation periods retain H1/H2. NCAAW half **markets**
are graded from Q1+Q2/Q3+Q4 without relabelling the official quarters.
Historical finals without usable linescores
are stored as final-score-only evidence, not period-reconciled results. See
[ESPN acquisition and history](ESPN-ACQUISITION.md) for the bounded cursor and
verified index floors.

### EuroLeague acquisition details (verified 2026-11-16)

- Season history is acquired one whole season per request from the official
  schedule (`api-live.euroleague.net/v2/competitions/E/seasons/{code}/games`):
  every played row already carries the official score and per-quarter
  partials, so a full season of reconcilable finals costs one request and
  there is no league-wide game cap. Bounded batches walk seasons backwards
  under a durable `euroleague-history` cursor. The current-season parser is
  fixture-tested; older-season completeness should be confirmed from the
  per-season backfill reports rather than inferred from the schedule endpoint's
  existence.
- Two silent traps are guarded: a bare-year season code (`2026` not `E2026`)
  answers HTTP 200 with an **empty** data list (refused outright), and
  `winner` is `null` for current-season rows (winners are derived from
  reconciled scores, never from the field).
- Standings use the v3 per-round `basicstandings` endpoint (the v2
  `standings` path answers 404; verified live). It answers only completed
  rounds; the latest completed round is derived from the schedule when not
  given.
- Boxscore results (`live.euroleague.net/api/Boxscore`) remain the deeper
  per-game evidence path for slate results.

### CBA official data API — licensing and terms caveats

`data-server.cbaleague.com` is the data site's own undocumented API. It is
keyless and publicly served to every visitor's browser; response bodies are
AES-128-ECB encrypted with a 16-byte constant shipped inside the site's
front-end bundle (function `TIe`), which the vendored stdlib decryptor in
`ingestion/providers/_cba_aes.py` removes (copied from SportsDataverse
`sportsdataverse-py`, MIT; code license is not a data license). No
credentials, proxies, or access controls are used. This does not establish
permission to automate, store, or redistribute the returned data: the source's
terms have not been reviewed, and the scheduled workflow can persist normalized
facts. Keep data rights unresolved until the CBA site's terms or written
permission are checked. SportsDataverse has also not ruled on the upstream
terms. The wrapper's documented "datacenter-IP TCP block" did **not** reproduce
from this datacenter IP on 2026-11-16 (HTTP 200 in under a second), but it may
be geographic or conditional: an undecryptable or refused body is surfaced as
a provider error, never banked as data.

Chinese team display names are season sponsor names (e.g. `长白山恩都里` is
the Jilin club in 2026-27). Known variants map through the current alias table;
unknown names remain unchanged and will not safely join across providers. Do
not assume they are canonical franchise identities for historical modeling;
add a verified alias/team-ID mapping before using those rows as joined history.
The captured season list spans 1995–2026, but that list alone does not prove
daily schedule/results coverage for every season. Checked-in day/result
fixtures cover 2026 only; depth of successful historical CBA backfill remains
unverified until older dates are actually acquired and reconciled.

## Prediction providers

Priority is `espn`, `hardwood`, `predixsport`, then `sportradar` by default.
One coherent provider bundle is selected; probabilities and expected scores
from different providers are never averaged. Bookmaker prices are never
converted into prediction records.

| Provider | League | Prediction fields | Key / freshness |
| --- | --- | --- | --- |
| ESPN BPI | NBA, WNBA, NCAA_M, NCAA_W | Home/away probabilities; expected scores, total, and margin when present | Public; per-prediction source modification time; adapter version `adapter:espn-bpi-v1` |
| Hardwood | NBA | One artifact's probabilities, expected scores, total, margin, model provenance | Public artifact; artifact-declared model version, generated-at freshness, and team/time/ID join |
| Predixsport | NBA and EuroLeague | Probabilities; detail endpoint can add expected total/margin and derived scores | `PREDIXSPORT_API_KEY`; disabled without a key; adapter version `adapter:predixsport-v1`; documentation-only endpoint schema |
| Sportradar Global Basketball | CBA | Independent pre-match 2-way home/away probabilities; no expected scores or totals | `SPORTRADAR_API_KEY`, optional `SPORTRADAR_ACCESS_LEVEL=trial|production`; declared adapter model tag `global-basketball-v2`; `DOCUMENTATION_ONLY` |

Sportradar discovery requires one unique men's `CBA` competition in category
`CHN`, then one non-disabled season containing the target date. Its documented
Season Probabilities feed is paginated. An event joins an acquired CBA fixture
only through a unique one-to-one match on both canonical team identities and
the exact confirmed UTC start time. Missing, duplicate, unconfirmed, live, or
removed probability evidence remains a gap; it is never replaced with implied
book probabilities.

ESPN BPI and Predixsport do not expose a stable upstream model release in the
fields this adapter consumes. Their `adapter:` model versions fence known
adapter/provider changes and must be bumped when those changes are observed;
they do not automatically detect a silent upstream model replacement.
`produced_by` is retained as prediction provenance but arbitrary payload labels
are excluded from calibration identity because they can vary row by row. The
fingerprint uses only a closed Predixsport mapping for the recognized
`earlyseason` and `mature` model phases. If a provider's model changes, use a
stable `model_version` or bump its adapter-owned version; do not use free-text
producer labels to split samples.

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
| ESPN scoreboard | Moneyline, totals, spreads when explicitly offered | Full game | NBA, WNBA, NCAA_M, NCAA_W; public, independently timestamped canonical snapshots |
| BetExplorer | Representative average two-way moneyline prices | Full game | EuroLeague and CBA; public HTML pages; observed live 2026-11-16, but source terms/reuse rights are unreviewed; pre-tip rows only, no historical archive or per-book prices in this adapter |
| TheOddsAPI | h2h, totals, spreads; half markets use native H1/H2 | NBA: full game, Q1–Q4, H1 h2h; NCAA_M/NCAA_W: full game and H1/H2 markets | NBA, NCAA_M, NCAA_W; `THE_ODDS_API_KEY` |
| iSportsAPI | Moneyline, totals, spreads | Full game, H1 spread/total, Q1–Q4 spread/total | CBA; `ISPORTS_API_KEY`; `DOCUMENTATION_ONLY` |
| SportsGameOdds | Totals only | Full game and Q1–Q4 | NBA; `SPORTSGAMEODDS_API_KEY` |
| TheRundown | Moneyline and totals | Full game only | NBA; `THERUNDOWN_API_KEY` |

Important restrictions:

- ESPN scoreboard books are stored even when no prediction exists. Opening
  values without their own observation timestamps remain context and cannot be
  used as closing-line-value evidence; betting links are discarded.
- TheRundown has no code path that can emit quarter or half markets. Its
  requested spread reference data is deliberately not mapped to canonical
  bets.
- SportsGameOdds does not infer quarter coverage from an oddID pattern alone;
  only literal documented total IDs in a response are parsed.
- TheOddsAPI quarter/half markets come from the per-event endpoint. H1 is not
  a Q1 book and Q1 is never relabelled as H1. NCAA requests use the configured
  H1/H2 book markets and do not request quarter markets. For NCAAW, official
  quarter results derive the corresponding H1/H2 scores for settlement; a
  full-game line is never substituted for either half.
- iSportsAPI prices are Hong Kong odds and are converted to decimal by adding
  one. Spread, total, and European-moneyline bookmaker IDs remain in their
  separate documented namespaces. In-play fields never become pre-match facts.
- BetExplorer's fixture table exposes only a representative average moneyline
  pair in the observed HTML. This adapter skips rows with missing quotes or a
  tip time already passed; it does not retain post-tip, opening/closing, or
  historical prices, nor per-bookmaker, spread, or total markets. CBA future
  fixtures had no prices in the 2026-11-16 capture, so CBA quote availability
  has not been demonstrated there.
- BetExplorer's and the CBA data site's reuse/automation terms have not been
  reviewed. Public/keyless access does not establish a right to store or
  redistribute data. Do not treat either source as cleared for unattended
  production reuse until terms are checked; the `LIVE_VERIFIED` labels describe
  endpoint/parser observations only, not rights clearance.
- No current book adapter supplies highest-quarter prices. Highest-quarter
  forecasts therefore remain research-only.
- Full-game spread observations may be stored from TheOddsAPI, but the
  replacement forecast contract currently predicts quarter spreads, not a
  full-game spread pick.

Book coverage and quota details are source-contract claims until an
authenticated smoke run verifies them. Secrets travel through the configured
HTTP authentication mechanism and are redacted from surfaced errors. Do not
print request URLs containing credentials.

## Free player gamelog (documented, not yet featurized)

`site.web.api.espn.com/apis/common/v3/sports/basketball/{nba|mens-college-basketball}/athletes/{id}/gamelog` is verified 200 for both leagues (2026-11-09 probes: NBA + NCAA_M). It returns per-event `MIN/FG/3PT/FT/REB/AST/BLK/STL/PF/TO/PTS` via `names/labels` aligned to `seasonTypes[].categories[].events[]` stat vectors and an `events` map for date/opponent. Past games only → pre-tip safe by construction. Team `roster` endpoint supplies athlete ids. Bounded use (priced/watchlist games, top rotation) is required to stay within `context_request_budget` / `http_request_budget`; walk-forward validation is required before it can affect probabilities.

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
python -m ingestion.cli smoke --league WNBA --date 2026-11-07
python -m ingestion.cli smoke --league EuroLeague --date 2026-11-07
python -m ingestion.cli smoke --league CBA --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_M --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_W --date 2026-11-07
```

Smoke is read-only; `--write` is rejected. Provider health distinguishes
missing credentials, stale evidence, provider errors, missing identity, and
successful coverage. A provider returning no rows is not automatically proof
that the league or market is unsupported.
