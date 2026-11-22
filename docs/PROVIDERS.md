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
| ESPN | NCAA_M | Division-I men's scoreboard/books, BPI predictions, summaries, two-half/OT results, recent and indexed deep history | Public endpoints | Historical probes and hermetic endpoint/two-half/OT regressions; external-trigger smoke still required |
| ESPN | NCAA_W | Division-I women's scoreboard/books, BPI predictions, summaries, modern quarters and declared legacy halves, recent and indexed deep history | Public endpoints | Live four-quarter final probe plus hermetic quarter/legacy-half and one-time reparse regressions |
| ESPN | NBA G League | `basketball/nba-development` date-bounded scoreboard discovery, summaries, reconciled Q1–Q4 results, and bounded recent-result backfill | Public endpoints | Captured public scoreboard and summary on 2026-10-29 plus hermetic identity/status/period reconciliation regressions; BPI, deep core-index history, and named-book pricing are **not** validated |
| EuroLeague | EuroLeague | Schedule discovery, game results, quarter/OT reconciliation, standings | Public endpoints | Schedule/boxscore live captures 2026-11-06; season schedule scores/partials, empty-season trap, and v3 per-round standings live-verified 2026-11-16; upstream reuse terms not reviewed |
| FIBA GDAP + game page | EuroLeague Women (`EUROLEAGUE_W`) | Competition schedule, canonical UTC fixtures/status/results, official game-page play-by-play/period scores, player/team detail | Public APIM key shipped in FIBA's web client | Live facts and fixture-only cases are separated below; independent namespace, no men's EuroLeague reuse. Automation/reuse rights and future route/status vocabulary remain unresolved. |
| CBA | CBA | Official league data API: season/day discovery, Beijing-local UTC times, quarter/OT reconciliation via observed period-score POST endpoint | Public endpoints; AES-128-ECB bodies decrypted with the key constant shipped in the site's own bundle | Live-verified 2026-11-16 from a datacenter IP (decryption, statuses, quarter sums, OT shape); unofficial/undocumented API and terms not reviewed — see caveats |

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

### FIBA EuroLeague Women acquisition details (observations checked 2026-11-18)

**Live-confirmed transport/schema facts — not an availability or rights
promise.**

- The canonical fixture/result index is the official GDAP route
  `GET https://digital-api.fiba.basketball/hapi/getgdapgamesbycompetitionid`
  with `gdapCompetitionId=209127`. The reviewed response held 110 rows for the
  2026–27 edition, with 54 future second-round/playoff slots carrying
  `teamA: null` and `teamB: null`. Those rows are now skipped as *unassigned
  bracket-slot coverage telemetry*, never mapped to a made-up canonical game.
  A team-bearing, date-relevant row remains strict: malformed identity or UTC
  time is a provider error rather than an empty slate.
- GDAP supplies `gameDateTimeUTC`, an IANA venue timezone, FIBA game/team IDs,
  official scores, and independent result/statistic/play-by-play status flags.
  Acquisition accepts only the UTC field; neither game-page display time nor
  venue-local text is converted into a canonical tip. The observed final
  signature was `gameResultStatusCode=VALID`, `liveGameStatus=999`, and two
  numeric scores. `liveGameStatus=1` was also observed on future games, so it
  is not used as an in-play shortcut.
- FIBA's browser application ships a public APIM subscription key. The provider
  and dependency-free gate share one contract, send it in the APIM header (not
  the URL), and retry once with a page-refreshed public value on failure. This
  is not a Ma Golide secret and no key value is placed in reports or durable
  state. Key rotation cadence, quotas, and terms remain unverified.
- A final enters the canonical store only after the official game-page
  play-by-play supplies complete Q1–Q4 (and, when present, OT) cumulative
  scores that reconcile to FIBA's final. The PBP object must explicitly bind
  `gameTechnicalName` (or `gameTechnicalNameOrGameId`) to the requested GDAP
  game ID; a shape-valid decoy PBP is rejected. Team IDs use the dedicated
  `fiba-elw:<competition>:<team>` namespace and game IDs use
  `fiba-elw-<competition>-<game>`; they cannot collide with men’s EuroLeague.
- The event-page slug is derived from the GDAP edition ending year (for example
  `2027` → `euroleague-women-26-27`) rather than a frozen source literal. A
  future event-route change still fails visibly; playoff/Final Four route
  continuity has not been live-confirmed.

**Hermetic regression coverage — parser behavior only.** The committed FIBA
fixtures cover a completed OT final, a future team-less bracket row, terminal
GDAP statuses, explicit PBP game binding/decoys, failed-detail cursor retention,
and a late GDAP score-correction re-fetch. They do **not** establish that FIBA
will publish cancelled/abandoned rows, multi-OT detail, future slugs, or score
corrections in the same shapes.

**Operational and market limits.** The FIBA API's automation, storage, and
redistribution rights, quotas, and key-rotation policy have not been reviewed.
`teamA` is used as FIBA's schedule-side A/home orientation within this adapter,
but neutral-site/all-round semantic confirmation remains outstanding; do not
interpret this as independently verified venue-home proof. No independent
pre-tip prediction provider or named-book market source has been validated for
ELW. Its paper/operator lane is explicitly labelled **unpriced** and never
makes price, value, or EV claims or inherits men’s EuroLeague pricing.

### CBA official data API — licensing and terms caveats

`data-server.cbaleague.com` is the data site's own undocumented API. It is
keyless and publicly served to every visitor's browser; response bodies are
AES-128-ECB encrypted with a 16-byte constant shipped inside the site's
front-end bundle (function `TIe`), which the vendored stdlib decryptor in
`ingestion/providers/_cba_aes.py` removes (copied from SportsDataverse
`sportsdataverse-py`, MIT; code license is not a data license). No
credentials, proxies, or access controls are used. This does not establish
permission to automate, store, or redistribute the returned data: the source's
terms have not been reviewed, and the named external-trigger workflow can
persist normalized facts. Keep data rights unresolved until the CBA site's
terms or written permission are checked. SportsDataverse has also not ruled on
the upstream terms. The wrapper's documented "datacenter-IP TCP block" did **not** reproduce
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

Priority is `espn`, `hardwood`, then `predixsport` by default.
One coherent provider bundle is selected; probabilities and expected scores
from different providers are never averaged. Bookmaker prices are never
converted into prediction records.

| Provider | League | Prediction fields | Key / freshness |
| --- | --- | --- | --- |
| ESPN BPI | NBA, WNBA, NCAA_M, NCAA_W | Home/away probabilities; expected scores, total, and margin when present | Public; per-prediction source modification time; adapter version `adapter:espn-bpi` |
| Hardwood | NBA | One artifact's probabilities, expected scores, total, margin, model provenance | Public artifact; artifact-declared model version, generated-at freshness, and team/time/ID join |
| Predixsport | NBA and EuroLeague | Win probabilities (`home_team`/`away_team` with `win_probability`); `predicted_point_diff`; detail endpoint can add `predicted_total_points` and derived scores | `PREDIXSPORT_API_KEY`; disabled without a key; adapter version `adapter:predixsport`; documentation-only schema; upcoming horizon is **1 day** (the provider publishes NBA predictions on game day, and documents multi-day windows as empty in steady state); upstream calls are spaced **18 hours per league** because the 200-request monthly allowance is shared — roughly 40 calls/month per league, 80 for NBA + EuroLeague. The spacing is recorded as quota telemetry in the `predixsport` `provider_health` row (never as an acquisition cursor), is written even for dry runs and failed requests, and a skipped window reports `skipped_quota_window` while stored pre-tip observations stay available to the forecast stage |
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

The published Predixsport examples carry no tip-off timestamp — only
`prediction_date` and a `match_id` of the form `HOME_YYYYMMDD_AWAY.csv`. A date
is not promoted into a tip-off, so such a row stays unmatched and is reported as
a `predixsport_prediction_identity` coverage gap rather than being joined by
date or array position.

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
| ESPN scoreboard | Moneyline, totals, spreads when explicitly offered | Full game | NBA, WNBA, NCAA_M, NCAA_W, NBA G League; named-book observations are independently timestamped but relayed by ESPN, so registry policy preserves `unresolved` / `unverified` audit metadata. An exact current named-book quote can enter the paper lane only after the ordinary model, calibration, and ROI gates pass. |
| BetExplorer aggregate (`betexplorer`) | Representative average two-way moneyline prices | Full game | EuroLeague and CBA are `LIVE_VERIFIED` from captured 2026-11-16 pages. ABA League uses the same route as a probationary aggregate research observation, separate from its `aba_liga` schedule/result adapter. Every aggregate adapter quote is `aggregate_consensus` with `unresolved` permission / `probation` validation, never a named-book ticket quote. |
| BetExplorer direct (`betexplorer_direct`) | Response-labelled bookmaker H/A moneyline, O/U totals, and Asian Handicap spreads | Full game | Universal fallback for canonical future basketball fixtures. It reads one dated basketball board then up to three exact comparison responses per strict ordered-team match: `ha` for the two-way H/A moneyline, `ou` for line-specific Over/Under prices, and `ah` for home-oriented spread prices. Each emitted price retains the response's bookmaker label, exact source line/side order, source-designated main line when present, and a capped URL/SHA-256 receipt. It is operator-authorised for this single-user project and marked `named_book` with `unresolved` permission / `probation` validation; no average, cross-book best price, relay, proxy, browser automation, login, or post-tip quote is used. |
| The Odds API | `h2h`, totals, spreads | Routine acquisition: full game only | NBA, WNBA, NCAA_M, NCAA_W; a scheduled runner receives `THE_ODDS_API_KEY` only with a matching three-credit permit for date-bounded named-DraftKings collection. Australia NBL remains a documented but `probation` mapping and receives no automated permit pending credentialed region/bookmaker coverage. |

The Odds API adapter remains `DOCUMENTATION_ONLY` until a bounded authenticated
smoke verifies account, region, named bookmaker, event identity, pre-tip
timestamp, and usable quotes. The provider's documented current-odds cost is
requested markets times effective bookmaker-region groups; the scheduler
therefore reserves the maximum three-credit cost before it exposes a key. Its
shared 500-credit allowance is guarded by a 360-credit local reservation cap,
a 140-credit provider-reported reserve, and a 12-credit SAST daily cap. See
[Operations](OPERATIONS.md#the-odds-api-free-tier-reservation-guard) for the
permit and reconciliation contract.

Important restrictions:

- ESPN scoreboard books are stored even when no prediction exists. Opening
  values without their own observation timestamps remain context and cannot be
  used as closing-line-value evidence; betting links are discarded. The
  pipeline stamps these relayed observations from the central league registry,
  not from adapter defaults, so their unresolved/unverified audit metadata is
  retained rather than silently promoted to direct-provider evidence.
- The Odds API routine uses one date-bounded list request for full-game
  `h2h,totals,spreads` and the named DraftKings bookmaker group. Native H1/H2
  and quarter markets require the provider's charged per-event endpoint, so
  scheduled free-tier work never fans out into per-event calls. Their parsers
  are retained only for separately supplied payloads; no full-game line is
  relabelled or substituted for a half/quarter market. For NCAAW, official
  quarter results may still derive the corresponding H1/H2 scores for
  settlement when a valid native half ticket exists.
- The original BetExplorer fixture-table adapter exposes only a representative
  average moneyline pair. It remains separate from `betexplorer_direct`, skips
  missing/post-tip table rows, and never emits a named-book ticket price.
- The direct BetExplorer fallback uses only exact full-game source surfaces:
  H/A (`ha`) for its two-way moneyline, O/U (`ou`) with the table's explicit
  total threshold and Over/Under order, and Asian Handicap (`ah`) with the
  table's explicit home-oriented handicap and home/away order. It retains each
  exact source line in the displayed ladder. A total/spread receives a
  main-line flag only when an active source-nav marker names a
  `data-handicap` line actually carried in that same response: basketball
  responses can expose stale cross-sport nav ids, so a marker alone is not
  evidence. It never manufactures a line from an average or cross-book
  calculation. It makes no more than the configured number of direct requests,
  waits at least three seconds between live requests, and stops for the rest of
  the run after 403/429 or a challenge/verification surface. It reads a
  source-supplied comparison-nav page parameter before falling back to the
  guarded match-page parameter read. It captures only a dated board and
  comparison pages for canonical, ordered-team matches; ambiguous, reversed,
  unmatched, empty, malformed, or post-tip rows stay visible as receipt/gap
  diagnostics. It does not use a proxy, relay, browser automation, login, or
  alternate endpoint.
- BetExplorer's and the CBA data site's reuse/automation terms remain
  `unresolved`; public/keyless access does not make either source a cleared API
  integration. `LIVE_VERIFIED` labels describe endpoint/parser observations,
  not rights clearance. The direct fallback's named-book classification means
  only that the comparison response explicitly labels the bookmaker—it never
  relabels the fixture-table average or a cross-book best price.
- No current book adapter supplies highest-quarter prices. Highest-quarter
  forecasts therefore remain research-only.
- Full-game spread observations may be stored from The Odds API, but the
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
python -m ingestion.cli smoke --league EUROLEAGUE_W --date 2026-11-07
python -m ingestion.cli smoke --league CBA --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_M --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_W --date 2026-11-07
```

Smoke is read-only; `--write` is rejected. Provider health distinguishes
missing credentials, stale evidence, provider errors, missing identity, and
successful coverage. A provider returning no rows is not automatically proof
that the league or market is unsupported.
