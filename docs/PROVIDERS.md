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
| Israel Winner League JSON | Israel Premier League (`ISRAEL_PREMIER_LEAGUE`) | Game board with per-quarter scores, OT totals, finals, team ids and English names | `basket.co.il/ws/ws.asmx/Games` (browser UA required) | **Captured 2026-11-22** by the operator's external probe: four rows committed verbatim, confirming points-in-period quarters, a double-overtime row, `game_id`, and the `year_id` season field; tip times carry no UTC offset, so `Asia/Jerusalem` is applied as a recorded assumption |
| ACB JSON API | Spain ACB (`SPAIN_ACB`) | Liga Endesa season calendar, per-game header with names, status, final score, quarter/OT reconciliation | Public endpoints (`api2.acb.com`) | Route contract referenced from SportsDataverse's generated `acb` module; **no Ma Golide first-party capture**; source automation/reuse terms unverified |
| LBA JSON API | Italy LBA (`ITALY_LBA`) | Serie A championships list, focus-day calendar with UTC-offset tips, finals and status codes, per-match play-by-play reconciled into quarters | Public endpoints (`www.legabasket.it/api`) | **Captured 2026-11-22** by the operator's external probe: the championships (867 B) and focus-day calendar (11636 B) bodies are committed byte-exact and re-hashed in `tests/test_capture_fidelity.py`; play-by-play is committed as verbatim period-end actions. The calendar serves only the league's current focus day, so no season enumeration and no history backfill |
| EuroLeague | EuroLeague, EuroCup (`EUROCUP`) | Schedule discovery, game results, quarter/OT reconciliation, standings | Public endpoints | EuroLeague: schedule/boxscore live captures 2026-11-06; season schedule scores/partials, empty-season trap, and v3 per-round standings live-verified 2026-11-16. EuroCup: same host and route shape with competition code `U`, fixture-tested only — no Ma Golide first-party capture. Upstream reuse terms not reviewed |
| FIBA GDAP + game page | EuroLeague Women (`EUROLEAGUE_W`) | Competition schedule, canonical UTC fixtures/status/results, official game-page play-by-play/period scores, player/team detail | Public APIM key shipped in FIBA's web client | Live facts and fixture-only cases are separated below; independent namespace, no men's EuroLeague reuse. Automation/reuse rights and future route/status vocabulary remain unresolved. |
| CBA | CBA | Official league data API: season/day discovery, Beijing-local UTC times, quarter/OT reconciliation via observed period-score POST endpoint | Public endpoints; AES-128-ECB bodies decrypted with the key constant shipped in the site's own bundle | Live-verified 2026-11-16 from a datacenter IP (decryption, statuses, quarter sums, OT shape); unofficial/undocumented API and terms not reviewed — see caveats |

Negative evidence is captured too: the KBL board, `team-record` and `text-cast`
bodies were transmitted and contain finals, `gmkey`s and per-action quarter
labels but **no per-quarter totals**, and `season/list` answers `400` with
`errorCode -1101` unless the channel headers are sent. B.League's S3 listing
answers `403 AccessDenied`, and the sportsdataverse sdist ships no `nbl` module
at all. Those lanes stay refused on that evidence.

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

### Israel Premier League acquisition details (referenced route contract, fixture-tested)

- ``GET https://basket.co.il/ws/ws.asmx/Games?cYear={end_year}&board_id=5&board_round=0&team_id=0&team_uid=0&game_id=0``
  returns one season board as JSON. Boards: 5 Winner League, 10 State Cup, 33
  play-in, 34 Super Cup. ``cYear`` is the season **end** year.
- Periods arrive as pipe-delimited strings — ``score_by_q_1`` (home) and
  ``score_by_q_2`` (away) — with ``total_ot``/``overtime``, finals in
  ``score_team1``/``score_team2``, and team identity in ``team1``/``team2`` plus
  ``team1_name_eng``/``team2_name_eng``. Quarter values are points scored in
  that period — an external 2026-11-22 audit of a captured board checked all 45
  rows and found none cumulative, which is how the adapter already treats them.
  That audit also confirmed this host publishes a per-game ``game_id`` (read
  preferentially for identity) and exposes overtime three ways (extra pipe
  elements, ``total_ot``, ``overtime``); the adapter's stricter "refuse an OT
  total without a fifth element" rule is deliberate and keeps the lane
  fail-closed rather than loosening to the source's least reliable channel.
- The host is Cloudflare-fronted and rejects a non-browser client with error
  1010, so the adapter's own client sends a browser ``User-Agent``. An injected
  fixture transport in tests bypasses it.
- **The provider exposes no UTC offset.** ``game_date_txt`` is ``dd/mm/yyyy``
  and ``TimeOfGame`` is local Israeli wall clock. The adapter applies
  ``Asia/Jerusalem`` explicitly, records the assumption in
  ``provider_ids``/``evidence``, refuses a wall time that falls in a DST gap,
  and flags an ambiguous (repeated) wall time.
- Overtime is reconciled only when its points are actually exposed: a row that
  reports ``total_ot`` without a fifth period is refused, never treated as a
  regulation final.
- **Captured rows, 2026-11-22:** four board rows are committed verbatim in
  `tests/fixtures/israel_ibl/board5_captured_rows.json` and are re-tested on
  every run. The probe reported that all 45 played rows of the full board sum to
  their finals; the four committed rows re-verify the reading here, including a
  double-overtime row of ``21|18|18|18|11|10`` against ``20|19|21|15|11|17``
  (96-103, ``total_ot`` "2"). They also show that a row's season field is
  ``year_id`` — not the ``cYear`` query parameter. Reading only ``cYear`` left
  every discovered game with an empty season, so a board refresh refused
  identity for the whole lane; the row season is now read from ``year_id``,
  then ``cYear``, then the requested board season.
- Pricing stays on the universal direct BetExplorer fallback.

### Italy LBA acquisition details (referenced route contract, fixture-tested)

- ``GET https://www.legabasket.it/api/championships/get-championships?cs_id=1&items=1000&s={year}``
  lists competitions for a season. ``s`` is the season's **first** year
  (``2026`` == 2026/27). The adapter requires exactly one ``cserie_code == "A1"``
  row; zero or several is refused rather than picking one.
- ``GET .../championships/get-championships-calendar-by-id?id={championship}``
  returns the games: match id, home/visitor ids and names, finals, status and
  ``match_datetime`` with an explicit offset. A timestamp without an offset, or
  a row naming the same team on both sides, is recorded as a gap and skipped.
- ``GET .../championships/get-championships-matches-play-by-play?id={match}``
  returns actions carrying ``period``, ``order`` and a cumulative ``score``. The
  adapter takes the highest-``order`` scored action per period, differences
  consecutive period ends into per-period points, and requires periods 1–4 plus
  overtime that is contiguous from period 5. A jump from regulation to period 6
  is refused, not filled in.
- **The probed day selectors are silently ignored server-side.** The calendar
  answers with the league's current focus day whatever date is requested, so the
  adapter treats a run's target date as a filter over that one slate, never as a
  date-bounded query. That also means a season cannot be enumerated through this
  route: history backfill is deliberately not wired.
- ``game_status`` is a numeric code, not a label: ``"2"`` final, ``"1"`` live,
  ``"0"`` not started. The captured focus day mixed all three and two live rows
  already carried real scores (63-62, 49-48), so status comes from the code —
  an unknown code maps to ``UNKNOWN`` and is never promoted to a final.
- One focus day is a **round**, not a date: the captured slate spanned
  2026-11-21 and 2026-11-22 while ``filters.days`` marked round 9 as focused.
  A run's target date filters that one slate; it is not a date-bounded query.
- ``additional_time`` was ``0`` on every captured row, so overtime is derived
  from the play-by-play period count alone and is never cross-checked against a
  field that was never observed to vary.
- ``get-championships-matches-play-by-play`` answered **500** for one match
  without play-by-play, not 404 — an unavailable body is a transient provider
  error and the game stays a recorded gap.
- The adapter needs its own client with a browser ``User-Agent``; tests inject a
  fixture transport instead.
- Pricing stays on the universal direct BetExplorer fallback.

### Spain ACB acquisition details (referenced route contract, fixture-tested)

- ``GET https://api2.acb.com/api/seasondata/Calendar/season-calendar?season={end_year}``
  returns every round of a season with its games: match id, season-team ids,
  scores, ``startDateTime``, ``matchStatus`` and ``currentQuarter``. The season
  value is the season **end** year (2027 = 2026-27) and the API's edition id is
  that year minus 1936; anything below the 1984 floor is refused before a
  request is built.
- ``GET https://api2.acb.com/api/matchdata/MatchHeader/match-header?matchId={id}``
  returns one game's header: team names and stable ids, ``matchStatus``
  (``FINALIZED`` once over), ``currentHomeScore``/``currentAwayScore`` and
  ``quarterScores[]`` — points scored *in* each period, with overtime as
  quarters 5+.
- The calendar carries no team names, so identity is resolved from the header
  and a slate costs one calendar request plus one header request per candidate
  game. A header that fails identity, a timestamp without a UTC offset, or a
  body without four regulation quarters is recorded as a gap and skipped — never
  folded, guessed from list order, or treated as missing points.
- No history backfill is wired: a season of history is one header request per
  game, so it needs its own bounded plan rather than an unbounded walk.
- **Authorisation, measured 2026-11-22:** the season calendar answers publicly,
  but ``matchdata/MatchHeader/match-header`` answers **401** with an *empty*
  body (``content-length: 0``), server ``Kestrel``, and **no**
  ``WWW-Authenticate`` header — so the response names no scheme to satisfy and
  the gate cannot be cleared from documentation. Because the calendar carries
  neither team names nor period scores, a gated header means the lane cannot
  currently produce a canonical game. That is reported as a distinct
  ``auth_required`` status with one clear gap, not as an empty season.
- Pricing stays on the universal direct BetExplorer fallback.

### EuroCup acquisition details (referenced route contract, fixture-tested)

- EuroLeague Basketball serves EuroCup from the same host with an identical
  route shape; the first letter of the season code selects the competition
  (`E` EuroLeague, `U` EuroCup). The adapter derives the competition from the
  season code rather than accepting it as a second, disagreeable parameter, so
  a EuroCup run cannot silently inherit the men's EuroLeague code. Bare-year
  and unknown-letter season codes are refused before a request is built.
- The lane reuses the validated EuroLeague acquisition code paths: one request
  per season of reconcilable finals, the `euroleague-history` cursor keyed by
  league, the date-bounded discovery filter, and the boxscore result path.
  EuroCup canonical game ids are `U{season}-{gameCode}` and cannot collide with
  EuroLeague ids.
- Evidence status is honestly limited: the `U` competition code is referenced
  from the provider host and two independent public clients, but no Ma Golide
  first-party EuroCup capture exists. Tests use a structural transform of the
  live-captured EuroLeague schedule fixture (only competition identity is
  rewritten), which exercises reconciliation against real provider row shapes
  without claiming a EuroCup capture.
- Pricing stays on the universal direct BetExplorer fallback; no EuroCup price
  contract has been validated.

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
| BetExplorer direct (`betexplorer_direct`) | Response-labelled bookmaker H/A moneyline, O/U totals, and Asian Handicap spreads | Full game | **Named-book per-fixture depth remains unproven (probe 2026-11-22):** the league board and finished-fixture rows carry only *average* odds (`data-odd`), the comparison loads client-side, and the route the site's own bundle uses (`/match-odds/{event}/{page}/{tab}/bestOdds/`) answered **404 with an empty body** to a headerless client for every token order tried (two 429s during rapid probing; the 404 reproduced after backoff), while the match-content fragment returned a shell with an empty `#odds-all`. Our adapter's route shape is that same one, so an unauthorised price fails closed to *unpriced* — never to a guessed price. Universal fallback for canonical future basketball fixtures. It reads one dated basketball board then up to three exact comparison responses per strict ordered-team match: `ha` for the two-way H/A moneyline, `ou` for line-specific Over/Under prices, and `ah` for home-oriented spread prices. Each emitted price retains the response's bookmaker label, exact source line/side order, source-designated main line when present, and a capped URL/SHA-256 receipt. It is operator-authorised for this single-user project and marked `named_book` with `unresolved` permission / `probation` validation; no average, cross-book best price, relay, proxy, browser automation, login, or post-tip quote is used. |
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
