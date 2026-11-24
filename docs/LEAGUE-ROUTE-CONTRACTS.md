# League route contracts: verified leads and blockers

This page records what has actually been verified for each lane — implemented
or still open — so the remaining work is mechanical rather than speculative. A
route lead is **not** an integration: a lane leaves `source_unavailable` only
when its adapter, identity checks, and fixtures exist and pass.

Hosts and endpoint shapes below were read from published documentation,
third-party client source, and a probe run by the operator's external agent on
2026-11-22. Where that probe transmitted a whole body it is committed byte-exact
and is labelled **captured**; where only verbatim excerpts were transmissible the
fixture says so. Everything else stays a `referenced route contract`. The two
classes are listed per file in `tests/fixtures/CAPTURE-PROVENANCE.md`.

## Ready pattern: a documented period-level route exists

| Lane | Host | Documented routes | Notes / blocker |
| --- | --- | --- | --- |
| Spain ACB | `api2.acb.com` | `seasondata/Calendar/season-calendar?season={end_year}`, `matchdata/MatchHeader/match-header?matchId=` | **Implemented** in `ingestion/providers/acb.py`, but the header route measured **401 with an empty body, no `WWW-Authenticate`, server Kestrel** on 2026-11-22, so the lane reports `auth_required` until that route is authorised — the response names no scheme to satisfy. |
| Israel Premier League | `basket.co.il/ws/ws.asmx/Games` | `?cYear=&board_id=5&board_round=0&team_id=0&team_uid=0&game_id=0` | **Implemented** in `ingestion/providers/israel_ibl.py`, now backed by four **captured** rows (2026-11-22): pipe-delimited `score_by_q_1/2` are points-in-period, a double-overtime row is explicit, rows carry `game_id` and `year_id`. No UTC offset — `Asia/Jerusalem` applied as a recorded assumption. Browser UA required (Cloudflare 1010). |
| Italy LBA | `www.legabasket.it/api` | `championships/get-championships?cs_id=1&items=1000&s={first_year}`, `get-championship-match-recap-by-id` | **Probed 2026-11-22:** live. The season key is the **first** year of the pair (2026 = "2026/27"), correcting this page's earlier end-year note. **Implemented** in `ingestion/providers/lba.py`, now **captured** (2026-11-22): the championships (867 B) and focus-day calendar (11636 B) bodies are committed byte-exact, and the play-by-play period ends for two matches (regulation 100-86 and a real overtime 92-106, `87-87` after regulation) are committed as verbatim actions. `game_status` is a numeric code (`2` final, `1` live, `0` not started) and the captured slate spans two dates while marking one *round* focused. Day selectors are ignored server-side, so the route cannot enumerate a season — history backfill is not wired. |
| Japan B.League | `b-league.s3.amazonaws.com/web_json` | `contexts/{schedule_key}` | **Probed 2026-11-22: blocked.** `?list-type=2&prefix=web_json/` answers **403 `AccessDenied`** (body captured), `www.bleague.jp/schedule/` exposes no key strings, and `api.bleague.jp` is unreachable, so the schedule key has no publishable discovery step. Stays `source_unavailable`. |
| South Korea KBL | `api.kbl.or.kr` | `season/list` plus schedule routes; requires `Channel`, `Device` and `TeamCode` headers | **Probed 2026-11-22: rejected for this pipeline.** The schedule is live with finals and stable team codes, but no API route exposes quarter scores, so a reconcilable `ActualResult` is impossible. Stays `source_unavailable`. |
| Australia NBL | module exists in `sportsdataverse-py` (`nbl_au`) | file path and routes not confirmed from documentation | **Probed 2026-11-22: no route.** The published sdist contains 0 `nbl` files — the module is not shipped — so there is no compliant route to reference. Stays `source_unavailable`; The Odds API NBL key remains probationary and never supplies fixtures. |

## No packaged adapter found

These have no maintained open-source period-level route. Each needs its own
site contract (calendar + period scores + stable team identity) or a
browser-automation fallback, which this project has ruled out as a primary
route.

- Basketball Champions League — **probed 2026-11-22:** GDAP competition
  `209123` is live (136 games, `competitionCode` "BCL", scores,
  `gameDateTimeUTC`), so fixtures and finals are available through the FIBA
  GDAP pattern already used for EuroLeague Women. Period-level extraction from
  the game page was not achieved, and without periods the lane cannot reconcile
  — it stays unregistered until that is solved.
- France LNB Pro A, Germany BBL, Turkey BSL, Greece Basket League, Brazil NBB,
  Argentina Liga Nacional. Probed 2026-11-22: NBB's schedule JSON is live but
  carries no period JSON; Greece exposes quarter scores in server-rendered HTML
  only (no JSON route found); Turkey's federation API sits behind a Nuxt SPA and
  no schedule endpoint was extracted. Each needs its own contract.

## Evidence that no pricing work is needed first

Every one of the lanes above already registers `betexplorer_direct` as a price
source, and for most it is the only one. Pricing was never the blocker: the
missing half is fixtures plus reconciled periods. BetExplorer itself cannot
close that gap — its adapter parses day boards and full-game moneylines/lines
but no quarter data, so it can never be a result source.

**League-page probe, 2026-11-22:** the basketball league pages for
`italy/lega-a`, `israel/super-league` and `spain/acb` answered 200 with fixture
tables and odds markers, and the same slug shape exists for `germany/bbl`,
Greece, Turkey, `brazil/nbb`, Argentina and `australia/nbl`. That confirms a
page exists, *not* that the direct fallback yields a named-book moneyline per
fixture on those boards — that per-fixture pricing behaviour is **unprobed** and
stays probationary.

## Lane status: retired, off-season, active

`ingestion/league_registry.py` carries a `LaneStatus` per competition, applied
from one table (`_RETIRED_ON_2026_11_24`, `_OFFSEASON_ON_2026_11_24`) so the
decision is auditable in a single place rather than scattered across specs.

The rule the table encodes: **retire a lane when its source refuses us or
serves nothing usable — not when its season is merely quiet, and not when it
merely lacks a price route.** Measured zero priced rows on an in-season
fixtures page, or no compliant schedule/result route at all, is a fight worth
ending. A lane with a working adapter keeps its place even while unpriced,
because prices can arrive later without re-deriving the fixtures.

Measured on 2026-11-24 and **retired**: Spain ACB, Italy LBA, Basketball
Champions League, Basketball Africa League, Argentina Liga Nacional, France LNB
Pro A, Germany BBL, Turkey BSL, Brazil NBB. Each of those ran for roughly 750
scheduled runs and acquired **six games between all nine of them** (ACB, BCL,
BAL, Argentina, France, Germany, Brazil: zero; LBA: six), per their own committed
archive manifests.

**ABA League was considered and rejected.** It looked like a fighter on the
fixtures page — 0 of 100 rows priced — but that page only ever serves a site
*average*, which is research-only here. Its ticket route is the direct
named-book comparison, and it demonstrably worked: on 2026-11-22 it froze three
priced picks (Stake.com 1.06 / 1.78 / 1.25, `named_book`, pre-match) out of 30
logical picks, and it holds 4,252 games with finals reconciled back to 2004.
Those three were CAUTION rather than clean only because calibration held n=13
against a minimum of 30 walk-forward samples. A lane waiting to mature is not a
source refusing us. `_KEPT_AFTER_REVIEW` records this, and a test asserts that
no retired lane has a matched-price pick in its own state — the guard that
caught the wrong call.
The WNBA is **off-season** (April–October calendar), which is a calendar fact
and must never be reported as a failure. The Chinese CBA stays **active**: its
empty fixture page on the 17th–18th was the league's own break between round 8
(2 November) and round 9 (4 December), not a refusal and not an off-season.

Retirement is operational, not destructive. A retired lane keeps its enum
member, contract, state key, price policies, adapter, stored finals and
`localdata/` history — so it can still settle what it already holds — but it is
never scheduled, never probed for a schedule, and never counted as work:

* `adaptive_gate` decides a dark lane **before** any state read or schedule
  probe and reports `should_run: false` with `lane_retired: …` /
  `lane_offseason: …`; the lane therefore leaves the workflow matrix without
  anyone editing a workflow file.
* `scripts/daily.py` refuses a retired league with exit code 2 and a JSON
  reason, so a stale cron line cannot silently resurrect it. `--run-retired`
  is the deliberate override, for inspection or revival only.
* `scripts/daily.py --league NBA --list-lanes` prints the whole table
  (`league`, status, schedule source, price provenance/permission, reason).

Revival is a one-line flip: `dataclasses.replace(spec, status=LaneStatus.ACTIVE)`
restores a fully schedulable lane, and a test asserts exactly that.

## The fixtures-table lanes (added 2026-11-24)

Four competitions that had no schedule route of their own, plus Greece revived
on the same adapter, now take both fixtures and prices from a single fixtures
page each (`ingestion/providers/betexplorer_fixtures.py`, which reuses the row
parser in `betexplorer.py` rather than growing a second one):

| lane | page | measured 2026-11-24 (rows / priced / blank kickoff) |
|---|---|---|
| WCBA Women (`CHINA_WCBA_W`) | `/basketball/china/wcba-women/fixtures/` | 255 / 3 / 2 |
| Lithuania NKL (`LITHUANIA_NKL`) | `/basketball/lithuania/nkl/fixtures/` | 167 / 7 / 99 |
| WNBL Women (`AUSTRALIA_WNBL_W`) | `/basketball/australia/wnbl-women/fixtures/` | 67 / 3 / 0 |
| Israel Liga Leumit (`ISRAEL_LIGA_LEUMIT`) | `/basketball/israel/liga-leumit/fixtures/` | 161 / 5 / 83 |
| Greece Basket League (`GREECE_BASKET_LEAGUE`) | `/basketball/greece/basket-league/fixtures/` | 126 / 7 / 2 |

Five properties of that page decide how the lane behaves, each measured rather
than assumed:

1. **The clock is the site's own CET/CEST wall clock for every competition** —
   not the league's local time, and there is no ISO attribute on the page to
   read. The label is parsed as naive `Europe/Paris` and converted to UTC, and
   every game records that reconstruction in its evidence.
2. **Identity is the 8-character match code in the row href.** The table serves
   transliterated names with sponsor suffixes, colliding city fragments
   (`Paris`, `Dubai`) and a ` W` suffix on women's clubs, so a name join would
   be a guess. Games and their price observations share
   `betexplorer:<match_id>`, which is what makes the join exact.
3. **Rows with a blank kickoff cell carry no time the server ever sends.** They
   are counted and reported (`fixtures_kickoff_missing`), never dated by
   assumption and never emitted as games.
4. **A slate with no quote says so** (`fixtures_price_coverage`) instead of
   looking like a page that was never read; a page that failed to serve a table
   raises rather than returning an empty slate.
5. **No settlement route.** These pages serve no final score, so
   `results_source` is `source_unavailable` and the run emits an
   `official_results` gap by name. The lanes produce fixtures, forecasts and
   research-only prices; they cannot settle, and the code says that instead of
   implying otherwise.

Prices on these lanes are an aggregate consensus from an unreviewed source
(`SourcePermission.UNRESOLVED`), so they are research-only and never
ticket-eligible. The gate fails open for this source — the page is season-wide,
so a half-hour gate must not pull it and call the read a probe.

## Order of work

Done: EuroCup, Spain ACB (schedule only, gated header), Israel Premier League,
Italy LBA. Rejected on evidence: South Korea KBL (no period route), Japan
B.League (no manifest), Australia NBL (module not shipped), Greece and Turkey
(no JSON route / blocked SPA).

0. **Open lead, unverified:** `get-championships-matches-by-id?id={championship}`
   answered 200 (9975 B for the whole season 602 request) and the same route with
   `id={match}` returns a match detail whose `prev_matches` rows carry ids,
   scores and UTC timestamps. If that season response really enumerates matches
   it would remove the one gap that keeps LBA from history — but a 9975-byte body
   for a 30-round season is more consistent with a page or a summary, and no body
   was transmitted for it, so nothing is built on it yet.
1. Basketball Champions League period extraction: GDAP already yields the
   schedule, so only per-game period scores are missing.
2. Any B.League manifest, if the league ever publishes a listing.
3. A date-bounded route for Italy LBA, the one implemented lane whose calendar
   cannot enumerate a season; until then it carries no history backfill.
4. Named-book moneyline depth on the probed BetExplorer league boards, which
   stays probationary until a per-fixture price is actually observed.
6. The seven domestic leagues as individual contracts; do not batch them
   behind a generic scraper, because each needs its own identity crosswalk and
   its own failure semantics.
