# Results-route assessment for the five fixtures-only lanes

**Session:** 2026-11-24 UTC, external probe run (this sandbox has no egress).
**Question asked:** for each lane whose `results_source` is `source_unavailable`,
is there a permitted, free, keyless route that returns finals keyed to a stable
game identity — and what are the real board spellings while you are in there?

**Rule applied to every candidate:** a route is promoted only when *both* gates
resolve from quoted bytes — `robots.txt` (absence is an allow state under
RFC 9309, and is reported as absence, not as permission) and a terms document
that actually licenses the use. **A 200 is not permission.** Coverage is
reported as arithmetic (how many of the last ten known fixtures returned a
final), never as a verdict.

## Verdicts

| Lane | Verdict | Route | Coverage | What stops it |
|---|---|---|---|---|
| `GREECE_BASKET_LEAGUE` | **PASS candidate** | `esake.gr/el/action/EsakeResults?idchampionship=184645B9&idseason=00000001&series=N` + `EsakegameView?idgame=<HEX>` | **10/10** (incl. an OT game) | Only the ToS reading (below). No code blocker. |
| `CHINA_WCBA_W` | Endpoint proven, **STOPPED at the ToS gate** | `cba.net.cn/datahub/cbamatch/games/matchschedules?competitionId=56&seasonId=2286` | 10/10 | Host publishes **no terms document at all** |
| `LITHUANIA_NKL` | Endpoint family proven, **STOPPED at the ToS gate** | `nkl.lt/wp-json/nkl-calendar/v1/teams?season=` + `/game/{id}` | not computed (stopped) | Host publishes **no terms document**; sitemap lists only cookie/privacy pages |
| `ISRAEL_LIGA_LEUMIT` | **Nothing passes — documented** | none | 0/10 on every candidate | `one.co.il` disallows `/api/` in robots (the only path with scores); `safsal.co.il` has no ToS page; the federation host `basketball.org.il` timed out ×3 over two days; `basket.co.il` serves only the Premier board |
| `AUSTRALIA_WNBL_W` | **Not delivered** | — | — | The probe log shows real work (`prod.rosetta.nbl.com.au/get/wnbl/matches/in/season/{2025,2026}/regular` → 200; `wnbl.com.au/terms-and-conditions` → 200) but no report or verdict was handed back |

This commit changes no `results_source` value; it writes the verdict into each
lane's own `source_detail` (so the dark lane explains itself where the reports
read it) and into this document. A lane flips from `source_unavailable` only
when an adapter exists **and** has been replayed against captured bytes.

## Greece: the one that can be promoted

`esake.gr` is the league's own operator. `robots.txt` answers 404 (no rules
published → allow state, quoted as absent). The terms page («Όροι Χρήσης», 200,
104,321 B) expressly licenses copying with source credit:

> «Ο ΕΣΑΚΕ επιτρέπει στους χρήστες να δημιουργούν αντίγραφα … για δική τους
> χρήση, με την προϋπόθεση ότι αναφέρεται η πηγή … και δεν γίνονται αλλαγές»
> — ESAKE permits copies for one's own use, provided the source is credited and
> the content is not modified.

The same document scopes the copying clause to personal, non-commercial use, and
contains no automated-access prohibition. **That is a reading, not a licence we
can grant ourselves** — so it was put to the operator, who accepted the scope for
this system's paper record on 2026-11-24. That acceptance is recorded here as an
operator decision, not as a legal conclusion; the reader keeps the cheap half of
the bargain by stamping the source and the sha256 of the bytes on every receipt.

What the bytes already support: stable `idgame` identity per game
(`EsakegameView?idgame=1439C9D6` refetched identically), finals in
`esake-program-game-final-score` blocks, a server-honoured date/matchday window
(`series=1`/`series=2`/unsuffixed all differ), and 10/10 joins against the
BetExplorer baseline by score pair + day (BetExplorer is the naming baseline
only — its ToS prohibits scraping, and it is never the route).

**What shipped instead of an adapter** (same day, after the operator's
decision): `ingestion/providers/esake.py` reads a matchday page, and
`scripts/probe_esake_results.py` fetches a bounded number of them, saves each body
verbatim with its sha256, and reports two things — whether the parser's assumed
container classes exist, and whether a final can be attached to one of this
lane's games **without any club-name map at all**, by date + tip in Athens local
time. The reader resolves, or refuses, everything: an unreadable month, a
three-club row or a bare score outside the `final-score` block is counted, never
guessed. Two finals at the same minute are refused, not tie-broken.

The registry is still `source_unavailable`. One live probe run decides the next
step: if the counts come back with rows and finals, the adapter is a short step;
if the container class is absent, the parse failure is specific and the fix is
against real bytes rather than imagination.

## China and Lithuania: proven endpoints behind a missing ToS

Both hosts publish no terms document. Under the rule above that is a stop, not a
technicality: the payloads are excellent (CBA's data hub returns the whole
2026-27 season, 297,504 B, with integer `gameId`, Chinese names, scores and a
`gameStatus`; NKL exposes a WP-JSON namespace whose own index declares
`/teams` and `/game/{id}`), and both would need a terms document — or an
operator's written acceptance of that risk — before a single row enters the
corpus.

Two facts worth keeping from those probes:

- CBA's endpoint is **season-scoped, not date-scoped** (scoped variants returned
  42-byte empty lists), so it cannot satisfy a per-date read contract as-is.
- `slamsports.cn` (the WCBA operator CMS) is date-scoped and proven, but its
  public season config stops at 2025-26 — a backfill source only.

## The alias deliverables did not survive review

Both alias files answer a question this codebase does not ask, and **no
`_LEAGUE_NAME_PAIRS` rows were added**:

- **KBL.** The file pairs `teamNameEn` codes (`SONO`, `SK`, `DB`) against board
  names. `ingestion/providers/kbl.py` emits `teamNameFullEn` — the four-letter
  code is never a candidate. With the full name the captured pairs already
  agree: `WONJU DB PROMY` ↔ `Wonju DB` matches by whole-token containment, and
  the board spelling the earlier diagnosis quoted (`Wonju Dongbu Promy`) does
  not appear in any 2026-11-23/24 capture.
- **Japan.** The file pairs the schedule page's Japanese short labels (`福岡`,
  `越谷`) against board names. That page contributes keys and tips, not names:
  `bleague.py` takes names from the S3 context's `HomeTeamNameE`. English
  candidate against Latin board agrees by ordinary normalisation.

`tests/test_kbl_lane.py` and `tests/test_bleague_lane.py` now pin both contracts,
including the negative direction — the short code and the Japanese label must
stay unmatched, so a future edit that switches fields fails loudly instead of
quietly needing a twenty-row alias table per division.

## The Japan mismatch this session *did* find — resolved by splitting the lane

The probe surfaced a real, load-bearing bug: this lane read the provider's **top**
division (`tab=1`, now branded B.LEAGUE PREMIER) while the board slug it priced
against, `japan/b-league-one`, is the **second** division (B.LEAGUE ONE — the
captured tab=2 slate matched that board page game-for-game). Different clubs, so
no name could ever agree and the aggregate price source was dead by construction.

One lane could not be two competitions. The choice was left open then; on
2026-11-24 it was made by splitting rather than picking:

| lane | division | schedule | aggregate price |
|---|---|---|---|
| `JAPAN_B_LEAGUE` | B.LEAGUE PREMIER (tab=1) | `bleague` | **none claimed** — no PREMIER board page has been captured, so it guesses no slug; prices come only from the bounded direct day-board fallback |
| `JAPAN_B_ONE` | B.LEAGUE ONE (tab=2) | `bleague` | `basketball/japan/b-league-one/fixtures/` — the one Japan page captured, which matched this division game-for-game |

The division is now derived from the lane (`bleague.TAB_BY_LEAGUE`) and refused
when a caller asks for the other one, so the mismatch cannot be reintroduced by a
forgotten argument. `tests/test_japan_division_split.py` holds that invariant,
including that re-reading a B.ONE game's context cannot relabel it PREMIER.

**Closed 2026-11-24:** a B.LEAGUE PREMIER board page was read by an external browser
agent -- `basketball/japan/b-league-premier/fixtures/`, HTTP 200, served heading
*"B.League Premier 2026/2027 Fixtures"*, dated rows from 04.12. -- so that lane now
claims the slug and its aggregate price. The provenance is stated in the code: a
third-party read, so this pipeline's first live fetch confirms the slug rather than
assuming it, and a wrong slug fails loudly as a named gap instead of quietly pricing
the wrong clubs.

## The operator decisions of 2026-11-24

The licence gate on four lanes is now opened by a person rather than by a host,
and recorded as data in the registry (`LeagueSpec.operator_acceptance`), not as a
comment: who decided, when, what may be read, and on what basis.

| Lane | Basis | What was accepted |
|---|---|---|
| `GREECE_BASKET_LEAGUE` | esake.gr terms scope copying to personal, non-commercial use | that clause, for this paper record |
| `CHINA_WCBA_W` | cba.net.cn publishes **no terms document at all** | read the documented data-hub endpoint, paper-only, no redistribution |
| `LITHUANIA_NKL` | nkl.lt publishes **no terms document** (sitemap: cookie + privacy only) | read the documented WP-JSON surface, honouring `Crawl-delay: 10` |
| `AUSTRALIA_WNBL_W` | route assessment never completed; wnbl.com.au publishes terms | read the rosetta per-season route and those terms, paper-only |

**An acceptance is not an activation.** All four lanes still report
`results_source = source_unavailable`, because the remaining gate is not a licence
— it is a served page the parser has been replayed against. `tests/test_operator_acceptance.py`
asserts exactly that, so a future edit cannot turn a decision into a fake lane.

`ISRAEL_LIGA_LEUMIT` was retired on the same day, for the same reason as the nine
retired lanes before it: the schedule side is served, every candidate results
route is closed, so it could never settle and its fixture cards graded nothing.

## What an egress host changed on 2026-11-24 (14:02–14:03Z)

The probes that produced this document ran on a host without egress. Re-running
three of the routes from GitHub Codespaces — which has egress — moved the verdicts
on all three, and not in the direction anyone wanted:

| Lane | Earlier reading | Re-read from egress | What it means |
|---|---|---|---|
| `CHINA_WCBA_W` | 297,504 B of season JSON, 10/10 coverage | **HTTP 500, 179 B**: `param name: roundId value not exists` | the documented URL is not sufficient; a `roundId` is required, and the error names no values |
| `LITHUANIA_NKL` | "endpoint family proven" | `teams?season=2026` → 200 (1,390 B); `game/125863` → 200 but **schedule only**: `id, summary, location, start_local, start_utc, end_utc, url, ics` — no score, no status | `nkl-calendar` is a calendar, not a results source |
| `AUSTRALIA_WNBL_W` | 200, 493,824 B via referer | **HTTP 401, 88 B** even with the referer | the referer alone no longer opens it |

Two of the three earlier verdicts were therefore **overstated**, and the registry
now says so in each lane's own `source_detail` rather than quietly improving. The
lesson is worth keeping: a route proven once is a hypothesis about a host, not a
fact about it.

Greece was re-read the same day and came back the other way: `esake.gr/robots.txt` is
**404 (absent, an allow state)**, and the series=8 results page answers **200** with the
matchday label (*08η Αγωνιστική*), *Σαβ 21 Νοε - 15:00*, a venue, and
`ΗΡΑΚΛΗΣ 84 - 87 ΑΕΚ` -- scores, server-rendered, as plain text. What that read did not
yield is a container class or a game id, because the agent returned rendered text
rather than HTML. Saving the HTML is exactly what `scripts/probe_esake_results.py` does.

`GREEK` remains the shortest path: `scripts/probe_esake_results.py` was written for
exactly this, and no egress read of `esake.gr` has yet succeeded.

## Open items

1. **Greece:** run `scripts/probe_esake_results.py --series 8` once from a host with
   egress and hand back `esake_probe_receipt.json` plus the saved bodies. That
   replaces the derived fixture with a served page and answers the date+tip join.
2. **WNBL:** request the missing report — the probe log suggests it got furthest
   of all five.
3. **Japan:** choose the division.
4. **China:** run `scripts/probe_cba_results.py --rounds 1-40` from an egress host;
   it enumerates the round space the 500 names and saves every body.
5. **Lithuania:** read the two namespaces the index lists and this probe did not
   test (`nkl/v1`, `nkl-advanced-statistics/v1`) for a score-bearing route.
6. **WNBL:** capture the 401 body and the headers the site's own client sends; the
   referer-only read that worked once no longer does.
7. **Israel Leumit:** retired 2026-11-24; no candidate route is permitted.
