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

## Greece, first replay against a served page (2026-11-24T14:21Z)

`scripts/probe_esake_results.py --series 8` ran on an egress host and saved the
page: **139,358 B**, sha256 `a721bd89feea4da44e15602136350062ced3c9b4f0d0948bf37a0a5f9dfa8b43`,
robots.txt absent (404).

| signal | count | reading |
|---|---|---|
| `esake-program-game` containers | 14 | the row class guess was right |
| rows assembled | 7 | two containers per game |
| rows with an `EsakegameView?idgame=` id | **0** | the identity pattern is wrong for this markup |
| rows with a final-score block | **0** | the score pattern is wrong for this markup |
| finals written | 0 | correct: nothing is written from a guess |

So Greece moved from "route unproven" to "route served, parser wrong" — a much
better problem, and the first of the five lanes to be replayed against real bytes.
The next read needs two focused questions of the saved page rather than another
whole probe: what the game link actually looks like, and where the score sits.

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

## Greece, the split point found (2026-11-24T14:37Z)

The replay wrote zero finals, and the reason was in the page's own bytes rather than
in its contents. Printing 900 bytes either side of one game anchor answered both open
questions at once.

**The identity was there.** `EsakegameView?idgame=3C7521D5&mode=3` and `&mode=2`, eight
uppercase hex digits, two anchors per game. The reader had reported `no_identity: 7` on
a page holding fourteen anchors.

**Fourteen containers for seven games.** The toggle handler writes the container name
inside a script attribute:

```html
<a href="#" onclick="$j(this).closest('.esake-program-game').find('.esake-program-game-hidden-row')…">
```

`_ROW_RE` split on the bare token `esake-program-game(?![-\w])`, and `'` after `game` in
`.closest('.esake-program-game')` satisfies that lookahead. Every row was therefore cut
at its own toggle: the first half held both anchors and no score, the second half held
the final-score block and no anchor. Seven plus seven is the fourteen the receipt
reported, and it explains all four numbers at once — `containers: 14`, `rows: 7`,
`no_identity: 7`, `with_final: 0`.

The fix is one requirement: a container is `class="…esake-program-game…"`, not any
occurrence of the token. The lookahead still excludes `-row-1`, `-final-score` and
`-col-1`, which are parts of a row. Clubs and totals are then read from the
`final-score` block alone, and the hidden per-quarter row — which repeats both club
names and carries four numbers each — is cut before names are counted, because four
repeats of two names is still two clubs.

**What replays, and what does not.** The captured fragment replays to one container, one
row and identity `3C7521D5`, which is the regression that mattered. It does not replay to
a total, because the printout stopped inside the score cell. A full-page replay against
the saved `esake_results_series-8.html` is the remaining gate, and until it passes the
lane stays `source_unavailable`: the reader is fixed, not yet proven.

## Lithuania, every public namespace closed (2026-11-24T14:33Z)

`urbo-nkl/v1/games` → **401** `rest_unauthorized`, "Trūksta Authorization Bearer
antraštės" (missing bearer header). `nkl/v1/live-scores` → **200** with `[]`, with and
without `?date=`. `nkl-amw/v1/games` → **401** `rest_forbidden`, "Deja, Jūs negalite to
daryti". `nkl-calendar/v1` is public and is a calendar: `start_local`, `start_utc`,
`url`, `ics`, and no score or status field.

So the WP-JSON surface has no public results route left to try. That is a *closing*
result rather than an open one: the next read for this lane is the site's own match
pages, not another namespace.

## China, both coordinates accepted and empty (2026-11-24T14:33Z)

Thirty requests — `roundId=1` crossed with `teamId=1..30` — all answered **200 in 42 B
with zero games**. The important part is what stopped happening: the host no longer says
`teamId value not exists`. Both parameters are now accepted, so the value space is the
(round, team) grid the two errors implied, and round 1 simply holds no game for those
team ids.

Two things follow. First, the grid is a real cost: 40 requests per run, and the tool
refuses a wider sweep rather than trimming it. Second, the 42-byte body is now the
thing to read, because an empty list and a refusal wearing a 200 look identical from
outside, and only one of them means "no games here".

## WNBL, the aggregator's own results page (2026-11-24T14:41Z)

Path-only, robots-allowed, 200 in 129,186 B (`f4249ca…6115c`): a table with **67
match-code links, 34 distinct codes and 35 colon scores**. The page also contains the
string "No matches found", which is a hidden empty-state template rather than a verdict
on the table — the reason `index_signals` reports that hint *beside* the counts instead
of letting it decide anything.

34 games is a season-to-date: this is the entry point the lane lacked, because its
fixtures table holds only future rows and the capture tool derives candidates from
finished rows on that table. `--index` was added for exactly this shape.

One thing the counts cannot answer is what a row looks like, and a parser written from
counts is a guess with a digest attached. `--index` therefore now saves a 2 KB window
around the first match code as `results-index-window.txt`. That is the artifact the
reader gets written against, and the same command yields one for WCBA and NKL.

## China, the 42 bytes (2026-11-24T14:41Z)

```json
{"data":[],"status":0,"message":"success"}
```

Not a refusal wearing a 200: `status` is 0 and the message says success. Round 1 holds
no game for team 1, which is a statement about the round rather than about the route.
So the grid is accepted end to end and the remaining question is *which* round carries
games — a scoping question, answerable by walking one team across rounds inside the same
40-request budget.

## Greece, the route reads (2026-11-24T14:45Z)

The full-page replay on the fixed reader, against the same 139,358 B body:

```
{'containers': 7, 'rows': 7, 'with_final': 0, 'without_final': 0,
 'names_ambiguous': 0, 'unresolved_date': 7, 'duplicate_idgame': 0, 'no_identity': 0}
84-87, 100-64, 83-87, 99-77 …
```

Every count that used to be wrong is now right. Seven containers for seven games, seven
identities, seven totals, zero ambiguities — and `with_final: 0`, because the only thing
still missing is the **date**.

The served date cell breaks *inside* the month name: the raw bytes carry newlines and
indentation between «Νοε» and «μβρίου», so a pattern demanding `21<space>Νοεμβρίου` finds
nothing on a page that plainly shows a date. Whitespace inside a word is a rendering
artifact, so the reader now retries with all whitespace removed — «Σαβ21Νοεμβρίου19:00»
still bounds the day, because a month name is letters and a digit run is not.

A second, quieter defect surfaced in the same output: one club read as the tail of its
own name. `<span>(.*?)</span>` stops at the *inner* close tag, so a club cell that nests
a span loses everything after it. Span extraction is now depth-aware.

Both fixes are unit-tested on the shape that produced them, and neither moves the lane
off `source_unavailable`: what gates it now is a dated final read from the served bytes.

## The aggregator's results pages, three lanes (2026-11-24T14:45Z)

| lane | bytes | codes | distinct matches | colon scores |
|---|---|---|---|---|
| WNBL | 128,973 | 67 | 34 | 35 |
| WCBA | 122,154 | 49 | 25 | 26 |
| NKL | 192,203 | 279 | 140 | 141 |

Three lanes that could not settle now have a page, in one command each, path-only and
robots-allowed. NKL is the striking one: the host refuses every WP-JSON namespace that
carries a score, and its public namespace is a calendar with no score field — while a
140-match season sits readable on the aggregator's own results page.

All three pages carry the string "No matches found" and all three carry matches, which is
the empty-state template again. A hint is not a verdict.

## China, the grid exhausted (2026-11-24T14:45Z)

`teamId=1` walked across `roundId` 1–30: thirty requests, every one `200` in **42 B** with
zero games. Combined with the earlier `roundId=1 × teamId=1..30` sweep, no coordinate in
the ranges the errors implied contains a game.

The route accepts both parameters and answers politely; what is missing is a *real* team
id or a *real* season id, and the errors do not disclose either. That is the difference
between "unexplored" and "exhausted", and CBA is now the second: further guessing spends
requests without buying information.

## Lithuania's sitemap (2026-11-24T14:45Z)

`https://nkl.lt/wp-sitemap.xml` answers **301 in 0 B** — a redirect, not an absence. Any
read of it needs `-L`, and the hop is the one thing to record before trusting a body.

## The aggregator's results reader (2026-11-24T14:52Z)

The finals are on disk for three lanes; nothing reads them yet. `ingestion/providers/
betexplorer_results.py` is the reader, and it is deliberately built from the only markup
on that host this project has captured in full: the **fixtures** row — `in-match` link,
two `<span>` names, a `table-main__datetime` cell — plus a result cell.

The rules it holds:

* a row is read only with a code, exactly two club names and exactly one score-shaped
  pair; everything else is counted by reason, never guessed;
* **a clock is not a score.** `20:00` satisfies the colon shape, so a pair is a final
  only when both sides reach 25 points — an hour cannot. Forfeits (`20:0`) fall outside
  that rule and are excluded on purpose: a missed final is recoverable, an invented one
  is not;
* two score-shaped pairs are an ambiguity, not a choice;
* an undated result still reads, but carries no `played_on`, so nothing settles on it.

`--read` replays the reader against the saved `results-index.html` with **no network and
no request**, so proving it costs the host nothing. Until that replay yields rows, all
three lanes stay `source_unavailable` and this module is a hypothesis with a digest.

Provenance travels in the row: every result is labelled *aggregator result, matched by
provider code*, so no downstream record can quietly promote an aggregator's number into a
federation final.

## A note on reading Greek in a terminal (2026-11-24T14:51Z)

The Greek replay printed `"home": "ÎÎ¡ÎÎÎÎ£"` — UTF-8 bytes rendered as latin-1 by the
terminal. That is a *display* fault, not a data fault, and the distinction matters when
the question is whether a club name is intact. `row.home.encode("unicode_escape")` prints
`\u039a\u039f…` and survives any terminal, which settles it without a base64 round trip.

One name did look genuinely short in that output (`ΟΣΣΣ H HOTELS COLLECTION`). The
depth-aware span fix addresses the mechanism that would cause it — a nested span closing
early — but whether the name now reads whole is answerable only from the next replay, on
a tree that has the fix.

## Greece, the names read whole (2026-11-24T14:56Z)

The replay on the fixed tree, with the names printed as escapes so the terminal cannot
lie about them:

```
containers 7 · rows 7 · no_identity 0 · names_ambiguous 0 · with_final 0 · unresolved_date 7
ΚΟΛΟΣΣΟΣ H HOTELS COLLECTION   100 64   ΠΡΟΜΗΘΕΑΣ ΠΑΤΡΑΣ ΒΙΚΟΣ COLA
ΔΟΞΑ ΛΕΥΚΑΔΑΣ ELABET            83 87   ΠΕΡΙΣΤΕΡΙ Betsson
ΠΑΝΑΘΗΝΑΪΚΟΣ AKTOR              85 72   ΜΑΡΟΥΣΙ
```

Two things are now settled. The mojibake was the terminal, not the file — the escapes
prove it, and that costs one line rather than a base64 round trip. And the depth-aware
span fix is **proven against served bytes**: these are complete club names including
their title sponsors, where the old non-greedy span match trimmed them.

The date is the last blocker and it has a shape now: `day=None month=None` on all seven
rows, on a page whose *rendered* text demonstrably shows «Σαβ 21 Νοε - 15:00». So the day
and month live somewhere `_text` does not reach — an attribute, or a cell this reader
mistakes for markup. The next read is that cell's own escaped bytes, not another
hypothesis about them.

## A capture that taught nothing (2026-11-24T14:56Z)

The first `results-index-window.txt` spent 600 characters on the page's tab bar and pushed
the match row it existed to show out of the printed excerpt. The lead-in is now 300
characters and the tail 1900, so the code sits near the start of the file: a window that
ends before the row is a window that costs a round trip and returns nothing.

## The aggregator row, captured and read (2026-11-24T14:59Z)

```
read: 0   rows: 40   without_code: 40
```

Zero rows on all three pages, and the reason was **attribute order**. The fixtures parser
this project already trusts requires `href` immediately before `class="in-match"`; the
results page writes:

```html
<a data-test="11015187" href="/basketball/australia/wnbl-women/perth-lynx-canberra-capitals/EaZEfXoI/" class="in-match">
```

`data-test` between them, so a pattern anchored on the file's own fixtures markup matched
nothing. The page was never the problem; the assumption that one page's attribute order is
the host's contract was.

Three corrections followed from the same capture:

* the anchor is matched **whole** and its `href` read out of the tag, so order stops
  mattering;
* both clubs live **inside** that one anchor (`<span><strong>Perth Lynx W</strong></span> -
  <span>Canberra Capitals W</span>`), so names are read from the anchor's content and a
  span elsewhere in the row — a round label, a tab — is no longer a third club;
* the total is a **second unclassed link** in the next cell, not a `table-main__result`
  cell, and it can reach three digits. A pair regex bounded at two digits silently skipped
  every century: `113:90` produced no candidate at all. Bound is now three, with the same
  rule that both sides must reach 25 so no clock qualifies.

The captured row now reads `Perth Lynx W 113 - 90 Canberra Capitals W`.

What the page does **not** carry: a date per row. Results are grouped under
`<th colspan="2">6. Round</th>` headers, so every result from this page is undated. That
is a design consequence, not a defect: the date must come from the forecast the result is
matched against, and a result that cannot be dated that way cannot settle.

## Greece, the date cell as served (2026-11-24T14:59Z)

```html
<div class="esake-program-game-info one-col"><img src="/skn/clock.svg" ...>Σαβ 21 Νοε - 15:00</div>
```

The day and month are plain text in the row, and a reader given exactly this prints
`played_on = 2026-11-21`. So the remaining `unresolved_date: 7` is not about this cell's
contents — it is about what `_text` over the *whole chunk* yields, which is the next thing
to print rather than assume.

## Lithuania reads 140 finals (2026-11-24T15:0xZ)

The corrected reader on the captured NKL page:

```
read: 140   rows: 163   names_not_two: 0   score_not_one: 0
BC Olimpas Palanga 125 - 124 Omega-Tauras
Plunges Olimpas    101 -  78 Zalgiris Kaunas 2
Silute              98 -  79 Mazeikiai
```

One hundred and forty finals, every one with a code, both clubs and both totals, from a
host that refuses every score-bearing WP-JSON namespace and whose public namespace is a
calendar with no score field. The lane was dark this morning; it has a results route now.

Two properties worth naming, because they decide what wiring costs:

* **the names are the aggregator's own** ("BC Olimpas Palanga", "Zalgiris Kaunas 2"), and
  so are the fixtures this lane already forecasts from. Joining a result to a forecast is
  therefore a same-source match, not a cross-provider alias problem — which is the single
  most expensive thing this project knows how to do badly;
* **every row is undated.** Results are grouped under round headers, so the date has to
  come from the forecast being graded. A result that cannot be attached to a dated
  forecast cannot settle, and that is a wiring question rather than a parsing one.

## Greece, the date that should have matched (2026-11-24T15:0xZ)

The clock cell as served is plain text — `Σαβ 21 Νοε - 15:00` — and a reader handed
exactly that string returns `(21, 11)`, verified three ways here including from the
operator's own escaped bytes. Yet the same call on the whole row returns `(None, None)` on
a page where the same text is present. That is a contradiction, not a difficult parse, and
guessing at it would cost more than asking.

So two things changed rather than a theory: the date is now read **from the clock cell
first** and only then from the row, which both shrinks the input and stops an unrelated
day elsewhere on the page from being read as this game's; and `esake.date_probe(chunk)`
reports the cell text, the cell's matches, the row's first 160 characters and the row's
matches side by side. "The page shows a date and the reader says None" has two causes —
text never received, and a pattern that did not match text it did receive — and one call
now tells them apart.

Reading the cell first is also right on its own terms: a fixture's date belongs to the
clock cell, not to whichever day happens to appear earliest in the row's markup.

## The Greek date was never the page's fault (2026-11-24T15:1xZ)

`date_probe` answered it in one call:

```
cell_text:    'Î£Î±Î² 21 Î\x9dÎ¿Îµ - 15:00'
cell_matches: []
text_matches: [['22', 'PTS'], ['13', 'RBS'], ...]
```

Look at the second character of «Ν»: `\x9d`. That is **U+009D, a control code**, not a
Greek letter. The page was fine; the *file* holds mojibake. `esake.gr` serves UTF-8 and
declares no charset, so `requests` fell back to ISO-8859-1 — its default when a host says
nothing — and every non-ASCII character became two latin-1 characters. «Νοε» arrived as
`Î` + `U+009D` + `Î¿` + `Îµ`: four characters, one of them invisible.

Everything follows from that:

* the month table could never match, so `day=None month=None` on a page plainly showing a
  date — and the whitespace fallback could not help, because the problem was characters;
* the club names printed as `ÎÎ¡ÎÎÎ£` were **not** a terminal rendering UTF-8 as latin-1.
  The data really was mojibake, and this session said otherwise for three turns. The
  escapes looked like UTF-8 bytes because latin-1 characters, encoded as UTF-8, *are* the
  original bytes;
* `text_matches` finding `['22','PTS']` and `['13','RBS']` explains itself too: those were
  the only day+letters pairs left once the month name had dissolved.

The fix belongs at the boundary, not in any parser: `ingestion.http.decode_body_bytes`
honours a declared charset, otherwise decodes UTF-8, and falls back to latin-1 only for
bytes that are not valid UTF-8. A parser that silently re-decodes its input is how one
corrupts a record; a client that decodes correctly is how one never has to.

Existing captures stay broken until re-fetched. The one-line repair, for a body that was
decoded as latin-1 and saved as UTF-8:

```python
html = html.encode("latin-1").decode("utf-8")
```

A test now holds the whole shape: the same row read from mojibake yields totals but no
date, and read after that repair yields `2026-11-21` with the clubs intact.

## WNBL and WCBA read (2026-11-24T15:1xZ)

| lane | rows | read | header rows |
|---|---|---|---|
| WNBL | 40 | 33 | 7 |
| WCBA | 29 | 24 | 5 |
| NKL | 163 | 140 | 23 |

`without_code` is exactly the round-header rows in every table, and `score_not_one` is zero
throughout: no row that had a code and two clubs failed to yield a total. Three lanes that
could not settle this morning now read 197 finals between them.

## Greece reads seven dated finals (2026-11-24T15:2xZ)

```
containers 7 · rows 7 · with_final 7 · without_final 0 · names_ambiguous 0
unresolved_date 0 · duplicate_idgame 0 · no_identity 0

3C7521D5 2026-11-21 84-87  ΗΡΑΚΛΗΣ                    | ΑΕΚ
1439C9D6 2026-11-21 100-64 ΚΟΛΟΣΣΟΣ H HOTELS COLLECTION | ΠΡΟΜΗΘΕΑΣ ΠΑΤΡΑΣ ΒΙΚΟΣ COLA
BE8E8A17 2026-11-21 83-87  ΔΟΞΑ ΛΕΥΚΑΔΑΣ ELABET        | ΠΕΡΙΣΤΕΡΙ Betsson
95F2FDAA 2026-11-21 99-77  ΠΑΟΚ                       | ΑΡΗΣ
779E922D 2026-11-21 85-72  ΠΑΝΑΘΗΝΑΪΚΟΣ AKTOR          | ΜΑΡΟΥΣΙ
6AA68E52 2026-11-22 96-83  ΜΥΚΟΝΟΣ Betsson BC          | VIKOS ΦALCONS
C02F68B7 2026-11-22 75-81  ΚΑΡΔΙΤΣΑ ΙΑΠΩΝΙΚΗ           | ΟΛΥΜΠΙΑΚΟΣ
```

Every count is now the one the page promised, and the same run proved the encoding fix at
the source rather than on a repaired file: one request, **200 in 133,247 B**, sha256
`8a2ebdb9…`. The byte count is its own evidence — the file was 139,358 B while it was
double-encoded and is 133,247 B now, and 6,111 bytes is very close to one extra byte for
every Greek letter in the page.

Two things this does **not** establish, kept separate on purpose:

* **No settlement.** `ActualResult.reconciles()` needs Q1–Q4 (or H1/H2) summing to the
  final. This reader parses finals and cuts the hidden per-quarter row away precisely
  because it repeats both club names. The page *has* quarters — the row's collapsed text
  shows `Q1 15 23 Q2 17 19 Q3 25 20 Q4 20 7` — but that text is interleaved with player
  stat lines and does not sum to any of these finals, so reading it from collapsed text
  would be a guess.
* **No join.** The receipt's join block is `stored_games: 0, attachable: 0, unmatched: 0`:
  the lane has stored no games, so there is nothing for a final to attach to. That is the
  fixtures side, not the results side.

Hence `results_source` stays `source_unavailable`. The reader is proven; the *adapter* it
would need — quarters reconciled to a final, attached to a stored game — does not exist,
and relabelling the lane before it does would be a status string with no behaviour behind
it. The next capture is the hidden row's own markup, ~900 escaped characters.

## A final is not a settlement record (2026-11-24T15:3xZ)

Worth stating plainly, because 197 parsed finals make it easy to skip past:

`ActualResult.reconciles()` requires Q1–Q4 (or H1/H2) that sum to the final. It refuses
anything else — `period kind unknown is not reconcilable`, `expected exactly four quarter
values per team`. So a results **index**, which carries finals and nothing else, cannot
settle a forecast in this corpus. Not for NKL, not for WNBL, not for WCBA. The 197 rows are
cross-check material and an evidence trail; they are not settlement.

That leaves exactly two places a reconcilable record can come from:

* the **federation** page, which for Greece carries the quarters in its hidden row — the
  reason that lane's adapter is a quarter parser and not a totals parser;
* the **aggregator's per-match page**, which is where period scores live on that host.

The second is one request to answer, and `scripts/probe_betexplorer_results.py --match N`
now answers it: it takes the first N codes from the captured index, fetches each match page
at the tool's own spacing, saves the body verbatim, and reports `period_markers` — a count
of `Q1`/`1Q`/`OT`-shaped tokens. Zero means the host does not publish what reconciliation
needs, whatever its results table says; a positive count means the lane can settle from one
request per game.

That is deliberately one request per league and not one per game: the question is whether
this host carries period scores at all, and answering it with 140 fetches before knowing
would be the impatience this project keeps out of its probes.

## Greece's hidden row is player cards before it is quarters (2026-11-24T15:3xZ)

The 900 characters after `esake-program-game-hidden-row` are MVP cards —
`<div class="mvp-player"><div class="mvp-player-name"><a href="/el/action/EsakeplayerView?
idplayer=9DBC8FE2&mode="><span>FORRESTER</span><br/>JAKOB</a></div><div class="mvp-player-
stats">22<span>PTS</span></div>…` — and the per-quarter block sits further in. So the
window was aimed at the right element and too early inside it; the next read targets `Q1`
inside the hidden row rather than the row's start.

Nothing in this changes the finals: seven games, seven dates, both clubs each. It changes
only what the next 900 characters should be centred on.

## Greece reconciles (2026-11-24T15:3xZ)

The quarter block, captured:

```html
<table class="esake-program-game-half-scores">
<tr><td>15</td><td>Q1</td><td>15</td></tr>   <tr><td>23</td><td>Q2</td><td>17</td></tr>
<tr><td>19</td><td>Q3</td><td>25</td></tr>   <tr><td>20</td><td>Q4</td><td>20</td></tr>
<tr><td> 7</td><td>ΟΦ1</td><td>10</td></tr>
</table>
```

Add it up and the mystery of a 84-87 game with quarters that sum to 77-77 disappears: the
overtime is «ΟΦ1», 7-10. A totals-only reader had no way to know an extra period existed;
a reader that ignores the period table would have written a record that failed its own
reconciliation on every overtime game in the league.

`ResultRow.periods` now carries it, and `to_actual_result()` builds the `ActualResult` — or
returns `None` when there is no complete period table, because a total is not a settlement
record. Three tests hold the edges: the captured row reconciles `(True, '')`; a row without
a period table yields no record at all; and a table that does not add up reaches
`reconciles()` and is refused *there*, by the same check every other provider answers to.

## The nav tab that looked like a match (2026-11-24T15:3xZ)

`--match 1` reported `event_id: "fixtures"` and opened
`https://www.betexplorer.com/basketball/lithuania/nkl/fixtures/` — 183,036 bytes of
captured page, `period_markers: 0`, and none of it a match. `fixtures` is eight lowercase
letters and the code pattern asks for eight alphanumerics, so the tab bar matched.

Two corrections: a code must mix case or contain a digit (a nav word never does), and the
URL is now taken from the index's own `href` rather than rebuilt from a code — so a match
page is opened at the address the host published for it, and `match_url` survives only as
a fallback.

The three `period_markers: 0` results in that run are therefore **not** evidence about
period scores. They are evidence that the tool opened the wrong pages, which is worth
recording precisely because it is the second time today a plausible-looking success
(`read: 0` for three leagues, `event_id: fixtures`) carried no information at all.

## The href was relative (2026-11-24T15:2xZ)

`--match 1` then failed on all three lanes with one message:

```
HTTP 0 for /basketball/lithuania/nkl/alytaus-patriotai-neptunas/nRzm7OsT/:
Invalid URL: No scheme supplied.
```

The index publishes its hrefs as paths, and taking the href verbatim -- which was the
point of the previous fix, so a page opens at the address the host gave -- handed the
client `/basketball/...` with no scheme. `absolute_url()` now prefixes the site, because a
path from this host is a path on this host.

Three tool faults in one afternoon, and each was caught by the receipt rather than by a
review: `read: 0` for three leagues (attribute order), `event_id: fixtures` (a nav word is
eight letters), and now `HTTP 0 ... No scheme supplied` (a relative href). The receipts are
the reason none of them cost more than one run each -- a fetch that fails and says why is
worth more than a fetch that succeeds and says nothing.

## The match page is not a table (2026-11-24T15:24Z)

Two of three match pages were captured — NKL 110,391 B, WCBA 109,961 B — and both report:

```
td_cells: 0    integer_cells: 0    span_cells: 32    period_markers: 0
```

A hundred kilobytes with **no table cells at all**. So the scoreboard on this host is not a
`<tr>/<td>` structure and not a `Q1…Q4` label set; `td_cells` and `period_markers` were the
wrong instruments, and `period_markers: 0` here says only that this page does not use the
labels I chose to count. That is the third time today a zero needed reading carefully, and
the lesson is the same each time: a count measures the counter's vocabulary, not the page.

The WNBL request was **429**, six seconds after the NKL one. That is the host naming its own
limit, so the probe now waits ten seconds between match pages and **stops on 429** instead of
working through the rest of its list. Retrying would spend the budget arguing with a host
that has already answered.

What is still unknown is what those 110 KB *are*. The next read is a window of the saved
NKL body around whatever carries the score — one command, no request, on a file that is
already on disk.

## The match page was an age gate (2026-11-24T15:2xZ)

`scoreboard_mentions: 0`, and the 110 KB is this:

```html
<h1 id="ageTitle">Verificação de idade obrigatória</h1>
… Você deve ter pelo menos 18 anos para acessar este site …
```

An eighteen-plus modal in Brazilian Portuguese, wrapped in ad banners. Not a scoreboard, not
a table, not a match page — a consent gate. Every shape signal I had said "captured",
because 110 KB of modal is 110 KB.

Two readings are possible and the receipts cannot separate them:

* match pages on this host are gated by design, and results pages are not; or
* the host began serving the gate **after our burst** — NKL captured, WNBL 429, WCBA
  captured — in which case it is a soft block answering our request pattern rather than a
  property of the page.

The polite response is the same either way, so nothing is attempted past it: the probe now
labels such a body `status: "gated"` with the marker that matched, saves it as evidence, and
counts it in `gated` rather than in `captured`. A gate that reports as a capture is the
fourth version of today's recurring failure — a number that looks like an answer.

## Where the aggregator lane now stands, honestly

| what | state |
|---|---|
| results index (3 lanes) | **197 finals read** — codes, clubs, totals, no dates |
| match pages | **gated**; period scores unread |
| settlement | **impossible from what is read**, since a final does not reconcile |

So the aggregator is, today, a **verification** source and not a settlement source. That is
a scope statement, not a failure: 197 finals will catch a settlement error from any other
source, and they cost ten requests in total across three leagues.

Whether to go further is an operator decision and it has two branches: pass the age gate
(a self-declaration of 18+, which is their consent mechanism rather than a robots rule), or
leave the lane as verification and take settlement from federations. This project has
already recorded one such decision — the Greek terms — as the operator's, in the lane's own
`source_detail`, rather than making it quietly.

## The age gate, authorised (2026-11-24T15:3xZ)

The operator states they meet the site's 18+ requirement, so the self-declaration is
theirs. That authority is now carried in every capture as
`age_gate_authorisation`, for the same reason the Greek terms were written into the lane's
own registry entry: a gate passed without a named authority is exactly the kind of quiet
decision this project keeps out of its pipelines.

It is explicitly **not** a robots exception. Path-only URLs, no query shapes, no per-game
burst, and the spacing rules are unchanged. What changes is only the consent the site asks
a person for, which a person has now given.

### What passing the gate actually requires

Not a guess. `--gate-hints PATH` reads a saved gate page and reports what its accept button
writes — `document.cookie=…`, `localStorage.setItem(…)` — because the way to pass a gate is
to reproduce what it records, and the way to fail is to try plausible names until the host
gets annoyed. No request is made to answer this.

The output keeps two things apart that a single list would merge: `hints` are writes (a
cookie value, a storage key), `mentions` are identifiers that merely contain "age" — a CSS
class, an element id, a vendor name. Only the first kind is a way in; a list that mixed them
would invite replaying `age__modal`. `replayable` says whether an HTTP client can act on the
hints at all.

`--gate-accept` then replays exactly that cookie, once. Two requests per gated page, never a
loop, and the receipt says which happened:

| what the gate recorded | what the capture says |
|---|---|
| a cookie | `captured`, `gate_passed_by: ["age_verified"]`, `requests_made: 2` |
| only localStorage | `gated`, `… no cookie to replay, storage keys ['ageVerified'] -- needs a browser` |
| a cookie, still gated | `gated`, `… still gated after replaying 1 cookie(s)` |
| cookie writes, none near the gate | `gated`, `… no cookie written near the accept control (unrelated writes: ['lang'])` |
| nothing, not asked | `gated`, `requests_made: 1`, `gate_accept: false` |

The third row is the one worth keeping. A gate that answers the cookie and shows the modal
anyway is not a consent gate at all — it is the host refusing, and the honest reading is the
same 429 lesson: stop, and write down that it answered.

What is *not* claimed: that passing the gate yields period scores. The 110 KB bodies carry a
0 in every shape signal, so the modal may be incidental to a scoreboard that is drawn by a
script we have not read yet. The next run answers that; this change only makes the question
askable.

### The gate answered, and it was not a cookie (2026-11-24T15:45Z)

The first real receipt, run with the operator's authority and the cooldown in place:

```
hints:    ["lang=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/", "mode"]
replay:   110,501 bytes, sha 8a1e9f05…, "still gated after replaying 1 cookie(s)"
```

`lang=` with an epoch expiry is a language reset, and a bare `mode` is not a cookie at all.
Both are writes on that page; neither is consent. Replaying them bought a second identical
age modal, which is the second row of the table above working as intended — the host answered
the cookie and showed the gate anyway.

So the extractor was the thing that lied, not the host: `document.cookie = …` matched, and a
match is not the same as "the accept button wrote this". `gate_cookie_writes` now measures
each write's distance to the gate's own markup (its markers, and ids like `ageYes`); only
writes within 3,000 characters count as hints, everything else is listed with its distance as
`unrelated_writes`, and `--gate-hints` exits non-zero when nothing replayable is near the
gate. The failure this removes is precise: sending `lang=` to a host that then throttles us,
and recording the throttle as policy.

One incidental datum worth keeping: the same page was 110,391 bytes with sha `7cf27f25…` a
session earlier and 110,501 with sha `8a1e9f05…` now. The body varies per request, so byte
size and digest are evidence that *a* body was saved, never that it is the same body.

Where the consent actually lives is still unknown, and guessing again is not the plan.
`--gate-script` reads a saved gate page offline and prints what the accept control does —
inline handler excerpts, the external bundle list, how many inline blocks there are. If the
answer is a bundle, the gate cannot be passed from a saved body at all, and the honest
conclusion for these lanes is the one already written above: verification only, no settlement.

### Not being throttled, in one mechanism

Three match pages a few seconds apart drew a 429, and the pause inside a run could not have
helped: the *runs* were the burst. So the cooldown is a stamp on disk —
`results_probe/.host-cooldown` — holding the earliest instant the next run may act:

| run | asks at | waits | acts at |
|---|---|---|---|
| 1 | 1000 | 0 | 1000 |
| 2 | 1005 | 25 | 1030 |
| 3 | 1050 | 10 | 1060 |

The gap between actions stays 30 seconds, and a wait cannot be banked and spent early. Two
independent runs invoked a second apart now behave like one polite client, because the
stamp is the only place they meet.

Three other behaviours, all already in place and all cheap: a 429 ends the run rather than
retrying into a host that has answered; match pages are spaced ten seconds apart *inside* a
run; and `--match N` caps at twelve pages so a typo cannot turn into a crawl.
