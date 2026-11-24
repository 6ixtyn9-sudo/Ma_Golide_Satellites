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
can grant ourselves**: whether a paper-record research pipeline with attribution
sits inside "personal, non-commercial" is an operator decision, and it is the
only thing standing between this lane and settlement.

What the bytes already support: stable `idgame` identity per game
(`EsakegameView?idgame=1439C9D6` refetched identically), finals in
`esake-program-game-final-score` blocks, a server-honoured date/matchday window
(`series=1`/`series=2`/unsuffixed all differ), and 10/10 joins against the
BetExplorer baseline by score pair + day (BetExplorer is the naming baseline
only — its ToS prohibits scraping, and it is never the route).

**To build the adapter, this repo needs the two captured HTML files** (one
`EsakeResults` series page and one `EsakegameView` page). A parser written from
the report's snippets alone would be a guess dressed as code, and would be
untestable without egress.

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

## The Japan mismatch this session *did* find

The probe surfaced a real, load-bearing bug: this lane reads the provider's
**top** division (`tab=1`, now branded B.LEAGUE PREMIER) while the board slug it
prices against, `japan/b-league-one`, is the **second** division (B.LEAGUE ONE —
the captured tab=2 slate matched that board page game-for-game). Different
clubs, so no name can ever agree and the direct-price fallback is dead by
construction for this lane.

Fixing it is a scope decision, not a patch: either this lane is B.ONE (point the
provider at `tab=2`), or it stays PREMIER and needs its own board slug. Both are
one-line changes with different consequences for what the lane claims to cover,
so the constant is annotated and the choice is left open.

## Open items

1. **Greece:** operator decision on the ToS reading; then upload the two captured
   pages so the adapter can be written and fixture-tested.
2. **WNBL:** request the missing report — the probe log suggests it got furthest
   of all five.
3. **Japan:** choose the division.
4. **China / Lithuania / Israel:** unchanged until a terms document exists; no
   adapter will be written against imagined markup or an unlicensed payload.
