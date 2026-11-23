# Capture provenance for the committed provider fixtures

Everything here comes from one probe package produced **outside this repository**
by the operator's external agent on 2026-11-22 and delivered as text: each
artifact carried its own byte count and sha256, and every body was re-hashed on
arrival before it was committed. 24 of the 28 transmitted artifacts reproduced
byte-exactly; the four that did not are excluded rather than "fixed up" (see the
bottom of this page).

This file exists so the evidence class of a fixture is checkable rather than
asserted. `tests/test_capture_fidelity.py` re-hashes the two whole bodies on
every run: editing a capture now fails a test instead of quietly turning a
capture back into a shape.

## Committed, byte-exact whole bodies

| Fixture | Provider route | Bytes | sha256 |
| --- | --- | --- | --- |
| `lba/championships_2026.json` | `legabasket.it/api/v1/championships?cs_id=1&items=1000&s=2026` | 867 | `e6d648fb4f57e5a44c6fc3053c974402d103866a95291d758aba6f3a266e70cc` |
| `lba/calendar_602.json` | `legabasket.it/api/v1/championships/get-championships-calendar-by-id?id=602` | 11636 | `9a0e5d28c18eac20d06aeef94b9ca354e0597319301b2ba48628c5853c155101` |

The calendar is the league's current focus day at capture time: eight rows
spanning 2026-11-21 and 2026-11-22, with `filters.days` marking round 9 (id
4494) as focused rather than any single date.

## Committed, container written here around verbatim captured rows

The provider's play-by-play body is 314 KB and its season board is 159 KB; the
package transmitted verbatim byte excerpts instead of whole bodies. The row
objects below are copied unchanged from those excerpts — only the JSON envelope
(`{"pbp": {"actions": [...]}}`, `{"games": [...]}`, `{"data": [...]}`) and
the whitespace of the re-serialisation are this repository's.

| Fixture | Captured source | Rows used | Bytes (as committed) | sha256 (as committed) |
| --- | --- | --- | --- | --- |
| `lba/pbp_25426.json` | period-end actions of `get-championships-matches-play-by-play?id=25426` (Trento 100-86 Napoli) | 4 | 2389 | `6ba45b151318d157fb77fa1a86982ed8...` |
| `lba/pbp_25422_ot.json` | same route, `id=25422` (Cantù 92-106 Tortona, after overtime) | 5 | 2976 | `f4c0d57b6ed61e75e1e319e585ecf6b1...` |
| `israel_ibl/board5_captured_rows.json` | `basket.co.il/ws/ws.asmx/Games?cYear=2027&board_id=5...`, four rows | 4 | 4443 | `90804be6d6a85b0ce4f19fe0955e6058...` |
| `eurocup/U2026_ot_game.json` | `api-live.euroleague.net/v2/competitions/U/seasons/U2026/games`, the first row with `extraPeriods` | 1 | 3021 | `5f96e379955eaa3c546a974280fa16f5...` |

`lba/pbp_synthetic_noncontiguous_ot.json` is **not** a capture: it is a
hand-built refusal case (regulation anchors plus a lone period 6) and is named
accordingly.

## Not committed

Oversized original bodies that were never transmissible whole (retained only by
the external agent): the four LBA play-by-play bodies (313–372 KB), the ACB
2027 calendar (329 KB), the Israel board (159 KB), the EuroCup `U2026` schedule
(457 KB), the BCL GDAP response (309 KB), and three BetExplorer pages
(110–128 KB).

Four transmitted artifacts did **not** reproduce their stated hash on arrival and
are therefore not committed as captures: `acb/matchheader_401.txt` (366/372
bytes), `betexplorer/matchodds_404.txt` (910/917), `israel/board5_rows.slice.txt`
(4723/4725 — the four row objects inside it were intact and are what the Israel
fixture uses), and `nbl/sdist_nbl_listing.txt` (204/206). Their claims are
recorded in `docs/PROVIDERS.md` as *reported by the probe*, not as
repository-verified bytes.

## Shape fixtures, not captures (2026-11-23 source policy)

The 2026-11-23 roadmap permits scraped and undocumented public sources for
non-NBA expansion lanes on condition that a lane is promoted on a **live
receipt**, not on a fixture. Two classes therefore exist and must not be
confused:

| class | what it proves | promotion authority |
| --- | --- | --- |
| `capture` (the files above) | the provider served this exact body | may justify wiring a parser |
| `shape fixture` | the parser handles the documented column contract | **no** — a live receipt is still required |

`kbl/match_list_shape.json` and `kbl/team_list_shape.json` are **shape
fixtures**. They are written to the column contract documented by the
SportsDataverse `kbl` module (MIT, read 2026-11-23) — `gmkey`, `game_date`,
`game_start`, team codes and names, `home_score`/`away_score`, `is_started`,
`is_ended`, `playing_quarter` — and no Ma Golide first-party body exists for
that host yet.

What stands in for a capture is the run-time receipt: `kbl.last_probe()`
records the URL, byte count, sha256, envelope shape and row keys of the body
the adapter actually read, and `Pipeline._discover_games` stores it as
`provider_health`. An empty read is recorded as `CoverageState.MISSING` rather
than `OK`, so a lane that stops returning rows cannot look healthy.

`acb/calendar_2027.json` and `acb/match_headers_2027.json` are likewise **shape
fixtures** written to the documented ACB field contract — the calendar's match
ids, dates and scores, and the header's `teams`, `matchStatus`, current score
and `quarterScores[]`. The lane had passed those parser tests while reading
zero live games, which is how the missing `X-APIKEY` header stayed invisible for
728 runs. What makes the fix checkable is `acb.last_probe()`: every run records
the calendar URL, byte count, sha256, round count and row count it actually saw,
persisted as `provider_health`, with an empty read recorded as
`CoverageState.MISSING` rather than `OK`.

## Real capture: Turkey TBF daily matches (2026-11-23)

`tbf/captures/get-daily-matches-2026-10-11.json` is a **vendored real response
body**, not a shape fixture — a different evidence class from the KBL and ACB
files above.

- Source: `sportsdataverse/sportsdataverse-py`,
  `tests/fixtures/tbf/captures/Match__get-daily-matches.json` (MIT), vendored
  verbatim on 2026-11-23.
- Route: `GET miniappapi.tbf.org.tr/webapi-service/api/Match/get-daily-matches?match_date=2026-10-11`.
- Contents: the federation's envelope `{data, isSuccess, message,
  problemDetails}` with three grouped competitions — `kbsl`, `tbl`, `tkbl` —
  and rows carrying `matchId`, `activityCode`, naive-local `matchDate`,
  `homeTeam`/`awayTeam` ids, names and scores, `matchStatusId` 1 ("Oynandı"),
  `week`, `venueName` and `hasUzatma`.
- What it proves: the envelope, the row keys, the competition grouping, the
  local-time format and the played-status marker are read from a body the host
  actually served.
- What it does **not** prove: the body carries no `bsl` row, so the top flight's
  `activityCode` string is taken from the documented contract rather than
  observed. The lane's BSL filter is therefore checked for *behaviour* (it keeps
  BSL and nothing else) but the literal spelling is validated at run time by
  `tbf.last_probe()`, which records every competition code the daily body
  actually returned. If the spelling differs, the probe shows it on the first
  run and the lane reports an empty slate instead of silently borrowing TBL or
  KBSL games.
- Personal fields: none in this body (team and venue names only).

## Real capture: Korea KBL match detail (2026-11-23)

`kbl/captures/match__gmkey__S47G01N49.json` is a **vendored real response
body**, upgrading the KBL lane from shape-fixture-only evidence.

- Source: `sportsdataverse/sportsdataverse-py`,
  `tests/fixtures/kbl/captures/match__gmkey__S47G01N49.json` (MIT), vendored
  verbatim on 2026-11-23.
- Route: `GET api.kbl.or.kr/match/S47G01N49` — 2025-11-01, Goyang Sono v
  Changwon LG.
- Contents: `{game, teams, teamrecords}`; `game` carries `gmkey`, `gameDate`,
  `gameStart`, `isStarted`/`isEnded`, `playingQuarter`, team codes and names;
  `teamrecords.{home,away}` carries `scoreq1..scoreq4` and `scoreeq[]`.
- What it proves: the quarter split that the **list** route does not carry is
  real and parseable, so a settled KBL row can reconcile periods after all —
  quarter markets are gradeable from this route, not blocked forever.
- What it does **not** prove: the list route (`/match/list`) still has no
  vendored body, so the daily discovery shape remains contract-referenced and is
  checked at run time by `kbl.last_probe()`.

## Capture excerpts: Japan B.League (2026-11-23, out-of-sandbox probe)

Neither host is reachable from this repository's sandbox, so these two files
came from a probe run elsewhere and are labelled by what they actually are.

`bleague/captures/schedule-2026-10-04-excerpt.html`
- Source: `GET https://www.bleague.jp/schedule/?tab=1&year=2026&mon=10&day=4`,
  HTTP 200, returned to the probe on 2026-11-23T12:02:12Z (162,511 bytes).
- Class: **capture excerpt**. One `<li class="list-item">` game block is quoted
  verbatim; the `<ul id="schedule-b1">` wrapper and the closing `</li>` are
  synthesised because the excerpt stopped at `</a>`. It is not the full page.
- What it proves: the slate is server-rendered HTML, the block structure, the
  class names for names/scores/status/tip, and the honoured query spelling
  (`mon`, not `month` — the latter is ignored by the host).
- What it does not prove: multi-game pagination, live statuses beyond `FINAL`,
  or any markup outside this one block.

`bleague/captures/context-506418-113.json`
- Source: `GET https://b-league.s3.amazonaws.com/web_json/v2_genius_contexts/506418/latestid`
  → `113`, then `.../506418/113.json`, HTTP 200 (31,613 bytes), same probe run.
- Class: **partial capture**. The `Game` object is verbatim; its sibling keys
  (`PlayByPlays`, `HomeBoxscores`, `AwayBoxscores`, `Summaries`) are omitted
  because this adapter reads only `Game`.
- What it proves: club ids, long/short names, `.NET` `/Date(ms+0900)/` tip form,
  totals and `HomeTeamScore01..04` per quarter.
- What it does not prove: overtime encoding. `MaxPeriod` is documented as the OT
  marker but no OT body was captured, so an unreconciling split falls back to
  totals and records a gap rather than being interpreted.

The documented `web_json/contexts/{key}` tree was also tried and refused with a
403 `AccessDenied`; `v2_genius_contexts` is the keyless path this lane reads.

## Real capture: BCL GDAP season index (2026-11-23, out-of-sandbox probe)

`bcl/captures/gdap-games-209123-row.json` is one **verbatim row** of a captured
response.

- Source: `GET https://digital-api.fiba.basketball/hapi/getgdapgamesbycompetitionid?gdapCompetitionId=209123`,
  HTTP 200, 23,789 bytes, returned 2026-11-23T12:11:54Z, header
  `Ocp-Apim-Subscription-Key` set to the value fiba.basketball ships in its own
  Next.js bootstrap JSON (public, not account-scoped; the repo already holds it
  as `PUBLIC_FALLBACK_SUBSCRIPTION_KEY`).
- Class: **partial capture** — the first row of a 136-row season index. The
  other 135 rows, and any boxscore/detail route, are not vendored.
- What it proves: the array-shaped response, the 40 row fields including
  `gameDateTimeUTC` (offset-less UTC), `teamA`/`teamB` identity and codes,
  `teamAScore`/`teamBScore`, `statusCode`/`gameResultStatusCode`/`liveGameStatus`
  and the nested `round`/`competition` objects.
- What it does not prove: scheduled rows (no unplayed row was quoted — tests
  derive that state by mutating this row), bracket placeholders with null teams,
  or any period split. Since the index carries no quarter scores and no boxscore
  route is captured, settled BCL rows stay totals-only.
