# Provider matrix — acquisition sources (baseline 2026-11-06)

Evidence labels: `LIVE_VERIFIED` (observed response), `FIXTURE_VERIFIED`
(committed fixture), `DOCUMENTATION_ONLY` (docs/terms only). Revalidate at
implementation time; labels in smoke output update automatically for live
runs.

| Provider | Role | Endpoints used | Verification | Leagues/periods | Quota / rate limits | Auth | Licence / display | Freshness policy | Gaps & fallback |
|---|---|---|---|---|---|---|---|---|---|
| ESPN core/site APIs | NBA fixtures, finals (periods/OT), standings, BPI predictions | `site.api.espn.com/.../scoreboard?dates=`, `/summary?event=`, `/standings`, `sports.core.api.espn.com/.../predictor` | **LIVE_VERIFIED** 2026-11-06 (BPI 9/9 complete on slate; golden 401909938 exact) | NBA (WNBA/NCAA: probabilities only, no expected scores) | No documented SLA; undocumented endpoints | Keyless | Public endpoints; no SLA | BPI `lastModified` → `source_modified_utc`; fetch time → `observed_utc`; 12 h window default | Core `events?dates=` discovery returned **0 events** → site scoreboard is the discovery path. BPI orientation via team-ref IDs; unreconcilable refs → `IdentityError` quarantine |
| Hardwood (roni-altshuler/nba_predictor) | Independent NBA prediction fallback + shadow model | `raw.githubusercontent.com/.../game_forecasts.json` | **LIVE_VERIFIED** 2026-11-06 (1077 games, ingestion ok) | NBA full game | No service SLA; artifact is a public file | Keyless | MIT | `generated_at` within 12 h; `ingestion.status == ok`; `trained_through` sanity | Joins by ESPN game id + team/start validation; stale/mismatched → None + gap; never averaged with ESPN |
| Official EuroLeague | Schedule, facts, actual quarters + OT | `api-live.euroleague.net/v2/competitions/E/seasons/E2026/games`, `live.euroleague.net/api/Boxscore?gamecode=&seasoncode=` | **LIVE_VERIFIED** 2026-11-06 (380 games; OT game 91 RMB 112–115 PAO via `Extra1`) | EuroLeague full game + quarters | No documented quota | Keyless | Official public API | `played`/`utcDate` from schedule; boxscore at results time | `ByQuarter` = individual, `EndOfQuarter` = **cumulative** (never written as quarters); 59/98 played reconciled on live smoke, rest quarantined |
| Predixsport | Authenticated predictions (esp. EuroLeague) | `api.predixsport.com/v1/predictions/{nba\|euroleague}/upcoming`, `/v1/predictions/{league}/{match_id}` | **DOCUMENTATION_ONLY** (unauth call 2026-11-06 returned "Missing Authorization header"; fixtures synthetic) | NBA/EuroLeague (claimed) | 200 shared calls/calendar month, personal non-commercial | `Authorization: Bearer` (env `PREDIXSPORT_API_KEY`) | Free tier: personal/non-commercial | One list call per refresh + cached details | Disabled by default; detail scores `(total±diff)/2` (diff = home−away); winner from probability only; `produced_by` saved |
| TheRundown | NBA full-game moneyline + total | `therundown.io/api/v1/sports/{sport}/events/{date}` (markets 1,2,3) | **DOCUMENTATION_ONLY** | NBA full game only | 20k datapoints/day, 200k/month, 1 req/s, 5-min delayed; event cost `2*events+prices` | `THERUNDOWN_API_KEY` | Free: internal use only; no end-user display | Pre-match only (phase guard) | **No quarter markets on Free — adapter has no quarter code path + books.py hard guard;** NBA sport id constant requires verification before enabling |
| SportsGameOdds | Preferred candidate for NBA Q1–Q4 totals | `api.sportsgameodds.com/events?leagueID=NBA...`, `/account/usage` | **DOCUMENTATION_ONLY** (oddID pattern documented: `points-all-{game,1q,2q,3q,4q}-ou-{over,under}`) | NBA full game + quarters (claimed) | 2,500 objects/30 days, 10 req/min, 8 leagues/9 books; card at signup | `x-api-key` (env `SPORTSGAMEODDS_API_KEY`) | Free tier restrictions per FAQ | Pre-match only; main-line preference | oddIDs confirmed only through `/markets`; missing line ⇒ snapshot unusable (line=None); partial quarters allowed, never backfilled |
| The Odds API | Targeted Q1–Q4/full-game fallback | `api.the-odds-api.com/v4/sports/{sport}/odds/`, `/events/{id}/odds` (`totals_q1..q4`) | **DOCUMENTATION_ONLY** | NBA (sport key `basketball_nba` to verify); EuroLeague key unknown | ~1 credit per market per region per event; `x-requests-remaining/used` captured | `apiKey` query param (env `THE_ODDS_API_KEY`) | Per current terms at enablement | Pre-match via commence_time | Narrow region/book; per-event calls only where coverage justifies |
| Manual CSV | Quarter-line fallback | local file | n/a (operator input) | Any quarter | n/a | none | n/a | Timestamp required per row | Rows missing identity/book/period/line/observed_utc are rejected — never anonymous stale data |

## Golden mapping (tested)

ESPN event `401909938` (DEN @ IND, lastModified 2026-11-06T21:51Z, stats on
`awayTeam`) renders exactly `23.1 - 76.9`, `2`, `111-120`, `230.6`. Stats on
either side orient through team-ref IDs; unreconcilable → quarantine.

## Semantics enforced by tests (selection)

- Full-game total **line value** → upcoming `FT Score`; Q1–Q4 lines map
  individually; moneyline → decimal price of the predicted team.
- A model expected total (`Avg`) is never a book line; no full-game/half
  value is ever used to synthesize a quarter total.
- A missing refresh never erases a valid existing manual/book value
  (sheet-aware merge; manual Odds preserved too).
- TheRundown Free cannot fill quarter cells (adapter + `books.py` guard).
