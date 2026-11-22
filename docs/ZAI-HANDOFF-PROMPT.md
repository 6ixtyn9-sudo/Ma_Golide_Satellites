# Handoff prompt for z.ai — remaining league adapters

Copy everything below the horizontal rule into z.ai as a single prompt. It is
self-contained: it states the repository, the standing constraints, the code
that must be produced, the things Ma Golide could not confirm, and the single
markdown file the answer must be delivered in.

---

## Your task

You are extending **`6ixtyn9-sudo/Ma_Golide_Satellites`** (Python 3.11, pytest, ruff), a production paper-betting pipeline for basketball. `main` is at commit `73ff50f`. Read the repository first; do not guess at its conventions.

Ma Golide already has canonical schedule/result lanes for: NBA, WNBA, NCAA M/W, EuroLeague, **EuroCup**, **Spain ACB**, ABA League, CBA, EuroLeague Women, NBA G League. Your job is to build the **remaining eleven** and to **verify the things Ma Golide could not verify** because its sandbox had no outbound access to those hosts.

You DO have network access. Use it. Every claim you make must be backed by a probe you actually ran, with the URL, HTTP status, and the field names you saw. If a probe fails, say so and say what that blocks — never fill the gap with a plausible-looking endpoint.

**If the repository is not visible to you** (it may be private), do not invent its conventions: everything you strictly need is in this prompt, and for anything else ask for the specific file — the three reference adapters named below are the ones that matter — instead of guessing. A verified probe beats a guess, and a stated gap beats a fabricated adapter, in every section of the deliverable.

**Deliver everything as ONE markdown file.** No other files. See "Output format".

## Non-negotiable project rules

1. **No fabricated capability.** A league only leaves `source_unavailable` in `ingestion/league_registry.py` when its adapter exists AND its tests pass. A registry entry without a working adapter is forbidden.
2. **No changes to tickets, staking, ROI, pricing, prediction combination, calibration, or gate decisions.** You are adding acquisition only. If a change would alter what gets bet, stop.
3. **No new state/cursor subsystem.** History cursors reuse `provider_health` keyed by `(provider, league)` — the existing `_history_cursor` / `_record_history_cursor` helpers. No new registry, no new file layout for state.
4. **Fail closed.** 429/5xx/timeouts/network errors are transient: never advance a cursor, record a coverage gap. Only HTTP 404 means "this resource does not exist" and may advance past it.
5. **Never invent a period, a final, a price, or a team identity.** A body that cannot be reconciled becomes a coverage gap, not a zero-filled result.
6. **Docs may only describe what is implemented and tested.** No roadmap claims dressed as facts.
7. **Hermetic tests only.** No network in tests; fixtures live in `tests/fixtures/<source>/`. Label every non-captured fixture as shaped-from-contract, not recorded.

## Architecture you must follow

Key files:

- `ingestion/league_registry.py` — `LeagueSpec(league, label, state_key, timezone, acquisition_timezone, schedule_source, results_source, gate_schedule_source, price_sources=(), source_detail="")`. `book_providers` always appends `betexplorer_direct` as the universal fallback.
- `ingestion/models.py` — `Game`, `ActualResult`, `GameStatus`, `PeriodKind`.
- `ingestion/pipeline.py` — dispatch points you must extend (all of them, for each lane):
  - `_discover_games()` — per-league discovery branch.
  - official results dispatch (the `if league is League.X: return self._x_finals(...)` chain).
  - `backfill()` — the allow-list `set` of leagues, plus the `deep` dispatch chain (`_backfill_euroleague_history`, `_backfill_aba_liga_history`, …).
  - `smoke()` — per-league string branch (note: this compares `league.value`, e.g. `"SPAIN_ACB"`, not the enum).
- `ingestion/adaptive_gate.py::fetch_schedule` — add a branch per new `schedule_source`. A season-scoped calendar has no date-bounded probe: raise `GateUncertain("<League> has no verified date-bounded gate source")`, exactly as EuroLeague and ACB do.
- `ingestion/providers/*.py` — one module per source.

**Study these three completed adapters before writing anything. Copy their discipline, not their code:**

- `ingestion/providers/aba_liga.py` — HTML, season calendar, count-only reconciliation receipts, explicit "no first-party capture" provenance.
- `ingestion/providers/euroleague.py` — JSON, season-scoped, competition letter derived from the season code (`E` EuroLeague, `U` EuroCup), bare-year/unknown-letter codes refused before a request is built.
- `ingestion/providers/acb.py` — JSON, season calendar (identity-free) + per-game header for names/status/finals/`quarterScores[]`.

### Adapter interface the pipeline expects

```python
def discover_games(date_iso: str | None, fetch: FetchFn, ...) -> list[Game]: ...
def fetch_game_result(game: Game, fetch: FetchFn) -> Game: ...
def season_results(season_key, fetch: FetchFn, failures: list | None = None, limit: int | None = None) -> list[Game]: ...
def last_gaps() -> list[dict]: ...          # identity/period problems from the last call
def smoke(fetch: FetchFn, ...) -> dict: ...  # never raises; returns counts + failures
```

`Game` rules that will bite you:

- `scheduled_utc` must be timezone-aware ISO in UTC; `Game.__post_init__` rejects naive values. If the provider serves local wall-clock without an offset, **raise** — do not assume a timezone.
- `game_id` must be stable and prefixed (`ACB-105401`, `U2026-97`), `season` is a `str`, provider ids go in `provider_ids`.
- `home_name`/`away_name` are canonicalised with `ingestion.identity.canonical_team_name`; identical home/away names are an `IdentityError`.

`ActualResult` rules:

- `period_kind=PeriodKind.QUARTERS` needs **exactly four** regulation values per side plus `home_final`/`away_final`; `reconciles()` checks the sum.
- Overtime: sum all OT periods into `home_overtime`/`away_overtime`, set `overtime_periods`. Non-contiguous OT indices are an error.
- Halves-only competitions use `PeriodKind.HALVES` (two values).
- If the body has fewer than four quarters, raise/record a gap — a partial body is not a result.

### Registry changes and the tests they break

When a lane leaves `source_unavailable`, these tests must be updated in the same change:

- `tests/test_basketball_expansion.py::test_final_expansion_is_typed_and_registry_backed_everywhere` — add an explicit branch for your lane asserting `(schedule_source, results_source, gate_schedule_source)`.
- `tests/test_basketball_expansion.py::test_adaptive_gate_schedules_every_registry_lane_and_fails_open_for_unavailable_sources` — the "unavailable" exemplar is now `FRANCE_LNB_PRO_A`; keep it pointed at a lane that is still unavailable.
- `tests/test_adaptive_gate.py::test_source_unavailable_lane_runs_once_for_a_durable_coverage_gap_then_skips` — same exemplar, same reason.
- `tests/test_daily_duckdb.py::test_source_unavailable_coverage_marker_is_stable_and_write_gated` — same exemplar, same reason.

Run before delivering: `python -m pytest -q` (must be green; the suite is 792 passing today) and `ruff check .` (config: line-length 100, select E4,E7,E9,F).

## The eleven lanes

For each, produce a working adapter + fixtures + tests + registry entry + pipeline wiring + gate branch + docs row. Order below is the recommended sequence.

### 1. Italy LBA — closest to done

- Host confirmed from `sportsdataverse-py` `lba/legabasket.py`: `https://www.legabasket.it/api`.
- Confirmed endpoint: `GET /championships/get-championships?cs_id=1&items=1000&s=2027` (season as end year).
- **Unverified:** the match-list route for a championship/round, and where period scores live (box score vs match detail). Probe them.
- Build: `ingestion/providers/lba.py`, `schedule_source="lba"`, `results_source="lba"`.

### 2. South Korea KBL — smallest surface

- Host confirmed: `https://api.kbl.or.kr`.
- Confirmed endpoint: `GET /season/list`.
- **Unverified:** game/schedule routes, period scores, whether an API key or header is required.
- Build: `ingestion/providers/kbl.py`, source `"kbl"`.

### 3. Japan B.League

- Host confirmed: `https://b-league.s3.amazonaws.com/web_json`.
- Confirmed route shape: `GET /contexts/{schedule_key}`.
- **Unverified (this is the blocker):** how the `schedule_key` set is discovered — find the index/manifest object that enumerates schedule keys. Static S3 JSON is the most stable of the remaining leads if the index exists.
- Build: `ingestion/providers/bleague.py`, source `"bleague"`.

### 4. Australia NBL — prove it exists or drop it

- `sportsdataverse-py` has a module (`sportsdataverse/basketball/nbl_au`), but Ma Golide could not confirm its file path or any route. Treat as unknown.
- The Odds API NBL sport key is **probationary pricing only** and must never supply fixtures or results.
- If no documented period-level route exists, do NOT build a browser-automation scraper. Report it as "no compliant route" with your probe evidence, and leave the lane `source_unavailable`.

### 5. Basketball Champions League (FIBA family)

- No verified route. The validated FIBA pattern lives in `ingestion/providers/fiba_euroleague_w.py` (APIM key shipped in FIBA's web client, competition id parameter). Find BCL's competition id and confirm the same route serves it, or report failure.
- Build: `ingestion/providers/bcl.py`, source `"bcl"`.

### 6–12. France LNB Pro A, Germany BBL, Turkey BSL, Greece Basket League, Israel Premier League, Brazil NBB, Argentina Liga Nacional

No packaged period-level route exists for any of these. For each you must:

1. Find the league's own public schedule + result source (official site API is strongly preferred over an aggregator).
2. Confirm it exposes **period scores**, not just finals. This is the hard requirement — a finals-only source cannot produce a reconcilable `ActualResult`.
3. Establish stable team identity (ids preferred over names) and UTC kickoff times with an explicit offset.
4. Only then write the adapter.

Do these one at a time and report per lane. If a league has no period-level public route, say so with evidence and leave it unregistered. Do not batch them behind a generic scraper: each lane needs its own identity crosswalk and its own failure semantics.

## Things Ma Golide could not confirm — verify these first

These are cheap probes with high information value. Report each result explicitly.

1. **EuroCup live rows.** `GET https://api-live.euroleague.net/v2/competitions/U/seasons/U2026/games`. Does it return a data list? Is `U` the correct competition code, and is `U2026` the active EuroCup season? Ma Golide shipped this lane on documentation alone (`giasemidis/euroleague_api` restricts competition to `E`/`U`; `sportsdataverse-py` documents "E = EuroLeague, U = EuroCup").
2. **ACB without auth.** `GET https://api2.acb.com/api/seasondata/Calendar/season-calendar?season=2027&competition=1`, then `GET https://api2.acb.com/api/matchdata/MatchHeader/match-header?matchId=<id from calendar>`. Confirm: does the calendar require headers/cookies; does `startDateTime` carry a UTC offset; is `quarterScores[].homeScore` points-in-period (not cumulative); is `matchStatus` `FINALIZED` at end of game; is the edition-id offset (season − 1936) real.
3. **BetExplorer coverage for the new leagues.** For each lane you build, probe whether `betexplorer_direct` can actually price its fixtures (the day board + a named-book comparison response). This is the true gate on whether the lane can ever produce a ticket; a lane with fixtures and no prices emits unpriced rows forever.
4. **Terms and rate limits.** Record `robots.txt`, any documented rate limit, and whether automation is permitted, for every host you touch. One line per host. Ma Golide's existing adapters carry the label "source's automation/reuse terms remain unresolved" — match that honesty.

## Output format — one markdown file, nothing else

Produce **exactly one** file, e.g. `zai-handoff.md`, with these sections in this order. Fenced code blocks are expected here; this is a code handoff.

1. **Verification log** — a table: probe URL, method, HTTP status, what it proved, timestamp. Include the failures. This section is the most important one; a claim without a probe is worthless.
2. **Confirmed contracts** — per lane: host, endpoints, required params/headers, field paths for teams, kickoff, status, finals, periods, and the season-key convention.
3. **Code** — the complete contents of every new file, each preceded by its repository-relative path on its own line: `ingestion/providers/<lane>.py`, `tests/test_<lane>_lane.py`, `tests/fixtures/<lane>/*.json`.
4. **Patches** — for files you modify (`ingestion/league_registry.py`, `ingestion/pipeline.py`, `ingestion/adaptive_gate.py`, the four test files listed above), give unified diffs with enough context to apply, or the full replaced function bodies with the function name stated.
5. **Docs** — exact text to add to `docs/PROVIDERS.md` (acquisition details per lane) and `docs/LEAGUE-ROUTE-CONTRACTS.md` (move the lane from "lead" to "implemented", or state why it stays a lead), plus the `README.md` and `docs/LEAGUE-ADOPTION-ROADMAP.md` rows.
6. **Could not verify** — per lane: what failed, the error, and what it blocks. Do not soften this; an honest gap is more useful than a plausible adapter.
7. **Test evidence** — the pytest and ruff output you obtained after applying everything.

Constraints on the deliverable:

- Complete files, not snippets. The maintainer will paste them in verbatim.
- Repository-relative paths, no absolute paths.
- Python 3.11, PEP 604 unions fine, ruff line-length 100.
- If you cannot finish a lane, ship the lanes you did finish and list the rest in section 6. Do not pad section 3 with untested code.
- Do not include any credential, cookie, or token you used while probing.
