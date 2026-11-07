# ingestion — provider authoring guide

Read this fully before writing a provider adapter. It is the binding
interface contract between provider modules and the rest of the pipeline.

## Ground rules (from the master build prompt)

1. **Forebet is excluded.** Never scrape/call/build adapters for it or
   relabel another provider as Forebet. The acquisition layer is
   provider-generic; never populate legacy Forebet-specific inputs with
   non-Forebet data.
2. **No live calls in unit tests.** Every provider function receives a
   `fetch: FetchFn` callable (`FetchFn = Callable[[str], Any]`) that returns
   parsed JSON for a URL. In production the pipeline passes
   `HttpClient().get_json`; in tests it passes a fixture loader. Tests that
   need network must be marked `@pytest.mark.live` and skipped by default.
3. **No silent imputation.** Missing field → `None` + a `CoverageGap`.
   Never fabricate expected scores from probabilities, never derive quarter
   totals from full-game totals/halves, never relabel halves as quarters
   (`PeriodKind.HALVES` stays HALVES; cumulative rows stay
   `PeriodKind.CUMULATIVE`).
4. **Orientation through team IDs.** ESPN stats can sit on either side.
   Resolve the side that carries statistics against the event's
   competitors' team IDs. If a team reference cannot be reconciled, raise
   `IdentityError` (quarantine) — never swap home/away to make a test pass.
5. **UTC ISO 8601 with `Z`** everywhere (`2026-11-07T22:00:00Z`). Never
   naive datetimes. Never parse legacy `M/D/YYYY` display strings for
   scheduling decisions.
6. **No secrets.** Auth headers come from environment variables read inside
   the provider function; missing optional credentials disable the provider
   (return an `unauthorized` coverage state, don't raise to crash the run).
   Never log credential values (`http.redact` on any error text).
7. **Evidence labels.** Report coverage as `LIVE_VERIFIED`,
   `FIXTURE_VERIFIED`, or `DOCUMENTATION_ONLY` in smoke output. Every
   authenticated coverage claim must be backed by an observed response and
   timestamp; otherwise it stays documentation-only.
8. **DuckDB is canonical.** Provider adapters return domain records; they do
   not write Google Sheets, WhatsApp, or application state. The daily
   orchestrator persists those records in DuckDB, materializes CSV.GZ, and
   only then builds derived research reports. Sheets access is explicit legacy
   migration/compatibility only.

## Module shape

Each provider lives in `ingestion/providers/<name>.py` and exposes plain
functions (no classes required). Common helpers live in
`ingestion/providers/_common.py`:

- `FetchFn` type alias
- `IdentityError` — quarantine signal for irreconcilable identity
- `CoverageAccumulator` — small helper collecting `complete/total` counts
  and `CoverageGap`s per field
- `fixture_fetch(*parts)` — test helper returning a callable that serves
  JSON from `tests/fixtures/**`

### ESPN (`espn.py`)

```python
SCOREBOARD_URL = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/scoreboard?dates={yyyymmdd}&limit=100"
SUMMARY_URL    = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/summary?event={event_id}"
STANDINGS_URL  = "https://site.api.espn.com/apis/v2/sports/basketball/nba/standings"
BPI_URL        = "https://sports.core.api.espn.com/v2/sports/basketball/leagues/nba/events/{event_id}/competitions/{event_id}/predictor"

def discover_games(date_yyyymmdd: str, fetch: FetchFn, league=League.NBA) -> List[Game]
def fetch_bpi(event_id: str, fetch: FetchFn, game: Game | None = None) -> Prediction
def fetch_game_result(event_id: str, fetch: FetchFn, game: Game | None = None) -> Game   # fills .actual via summary linescores
def fetch_recent_results(date_yyyymmdd: str, fetch: FetchFn) -> List[Game]
def fetch_standings(fetch: FetchFn) -> List[StandingsRow]   # define StandingsRow in espn.py or derive.py
def smoke(date_yyyymmdd: str, fetch: FetchFn) -> Dict[str, Any]  # actual complete/total coverage per field
```

Notes:
- Scoreboard `dates=` is the **US Eastern calendar date** of the slate
  (observed: `dates=20261105` returns games with UTC datetimes
  `2026-11-06T02:00Z/03:00Z`). Response event dates are UTC. Document this
  in the docstring; conversions use UTC internally.
- Core `events?dates=` discovery returned **0 events** in the 2026-11-06
  live sweep — use the site scoreboard for discovery and record this in the
  provider matrix.
- BPI golden mapping (event `401909938`, DEN @ IND, observed
  2026-11-06T21:51Z, stats on `awayTeam`):
  - side carrying stats: `gameProjection`=76.938 (its win prob),
    `teamChanceLoss`=23.062 (opponent win prob), `teamExpectedPts`=120.046,
    `oppExpectedPts`=110.540, `teamPredPtDiff`=9.506.
  - home prob 23.1, away prob 76.9 → `Prob %` `23.1 - 76.9`, `Pred` `2`,
    `Pred Score` `111-120`, `Avg` `230.6` (one decimal, half-up).
  - Validate: probabilities in 0–100 in payload → store 0–1;
    `|teamExpectedPts - oppExpectedPts - teamPredPtDiff|` ≤ 0.25; probabilities
    sum to 100 ± 0.01 (as percentages). Violations → raise `ValueError`
    (caller quarantines).
- Missing `teamExpectedPts`/`oppExpectedPts` (WNBA/NCAA pattern): keep
  probabilities, leave scores/total `None`, emit gap `expected_scores`.
- `records[].summary` (e.g. `16-22`) and `lastFiveGames` come from summary;
  recent results come from scoreboard of past dates (`state == 'post'`) —
  parse linescores as quarters (NBA has exactly 4 + optional 5th OT element).
  A 5th linescore element = OT; aggregate multiple OT elements into the
  single combined OT value and set `overtime_periods`.

### Hardwood (`hardwood.py`)

```python
FORECASTS_URL = "https://raw.githubusercontent.com/roni-altshuler/nba_predictor/main/backend/data/predictions/game_forecasts.json"

def load_forecasts(fetch: FetchFn) -> ForecastsArtifact   # validates metadata
def fetch_prediction(game: Game, fetch: FetchFn, freshness_hours: float = 12.0, now_utc: str | None = None) -> Prediction | None
```

- Validate `ingestion.status == 'ok'`, `generated_at` freshness,
  `trained_through` not in the future relative to `generated_at`.
- Join by `game_id == ESPN event id` first; then validate teams
  (home/away names) and `date_utc` within tolerance of the game's
  scheduled time; on any mismatch return None + gap, never a best-guess join.
- `p_home/p_away` are 0–1 already. `exp_total` → `expected_total`;
  `exp_home_score`/`exp_away_score`; `exp_margin` → `expected_margin`.
  Convert probabilities from 0–1 to display percentages only at render.
- Record `artifact_sha256`, `generated_at`, `model_version` as provenance.
- Licence: MIT. Never average Hardwood with ESPN outputs — selection is in
  `predictions.py`, not here.

### Predixsport (`predixsport.py`) — disabled by default

```python
BASE_URL = "https://api.predixsport.com"
def is_enabled() -> bool                     # PREDIXSPORT_API_KEY present
def fetch_upcoming(league: str, days: int, fetch_factory) -> List[Prediction]  # league: 'nba'|'euroleague'
def fetch_match_detail(match_id: str, league: str, fetch_factory) -> Prediction
```

- `Authorization: Bearer <PREDIXSPORT_API_KEY>`. If key missing, functions
  return `[]`/None and record `CoverageState.UNAUTHORIZED` in health — the
  keyless pipeline must not fail.
- Derived scores: `home = (total + diff) / 2`, `away = (total - diff) / 2`
  where diff is **home minus away**. Winner/Pred from win probability only.
  Independent model disagreements (winner vs margin sign) are preserved.
- Save `produced_by` and honour the 200 shared calls/month budget: one
  list call per refresh + cached detail calls. Cache is the store's job;
  this module just performs calls when asked.
- Until a real-key smoke test occurs, tests are fixture-based and output
  must be labelled `FIXTURE_VERIFIED` / `DOCUMENTATION_ONLY`.

### EuroLeague (`euroleague.py`)

```python
SCHEDULE_URL = "https://api-live.euroleague.net/v2/competitions/E/seasons/{seasoncode}/games"
BOXSCORE_URL = "https://live.euroleague.net/api/Boxscore?gamecode={gamecode}&seasoncode={seasoncode}"

def discover_games(date_iso: str | None, fetch: FetchFn, seasoncode: str = "E2026") -> List[Game]
def fetch_game_result(game: Game, fetch: FetchFn) -> Game     # boxscore → ActualResult
def fetch_standings(fetch: FetchFn, seasoncode: str = "E2026") -> List[StandingsRow]  # if endpoint available; else gap
def smoke(fetch: FetchFn, seasoncode: str = "E2026") -> Dict[str, Any]
```

- Schedule gives `utcDate`, `local`/`road` club objects (`code`, `name`,
  `tvCode`), `score`, `partials`, `played`, `gameCode`, `round`, `winner`.
  Home/away resolution uses `local`/`road` fields — never array order.
- Boxscore `ByQuarter` = individual quarters (`Quarter1..Quarter4`,
  `Extra1..` for OT); `EndOfQuarter` = **cumulative** — never write
  EndOfQuarter values as individual quarters.
- OT: `Extra1`(+`Extra2`…) sum into the single combined OT pair;
  `overtime_periods` = number of Extra periods present.
- Team names from boxscore are UPPERCASE ("REAL MADRID"); canonicalise via
  `identity.canonical_team_name` and join to the schedule game via
  `gameCode`/`seasoncode` (authoritative), validating team codes.

### Bookmakers (`therundown.py`, `sportsgameodds.py`, `theoddsapi.py`)

Each exposes:

```python
def is_enabled() -> bool
def fetch_odds(league: str, date_yyyymmdd: str, fetch_factory) -> List[BookSnapshot]
def smoke(league: str, fetch_factory) -> Dict[str, Any]   # coverage per market/period, quota headers
```

- Produce `BookSnapshot` records only through normalised
  market/period/side fields. Totals: `line` = the line value (e.g. 232.0);
  prices go to `over/under_price_decimal`. Moneyline: `home/away_price_decimal`.
  American → decimal via `books.american_to_decimal`.
- TheRundown: sport ID for NBA resolved from provider data, **never** the
  guide's NFL id. Full-game markets are ids `1,2,3`. Free plan has **no
  quarter markets** — `therundown` is in `books.PROVIDERS_BLOCKED_FROM_QUARTERS`
  and its adapter must never emit a `Period.Q*` snapshot (there is no code
  path; assert this in tests).
- SportsGameOdds: oddIDs confirmed through `/markets` only — do not infer
  coverage from the oddID pattern. Quarter oddIDs
  `points-all-1q-ou-over/-under` etc. map to exact periods.
- The Odds API: `/events/{id}/odds` for period markets; capture
  `x-requests-remaining`/`x-requests-used` into health quota.
- All three: `pre_match`/`main_line` flags set from payload; feature-flagged
  off until authenticated coverage is measured (report `DOCUMENTATION_ONLY`).

## Testing requirements (per provider)

- `tests/test_<provider>.py` with small committed fixtures under
  `tests/fixtures/<provider>/`.
- No network in ordinary tests. `fetch` = fixture loader or
  `http.HttpClient(session=_FakeSession(...))`.
- Prove: orientation through team IDs (both placements), missing-expected-
  scores → blank + gap, quarters vs cumulative, OT aggregation, identity
  mismatch → quarantine, provider idempotency of mapping.
- ESPN golden: fixture `tests/fixtures/espn/predictor_401909938.json`
  renders exactly `23.1 - 76.9`, `2`, `111-120`, `230.6`.
