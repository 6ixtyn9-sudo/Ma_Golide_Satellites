"""Official EuroLeague provider adapter (schedule / boxscore / standings).

Evidence status (see docs/PROVIDERS.md)
--------------------------------------------------
- Schedule + Boxscore endpoints are ``LIVE_VERIFIED``: observed live on
  2026-11-06 and captured into the committed fixtures under
  ``tests/fixtures/euroleague/`` (``smoke()`` therefore reports
  ``"LIVE_VERIFIED (live captures 2026-11-06)"``).
- The v2 standings endpoint was NOT confirmed at baseline: ``fetch_standings``
  attempts it defensively, never raises, and degrades to ``[]`` plus a
  reason; its evidence label stays ``DOCUMENTATION_ONLY`` until an observed
  response proves the schema.

Binding rules implemented here
------------------------------
- Schedule home/away comes from the ``local``/``road`` club objects, never
  from array order; a row whose local and road club codes coincide raises
  ``IdentityError`` instead of guessing orientation.
- ``scheduled_utc`` is the schedule's ``utcDate`` normalised to
  timezone-aware ``...Z`` form.
- Boxscore ``ByQuarter`` rows carry INDIVIDUAL quarter points
  (``Quarter1..Quarter4`` plus ``Extra1..ExtraN`` for overtime).
  ``EndOfQuarter`` rows are CUMULATIVE and are used ONLY for the running
  final (their last value = the official final score); they are never
  written as individual quarters, and a payload without ``ByQuarter`` is
  rejected instead of improvised from cumulative rows.
- Boxscore ``Team`` names are UPPERCASE (e.g. ``REAL MADRID``); they are
  matched against the schedule game's home/away names case-insensitively
  (normalized exact or containment match, plus the canonical alias table as
  a fallback for short codes such as ``RMB``).  A row matching neither side
  raises ``IdentityError`` — callers quarantine, never assume array order.
- Multiple OT periods are aggregated into the single combined OT pair
  (``home_overtime``/``away_overtime``) and ``overtime_periods`` counts them,
  per the ResultsClean ``OT`` cell semantics.
- Final scores: explicit totals on the boxscore rows when present, else the
  last cumulative ``EndOfQuarter`` value per team, else the sum of the
  individual quarters + OT.  The schedule payload's own scores are not
  reachable here (``Game`` carries no schedule-score slot), so the official
  cumulative ``EndOfQuarter`` final is the authoritative fallback — and
  ``ActualResult.reconciles()`` is enforced on every source.
- OT discovery: ``Game`` has no discovery-time OT slot, so the number of OT
  periods visible in the schedule ``partials.extraPeriods`` is stored under
  the provider-namespaced ``provider_ids['euroleague_ot_periods']`` key (see
  :func:`schedule_ot_periods`) — discovery itself flags OT games.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from ..http import HttpClient
from ..identity import canonical_team_name, normalize_team_name, parse_utc, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

# EuroLeague Basketball serves both competitions from one host; the first
# letter of the season code is the competition code (``E`` EuroLeague, ``U``
# EuroCup). The route shape is identical, so the competition is derived from
# the season code rather than threaded as a second, disagreeable parameter.
COMPETITION_EUROLEAGUE = "E"
COMPETITION_EUROCUP = "U"
COMPETITIONS = {COMPETITION_EUROLEAGUE, COMPETITION_EUROCUP}

SCHEDULE_URL = (
    "https://api-live.euroleague.net/v2/competitions/{competition}"
    "/seasons/{seasoncode}/games"
)
BOXSCORE_URL = (
    "https://live.euroleague.net/api/Boxscore?gamecode={gamecode}&seasoncode={seasoncode}"
)
# v3 per-round standings replaced the dead v2 path (live 2026-11-16: the v2
# ``standings`` endpoint answers 404; the v3 per-round basicstandings answers
# 200 for completed rounds and 404 for rounds not yet finished).
STANDINGS_URL = (
    "https://api-live.euroleague.net/v3/competitions/{competition}"
    "/seasons/{seasoncode}/rounds/{round}/basicstandings"
)

# Evidence labels used by smoke reports.
VERIFICATION_LIVE = "LIVE_VERIFIED (live captures 2026-11-06; standings 2026-11-16)"
VERIFICATION_FIXTURE = "FIXTURE_VERIFIED"
EVIDENCE_STANDINGS = "LIVE_VERIFIED (v3 per-round standings 2026-11-16)"

OT_PERIODS_KEY = "euroleague_ot_periods"

_QUARTER_KEYS = ("Quarter1", "Quarter2", "Quarter3", "Quarter4")
_EXTRA_RE = re.compile(r"^Extra(\d+)$")


def competition_of(seasoncode: str) -> str:
    """Competition code carried by a season code (``U2026`` -> ``U``).

    ``E`` EuroLeague and ``U`` EuroCup are documented by the provider host and
    by two independent public clients; anything else fails closed rather than
    being sent to an unvalidated path.
    """
    code = str(seasoncode or "").strip().upper()
    if len(code) != 5 or not code[1:].isdigit() or code[0] not in COMPETITIONS:
        raise ValueError(
            f"euroleague season code {seasoncode!r} must look like E2026 or U2026"
        )
    return code[0]


def season_code(competition: str, year: int) -> str:
    """Canonical season code for one competition and season start year."""
    code = str(competition or "").strip().upper()
    if code not in COMPETITIONS:
        raise ValueError(f"unsupported euroleague competition code {competition!r}")
    return f"{code}{int(year)}"


def _season_display(seasoncode: str) -> str:
    """``E2026``/``U2026`` -> ``2026`` (the canonical Game.season display value)."""
    code = str(seasoncode or "")
    head = code[:1].upper()
    if head in COMPETITIONS and code[1:].isdigit():
        return code[1:]
    return code


# ---------------------------------------------------------------------------
# Discovery (schedule)
# ---------------------------------------------------------------------------


def _fetch_season_rows(fetch: FetchFn, seasoncode: str) -> List[Dict[str, Any]]:
    """Fetch and validate the raw season schedule rows (one request)."""
    payload = fetch(
        SCHEDULE_URL.format(competition=competition_of(seasoncode), seasoncode=seasoncode)
    )
    rows = payload.get("data") if isinstance(payload, dict) else payload
    if not isinstance(rows, list):
        raise ValueError(
            f"euroleague schedule for {seasoncode} has no data list (got {type(payload).__name__})"
        )
    if not rows:
        # Live trap (verified 2026-11-16): a bare-year season code (``2026``
        # instead of ``E2026``) answers HTTP 200 with an EMPTY data list.
        # Silence would poison backfill with a season of zero games, so an
        # empty season payload is refused outright.
        raise ValueError(
            f"euroleague schedule for {seasoncode} is empty; a wrong season code "
            "(e.g. '2026' instead of 'E2026') answers 200 with no rows"
        )
    return [row for row in rows if isinstance(row, dict)]


def discover_games(
    date_iso: Optional[str], fetch: FetchFn, seasoncode: str = "E2026"
) -> List[Game]:
    """Discover games from the official season schedule.

    ``date_iso`` (``YYYY-MM-DD``) filters to games whose ``utcDate`` falls on
    that UTC date; ``None`` returns the whole season payload.  Game ids are
    canonical ``{seasoncode}-{gameCode}`` (e.g. ``E2026-97``); the euroleague
    gamecode and season code are kept in ``provider_ids`` (``euroleague``,
    ``euroleague_season``), alongside ``euroleague_ot_periods`` when the
    schedule partials already show overtime.
    """
    games: List[Game] = []
    for row in _fetch_season_rows(fetch, seasoncode):
        utc_date = row.get("utcDate")
        if date_iso and not str(utc_date or "").startswith(date_iso):
            continue
        games.append(_game_from_schedule_row(row, seasoncode))
    return games


def _extra_period_count(row: Dict[str, Any]) -> int:
    """Number of OT periods visible in a schedule row's partials."""
    indices: set = set()
    for side in ("local", "road"):
        extras = ((row.get(side) or {}).get("partials") or {}).get("extraPeriods") or {}
        if isinstance(extras, dict):
            indices.update(str(k) for k in extras)
    return len(indices)


def _game_from_schedule_row(row: Dict[str, Any], seasoncode: str) -> Game:
    """Map one schedule row to a canonical Game.

    Status: ``played=True`` -> FINAL, else SCHEDULED.  Phase is the official
    ``roundName`` (e.g. ``Round 38``), used as the canonical game phase.
    Home/away orientation is validated: identical local/road club codes (or,
    when codes are absent, identical club names) raise ``IdentityError``.
    """
    gamecode = row.get("gameCode")
    if gamecode is None:
        raise ValueError(f"euroleague schedule row without gameCode: {row.get('id')!r}")
    gamecode = str(gamecode)
    utc_date = row.get("utcDate")
    if not utc_date:
        raise ValueError(f"euroleague schedule row {gamecode} has no utcDate; refusing to guess")
    try:
        scheduled_utc = to_utc_iso(parse_utc(str(utc_date)))
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"euroleague schedule row {gamecode} has unparseable utcDate {utc_date!r}: {exc}"
        ) from exc

    local = row.get("local") or {}
    road = row.get("road") or {}
    local_club = local.get("club") or {}
    road_club = road.get("club") or {}
    home_name = canonical_team_name(str(local_club.get("name") or ""))
    away_name = canonical_team_name(str(road_club.get("name") or ""))
    home_code = local_club.get("code")
    road_code = road_club.get("code")
    if (home_code and road_code and str(home_code) == str(road_code)) or (
        not (home_code and road_code)
        and home_name
        and normalize_team_name(home_name) == normalize_team_name(away_name)
    ):
        raise IdentityError(
            f"euroleague schedule row {gamecode}: local and road refer to the "
            f"same club (codes {home_code!r}/{road_code!r}, names "
            f"{home_name!r}/{away_name!r}); home/away cannot be oriented"
        )

    provider_ids = {
        "euroleague": gamecode,
        "euroleague_season": str(seasoncode),
        "euroleague_competition": competition_of(seasoncode),
    }
    ot_periods = _extra_period_count(row)
    if ot_periods:
        provider_ids[OT_PERIODS_KEY] = str(ot_periods)

    return Game(
        game_id=f"{seasoncode}-{gamecode}",
        league=League.EUROLEAGUE,
        season=_season_display(seasoncode),
        home_team_id=str(home_code or home_name),
        away_team_id=str(road_code or away_name),
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=scheduled_utc,
        status=GameStatus.FINAL if row.get("played") else GameStatus.SCHEDULED,
        phase=str(row.get("roundName") or "League"),
        provider_ids=provider_ids,
    )


def schedule_ot_periods(game: Game) -> int:
    """OT periods the schedule already showed for this game (0 when none)."""
    try:
        return int(game.provider_ids.get(OT_PERIODS_KEY, "0") or 0)
    except (TypeError, ValueError):
        return 0


# ---------------------------------------------------------------------------
# Schedule-row results (whole-season reconcilable finals, one request)
# ---------------------------------------------------------------------------


def _partial(row: Dict[str, Any], key: str, game_ref: str) -> float:
    value = (row.get("partials") or {}).get(key)
    if value is None or isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(
            f"euroleague schedule result {game_ref} lacks numeric {key}; refusing "
            "a final-score-only record"
        )
    return float(value)


def _actual_from_schedule_row(row: Dict[str, Any], seasoncode: str) -> ActualResult:
    """Reconciled result built from a played schedule row only.

    Schedule rows carry the official final score per side plus the official
    per-quarter partials (``partials1..4``) and any overtime partials
    (``extraPeriods``, e.g. ``{'1': 19}``). All four regulation partials are
    required per side — a row without them is refused (never stored as a
    final-score-only record). Reconciliation is enforced exactly as for the
    boxscore path.
    """
    game_ref = f"{seasoncode}-{row.get('gameCode')}"

    def side(side_key: str) -> Tuple[List[float], Optional[float], int]:
        block = row.get(side_key) or {}
        quarters = [_partial(block, f"partials{i}", game_ref) for i in (1, 2, 3, 4)]
        extras = (block.get("partials") or {}).get("extraPeriods") or {}
        values = []
        for key in sorted(extras, key=lambda item: str(item)):
            value = extras[key]
            if value is None or isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ValueError(
                    f"euroleague schedule result {game_ref} has invalid extraPeriods "
                    f"entry {key!r}: {value!r}"
                )
            values.append(float(value))
        return quarters, (sum(values) if values else None), len(values)

    home_quarters, home_ot, home_extra = side("local")
    away_quarters, away_ot, away_extra = side("road")
    overtime_periods = max(home_extra, away_extra)
    if (home_ot is None) != (away_ot is None):
        raise ValueError(
            f"euroleague schedule result {game_ref}: only one side shows overtime partials"
        )
    home_final = _partial_score(row, "local", game_ref)
    away_final = _partial_score(row, "road", game_ref)
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=home_final,
        away_final=away_final,
        overtime_periods=overtime_periods,
        source=f"euroleague schedule season={seasoncode}",
    )
    ok, reason = actual.reconciles()
    if not ok:
        raise ValueError(f"euroleague schedule result {game_ref} failed reconciliation: {reason}")
    return actual


def _partial_score(row: Dict[str, Any], side_key: str, game_ref: str) -> float:
    block = row.get(side_key) or {}
    value = block.get("score")
    if value is None or isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"euroleague schedule result {game_ref} lacks the {side_key} score")
    return float(value)


def season_results(
    seasoncode: str,
    fetch: FetchFn,
    failures: Optional[List[Dict[str, str]]] = None,
) -> List[Game]:
    """Reconciled finals for every played game of one season.

    One schedule request covers the whole season: every played row already
    carries the official score and per-quarter partials, so a full season
    of reconcilable history costs a single request. Rows that fail identity
    or reconciliation are skipped and, when ``failures`` is given, recorded
    there — never folded, never silently dropped.
    """
    games: List[Game] = []
    for row in _fetch_season_rows(fetch, seasoncode):
        try:
            game = _game_from_schedule_row(row, seasoncode)
        except (IdentityError, ValueError) as exc:
            if failures is not None:
                failures.append(
                    {"game_id": str(row.get("gameCode")), "error": f"schedule row: {exc}"}
                )
            continue
        if game.status is not GameStatus.FINAL:
            continue
        try:
            game.actual = _actual_from_schedule_row(row, seasoncode)
        except ValueError as exc:
            if failures is not None:
                failures.append({"game_id": game.game_id, "error": str(exc)})
            continue
        games.append(game)
    return games


def ot_games_from_schedule(fetch: FetchFn, seasoncode: str = "E2026") -> List[str]:
    """Canonical ids of games whose schedule partials show overtime."""
    return [
        game.game_id
        for game in discover_games(None, fetch, seasoncode)
        if schedule_ot_periods(game) > 0
    ]


def _gamecode_of(game: Game) -> str:
    code = game.provider_ids.get("euroleague")
    if code:
        return str(code)
    match = re.search(r"(\d+)$", game.game_id)
    if match:
        return match.group(1)
    raise IdentityError(
        f"game {game.game_id} carries no euroleague gamecode; cannot fetch boxscore"
    )


def _seasoncode_of(game: Game) -> str:
    """Season code for the boxscore URL: explicit id first, else derived.

    The derived path is competition-aware: a stored ``euroleague_competition``
    or an owned ``U`` game id yields a EuroCup season code, never a EuroLeague
    one. Guessing the wrong competition would silently query another
    competition's season.
    """
    code = game.provider_ids.get("euroleague_season")
    if code:
        return str(code)
    competition = str(game.provider_ids.get("euroleague_competition") or "").upper()
    season = str(game.season or "")
    if season[:1].upper() in COMPETITIONS and season[1:].isdigit():
        return season
    if competition not in COMPETITIONS:
        competition = COMPETITION_EUROLEAGUE
    return f"{competition}{season}"


# ---------------------------------------------------------------------------
# Results (boxscore)
# ---------------------------------------------------------------------------


def fetch_game_result(game: Game, fetch: FetchFn) -> Game:
    """Fetch the official boxscore and attach a reconciled ``ActualResult``.

    Raises ``IdentityError`` when boxscore teams cannot be matched to the
    schedule game's home/away, and ``ValueError`` when the payload cannot
    produce individual quarters or fails reconciliation (callers quarantine
    both).  On success the game status becomes FINAL.
    """
    gamecode = _gamecode_of(game)
    seasoncode = _seasoncode_of(game)
    payload = fetch(BOXSCORE_URL.format(gamecode=gamecode, seasoncode=seasoncode))
    if not isinstance(payload, dict):
        raise ValueError(f"boxscore for {game.game_id} is not a JSON object")

    by_quarter = payload.get("ByQuarter")
    if not isinstance(by_quarter, list) or not by_quarter:
        # EndOfQuarter is cumulative and must NEVER stand in for individual
        # quarters; without ByQuarter there is nothing safe to parse.
        raise ValueError(
            f"boxscore for {game.game_id} has no individual ByQuarter rows "
            "(cumulative EndOfQuarter must not be used as quarters)"
        )

    home_row, away_row = _match_team_rows(game, by_quarter)
    home_quarters = [_period_value(home_row, k) for k in _QUARTER_KEYS]
    away_quarters = [_period_value(away_row, k) for k in _QUARTER_KEYS]
    home_ot, away_ot, overtime_periods = _extra_periods(home_row, away_row)
    home_final, away_final = _final_scores(
        game,
        payload,
        home_row,
        away_row,
        home_quarters,
        away_quarters,
        home_ot,
        away_ot,
    )

    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=home_final,
        away_final=away_final,
        overtime_periods=overtime_periods,
        source=f"euroleague boxscore gamecode={gamecode} season={seasoncode}",
    )
    ok, reason = actual.reconciles()
    if not ok:
        raise ValueError(f"boxscore for {game.game_id} failed reconciliation: {reason}")
    game.actual = actual
    game.status = GameStatus.FINAL
    return game


def _team_matches(row_team: str, game_name: str) -> bool:
    """Case-insensitive match, tolerant of containment and short codes.

    Exact normalized comparison covers the observed boxscore convention
    (UPPERCASE full names vs mixed-case schedule names).  A containment
    match (either direction) tolerates suffixes such as city qualifiers, and
    the canonical alias table additionally reconciles short codes/tv codes
    (e.g. ``RMB`` -> ``real madrid``).  Anything else stays unmatched and
    quarantines — never matched by array order.
    """
    rt = normalize_team_name(row_team)
    gt = normalize_team_name(game_name)
    if not rt or not gt:
        return False
    if rt == gt:
        return True
    if gt in rt or rt in gt:
        return True
    return canonical_team_name(row_team) == canonical_team_name(game_name)


def _match_team_rows(game: Game, rows: List[Any]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Match the two ByQuarter rows to the schedule game's home/away names.

    Never assumes array order: rows are matched by team name; ambiguity or a
    non-matching team raises ``IdentityError`` (quarantine signal).
    """
    dict_rows = [r for r in rows if isinstance(r, dict)]
    if len(dict_rows) != 2:
        raise IdentityError(
            f"boxscore for {game.game_id} does not have exactly two team rows "
            f"(got {len(dict_rows)})"
        )
    home_row: Optional[Dict[str, Any]] = None
    away_row: Optional[Dict[str, Any]] = None
    for row in dict_rows:
        team = str(row.get("Team") or "")
        if _team_matches(team, game.home_name):
            if home_row is not None:
                raise IdentityError(
                    f"boxscore for {game.game_id}: both rows match home team {game.home_name!r}"
                )
            home_row = row
        elif _team_matches(team, game.away_name):
            if away_row is not None:
                raise IdentityError(
                    f"boxscore for {game.game_id}: both rows match away team {game.away_name!r}"
                )
            away_row = row
        else:
            raise IdentityError(
                f"boxscore for {game.game_id}: team {team!r} matches neither "
                f"home {game.home_name!r} nor away {game.away_name!r}"
            )
    if home_row is None or away_row is None:
        raise IdentityError(
            f"boxscore for {game.game_id}: home/away rows not resolved for "
            f"{game.home_name!r} vs {game.away_name!r}"
        )
    return home_row, away_row


def _period_value(row: Dict[str, Any], key: str) -> Optional[float]:
    value = row.get(key)
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _extra_periods(
    home_row: Dict[str, Any], away_row: Dict[str, Any]
) -> Tuple[Optional[float], Optional[float], int]:
    """Aggregate Extra1..ExtraN into the single combined OT pair.

    Returns (home_ot, away_ot, overtime_periods); OT is None when no Extra
    period exists anywhere.
    """
    indices = set()
    for row in (home_row, away_row):
        for key in row or {}:
            m = _EXTRA_RE.match(str(key))
            if m:
                indices.add(int(m.group(1)))
    if not indices:
        return None, None, 0

    def _sum_ot(row: Dict[str, Any]) -> Optional[float]:
        values = [_period_value(row, f"Extra{i}") for i in sorted(indices)]
        numeric = [v for v in values if v is not None]
        return sum(numeric) if numeric else None

    return _sum_ot(home_row), _sum_ot(away_row), len(indices)


def _endofquarter_final(game: Game, payload: Dict[str, Any], team_name: str) -> Optional[float]:
    """Official final for one team from the CUMULATIVE EndOfQuarter rows.

    The last value in ``Quarter1..Quarter4, Extra1..ExtraN`` order is the
    running total through the final period (67 for Besiktas in game 97; 112
    for Real Madrid in OT game 91).  Used ONLY as a final score — never as
    individual quarters.  Returns None when no row matches ``team_name``.
    """
    rows = payload.get("EndOfQuarter")
    if not isinstance(rows, list):
        return None
    for row in rows:
        if not isinstance(row, dict) or not _team_matches(str(row.get("Team") or ""), team_name):
            continue
        ordered_keys = list(_QUARTER_KEYS) + [f"Extra{i}" for i in sorted(_extra_indices(row))]
        present = [_period_value(row, k) for k in ordered_keys]
        numeric = [v for v in present if v is not None]
        return numeric[-1] if numeric else None
    return None


def _extra_indices(row: Dict[str, Any]) -> List[int]:
    indices = set()
    for key in row or {}:
        m = _EXTRA_RE.match(str(key))
        if m:
            indices.add(int(m.group(1)))
    return sorted(indices)


def _final_scores(
    game: Game,
    payload: Dict[str, Any],
    home_row: Dict[str, Any],
    away_row: Dict[str, Any],
    home_quarters: List[Optional[float]],
    away_quarters: List[Optional[float]],
    home_ot: Optional[float],
    away_ot: Optional[float],
) -> Tuple[Optional[float], Optional[float]]:
    """Final scores: explicit totals, then cumulative EndOfQuarter, then sum.

    Order of authority (see module docstring): (1) a ``Total`` key on the
    team rows or a top-level ``Totals``/``Total`` row list (matched by team,
    never by order); (2) the last cumulative ``EndOfQuarter`` value per team
    (the official running final); (3) the sum of the individual quarters +
    aggregated OT.  ``ActualResult.reconciles()`` is enforced by the caller
    regardless of the source used.
    """
    home_total = _period_value(home_row, "Total")
    away_total = _period_value(away_row, "Total")
    if home_total is None or away_total is None:
        totals_rows = payload.get("Totals") or payload.get("Total")
        if isinstance(totals_rows, list) and len(totals_rows) == 2:
            try:
                t_home, t_away = _match_team_rows(game, totals_rows)
            except (IdentityError, ValueError):
                t_home = t_away = None
            if t_home is not None:
                home_total = (
                    home_total if home_total is not None else _period_value(t_home, "Total")
                )
                away_total = (
                    away_total if away_total is not None else _period_value(t_away, "Total")
                )
    if home_total is not None and away_total is not None:
        return home_total, away_total

    home_final = _endofquarter_final(game, payload, game.home_name)
    away_final = _endofquarter_final(game, payload, game.away_name)
    if home_final is not None and away_final is not None:
        return home_final, away_final

    def _sum_periods(quarters: List[Optional[float]], ot: Optional[float]) -> Optional[float]:
        if any(v is None for v in quarters):
            return None
        total = sum(v for v in quarters if v is not None)
        return total + ot if ot is not None else total

    return _sum_periods(home_quarters, home_ot), _sum_periods(away_quarters, away_ot)


# ---------------------------------------------------------------------------
# Standings (v3 per-round basicstandings; live-verified 2026-11-16)
# ---------------------------------------------------------------------------


def _latest_completed_round(fetch: FetchFn, seasoncode: str) -> Optional[int]:
    """Highest round with played games in the season schedule (1 request)."""
    rounds = [
        int(row.get("round"))
        for row in _fetch_season_rows(fetch, seasoncode)
        if row.get("played") and isinstance(row.get("round"), (int, float))
    ]
    return max(rounds) if rounds else None


def fetch_standings(
    fetch: FetchFn,
    seasoncode: str = "E2026",
    round: Optional[int] = None,
    report: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    """Standings from the v3 per-round basicstandings endpoint.

    The endpoint answers 200 only for COMPLETED rounds (verified 2026-11-16:
    round 11 OK, round 12 → 404 ``{"detail": "Not found standings for
    season: E2026, round: 12"}``). With ``round=None`` the latest completed
    round is derived from the season schedule (one extra request). The call
    stays defensive — standings must never break a run — and every failure
    fills ``report`` with ``{'status', 'reason', 'evidence'}``.
    """

    def _fail(reason: str) -> List[Dict[str, Any]]:
        if report is not None:
            report.update({"status": "unavailable", "reason": reason, "evidence": EVIDENCE_STANDINGS})
        return []

    try:
        if round is None:
            round = _latest_completed_round(fetch, seasoncode)
            if round is None:
                return _fail(f"season {seasoncode} has no played rounds yet")
        payload = fetch(
            STANDINGS_URL.format(
                competition=competition_of(seasoncode), seasoncode=seasoncode, round=int(round)
            )
        )
    except Exception as exc:  # deliberate: standings must never break a run
        return _fail(f"standings fetch failed: {type(exc).__name__}: {exc}")

    teams = payload.get("teams") if isinstance(payload, dict) else payload
    if not isinstance(teams, list):
        return _fail(
            "unexpected standings payload shape "
            f"(got {type(payload).__name__}); schema stays fixture-labelled"
        )
    out: List[Dict[str, Any]] = []
    for row in teams:
        if not isinstance(row, dict):
            continue
        club = row.get("club") or row.get("team")
        out.append(
            {
                "position": row.get("position"),
                "team": club.get("name") if isinstance(club, dict) else club,
                "games": row.get("gamesPlayed", row.get("games")),
                "wins": row.get("gamesWon", row.get("won", row.get("wins"))),
                "losses": row.get("gamesLost", row.get("lost", row.get("losses"))),
                "points_for": row.get("pointsFor", row.get("points_for")),
                "points_against": row.get("pointsAgainst", row.get("points_against")),
            }
        )
    if report is not None:
        report.update(
            {"status": "ok", "reason": "", "evidence": EVIDENCE_STANDINGS, "round": str(round)}
        )
    return out


# ---------------------------------------------------------------------------
# Smoke
# ---------------------------------------------------------------------------


def smoke(fetch: FetchFn, seasoncode: str = "E2026") -> Dict[str, Any]:
    """Coverage summary over the schedule + reconciled boxscores.

    Returns the required shape::

        {"games_found": n,
         "results_reconciled": {"complete": x, "total": y},
         "ot_games": [...],              # canonical ids flagged by the schedule
         "verification": "LIVE_VERIFIED (live captures 2026-11-06)"}

    plus a ``failures`` list so per-game boxscore problems are recorded, not
    silently dropped (smoke never raises).  ``total`` counts only games the
    schedule marks played — the games a result fetch could legitimately
    reconcile.  A real ``HttpClient`` run is labelled with the live-capture evidence;
    fixture replay is labelled ``FIXTURE_VERIFIED`` and is never presented as
    a new live call.
    """
    games = discover_games(None, fetch, seasoncode)
    played = [g for g in games if g.status is GameStatus.FINAL]
    ot_games = [g.game_id for g in games if schedule_ot_periods(g) > 0]

    complete = 0
    failures: List[Dict[str, str]] = []
    for game in played:
        try:
            fetch_game_result(game, fetch)
        except Exception as exc:
            failures.append({"game_id": game.game_id, "error": f"{type(exc).__name__}: {exc}"})
            continue
        if game.actual is not None and game.actual.reconciles()[0]:
            complete += 1

    fetch_target = getattr(fetch, "__self__", fetch)
    verification = (
        VERIFICATION_LIVE if isinstance(fetch_target, HttpClient) else VERIFICATION_FIXTURE
    )
    return {
        "games_found": len(games),
        "results_reconciled": {"complete": complete, "total": len(played)},
        "ot_games": ot_games,
        "verification": verification,
        "failures": failures,
    }
