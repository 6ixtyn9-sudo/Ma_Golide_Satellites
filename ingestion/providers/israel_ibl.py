"""Israel Premier League (Winner League) adapter for a referenced JSON route.

Contract reference
------------------
Route and field contract referenced from the league's own site and from
SportsDataverse ``hoopR`` documentation (``ibba_games``); Ma Golide retains no
first-party response in this repository, so live behaviour and the source's
automation/reuse terms remain unverified.

* The league publishes its game list as JSON at
  ``GET https://basket.co.il/ws/ws.asmx/Games`` with ``cYear`` (season END
  year), ``board_id`` (5 Winner League, 10 cup, 33 play-in, 34 supercup) and
  zeroed identity filters. The host is Cloudflare-fronted and refuses a plain
  client; it requires a browser ``User-Agent``.
* Rows carry per-quarter scores as pipe-delimited strings: ``score_by_q_1``
  (home quarters, e.g. ``"41|29|32|18"``) and ``score_by_q_2`` (away), plus
  ``total_ot``/``overtime``, finals in ``score_team1``/``score_team2``, team ids
  and English names in ``team1``/``team2``/``team1_name_eng``/``team2_name_eng``,
  and the tip as ``game_date_txt`` (``dd/mm/yyyy``) plus ``TimeOfGame``
  (``HH:MM``).
* **The provider exposes no UTC offset.** The wall clock is local Israeli time,
  so this adapter applies ``Asia/Jerusalem`` explicitly and records that the
  offset is an assumption, never a provider fact. A game whose wall time is
  ambiguous or nonexistent at a DST boundary is refused rather than shifted.

Captured rows (2026-11-22; four board rows committed as
``tests/fixtures/israel_ibl/board5_captured_rows.json`` and re-tested on every
run) settle two things: the quarter strings are points-in-period — the probe
reported all 45 played rows of the full board summing to their finals, and the
committed double-overtime row is ``21|18|18|18|11|10`` against
``20|19|21|15|11|17`` — and a row's season is ``year_id`` rather than the
``cYear`` query parameter. Reading only ``cYear`` left every discovered game with
an empty season, so a later board refresh refused identity for the whole lane.

Every identity, period, and final check below fails closed.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

BASE_URL = "https://basket.co.il/ws/ws.asmx/Games"
PROVIDER = "israel_ibl"

#: The provider serves local wall-clock time with no offset. This is the
#: documented assumption, not a value supplied by the source.
ASSUMED_LEAGUE_TZ = "Asia/Jerusalem"

BOARD_WINNER_LEAGUE = 5
BOARDS = {
    5: "Winner League",
    10: "State Cup",
    33: "Play-in",
    34: "Super Cup",
}
SEASON_FLOOR = 2000
#: The host is Cloudflare-fronted and rejects a plain client with error 1010.
BROWSER_USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0 Safari/537.36"
)

_last_gaps: List[Dict[str, Any]] = []


def last_gaps() -> List[Dict[str, Any]]:
    """Gaps recorded by the most recent call (identity/period/time problems)."""

    return list(_last_gaps)


def _reset_gaps() -> None:
    _last_gaps.clear()


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def season_end_year(target: datetime) -> int:
    """Season END year for a date (the provider's ``cYear`` convention)."""

    return target.year + 1 if target.month >= 7 else target.year


def games_url(end_year: int, board: int = BOARD_WINNER_LEAGUE) -> str:
    """The official board route for one season."""

    if end_year < SEASON_FLOOR:
        raise ValueError(f"Israel Premier League season {end_year} precedes the {SEASON_FLOOR} floor")
    if board not in BOARDS:
        raise ValueError(f"unsupported Israel board id {board!r}")
    return (
        f"{BASE_URL}?cYear={int(end_year)}&board_id={int(board)}"
        "&board_round=0&team_id=0&team_uid=0&game_id=0"
    )


def _int(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _quarter_list(value: Any) -> List[Optional[float]]:
    """Parse a pipe-delimited quarter string (``"41|29|32|18"``)."""

    if value is None:
        return []
    parts = [part.strip() for part in str(value).split("|") if part.strip() != ""]
    quarters: List[Optional[float]] = []
    for part in parts:
        try:
            quarters.append(float(part))
        except ValueError:
            quarters.append(None)
    return quarters


# The Winner League board fills ``TimeOfGame`` with a Hebrew placeholder while a
# fixture's tip time is still unannounced. Observed in run 45185698512
# (2026-11-23): row IBL-26679 carried "טרם נקבע" and the lane rejected it as a
# malformed time. It is not malformed -- it is a provider telling us the time is
# not decided yet, which is a different fact from "the payload is corrupt".
_UNANNOUNCED_TIP_MARKERS = frozenset(
    {
        "טרם נקבע",
        "טרם נקבעה",
        "טרם נקבעו",
        "לא נקבע",
        "tbd",
        "tba",
        "n/a",
        "na",
        "-",
    }
)


class TipTimeUnannounced(ValueError):
    """The provider states that this fixture's tip time is not set yet."""


def _is_unannounced_tip(time_text: str) -> bool:
    normalised = " ".join(time_text.split()).casefold()
    return normalised in _UNANNOUNCED_TIP_MARKERS


def _tip_utc(date_text: Any, time_text: Any, *, played: bool = False) -> tuple[str, bool]:
    """Convert the provider's naive wall clock into UTC.

    Returns ``(iso_utc, dst_edge)``. The wall clock is interpreted in
    :data:`ASSUMED_LEAGUE_TZ`; a wall time that falls in a DST gap is refused,
    and an ambiguous (repeated) wall time is flagged so callers can treat it as
    non-canonical rather than silently picking the first offset.
    """

    raw_date = str(date_text or "").strip()
    raw_time = str(time_text or "").strip()
    if not raw_date:
        raise ValueError("Israel row has no game date")
    try:
        day = datetime.strptime(raw_date, "%d/%m/%Y").date()
    except ValueError as exc:
        raise ValueError(f"Israel game date {raw_date!r} is not dd/mm/yyyy") from exc
    hour, minute = 0, 0
    if raw_time and _is_unannounced_tip(raw_time):
        # Fail closed on the fixture, but say what the provider actually said.
        # An unannounced tip is not a midnight tip and must never be stored as
        # one: it would place a scheduled game at 00:00 local and let it look
        # like a real pre-tip candidate.
        raise TipTimeUnannounced(
            f"tip_time_unannounced: provider reports {raw_time!r} for {raw_date}"
        )
    if raw_time:
        try:
            parsed_time = datetime.strptime(raw_time, "%H:%M")
        except ValueError as exc:
            raise ValueError(f"Israel tip time {raw_time!r} is not HH:MM") from exc
        hour, minute = parsed_time.hour, parsed_time.minute
    elif not played:
        # Same rule for a blank that is not yet a result: no announced tip means
        # the fixture cannot be placed on a clock, so withhold it rather than
        # invent midnight. A row that already carries a final score is different
        # -- its tip time is incidental to grading, and dropping it would lose
        # a settled result.
        raise TipTimeUnannounced(
            f"tip_time_unannounced: provider left the tip time blank for {raw_date}"
        )
    zone = ZoneInfo(ASSUMED_LEAGUE_TZ)
    naive = datetime(day.year, day.month, day.day, hour, minute)
    aware = naive.replace(tzinfo=zone)
    # A DST gap maps a nonexistent wall time onto a neighbouring offset; detect
    # it by round-tripping through UTC.
    round_trip = aware.astimezone(timezone.utc).astimezone(zone).replace(tzinfo=None)
    if round_trip != naive:
        raise ValueError(f"Israel tip {raw_date} {raw_time} falls in a DST gap")
    # Ambiguity: the same wall time occurs twice. ``fold`` distinguishes them.
    other = naive.replace(tzinfo=zone, fold=1)
    dst_edge = aware.utcoffset() != other.utcoffset()
    return to_utc_iso(aware.astimezone(timezone.utc)), dst_edge


def _periods(row: Dict[str, Any], game_ref: str) -> ActualResult:
    """Reconciled regulation quarters (plus overtime when the row exposes it)."""

    home = _quarter_list(row.get("score_by_q_1"))
    away = _quarter_list(row.get("score_by_q_2"))
    if len(home) < 4 or len(away) < 4:
        raise ValueError(f"Israel game {game_ref} exposes fewer than four quarters")
    home_regulation, away_regulation = home[:4], away[:4]
    if any(value is None for value in home_regulation + away_regulation):
        raise ValueError(f"Israel game {game_ref} has an unparseable quarter value")
    home_final = _int(row.get("score_team1"))
    away_final = _int(row.get("score_team2"))
    if home_final is None or away_final is None:
        raise ValueError(f"Israel game {game_ref} has no final score")
    home_overtime: Optional[float] = None
    away_overtime: Optional[float] = None
    overtime_periods = 0
    extra_home = [value for value in home[4:] if value is not None]
    extra_away = [value for value in away[4:] if value is not None]
    if extra_home or extra_away:
        if len(extra_home) != len(extra_away) or not extra_home:
            raise ValueError(f"Israel game {game_ref} has asymmetric overtime periods")
        home_overtime = float(sum(extra_home))
        away_overtime = float(sum(extra_away))
        overtime_periods = len(extra_home)
    total_ot = _int(row.get("total_ot")) or 0
    if total_ot and overtime_periods == 0:
        # The row reports overtime but does not expose its points; refuse rather
        # than pretending the regulation quarters are the whole game.
        raise ValueError(
            f"Israel game {game_ref} reports {total_ot} overtime period(s) without "
            "period points; overtime is unavailable"
        )
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_regulation,
        away_quarters=away_regulation,
        home_overtime=home_overtime,
        away_overtime=away_overtime,
        home_final=home_final,
        away_final=away_final,
        overtime_periods=overtime_periods,
        source=f"israel_ibl board row game={game_ref}",
    )


def _team(row: Dict[str, Any], side: int) -> tuple[str, str]:
    team_id = _int(row.get(f"team{side}"))
    name = row.get(f"team{side}_name_eng") or row.get(f"team{side}_name") or ""
    canonical = canonical_team_name(str(name))
    if team_id is None or not canonical:
        raise IdentityError(f"Israel row lacks team{side} identity")
    return str(team_id), canonical


def _game_ref(row: Dict[str, Any]) -> str:
    for key in ("game_id", "GameId", "id"):
        value = _int(row.get(key))
        if value is not None:
            return str(value)
    home, away = _int(row.get("team1")), _int(row.get("team2"))
    date_text = str(row.get("game_date_txt") or "").replace("/", "")
    if home is None or away is None or not date_text:
        raise IdentityError("Israel row has no usable game identity")
    return f"{date_text}-{home}-{away}"


def _row_season(row: Dict[str, Any], game_ref: str, hint: Optional[int]) -> int:
    """Season end year for a board row.

    The captured board publishes ``year_id``; the query parameter is ``cYear``
    and older or proxied bodies have been seen carrying that name on the row
    too. A row with neither falls back to the season the board was requested
    for, and only then is a missing season refused — an empty season made every
    later result refresh impossible, silently turning the whole lane into gaps.
    """

    for key in ("year_id", "cYear"):
        value = _int(row.get(key))
        if value:
            return value
    if hint:
        return int(hint)
    raise IdentityError(f"Israel game {game_ref} carries no season for a board refresh")


def _game_from_row(
    row: Dict[str, Any], phase: str, season_hint: Optional[int] = None
) -> Game:
    game_ref = _game_ref(row)
    home_id, home_name = _team(row, 1)
    away_id, away_name = _team(row, 2)
    if home_id == away_id or home_name.casefold() == away_name.casefold():
        raise IdentityError(f"Israel game {game_ref} names the same team on both sides")
    home_final = _int(row.get("score_team1"))
    away_final = _int(row.get("score_team2"))
    played = home_final is not None and away_final is not None and (home_final + away_final) > 0
    scheduled_utc, dst_edge = _tip_utc(
        row.get("game_date_txt"), row.get("TimeOfGame"), played=played
    )
    game = Game(
        game_id=f"IBL-{game_ref}",
        league=League.ISRAEL_PREMIER_LEAGUE,
        season=str(_row_season(row, game_ref, season_hint)),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=scheduled_utc,
        status=GameStatus.FINAL if played else GameStatus.SCHEDULED,
        phase=phase,
        provider_ids={
            PROVIDER: game_ref,
            "israel_assumed_timezone": ASSUMED_LEAGUE_TZ,
        },
        evidence={
            "tip_timezone_assumed": ASSUMED_LEAGUE_TZ,
            "tip_ambiguous_at_dst_edge": dst_edge,
        },
    )
    if played:
        game.actual = _periods(row, game_ref)
    return game


def default_fetch() -> FetchFn:
    """Adapter-owned JSON client carrying the browser ``User-Agent``.

    The host is Cloudflare-fronted and answers a non-browser client with error
    1010, so the live transport must send a browser agent. An injected fixture
    transport in tests bypasses this entirely.
    """

    client = HttpClient(timeout=20.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return client.get_json


def fetch_board(
    end_year: int, fetch: Optional[FetchFn] = None, board: int = BOARD_WINNER_LEAGUE
) -> List[Dict[str, Any]]:
    """Raw game rows for one board and season (one request)."""

    url = games_url(end_year, board)
    fetcher = fetch or default_fetch()
    payload = fetcher(url)
    if isinstance(payload, dict):
        rows = payload.get("games")
    else:
        rows = payload
    if not isinstance(rows, list) or not rows:
        raise ValueError(
            f"Israel board {board} season {end_year} returned no games; an empty body is "
            "refused rather than treated as a season without fixtures"
        )
    return [row for row in rows if isinstance(row, dict)]


def discover_games(
    date_iso: Optional[str],
    fetch: FetchFn,
    end_year: int,
    board: int = BOARD_WINNER_LEAGUE,
) -> List[Game]:
    """Games for one UTC date (or all rows when ``date_iso`` is ``None``)."""

    _reset_gaps()
    phase = BOARDS.get(int(board), f"Board {board}")
    games: List[Game] = []
    for row in fetch_board(end_year, fetch, board):
        try:
            game = _game_from_row(row, phase, end_year)
        except (IdentityError, ValueError) as exc:
            _record_gap(str(row.get("game_id") or ""), str(exc))
            continue
        if date_iso and not game.scheduled_utc.startswith(date_iso):
            continue
        games.append(game)
    return games


def fetch_game_result(game: Game, fetch: FetchFn) -> Game:
    """Refresh one game from its board row, refusing a mismatched identity."""

    game_ref = str(game.provider_ids.get(PROVIDER) or "")
    if not game_ref:
        raise IdentityError(f"game {game.game_id} carries no Israel board id")
    season = _int(game.season)
    if season is None:
        raise IdentityError(f"game {game.game_id} has no season for a board refresh")
    for row in fetch_board(season, fetch):
        if _game_ref(row) != game_ref:
            continue
        refreshed = _game_from_row(row, game.phase or "Winner League", season)
        if refreshed.home_name.casefold() != game.home_name.casefold():
            raise IdentityError(
                f"Israel game {game_ref} now names {refreshed.home_name!r}, "
                f"expected {game.home_name!r}"
            )
        return refreshed
    raise IdentityError(f"Israel game {game_ref} is absent from its board response")


def season_results(
    end_year: int,
    fetch: FetchFn,
    failures: Optional[List[Dict[str, str]]] = None,
    board: int = BOARD_WINNER_LEAGUE,
) -> List[Game]:
    """Reconciled finals for one season board."""

    _reset_gaps()
    phase = BOARDS.get(int(board), f"Board {board}")
    results: List[Game] = []
    for row in fetch_board(end_year, fetch, board):
        try:
            game = _game_from_row(row, phase, end_year)
        except (IdentityError, ValueError) as exc:
            if failures is not None:
                failures.append({"game_id": str(row.get("game_id") or ""), "error": str(exc)})
            continue
        if game.status is not GameStatus.FINAL or game.actual is None:
            continue
        ok, reason = game.actual.reconciles()
        if not ok:
            if failures is not None:
                failures.append({"game_id": game.game_id, "error": reason})
            continue
        results.append(game)
    return results


def smoke(fetch: FetchFn, end_year: int, board: int = BOARD_WINNER_LEAGUE) -> Dict[str, Any]:
    """Coverage summary over one season board (never raises)."""

    report: Dict[str, Any] = {
        "games_found": 0,
        "results_reconciled": {"complete": 0, "total": 0},
        "failures": [],
        "verification": "FIXTURE_VERIFIED",
        "season_end_year": int(end_year),
        "board_id": int(board),
        "assumed_timezone": ASSUMED_LEAGUE_TZ,
    }
    try:
        rows = fetch_board(end_year, fetch, board)
    except Exception as exc:  # smoke never raises
        report["status"] = "provider_error"
        report["error_type"] = type(exc).__name__
        report["detail"] = str(exc)[:200]
        return report
    report["games_found"] = len(rows)
    failures: List[Dict[str, str]] = []
    finished = 0
    complete = 0
    for row in rows:
        try:
            game = _game_from_row(row, BOARDS.get(int(board), "Winner League"))
        except (IdentityError, ValueError) as exc:
            failures.append({"game_id": str(row.get("game_id") or ""), "error": str(exc)})
            continue
        if game.status is not GameStatus.FINAL or game.actual is None:
            continue
        finished += 1
        ok, reason = game.actual.reconciles()
        if ok:
            complete += 1
        else:
            failures.append({"game_id": game.game_id, "error": reason})
    report["results_reconciled"] = {"complete": complete, "total": finished}
    report["failures"] = failures
    return report
