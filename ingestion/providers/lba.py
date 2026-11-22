"""Italy LBA (Lega Basket Serie A) schedule/result adapter for legabasket.it.

Contract reference
------------------
Route and field contract referenced from a live probe of
``https://www.legabasket.it/api`` performed by an external agent on 2026-11-22
and reported to this repository; **Ma Golide retains no first-party response**,
so live behaviour and the source's automation/reuse terms remain unverified.

* ``GET /championships/get-championships?cs_id=1&items=1000&s={first_year}``
  returns one row per competition of that season. ``s`` is the season's FIRST
  year (``s=2026`` returned ``Regular Season 2026/27``, id 602; ``s=2027``
  returned an empty list), so the season key is *not* an end year.
* ``GET /championships/get-championships-calendar-by-id?id={championship_id}``
  returns the league's CURRENT focus day: stable ``id``, ``h_team_id``/
  ``h_team_name``/``v_team_id``/``v_team_name`` (plus club codes),
  ``home_final_score``/``visitor_final_score``, ``game_status`` and
  ``match_datetime`` carrying an explicit offset. Probed day selectors
  (``matchDay``, ``day``/``dayId``, explicit dates) are ignored server-side, so
  this route cannot enumerate a season and history backfill is refused rather
  than guessed.
* ``GET /championships/get-championships-matches-play-by-play?id={match_id}``
  returns reverse-chronological ``pbp.actions[]`` carrying ``period`` (1-4
  regulation, 5+ overtime), ``order`` (descending) and ``score``
  (``"100 - 86"``, cumulative). The highest-``order`` scored action of a period
  is that period's end; regulation quarters are the consecutive differences and
  must reconcile against the final cumulative.

Every scope, identity, timestamp, period, and final check fails closed: a body
that cannot be reconciled becomes a recorded gap, never a fabricated result.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Dict, List, Optional

from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

BASE_URL = "https://www.legabasket.it/api"
CHAMPIONSHIPS_URL = BASE_URL + "/championships/get-championships"
CALENDAR_URL = BASE_URL + "/championships/get-championships-calendar-by-id"
PLAY_BY_PLAY_URL = BASE_URL + "/championships/get-championships-matches-play-by-play"
SERIE_A_CS_ID = 1
SEASON_FLOOR = 2010
PROVIDER = "lba"

_last_gaps: List[Dict[str, Any]] = []


def last_gaps() -> List[Dict[str, Any]]:
    """Gaps recorded by the most recent call (identity/period problems)."""

    return list(_last_gaps)


def _reset_gaps() -> None:
    _last_gaps.clear()


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def season_start_year(target: datetime) -> int:
    """Season FIRST year for an LBA date (``2026`` == 2026/27)."""

    return target.year if target.month >= 7 else target.year - 1


def _int(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _iso_utc(value: Any, match_id: str) -> str:
    """Parse ``match_datetime`` into UTC, refusing naive wall-clock values."""

    text = str(value or "").strip()
    if not text:
        raise ValueError(f"LBA match {match_id} timestamp is empty")
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"LBA match {match_id} timestamp {text!r} is not ISO 8601") from exc
    if parsed.tzinfo is None:
        # Every captured row carried an explicit offset; a naive value would
        # silently shift the tip, so it is refused rather than assumed.
        raise ValueError(f"LBA match {match_id} timestamp {text!r} has no UTC offset")
    return to_utc_iso(parsed)


def championship_id(start_year: int, fetch: FetchFn) -> int:
    """Resolve Serie A's regular-season championship id for one season."""

    if start_year < SEASON_FLOOR:
        raise ValueError(f"LBA season {start_year} precedes the {SEASON_FLOOR} floor")
    payload = fetch(
        f"{CHAMPIONSHIPS_URL}?cs_id={SERIE_A_CS_ID}&items=1000&s={int(start_year)}"
    )
    if not isinstance(payload, dict):
        raise ValueError(f"LBA championships {start_year} is not a JSON object")
    competitions = payload.get("competitions")
    if not isinstance(competitions, list):
        raise ValueError(f"LBA championships {start_year} has no competitions list")
    candidates = []
    for row in competitions:
        if not isinstance(row, dict) or _int(row.get("id")) is None:
            continue
        name = str(row.get("name") or "")
        if str(row.get("cserie_code") or "") == "A1" and str(start_year) in name:
            candidates.append((_int(row.get("id")), name))
    if len(candidates) != 1:
        raise ValueError(
            f"LBA championships {start_year} is ambiguous: {candidates}; refusing to guess"
        )
    return int(candidates[0][0])


def fetch_calendar(championship_id_value: int, fetch: FetchFn) -> List[Dict[str, Any]]:
    """The league's current focus-day matches for one championship."""

    payload = fetch(f"{CALENDAR_URL}?id={int(championship_id_value)}")
    if not isinstance(payload, dict):
        raise ValueError(f"LBA calendar {championship_id_value} is not a JSON object")
    rows = payload.get("matches")
    if not isinstance(rows, list) or not rows:
        raise ValueError(
            f"LBA calendar {championship_id_value} has no matches; an empty body is "
            "refused rather than treated as a free day"
        )
    return [row for row in rows if isinstance(row, dict)]


def _score_pair(text: Any, match_id: str) -> tuple[int, int]:
    match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", str(text or "").strip())
    if not match:
        raise ValueError(f"LBA match {match_id} has an unparseable pbp score {text!r}")
    return int(match.group(1)), int(match.group(2))


def _periods(pbp: Dict[str, Any], match_id: str) -> ActualResult:
    """Regulation quarters plus overtime from cumulative play-by-play."""

    actions = ((pbp.get("pbp") or {}).get("actions")) or []
    ends: Dict[int, tuple[int, tuple[int, int]]] = {}
    for action in actions:
        if not isinstance(action, dict) or not action.get("score"):
            continue
        period = _int(action.get("period"))
        order = _int(action.get("order"))
        if period is None or order is None:
            continue
        pair = _score_pair(action.get("score"), match_id)
        if period not in ends or order > ends[period][0]:
            ends[period] = (order, pair)
    if not all(period in ends for period in (1, 2, 3, 4)):
        raise ValueError(
            f"LBA match {match_id} play-by-play lacks a cumulative end for all four quarters"
        )
    overtime = sorted(period for period in ends if period > 4)
    if overtime and overtime != list(range(5, 5 + len(overtime))):
        raise ValueError(f"LBA match {match_id} has non-contiguous overtime periods {overtime}")
    ordered = [ends[period][1] for period in sorted(ends)]
    home_final, away_final = ordered[-1]
    quarters = [ordered[0]]
    for previous, current in zip(ordered, ordered[1:]):
        quarters.append((current[0] - previous[0], current[1] - previous[1]))
    home_extra = sum(quarter[0] for quarter in quarters[4:])
    away_extra = sum(quarter[1] for quarter in quarters[4:])
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[float(q[0]) for q in quarters[:4]],
        away_quarters=[float(q[1]) for q in quarters[:4]],
        home_overtime=float(home_extra) if overtime else None,
        away_overtime=float(away_extra) if overtime else None,
        home_final=home_final,
        away_final=away_final,
        overtime_periods=len(overtime),
        source=f"lba play-by-play matchId={match_id}",
    )


def fetch_play_by_play(match_id: str, fetch: FetchFn) -> Dict[str, Any]:
    """One match's play-by-play body (period ends derive quarters)."""

    payload = fetch(f"{PLAY_BY_PLAY_URL}?id={match_id}")
    if not isinstance(payload, dict):
        raise ValueError(f"LBA play-by-play {match_id} is not a JSON object")
    pbp = payload.get("pbp")
    if not isinstance(pbp, dict) or not isinstance(pbp.get("actions"), list):
        raise ValueError(f"LBA play-by-play {match_id} has no actions")
    return payload


def _game_from_row(row: Dict[str, Any], season: int, phase: str) -> Game:
    match_id = str(row.get("id") or "")
    if not match_id:
        raise IdentityError("LBA calendar row lacks a match id")
    home_name = canonical_team_name(str(row.get("h_team_name") or ""))
    away_name = canonical_team_name(str(row.get("v_team_name") or ""))
    if not home_name or not away_name:
        raise IdentityError(f"LBA match {match_id} row lacks a team name")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"LBA match {match_id} names the same team on both sides")
    home_id = str(row.get("h_team_id") or "")
    away_id = str(row.get("v_team_id") or "")
    if not home_id or not away_id or home_id == away_id:
        raise IdentityError(f"LBA match {match_id} row lacks distinct team ids")
    home_final = _int(row.get("home_final_score"))
    away_final = _int(row.get("visitor_final_score"))
    status = (
        GameStatus.FINAL if home_final is not None and away_final is not None
        else GameStatus.SCHEDULED
    )
    return Game(
        game_id=f"LBA-{match_id}",
        league=League.ITALY_LBA,
        season=str(int(season)),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=_iso_utc(row.get("match_datetime"), match_id),
        status=status,
        phase=phase,
        provider_ids={PROVIDER: match_id},
    )


def discover_games(date_iso: Optional[str], fetch: FetchFn, start_year: int) -> List[Game]:
    """Games of one UTC date on the league's focus-day calendar.

    The calendar serves only the league's current focus day (probed day
    selectors are ignored server-side), so a ``date_iso`` outside that day
    yields an honest empty slate rather than a guessed season walk.
    """

    _reset_gaps()
    games: List[Game] = []
    for row in fetch_calendar(championship_id(start_year, fetch), fetch):
        try:
            game = _game_from_row(row, start_year, "Serie A")
        except (IdentityError, ValueError) as exc:
            _record_gap(str(row.get("id") or ""), str(exc))
            continue
        if date_iso and not game.scheduled_utc.startswith(date_iso):
            continue
        games.append(game)
    return games


def fetch_game_result(game: Game, fetch: FetchFn) -> Game:
    """Attach reconciled periods derived from the official play-by-play."""

    match_id = str(game.provider_ids.get(PROVIDER) or "")
    if not match_id:
        raise IdentityError(f"game {game.game_id} carries no LBA match id")
    actual = _periods(fetch_play_by_play(match_id, fetch), match_id)
    ok, reason = actual.reconciles()
    if not ok:
        raise ValueError(f"LBA match {match_id} periods do not reconcile: {reason}")
    refreshed = Game(**{**game.__dict__, "actual": actual, "status": GameStatus.FINAL})
    return refreshed


def smoke(fetch: FetchFn, start_year: int) -> Dict[str, Any]:
    """Coverage summary over the focus-day calendar (never raises)."""

    report: Dict[str, Any] = {
        "games_found": 0,
        "games_final": 0,
        "verification": "FIXTURE_VERIFIED",
        "season_start_year": int(start_year),
    }
    try:
        rows = fetch_calendar(championship_id(start_year, fetch), fetch)
    except Exception as exc:  # smoke never raises
        report["status"] = "provider_error"
        report["error_type"] = type(exc).__name__
        report["detail"] = str(exc)[:200]
        return report
    report["games_found"] = len(rows)
    report["games_final"] = sum(
        1
        for row in rows
        if _int(row.get("home_final_score")) is not None
        and _int(row.get("visitor_final_score")) is not None
    )
    return report
