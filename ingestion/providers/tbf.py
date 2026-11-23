"""Turkey TBF schedule/result adapter for ``miniappapi.tbf.org.tr``.

Contract reference
------------------
Route and column contract referenced from the SportsDataverse ``tbf`` module
(MIT, read 2026-11-23) which wraps the federation's own mini-app API. Unlike
KBL, a **real response body is vendored** for this lane
(``tests/fixtures/tbf/captures/get-daily-matches-2026-10-11.json``, captured by
the upstream fixture set for 2026-10-11 and attributed in
``tests/fixtures/CAPTURE-PROVENANCE.md``), so the parser below is checked
against a body the host actually served rather than a reconstruction of one.

* ``GET /Match/get-daily-matches?match_date={YYYY-MM-DD}`` answers every game
  the federation lists for one **Turkey local date**, grouped by competition:

  ``{"data": [{"groupName": ..., "totalCount": N,
              "matches": [<game row>, ...]}, ...],
     "isSuccess": true, "message": "Success", "problemDetails": ...}``

  A game row carries ``matchId``, ``matchCode``, ``genuisId``, ``activityCode``
  (``bsl`` / ``tbl`` / ``kbsl`` / ``tkbl``), ``activityName``, ``season``
  (label form ``2026-2027``), ``matchDate`` (naive local ``YYYY-MM-DDTHH:MM:SS``),
  ``matchDateOnly``, ``matchTime``, ``week``, ``venueName``,
  ``homeTeam``/``awayTeam`` (``{id, name, score, isHome, ...}``),
  ``matchStatusId`` / ``matchStatus`` (``1`` = played, ``"Oynandı"``) and
  ``hasUzatma``.

* One route covers the top flight (``bsl``), the second tier (``tbl``), the
  women's Super Ligi (``kbsl``) and its second tier (``tkbl``); the competition
  is selected from ``activityCode`` on the row, which is why a single adapter
  serves them all.
* No key and no special header is required. ``Accept: application/json`` and a
  browser user-agent are sent anyway.

What this adapter refuses
-------------------------
``matchStatusId == 1`` is the only settled marker this adapter trusts, and only
when **both** scores parse as integers; a played row without both scores is
refused as a final and recorded as a gap. No live marker is documented for this
route, so a row that is not played stays ``SCHEDULED`` rather than being guessed
into ``LIVE``.

The daily route carries no quarter split, so a settled row stores a totals-only
:class:`ActualResult` (``PeriodKind.UNKNOWN``) and is **not
quarter-reconcilable**. ``hasUzatma`` is recorded as evidence only: it says an
overtime happened, which is not the same as knowing the period scores.

Every scope, identity and timestamp problem is recorded as a gap; a body that
cannot be read raises instead of being treated as an empty slate.
"""

from __future__ import annotations

import hashlib
import json
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

PROVIDER = "tbf"

BASE_URL = "https://miniappapi.tbf.org.tr/webapi-service/api"
DAILY_MATCHES_URL = BASE_URL + "/Match/get-daily-matches"

#: ``activityCode`` values known to appear in the daily body. The lane's own
#: competition is passed explicitly; the rest are here so a caller can widen the
#: filter without guessing the strings.
COMPETITION_BSL = "bsl"  # Basketbol Süper Ligi (top flight; named TBL before 2007-08)
COMPETITION_TBL = "tbl"  # second tier
COMPETITION_KBSL = "kbsl"  # women's Super Ligi
COMPETITION_TKBL = "tkbl"  # women's second tier
COMPETITIONS = (COMPETITION_BSL, COMPETITION_TBL, COMPETITION_KBSL, COMPETITION_TKBL)

#: ``matchStatusId`` 1 == played (``matchStatus`` "Oynandı"). No live marker is
#: documented for this route; anything else is treated as not yet played.
STATUS_PLAYED = 1

#: Turkey is UTC+03:00 all year (no DST since 2016), so the local offset is a
#: constant rather than a zone lookup.
TRT = timezone(timedelta(hours=3), "Europe/Istanbul")

BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"

REQUIRED_HEADERS = {"Accept": "application/json"}

_last_gaps: List[Dict[str, Any]] = []
_last_probe: Dict[str, Any] = {}


def _reset() -> None:
    _last_gaps.clear()
    _last_probe.clear()


def last_gaps() -> List[Dict[str, Any]]:
    return [dict(gap) for gap in _last_gaps]


def last_probe() -> Dict[str, Any]:
    """What the last daily response actually looked like (never a body)."""

    return dict(_last_probe)


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def _record_probe(url: str, payload: Dict[str, Any], groups: int, rows: int) -> None:
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    competitions = sorted(
        {
            str(match.get("activityCode") or "").strip().lower()
            for group in payload.get("data") or []
            if isinstance(group, dict)
            for match in group.get("matches") or []
            if isinstance(match, dict)
        }
    )
    _last_probe.update(
        {
            "url": url,
            "bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
            "groups": groups,
            "rows": rows,
            "competitions": competitions,
            "is_success": payload.get("isSuccess"),
        }
    )


def default_fetch() -> FetchFn:
    """Adapter-owned JSON client for the federation's mini-app API."""

    client = HttpClient(timeout=20.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_json(url, headers=REQUIRED_HEADERS)


def season_end_year(label: Any) -> Optional[int]:
    """``"2026-2027"`` -> 2027, matching how the lane keys a season."""

    text = str(label or "").strip()
    if len(text) >= 4 and text[:4].isdigit():
        tail = text.split("-")[-1]
        if tail[:4].isdigit():
            return int(tail[:4])
        return int(text[:4])
    return None


def _int(value: Any) -> Optional[int]:
    """Scores and ids arrive as ``"85"``, ``85.0`` or ``1.0``; never guess."""

    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def _tip_utc(row: Dict[str, Any], match_id: str) -> str:
    """``matchDate`` is naive Turkey local time; refuse anything else."""

    raw = str(row.get("matchDate") or "").strip()
    if raw:
        try:
            parsed = datetime.fromisoformat(raw.replace("Z", ""))
        except ValueError as exc:
            raise ValueError(f"TBF match {match_id} matchDate {raw!r} is not ISO: {exc}") from exc
        if parsed.tzinfo is not None:
            return to_utc_iso(parsed)
        return to_utc_iso(parsed.replace(tzinfo=TRT))

    date_only = str(row.get("matchDateOnly") or "").strip()
    time_only = str(row.get("matchTime") or "").strip()
    if not date_only or not time_only:
        raise ValueError(f"TBF match {match_id} carries no date/time")
    try:
        parsed = datetime.strptime(f"{date_only} {time_only.zfill(5)}", "%Y-%m-%d %H:%M")
    except ValueError as exc:
        raise ValueError(f"TBF match {match_id} date/time is unreadable: {exc}") from exc
    return to_utc_iso(parsed.replace(tzinfo=TRT))


def _team(side: Dict[str, Any], match_id: str, label: str) -> tuple[str, str]:
    if not isinstance(side, dict):
        raise IdentityError(f"TBF match {match_id} lacks a {label} team object")
    team_id = _int(side.get("id"))
    name = canonical_team_name(str(side.get("name") or ""))
    if team_id is None:
        raise IdentityError(f"TBF match {match_id} {label} team lacks an id")
    if not name:
        raise IdentityError(f"TBF match {match_id} {label} team lacks a name")
    return str(team_id), name


def _game_from_row(row: Dict[str, Any], competition: Optional[str]) -> Game:
    match_id = _int(row.get("matchId"))
    if match_id is None:
        raise IdentityError("TBF row lacks a matchId")
    key = str(match_id)
    code = str(row.get("activityCode") or "").strip().lower()
    if competition is not None and code != competition:
        raise IdentityError(f"TBF match {key} is activityCode {code!r}, not {competition!r}")

    home_id, home_name = _team(row.get("homeTeam") or {}, key, "home")
    away_id, away_name = _team(row.get("awayTeam") or {}, key, "away")
    if home_id == away_id:
        raise IdentityError(f"TBF match {key} names the same team id on both sides")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"TBF match {key} resolves the same name on both sides")

    played = _int(row.get("matchStatusId")) == STATUS_PLAYED
    home_score = _int((row.get("homeTeam") or {}).get("score"))
    away_score = _int((row.get("awayTeam") or {}).get("score"))
    actual: Optional[ActualResult] = None
    if played and home_score is not None and away_score is not None:
        status = GameStatus.FINAL
        # Totals only: the daily route carries no quarter split, so this record
        # is deliberately not quarter-reconcilable.
        actual = ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=float(home_score),
            away_final=float(away_score),
            source=f"tbf get-daily-matches matchId={key}",
        )
    elif played:
        status = GameStatus.SCHEDULED
        _record_gap(key, "matchStatusId=1 without both scores; refused as a final")
    else:
        status = GameStatus.SCHEDULED

    return Game(
        game_id=f"TBF-{key}",
        league=League.TURKEY_BSL,
        season=str(season_end_year(row.get("season")) or ""),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=_tip_utc(row, key),
        status=status,
        provider_ids={PROVIDER: key},
        actual=actual,
        evidence={
            "provider": PROVIDER,
            "activity_code": code,
            "activity_name": row.get("activityName"),
            "match_code": row.get("matchCode"),
            "week": row.get("week"),
            "venue": row.get("venueName"),
            "status_text": row.get("matchStatus"),
            "has_overtime": row.get("hasUzatma"),
            "url": DAILY_MATCHES_URL,
        },
    )


def _rows(payload: Any) -> List[Dict[str, Any]]:
    """Flatten the grouped envelope; an unreadable body is never an empty slate."""

    if not isinstance(payload, dict):
        raise ValueError("TBF daily-matches body is not a JSON object")
    if payload.get("isSuccess") is False:
        raise ValueError(
            f"TBF daily-matches answered isSuccess=false: {payload.get('message')!r}"
        )
    groups = payload.get("data")
    if not isinstance(groups, list):
        raise ValueError("TBF daily-matches has no data list")
    rows: List[Dict[str, Any]] = []
    for group in groups:
        if not isinstance(group, dict):
            continue
        matches = group.get("matches")
        if isinstance(matches, list):
            rows.extend(match for match in matches if isinstance(match, dict))
    return rows


def discover_games(
    date_iso: Optional[str],
    fetch: Optional[FetchFn] = None,
    competition: Optional[str] = COMPETITION_BSL,
) -> List[Game]:
    """Games of one UTC date for one competition (``None`` = every competition).

    The host is asked for a **Turkey local date**, so a UTC target day is read
    through the two local dates that can contain a tip inside it; rows whose UTC
    tip falls outside ``date_iso`` are dropped afterwards.
    """

    _reset()
    transport = fetch or default_fetch()
    if date_iso:
        target = datetime.strptime(date_iso, "%Y-%m-%d")
    else:
        target = datetime.now(timezone.utc).replace(tzinfo=None)
    local_day = (target.replace(tzinfo=timezone.utc) + timedelta(hours=3)).date()

    wanted = competition.strip().lower() if competition else None
    games: List[Game] = []
    seen: set[str] = set()
    for offset in (0, 1):
        day = local_day + timedelta(days=offset)
        url = f"{DAILY_MATCHES_URL}?match_date={day.isoformat()}"
        payload = transport(url)
        rows = _rows(payload)
        _record_probe(url, payload, len((payload.get("data") or [])), len(rows))
        for row in rows:
            code = str(row.get("activityCode") or "").strip().lower()
            if wanted is not None and code != wanted:
                continue
            try:
                game = _game_from_row(row, wanted)
            except (IdentityError, ValueError) as exc:
                _record_gap(str(_int(row.get("matchId")) or ""), str(exc))
                continue
            if date_iso and not game.scheduled_utc.startswith(date_iso):
                continue
            if game.game_id in seen:
                continue
            seen.add(game.game_id)
            games.append(game)
    return games


def fetch_game_result(game: Game, fetch: Optional[FetchFn] = None) -> Game:
    """Re-read one game's daily row to attach a settled score.

    The daily route is the only captured route for this lane, so this is the
    same read as discovery restricted to one id; it stays totals-only.
    """

    match_id = str(game.provider_ids.get(PROVIDER) or "")
    if not match_id:
        raise IdentityError(f"game {game.game_id} carries no TBF match id")
    transport = fetch or default_fetch()
    for offset in (-1, 0, 1):
        day = (datetime.now(timezone.utc) + timedelta(days=offset)).date()
        for row in _rows(transport(f"{DAILY_MATCHES_URL}?match_date={day.isoformat()}")):
            if _int(row.get("matchId")) == _int(match_id):
                return _game_from_row(row, None)
    raise ValueError(f"TBF match {match_id} is not on any recent daily slate")
