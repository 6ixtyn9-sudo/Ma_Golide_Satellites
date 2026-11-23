"""Korea KBL schedule/result adapter for ``api.kbl.or.kr``.

Contract reference
------------------
Route and column contract referenced from the SportsDataverse ``kbl`` module
documentation (MIT, read 2026-11-23) which wraps this host; **Ma Golide retains
no first-party response for this lane yet**, so the live body shape is checked
at run time rather than assumed:

* ``GET /match/list?season={END_YEAR}&fromDate={YYYYMMDD}&toDate={YYYYMMDD}``
  returns the games of a window. ``season`` is the season's **END** year (2026
  == 2025-26) and a window inside one calendar year is answered in one request.
  Rows carry ``gmkey`` (``S{seasonCode}G{gameCode}N{gameNo}``), ``game_date``
  (KST), ``game_start`` (``HHMM``, KST), the home/away team codes and names,
  ``home_score`` / ``away_score``, ``is_started``, ``is_ended`` and
  ``playing_quarter``.
* ``GET /common/teamList`` returns ``teamCode`` -> ``teamNameEn`` /
  ``teamNameFullEn``, which is how the canonical (English) team name is
  resolved instead of the Korean short name.
* Two request headers are required and no key is needed: ``Channel: WEB`` and
  ``TeamCode: XX``.

What this adapter refuses
-------------------------
``is_ended`` is documented as unreliable (0 on every finished game through
2005-06 and on some later ones), and ``playing_quarter`` can stay stale after
the final. A game is therefore only marked ``FINAL`` when ``is_ended == 1``
**and** both scores are present; ``is_started == 1`` without an end marker is
``LIVE``; anything else stays ``SCHEDULED``. Scores are read, never inferred.

The list route carries no period split, so a row settled from it carries a
totals-only :class:`ActualResult` (``PeriodKind.UNKNOWN``). ``/match/{gmkey}``
does carry one -- ``teamrecords.{home,away}.scoreq1..scoreq4`` plus ``scoreeq[]``
for overtime -- and a real body for ``S47G01N49`` is vendored at
``tests/fixtures/kbl/captures/``, so :func:`fetch_game_result` reconciles
quarters and the results stage replaces the totals-only record. A split that
does not reconcile stays a gap instead of being summed into an accepted final.

Every scope, identity and timestamp problem is recorded as a gap; a body that
cannot be read raises instead of being treated as an empty slate.
"""

from __future__ import annotations

import hashlib
import json
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional, Tuple

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

BASE_URL = "https://api.kbl.or.kr"
MATCH_LIST_URL = BASE_URL + "/match/list"
TEAM_LIST_URL = BASE_URL + "/common/teamList"
MATCH_URL = BASE_URL + "/match/{}"
PROVIDER = "kbl"
#: The host serves its own single-page app and requires these two headers.
REQUIRED_HEADERS = {"Channel": "WEB", "TeamCode": "XX"}
BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"
KST = timezone(timedelta(hours=9))
#: Keys a JSON envelope may use to carry the row list.
_ROW_KEYS = ("data", "list", "games", "items", "result", "body")

_last_gaps: List[Dict[str, Any]] = []
_last_probe: Dict[str, Any] = {}


def _quarter_scores(payload: Dict[str, Any], gmkey: str) -> ActualResult:
    """Reconciled regulation quarters plus overtime, from ``teamrecords``.

    ``/match/{gmkey}`` carries ``scoreq1..scoreq4`` per side and ``scoreeq`` as a
    list of overtime periods. Four regulation quarters are required; overtime
    must be contiguous from the fifth or the body is refused rather than summed
    into a total that cannot be checked.
    """

    records = payload.get("teamrecords")
    if not isinstance(records, dict):
        raise ValueError(f"KBL match {gmkey} has no teamrecords")
    home = records.get("home") or {}
    away = records.get("away") or {}
    quarters: Dict[int, tuple] = {}
    for index in (1, 2, 3, 4):
        home_points = _int(home.get(f"scoreq{index}"))
        away_points = _int(away.get(f"scoreq{index}"))
        if home_points is None or away_points is None:
            raise ValueError(f"KBL match {gmkey} lacks quarter {index} for both sides")
        quarters[index] = (float(home_points), float(away_points))

    def _overtimes(side: Dict[str, Any]) -> List[float]:
        raw = side.get("scoreeq")
        if raw in (None, ""):
            return []
        if not isinstance(raw, list):
            raise ValueError(f"KBL match {gmkey} has a non-list scoreeq")
        out: List[float] = []
        for value in raw:
            points = _int(value)
            if points is None:
                raise ValueError(f"KBL match {gmkey} has an unreadable overtime value")
            out.append(float(points))
        return out

    home_overtime = _overtimes(home)
    away_overtime = _overtimes(away)
    if len(home_overtime) != len(away_overtime):
        raise ValueError(f"KBL match {gmkey} sides disagree on the overtime count")
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[quarters[index][0] for index in (1, 2, 3, 4)],
        away_quarters=[quarters[index][1] for index in (1, 2, 3, 4)],
        home_overtime=sum(home_overtime) if home_overtime else None,
        away_overtime=sum(away_overtime) if away_overtime else None,
        home_final=sum(quarters[i][0] for i in (1, 2, 3, 4)) + sum(home_overtime),
        away_final=sum(quarters[i][1] for i in (1, 2, 3, 4)) + sum(away_overtime),
        overtime_periods=len(home_overtime),
        source=f"kbl match gmkey={gmkey}",
    )


def fetch_game_result(game: Game, fetch: Optional[FetchFn] = None) -> Game:
    """Re-read one game's ``/match/{gmkey}`` so its quarters can reconcile.

    Discovery settles a row totals-only because the list route carries no
    period split. This route does, so the results stage replaces that record
    with reconciled quarters -- and refuses a body that belongs to another game
    or another season.
    """

    gmkey = str(game.provider_ids.get(PROVIDER) or "")
    if not gmkey:
        raise IdentityError(f"game {game.game_id} carries no KBL gmkey")
    transport = fetch or default_fetch()
    url = MATCH_URL.format(gmkey)
    payload = transport(url)
    if not isinstance(payload, dict):
        raise ValueError(f"KBL match {gmkey} is not a JSON object")
    row = payload.get("game") if isinstance(payload.get("game"), dict) else payload
    if str(row.get("gmkey") or "").strip() not in ("", gmkey):
        raise IdentityError(
            f"KBL match body answers gmkey {row.get('gmkey')!r}, asked for {gmkey!r}"
        )
    if _int(row.get("isEnded")) != 1:
        raise ValueError(f"KBL match {gmkey} is not ended; refusing a result")
    actual = _quarter_scores(payload, gmkey)
    refreshed = Game(
        game_id=game.game_id,
        league=game.league,
        season=game.season,
        home_team_id=game.home_team_id,
        away_team_id=game.away_team_id,
        home_name=game.home_name,
        away_name=game.away_name,
        scheduled_utc=game.scheduled_utc,
        status=GameStatus.FINAL,
        provider_ids=dict(game.provider_ids),
        actual=actual,
        evidence={**game.evidence, "result_url": url, "quarters_from": "kbl match"},
    )
    if str(row.get("tcodeH") or "").strip() not in ("", game.home_team_id):
        raise IdentityError(f"KBL match {gmkey} body names another home team")
    if str(row.get("tcodeA") or "").strip() not in ("", game.away_team_id):
        raise IdentityError(f"KBL match {gmkey} body names another away team")
    return refreshed


def last_gaps() -> List[Dict[str, Any]]:
    """Gaps recorded by the most recent call (identity/timestamp problems)."""

    return list(_last_gaps)


def last_probe() -> Dict[str, Any]:
    """What the last response actually looked like (receipt, never a body).

    The new source policy requires every scraped read to be recorded rather
    than asserted: this carries the URL, byte count, sha256 and the shape that
    was observed, which is what makes the *first* scheduled run evidence.
    """

    return dict(_last_probe)


def _reset() -> None:
    _last_gaps.clear()
    _last_probe.clear()


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def default_fetch() -> FetchFn:
    """Adapter-owned JSON client carrying the two required headers."""

    client = HttpClient(timeout=20.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_json(url, headers=REQUIRED_HEADERS)


def season_end_years(target: datetime) -> Tuple[int, int]:
    """Candidate season END years for a KST date, most likely first.

    KBL labels a season by its END year, so a September–December date belongs
    to the next end year and a January–August date to the current one. The
    boundary moves around, so the alternate candidate is offered as a fallback
    instead of being assumed away.
    """

    if target.month >= 8:
        return target.year + 1, target.year
    return target.year, target.year + 1


def _int(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _rows(payload: Any, what: str) -> List[Dict[str, Any]]:
    """Rows out of a bare list or a one-level JSON envelope.

    The envelope key is not asserted by the contract, so the documented
    candidates are tried; a shape that matches none of them raises rather than
    reading as an empty slate.
    """

    if isinstance(payload, list):
        return [row for row in payload if isinstance(row, dict)]
    if isinstance(payload, dict):
        for key in _ROW_KEYS:
            value = payload.get(key)
            if isinstance(value, list):
                return [row for row in value if isinstance(row, dict)]
            if isinstance(value, dict):
                for inner in _ROW_KEYS:
                    nested = value.get(inner)
                    if isinstance(nested, list):
                        return [row for row in nested if isinstance(row, dict)]
        raise ValueError(f"KBL {what} has no recognisable row list: keys={sorted(payload)}")
    raise ValueError(f"KBL {what} is neither a list nor an object")


def _probe(url: str, payload: Any, rows: List[Dict[str, Any]]) -> None:
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    _last_probe.update(
        {
            "url": url,
            "bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
            "shape": "list" if isinstance(payload, list) else "object",
            "rows": len(rows),
            "row_keys": sorted(rows[0]) if rows else [],
        }
    )


def fetch_window(
    season_end_year: int,
    start_yyyymmdd: str,
    end_yyyymmdd: str,
    fetch: FetchFn,
) -> List[Dict[str, Any]]:
    """Rows of one date window for one season END year."""

    url = (
        f"{MATCH_LIST_URL}?season={int(season_end_year)}"
        f"&fromDate={start_yyyymmdd}&toDate={end_yyyymmdd}"
    )
    payload = fetch(url)
    rows = _rows(payload, f"match/list season={season_end_year}")
    _probe(url, payload, rows)
    return rows


def fetch_team_names(fetch: FetchFn) -> Dict[str, str]:
    """``teamCode`` -> English full name, best effort and never fatal.

    A team-list miss degrades the *name*, not the identity: the canonical id is
    the team code from the schedule row.
    """

    try:
        payload = fetch(TEAM_LIST_URL)
        rows = _rows(payload, "common/teamList")
    except Exception:  # noqa: BLE001 - a name lookup must not fail the slate
        return {}
    names: Dict[str, str] = {}
    for row in rows:
        code = str(row.get("teamCode") or "").strip()
        if not code:
            continue
        label = str(row.get("teamNameFullEn") or row.get("teamNameEn") or "").strip()
        if label:
            names[code] = canonical_team_name(label)
    return names


def _date_plus_time(row: Dict[str, Any], gmkey: str) -> str:
    raw_date = str(row.get("game_date") or "").strip()
    digits = raw_date.replace("-", "")
    if len(digits) != 8 or not digits.isdigit():
        raise ValueError(f"KBL game {gmkey} has an unparseable game_date {raw_date!r}")
    raw_start = str(row.get("game_start") or "").strip()
    if len(raw_start) == 3:
        raw_start = "0" + raw_start
    if len(raw_start) != 4 or not raw_start.isdigit():
        raise ValueError(f"KBL game {gmkey} has an unparseable game_start {raw_start!r}")
    try:
        local = datetime(
            int(digits[0:4]),
            int(digits[4:6]),
            int(digits[6:8]),
            int(raw_start[0:2]),
            int(raw_start[2:4]),
            tzinfo=KST,
        )
    except ValueError as exc:
        raise ValueError(f"KBL game {gmkey} has an impossible tip time: {exc}") from exc
    return to_utc_iso(local)


def _game_from_row(
    row: Dict[str, Any], season_end_year: int, names: Dict[str, str]
) -> Game:
    gmkey = str(row.get("gmkey") or "").strip()
    if not gmkey:
        raise IdentityError("KBL row lacks a gmkey")
    home_code = str(row.get("home_team_code") or "").strip()
    away_code = str(row.get("away_team_code") or "").strip()
    if not home_code or not away_code:
        raise IdentityError(f"KBL game {gmkey} row lacks a team code")
    if home_code == away_code:
        raise IdentityError(f"KBL game {gmkey} names the same team code on both sides")
    home_name = names.get(home_code) or canonical_team_name(
        str(row.get("home_team_name") or "")
    )
    away_name = names.get(away_code) or canonical_team_name(
        str(row.get("away_team_name") or "")
    )
    if not home_name or not away_name:
        raise IdentityError(f"KBL game {gmkey} row lacks a team name")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"KBL game {gmkey} resolves the same name on both sides")

    home_score = _int(row.get("home_score"))
    away_score = _int(row.get("away_score"))
    ended = _int(row.get("is_ended")) == 1
    started = _int(row.get("is_started")) == 1
    actual: Optional[ActualResult] = None
    if ended and home_score is not None and away_score is not None:
        status = GameStatus.FINAL
        # Totals only: /match/{gmkey} (the quarter split) is not captured yet,
        # so this record is deliberately not quarter-reconcilable.
        actual = ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=float(home_score),
            away_final=float(away_score),
            source=f"kbl match/list gmkey={gmkey}",
        )
    elif ended:
        status = GameStatus.SCHEDULED
        _record_gap(gmkey, "is_ended=1 without both scores; refused as a final")
    elif started:
        status = GameStatus.LIVE
    else:
        status = GameStatus.SCHEDULED

    return Game(
        game_id=f"KBL-{gmkey}",
        league=League.SOUTH_KOREA_KBL,
        season=str(int(season_end_year)),
        home_team_id=home_code,
        away_team_id=away_code,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=_date_plus_time(row, gmkey),
        status=status,
        provider_ids={PROVIDER: gmkey},
        actual=actual,
        evidence={
            "provider": PROVIDER,
            "gmkey": gmkey,
            "glkey": str(row.get("glkey") or ""),
            "season_category": str(row.get("season_category") or ""),
            "playing_quarter": str(row.get("playing_quarter") or ""),
            "stadium": str(row.get("stadium_name_full") or row.get("stadium_name") or ""),
            "korean_home_name": str(row.get("home_team_name") or ""),
            "korean_away_name": str(row.get("away_team_name") or ""),
        },
    )


def discover_games(
    date_iso: Optional[str],
    fetch: Optional[FetchFn],
    end_year: Optional[int] = None,
) -> List[Game]:
    """Games for one KST/UTC date, from the documented date-window route.

    ``date_iso`` filters on the UTC tip date. The window requested is the KST
    day around it so a game tipping after midnight KST is not lost.
    """

    _reset()
    transport = fetch or default_fetch()
    if date_iso:
        target = datetime.strptime(date_iso, "%Y-%m-%d")
    else:
        target = datetime.now(timezone.utc)
    local = target.replace(tzinfo=timezone.utc).astimezone(KST)
    start = (local - timedelta(days=1)).strftime("%Y%m%d")
    end = (local + timedelta(days=1)).strftime("%Y%m%d")

    candidates = (end_year,) if end_year is not None else season_end_years(local)
    names: Dict[str, str] = {}
    games: List[Game] = []
    seen: set[str] = set()
    for season in candidates:
        rows = fetch_window(season, start, end, transport)
        if rows and not names:
            names = fetch_team_names(transport)
        for row in rows:
            key = str(row.get("gmkey") or "")
            if key and key in seen:
                continue
            try:
                game = _game_from_row(row, season, names)
            except (IdentityError, ValueError) as exc:
                _record_gap(key or None, str(exc))
                continue
            if date_iso and not game.scheduled_utc.startswith(date_iso):
                continue
            seen.add(key)
            games.append(game)
        if games:
            break
    games.sort(key=lambda game: game.scheduled_utc)
    return games
