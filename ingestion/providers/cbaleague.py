"""Official CBA data-server adapter (keyless; AES-128-ECB decrypted bodies).

Evidence status (see docs/PROVIDERS.md)
--------------------------------------------------
``LIVE_VERIFIED (live captures 2026-11-16)``: the endpoints below were
fetched and decrypted live on 2026-11-16 and captured into
``tests/fixtures/cbaleague/``. Verified conventions:

- ``GET com-code-tables/getSeason`` lists seasons as
  ``{"key": "2026-2027", "value": "2026", "isDefault": n}``; the API's
  season value is the START year and ``isDefault == 1`` marks the current
  season. Seasons reach back to 1995 (live observation).
- ``GET schedules/day?auto=0&day=YYYY-MM-DD`` answers ``[]`` on days
  without games and otherwise rows with ``scheduleId``, ``season`` (START
  year), ``homeTeamId``/``visitingTeamId``, Chinese team display names
  (season sponsor names, e.g. ``长白山恩都里`` = the Jilin club in 2026-27),
  ``status`` (1 = scheduled, 4 = final in every observed row), and
  ``scheduleTime`` as Beijing-local ``YYYY-MM-DD HH:MM:SS``.
- ``POST schedule-quarters/getQuarterInfo`` with ``{"scheduleId": n}``
  answers period scores ``oneScore..sevenScore`` per side: four regulation
  quarters plus up to three overtime periods. Verified sums equal the
  finals exactly (79 and 90; the OT sample 97/99 with ``fiveScore`` 10/12).
- Every body is a JSON string literal holding base64 of AES-128-ECB/PKCS7
  ciphertext under the public key constant shipped in the site's own
  front-end bundle (function ``TIe``); the stdlib decryptor is vendored in
  ``_cba_aes.py``. A body that does not decrypt is refused, never banked.

This is the data site's own undocumented API: keyless and publicly served,
but source terms have not been reviewed. No credentials, proxies, or access
controls are used; public access does not establish permission to automate,
store, or redistribute the data. Team display names change with sponsors, so
known variants map through the canonical alias table. Unknown Chinese names
remain unchanged and must not be assumed to identify a stable franchise for
historical modeling or cross-provider joins.
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional
from zoneinfo import ZoneInfo

from ..identity import canonical_team_name, to_utc_iso
from ..models import (
    ActualResult,
    CoverageGap,
    CoverageState,
    Game,
    GameStatus,
    League,
    PeriodKind,
)
from ._cba_aes import Undecryptable, decrypt_body
from . import FetchFn, IdentityError

PostFn = Callable[[str, Dict[str, Any]], str]

PROVIDER = "cbaleague"
BASE_URL = "https://data-server.cbaleague.com/api"
REFERER = "https://www.cbaleague.com/data/"
DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"
EVIDENCE = "LIVE_VERIFIED (live captures 2026-11-16; source terms unreviewed)"
LEAGUE_TZ = ZoneInfo("Asia/Shanghai")

#: Wire statuses observed live: 1 = scheduled, 4 = final. Anything else
#: (postponed/cancelled codes are not publicly documented) stays UNKNOWN.
STATUS_MAP = {1: GameStatus.SCHEDULED, 4: GameStatus.FINAL}
OT_FIELDS = ("fiveScore", "sixScore", "sevenScore")
QUARTER_FIELDS = ("oneScore", "twoScore", "threeScore", "fourScore")
_MAX_OT_PERIODS = 3

_LAST_GAPS: List[CoverageGap] = []


def is_enabled() -> bool:
    return True  # keyless official public API


def last_gaps() -> List[CoverageGap]:
    return list(_LAST_GAPS)


def last_quota() -> Dict[str, Any]:
    return {}


def _reset_gaps() -> None:
    _LAST_GAPS.clear()


def _gap(game_id: Optional[str], field: str, detail: str) -> None:
    _LAST_GAPS.append(
        CoverageGap(
            game_id=game_id,
            required_field=field,
            reason=CoverageState.MISSING,
            provider=PROVIDER,
            detail=detail,
        )
    )


def _require_cba(league: League | str) -> League:
    value = league.value if isinstance(league, League) else str(league).strip().upper()
    if value != "CBA":
        raise ValueError(f"{PROVIDER} is enabled only for Chinese CBA, not {value!r}")
    return League.CBA


def _decrypt(text: Any) -> Any:
    if not isinstance(text, str):
        raise IdentityError(f"{PROVIDER}: expected a raw encrypted body string")
    try:
        return decrypt_body(text, _key())
    except Undecryptable as exc:
        raise IdentityError(
            f"{PROVIDER}: body did not decrypt ({exc}); the site key may have rotated"
        ) from exc


def _key() -> bytes:
    """The public bundle constant, overridable for tests/key rotation."""
    env = os.environ.get("CBA_DECRYPT_KEY", "")
    return env.encode() if env else b"uVayqL4ONKjFbVzQ"


def _headers() -> Dict[str, str]:
    return {"Referer": REFERER}


def _iso_date(date_yyyymmdd: str) -> str:
    text = str(date_yyyymmdd).replace("-", "")
    datetime.strptime(text, "%Y%m%d")
    return f"{text[:4]}-{text[4:6]}-{text[6:8]}"


def available_seasons(fetch: FetchFn) -> List[int]:
    """Season start years exposed by the official season-list endpoint."""
    rows = _decrypt(fetch(f"{BASE_URL}/com-code-tables/getSeason"))
    if not isinstance(rows, list):
        raise IdentityError("getSeason did not answer a list")
    seasons = sorted(
        {
            int(row["value"])
            for row in rows
            if isinstance(row, dict) and str(row.get("value") or "").isdigit()
        }
    )
    if not seasons:
        raise IdentityError("getSeason returned no parseable season years")
    return seasons


def current_season(fetch: FetchFn) -> int:
    """The current season's START year (``isDefault == 1`` row)."""
    rows = _decrypt(fetch(f"{BASE_URL}/com-code-tables/getSeason"))
    if not isinstance(rows, list):
        raise IdentityError("getSeason did not answer a list")
    defaults = [
        row
        for row in rows
        if isinstance(row, dict)
        and row.get("isDefault") == 1
        and str(row.get("value") or "").isdigit()
    ]
    if len(defaults) != 1:
        raise IdentityError(
            f"getSeason returned {len(defaults)} current-season markers; refusing to guess"
        )
    return int(defaults[0]["value"])


def day_rows(day_iso: str, fetch: FetchFn) -> List[Dict[str, Any]]:
    """Decrypted schedule rows for one Beijing calendar day."""
    payload = _decrypt(fetch(f"{BASE_URL}/schedules/day?auto=0&day={day_iso}"))
    if not isinstance(payload, list):
        raise IdentityError(f"schedules/day {day_iso} did not answer a list")
    return [row for row in payload if isinstance(row, dict)]


def _scheduled_utc(value: Any) -> str:
    text = str(value or "").strip()
    try:
        local = datetime.strptime(text, "%Y-%m-%d %H:%M:%S").replace(tzinfo=LEAGUE_TZ)
    except ValueError:
        raise IdentityError(f"invalid scheduleTime {value!r}") from None
    return to_utc_iso(local)


def _team_name(value: Any, role: str, game_ref: str) -> str:
    name = str(value or "").strip()
    if not name:
        raise IdentityError(f"CBA {game_ref} lacks {role} team name")
    return canonical_team_name(name)


def _period(row: Dict[str, Any], key: str, game_ref: str) -> Optional[float]:
    value = row.get(key)
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise IdentityError(f"CBA {game_ref} has invalid {key}: {value!r}")
    value = float(value)
    if value < 0:
        raise IdentityError(f"CBA {game_ref} has negative {key}: {value!r}")
    return value


def _quarters_from_info(info: Dict[str, Any], side: str, game_ref: str):
    block = info.get(side)
    if not isinstance(block, dict):
        raise IdentityError(f"CBA {game_ref} quarter info lacks the {side} block")
    quarters = [_period(block, key, game_ref) for key in QUARTER_FIELDS]
    ot_values = [_period(block, key, game_ref) for key in OT_FIELDS[:_MAX_OT_PERIODS]]
    numeric_ot = [value for value in ot_values if value is not None]
    overtime = sum(numeric_ot) if numeric_ot else None
    overtime_periods = len(numeric_ot)
    return quarters, overtime, overtime_periods


def _actual(schedule_id: str, info: Dict[str, Any], final_home, final_away) -> ActualResult:
    home_quarters, home_ot, ot_periods = _quarters_from_info(info, "home", schedule_id)
    away_quarters, away_ot, ot_periods_away = _quarters_from_info(info, "visiting", schedule_id)
    if ot_periods != ot_periods_away:
        raise IdentityError(f"CBA {schedule_id}: sides disagree on overtime period count")
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=final_home,
        away_final=final_away,
        overtime_periods=ot_periods,
        source=f"cbaleague quarterinfo scheduleId={schedule_id}",
    )
    ok, reason = actual.reconciles()
    if not ok:
        raise IdentityError(f"CBA {schedule_id} failed reconciliation: {reason}")
    return actual


def _game_from_row(row: Dict[str, Any]) -> Game:
    schedule_id = str(row.get("scheduleId") or "").strip()
    if not schedule_id:
        raise IdentityError("CBA schedule row lacks scheduleId")
    home_id = str(row.get("homeTeamId") or "").strip()
    away_id = str(row.get("visitingTeamId") or "").strip()
    if not home_id or not away_id:
        raise IdentityError(f"CBA {schedule_id} lacks team ids")
    if home_id == away_id:
        raise IdentityError(f"CBA {schedule_id}: home and away team ids coincide")
    status = STATUS_MAP.get(row.get("status"), GameStatus.UNKNOWN)
    scheduled_utc = _scheduled_utc(row.get("scheduleTime"))
    season = row.get("season")
    return Game(
        game_id=schedule_id,
        league=League.CBA,
        season=str(season) if season is not None else "",
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=_team_name(row.get("homeTeamName"), "home", schedule_id),
        away_name=_team_name(row.get("visitingTeamName"), "visiting", schedule_id),
        scheduled_utc=scheduled_utc,
        status=status,
        phase=str(row.get("scheduleTypeName") or "League"),
        provider_ids={PROVIDER: schedule_id},
    )


def discover_games(
    date_yyyymmdd: str, fetch: FetchFn, league: League | str = League.CBA
) -> List[Game]:
    """Discover the day's games from the official Beijing-calendar slate."""
    _require_cba(league)
    _reset_gaps()
    games: List[Game] = []
    for row in day_rows(_iso_date(date_yyyymmdd), fetch):
        try:
            games.append(_game_from_row(row))
        except IdentityError as exc:
            _gap(str(row.get("scheduleId") or "") or None, "schedule_row", str(exc))
    return games


def fetch_recent_results(
    date_yyyymmdd: str,
    fetch: FetchFn,
    league: League | str = League.CBA,
    post_fn: Optional[PostFn] = None,
) -> List[Game]:
    """Reconciled finals for one Beijing calendar day (quarters included).

    Every final costs one observed period-score request; a game whose
    periods do not reconcile is dropped with a gap, never folded or
    repaired.
    """
    _require_cba(league)
    _reset_gaps()
    finals: List[Game] = []
    for row in day_rows(_iso_date(date_yyyymmdd), fetch):
        if row.get("status") != 4:
            continue
        try:
            game = _game_from_row(row)
        except IdentityError as exc:
            _gap(str(row.get("scheduleId") or "") or None, "schedule_row", str(exc))
            continue
        try:
            finals.append(attach_result(game, row, fetch, post_fn))
        except IdentityError as exc:
            _gap(game.game_id, "actual_result", str(exc))
    return finals


def fetch_quarters(
    schedule_id: str, post_fn: Optional[PostFn] = None
) -> Dict[str, Any]:
    """Decrypted period-score payload for one game.

    The observed contract is a POST with ``{"scheduleId": n}``. The
    repo's ``FetchFn`` is GET-only, so the POST runs through ``post_fn``
    (default: a bounded ``requests.post`` with the public referer header,
    no credentials, timeout, single attempt like the other adapters' raw
    paths). Tests inject a fake ``post_fn`` returning recorded wire text.
    """
    text = _post(
        f"{BASE_URL}/schedule-quarters/getQuarterInfo",
        {"scheduleId": int(schedule_id)},
        post_fn,
    )
    info = _decrypt(text)
    if not isinstance(info, dict):
        raise IdentityError(f"CBA {schedule_id} quarter info is not an object")
    if info.get("style") is None:
        raise IdentityError(f"CBA {schedule_id} quarter info lacks the style marker")
    return info


def _post(url: str, body: Dict[str, Any], post_fn: Optional[PostFn] = None) -> str:
    if post_fn is not None:
        return post_fn(url, body)
    import json

    import requests

    response = requests.post(
        url,
        data=json.dumps(body),
        headers={**_headers(), "Content-Type": "application/json"},
        timeout=20.0,
    )
    if response.status_code != 200:
        raise IdentityError(f"CBA quarter-info POST returned HTTP {response.status_code}")
    return response.text


def attach_result(
    game: Game, row: Dict[str, Any], fetch: FetchFn, post_fn: Optional[PostFn] = None
) -> Game:
    """Attach a reconciled ``ActualResult`` from the day row + period scores.

    Period scores exist only through the observed quarter-info call. A
    game whose periods do not reconcile (or whose payload fails to decrypt)
    raises ``IdentityError``: the caller quarantines it — never a
    final-score-only result, never a repaired sum.
    """
    schedule_id = str(game.provider_ids.get(PROVIDER) or game.game_id)
    info = fetch_quarters(schedule_id, post_fn)
    final_home = _period(row, "homeTeamScore", schedule_id)
    final_away = _period(row, "visitingTeamScore", schedule_id)
    if final_home is None or final_away is None:
        raise IdentityError(f"CBA {schedule_id} final row lacks both scores")
    game.actual = _actual(schedule_id, info, final_home, final_away)
    game.status = GameStatus.FINAL
    return game


def result_from_row(
    row: Dict[str, Any], fetch: FetchFn, post_fn: Optional[PostFn] = None
) -> Game:
    """Build the game from a day row and attach its reconciled result."""
    game = _game_from_row(row)
    return attach_result(game, row, fetch, post_fn)


def schedule_smoke(
    date_yyyymmdd: str, fetch: FetchFn, league: League | str = League.CBA
) -> Dict[str, Any]:
    _require_cba(league)
    games = discover_games(date_yyyymmdd, fetch, league)
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": "CBA",
        "verification": EVIDENCE,
        "season": current_season(fetch),
        "fixtures_found": len(games),
        "gaps": [
            {"game_id": gap.game_id, "required_field": gap.required_field, "detail": gap.detail}
            for gap in _LAST_GAPS
        ],
    }


__all__ = [
    "BASE_URL",
    "DOCUMENTATION_ONLY",
    "EVIDENCE",
    "LEAGUE_TZ",
    "PROVIDER",
    "STATUS_MAP",
    "attach_result",
    "available_seasons",
    "current_season",
    "day_rows",
    "discover_games",
    "fetch_quarters",
    "fetch_recent_results",
    "is_enabled",
    "last_gaps",
    "last_quota",
    "result_from_row",
    "schedule_smoke",
]
