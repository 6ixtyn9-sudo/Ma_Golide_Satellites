"""Basketball Africa League adapter for ``bal.nba.com/schedule``.

Contract reference
------------------
BAL is served by the NBA's site stack: ``https://bal.nba.com/schedule``
server-renders the **whole season** inside its Next.js ``__NEXT_DATA__`` blob, at
``props.pageProps.gameDates[]`` → ``{gameDate, seasonYear, games[]}``. A captured
body (HTTP 200, 379,284 bytes, 23 game dates, 2026-11-23T12:12:51Z, out-of-sandbox
probe) is vendored as one verbatim game object under
``tests/fixtures/bal/captures/``; the surrounding envelope there is synthesised
and labelled as such.

A game row carries ``gameId``, ``gameCode`` (``YYYYMMDD/AWAYHOME``),
``gameStatus``, ``gameStatusText``, ``gameDateUTC``/``gameDateTimeUTC``,
``arenaName``/``arenaCity``, ``postponedStatus``, ``ifNecessary``,
``seriesGameNumber`` and ``homeTeam``/``awayTeam`` with ``teamId``,
``teamName``, ``teamCity``, ``teamTricode``, ``teamSlug``, ``wins``, ``losses``
and ``score``. No key, no special header.

The FIBA GDAP route does **not** serve this competition: ``gdapCompetitionId=48286``
(the id BAL detail links carry) answers an empty array, so this lane reads the
NBA-hosted schedule instead. That dead end is recorded here so it is not retried.

What this adapter refuses
-------------------------
State is derived from **scores**, not from the unverified ``gameStatus`` code:
the captured row is ``gameStatus: 1`` with a 0-0 scoreline and a future tip, so a
single sample is not enough to assert what ``3`` means. A row is ``FINAL`` only
when both scores parse and at least one is non-zero — a settled basketball game
cannot be 0-0, so no result can be fabricated from an unplayed row.
``postponedStatus`` marks a postponement; a live ``gameStatus`` of 2 is treated
as live only when no final score is present. Everything else stays ``SCHEDULED``.

The schedule body carries no quarter split, so a settled row stores a totals-only
:class:`ActualResult` (``PeriodKind.UNKNOWN``) and is not quarter-reconcilable.
A body without ``gameDates`` raises rather than reading as an empty season.
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

PROVIDER = "bal"

SCHEDULE_URL = "https://bal.nba.com/schedule"

BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"
REQUIRED_HEADERS = {"Accept": "text/html,application/json"}

_NEXT_DATA = re.compile(
    r'<script[^>]*id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S
)

_last_gaps: List[Dict[str, Any]] = []
_last_probe: Dict[str, Any] = {}


def _reset() -> None:
    _last_gaps.clear()
    _last_probe.clear()


def last_gaps() -> List[Dict[str, Any]]:
    return [dict(gap) for gap in _last_gaps]


def last_probe() -> Dict[str, Any]:
    return dict(_last_probe)


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def default_fetch() -> FetchFn:
    """Text transport: the schedule is HTML, so the blob is extracted here."""

    client = HttpClient(timeout=25.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_text(url, headers=REQUIRED_HEADERS)


def extract_schedule(html: str) -> Dict[str, Any]:
    """The parsed ``__NEXT_DATA__`` payload, or a refusal."""

    if not isinstance(html, str):
        raise ValueError("BAL schedule page was not text")
    match = _NEXT_DATA.search(html)
    if not match:
        raise ValueError(
            "BAL schedule page carries no __NEXT_DATA__ blob; refusing to read it "
            "as an empty season"
        )
    try:
        payload = json.loads(match.group(1))
    except ValueError as exc:
        raise ValueError(f"BAL __NEXT_DATA__ is not JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("BAL __NEXT_DATA__ is not a JSON object")
    return payload


def _game_dates(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    page = ((payload.get("props") or {}).get("pageProps")) or {}
    dates = page.get("gameDates") if isinstance(page, dict) else None
    if not isinstance(dates, list):
        raise ValueError("BAL schedule payload has no props.pageProps.gameDates list")
    return [group for group in dates if isinstance(group, dict)]


def _int(value: Any) -> Optional[int]:
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


def _tip_utc(row: Dict[str, Any], game_id: str) -> str:
    raw = str(row.get("gameDateTimeUTC") or "").strip()
    if not raw:
        raise ValueError(f"BAL game {game_id} carries no gameDateTimeUTC")
    try:
        return to_utc_iso(datetime.fromisoformat(raw.replace("Z", "+00:00")))
    except ValueError as exc:
        raise ValueError(
            f"BAL game {game_id} gameDateTimeUTC {raw!r} is not ISO 8601: {exc}"
        ) from exc


def _team(side: Any, game_id: str, label: str) -> Tuple[str, str]:
    if not isinstance(side, dict):
        raise IdentityError(f"BAL game {game_id} lacks a {label} team object")
    team_id = _int(side.get("teamId"))
    if team_id is None:
        raise IdentityError(f"BAL game {game_id} {label} team lacks a teamId")
    name = canonical_team_name(
        " ".join(
            part
            for part in (
                str(side.get("teamCity") or "").strip(),
                str(side.get("teamName") or "").strip(),
            )
            if part
        )
    )
    if not name:
        raise IdentityError(f"BAL game {game_id} {label} team lacks a name")
    return str(team_id), name


def game_from_row(row: Dict[str, Any], season_year: Optional[int] = None) -> Game:
    """Map one strictly-scoped schedule row to a canonical BAL game."""

    game_id = str(row.get("gameId") or "").strip()
    if not game_id:
        raise IdentityError("BAL schedule row lacks a gameId")
    home_id, home_name = _team(row.get("homeTeam"), game_id, "home")
    away_id, away_name = _team(row.get("awayTeam"), game_id, "away")
    if home_id == away_id:
        raise IdentityError(f"BAL game {game_id} names the same team id on both sides")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"BAL game {game_id} resolves the same name on both sides")

    home_score = _int((row.get("homeTeam") or {}).get("score"))
    away_score = _int((row.get("awayTeam") or {}).get("score"))
    played = (
        home_score is not None
        and away_score is not None
        and (home_score > 0 or away_score > 0)
    )
    postponed = str(row.get("postponedStatus") or "N").strip().upper() == "P"
    actual: Optional[ActualResult] = None
    if played:
        status = GameStatus.FINAL
        # Totals only: the schedule body carries no quarter split.
        actual = ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=float(home_score),
            away_final=float(away_score),
            source=f"bal schedule gameId={game_id}",
        )
    elif postponed:
        status = GameStatus.POSTPONED
    elif _int(row.get("gameStatus")) == 2:
        status = GameStatus.LIVE
    else:
        status = GameStatus.SCHEDULED

    scheduled_utc = _tip_utc(row, game_id)
    year = season_year
    if year is None:
        code = str(row.get("gameCode") or "")
        year = _int(code[:4]) if code[:4].isdigit() else _int(scheduled_utc[:4])
    return Game(
        game_id=f"BAL-{game_id}",
        league=League.BASKETBALL_AFRICA_LEAGUE,
        season=str(year or ""),
        home_team_id=f"bal:{home_id}",
        away_team_id=f"bal:{away_id}",
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=scheduled_utc,
        status=status,
        provider_ids={
            PROVIDER: game_id,
            "bal_game_code": str(row.get("gameCode") or ""),
            "bal_team_a": home_id,
            "bal_team_b": away_id,
            "bal_tricode_a": str((row.get("homeTeam") or {}).get("teamTricode") or ""),
            "bal_tricode_b": str((row.get("awayTeam") or {}).get("teamTricode") or ""),
        },
        actual=actual,
        evidence={
            "provider": PROVIDER,
            "arena": row.get("arenaName"),
            "city": row.get("arenaCity"),
            "status_text": row.get("gameStatusText"),
            "neutral": row.get("isNeutral"),
            "series_game": row.get("seriesGameNumber"),
            "url": SCHEDULE_URL,
        },
    )


def discover_games(date_iso: Optional[str], fetch: Optional[FetchFn] = None) -> List[Game]:
    """Games of one UTC date, filtered from the season blob before mapping."""

    _reset()
    transport = fetch or default_fetch()
    html = transport(SCHEDULE_URL)
    payload = extract_schedule(html)
    groups = _game_dates(payload)
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    _last_probe.update(
        {
            "url": SCHEDULE_URL,
            "bytes": max(len(body), len(html.encode("utf-8"))),
            "sha256": hashlib.sha256(html.encode("utf-8")).hexdigest(),
            "game_dates": len(groups),
            "rows": sum(len(group.get("games") or []) for group in groups),
        }
    )

    games: List[Game] = []
    # Every row the date filter removes is counted under its own reason. The
    # old bare ``continue`` sat upstream of every counter, so run 45250804751
    # could report ``rows: 42`` beside ``games: 0`` and ``off_date_dropped: 0``
    # with nothing in between to explain the difference.
    off_date_dropped = 0
    missing_stamp_dropped = 0
    non_dict_rows = 0
    for group in groups:
        season_year = _int(group.get("seasonYear"))
        for row in group.get("games") or []:
            if not isinstance(row, dict):
                non_dict_rows += 1
                continue
            game_id = str(row.get("gameId") or "") or None
            stamp = row.get("gameDateTimeUTC")
            if date_iso:
                if not isinstance(stamp, str) or not stamp:
                    # A row with no parseable stamp cannot be placed on a day;
                    # dropping it is a decision, so it is counted as one.
                    missing_stamp_dropped += 1
                    continue
                if not stamp.startswith(date_iso):
                    off_date_dropped += 1
                    continue
            try:
                game = game_from_row(row, season_year)
            except (IdentityError, ValueError) as exc:
                _record_gap(game_id, str(exc))
                continue
            if date_iso and not game.scheduled_utc.startswith(date_iso):
                _record_gap(game_id, "tip does not fall on the requested UTC date")
                continue
            games.append(game)
    _last_probe.update(
        {
            "games": len(games),
            "off_date_dropped": off_date_dropped,
            "missing_stamp_dropped": missing_stamp_dropped,
            "non_dict_rows": non_dict_rows,
        }
    )
    return games


def fetch_game_result(game: Game, fetch: Optional[FetchFn] = None) -> Game:
    """Re-read the season blob for one game so a settled score is not guessed."""

    game_id = str(game.provider_ids.get(PROVIDER) or "")
    if not game_id:
        raise IdentityError(f"game {game.game_id} carries no BAL gameId")
    transport = fetch or default_fetch()
    payload = extract_schedule(transport(SCHEDULE_URL))
    for group in _game_dates(payload):
        for row in group.get("games") or []:
            if isinstance(row, dict) and str(row.get("gameId")) == game_id:
                refreshed = game_from_row(row, _int(group.get("seasonYear")))
                if refreshed.status is not GameStatus.FINAL:
                    raise ValueError(f"BAL game {game_id} is not settled; refusing a result")
                if refreshed.home_team_id != game.home_team_id:
                    raise IdentityError(
                        f"BAL game {game_id} names home {refreshed.home_team_id}, "
                        f"expected {game.home_team_id}"
                    )
                return refreshed
    raise ValueError(f"BAL game {game_id} is not in the schedule blob")
