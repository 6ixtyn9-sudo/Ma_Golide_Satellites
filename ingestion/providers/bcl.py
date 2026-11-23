"""Basketball Champions League adapter on the FIBA GDAP platform.

BCL runs on the same FIBA ``digital-api.fiba.basketball/hapi`` service the
EuroLeague Women lane already reads, so this module **reuses** that contract
rather than restating it: the APIM header name, the public fallback key, the
UTC-timestamp parser, the conservative status vocabulary, the unassigned-slot
test and the score test all come from
:mod:`ingestion.providers.fiba_euroleague_w_contract`.

What is BCL-specific
--------------------
* Competition id **209123** (EuroLeague Women is 209127), on
  ``getgdapgamesbycompetitionid``. One call returns the whole season — 136 rows
  for 2026-27 in the captured body — so the date filter happens client-side on
  ``gameDateTimeUTC``.
* Its own identity namespace (``bcl:{teamId}``). A men's club must never share an
  external team identity with the women's competition, even where the same
  organisation appears in both.
* No period split: the schedule index carries ``teamAScore``/``teamBScore`` and
  nothing per quarter, and the boxscore route has not been captured. A settled
  row therefore stores a totals-only :class:`ActualResult`
  (``PeriodKind.UNKNOWN``) and is not quarter-reconcilable. Quarter markets must
  wait for a captured detail body rather than infer splits from a total.

Evidence
--------
``tests/fixtures/bcl/captures/gdap-games-209123-row.json`` is one **verbatim row**
from the captured response (HTTP 200, 23,789 bytes, 2026-11-23T12:11:54Z). The
sibling rows were not vendored; tests derive other states by mutating that row,
which is labelled as such.

Fail-closed rules
-----------------
An unassigned bracket slot (``teamA``/``teamB`` null) is a truthful schedule row
but has no canonical identity, so it is skipped and recorded. A row whose status
the shared vocabulary cannot classify is skipped, never guessed into scheduled or
final. A row whose competition object disagrees with the requested id is refused.
"""

from __future__ import annotations

import hashlib
import json
from typing import Any, Dict, List, Optional

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError
from .fiba_euroleague_w_contract import (
    GDAP_BASE_URL,
    KEY_HEADER,
    PUBLIC_FALLBACK_SUBSCRIPTION_KEY,
    game_datetime_utc,
    has_complete_team_identity,
    has_unassigned_team_slot,
    schedule_status,
    scores_present,
)

PROVIDER = "bcl"

COMPETITION_ID = 209123
COMPETITION_CODE = "BCL"
SCHEDULE_ENDPOINT = "getgdapgamesbycompetitionid"
SCHEDULE_URL = f"{GDAP_BASE_URL}/{SCHEDULE_ENDPOINT}"

BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"

_STATUS_MAP = {
    "scheduled": GameStatus.SCHEDULED,
    "live": GameStatus.LIVE,
    "final": GameStatus.FINAL,
    "postponed": GameStatus.POSTPONED,
    "cancelled": GameStatus.CANCELLED,
    "abandoned": GameStatus.ABANDONED,
}

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


def schedule_url(competition_id: int = COMPETITION_ID) -> str:
    return f"{SCHEDULE_URL}?gdapCompetitionId={int(competition_id)}"


def request_headers() -> Dict[str, str]:
    return {
        KEY_HEADER: PUBLIC_FALLBACK_SUBSCRIPTION_KEY,
        "Accept": "application/json",
    }


def default_fetch() -> FetchFn:
    client = HttpClient(timeout=30.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_json(url, headers=request_headers())


def _rows(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, list):
        rows = payload
    elif isinstance(payload, dict) and isinstance(payload.get("value"), list):
        rows = payload["value"]
    else:
        raise ValueError("BCL GDAP schedule is neither a list nor a value envelope")
    return [row for row in rows if isinstance(row, dict)]


def _team(side: Any, game_id: str, label: str) -> tuple[str, str, str]:
    if not isinstance(side, dict):
        raise IdentityError(f"BCL game {game_id} has no {label} object")
    team_id = side.get("teamId")
    if team_id is None or not str(team_id).isdigit():
        raise IdentityError(f"BCL game {game_id} {label} lacks an integer teamId")
    code = str(side.get("code") or "").strip()
    name = canonical_team_name(str(side.get("officialName") or side.get("shortName") or ""))
    if not name:
        raise IdentityError(f"BCL game {game_id} {label} lacks a name")
    return str(team_id), name, code


def game_from_row(row: Dict[str, Any], competition_id: int = COMPETITION_ID) -> Game:
    """Map one strictly-scoped GDAP row to a canonical BCL game."""

    game_id = row.get("gameId")
    if game_id is None or not str(game_id).isdigit():
        raise IdentityError("BCL schedule row lacks an integer gameId")
    game_id = str(game_id)

    competition = row.get("competition") or {}
    if int(competition.get("competitionId") or competition_id) != int(competition_id):
        raise IdentityError(
            f"BCL game {game_id} belongs to competition "
            f"{competition.get('competitionId')!r}, expected {competition_id}"
        )

    home_id, home_name, home_code = _team(row.get("teamA"), game_id, "teamA")
    away_id, away_name, away_code = _team(row.get("teamB"), game_id, "teamB")
    if home_id == away_id or (home_code and home_code == away_code):
        raise IdentityError(f"BCL game {game_id} has indistinguishable teamA/teamB identity")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"BCL game {game_id} resolves the same name on both sides")

    vocabulary = schedule_status(row)
    status = _STATUS_MAP.get(vocabulary)
    if status is None:
        raise ValueError(
            f"BCL game {game_id} status is {vocabulary!r}; refusing to guess a state"
        )

    actual: Optional[ActualResult] = None
    if status is GameStatus.FINAL:
        if not scores_present(row):
            raise ValueError(f"BCL game {game_id} is final without both scores")
        # Totals only: the schedule index carries no quarter split and the
        # boxscore route is not captured, so this record is deliberately not
        # quarter-reconcilable.
        actual = ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=float(row["teamAScore"]),
            away_final=float(row["teamBScore"]),
            source=f"bcl getgdapgamesbycompetitionid gameId={game_id}",
        )

    scheduled_utc = to_utc_iso(game_datetime_utc(row.get("gameDateTimeUTC")))
    round_row = row.get("round") or {}
    return Game(
        game_id=f"BCL-{game_id}",
        league=League.BASKETBALL_CHAMPIONS_LEAGUE,
        season=str(competition.get("season") or ""),
        home_team_id=f"bcl:{home_id}",
        away_team_id=f"bcl:{away_id}",
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=scheduled_utc,
        status=status,
        phase=str(round_row.get("roundName") or COMPETITION_CODE),
        provider_ids={
            PROVIDER: game_id,
            "bcl_competition": str(competition_id),
            "bcl_team_a": home_id,
            "bcl_team_b": away_id,
            "bcl_team_a_code": home_code,
            "bcl_team_b_code": away_code,
            "bcl_timezone": str(row.get("ianaTimeZone") or ""),
        },
        actual=actual,
        evidence={
            "provider": PROVIDER,
            "venue": row.get("venueName"),
            "round": round_row.get("roundName"),
            "status_text": row.get("statusCode"),
            "url": schedule_url(competition_id),
        },
    )


def discover_games(
    date_iso: Optional[str],
    fetch: Optional[FetchFn] = None,
    competition_id: int = COMPETITION_ID,
) -> List[Game]:
    """Games of one UTC date, filtered from the season index before mapping."""

    _reset()
    transport = fetch or default_fetch()
    url = schedule_url(competition_id)
    payload = transport(url)
    rows = _rows(payload)
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    _last_probe.update(
        {
            "url": url,
            "bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
            "rows": len(rows),
        }
    )

    games: List[Game] = []
    for row in rows:
        game_id = str(row.get("gameId") or "") or None
        if has_unassigned_team_slot(row):
            # Future bracket placeholders are truthful but have no identity.
            _record_gap(game_id, "unassigned bracket slot; no canonical fixture")
            continue
        if not has_complete_team_identity(row):
            _record_gap(game_id, "row lacks two assigned teams")
            continue
        try:
            game = game_from_row(row, competition_id)
        except (IdentityError, ValueError) as exc:
            _record_gap(game_id, str(exc))
            continue
        if date_iso and not game.scheduled_utc.startswith(date_iso):
            continue
        games.append(game)
    _last_probe["games"] = len(games)
    return games


def _find_row(fetch: FetchFn, game_id: str, competition_id: int) -> Dict[str, Any]:
    for row in _rows(fetch(schedule_url(competition_id))):
        if str(row.get("gameId")) == game_id:
            return row
    raise ValueError(f"BCL game {game_id} is not in the competition index")


def fetch_game_result(
    game: Game, fetch: Optional[FetchFn] = None, competition_id: int = COMPETITION_ID
) -> Game:
    """Re-read one game's index row so a settled result is not lost/guessed."""

    game_id = str(game.provider_ids.get(PROVIDER) or "")
    if not game_id:
        raise IdentityError(f"game {game.game_id} carries no BCL gameId")
    transport = fetch or default_fetch()
    refreshed = game_from_row(_find_row(transport, game_id, competition_id), competition_id)
    if refreshed.status is not GameStatus.FINAL:
        raise ValueError(f"BCL game {game_id} is not final; refusing a result")
    if refreshed.home_team_id != game.home_team_id:
        raise IdentityError(
            f"BCL game {game_id} row names home {refreshed.home_team_id}, "
            f"expected {game.home_team_id}"
        )
    return refreshed
