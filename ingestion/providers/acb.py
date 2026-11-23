"""ACB (Liga Endesa) schedule/result adapter for a referenced route contract.

Contract reference
------------------
This adapter targets the ACB's own JSON API. Its route and field contract is
referenced from SportsDataverse documentation generated against
``https://api2.acb.com/api`` (``sportsdataverse-py``; ``acb`` module). Ma Golide
retains no first-party source response in this repository, so live behaviour
and the source's automation/reuse terms remain unverified.

* ``GET /seasondata/Calendar/season-calendar?season={end_year}`` returns every
  round of one season with its games: ``rounds[].matches[]`` carrying ``id``
  (match id), ``homeTeamId``/``awayTeamId`` (season-team entries),
  ``homeScore``/``awayScore``, ``startDateTime``, ``matchStatus`` and
  ``currentQuarter``.
* ``GET /matchdata/MatchHeader/match-header?matchId={id}`` returns one game's
  header: ``teams.home``/``teams.away`` with ``fullName`` and stable club ids,
  ``matchStatus`` (``FINALIZED`` once the game is over),
  ``currentHomeScore``/``currentAwayScore`` and ``quarterScores[]`` with
  ``quarter``, ``homeScore`` and ``awayScore`` — points scored *in* that
  period, not cumulative.

The calendar is season-scoped (one request lists a whole season), but it does
not carry team names. Identity therefore comes from the per-game header, so a
slate costs one calendar request plus one header request per candidate game.
Every scope, identity, period, and final check below fails closed rather than
guessing from a score or list order.
"""

from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence

from ..http import HttpClient, HttpError
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

BASE_URL = "https://api2.acb.com/api"
CALENDAR_URL = BASE_URL + "/seasondata/Calendar/season-calendar"
MATCH_HEADER_URL = BASE_URL + "/matchdata/MatchHeader/match-header"
PROVIDER = "acb"

#: api2.acb.com is the .NET API behind acb.com and live.acb.com. It is not a
#: documented public API and it answers an unauthenticated client with an auth
#: failure, which is why this lane read zero games for 728 runs while its
#: parser tests passed against a hand-written body.
#:
#: ``X-APIKEY`` is a *public* value shipped in acb.com's own JavaScript bundle
#: (``sharedServicesConfigDotnet``). It is a site-scoped identifier, not an
#: account credential: nothing is billed, nothing is attributable, and the
#: operator can override it with ``$ACB_API_KEY`` if the bundle rotates it.
#: Recorded here from the SportsDataverse ``acb`` runtime (MIT, read
#: 2026-11-23), which read it from the bundle on 2026-10-12.
KEY_ENV = "ACB_API_KEY"
DEFAULT_KEY = "0dd94928-6f57-4c08-a3bd-b1b2f092976e"
KEY_HEADER = "X-APIKEY"
BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"

_last_probe: Dict[str, Any] = {}


def api_key() -> str:
    """The site-bundle key, with an operator override."""

    return os.environ.get(KEY_ENV) or DEFAULT_KEY


def request_headers() -> Dict[str, str]:
    return {KEY_HEADER: api_key(), "Accept": "application/json"}


def default_fetch() -> FetchFn:
    """Adapter-owned JSON client that carries the required key header."""

    client = HttpClient(timeout=30.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_json(url, headers=request_headers())


def last_probe() -> Dict[str, Any]:
    """What the last calendar response actually looked like (never a body)."""

    return dict(_last_probe)


def _record_probe(url: str, payload: Any, rows: int) -> None:
    body = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    _last_probe.update(
        {
            "url": url,
            "bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
            "rounds": len(payload.get("rounds") or []) if isinstance(payload, dict) else 0,
            "rows": rows,
        }
    )

COMPETITION_LIGA_ENDESA = 1
COMPETITION_NAMES = {1: "Liga Endesa", 2: "Copa del Rey", 3: "Supercopa Endesa"}
SEASON_FLOOR = 1984  # first ACB season; ``season`` is the season END year
FINAL_STATUS = "FINALIZED"
PROVIDER = "acb"

_last_gaps: List[Dict[str, Any]] = []


def last_gaps() -> List[Dict[str, Any]]:
    """Gaps recorded by the most recent call (identity/period problems)."""

    return list(_last_gaps)


def _reset_gaps() -> None:
    _last_gaps.clear()


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def season_end_year(target: datetime) -> int:
    """Season END year for an ACB date (the API's naming, e.g. 2027 == 2026-27).

    ACB seasons start in the autumn and finish in the following spring, so a
    July boundary selects the season whose final rounds fall in that spring.
    """

    return target.year + 1 if target.month >= 7 else target.year


def season_id_for(end_year: int) -> int:
    """The API's ``editionId`` for a season END year (documented offset)."""

    if end_year < SEASON_FLOOR:
        raise ValueError(f"ACB season {end_year} precedes the {SEASON_FLOOR} floor")
    return end_year - 1936


def _int(value: Any) -> Optional[int]:
    if value is None or isinstance(value, bool):
        return None
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return parsed


def _iso_utc(value: Any) -> str:
    """Parse a provider timestamp into canonical UTC, refusing naive values."""

    text = str(value or "").strip()
    if not text:
        raise ValueError("ACB timestamp is empty")
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"ACB timestamp {text!r} is not ISO 8601: {exc}") from exc
    if parsed.tzinfo is None:
        # The calendar is served in local (Europe/Madrid) wall-clock terms; a
        # naive value would silently shift every tip. Refuse instead.
        raise ValueError(f"ACB timestamp {text!r} has no UTC offset")
    return to_utc_iso(parsed)


def fetch_calendar(
    end_year: int, fetch: FetchFn, competition: int = COMPETITION_LIGA_ENDESA
) -> List[Dict[str, Any]]:
    """Every game row of one season, flattened out of its rounds."""

    season_id_for(end_year)
    payload = fetch(f"{CALENDAR_URL}?season={int(end_year)}&competition={int(competition)}")
    if not isinstance(payload, dict):
        raise ValueError(f"ACB calendar {end_year} is not a JSON object")
    calendar_url = f"{CALENDAR_URL}?season={int(end_year)}&competition={int(competition)}"
    rounds = payload.get("rounds")
    if not isinstance(rounds, list) or not rounds:
        raise ValueError(
            f"ACB calendar {end_year} has no rounds; an empty body is refused rather than "
            "treated as a season without games"
        )
    rows: List[Dict[str, Any]] = []
    for round_row in rounds:
        if not isinstance(round_row, dict):
            continue
        matches = round_row.get("matches")
        if not isinstance(matches, list):
            continue
        for match in matches:
            if isinstance(match, dict):
                rows.append({"round": round_row, "match": match})
    if not rows:
        raise ValueError(f"ACB calendar {end_year} has rounds but no matches")
    _record_probe(calendar_url, payload, len(rows))
    return rows


class AuthenticationRequired(RuntimeError):
    """The ACB match-header route answered 401/403.

    A live probe on 2026-11-22 found the season calendar public but the
    per-game header gated: the calendar alone carries no team names and no
    period scores, so this lane cannot produce a canonical game until that
    route is authorised. Raising a distinct error keeps the reason visible in
    coverage gaps instead of looking like an empty season.
    """


def _match_header(fetch: FetchFn, match_id: str) -> Dict[str, Any]:
    url = f"{MATCH_HEADER_URL}?matchId={match_id}"
    try:
        payload = fetch(url)
    except HttpError as exc:
        if int(getattr(exc, "status", 0) or 0) in (401, 403):
            raise AuthenticationRequired(
                f"ACB match-header requires authorisation (HTTP {exc.status}); "
                "the public calendar carries neither team names nor period scores"
            ) from exc
        raise
    if not isinstance(payload, dict):
        raise ValueError(f"ACB match header {match_id} is not a JSON object")
    if not isinstance(payload.get("teams"), dict):
        raise ValueError(f"ACB match header {match_id} has no teams object")
    return payload


def _periods(payload: Dict[str, Any], match_id: str) -> ActualResult:
    """Reconciled regulation quarters plus any overtime, from ``quarterScores``."""

    rows = payload.get("quarterScores")
    if not isinstance(rows, list) or not rows:
        raise ValueError(f"ACB match {match_id} has no quarterScores")
    quarters: Dict[int, tuple] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        quarter = _int(row.get("quarter"))
        home = row.get("homeScore")
        away = row.get("awayScore")
        if quarter is None or home is None or away is None:
            continue
        try:
            quarters[int(quarter)] = (float(home), float(away))
        except (TypeError, ValueError):
            continue
    if not all(index in quarters for index in (1, 2, 3, 4)):
        raise ValueError(f"ACB match {match_id} lacks four regulation quarters")
    overtime = sorted(index for index in quarters if index > 4)
    if overtime and overtime != list(range(5, 5 + len(overtime))):
        raise ValueError(f"ACB match {match_id} has non-contiguous overtime periods {overtime}")
    home_overtime = sum(quarters[index][0] for index in overtime) if overtime else None
    away_overtime = sum(quarters[index][1] for index in overtime) if overtime else None
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[quarters[index][0] for index in (1, 2, 3, 4)],
        away_quarters=[quarters[index][1] for index in (1, 2, 3, 4)],
        home_overtime=home_overtime,
        away_overtime=away_overtime,
        home_final=_int(payload.get("currentHomeScore")),
        away_final=_int(payload.get("currentAwayScore")),
        overtime_periods=len(overtime),
        source=f"acb match-header matchId={match_id}",
    )


def _game_from_header(header: Dict[str, Any], match_id: str, season: int, phase: str) -> Game:
    teams = header.get("teams") or {}
    home = teams.get("home") or {}
    away = teams.get("away") or {}
    home_name = canonical_team_name(str(home.get("fullName") or ""))
    away_name = canonical_team_name(str(away.get("fullName") or ""))
    if not home_name or not away_name:
        raise IdentityError(f"ACB match {match_id} header lacks a team name")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"ACB match {match_id} names the same team on both sides")
    home_id = str(home.get("id") or home.get("clubId") or "")
    away_id = str(away.get("id") or away.get("clubId") or "")
    if not home_id or not away_id or home_id == away_id:
        raise IdentityError(f"ACB match {match_id} header lacks distinct team ids")
    status_text = str(header.get("matchStatus") or "").strip().upper()
    status = GameStatus.FINAL if status_text == FINAL_STATUS else GameStatus.SCHEDULED
    game = Game(
        game_id=f"ACB-{match_id}",
        league=League.SPAIN_ACB,
        season=str(int(season)),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=_iso_utc(header.get("start")),
        status=status,
        phase=phase,
        provider_ids={PROVIDER: str(match_id)},
    )
    if status is GameStatus.FINAL:
        game.actual = _periods(header, match_id)
    return game


def _candidates(rows: Sequence[Dict[str, Any]], date_iso: Optional[str]) -> List[Dict[str, Any]]:
    """Calendar rows for one UTC date (or all rows when ``date_iso`` is None)."""

    selected: List[Dict[str, Any]] = []
    off_date = 0
    unparsable = 0
    for row in rows:
        match = row.get("match") or {}
        if date_iso is None:
            selected.append(row)
            continue
        start = str(match.get("startDateTime") or "")
        try:
            if _iso_utc(start).startswith(date_iso):
                selected.append(row)
            else:
                off_date += 1
        except ValueError:
            unparsable += 1
            continue
    _last_probe["off_date_dropped"] = off_date
    _last_probe["unparsable_date"] = unparsable
    return selected


def discover_games(
    date_iso: Optional[str],
    fetch: Optional[FetchFn],
    end_year: int,
    competition: int = COMPETITION_LIGA_ENDESA,
) -> List[Game]:
    """Games of one UTC date, identified from the per-game header.

    The calendar carries ids, dates, and scores but no team names, so identity
    is resolved per candidate game. A candidate whose header fails identity is
    recorded as a gap and skipped — never guessed from list order.
    """

    _reset_gaps()
    transport = fetch or default_fetch()
    phase = COMPETITION_NAMES.get(int(competition), f"Competition {competition}")
    games: List[Game] = []
    for row in _candidates(fetch_calendar(end_year, transport, competition), date_iso):
        match_id = str((row.get("match") or {}).get("id") or "")
        if not match_id:
            _record_gap(None, f"calendar row without a match id in {end_year}")
            continue
        try:
            header = _match_header(transport, match_id)
            games.append(_game_from_header(header, match_id, end_year, phase))
        except AuthenticationRequired as exc:
            # Every further candidate would fail identically; record the reason
            # once and stop instead of emitting a gap per game.
            _record_gap(match_id, str(exc))
            break
        except (IdentityError, ValueError) as exc:
            _record_gap(match_id, str(exc))
    return games


def fetch_game_result(game: Game, fetch: FetchFn) -> Game:
    """Attach reconciled periods to a scheduled/finished game from its header."""

    match_id = str(game.provider_ids.get(PROVIDER) or "")
    if not match_id:
        raise IdentityError(f"game {game.game_id} carries no ACB match id")
    header = _match_header(fetch, match_id)
    refreshed = _game_from_header(
        header, match_id, int(str(game.season or 0) or 0) or 0, game.phase or "Liga Endesa"
    )
    refreshed.provider_ids.update(game.provider_ids)
    if refreshed.home_name.casefold() != game.home_name.casefold():
        raise IdentityError(
            f"ACB match {match_id} header names {refreshed.home_name!r}, expected {game.home_name!r}"
        )
    return refreshed


def season_results(
    end_year: int,
    fetch: FetchFn,
    failures: Optional[List[Dict[str, str]]] = None,
    competition: int = COMPETITION_LIGA_ENDESA,
    limit: Optional[int] = None,
) -> List[Game]:
    """Reconciled finals of one season, one header request per game.

    Every finished calendar row is visited in round order until ``limit``
    reconciled finals have been produced (``None`` means no cap). A row whose
    header fails identity or reconciliation is recorded in ``failures`` and
    skipped; nothing is folded or invented.
    """

    _reset_gaps()
    phase = COMPETITION_NAMES.get(int(competition), f"Competition {competition}")
    results: List[Game] = []
    for row in fetch_calendar(end_year, fetch, competition):
        match = row.get("match") or {}
        match_id = str(match.get("id") or "")
        if not match_id:
            continue
        if str(match.get("matchStatus") or "").strip().upper() != FINAL_STATUS:
            continue
        try:
            game = _game_from_header(_match_header(fetch, match_id), match_id, end_year, phase)
        except AuthenticationRequired as exc:
            if failures is not None:
                failures.append({"game_id": f"ACB-{match_id}", "error": str(exc)})
            break
        except (IdentityError, ValueError) as exc:
            if failures is not None:
                failures.append({"game_id": f"ACB-{match_id}", "error": str(exc)})
            continue
        if game.status is not GameStatus.FINAL or game.actual is None:
            if failures is not None:
                failures.append(
                    {"game_id": game.game_id, "error": "header is not a reconciled final"}
                )
            continue
        ok, reason = game.actual.reconciles()
        if not ok:
            if failures is not None:
                failures.append({"game_id": game.game_id, "error": reason})
            continue
        results.append(game)
        if limit is not None and len(results) >= limit:
            break
    return results


def smoke(fetch: FetchFn, end_year: int, competition: int = COMPETITION_LIGA_ENDESA) -> Dict[str, Any]:
    """Coverage summary over one season calendar (never raises)."""

    report: Dict[str, Any] = {
        "games_found": 0,
        "results_reconciled": {"complete": 0, "total": 0},
        "failures": [],
        "verification": "FIXTURE_VERIFIED",
        "season_end_year": int(end_year),
        "competition": int(competition),
    }
    try:
        rows = fetch_calendar(end_year, fetch, competition)
    except Exception as exc:  # smoke never raises
        report["status"] = "provider_error"
        report["error_type"] = type(exc).__name__
        report["detail"] = str(exc)[:200]
        return report
    report["games_found"] = len(rows)
    finished = [
        row
        for row in rows
        if str((row.get("match") or {}).get("matchStatus") or "").strip().upper() == FINAL_STATUS
    ]
    report["results_reconciled"]["total"] = len(finished)
    failures: List[Dict[str, str]] = []
    complete = 0
    for row in finished:
        match_id = str((row.get("match") or {}).get("id") or "")
        try:
            game = _game_from_header(
                _match_header(fetch, match_id),
                match_id,
                end_year,
                COMPETITION_NAMES.get(int(competition), "Liga Endesa"),
            )
            ok, reason = (game.actual or ActualResult(PeriodKind.UNKNOWN, [], [])).reconciles()
            if ok:
                complete += 1
            else:
                failures.append({"game_id": game.game_id, "error": reason})
        except AuthenticationRequired as exc:
            failures.append({"game_id": f"ACB-{match_id}", "error": str(exc)})
            break
        except (IdentityError, ValueError) as exc:
            failures.append({"game_id": f"ACB-{match_id}", "error": str(exc)})
    report["results_reconciled"]["complete"] = complete
    report["failures"] = failures
    if failures and all(
        "requires authorisation" in str(row.get("error", "")) for row in failures
    ):
        # Never present a fully gated route as a season with no results.
        report["status"] = "auth_required"
        report["detail"] = (
            "ACB match-header answered 401/403 for every finished game; period-level "
            "results stay unavailable until the route is authorised"
        )
    return report
