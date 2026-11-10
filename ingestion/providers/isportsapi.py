"""iSportsAPI Chinese Basketball Association acquisition.

Evidence status: DOCUMENTATION_ONLY until an authenticated smoke test runs.
The adapter follows the published ``/sport/basketball`` contracts for league
 discovery, schedule/results, and pre-match fulltime/half/quarter odds. CBA is
resolved through the provider's league directory by an exact, unique identity;
no league ID is hardcoded or guessed.

Only documented pre-match ``instant*`` fields are used (with opening quarter
fields as a fallback when the current field group is empty). In-play fields
are never relabelled as pre-match. Prices arrive as Hong Kong odds and are
converted to decimal by adding one. Book observations are never converted
into Prediction records; independent CBA probabilities come from Sportradar.
"""

from __future__ import annotations

from datetime import datetime, timezone
import math
import os
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence
from urllib.parse import urlencode

from ..identity import canonical_team_name, now_utc_iso, parse_utc
from ..models import (
    ActualResult,
    BookSnapshot,
    CoverageGap,
    CoverageState,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    payload_hash,
)
from . import FetchFn, IdentityError

PROVIDER = "isportsapi"
ENV_KEY = "ISPORTS_API_KEY"
BASE_URL = "https://api.isportsapi.com"
DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"

_CBA_FULL_NAME = "chinesebasketballassociation"
_CBA_SHORT_NAME = "cba"
_FINAL = -1
_STATUS = {
    0: GameStatus.SCHEDULED,
    -1: GameStatus.FINAL,
    -2: GameStatus.UNKNOWN,
    -3: GameStatus.ABANDONED,
    -4: GameStatus.CANCELLED,
    -5: GameStatus.POSTPONED,
}
_PHASE = {1: "League", 2: "Playoffs", 3: "Preseason", -1: "League"}
_SPREAD_BOOKS = {
    "1": "Macauslot",
    "2": "Easybets",
    "3": "Crown",
    "8": "Bet365",
    "9": "Vcbet",
    "10": "William Hill",
    "19": "Interwetten",
    "20": "Ladbrokes",
    "31": "Sbobet",
    "24": "12bet",
    "30": "China Sports Lottery",
    "49": "BWin",
}
_TOTAL_BOOKS = {
    "4": "Macauslot",
    "5": "Easybets",
    "6": "Crown",
    "11": "Bet365",
    "12": "Vcbet",
    "13": "William Hill",
    "22": "Interwetten",
    "23": "Ladbrokes",
    "34": "Sbobet",
    "27": "12bet",
    "33": "China Sports Lottery",
    "52": "BWin",
}
_LAST_GAPS: List[CoverageGap] = []


def is_enabled() -> bool:
    return bool(os.environ.get(ENV_KEY, "").strip())


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


def _base_url() -> str:
    return os.environ.get("ISPORTS_API_BASE_URL", BASE_URL).rstrip("/")


def _url(path: str, **params: Any) -> str:
    query = {"api_key": os.environ.get(ENV_KEY, ""), **params}
    return f"{_base_url()}/sport/basketball/{path}?{urlencode(query)}"


def _payload(fetch: FetchFn, path: str, **params: Any) -> Dict[str, Any]:
    payload = fetch(_url(path, **params)) or {}
    if not isinstance(payload, dict):
        raise IdentityError(f"{PROVIDER} {path}: expected an object response")
    code = payload.get("code", 0)
    if code not in (0, "0", None):
        raise IdentityError(
            f"{PROVIDER} {path}: provider code {code!r}: {str(payload.get('message') or '')[:160]}"
        )
    data = payload.get("data")
    if data is None:
        raise IdentityError(f"{PROVIDER} {path}: response has no data field")
    return payload


def _normalized(value: Any) -> str:
    return "".join(character for character in str(value or "").lower() if character.isalnum())


def resolve_cba_league_id(fetch: FetchFn) -> str:
    """Resolve exactly one Chinese Basketball Association directory entry."""
    rows = _payload(fetch, "league/basic").get("data") or []
    if not isinstance(rows, list):
        raise IdentityError("league/basic data is not a list")
    full = [
        row
        for row in rows
        if isinstance(row, dict) and _normalized(row.get("leagueName")) == _CBA_FULL_NAME
    ]
    candidates = full or [
        row
        for row in rows
        if isinstance(row, dict) and _normalized(row.get("leagueShortName")) == _CBA_SHORT_NAME
    ]
    identities = {str(row.get("leagueId") or "").strip() for row in candidates}
    identities.discard("")
    if len(identities) != 1:
        raise IdentityError(
            f"Chinese CBA league resolution returned {len(identities)} unique IDs; refusing to guess"
        )
    return identities.pop()


def _iso_date(date_yyyymmdd: str) -> str:
    text = str(date_yyyymmdd).replace("-", "")
    datetime.strptime(text, "%Y%m%d")
    return f"{text[:4]}-{text[4:6]}-{text[6:8]}"


def _number(value: Any, field: str, *, required: bool = True) -> Optional[float]:
    if value is None or value == "":
        if required:
            raise IdentityError(f"missing numeric {field}")
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise IdentityError(f"unparseable numeric {field}: {value!r}") from None
    if not math.isfinite(number) or number < 0:
        raise IdentityError(f"invalid numeric {field}: {value!r}")
    return number


def _timestamp(value: Any) -> str:
    seconds = _number(value, "matchTime")
    try:
        return datetime.fromtimestamp(seconds, tz=timezone.utc).isoformat().replace("+00:00", "Z")
    except (OverflowError, OSError, ValueError):
        raise IdentityError(f"invalid matchTime: {value!r}") from None


def _status(value: Any) -> GameStatus:
    try:
        code = int(value)
    except (TypeError, ValueError):
        return GameStatus.UNKNOWN
    return _STATUS.get(code, GameStatus.LIVE if code > 0 else GameStatus.UNKNOWN)


def _actual(row: Dict[str, Any]) -> ActualResult:
    if int(row.get("quarterCount") or 0) != 4:
        raise IdentityError("CBA result does not declare a four-quarter format")
    names = ("First", "Second", "Third", "Fourth")
    home = [
        _number(row.get(f"home{name}QuarterScore"), f"home{name}QuarterScore") for name in names
    ]
    away = [
        _number(row.get(f"away{name}QuarterScore"), f"away{name}QuarterScore") for name in names
    ]
    try:
        overtime_periods = int(row.get("overTimeCount") or 0)
    except (TypeError, ValueError):
        raise IdentityError("invalid overTimeCount") from None
    if overtime_periods < 0 or overtime_periods > 3:
        raise IdentityError(f"unsupported CBA overtime count {overtime_periods}")
    overtime_names = ("First", "Second", "Third")
    home_ot = away_ot = None
    if overtime_periods:
        home_ot = sum(
            _number(row.get(f"home{overtime_names[index]}OverTimeScore"), "home overtime") or 0.0
            for index in range(overtime_periods)
        )
        away_ot = sum(
            _number(row.get(f"away{overtime_names[index]}OverTimeScore"), "away overtime") or 0.0
            for index in range(overtime_periods)
        )
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home,
        away_quarters=away,
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=_number(row.get("homeScore"), "homeScore"),
        away_final=_number(row.get("awayScore"), "awayScore"),
        overtime_periods=overtime_periods,
        source=PROVIDER,
    )


def _game(row: Dict[str, Any]) -> Game:
    match_id = str(row.get("matchId") or "").strip()
    home_id = str(row.get("homeId") or "").strip()
    away_id = str(row.get("awayId") or "").strip()
    home_name = str(row.get("homeName") or "").strip()
    away_name = str(row.get("awayName") or "").strip()
    if not all((match_id, home_id, away_id, home_name, away_name)):
        raise IdentityError("CBA schedule row lacks match/team identity")
    status = _status(row.get("status"))
    actual = _actual(row) if status is GameStatus.FINAL else None
    try:
        match_type = int(row.get("matchType", -1))
    except (TypeError, ValueError):
        match_type = -1
    return Game(
        game_id=match_id,
        league=League.CBA,
        season=str(row.get("leagueSeason") or ""),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=canonical_team_name(home_name),
        away_name=canonical_team_name(away_name),
        scheduled_utc=_timestamp(row.get("matchTime")),
        status=status,
        phase=_PHASE.get(match_type, "League"),
        provider_ids={PROVIDER: match_id},
        actual=actual,
    )


def _schedule(date_yyyymmdd: str, fetch: FetchFn, league_id: str) -> List[Game]:
    rows = _payload(fetch, "schedule", date=_iso_date(date_yyyymmdd)).get("data") or []
    if not isinstance(rows, list):
        raise IdentityError("schedule data is not a list")
    games: List[Game] = []
    for row in rows:
        if not isinstance(row, dict) or str(row.get("leagueId") or "") != str(league_id):
            continue
        try:
            games.append(_game(row))
        except IdentityError as exc:
            _gap(str(row.get("matchId") or "") or None, "schedule_or_result", str(exc))
    return games


def discover_games(
    date_yyyymmdd: str, fetch: FetchFn, league: League | str = League.CBA
) -> List[Game]:
    _require_cba(league)
    _reset_gaps()
    return _schedule(date_yyyymmdd, fetch, resolve_cba_league_id(fetch))


def fetch_recent_results(
    date_yyyymmdd: str, fetch: FetchFn, league: League | str = League.CBA
) -> List[Game]:
    return [
        game
        for game in discover_games(date_yyyymmdd, fetch, league)
        if game.status is GameStatus.FINAL and game.actual is not None
    ]


def fetch_game_result(
    match_id: str,
    fetch: FetchFn,
    game: Optional[Game] = None,
    league: League | str = League.CBA,
) -> Game:
    _require_cba(league)
    rows = _payload(fetch, "schedule", matchId=str(match_id)).get("data") or []
    matches = [
        row
        for row in rows
        if isinstance(row, dict) and str(row.get("matchId") or "") == str(match_id)
    ]
    if len(matches) != 1:
        raise IdentityError(f"CBA match {match_id}: expected one schedule row, got {len(matches)}")
    result = _game(matches[0])
    if result.status is not GameStatus.FINAL or result.actual is None:
        raise IdentityError(f"CBA match {match_id} is not a complete final")
    if game is not None:
        if game.league is not League.CBA or str(
            game.provider_ids.get(PROVIDER, game.game_id)
        ) != str(match_id):
            raise IdentityError("CBA result identity contradicts the supplied game")
        if (result.home_team_id, result.away_team_id) != (game.home_team_id, game.away_team_id):
            raise IdentityError("CBA result home/away identity contradicts the scheduled game")
        game.status = GameStatus.FINAL
        game.actual = result.actual
        return game
    return result


def _rows(payload: Dict[str, Any], key: str) -> Iterable[List[str]]:
    data = payload.get("data") or {}
    for raw in data.get(key) or []:
        if not isinstance(raw, str):
            continue
        yield raw.split(",")


def _hk_decimal(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        price = float(value) + 1.0
    except (TypeError, ValueError):
        return None
    return price if math.isfinite(price) and price >= 1.0 else None


def _line(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        line = float(value)
    except (TypeError, ValueError):
        return None
    return line if math.isfinite(line) else None


def _book(mapping: Dict[str, str], company_id: str) -> Optional[str]:
    name = mapping.get(str(company_id))
    if name is None:
        _gap(None, "bookmaker", f"unrecognised companyId {company_id!r}; row skipped")
    return name


def _pregame(game: Game, now_utc: str) -> bool:
    return parse_utc(game.scheduled_utc) > parse_utc(now_utc)


def _snapshot(
    game: Game,
    *,
    book: str,
    market: Market,
    period: Period,
    observed_utc: str,
    line: Optional[float],
    first_price: Optional[float],
    second_price: Optional[float],
    raw: Sequence[str],
) -> Optional[BookSnapshot]:
    if first_price is None or second_price is None:
        return None
    if market in {Market.SPREAD, Market.TOTAL} and line is None:
        return None
    return BookSnapshot(
        game_id=game.game_id,
        provider=PROVIDER,
        provider_event_id=game.game_id,
        book=book,
        market=market,
        period=period,
        observed_utc=observed_utc,
        line=line,
        home_price_decimal=first_price if market in {Market.MONEYLINE, Market.SPREAD} else None,
        away_price_decimal=second_price if market in {Market.MONEYLINE, Market.SPREAD} else None,
        over_price_decimal=first_price if market is Market.TOTAL else None,
        under_price_decimal=second_price if market is Market.TOTAL else None,
        pre_match=_pregame(game, observed_utc),
        main_line=True,
        payload_hash=payload_hash(list(raw)),
        home_name=game.home_name,
        away_name=game.away_name,
        scheduled_utc=game.scheduled_utc,
    )


def _parse_fulltime(
    payload: Dict[str, Any],
    games: Dict[str, Game],
    observed_utc: str,
    moneyline_books: Dict[str, str],
) -> List[BookSnapshot]:
    output: List[BookSnapshot] = []
    for raw in _rows(payload, "moneyLine"):
        if len(raw) != 6 or raw[0] not in games:
            continue
        book = _book(moneyline_books, raw[1])
        if book:
            snapshot = _snapshot(
                games[raw[0]],
                book=book,
                market=Market.MONEYLINE,
                period=Period.FULL_GAME,
                observed_utc=observed_utc,
                line=None,
                first_price=_hk_decimal(raw[4]),
                second_price=_hk_decimal(raw[5]),
                raw=raw,
            )
            if snapshot:
                output.append(snapshot)
    for key, market, mapping in (
        ("spread", Market.SPREAD, _SPREAD_BOOKS),
        ("total", Market.TOTAL, _TOTAL_BOOKS),
    ):
        for raw in _rows(payload, key):
            if len(raw) != 11 or raw[0] not in games:
                continue
            book = _book(mapping, raw[1])
            if book:
                snapshot = _snapshot(
                    games[raw[0]],
                    book=book,
                    market=market,
                    period=Period.FULL_GAME,
                    observed_utc=observed_utc,
                    line=_line(raw[5]),
                    first_price=_hk_decimal(raw[6]),
                    second_price=_hk_decimal(raw[7]),
                    raw=raw,
                )
                if snapshot:
                    output.append(snapshot)
    return output


def _parse_half(
    payload: Dict[str, Any], games: Dict[str, Game], observed_utc: str
) -> List[BookSnapshot]:
    output: List[BookSnapshot] = []
    for key, market, mapping in (
        ("spread", Market.SPREAD, _SPREAD_BOOKS),
        ("total", Market.TOTAL, _TOTAL_BOOKS),
    ):
        for raw in _rows(payload, key):
            if len(raw) != 8 or raw[0] not in games:
                continue
            book = _book(mapping, raw[1])
            if book:
                snapshot = _snapshot(
                    games[raw[0]],
                    book=book,
                    market=market,
                    period=Period.H1,
                    observed_utc=observed_utc,
                    line=_line(raw[5]),
                    first_price=_hk_decimal(raw[6]),
                    second_price=_hk_decimal(raw[7]),
                    raw=raw,
                )
                if snapshot:
                    output.append(snapshot)
    return output


def _parse_quarters(
    payload: Dict[str, Any], games: Dict[str, Game], observed_utc: str
) -> List[BookSnapshot]:
    output: List[BookSnapshot] = []
    periods = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)
    for key, market, mapping in (
        ("spread", Market.SPREAD, _SPREAD_BOOKS),
        ("total", Market.TOTAL, _TOTAL_BOOKS),
    ):
        for raw in _rows(payload, key):
            if len(raw) != 26 or raw[0] not in games:
                if raw and raw[0] in games:
                    _gap(raw[0], f"{key}_quarters", f"expected 26 fields, got {len(raw)}")
                continue
            book = _book(mapping, raw[1])
            if not book:
                continue
            for index, period in enumerate(periods):
                current = 2 + index * 3
                opening = 14 + index * 3
                group = current if any(raw[current : current + 3]) else opening
                snapshot = _snapshot(
                    games[raw[0]],
                    book=book,
                    market=market,
                    period=period,
                    observed_utc=observed_utc,
                    line=_line(raw[group]),
                    first_price=_hk_decimal(raw[group + 1]),
                    second_price=_hk_decimal(raw[group + 2]),
                    raw=raw,
                )
                if snapshot:
                    output.append(snapshot)
    return output


def _games_by_id(date_yyyymmdd: str, fetch: FetchFn) -> Dict[str, Game]:
    return {game.game_id: game for game in discover_games(date_yyyymmdd, fetch)}


def _moneyline_bookmakers(fetch: FetchFn) -> Dict[str, str]:
    """Resolve companyIdEu names without mixing spread/total ID namespaces."""
    rows = _payload(fetch, "bookmaker").get("data") or []
    if not isinstance(rows, list):
        raise IdentityError("bookmaker data is not a list")
    books: Dict[str, str] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        company_id = row.get("companyIdEu")
        name = str(row.get("companyName") or "").strip()
        if company_id in (None, "", 0, "0") or not name:
            continue
        key = str(company_id)
        if key in books and books[key] != name:
            raise IdentityError(f"conflicting names for companyIdEu {key}")
        books[key] = name
    return books


def fetch_odds(
    league: str,
    date_yyyymmdd: str,
    fetch_factory: Callable[[str], FetchFn],
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> List[BookSnapshot]:
    _require_cba(league)
    if not is_enabled():
        return []
    fetch = fetch_factory(PROVIDER)
    observed = now_utc or now_utc_iso()
    games = _games_by_id(date_yyyymmdd, fetch)
    try:
        moneyline_books = _moneyline_bookmakers(fetch)
    except Exception as exc:
        moneyline_books = {}
        _gap(None, "moneyline_bookmakers", str(exc))
    snapshots = _parse_fulltime(_payload(fetch, "odds/fulltime"), games, observed, moneyline_books)
    snapshots.extend(_parse_half(_payload(fetch, "odds/half"), games, observed))
    snapshots.extend(_parse_quarters(_payload(fetch, "odds/quarter"), games, observed))
    return snapshots


def schedule_smoke(
    date_yyyymmdd: str, fetch: FetchFn, league: League | str = League.CBA
) -> Dict[str, Any]:
    _require_cba(league)
    if not is_enabled():
        return {
            "enabled": False,
            "provider": PROVIDER,
            "league": "CBA",
            "verification": "UNAUTHORIZED",
            "reason": f"missing {ENV_KEY}",
        }
    games = discover_games(date_yyyymmdd, fetch, league)
    finals = [game for game in games if game.status is GameStatus.FINAL and game.actual]
    reconciled = [game for game in finals if game.actual and game.actual.reconciles()[0]]
    return {
        "enabled": is_enabled(),
        "provider": PROVIDER,
        "league": "CBA",
        "verification": DOCUMENTATION_ONLY,
        "fixtures_found": len(games),
        "results_reconciled": {
            "finals": len(finals),
            "reconciled": len(reconciled),
            "flagged": len(finals) - len(reconciled),
        },
        "gaps": [
            {"game_id": gap.game_id, "required_field": gap.required_field, "detail": gap.detail}
            for gap in _LAST_GAPS
        ],
    }


def smoke(
    league: str,
    fetch_factory: Callable[[str], FetchFn],
    date_yyyymmdd: Optional[str] = None,
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> Dict[str, Any]:
    _require_cba(league)
    if not is_enabled():
        return {
            "enabled": False,
            "provider": PROVIDER,
            "league": "CBA",
            "verification": "UNAUTHORIZED",
        }
    date = date_yyyymmdd or datetime.now(timezone.utc).strftime("%Y%m%d")
    snapshots = fetch_odds(league, date, fetch_factory, now_utc=now_utc, tz_name=tz_name)
    by_period: Dict[str, set] = {}
    for snapshot in snapshots:
        if snapshot.pre_match and (
            snapshot.market is not Market.TOTAL or snapshot.line is not None
        ):
            by_period.setdefault(snapshot.period.value, set()).add(snapshot.game_id)
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": "CBA",
        "date": date,
        "verification": DOCUMENTATION_ONLY,
        "games_with_prematch_evidence": len({s.game_id for s in snapshots if s.pre_match}),
        "native_periods": {
            period: len(by_period.get(period, set()))
            for period in ("full_game", "H1", "Q1", "Q2", "Q3", "Q4")
        },
        "gaps": len(_LAST_GAPS),
    }


__all__ = [
    "DOCUMENTATION_ONLY",
    "ENV_KEY",
    "PROVIDER",
    "discover_games",
    "fetch_game_result",
    "fetch_odds",
    "fetch_recent_results",
    "is_enabled",
    "last_gaps",
    "last_quota",
    "resolve_cba_league_id",
    "schedule_smoke",
    "smoke",
]
