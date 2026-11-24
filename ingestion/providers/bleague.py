"""Japan B.League (B1) schedule/result adapter for ``www.bleague.jp`` + S3.

Contract reference
------------------
The slate is **server-rendered HTML**, not an inline JSON blob, and the club
identity, tip-off instant and period split live on a keyless S3 snapshot. Both
halves were captured by an out-of-sandbox probe on 2026-11-23 (this sandbox has
no egress to either host) and are vendored under
``tests/fixtures/bleague/captures/`` with the excerpts labelled there.

* ``GET /schedule/?tab={1|2|3}&year={START_YEAR}&mon={M}&day={D}`` renders one
  day of one division. Each game is ``<li class="list-item" id="{ScheduleKey}">``
  inside ``<ul id="schedule-b1">``; the block carries
  ``span.team.home > span.team-name`` / ``span.team.away > span.team-name``,
  ``span.number.home-score`` / ``span.number.away-score``, the local tip time as
  the third ``span`` of ``div.info-arena`` (JST, no offset in the markup) and the
  status text in ``div.info-scorestate`` (``FINAL``, ``見どころ`` before tip).
  **Note the parameter names**: ``month`` is ignored by the host; ``mon`` is the
  honoured spelling, and ``year`` is the season's START year.
* ``GET /direct/bleague/get_today_games_score/?ScheduleKey={k1,k2,...}`` answers
  one JSON object keyed by schedule key with ``HomeTeamScore``,
  ``AwayTeamScore``, ``GameEndedFlg`` and ``RecordFixedFlg``.
* ``GET https://b-league.s3.amazonaws.com/web_json/v2_genius_contexts/{key}/latestid``
  answers the current snapshot id (a bare integer), and
  ``.../{key}/{snapshot}.json`` answers the game context
  ``{PlayByPlays, HomeBoxscores, AwayBoxscores, Summaries, Game}``. ``Game``
  carries ``HomeTeamID``/``AwayTeamID``, the long and short names, the totals,
  per-quarter ``HomeTeamScore01..04`` and ``GameDateTime`` in .NET JSON form
  ``/Date({epoch_ms}{+0900})/``.

The documented ``web_json/contexts/{key}`` tree refuses anonymous reads (403
``AccessDenied``); ``v2_genius_contexts`` is the working keyless path, and it is
the one this adapter reads. No key, no login, no special header.

What this adapter refuses
-------------------------
``GameEndedFlg`` is the only settled marker trusted, and only with both totals.
A quarter split is written only when Q1..Q4 are all present **and** they sum to
the total on both sides; an overtime or a missing period therefore leaves a
totals-only record plus a gap instead of a split that cannot reconcile. A body
answering another ``ScheduleKey``, or naming other clubs, is refused rather than
attached to the wrong game.

Every scope, identity and timestamp problem is recorded as a gap; a slate page
with no recognisable game block raises instead of reading as an empty slate.
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import date, datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from ..http import HttpClient
from ..identity import canonical_team_name, to_utc_iso
from ..models import ActualResult, Game, GameStatus, League, PeriodKind
from . import FetchFn, IdentityError

PROVIDER = "bleague"

SITE = "https://www.bleague.jp"
SCHEDULE_URL = SITE + "/schedule/"
SCORE_URL = SITE + "/direct/bleague/get_today_games_score/"
CONTEXT_BASE = "https://b-league.s3.amazonaws.com/web_json/v2_genius_contexts"

#: ``tab`` is the division, top-first: 1 = B1, 2 = B2, 3 = B3. The numbering
#: survived the 2026-27 rebrand -- tab=1 still serves the top division -- but the
#: *names* did not: tab=1 now renders "B.LEAGUE PREMIER" and tab=2 "B.LEAGUE
#: ONE" (captured 2026-11-24, probes24/aliases/bleague-schedule*-20261127.html).
#:
#: One adapter, two lanes, one division each. BetExplorer's ``japan/b-league-one``
#: slug is the **second** division and matched the tab=2 slate game-for-game, so
#: it belongs to ``JAPAN_B_ONE``; the top division is its own lane and claims no
#: board slug until a page of it is captured. Sharing one lane across both
#: divisions is what priced PREMIER fixtures against B.ONE clubs and made every
#: name disagree (run 45385819819). See ``docs/RESULTS-ROUTE-ASSESSMENT.md``.
TAB_B1 = 1
TAB_B2 = 2

#: The division each Japan lane reads. A lane is exactly one competition; this
#: mapping is the only place the two are allowed to be related.
TAB_BY_LEAGUE: Dict[League, int] = {
    League.JAPAN_B_LEAGUE: TAB_B1,  # B.LEAGUE PREMIER (top division)
    League.JAPAN_B_ONE: TAB_B2,  # B.LEAGUE ONE (second division)
}

#: Japan is UTC+09:00 all year (no DST), so the local offset is a constant.
JST = timezone(timedelta(hours=9))

BROWSER_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) MaGolide/1.0"
REQUIRED_HEADERS = {"Accept": "text/html,application/json"}

_last_gaps: List[Dict[str, Any]] = []
_last_probe: Dict[str, Any] = {}

_LI_BLOCK = re.compile(r'<li[^>]*class="[^"]*list-item[^"]*"[^>]*id="(\d+)"[^>]*>(.*?)</li>', re.S)
_HOME_NAME = re.compile(r'<span class="team home">.*?<span class="team-name">(.*?)</span>', re.S)
_AWAY_NAME = re.compile(r'<span class="team away">.*?<span class="team-name">(.*?)</span>', re.S)
_HOME_SCORE = re.compile(r'<span class="number home-score[^"]*">\s*<span>(\d*)</span>', re.S)
_AWAY_SCORE = re.compile(r'<span class="number away-score[^"]*">\s*<span>(\d*)</span>', re.S)
_STATUS_TEXT = re.compile(r'<div class="info-scorestate">\s*<span>(.*?)</span>', re.S)
_TIP_TEXT = re.compile(r"<span>\s*(\d{1,2}:\d{2})\s*</span>")
_DOTNET_DATE = re.compile(r"^/Date\((-?\d+)([+-]\d{4})?\)/$")


def _reset() -> None:
    _last_gaps.clear()
    _last_probe.clear()


def last_gaps() -> List[Dict[str, Any]]:
    return [dict(gap) for gap in _last_gaps]


def last_probe() -> Dict[str, Any]:
    """What the run actually read: URLs, sizes and digests, never a body."""

    return dict(_last_probe)


def _record_gap(game_id: Optional[str], detail: str) -> None:
    _last_gaps.append({"game_id": game_id, "detail": detail})


def _record_read(url: str, body: str) -> None:
    reads = _last_probe.setdefault("reads", [])
    reads.append(
        {
            "url": url,
            "bytes": len(body.encode("utf-8")),
            "sha256": hashlib.sha256(body.encode("utf-8")).hexdigest(),
        }
    )
    _last_probe["bytes"] = sum(item["bytes"] for item in reads)
    _last_probe["url"] = reads[0]["url"]
    _last_probe["sha256"] = reads[0]["sha256"]
    _last_probe["requests"] = len(reads)


def default_fetch() -> FetchFn:
    """Text transport: the slate and the snapshots are read as text, then parsed.

    ``get_json`` is deliberately not used: the slate is HTML, and the S3
    snapshot needs the same transport as the ``latestid`` pointer (a bare
    integer, not JSON).
    """

    client = HttpClient(timeout=25.0, retries=2, user_agent=BROWSER_USER_AGENT)
    return lambda url: client.get_text(url, headers=REQUIRED_HEADERS)


def season_end_year(day: date) -> int:
    """The season label's END year: September starts a new season."""

    return day.year + 1 if day.month >= 8 else day.year


def _plain(text: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", text)).strip()


def schedule_url(day: date, tab: int = TAB_B1) -> str:
    return f"{SCHEDULE_URL}?tab={int(tab)}&year={day.year}&mon={day.month}&day={day.day}"


def tab_for(league: Any) -> int:
    """The division tab for one Japan lane; an unknown lane is refused."""

    value = league.value if isinstance(league, League) else str(league).strip().upper()
    try:
        return TAB_BY_LEAGUE[League(value)]
    except (KeyError, ValueError) as exc:
        raise ValueError(f"no B.League division tab for {league!r}") from exc


def parse_slate(html: str) -> List[Dict[str, Any]]:
    """The day's game blocks: key, both names, scores, status and local tip."""

    blocks: List[Dict[str, Any]] = []
    for match in _LI_BLOCK.finditer(html):
        key, body = match.group(1), match.group(2)
        home = _HOME_NAME.search(body)
        away = _AWAY_NAME.search(body)
        if not home or not away:
            continue
        tip = _TIP_TEXT.search(body)
        status = _STATUS_TEXT.search(body)
        blocks.append(
            {
                "schedule_key": key,
                "home_name": _plain(home.group(1)),
                "away_name": _plain(away.group(1)),
                "home_score": (_HOME_SCORE.search(body).group(1) or None)
                if _HOME_SCORE.search(body)
                else None,
                "away_score": (_AWAY_SCORE.search(body).group(1) or None)
                if _AWAY_SCORE.search(body)
                else None,
                "status_text": _plain(status.group(1)) if status else "",
                "tip_local": tip.group(1) if tip else "",
            }
        )
    return blocks


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


def _tip_utc(raw: Any, key: str) -> str:
    text = str(raw or "").strip()
    match = _DOTNET_DATE.match(text)
    if not match:
        raise ValueError(f"B.LEAGUE game {key} GameDateTime {text!r} is not .NET JSON form")
    millis = int(match.group(1))
    return to_utc_iso(datetime.fromtimestamp(millis / 1000.0, tz=timezone.utc))


def status_from_text(text: str) -> str:
    upper = text.upper()
    if "FINAL" in upper or "終了" in text:
        return "final"
    if "LIVE" in upper or "Q" in upper or "第" in text:
        return "live"
    return "scheduled"


def _period_result(game: Dict[str, Any], key: str) -> Optional[ActualResult]:
    """A quarter split, only when it reconciles; otherwise ``None``."""

    quarters_home = [_int(game.get(f"HomeTeamScore0{index}")) for index in (1, 2, 3, 4)]
    quarters_away = [_int(game.get(f"AwayTeamScore0{index}")) for index in (1, 2, 3, 4)]
    if any(value is None for value in quarters_home + quarters_away):
        return None
    home_total = _int(game.get("HomeTeamScore"))
    away_total = _int(game.get("AwayTeamScore"))
    if home_total is None or away_total is None:
        return None
    if sum(quarters_home) != home_total or sum(quarters_away) != away_total:
        # An overtime (MaxPeriod > 4) or a corrected period: the four regulation
        # values do not add up, so no split is asserted.
        return None
    return ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[float(value) for value in quarters_home],
        away_quarters=[float(value) for value in quarters_away],
        home_final=float(home_total),
        away_final=float(away_total),
        source=f"bleague v2_genius_contexts ScheduleKey={key}",
    )


def _game_from_context(
    payload: Dict[str, Any],
    key: str,
    requested_date: Optional[str],
    league: League = League.JAPAN_B_LEAGUE,
) -> Optional[Game]:
    game = payload.get("Game") if isinstance(payload.get("Game"), dict) else payload
    if not isinstance(game, dict):
        raise ValueError(f"B.LEAGUE context {key} carries no Game object")
    if _int(game.get("ScheduleKey")) != _int(key):
        raise IdentityError(
            f"B.LEAGUE context answers ScheduleKey {game.get('ScheduleKey')!r}, asked for {key!r}"
        )
    home_id = str(game.get("HomeTeamID") or "").strip()
    away_id = str(game.get("AwayTeamID") or "").strip()
    if not home_id or not away_id or home_id == away_id:
        raise IdentityError(f"B.LEAGUE game {key} lacks distinct club ids")
    home_name = canonical_team_name(
        str(game.get("HomeTeamNameE") or game.get("HomeTeamNameJ") or "")
    )
    away_name = canonical_team_name(
        str(game.get("AwayTeamNameE") or game.get("AwayTeamNameJ") or "")
    )
    if not home_name or not away_name:
        raise IdentityError(f"B.LEAGUE game {key} lacks a club name")
    if home_name.casefold() == away_name.casefold():
        raise IdentityError(f"B.LEAGUE game {key} resolves the same name on both sides")

    scheduled_utc = _tip_utc(game.get("GameDateTime"), key)
    if requested_date and not scheduled_utc.startswith(requested_date):
        return None

    ended = bool(game.get("GameEndedFlg"))
    period = _int(game.get("GameCurrentPeriod")) or 0
    home_score = _int(game.get("HomeTeamScore"))
    away_score = _int(game.get("AwayTeamScore"))
    actual: Optional[ActualResult] = None
    if ended and home_score is not None and away_score is not None:
        status = GameStatus.FINAL
        actual = _period_result(game, key)
        if actual is None:
            actual = ActualResult(
                period_kind=PeriodKind.UNKNOWN,
                home_quarters=[],
                away_quarters=[],
                home_final=float(home_score),
                away_final=float(away_score),
                source=f"bleague v2_genius_contexts ScheduleKey={key}",
            )
            _record_gap(
                key,
                "quarter split does not reconcile with the totals (overtime or a "
                "missing period); totals only",
            )
    elif ended:
        status = GameStatus.SCHEDULED
        _record_gap(key, "GameEndedFlg set without both totals; refused as a final")
    elif period >= 1:
        status = GameStatus.LIVE
    else:
        status = GameStatus.SCHEDULED

    day = date.fromisoformat(scheduled_utc[:10])
    return Game(
        game_id=f"BLEAGUE-{key}",
        league=league,
        season=str(season_end_year(day)),
        home_team_id=home_id,
        away_team_id=away_id,
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=scheduled_utc,
        status=status,
        provider_ids={PROVIDER: key},
        actual=actual,
        evidence={
            "provider": PROVIDER,
            "competition": game.get("ConventionNameE"),
            "venue": game.get("StadiumNameE"),
            "max_period": game.get("MaxPeriod"),
            "current_period": game.get("GameCurrentPeriod"),
            "record_fixed": game.get("RecordFixedFlg"),
            "snapshot": CONTEXT_BASE + f"/{key}/latestid",
        },
    )


def _context(payload: str, key: str) -> Dict[str, Any]:
    try:
        parsed = json.loads(payload)
    except ValueError as exc:
        raise ValueError(f"B.LEAGUE context {key} is not JSON: {exc}") from exc
    if not isinstance(parsed, dict):
        raise ValueError(f"B.LEAGUE context {key} is not a JSON object")
    return parsed


def _fetch_context(key: str, transport: FetchFn) -> Dict[str, Any]:
    latest_url = f"{CONTEXT_BASE}/{key}/latestid"
    latest = str(transport(latest_url) or "").strip()
    _record_read(latest_url, latest)
    if not latest.isdigit():
        raise ValueError(f"B.LEAGUE {key}/latestid did not answer a snapshot id")
    snapshot_url = f"{CONTEXT_BASE}/{key}/{latest}.json"
    body = transport(snapshot_url)
    if not isinstance(body, str):
        body = json.dumps(body, sort_keys=True)
    _record_read(snapshot_url, body)
    return _context(body, key)


def discover_games(
    date_iso: Optional[str],
    fetch: Optional[FetchFn] = None,
    tab: Optional[int] = None,
    league: League = League.JAPAN_B_LEAGUE,
) -> List[Game]:
    """Games of one UTC date in one division, from the slate page plus contexts.

    ``tab`` defaults to the division ``league`` owns; passing a tab for another
    division is refused, because a lane that reads one competition and prices
    another cannot be diagnosed.

    The page is asked for in **JST terms** for both JST dates that can contain a
    tip inside the target UTC day, and rows whose tip falls outside
    ``date_iso`` are dropped afterwards.

    ``day`` does **not** scope the page's content *near today*. On 2026-11-23
    four fetches (``day=22/23/24/25``) returned 143,266 B each, parsed to the
    same ten game keys (506602-506611), and the day22/day23 bodies differed by
    exactly two bytes: the echoed ``day`` in a meta tag and in a JS assignment.
    The host serves whatever round its pointer is on, and every row carries its
    own ``GameDateTime``, which is what makes the UTC-day filter above correct.

    That observation is bounded by what was asked, and the bound matters: a
    separate probe on 2026-11-24 of a deep-past day (``year=2025&mon=5&day=1``)
    returned thirteen DIFFERENT ScheduleKeys with zero overlap against today's
    ten, and one of those keys resolved to an S3 snapshot carrying both teams'
    Q1-Q4. So the honest statement is "near-today days do not scope, a
    historical day does" -- a historical day is the only way to enumerate past
    games, and ``day_param`` below can only ever report the near-today world
    because this lane asks for today and tomorrow by design.

    The original text claimed the wider thing outright, which was wrong; see
    ``_record_read``'s ``day_param`` for how the narrower claim is measured.

    Both dates are still requested: the second costs one request, and it is the
    only thing that would find a tip on the neighbouring JST day. What the
    second page's contribution was is recorded per page (``new_keys``), and
    ``day_param`` says which world we are in -- ``ignored`` when the second page
    added nothing, ``scoped`` when it did.
    """

    _reset()
    # Entry markers. A lane that returns [] having recorded nothing is an
    # impossible state that nonetheless showed up on the runner: the guard in
    # Pipeline._discover_games reported "returned no fixtures and recorded
    # neither a coverage gap nor a read receipt" while this function is written
    # to raise rather than return empty. Recording how far execution got turns
    # that contradiction into evidence instead of another theory.
    expected_tab = tab_for(league)
    if tab is not None and int(tab) != expected_tab:
        raise ValueError(
            f"{league.value} reads tab={expected_tab}, not tab={int(tab)}; "
            "one lane is one division"
        )
    tab = expected_tab
    _last_probe["entered"] = True
    _last_probe["league"] = league.value
    _last_probe["tab"] = tab
    transport = fetch or default_fetch()
    _last_probe["transport"] = "injected" if fetch is not None else "default"
    if date_iso:
        target = datetime.strptime(date_iso, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    else:
        target = datetime.now(timezone.utc)
    local_day = target.astimezone(JST).date()

    keys: List[str] = []
    for offset in (0, 1):
        day = local_day + timedelta(days=offset)
        url = schedule_url(day, tab)
        _last_probe["stages"] = [*_last_probe.get("stages", []), f"request:{day.isoformat()}"]
        html = transport(url)
        if not isinstance(html, str):
            raise ValueError("B.LEAGUE schedule page was not text")
        _record_read(url, html)
        blocks = parse_slate(html)
        _last_probe["stages"] = [*_last_probe.get("stages", []), f"blocks:{len(blocks)}"]
        if not blocks:
            raise ValueError(
                f"B.LEAGUE schedule page for {day.isoformat()} has no recognisable game "
                "block; refusing to read it as an empty slate"
            )
        page = {
            "day": day.isoformat(),
            "sha256": hashlib.sha256(html.encode("utf-8")).hexdigest(),
            "bytes": len(html.encode("utf-8")),
            "blocks": len(blocks),
            "new_keys": 0,
        }
        for block in blocks:
            if block["schedule_key"] not in keys:
                keys.append(block["schedule_key"])
                page["new_keys"] += 1
        _last_probe.setdefault("pages", []).append(page)

    games: List[Game] = []
    off_date = 0
    for key in keys:
        try:
            payload = _fetch_context(key, transport)
            game = _game_from_context(payload, key, date_iso, league)
        except (IdentityError, ValueError) as exc:
            _record_gap(key, str(exc))
            continue
        if game is None:
            # A slate row whose tip is not inside the requested UTC day. That is
            # a correct drop, but it is also the reason an 11-23 local slate can
            # answer 11-24 with zero fixtures; without counting it, a read that
            # found nothing is indistinguishable from a read that never happened.
            off_date += 1
            continue
        games.append(game)
    pages = _last_probe.get("pages", [])
    _last_probe["day_param"] = (
        "scoped" if any(page["new_keys"] for page in pages[1:]) else "ignored"
    )
    _last_probe["slate_games"] = len(keys)
    _last_probe["off_date_dropped"] = off_date
    _last_probe["games"] = len(games)
    return games


def fetch_game_result(game: Game, fetch: Optional[FetchFn] = None) -> Game:
    """Re-read one game's context so its period split can reconcile."""

    key = str(game.provider_ids.get(PROVIDER) or "")
    if not key:
        raise IdentityError(f"game {game.game_id} carries no B.LEAGUE ScheduleKey")
    transport = fetch or default_fetch()
    refreshed = _game_from_context(_fetch_context(key, transport), key, None, game.league)
    if refreshed is None:
        raise ValueError(f"B.LEAGUE game {key} is not on a dated slate")
    if refreshed.home_team_id != game.home_team_id:
        raise IdentityError(
            f"B.LEAGUE game {key} context names home club {refreshed.home_team_id}, "
            f"expected {game.home_team_id}"
        )
    if refreshed.away_team_id != game.away_team_id:
        raise IdentityError(
            f"B.LEAGUE game {key} context names away club {refreshed.away_team_id}, "
            f"expected {game.away_team_id}"
        )
    if refreshed.status is not GameStatus.FINAL:
        raise ValueError(f"B.LEAGUE game {key} is not ended; refusing a result")
    return refreshed
