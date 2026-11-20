"""ABA League schedule/result adapter for an externally referenced route contract.

Contract reference
------------------
This adapter targets the league's own server-rendered PHP pages. Its route and
markup contract is referenced from SportsDataverse documentation describing
third-party captures from ``www.aba-liga.com`` on 2026-10-12
(``sportsdataverse-py`` commit ``2fdd061``); Ma Golide retains no first-party
source response in this repository:

* ``/calendar/{sez}/1/`` is a complete ABA League season calendar. ``sez`` is
  the season start year minus 2000 (25 means 2025-26); game numbers restart
  every season.
* ``/match/{game_no}/{sez}/1/Overview`` carries stable team IDs, final score,
  Q1--Q4 and any OT periods for completed games.
* The external reference reports that the site labels wall-clock times ``CET``
  even during daylight saving. The parser treats those values as
  ``Europe/Belgrade`` wall-clock time, never as a fixed UTC offset; this remains
  subject to first-party capture confirmation.

The source's automation/reuse terms remain unresolved, as they are for several
existing public schedule adapters. That fact remains explicit in provenance and
does not create a price, ticket, ROI, or execution entitlement. The adapter's
injected-text seam is for hermetic tests; live use reads the public official
pages. Every scope, identity, period, and final reconciliation check below
fails closed rather than guessing from a score or table order.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from html.parser import HTMLParser
import re
from typing import Callable, Dict, Iterable, List, Optional, Sequence
from zoneinfo import ZoneInfo

from ..http import HttpClient
from ..identity import canonical_team_name, normalize_team_name, parse_utc, to_utc_iso
from ..models import ActualResult, CoverageGap, CoverageState, Game, GameStatus, League, PeriodKind
from . import IdentityError

PROVIDER = "aba_liga"
LEAGUE = League.ABA_LEAGUE
HOST = "https://www.aba-liga.com"
SITE_TZ = "Europe/Belgrade"
LEAGUE_ID = 1
FIRST_SEASON_END_YEAR = 2002
EVIDENCE = (
    "ROUTE_CONTRACT_REFERENCED (external SportsDataverse documentation describes "
    "2026-10-12 captures; no Ma Golide first-party capture retained; tests are "
    "synthetic; source terms unreviewed)"
)

TextFetch = Callable[[str], str]

_LAST_GAPS: List[CoverageGap] = []


@dataclass
class _Node:
    tag: str
    attrs: Dict[str, str]
    parent: Optional["_Node"] = None
    children: List[object] = field(default_factory=list)

    def classes(self) -> set[str]:
        return set(self.attrs.get("class", "").split())

    def text(self) -> str:
        chunks: List[str] = []
        stack = list(reversed(self.children))
        while stack:
            item = stack.pop()
            if isinstance(item, _Node):
                stack.extend(reversed(item.children))
            else:
                chunks.append(str(item))
        return re.sub(r"\s+", " ", "".join(chunks)).strip()

    def descendants(self) -> Iterable["_Node"]:
        stack = list(reversed(self.children))
        while stack:
            item = stack.pop()
            if isinstance(item, _Node):
                yield item
                stack.extend(reversed(item.children))

    def find_all(
        self, *, tag: Optional[str] = None, class_name: Optional[str] = None, ident: Optional[str] = None
    ) -> List["_Node"]:
        return [
            node
            for node in self.descendants()
            if (tag is None or node.tag == tag)
            and (class_name is None or class_name in node.classes())
            and (ident is None or node.attrs.get("id") == ident)
        ]

    def find(
        self, *, tag: Optional[str] = None, class_name: Optional[str] = None, ident: Optional[str] = None
    ) -> Optional["_Node"]:
        rows = self.find_all(tag=tag, class_name=class_name, ident=ident)
        return rows[0] if rows else None

    def direct(self, *tags: str) -> List["_Node"]:
        return [
            item
            for item in self.children
            if isinstance(item, _Node) and (not tags or item.tag in tags)
        ]


class _Tree(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = self.current = _Node("root", {})

    def handle_starttag(self, tag: str, attrs) -> None:  # type: ignore[no-untyped-def]
        node = _Node(tag, {str(key): str(value or "") for key, value in attrs}, self.current)
        self.current.children.append(node)
        if tag not in {"area", "base", "br", "col", "img", "input", "link", "meta", "source"}:
            self.current = node

    def handle_startendtag(self, tag: str, attrs) -> None:  # type: ignore[no-untyped-def]
        self.current.children.append(_Node(tag, {str(key): str(value or "") for key, value in attrs}, self.current))

    def handle_endtag(self, tag: str) -> None:
        node: Optional[_Node] = self.current
        while node is not None and node.tag != tag:
            node = node.parent
        if node is not None and node.parent is not None:
            self.current = node.parent

    def handle_data(self, data: str) -> None:
        self.current.children.append(data)


def _dom(html: str) -> _Node:
    if not isinstance(html, str) or not html.strip():
        raise IdentityError(f"{PROVIDER}: empty or non-text HTML response")
    parser = _Tree()
    parser.feed(html)
    parser.close()
    return parser.root


def is_enabled() -> bool:
    """The canonical schedule/result adapter is enabled for ABA League."""

    return True


def last_quota() -> Dict[str, object]:
    return {}


def last_gaps() -> List[CoverageGap]:
    return list(_LAST_GAPS)


def _reset_gaps() -> None:
    _LAST_GAPS.clear()


def _gap(game_id: Optional[str], field: str, detail: str, *, reason: CoverageState = CoverageState.MISSING) -> None:
    _LAST_GAPS.append(
        CoverageGap(
            game_id=game_id,
            required_field=field,
            reason=reason,
            provider=PROVIDER,
            detail=detail[:200],
        )
    )


def _text_fetch(text_fetch: Optional[TextFetch]) -> TextFetch:
    """Use injected fixture text when supplied, otherwise the official page."""

    return text_fetch or HttpClient().get_text


def season_end_year(day: date) -> int:
    """ABA season end year for a league-local date (season starts in autumn)."""
    return day.year + 1 if day.month >= 8 else day.year


def season_key(season: int) -> int:
    if isinstance(season, bool) or not isinstance(season, int):
        raise ValueError("ABA season end year must be an integer")
    if season < FIRST_SEASON_END_YEAR:
        raise ValueError(f"ABA historical floor is {FIRST_SEASON_END_YEAR}")
    return season - 2001


def calendar_url(season: int) -> str:
    return f"{HOST}/calendar/{season_key(season)}/{LEAGUE_ID}/"


def match_url(game_no: int, season: int) -> str:
    if isinstance(game_no, bool) or not isinstance(game_no, int) or game_no < 1:
        raise ValueError("ABA game number must be a positive integer")
    return f"{HOST}/match/{game_no}/{season_key(season)}/{LEAGUE_ID}/Overview"


def _hidden(root: _Node, ident: str) -> Optional[str]:
    node = root.find(tag="input", ident=ident)
    return node.attrs.get("value") if node else None


def _int(value: Optional[str], label: str) -> int:
    if value is None or not re.fullmatch(r"\d+", value.strip()):
        raise IdentityError(f"{PROVIDER}: {label} is missing or non-numeric")
    return int(value)


def _season_label(season: int) -> str:
    return f"{season - 1}-{season % 100:02d}"


def _game_ref(season: int, game_no: int) -> str:
    return f"{PROVIDER}:{season_key(season)}:{LEAGUE_ID}:{game_no}"


def _provider_event_id(season: int, game_no: int) -> str:
    return f"{season_key(season)}:{LEAGUE_ID}:{game_no}"


def _pair(anchor: _Node, label: str) -> tuple[str, str]:
    span = next((item for item in anchor.direct("span") if item.text() == ":"), None)
    if span is None:
        raise IdentityError(f"{PROVIDER}: {label} has no explicit home/away separator")
    index = anchor.children.index(span)

    def text(items: Sequence[object]) -> str:
        output: List[str] = []
        for item in items:
            output.append(item.text() if isinstance(item, _Node) else str(item))
        return re.sub(r"\s+", " ", "".join(output)).strip()

    home, away = text(anchor.children[:index]), text(anchor.children[index + 1 :])
    if not home or not away:
        raise IdentityError(f"{PROVIDER}: {label} has an unresolved team")
    return home, away


def _parse_when(value: str, label: str) -> str:
    text = re.sub(r"\s+", " ", value or "").strip()
    match = re.fullmatch(r"(?:[A-Za-z]+, )?(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2}) CET", text)
    if match is None:
        raise IdentityError(f"{PROVIDER}: {label} has no readable CET/CEST timestamp: {text!r}")
    try:
        local = datetime(
            int(match.group(3)),
            int(match.group(2)),
            int(match.group(1)),
            int(match.group(4)),
            int(match.group(5)),
            tzinfo=ZoneInfo(SITE_TZ),
        )
    except ValueError as exc:
        raise IdentityError(f"{PROVIDER}: {label} has invalid timestamp {text!r}") from exc
    return to_utc_iso(local)


def _is_unpinned_date_range(value: str) -> bool:
    """Recognize an upcoming calendar placeholder with no canonical tip time."""

    text = re.sub(r"\s+", " ", value or "").strip()
    return bool(
        re.fullmatch(
            r"(?:[A-Za-z]+, )?\d{2}\.\d{2}\.\d{4}\s*-\s*"
            r"(?:[A-Za-z]+, )?\d{2}\.\d{2}\.\d{4}",
            text,
        )
    )


def _scope_calendar(root: _Node, season: int) -> None:
    selected = root.find(tag="select", ident="season")
    if selected is None:
        raise IdentityError(f"{PROVIDER}: calendar has no selected-season control")
    expected = f"{season_key(season)}/{LEAGUE_ID}/"
    options = selected.find_all(tag="option")
    chosen = next((option.attrs.get("value") for option in options if "selected" in option.attrs), None)
    if chosen != expected:
        raise IdentityError(f"{PROVIDER}: calendar scope {chosen!r} does not match {expected!r}")


@dataclass(frozen=True)
class CalendarGame:
    game_no: int
    season: int
    home_name: str
    away_name: str
    scheduled_utc: str
    phase: str
    has_final: bool


def parse_calendar(html: str, season: int) -> List[CalendarGame]:
    """Parse a scoped official season calendar without inventing identities."""
    root = _dom(html)
    _scope_calendar(root, season)
    games: List[CalendarGame] = []
    game_numbers: List[int] = []
    unpinned_rows = 0
    expected = (season_key(season), LEAGUE_ID)
    for table in root.find_all(tag="table", class_name="league_calendar_table"):
        panel = table.parent
        while panel is not None and "panel" not in panel.classes():
            panel = panel.parent
        title = panel.find(tag="h4", class_name="panel-title") if panel else None
        # Bootstrap's accordion toggle appends a private-use icon glyph to the
        # otherwise user-facing round label. It is presentation, not phase ID.
        phase = re.sub(r"[\ue000-\uf8ff]", "", title.text()).strip() if title else "League"
        header = table.find(tag="thead")
        body = table.find(tag="tbody")
        header_row = header.find(tag="tr") if header else None
        if header_row is None or body is None:
            raise IdentityError(f"{PROVIDER}: calendar table has no header/body")
        labels = [cell.text() for cell in header_row.direct("th", "td")]
        if labels not in (["Game", "Result", "Date and Time", "Info"], ["Game", "Result", "Date and Time", "Info", "Group"]):
            raise IdentityError(f"{PROVIDER}: calendar header drift {labels!r}")
        for row in body.direct("tr"):
            cells = row.direct("td", "th")
            if len(cells) != len(labels):
                raise IdentityError(f"{PROVIDER}: calendar row width {len(cells)} != {len(labels)}")
            full = cells[0].find(tag="p", class_name="hidden-xs")
            anchor = full.find(tag="a") if full else None
            href = anchor.attrs.get("href", "") if anchor else ""
            match = re.search(r"/match/(\d+)/(\d+)/(\d+)/Overview", href)
            if match is None or anchor is None:
                raise IdentityError(f"{PROVIDER}: calendar row lacks a scoped match link")
            game_no, sez, lea = (int(value) for value in match.groups())
            if (sez, lea) != expected:
                raise IdentityError(f"{PROVIDER}: calendar link {sez}/{lea} is outside requested scope")
            game_numbers.append(game_no)
            home, away = _pair(anchor, "calendar matchup")
            result = cells[1].text()
            score = re.fullmatch(r"(\d+)\s*:\s*(\d+)", result)
            if score is None and result:
                raise IdentityError(f"{PROVIDER}: calendar result is malformed: {result!r}")
            when = cells[2].text()
            if _is_unpinned_date_range(when):
                if score is not None:
                    raise IdentityError(
                        f"{PROVIDER}: completed calendar game {game_no} has a date range but no exact tip"
                    )
                unpinned_rows += 1
                # The calendar can publish a future pairing's multi-day window
                # before selecting one game tip. It is not a canonical fixture:
                # retain a visible, non-fatal coverage record and do not invent
                # a tip or attempt its match detail.
                _gap(
                    _game_ref(season, game_no),
                    "official_schedule_time",
                    f"calendar game {game_no} has an unpinned date range {when!r}; no canonical game emitted",
                    reason=CoverageState.NOT_APPLICABLE,
                )
                continue
            games.append(
                CalendarGame(
                    game_no=game_no,
                    season=season,
                    home_name=canonical_team_name(home),
                    away_name=canonical_team_name(away),
                    scheduled_utc=_parse_when(when, f"calendar game {game_no}"),
                    phase=phase,
                    has_final=score is not None,
                )
            )
    if len(set(game_numbers)) != len(game_numbers):
        raise IdentityError(f"{PROVIDER}: calendar has duplicate game numbers")
    if not games and not unpinned_rows:
        raise IdentityError(f"{PROVIDER}: calendar has no scoped games")
    return games


def _match_teams(root: _Node, season: int, game_no: int) -> tuple[str, str, str, str]:
    block = root.find(tag="div", ident="match_clubs_and_results_info_table")
    if block is None:
        raise IdentityError(f"{PROVIDER}: match {game_no} has no clubs/result header")
    expected = (season_key(season), LEAGUE_ID)
    teams: List[tuple[str, str]] = []
    for cell in block.find_all(tag="td", class_name="hidden-xs"):
        club = cell.find(tag="span", class_name="club")
        anchor = club.find(tag="a") if club else None
        match = re.search(r"/team/(\d+)/(\d+)/(\d+)/", anchor.attrs.get("href", "") if anchor else "")
        if match is None or anchor is None:
            continue
        team_id, sez, lea = (int(value) for value in match.groups())
        if (sez, lea) != expected:
            raise IdentityError(f"{PROVIDER}: match {game_no} team is outside requested scope")
        teams.append((str(team_id), canonical_team_name(anchor.text())))
    if len(teams) != 2 or teams[0][0] == teams[1][0]:
        raise IdentityError(f"{PROVIDER}: match {game_no} has unresolved home/away team IDs")
    return teams[0][0], teams[0][1], teams[1][0], teams[1][1]


def _match_scope(root: _Node, season: int, game_no: int) -> None:
    wanted = {"id": game_no, "sez": season_key(season), "lea": LEAGUE_ID}
    for key, expected in wanted.items():
        actual = _int(_hidden(root, f"match_data_{key}"), f"match_data_{key}")
        if actual != expected:
            raise IdentityError(f"{PROVIDER}: match scope {key}={actual} does not match {expected}")


def _match_scheduled_utc(root: _Node, game_no: int) -> str:
    info = root.find(tag="div", ident="basic_match_info")
    venue = info.find(tag="div", class_name="dateAndVenue_container") if info else None
    text = venue.text() if venue else ""
    match = re.search(
        r"(?:[A-Za-z]+, )?(\d{2}\.\d{2}\.\d{4} \d{2}:\d{2} CET)", text
    )
    if match is None:
        raise IdentityError(f"{PROVIDER}: match {game_no} has no readable scheduled timestamp")
    return _parse_when(match.group(1), f"match {game_no}")


def _match_actual(root: _Node, season: int, game_no: int) -> ActualResult:
    clubs = root.find(tag="div", ident="match_clubs_and_results_info_table")
    score_cell = clubs.find(tag="td", class_name="gameScore") if clubs else None
    final = re.fullmatch(r"(\d+)\s*:\s*(\d+)", score_cell.text() if score_cell else "")
    if final is None:
        raise IdentityError(f"{PROVIDER}: match {game_no} has no completed final")
    table = root.find(tag="table", ident="match_cetrtine_rezultat")
    body = table.find(tag="tbody") if table else None
    rows = (body.direct("tr") if body else []) or (table.direct("tr") if table else [])
    if len(rows) != 2:
        raise IdentityError(f"{PROVIDER}: match {game_no} period table shape is invalid")
    labels = [cell.text() for cell in rows[0].direct("td", "th")]
    values = [cell.text() for cell in rows[1].direct("td", "th")]
    if len(labels) != len(values) or labels[:4] != ["Q1", "Q2", "Q3", "Q4"]:
        raise IdentityError(f"{PROVIDER}: match {game_no} has invalid period labels {labels!r}")
    periods: List[tuple[int, int]] = []
    for label, value in zip(labels, values):
        score = re.fullmatch(r"(\d+)\s*:\s*(\d+)", value)
        if score is None or re.fullmatch(r"Q[1-4]|OT\d+", label) is None:
            raise IdentityError(f"{PROVIDER}: match {game_no} has malformed {label!r} period")
        periods.append((int(score.group(1)), int(score.group(2))))
    home_quarters = [float(home) for home, _away in periods[:4]]
    away_quarters = [float(away) for _home, away in periods[:4]]
    overtime = periods[4:]
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_overtime=sum(home for home, _away in overtime) if overtime else None,
        away_overtime=sum(away for _home, away in overtime) if overtime else None,
        home_final=float(final.group(1)),
        away_final=float(final.group(2)),
        overtime_periods=len(overtime),
        source=f"{PROVIDER} match/{game_no}/{season_key(season)}/{LEAGUE_ID}",
    )
    valid, detail = actual.reconciles()
    if not valid:
        raise IdentityError(f"{PROVIDER}: match {game_no} does not reconcile: {detail}")
    return actual


def _match_game(html: str, calendar: CalendarGame, *, require_final: bool) -> Game:
    root = _dom(html)
    _match_scope(root, calendar.season, calendar.game_no)
    match_tip = _match_scheduled_utc(root, calendar.game_no)
    if match_tip != calendar.scheduled_utc:
        raise IdentityError(
            f"{PROVIDER}: match {calendar.game_no} scheduled time contradicts calendar "
            f"({match_tip} != {calendar.scheduled_utc})"
        )
    home_id, home_name, away_id, away_name = _match_teams(root, calendar.season, calendar.game_no)
    if (
        normalize_team_name(home_name) != normalize_team_name(calendar.home_name)
        or normalize_team_name(away_name) != normalize_team_name(calendar.away_name)
    ):
        raise IdentityError(
            f"{PROVIDER}: match {calendar.game_no} header team identity contradicts calendar"
        )
    actual = _match_actual(root, calendar.season, calendar.game_no) if require_final else None
    return Game(
        game_id=_game_ref(calendar.season, calendar.game_no),
        league=LEAGUE,
        season=_season_label(calendar.season),
        home_team_id=f"{PROVIDER}:team:{home_id}",
        away_team_id=f"{PROVIDER}:team:{away_id}",
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=calendar.scheduled_utc,
        # Calendar finality is only an index signal here. ``actual`` stays
        # absent until the individual match page's periods reconcile.
        status=GameStatus.FINAL if calendar.has_final else GameStatus.SCHEDULED,
        phase=calendar.phase,
        provider_ids={PROVIDER: _provider_event_id(calendar.season, calendar.game_no)},
        actual=actual,
        evidence={
            PROVIDER: {
                "season_end_year": calendar.season,
                "season_key": season_key(calendar.season),
                "league_id": LEAGUE_ID,
                "game_no": calendar.game_no,
                "home_team_id": home_id,
                "away_team_id": away_id,
                "evidence": EVIDENCE,
            }
        },
    )


def _calendar(season: int, fetch: TextFetch) -> List[CalendarGame]:
    return parse_calendar(fetch(calendar_url(season)), season)


def calendar_final_count(
    season: int,
    fetch_factory=None,
    *,
    text_fetch: Optional[TextFetch] = None,
) -> int:
    """Count final-marked rows in one scoped calendar without fetching details."""

    del fetch_factory
    _reset_gaps()
    return sum(item.has_final for item in _calendar(season, _text_fetch(text_fetch)))


def discover_games(
    date_yyyymmdd: str,
    fetch_factory=None,
    *,
    text_fetch: Optional[TextFetch] = None,
    season: Optional[int] = None,
) -> List[Game]:
    """Discover one Belgrade-local ABA slate with stable team IDs.

    The calendar supplies fixture dates and event IDs; each target-day match
    header supplies stable team IDs. A malformed/unscoped header is a coverage
    gap, never a name-only canonical game.
    """
    del fetch_factory
    _reset_gaps()
    target = datetime.strptime(date_yyyymmdd.replace("-", ""), "%Y%m%d").date()
    season = season or season_end_year(target)
    fetch = _text_fetch(text_fetch)
    candidates = [
        game
        for game in _calendar(season, fetch)
        if parse_utc(game.scheduled_utc).astimezone(ZoneInfo(SITE_TZ)).date() == target
    ]
    games: List[Game] = []
    for calendar in candidates:
        try:
            games.append(_match_game(fetch(match_url(calendar.game_no, season)), calendar, require_final=False))
        except Exception as exc:
            _gap(_game_ref(season, calendar.game_no), "official_schedule_identity", str(exc))
    return games


def fetch_recent_results(
    date_yyyymmdd: str,
    fetch_factory=None,
    *,
    text_fetch: Optional[TextFetch] = None,
    season: Optional[int] = None,
) -> List[Game]:
    """Return only official, reconciling ABA finals for one Belgrade-local day."""
    del fetch_factory
    _reset_gaps()
    target = datetime.strptime(date_yyyymmdd.replace("-", ""), "%Y%m%d").date()
    season = season or season_end_year(target)
    fetch = _text_fetch(text_fetch)
    finals: List[Game] = []
    for calendar in _calendar(season, fetch):
        local_day = parse_utc(calendar.scheduled_utc).astimezone(ZoneInfo(SITE_TZ)).date()
        if local_day != target or not calendar.has_final:
            continue
        try:
            finals.append(_match_game(fetch(match_url(calendar.game_no, season)), calendar, require_final=True))
        except Exception as exc:
            _gap(_game_ref(season, calendar.game_no), "official_result", str(exc))
    return finals


def season_finals(
    season: int,
    *,
    start_after_game_no: int = 0,
    limit: Optional[int] = None,
    fetch_factory=None,
    text_fetch: Optional[TextFetch] = None,
) -> tuple[List[Game], List[int], bool, int]:
    """Fetch a bounded, monotonically ordered batch of official season finals.

    Returns ``(games, attempted_game_numbers, exhausted, calendar_final_count)``.
    The caller owns the durable cursor and advances it only after canonical
    writes; one malformed game therefore cannot silently skip future evidence.
    ``calendar_final_count`` covers the whole selected calendar, including
    final rows below a resumed cursor, so callers can detect a late-posted
    result instead of silently treating a completed season as immutable.
    """
    del fetch_factory
    _reset_gaps()
    if start_after_game_no < 0:
        raise ValueError("start_after_game_no must be non-negative")
    fetch = _text_fetch(text_fetch)
    completed = [item for item in _calendar(season, fetch) if item.has_final]
    candidates = [item for item in completed if item.game_no > start_after_game_no]
    candidates.sort(key=lambda item: item.game_no)
    batch = candidates if limit is None else candidates[: max(0, int(limit))]
    games: List[Game] = []
    attempted: List[int] = []
    stopped_on_error = False
    for calendar in batch:
        attempted.append(calendar.game_no)
        try:
            games.append(_match_game(fetch(match_url(calendar.game_no, season)), calendar, require_final=True))
        except Exception as exc:
            _gap(_game_ref(season, calendar.game_no), "historical_official_result", str(exc))
            stopped_on_error = True
            break
    return games, attempted, not stopped_on_error and len(batch) == len(candidates), len(completed)


def smoke(
    league: League | str = LEAGUE,
    fetch_factory=None,
    *,
    text_fetch: Optional[TextFetch] = None,
    season: Optional[int] = None,
) -> Dict[str, object]:
    """Read one official season calendar without canonical writes."""
    value = league.value if isinstance(league, League) else str(league).strip().upper()
    if value != LEAGUE.value:
        raise ValueError(f"{PROVIDER} supports only {LEAGUE.value}, not {value!r}")
    del fetch_factory
    _reset_gaps()
    season = season or season_end_year(datetime.now(ZoneInfo(SITE_TZ)).date())
    games = _calendar(season, _text_fetch(text_fetch))
    return {
        "provider": PROVIDER,
        "league": LEAGUE.value,
        "verification": EVIDENCE,
        "season_end_year": season,
        "calendar_games": len(games),
        "calendar_finals": sum(game.has_final for game in games),
        "gaps": len(_LAST_GAPS),
    }


__all__ = [
    "EVIDENCE",
    "FIRST_SEASON_END_YEAR",
    "HOST",
    "LEAGUE",
    "LEAGUE_ID",
    "PROVIDER",
    "SITE_TZ",
    "calendar_final_count",
    "calendar_url",
    "discover_games",
    "fetch_recent_results",
    "is_enabled",
    "last_gaps",
    "last_quota",
    "match_url",
    "parse_calendar",
    "season_end_year",
    "season_finals",
    "season_key",
    "smoke",
]
