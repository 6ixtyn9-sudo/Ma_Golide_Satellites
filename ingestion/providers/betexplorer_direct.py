"""Bounded direct BetExplorer named-book fallback for basketball prices.

This is deliberately separate from :mod:`ingestion.providers.betexplorer`.
That adapter records the fixture table's representative *average* as research
only.  This adapter reads the named bookmaker rows displayed in BetExplorer's
full-game comparison responses: Home/Away two-way moneyline, Over/Under
totals, and Asian Handicap spreads.

The route is an operator-authorised, direct public-page fallback for this
single-user project.  It is intentionally conservative operationally:

* one date-board fetch per distinct candidate Paris-local date, followed only
  by matched candidate fixtures;
* at most the source's three exact full-game comparison tabs (``ha``, ``ou``,
  and ``ah``) per matched fixture — no browser, proxy, relay, login, or
  alternate egress path;
* a three-second minimum request interval, hard request cap, and run-scoped
  403/429 cooldown; and
* strict ordered home/away and slate-date matching before a price is emitted.

Every response is represented in the capped run receipt with its source URL
and SHA-256 digest.  A named bookmaker is retained only when the response
itself supplies the bookmaker label; the adapter never attributes a fixture
average or a cross-book best price to a bookmaker.
"""

from __future__ import annotations

import hashlib
import html
import json
import math
import os
import re
import time
import unicodedata
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Iterable, List, Optional
from zoneinfo import ZoneInfo

from ..http import HttpClient, HttpError
from ..identity import canonical_team_key, canonical_team_name, now_utc_iso, parse_utc
from ..models import (
    BookSnapshot,
    CoverageGap,
    CoverageState,
    GameStatus,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
    payload_hash,
)

PROVIDER = "betexplorer_direct"
BASE_URL = "https://www.betexplorer.com"
USER_AGENT = "MaGolide BetExplorer fallback/1.0 (single-user price capture)"
MIN_REQUEST_INTERVAL_SECONDS = 3.0
# One dated board plus three full-game tabs per fixture.  This permits a large
# NCAA slate while remaining below the bounded 45-minute workflow lane.
MAX_REQUESTS_PER_RUN = int(os.environ.get("BETEXPLORER_DIRECT_MAX_REQUESTS", "180"))
MAX_RECEIPT_FIXTURES = 20
MAX_RECEIPT_CANDIDATES = 100
MAX_DIAGNOSTIC_BOARD_ROWS = 25
MAX_BOOKS_PER_FIXTURE = 40

TextFetch = Callable[[str], str]

# A daily basketball board has an explicit date parameter and supplies the
# exact match page URL/event id for each candidate.  It avoids probing every
# league endpoint merely to find a price for a canonical slate game.
_DAY_URL = BASE_URL + "/basketball/results/?year={year}&month={month}&day={day}"
_ODDS_URL = BASE_URL + "/match-odds/{event_id}/{page}/{tab}/bestOdds/?lang=en"
# The live basketball comparison response for Aa4XfIcJ (Austin Peay vs UTRGV,
# captured 2026-11-21) identifies these source tab names.  ``ha`` is the
# Home/Away moneyline tab; ``ah`` is Asian Handicap.
_MARKETS = (
    ("ha", Market.MONEYLINE, "moneyline"),
    ("ou", Market.TOTAL, "total"),
    ("ah", Market.SPREAD, "spread"),
)

_ROW_RE = re.compile(r"<tr\b(?P<attrs>[^>]*)>(?P<body>.*?)</tr>", re.I | re.S)
_LINK_RE = re.compile(
    r'<a\b[^>]*href=["\'](?P<href>/basketball/[^"\']+)["\'][^>]*>(?P<body>.*?)</a>',
    re.I | re.S,
)
_SPAN_RE = re.compile(r"<span\b[^>]*>(?P<name>.*?)</span>", re.I | re.S)
_DATA_DT_RE = re.compile(r'data-dt=["\'](?P<value>[^"\']+)["\']', re.I)
_ODD_RE = re.compile(r'data-odd=["\'](?P<value>[0-9]+(?:\.[0-9]+)?)["\']', re.I)
_BOOK_TITLE_RE = re.compile(
    r'<span\b[^>]*class=["\'][^"\']*in-bookmaker-logo[^"\']*["\'][^>]*'
    r'title=["\'](?P<book>[^"\']+)["\']',
    re.I | re.S,
)
_BOOK_LINK_RE = re.compile(
    r'<a\b[^>]*class=["\'][^"\']*in-bookmaker-logo-link[^"\']*["\'][^>]*>'
    r'(?P<book>.*?)</a>',
    re.I | re.S,
)
_BOOK_DATA_RE = re.compile(r'data-bookie=["\'](?P<book>[^"\']+)["\']', re.I)
_PAGE_PARAM_RE = re.compile(
    r"match_load_tabs\('\w+',\s*'ha',\s*'[^']*',\s*'[^']*',\s*'(?P<page>[^']+)'",
    re.I,
)
# Comparison responses themselves carry a source-supplied page value in every
# tab's navigation callback.  Prefer it over an extra match-page request when
# a comparison response unexpectedly has no named rows.
_ODDS_NAV_PAGE_PARAM_RE = re.compile(
    r"match_change_bettype_best\(\s*(?P<page>[1-9][0-9]{0,3})\s*,\s*['\"][^'\"]+['\"]\s*,\s*"
    r"['\"](?P<tab>ha|ou|ah)['\"]",
    re.I,
)
_ALL_BOOK_ROW_RE = re.compile(
    r'<div\b[^>]*class=["\'][^"\']*oddsComparisonAll__rowBookie[^"\']*["\'][^>]*>'
    r'(?P<body>.*?)(?=<div\b[^>]*class=["\'][^"\']*oddsComparisonAll__rowBookie|\Z)',
    re.I | re.S,
)
_TABLE_RE = re.compile(r"<table\b(?P<attrs>[^>]*)>(?P<body>.*?)</table>", re.I | re.S)
_HANDICAP_RE = re.compile(r'data-handicap=["\'](?P<line>[+-]?[0-9]+(?:\.[0-9]+)?)["\']', re.I)
_ACTIVE_LINE_RE = re.compile(
    r'<li\b(?P<attrs>[^>]*(?:class=["\'][^"\']*oddsComparison__activeSubLi[^"\']*["\'])[^>]*)>',
    re.I | re.S,
)
_ID_RE = re.compile(r'id=["\'](?P<line>[+-]?[0-9]+(?:\.[0-9]+)?)["\']', re.I)
# A verification page is not an empty BetExplorer market.  These deliberately
# specific markers cover the source's common challenge interstitials without
# treating ordinary odds prose as a refusal.
_CHALLENGE_MARKERS = (
    "captcha",
    "checking your browser",
    "verify you are human",
    "attention required",
    "cf-chl-",
    "challenge-platform",
)

_last_gaps: List[CoverageGap] = []
_last_receipt: Dict[str, Any] = {}
_last_request_at = 0.0


def is_enabled() -> bool:
    """The direct, operator-authorised fallback has no credential gate."""

    return True


def last_gaps() -> List[CoverageGap]:
    """Return gaps recorded by the most recent direct fallback run."""

    return list(_last_gaps)


def last_receipt() -> Dict[str, Any]:
    """Return a compact, source-URL/digest receipt for the latest run."""

    return dict(_last_receipt)


def last_quota() -> Dict[str, Any]:
    """Expose bounded request/cooldown telemetry through the pipeline quota slot."""

    return {
        key: _last_receipt.get(key)
        for key in (
            "status",
            "requests",
            "request_cap",
            "cooldown",
            "day_rows_seen",
            "candidate_games",
            "fixtures_matched",
            "fixtures_unmatched",
            "fixtures_with_named_books",
            "snapshots_emitted",
        )
        if key in _last_receipt
    }


def _reset_run(league: str, date_yyyymmdd: str, observed_utc: str) -> None:
    global _last_gaps, _last_receipt
    _last_gaps = []
    _last_receipt = {
        "provider": PROVIDER,
        "league": str(league).strip().upper(),
        "date": _date_text(date_yyyymmdd),
        "observed_utc": observed_utc,
        "status": "started",
        "requests": 0,
        "request_cap": MAX_REQUESTS_PER_RUN,
        "cooldown": False,
        "day_rows_seen": 0,
        "candidate_games": 0,
        "fixtures_matched": 0,
        "fixtures_unmatched": 0,
        "fixtures_with_named_books": 0,
        "snapshots_emitted": 0,
        "identity_rejections": {},
        "fixture_receipts": [],
    }


def _date_text(value: str) -> str:
    text = str(value or "").strip().replace("-", "")
    if len(text) != 8 or not text.isdigit():
        raise ValueError("date_yyyymmdd must be YYYYMMDD or YYYY-MM-DD")
    return f"{text[:4]}-{text[4:6]}-{text[6:]}"


def _default_text_fetch() -> TextFetch:
    client = HttpClient(timeout=20.0, retries=0, user_agent=USER_AGENT)
    return client.get_text


def _text_fetch(
    fetch_factory: Optional[Callable[[str], Callable[[str], Any]]],
    text_fetch: Optional[TextFetch],
) -> TextFetch:
    """Return raw HTML/JSON text transport; fixture callers inject text directly."""

    del fetch_factory  # JSON provider factories cannot safely represent HTML pages.
    return text_fetch or _default_text_fetch()


def _record_gap(
    game_id: Optional[str], field: str, detail: str, state: CoverageState = CoverageState.MISSING
) -> None:
    _last_gaps.append(
        CoverageGap(
            game_id=game_id,
            required_field=field,
            reason=state,
            provider=PROVIDER,
            detail=detail,
        )
    )


def _count_rejection(reason: str) -> None:
    counts = _last_receipt.setdefault("identity_rejections", {})
    counts[reason] = int(counts.get(reason, 0)) + 1


def _throttle() -> None:
    global _last_request_at
    elapsed = time.monotonic() - _last_request_at
    if elapsed < MIN_REQUEST_INTERVAL_SECONDS:
        time.sleep(MIN_REQUEST_INTERVAL_SECONDS - elapsed)
    _last_request_at = time.monotonic()


def _fetch_text(fetch: TextFetch, url: str, *, throttle: bool) -> str:
    """Fetch one bounded response or trip run cooldown on source refusal."""

    if _last_receipt.get("cooldown"):
        raise RuntimeError("BetExplorer direct fallback is cooling down after source refusal")
    if int(_last_receipt.get("requests", 0)) >= MAX_REQUESTS_PER_RUN:
        _last_receipt["status"] = "request_cap_reached"
        _record_gap(None, "request_budget", "bounded BetExplorer request cap reached", CoverageState.UNSUPPORTED)
        raise RuntimeError("BetExplorer direct fallback request cap reached")
    if throttle:
        _throttle()
    _last_receipt["requests"] = int(_last_receipt.get("requests", 0)) + 1
    try:
        body = fetch(url)
    except HttpError as exc:
        if exc.status in {403, 429}:
            _last_receipt["cooldown"] = True
            _last_receipt["status"] = f"http_{exc.status}_cooldown"
            _record_gap(None, "source_access", f"HTTP {exc.status}; no further requests this run", CoverageState.PROVIDER_ERROR)
        raise
    if not isinstance(body, str):
        raise TypeError("BetExplorer transport must return text")
    lowered = body.casefold()
    if any(marker in lowered for marker in _CHALLENGE_MARKERS):
        _last_receipt["cooldown"] = True
        _last_receipt["status"] = "source_challenge"
        _record_gap(
            None,
            "source_access",
            "BetExplorer challenge/verification surface; no further requests this run",
            CoverageState.PROVIDER_ERROR,
        )
        raise RuntimeError("BetExplorer returned a challenge/verification surface")
    encoded = body.encode("utf-8", "replace")
    if len(encoded) > 1_500_000:
        raise ValueError("BetExplorer response exceeds 1.5 MB bound")
    return body


def _visible_text(value: str) -> str:
    text = re.sub(r"<[^>]+>", " ", value or "")
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def _event_id(href: str) -> Optional[str]:
    parts = [part for part in str(href or "").split("?")[0].split("/") if part]
    if not parts:
        return None
    value = parts[-1]
    return value if re.fullmatch(r"[A-Za-z0-9_-]{4,64}", value) else None


def _teams(link_body: str) -> tuple[Optional[str], Optional[str]]:
    names = [_visible_text(item) for item in _SPAN_RE.findall(link_body)]
    names = [name for name in names if name]
    if len(names) == 2:
        return names[0], names[1]
    text = _visible_text(link_body)
    if " - " not in text:
        return None, None
    home, away = (part.strip() for part in text.split(" - ", 1))
    return (home or None), (away or None)


def _row_date(attrs: str) -> Optional[str]:
    match = _DATA_DT_RE.search(attrs)
    if not match:
        return None
    values = [value.strip() for value in match.group("value").split(",")]
    if len(values) < 3 or not all(value.isdigit() for value in values[:3]):
        return None
    day, month, year = (int(value) for value in values[:3])
    try:
        return datetime(year, month, day, tzinfo=timezone.utc).strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_day_board(body: str) -> List[Dict[str, str]]:
    """Read only explicit basketball fixture rows from a dated BetExplorer board."""

    fixtures: List[Dict[str, str]] = []
    for row in _ROW_RE.finditer(body or ""):
        link = _LINK_RE.search(row.group("body"))
        if link is None:
            continue
        event_id = _event_id(html.unescape(link.group("href")))
        home, away = _teams(link.group("body"))
        date = _row_date(row.group("attrs"))
        if not (event_id and home and away and date):
            continue
        fixtures.append(
            {
                "event_id": event_id,
                "home_name": canonical_team_name(home),
                "away_name": canonical_team_name(away),
                "date": date,
                "match_url": BASE_URL + html.unescape(link.group("href")),
            }
        )
    return fixtures


def _book_name(segment: str) -> Optional[str]:
    for pattern in (_BOOK_TITLE_RE, _BOOK_LINK_RE, _BOOK_DATA_RE):
        match = pattern.search(segment)
        if match is None:
            continue
        name = _visible_text(match.group("book"))
        if name:
            return name
    return None


def _decimal(value: str) -> Optional[float]:
    try:
        parsed = float(str(value))
    except (TypeError, ValueError):
        return None
    return parsed if math.isfinite(parsed) and parsed >= 1.01 else None


def parse_named_moneyline_rows(odds_html: str) -> List[Dict[str, Any]]:
    """Return exact named-book two-way H/A rows; ignore averages and unnamed rows."""

    rows: Dict[str, Dict[str, Any]] = {}
    # The compact comparison table uses tr elements.  The response's optional
    # all-bookmaker view uses div.oddsComparisonAll__rowBookie instead; retain
    # both shapes without treating the adjacent average widgets as a bookmaker.
    segments = [row.group(0) for row in _ROW_RE.finditer(odds_html or "")]
    segments.extend(row.group(0) for row in _ALL_BOOK_ROW_RE.finditer(odds_html or ""))
    for segment in segments:
        book = _book_name(segment)
        odds = [_decimal(value) for value in _ODD_RE.findall(segment)]
        if not book or len(odds) < 2 or any(value is None for value in odds[:2]):
            continue
        # Comparison responses can expose a compact best table followed by an
        # all-bookmaker block.  The latter occurs later and is the fuller view,
        # so it intentionally replaces the compact duplicate for one book.
        rows[book.casefold()] = {
            "book": book,
            "home_price_decimal": odds[0],
            "away_price_decimal": odds[1],
        }
    return list(rows.values())[:MAX_BOOKS_PER_FIXTURE]


def parse_named_line_rows(odds_html: str) -> List[Dict[str, Any]]:
    """Read named two-sided total/spread rows with the source's exact line.

    BetExplorer's basketball ``ou`` and ``ah`` responses put the full-game
    threshold/handicap in a table ``data-handicap`` attribute.  Each source row
    supplies the bookmaker name plus the two displayed decimal prices.  No
    average widget, unnamed row, or guessed line becomes a snapshot.
    """

    rows: Dict[tuple[str, float], Dict[str, Any]] = {}
    for table in _TABLE_RE.finditer(odds_html or ""):
        line_match = _HANDICAP_RE.search(table.group("attrs"))
        if line_match is None:
            continue
        try:
            line = float(line_match.group("line"))
        except ValueError:
            continue
        if not math.isfinite(line):
            continue
        for row in _ROW_RE.finditer(table.group("body")):
            segment = row.group(0)
            book = _book_name(segment)
            odds = [_decimal(value) for value in _ODD_RE.findall(segment)]
            if not book or len(odds) < 2 or any(value is None for value in odds[:2]):
                continue
            rows[(book.casefold(), line)] = {
                "book": book,
                "line": line,
                "first_price_decimal": odds[0],
                "second_price_decimal": odds[1],
            }
    return list(rows.values())[:MAX_BOOKS_PER_FIXTURE]


def _response_lines(odds_html: str) -> set[float]:
    """Return the exact O/U or AH lines that the response actually carries."""

    lines: set[float] = set()
    for value in _HANDICAP_RE.findall(odds_html or ""):
        try:
            line = float(value)
        except ValueError:
            continue
        if math.isfinite(line):
            lines.add(line)
    return lines


def _source_active_line(odds_html: str) -> Optional[float]:
    """Return a source-active O/U or AH line only when it is actually carried.

    Basketball comparison responses can expose stale cross-sport navigation
    ids, for example an active ``2.50`` alongside a total table at ``133.50``.
    A marker that does not match a table ``data-handicap`` is therefore not an
    active basketball line.  No average, midpoint, or book-count substitute is
    invented when the source has no trustworthy marker.
    """

    carried_lines = _response_lines(odds_html)
    for item in _ACTIVE_LINE_RE.finditer(odds_html or ""):
        line_match = _ID_RE.search(item.group("attrs"))
        if line_match is None:
            continue
        try:
            line = float(line_match.group("line"))
        except ValueError:
            continue
        if math.isfinite(line) and line in carried_lines:
            return line
    return None


def _response_odds_html(body: str) -> str:
    """Unwrap the page's comparison JSON without accepting invented shapes."""

    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return body
    if isinstance(payload, dict) and isinstance(payload.get("odds"), str):
        return payload["odds"]
    return ""


def _odds_nav_page_param(odds_html: str, tab: str) -> Optional[str]:
    """Read an exact tab's page value from its response navigation callback."""

    expected_tab = str(tab or "").casefold()
    for match in _ODDS_NAV_PAGE_PARAM_RE.finditer(odds_html or ""):
        if match.group("tab").casefold() == expected_tab:
            return match.group("page")
    return None


def _comparison_page_param(match_page: str) -> str:
    """Read the page-supplied basketball H/A tab parameter; default to page 1."""

    match = _PAGE_PARAM_RE.search(match_page or "")
    return match.group("page") if match and match.group("page") else "1"


def _scheduled_candidate_games(
    games: Optional[Iterable[Any]], observed_utc: str
) -> List[Any]:
    observed = parse_utc(observed_utc)
    output = []
    for game in games or ():
        scheduled = getattr(game, "scheduled_utc", None)
        status = getattr(game, "status", None)
        if not scheduled or status is not GameStatus.SCHEDULED:
            continue
        try:
            if parse_utc(scheduled) <= observed:
                continue
        except (TypeError, ValueError):
            continue
        if not (getattr(game, "home_name", None) and getattr(game, "away_name", None)):
            continue
        output.append(game)
    return output


def _candidate_board_date(game: Any) -> str:
    """Return the BetExplorer-local calendar date for a scheduled game."""

    scheduled = parse_utc(game.scheduled_utc)
    return scheduled.astimezone(ZoneInfo("Europe/Paris")).date().isoformat()


def _team_name_tokens(name: str) -> set[str]:
    """Tokenize names consistently for diagnostics and conservative matching."""

    decomposed = unicodedata.normalize("NFKD", str(name or "")).casefold()
    folded = "".join(character for character in decomposed if not unicodedata.combining(character))
    return set(re.findall(r"[^\W_]+", folded, flags=re.UNICODE))


def _team_name_words(name: str) -> List[str]:
    decomposed = unicodedata.normalize("NFKD", str(name or "")).casefold()
    folded = "".join(character for character in decomposed if not unicodedata.combining(character))
    return re.findall(r"[^\W_]+", folded, flags=re.UNICODE)


_SQUAD_MARKERS = {"b", "ii", "iii", "w", "women", "womens", "ladies", "reserve", "reserves", "res", "youth", "academy"}


def _matching_team_tokens(name: str, *, league: str) -> set[str]:
    """Apply narrowly scoped, deterministic sports-name word equivalences."""

    tokens = _team_name_tokens(name)
    if str(league or "").strip().upper() in {"NCAA_M", "NCAA_W"}:
        # In US college team labels, "St." is the conventional abbreviation
        # for "Saint" (e.g. St. John's / Saint John's). Scope this to NCAA so
        # unrelated source labels are not rewritten.
        tokens.discard("st")
        if "st" in _team_name_words(name):
            tokens.add("saint")
    return tokens


def _squad_markers(name: str) -> set[str]:
    markers = set()
    for word in _team_name_words(name):
        if re.fullmatch(r"u(?:1[4-9]|2[0-3])", word):
            markers.add(word)
        elif word in _SQUAD_MARKERS:
            markers.add(word)
    return markers


_SEASON_SCOPED_NAME_PAIRS = {
    # Confirmed by the 2026-11-22 NCAA_M board diagnostics. Keep these source-
    # specific spellings out of global identity aliases and expire after the
    # 2026-27 season unless re-verified.
    ("NCAA_M", "2026-27", "alice lloyd eagles", "alice loyd"),
    ("NCAA_M", "2026-27", "saint marys gaels", "st marys ca"),
}


def _team_names_agree(
    candidate_name: str,
    board_name: str,
    *,
    league: str,
    board_date: str = "",
) -> tuple[bool, str]:
    """Edge-inspired deterministic name agreement; never edit-distance fuzzy."""

    candidate = str(candidate_name or "")
    board = str(board_name or "")
    league_key = str(league or "").strip().upper()
    season_key = ""
    try:
        season_year, season_month, _season_day = (int(part) for part in board_date.split("-"))
        if (season_year == 2026 and season_month >= 7) or (
            season_year == 2027 and season_month <= 6
        ):
            season_key = "2026-27"
    except (TypeError, ValueError):
        pass
    candidate_key = canonical_team_key(candidate)
    board_key = canonical_team_key(board)
    for alias_league, alias_season, canonical_name, board_variant in _SEASON_SCOPED_NAME_PAIRS:
        if (
            league_key == alias_league
            and season_key == alias_season
            and (candidate_key, board_key)
            in {
                (canonical_team_key(canonical_name), canonical_team_key(board_variant)),
                (canonical_team_key(board_variant), canonical_team_key(canonical_name)),
            }
        ):
            return True, "season_scoped_alias"

    if league_key == "NCAA_W":
        # BetExplorer labels women's NCAA teams with a terminal W marker;
        # ESPN's NCAA_W schedule lane is already gender-scoped.
        candidate = re.sub(r"\s+W$", "", candidate, flags=re.IGNORECASE)
        board = re.sub(r"\s+W$", "", board, flags=re.IGNORECASE)

    if canonical_team_key(candidate) == canonical_team_key(board):
        return True, "normalized_exact"

    if _squad_markers(candidate) != _squad_markers(board):
        return False, ""
    candidate_digits = set(re.findall(r"\d+", candidate))
    board_digits = set(re.findall(r"\d+", board))
    if candidate_digits and board_digits and candidate_digits != board_digits:
        return False, ""

    candidate_tokens = _matching_team_tokens(candidate, league=league_key)
    board_tokens = _matching_team_tokens(board, league=league_key)
    shared = candidate_tokens & board_tokens
    # Retain Edge's conservative whole-name subset principle: one complete
    # token set must be contained in the other, with at least five shared
    # letters. No partial-token/edit-distance or side-swapping match.
    if shared and (shared == candidate_tokens or shared == board_tokens) and sum(map(len, shared)) >= 5:
        return True, "ordered_token_subset"
    return False, ""


def _nearest_board_rows(
    game: Any, date_rows: List[Dict[str, str]], limit: int = MAX_DIAGNOSTIC_BOARD_ROWS
) -> List[Dict[str, Any]]:
    """Return bounded token-overlap examples for diagnosis, not fixture matching."""

    home_tokens = _team_name_tokens(game.home_name)
    away_tokens = _team_name_tokens(game.away_name)
    ranked: List[tuple[int, int, str, List[str], Dict[str, str]]] = []
    for index, fixture in enumerate(date_rows):
        row_home = _team_name_tokens(fixture["home_name"])
        row_away = _team_name_tokens(fixture["away_name"])
        ordered = sorted(home_tokens & row_home) + sorted(away_tokens & row_away)
        reversed_names = sorted(home_tokens & row_away) + sorted(away_tokens & row_home)
        if len(reversed_names) > len(ordered):
            score, orientation, shared = len(reversed_names), "reversed_diagnostic_only", reversed_names
        else:
            score, orientation, shared = len(ordered), "ordered", ordered
        ranked.append((score, -index, orientation, shared, fixture))
    ranked.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return [
        {
            "home": fixture["home_name"],
            "away": fixture["away_name"],
            "shared_token_count": score,
            "shared_tokens": shared,
            "orientation": orientation,
        }
        for score, _index, orientation, shared, fixture in ranked[:limit]
    ]


def _match_fixtures(
    fixtures: Iterable[Dict[str, str]], games: Iterable[Any], *, league: str
) -> List[tuple[Dict[str, str], Any]]:
    """Bind unique board events by Paris date and guarded ordered team names."""

    candidates = list(games)
    raw_fixture_rows = list(fixtures)
    fixture_rows: List[Dict[str, str]] = []
    seen_rows = set()
    for fixture in raw_fixture_rows:
        # The two adjacent day boards can repeat the same event. Collapse only
        # identical event/date/ordered-team identities; distinct event IDs are
        # retained and will be refused as ambiguous if both match one candidate.
        signature = (
            fixture.get("event_id"),
            fixture.get("date"),
            canonical_team_key(fixture.get("home_name", "")),
            canonical_team_key(fixture.get("away_name", "")),
        )
        if signature in seen_rows:
            continue
        seen_rows.add(signature)
        fixture_rows.append(fixture)
    _last_receipt["duplicate_board_rows_removed"] = len(raw_fixture_rows) - len(fixture_rows)

    fixture_dates = {fixture.get("date") for fixture in fixture_rows}
    fixtures_by_date: Dict[str, List[Dict[str, str]]] = {}
    for fixture in fixture_rows:
        fixtures_by_date.setdefault(fixture.get("date", ""), []).append(fixture)

    diagnostics: Dict[str, Dict[str, Any]] = {}
    for game in candidates:
        board_date = _candidate_board_date(game)
        date_rows = fixtures_by_date.get(board_date, [])
        home_key = canonical_team_key(game.home_name)
        away_key = canonical_team_key(game.away_name)
        exact_rows = [
            fixture
            for fixture in date_rows
            if canonical_team_key(fixture["home_name"]) == home_key
            and canonical_team_key(fixture["away_name"]) == away_key
        ]
        compatible_rows = [
            fixture
            for fixture in date_rows
            if _team_names_agree(game.home_name, fixture["home_name"], league=league, board_date=fixture.get("date", ""))[0]
            and _team_names_agree(game.away_name, fixture["away_name"], league=league, board_date=fixture.get("date", ""))[0]
        ]
        reversed_rows = [
            fixture
            for fixture in date_rows
            if _team_names_agree(game.home_name, fixture["away_name"], league=league, board_date=fixture.get("date", ""))[0]
            and _team_names_agree(game.away_name, fixture["home_name"], league=league, board_date=fixture.get("date", ""))[0]
        ]
        diagnostics[str(game.game_id)] = {
            "game_id": str(game.game_id),
            "scheduled_utc": str(game.scheduled_utc),
            "board_date": board_date,
            "home": str(game.home_name),
            "away": str(game.away_name),
            "board_rows_on_date": len(date_rows),
            "ordered_pair_rows": len(exact_rows),
            "compatible_ordered_pair_rows": len(compatible_rows),
            "reversed_pair_rows_diagnostic_only": len(reversed_rows),
            "nearest_board_rows": _nearest_board_rows(game, date_rows),
            "status": "unmatched",
            "reason": (
                "no_fixture_rows_on_paris_board_date"
                if board_date not in fixture_dates
                else "ordered_team_pair_not_found_on_board_date"
            ),
        }

    hits_by_game: Dict[str, Dict[str, tuple[Dict[str, str], str]]] = {
        str(game.game_id): {} for game in candidates
    }
    ambiguous_game_ids: set[str] = set()
    for fixture in fixture_rows:
        same_date = [game for game in candidates if _candidate_board_date(game) == fixture.get("date")]
        pair: List[tuple[Any, str]] = []
        for game in same_date:
            home_ok, home_method = _team_names_agree(
                game.home_name, fixture["home_name"], league=league, board_date=fixture.get("date", "")
            )
            away_ok, away_method = _team_names_agree(
                game.away_name, fixture["away_name"], league=league, board_date=fixture.get("date", "")
            )
            if home_ok and away_ok:
                method = (
                    "normalized_exact"
                    if home_method == away_method == "normalized_exact"
                    else "ordered_token_subset"
                )
                pair.append((game, method))
        if len(pair) == 1:
            game, method = pair[0]
            event_id = str(fixture.get("event_id") or "")
            hits_by_game[str(game.game_id)].setdefault(event_id, (fixture, method))
        elif pair:
            for game, _method in pair:
                ambiguous_game_ids.add(str(game.game_id))

    matches: List[tuple[Dict[str, str], Any]] = []
    for game in candidates:
        game_id = str(game.game_id)
        hits = hits_by_game[game_id]
        diagnostic = diagnostics[game_id]
        if game_id in ambiguous_game_ids or len(hits) > 1:
            diagnostic.update(status="ambiguous", reason="ambiguous_ordered_team_pair")
            _count_rejection("ambiguous_ordered_team_pair")
            continue
        if len(hits) == 1:
            fixture, method = next(iter(hits.values()))
            matches.append((fixture, game))
            diagnostic.update(
                status="matched",
                reason=(
                    "exact_ordered_pair_on_paris_board_date"
                    if method == "normalized_exact"
                    else "ordered_token_subset_on_paris_board_date"
                ),
                match_method=method,
            )
            continue
        _count_rejection(diagnostic["reason"])

    _last_receipt["candidate_fixture_diagnostics"] = [
        diagnostics[str(game.game_id)] for game in candidates
    ][:MAX_RECEIPT_CANDIDATES]
    _last_receipt["candidate_fixture_diagnostics_truncated"] = max(
        0, len(candidates) - MAX_RECEIPT_CANDIDATES
    )
    return matches


def _receipt_fixture(
    fixture: Dict[str, str],
    game: Any,
    markets: List[Dict[str, Any]],
    *,
    page_param: str,
    match_page: Optional[str] = None,
) -> None:
    receipts = _last_receipt.setdefault("fixture_receipts", [])
    if len(receipts) >= MAX_RECEIPT_FIXTURES:
        return
    all_books = sorted(
        {book for market in markets for book in market.get("named_books", []) if isinstance(book, str)}
    )
    receipts.append(
        {
            "event_id": fixture["event_id"],
            "game_id": game.game_id,
            "scheduled_home_name": game.home_name,
            "scheduled_away_name": game.away_name,
            "board_home_name": fixture["home_name"],
            "board_away_name": fixture["away_name"],
            "match_url": fixture["match_url"],
            "comparison_page_param": page_param,
            "match_page_sha256": (
                hashlib.sha256(match_page.encode("utf-8", "replace")).hexdigest()
                if match_page is not None
                else None
            ),
            "named_books": all_books,
            "named_book_rows": sum(int(market.get("named_book_rows", 0)) for market in markets),
            "markets": markets,
        }
    )


def fetch_odds(
    league: str,
    date_yyyymmdd: str,
    fetch_factory: Optional[Callable[[str], Callable[[str], Any]]] = None,
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
    games: Optional[Iterable[Any]] = None,
    text_fetch: Optional[TextFetch] = None,
) -> List[BookSnapshot]:
    """Fetch direct named-book full-game H/A, total, and spread rows."""

    del tz_name  # BetExplorer fixture rows use Europe/Paris wall-clock dates.
    observed_utc = now_utc or now_utc_iso()
    date = _date_text(date_yyyymmdd)
    _reset_run(league, date, observed_utc)
    candidates = _scheduled_candidate_games(games, observed_utc)
    _last_receipt["candidate_games"] = len(candidates)
    if not candidates:
        _last_receipt["status"] = "no_future_candidate_games"
        return []

    fetch = _text_fetch(fetch_factory, text_fetch)
    throttled = text_fetch is None
    board_dates = sorted({_candidate_board_date(game) for game in candidates})
    fixtures: List[Dict[str, str]] = []
    day_receipts: List[Dict[str, Any]] = []
    for board_date in board_dates:
        year, month, day = board_date.split("-")
        url = _DAY_URL.format(year=year, month=month, day=day)
        try:
            board_body = _fetch_text(fetch, url, throttle=throttled)
        except Exception as exc:
            if not _last_receipt.get("cooldown"):
                _last_receipt["status"] = "board_fetch_error"
                _record_gap(
                    None,
                    "daily_board",
                    f"{type(exc).__name__}: {str(exc)[:160]}",
                    CoverageState.PROVIDER_ERROR,
                )
            return []
        parsed = parse_day_board(board_body)
        fixtures.extend(parsed)
        day_receipts.append(
            {
                "date": board_date,
                "url": url,
                "response_sha256": hashlib.sha256(
                    board_body.encode("utf-8", "replace")
                ).hexdigest(),
                "rows_seen": len(parsed),
            }
        )

    _last_receipt["board_dates"] = day_receipts
    # Keep the original single-board receipt fields for existing consumers.
    if day_receipts:
        _last_receipt["day_url"] = day_receipts[0]["url"]
        _last_receipt["day_response_sha256"] = day_receipts[0]["response_sha256"]
    _last_receipt["day_rows_seen"] = sum(item["rows_seen"] for item in day_receipts)
    matched = _match_fixtures(fixtures, candidates, league=league)
    _last_receipt["fixtures_matched"] = len(matched)
    _last_receipt["fixtures_unmatched"] = max(0, len(candidates) - len(matched))

    snapshots: List[BookSnapshot] = []
    for fixture, game in matched:
        page_param = "1"
        match_page: Optional[str] = None
        fixture_markets: List[Dict[str, Any]] = []
        fixture_has_named_rows = False
        for tab, market, receipt_name in _MARKETS:
            odds_url = _ODDS_URL.format(
                event_id=fixture["event_id"], page=page_param, tab=tab
            )
            try:
                response_body = _fetch_text(fetch, odds_url, throttle=throttled)
            except Exception as exc:
                if not _last_receipt.get("cooldown"):
                    _record_gap(
                        game.game_id,
                        f"named_book_{receipt_name}",
                        f"{type(exc).__name__}: {str(exc)[:160]}",
                        CoverageState.PROVIDER_ERROR,
                    )
                if _last_receipt.get("cooldown"):
                    break
                continue

            odds_html = _response_odds_html(response_body)
            rows = (
                parse_named_moneyline_rows(odds_html)
                if market is Market.MONEYLINE
                else parse_named_line_rows(odds_html)
            )
            if not rows and not _last_receipt.get("cooldown"):
                # A comparison response already carries the source's own page
                # parameter in its nav callback.  Prefer this no-extra-page
                # recovery, then use the guarded match-page read only when it
                # supplies a page that this tab has not already tried.
                attempted_pages = {page_param}
                try:
                    nav_page = _odds_nav_page_param(odds_html, tab)
                    if nav_page is not None and nav_page not in attempted_pages:
                        page_param = nav_page
                        attempted_pages.add(page_param)
                        odds_url = _ODDS_URL.format(
                            event_id=fixture["event_id"], page=page_param, tab=tab
                        )
                        response_body = _fetch_text(fetch, odds_url, throttle=throttled)
                        odds_html = _response_odds_html(response_body)
                        rows = (
                            parse_named_moneyline_rows(odds_html)
                            if market is Market.MONEYLINE
                            else parse_named_line_rows(odds_html)
                        )
                    if not rows and match_page is None:
                        match_page = _fetch_text(fetch, fixture["match_url"], throttle=throttled)
                        recovered_page = _comparison_page_param(match_page)
                        if recovered_page not in attempted_pages:
                            page_param = recovered_page
                            attempted_pages.add(page_param)
                            odds_url = _ODDS_URL.format(
                                event_id=fixture["event_id"], page=page_param, tab=tab
                            )
                            response_body = _fetch_text(fetch, odds_url, throttle=throttled)
                            odds_html = _response_odds_html(response_body)
                            rows = (
                                parse_named_moneyline_rows(odds_html)
                                if market is Market.MONEYLINE
                                else parse_named_line_rows(odds_html)
                            )
                except Exception as exc:
                    if not _last_receipt.get("cooldown"):
                        _record_gap(
                            game.game_id,
                            "comparison_page_parameter",
                            f"{type(exc).__name__}: {str(exc)[:160]}",
                            CoverageState.PROVIDER_ERROR,
                        )

            fixture_markets.append(
                {
                    "market": receipt_name,
                    "tab": tab,
                    "url": odds_url,
                    "response_sha256": hashlib.sha256(
                        response_body.encode("utf-8", "replace")
                    ).hexdigest(),
                    "named_books": [row["book"] for row in rows],
                    "named_book_rows": len(rows),
                    "main_line": (
                        _source_active_line(odds_html) if market is not Market.MONEYLINE else None
                    ),
                }
            )
            if not rows:
                _record_gap(
                    game.game_id,
                    f"named_book_{receipt_name}",
                    "comparison response had no complete named source row",
                )
                continue

            fixture_has_named_rows = True
            main_line = _source_active_line(odds_html) if market is not Market.MONEYLINE else None
            for row in rows:
                common = {
                    "game_id": f"{PROVIDER}:{fixture['event_id']}",
                    "provider": PROVIDER,
                    "book": row["book"],
                    "period": Period.FULL_GAME,
                    "observed_utc": observed_utc,
                    "pre_match": True,
                    "main_line": (
                        True if market is Market.MONEYLINE else float(row["line"]) == main_line
                    ),
                    "payload_hash": payload_hash(
                        {
                            "source": "betexplorer_comparison",
                            "event_id": fixture["event_id"],
                            "tab": tab,
                            "odds_url": odds_url,
                            "book": row["book"],
                            "line": row.get("line"),
                            "first_price_decimal": row.get("first_price_decimal"),
                            "second_price_decimal": row.get("second_price_decimal"),
                        }
                    ),
                    # The central registry re-stamps this direct observation.
                    # Defaults make direct adapter invocation truthful as well.
                    "price_provenance": PriceProvenance.NAMED_BOOK,
                    "source_permission": SourcePermission.UNRESOLVED,
                    "operational_validation": OperationalValidation.PROBATION,
                    "provider_event_id": fixture["event_id"],
                    # The guarded source matcher already bound this event to
                    # the canonical scheduled game. Carry that identity through
                    # downstream mapping; preserve raw board labels in the
                    # fixture receipt for audit instead of re-matching them.
                    "home_name": game.home_name,
                    "away_name": game.away_name,
                    # The canonical scheduled instant is supplied by the
                    # official/schedule adapter.  The dated public board and
                    # ordered teams bind the BetExplorer event to that game;
                    # no vague page-local clock is converted or guessed.
                    "scheduled_utc": game.scheduled_utc,
                }
                if market is Market.MONEYLINE:
                    snapshots.append(
                        BookSnapshot(
                            market=market,
                            home_price_decimal=float(row["home_price_decimal"]),
                            away_price_decimal=float(row["away_price_decimal"]),
                            **common,
                        )
                    )
                elif market is Market.TOTAL:
                    snapshots.append(
                        BookSnapshot(
                            market=market,
                            line=float(row["line"]),
                            over_price_decimal=float(row["first_price_decimal"]),
                            under_price_decimal=float(row["second_price_decimal"]),
                            **common,
                        )
                    )
                else:
                    snapshots.append(
                        BookSnapshot(
                            market=market,
                            line=float(row["line"]),
                            home_price_decimal=float(row["first_price_decimal"]),
                            away_price_decimal=float(row["second_price_decimal"]),
                            **common,
                        )
                    )
        _receipt_fixture(
            fixture,
            game,
            fixture_markets,
            page_param=page_param,
            match_page=match_page,
        )
        if fixture_has_named_rows:
            _last_receipt["fixtures_with_named_books"] = int(
                _last_receipt["fixtures_with_named_books"]
            ) + 1
        if _last_receipt.get("cooldown"):
            break
    _last_receipt["snapshots_emitted"] = len(snapshots)
    _last_receipt["partial_snapshots_preserved_on_cooldown"] = bool(
        _last_receipt.get("cooldown") and snapshots
    )
    if _last_receipt.get("status") == "started":
        _last_receipt["status"] = "success" if snapshots else "no_named_book_prices"
    return snapshots


def smoke(
    league: str,
    fetch_factory: Optional[Callable[[str], Callable[[str], Any]]] = None,
    date_yyyymmdd: Optional[str] = None,
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
    text_fetch: Optional[TextFetch] = None,
) -> Dict[str, Any]:
    """Bounded read-only board probe; no fixture comparison requests are made."""

    del tz_name
    observed_utc = now_utc or now_utc_iso()
    date = _date_text(date_yyyymmdd or observed_utc[:10])
    _reset_run(league, date, observed_utc)
    fetch = _text_fetch(fetch_factory, text_fetch)
    url = _DAY_URL.format(year=date[:4], month=date[5:7], day=date[8:])
    try:
        body = _fetch_text(fetch, url, throttle=text_fetch is None)
    except Exception as exc:
        if not _last_receipt.get("cooldown"):
            _last_receipt["status"] = "board_fetch_error"
        return {
            "enabled": True,
            "provider": PROVIDER,
            "league": str(league).upper(),
            "verification": "DIRECT_PUBLIC_PAGE_FALLBACK",
            "status": _last_receipt["status"],
            "detail": f"{type(exc).__name__}: {str(exc)[:160]}",
        }
    fixtures = parse_day_board(body)
    _last_receipt.update(
        {
            "status": "success",
            "day_url": url,
            "day_rows_seen": len(fixtures),
            "day_response_sha256": hashlib.sha256(body.encode("utf-8", "replace")).hexdigest(),
        }
    )
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": str(league).upper(),
        "verification": "DIRECT_PUBLIC_PAGE_FALLBACK",
        "date": date,
        "daily_board_rows": len(fixtures),
        "request_cap": MAX_REQUESTS_PER_RUN,
        "minimum_request_interval_seconds": MIN_REQUEST_INTERVAL_SECONDS,
        "receipt": last_receipt(),
    }


__all__ = [
    "BASE_URL",
    "MAX_REQUESTS_PER_RUN",
    "MIN_REQUEST_INTERVAL_SECONDS",
    "PROVIDER",
    "fetch_odds",
    "is_enabled",
    "last_gaps",
    "last_quota",
    "last_receipt",
    "parse_day_board",
    "parse_named_moneyline_rows",
    "smoke",
]
