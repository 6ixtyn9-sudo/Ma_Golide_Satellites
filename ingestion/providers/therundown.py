"""TheRundown bookmaker adapter — NBA full-game moneyline/total only (API v2).

Evidence status: DOCUMENTATION_ONLY (docs.therundown.io / free-tier terms)
until an authenticated smoke test occurs. Every coverage claim produced by
this module is DOCUMENTATION_ONLY until then. Documented v2 contract used
here:

- ``GET https://therundown.io/api/v2/sports`` resolves the NBA sport id from
  the provider's own reference data. There is deliberately NO hardcoded
  sport id: the id is resolved on every call, and a reference payload
  without an unambiguous NBA entry raises instead of guessing.
- ``GET https://therundown.io/api/v2/sports/{sport_id}/events/{YYYY-MM-DD}``
  with ``market_ids=1,2,3`` (the full-game markets: moneyline, spread,
  total) and ``affiliate_ids=...`` (documented book ids: DraftKings 19,
  BetMGM 22, FanDuel 23; default 23, configurable via
  ``THERUNDOWN_AFFILIATE_ID``). Event dates are ``YYYY-MM-DD``.
- Authentication: the API key (``THERUNDOWN_API_KEY``) travels ONLY in the
  ``X-TheRundown-Key`` request header (see :func:`auth_headers`); it is
  never placed in a URL, and ``http.redact`` still scrubs it from any error
  text upstream. The production fetch factory in ``pipeline.py`` attaches
  the header via ``HttpClient(default_headers=...)``.
- Free terms: 20,000 datapoints/UTC day, 200,000/UTC month, 1 req/sec,
  five-minute-delayed pre-match data; event cost ``2 * events + prices``
  (read the ``X-Datapoints`` headers). The Free plan excludes quarters,
  halves, alternates, team totals, live odds and end-user display.
- v2 payload traversal: ``events → markets → participants → lines →
  prices``. Prices arrive in American format and are converted with
  ``books.american_to_decimal``. The total ``line`` is the LINE VALUE
  (e.g. 232.0), never a price; a total without a line value yields
  ``line=None`` and is NOT usable for sheet writes (smoke reports it under
  ``not_usable``).

QUARTER PROHIBITION (hard rule): the Free plan has no quarter markets, and
``therundown`` is listed in ``books.PROVIDERS_BLOCKED_FROM_QUARTERS``. This
adapter contains NO code path that can produce a ``Period.Q1..Q4`` snapshot:
only the requested full-game market ids ``"1"``/``"3"`` are parsed (the
spread ``"2"`` is requested for completeness but deliberately unmapped),
every other market id and every line carrying a non-full-game period is
skipped with a coverage gap, the snapshot factory raises on any
non-full-game period, and the emitted list is re-asserted before returning.

FetchFn convention (shared with the pipeline's auth-aware client):
``fetch_factory(provider)`` returns a ``FetchFn`` (``Callable[[str], Any]``)
that returns parsed JSON for a URL. Optionally the FetchFn may return an
envelope dict ``{"json": <payload>, "headers": <response headers>}`` — when
it does, this module unwraps it and feeds ``headers`` through
:func:`parse_quota` so smoke() can report the ``X-Datapoints`` quota. A
provider payload would never legitimately contain both keys; the convention
is reserved for the client layer.
"""

from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

from .. import identity
from ..books import american_to_decimal
from ..models import BookSnapshot, CoverageGap, CoverageState, Market, Period, payload_hash
from ._common import bound_client_quota

PROVIDER = "therundown"
ENV_KEY = "THERUNDOWN_API_KEY"
BASE_URL = "https://therundown.io/api/v2"
AUTH_HEADER = "X-TheRundown-Key"

# Documented full-game market ids (also the requested market_ids parameter).
MARKET_MONEYLINE = "1"
MARKET_SPREAD = "2"  # requested for completeness; NOT mapped (pipeline needs ML/total only)
MARKET_TOTAL = "3"
FULL_GAME_MARKET_IDS = (MARKET_MONEYLINE, MARKET_SPREAD, MARKET_TOTAL)
MARKET_IDS_PARAM = "1,2,3"

# Market-name fallbacks when a market object carries no numeric id.
MARKET_NAME_IDS: Dict[str, str] = {
    "moneyline": MARKET_MONEYLINE,
    "ml": MARKET_MONEYLINE,
    "spread": MARKET_SPREAD,
    "total": MARKET_TOTAL,
    "totals": MARKET_TOTAL,
    "over_under": MARKET_TOTAL,
}

# Documented affiliate (book) ids on the free plan.
AFFILIATE_NAMES: Dict[int, str] = {19: "DraftKings", 22: "BetMGM", 23: "FanDuel"}
DEFAULT_AFFILIATE_ID = 23  # FanDuel

# Documented "off the board" sentinel — treated as a missing value.
OFF_THE_BOARD = 0.0001

# Documented event phases. Anything NOT equal to "pre-game" is not pre-match.
PREGAME_PHASE = "pre-game"

# Line-period values that mean full game. Any other value ("Q1", 4, "half",
# ...) is FORBIDDEN and skipped with a gap — there is no quarter code path.
FULL_GAME_PERIOD_VALUES = ("1", "full-game", "full_game", "fullgame", "main", "game")

DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"

_last_gaps: List[CoverageGap] = []
_last_quota: Dict[str, Any] = {}


def is_enabled() -> bool:
    """True only when THERUNDOWN_API_KEY is configured (never raises)."""
    return bool(os.environ.get(ENV_KEY, "").strip())


def auth_headers() -> Dict[str, str]:
    """Documented v2 auth: the key travels ONLY in the ``X-TheRundown-Key`` header.

    It is never placed in a URL; the tests prove the secret appears in no
    requested URL and reaches exactly this header. The production fetch
    factory (``pipeline._default_fetch_factory``) passes this mapping to
    ``HttpClient(default_headers=...)``.
    """
    return {AUTH_HEADER: os.environ.get(ENV_KEY, "").strip()}


def last_gaps() -> List[CoverageGap]:
    """Coverage gaps recorded by the most recent fetch_odds/smoke call."""
    return list(_last_gaps)


def last_quota() -> Dict[str, Any]:
    return dict(_last_quota)


def _reset_gaps() -> None:
    global _last_gaps, _last_quota
    _last_gaps = []
    _last_quota = {}


def _record_gap(
    game_id: Optional[str],
    field: str,
    detail: str,
    state: CoverageState = CoverageState.MISSING,
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


# ---------------------------------------------------------------------------
# Quota / fetch conventions / URLs
# ---------------------------------------------------------------------------


def parse_quota(headers: Dict[str, str]) -> Dict[str, Any]:
    """Normalise TheRundown ``X-Datapoints*`` response headers into a quota dict.

    Header names are matched case-insensitively; numeric values are converted
    to ``int`` (unparseable values pass through as strings). Empty/absent
    headers yield an empty dict.
    """
    out: Dict[str, Any] = {}
    for key, value in (headers or {}).items():
        name = str(key).strip().lower()
        if not name.startswith("x-datapoints"):
            continue
        text = str(value).strip()
        try:
            out[name] = int(text)
        except ValueError:
            out[name] = text
    return out


def _unwrap_fetch(response: Any) -> Tuple[Any, Dict[str, str]]:
    """Apply the fetch envelope convention (see module docstring).

    A FetchFn returns parsed JSON; alternatively ``{"json": ..., "headers":
    ...}`` unwraps to (payload, headers). Anything else is returned as
    (payload, {}).
    """
    if isinstance(response, dict) and "json" in response and isinstance(response.get("headers"), dict):
        return response["json"], response["headers"]
    return response, {}


def _call(fetch: Callable[[str], Any], url: str) -> Tuple[Any, Dict[str, str]]:
    payload, headers = _unwrap_fetch(fetch(url))
    headers = headers or bound_client_quota(fetch)
    if headers:
        _last_quota.update(parse_quota(headers))
    return payload, headers


def _require_supported_league(league: str) -> str:
    """NBA only (case-insensitive). EuroLeague coverage is unproven and is
    never fabricated: raise instead of returning empty data."""
    normalized = str(league or "").strip().upper()
    if normalized == "EUROLEAGUE":
        raise ValueError("not enabled for this league yet")
    if normalized != "NBA":
        raise ValueError(f"league {league!r} is not enabled for the {PROVIDER} adapter (NBA only)")
    return normalized


def _affiliate_id() -> int:
    """Configured book (affiliate id); default 23 = FanDuel."""
    raw = os.environ.get("THERUNDOWN_AFFILIATE_ID", "").strip()
    try:
        return int(raw) if raw else DEFAULT_AFFILIATE_ID
    except ValueError:
        return DEFAULT_AFFILIATE_ID


def normalize_date(date_yyyymmdd: str) -> str:
    """Accept ``YYYYMMDD`` or ``YYYY-MM-DD``; return the documented ``YYYY-MM-DD``."""
    compact = str(date_yyyymmdd).strip().replace("-", "")
    if len(compact) == 8 and compact.isdigit():
        return f"{compact[0:4]}-{compact[4:6]}-{compact[6:8]}"
    raise ValueError(f"unrecognised slate date {date_yyyymmdd!r} (expected YYYYMMDD or YYYY-MM-DD)")


def _sports_url() -> str:
    """Documented v2 sports reference endpoint (NBA id resolution)."""
    return f"{BASE_URL}/sports"


def _events_url(sport_id: Any, date_iso: str, affiliate_id: int) -> str:
    """Documented v2 events URL. The key NEVER appears here (header auth).

    Query parameters are the documented full-game ``market_ids=1,2,3`` and
    ``affiliate_ids=...``. Commas inside query values are legal and are left
    unencoded to match the documented parameter shape.
    """
    return (
        f"{BASE_URL}/sports/{sport_id}/events/{date_iso}"
        f"?market_ids={MARKET_IDS_PARAM}&affiliate_ids={affiliate_id}"
    )


def _resolve_nba_sport_id(fetch: Callable[[str], Any]) -> Any:
    """Resolve the NBA sport id from the provider's own /sports reference data.

    There is deliberately NO hardcoded fallback id: a reference payload
    without an unambiguous NBA entry raises (with a recorded gap) instead of
    guessing. The id is returned exactly as the provider spells it.
    """
    payload, _headers = _call(fetch, _sports_url())
    entries: List[Any] = []
    if isinstance(payload, dict):
        entries = payload.get("sports") or payload.get("data") or []
    elif isinstance(payload, list):
        entries = payload
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        names = [entry.get(key) for key in ("abbreviation", "sport_name", "league", "name", "slug")]
        if not any(str(name).strip().upper() == "NBA" for name in names if name is not None):
            continue
        sport_id = entry.get("id", entry.get("sport_id"))
        if sport_id is not None and str(sport_id).strip() != "":
            return sport_id
        break  # NBA entry present but id-less: do not keep searching
    _record_gap(
        None,
        "sport_id",
        "NBA sport id not found in /sports reference data; refusing to guess or hardcode",
        CoverageState.PROVIDER_ERROR,
    )
    raise ValueError("NBA sport id not found in /sports; refusing to hardcode or guess")


# ---------------------------------------------------------------------------
# Parsing (full-game only — no quarter code path exists)
# ---------------------------------------------------------------------------


def _number(value: Any) -> Optional[float]:
    """Parse a numeric payload value; the OTB sentinel (0.0001) is missing."""
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if abs(number - OFF_THE_BOARD) < 1e-9:
        return None
    return number


def _american(value: Any) -> Optional[float]:
    """Clean an American price (0 and the OTB sentinel mean missing)."""
    number = _number(value)
    if number is None or number == 0:
        return None
    return number


def _line_period_is_full_game(line: Dict[str, Any]) -> bool:
    """True only when every period-ish field on the line says FULL GAME.

    Fail-closed: a line carrying ``period_id=4``, ``period="Q1"`` or any
    other non-full-game marker is rejected (never mapped) — the adapter has
    no quarter code path. Lines with no period marker at all are accepted;
    the market-id gate (full-game market_ids only) is the primary guard.
    """
    for key in ("period_id", "period", "period_name"):
        value = line.get(key)
        if value is None:
            continue
        if str(value).strip().lower() not in FULL_GAME_PERIOD_VALUES:
            return False
    return True


def _participant_side(participant: Dict[str, Any]) -> Optional[str]:
    """Map a v2 market participant onto home/away/over/under (never guessed).

    Explicit side-ish fields win, then the participant name, then the
    ``is_home``/``is_away`` flags. Returns None when nothing identifies the
    side — the caller skips rather than guessing.
    """
    for key in ("side", "role", "type", "home_away"):
        value = participant.get(key)
        if value is None:
            continue
        text = str(value).strip().lower()
        if text in ("home", "away", "over", "under"):
            return text
        if text in ("home_team", "away_team", "over_line", "under_line"):
            return text[:-5]
    for key in ("name", "participant", "label"):
        value = participant.get(key)
        if value is not None and str(value).strip().lower() in ("home", "away", "over", "under"):
            return str(value).strip().lower()
    if participant.get("is_home") is True:
        return "home"
    if participant.get("is_away") is True or participant.get("is_home") is False:
        return "away"
    return None


def _participants(market: Dict[str, Any]) -> List[Any]:
    participants = market.get("participants")
    return participants if isinstance(participants, list) else []


def _lines_of(participant: Dict[str, Any]) -> List[Any]:
    lines = participant.get("lines")
    return lines if isinstance(lines, list) else []


def _line_prices(line: Dict[str, Any]) -> List[Any]:
    prices = line.get("prices")
    if isinstance(prices, list):
        return prices
    if prices is not None:
        return [prices]
    if line.get("price") is not None:
        return [line["price"]]
    return []


def _price_value(price: Any) -> Optional[float]:
    """American price from a raw price entry (dict with price/value/american, or number)."""
    if isinstance(price, dict):
        price = price.get("price", price.get("value", price.get("american")))
    return _american(price)


def _typed_price(line: Dict[str, Any], wanted: str) -> Optional[float]:
    """First usable price in the line typed as ``wanted`` (home/away/over/under)."""
    for raw in _line_prices(line):
        if not isinstance(raw, dict):
            continue
        typed = raw.get("type", raw.get("name", raw.get("side")))
        if typed is not None and str(typed).strip().lower() == wanted:
            value = _price_value(raw)
            if value is not None:
                return value
    return None


def _line_affiliate(line: Dict[str, Any]) -> Optional[int]:
    raw = line.get("affiliate_id", line.get("affiliate"))
    if raw is None:
        return None
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def _book_name(affiliate_id: Optional[int], requested: int) -> str:
    """Book name from the documented affiliate map; unknown ids stay explicit."""
    if affiliate_id is None:
        affiliate_id = requested
    return AFFILIATE_NAMES.get(int(affiliate_id), f"affiliate_{affiliate_id}")


def _phase_is_pregame(event: Dict[str, Any]) -> bool:
    """True only when the payload explicitly says ``pre-game``.

    Defensive fail-closed: a missing/unreadable phase is NOT pre-match, so a
    stale or live event can never be written into a pre-tip cell.
    """
    phase = str(event.get("event_status_phase") or "").strip().lower()
    return phase == PREGAME_PHASE


def _extract_teams(event: Dict[str, Any]) -> Optional[Tuple[str, str]]:
    """(home, away) names from the event-level ``participants`` (or ``teams``).

    Returns None when the two sides cannot be resolved unambiguously — the
    caller skips the event rather than guessing orientation.
    """
    sides = {"home": None, "away": None}
    roster = event.get("participants") if isinstance(event.get("participants"), list) else event.get("teams")
    if not isinstance(roster, list) or len(roster) != 2:
        return None
    for team in roster:
        if not isinstance(team, dict):
            return None
        name = team.get("full_name") or team.get("nickname") or team.get("name")
        if not name:
            return None
        side = _participant_side(team)
        if side not in sides or sides[side] is not None:  # unknown side or duplicate
            return None
        sides[side] = str(name)
    if sides["home"] is None or sides["away"] is None:
        return None
    return sides["home"], sides["away"]


def _snapshot(
    game_id: str,
    book: str,
    market: Market,
    period: Period,
    observed_utc: str,
    pre_match: bool,
    event: Dict[str, Any],
    source_modified_utc: Optional[str],
    **price_fields: Any,
) -> BookSnapshot:
    """Build a BookSnapshot with the hard full-game guard.

    Raises ``ValueError`` if a non-full-game period is ever requested: the
    Free plan has no quarter markets and this adapter must never map one.
    """
    if period is not Period.FULL_GAME:
        raise ValueError(
            f"therundown adapter attempted to build a {period.value} snapshot; "
            "the Free plan has no quarter markets and this adapter is full-game only"
        )
    return BookSnapshot(
        game_id=game_id,
        provider=PROVIDER,
        book=book,
        market=market,
        period=period,
        observed_utc=observed_utc,
        source_modified_utc=source_modified_utc,
        pre_match=pre_match,
        # Only documented full-game main markets (market_ids=1,2,3) are
        # requested and the free plan exposes main lines only.
        main_line=True,
        payload_hash=payload_hash(event),
        **price_fields,
    )


def _assert_full_game_only(snapshots: List[BookSnapshot]) -> None:
    """Runtime assertion: every built snapshot is Period.FULL_GAME."""
    for snapshot in snapshots:
        if snapshot.period is not Period.FULL_GAME:
            raise ValueError(
                f"therundown adapter produced a {snapshot.period.value} snapshot; "
                "TheRundown Free has no quarter markets and is blocked from quarters"
            )


def _moneyline_from_market(
    market: Dict[str, Any],
    game_id: str,
    requested_affiliate: int,
    observed_utc: str,
    pre_match: bool,
    event: Dict[str, Any],
) -> Optional[BookSnapshot]:
    """Full-game moneyline from market id "1" (participants → lines → prices).

    Sides are resolved per participant (home/away) or from typed prices; a
    side matching neither is skipped with a gap, never guessed.
    """
    prices: Dict[str, float] = {}
    book: Optional[str] = None
    modified: Optional[str] = None
    for participant in _participants(market):
        if not isinstance(participant, dict):
            _record_gap(game_id, "participant", "non-dict moneyline participant skipped", CoverageState.PROVIDER_ERROR)
            continue
        side = _participant_side(participant)
        for line in _lines_of(participant):
            if not isinstance(line, dict):
                _record_gap(game_id, "lines", "non-dict moneyline line skipped", CoverageState.PROVIDER_ERROR)
                continue
            if not _line_period_is_full_game(line):
                _record_gap(
                    game_id,
                    "lines",
                    f"line period {line.get('period_id', line.get('period'))!r} is not full-game; "
                    "skipped (this adapter has no quarter code path)",
                    CoverageState.UNSUPPORTED,
                )
                continue
            affiliate = _line_affiliate(line)
            typed = _typed_price(line, side) if side in ("home", "away") else None
            if typed is not None:
                prices.setdefault(side, typed)
                book = book or _book_name(affiliate, requested_affiliate)
                modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
                continue
            if side in ("home", "away") and side not in prices:
                for raw in _line_prices(line):
                    value = _price_value(raw)
                    if value is not None:
                        prices[side] = value
                        book = book or _book_name(affiliate, requested_affiliate)
                        modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
                        break
            elif side is not None and side not in ("home", "away"):
                _record_gap(game_id, "participant", f"moneyline participant side {side!r} unexpected; skipped")
    if not prices:
        return None
    home = prices.get("home")
    away = prices.get("away")
    return _snapshot(
        game_id,
        book or _book_name(None, requested_affiliate),
        Market.MONEYLINE,
        Period.FULL_GAME,
        observed_utc,
        pre_match,
        event,
        modified,
        home_price_decimal=american_to_decimal(home) if home is not None else None,
        away_price_decimal=american_to_decimal(away) if away is not None else None,
    )


def _total_from_market(
    market: Dict[str, Any],
    game_id: str,
    requested_affiliate: int,
    observed_utc: str,
    pre_match: bool,
    event: Dict[str, Any],
) -> Optional[BookSnapshot]:
    """Full-game total from market id "3" (participants → lines → prices).

    The LINE VALUE comes from the line's ``total`` field (never a price);
    over/under prices from typed prices or per-side participants.
    """
    over = under = None
    line_value: Optional[float] = None
    book: Optional[str] = None
    modified: Optional[str] = None
    for participant in _participants(market):
        if not isinstance(participant, dict):
            _record_gap(game_id, "participant", "non-dict total participant skipped", CoverageState.PROVIDER_ERROR)
            continue
        side = _participant_side(participant)
        for line in _lines_of(participant):
            if not isinstance(line, dict):
                _record_gap(game_id, "lines", "non-dict total line skipped", CoverageState.PROVIDER_ERROR)
                continue
            if not _line_period_is_full_game(line):
                _record_gap(
                    game_id,
                    "lines",
                    f"line period {line.get('period_id', line.get('period'))!r} is not full-game; "
                    "skipped (this adapter has no quarter code path)",
                    CoverageState.UNSUPPORTED,
                )
                continue
            affiliate = _line_affiliate(line)
            if line_value is None:
                candidate = _number(line.get("total"))
                if candidate is None:
                    candidate = _number(line.get("line"))
                if candidate is not None:
                    line_value = candidate
            typed_over = _typed_price(line, "over")
            typed_under = _typed_price(line, "under")
            if typed_over is not None and over is None:
                over = typed_over
                book = book or _book_name(affiliate, requested_affiliate)
                modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
            if typed_under is not None and under is None:
                under = typed_under
                book = book or _book_name(affiliate, requested_affiliate)
                modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
            if side in ("over", "under"):
                own = _typed_price(line, side)
                if own is None:
                    for raw in _line_prices(line):
                        value = _price_value(raw)
                        if value is not None:
                            own = value
                            break
                if own is not None:
                    if side == "over" and over is None:
                        over = own
                        book = book or _book_name(affiliate, requested_affiliate)
                        modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
                    if side == "under" and under is None:
                        under = own
                        book = book or _book_name(affiliate, requested_affiliate)
                        modified = modified or (str(line.get("last_modified")) if line.get("last_modified") else None)
            elif side is not None and side not in ("over", "under", "home", "away"):
                _record_gap(game_id, "participant", f"total participant side {side!r} unexpected; skipped")
    if line_value is None and over is None and under is None:
        _record_gap(game_id, "total", "no full-game total fields under market '3'")
        return None
    return _snapshot(
        game_id,
        book or _book_name(None, requested_affiliate),
        Market.TOTAL,
        Period.FULL_GAME,
        observed_utc,
        pre_match,
        event,
        modified,
        line=line_value,
        over_price_decimal=american_to_decimal(over) if over is not None else None,
        under_price_decimal=american_to_decimal(under) if under is not None else None,
    )


def _market_id(market: Dict[str, Any]) -> str:
    """Market id string, or the documented id derived from the market name."""
    raw = market.get("id", market.get("market_id"))
    if raw is not None and str(raw).strip() != "":
        return str(raw).strip()
    name = str(market.get("name") or market.get("type") or "").strip().lower()
    return MARKET_NAME_IDS.get(name, "")


def _parse_event(
    event: Dict[str, Any],
    observed_utc: str,
    requested_affiliate: int,
) -> List[BookSnapshot]:
    """One v2 event → full-game moneyline/total snapshots.

    Traversal: ``markets → participants → lines → prices``. Only market ids
    "1" (moneyline) and "3" (total) are mapped; "2" (spread) and every other
    id — including any quarter market — is skipped with an explicit gap.
    """
    event_id = event.get("event_id", event.get("id"))
    if event_id is None or str(event_id).strip() == "":
        _record_gap(None, "event_id", "event without event_id skipped", CoverageState.PROVIDER_ERROR)
        return []
    game_id = str(event_id)
    teams = _extract_teams(event)
    if teams is None:
        _record_gap(game_id, "teams", "event without resolvable home/away teams skipped", CoverageState.PROVIDER_ERROR)
        return []
    pre_match = _phase_is_pregame(event)
    markets = event.get("markets")
    if not isinstance(markets, list) or not markets:
        _record_gap(game_id, "markets", "event without markets array skipped")
        return []

    snapshots: List[BookSnapshot] = []
    for market in markets:
        if not isinstance(market, dict):
            _record_gap(game_id, "markets", "non-dict market skipped", CoverageState.PROVIDER_ERROR)
            continue
        market_id = _market_id(market)
        if market_id not in FULL_GAME_MARKET_IDS:
            _record_gap(
                game_id,
                "markets",
                f"market id {market_id or '<none>'!r} is not a requested full-game market "
                f"(market_ids={MARKET_IDS_PARAM}); skipped (this adapter has no quarter code path)",
                CoverageState.UNSUPPORTED,
            )
            continue
        if market_id == MARKET_SPREAD:
            _record_gap(
                game_id,
                "markets",
                "full-game spread (market '2') requested but not mapped; "
                "the pipeline needs moneyline/total only",
                CoverageState.UNSUPPORTED,
            )
            continue
        builder = _moneyline_from_market if market_id == MARKET_MONEYLINE else _total_from_market
        snapshot = builder(market, game_id, requested_affiliate, observed_utc, pre_match, event)
        if snapshot is not None:
            snapshots.append(snapshot)

    scheduled = str(event.get("event_date") or "").strip()
    try:
        # A date-only value is deliberately insufficient for a 30-minute
        # cross-provider identity join.
        identity.parse_utc(scheduled)
    except (TypeError, ValueError):
        scheduled = ""
    for snapshot in snapshots:
        snapshot.provider_event_id = game_id
        snapshot.home_name = teams[0]
        snapshot.away_name = teams[1]
        snapshot.scheduled_utc = scheduled or None
    return snapshots


def parse_events(
    payload: Any,
    observed_utc: str,
    affiliate_id: int = DEFAULT_AFFILIATE_ID,
) -> List[BookSnapshot]:
    """Normalise a therundown v2 events payload into BookSnapshots.

    ``book`` comes from the lines' affiliate id (documented map 19/22/23;
    unknown ids fall back to ``affiliate_<id>``). Malformed events are
    skipped with gaps; the returned list is asserted FULL_GAME-only.
    """
    events: List[Any] = []
    if isinstance(payload, dict):
        events = payload.get("events") or []
    elif isinstance(payload, list):
        events = payload
    snapshots: List[BookSnapshot] = []
    for event in events:
        if not isinstance(event, dict):
            _record_gap(None, "event", "non-dict event skipped", CoverageState.PROVIDER_ERROR)
            continue
        snapshots.extend(_parse_event(event, observed_utc, affiliate_id))
    _assert_full_game_only(snapshots)
    return snapshots


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def fetch_odds(
    league: str,
    date_yyyymmdd: str,
    fetch_factory: Callable[[str], Callable[[str], Any]],
    now_utc: Optional[str] = None,
) -> List[BookSnapshot]:
    """Fetch normalised pre-match full-game odds for one league/UTC date.

    Two documented calls: ``GET /api/v2/sports`` (NBA id resolution) and
    ``GET /api/v2/sports/{id}/events/{YYYY-MM-DD}?market_ids=1,2,3&affiliate_ids=...``.
    Returns ``[]`` when the provider is disabled (missing key) — the keyless
    pipeline must not fail. Raises ``ValueError`` for unsupported leagues
    (EuroLeague coverage is unproven) and when NBA cannot be resolved from
    the provider's reference data.
    """
    _reset_gaps()
    if not is_enabled():
        return []
    _require_supported_league(league)
    fetch = fetch_factory(PROVIDER)
    sport_id = _resolve_nba_sport_id(fetch)
    affiliate_id = _affiliate_id()
    date_iso = normalize_date(date_yyyymmdd)
    payload, _headers = _call(fetch, _events_url(sport_id, date_iso, affiliate_id))
    observed_utc = now_utc or identity.now_utc_iso()
    return parse_events(payload, observed_utc, affiliate_id)


def _count_prices(snapshots: List[BookSnapshot]) -> int:
    """Price values seen (documented cost formula counts prices, not lines)."""
    count = 0
    for s in snapshots:
        count += sum(1 for v in (s.home_price_decimal, s.away_price_decimal) if v is not None)
        count += sum(1 for v in (s.over_price_decimal, s.under_price_decimal) if v is not None)
    return count


def _period_presence(snapshots: List[BookSnapshot]) -> Dict[str, set]:
    """Per-period set of game_ids with a usable (line set, pre-match) total."""
    presence: Dict[str, set] = {}
    for s in snapshots:
        if s.market is not Market.TOTAL or s.line is None or not s.pre_match:
            continue
        presence.setdefault(s.period.value, set()).add(s.game_id)
    return presence


def smoke(
    league: str,
    fetch_factory: Callable[[str], Callable[[str], Any]],
    date_yyyymmdd: Optional[str] = None,
    now_utc: Optional[str] = None,
) -> Dict[str, Any]:
    """Coverage smoke for one league (DOCUMENTATION_ONLY until live-verified).

    Reports the resolved NBA sport id, the documented quarter-coverage shape
    (for TheRundown every quarter ``present`` is structurally 0 — no quarter
    code path exists), and the ``X-Datapoints`` quota from the fetch
    envelope convention. ``date`` defaults to today's UTC date.
    """
    _reset_gaps()
    if not is_enabled():
        return {
            "enabled": False,
            "provider": PROVIDER,
            "league": str(league),
            "verification": DOCUMENTATION_ONLY,
        }
    league_name = _require_supported_league(league)
    date_iso = normalize_date(date_yyyymmdd) if date_yyyymmdd else datetime.now(timezone.utc).strftime("%Y-%m-%d")
    fetch = fetch_factory(PROVIDER)
    sport_id = _resolve_nba_sport_id(fetch)
    affiliate_id = _affiliate_id()
    payload, headers = _call(fetch, _events_url(sport_id, date_iso, affiliate_id))
    observed_utc = now_utc or identity.now_utc_iso()
    snapshots = parse_events(payload, observed_utc, affiliate_id)

    events: List[Any] = payload.get("events") if isinstance(payload, dict) else payload
    events = events if isinstance(events, list) else []
    events_checked = len(events)
    pre_match_count = sum(1 for e in events if isinstance(e, dict) and _phase_is_pregame(e))
    presence = _period_presence(snapshots)
    not_usable = sum(
        1 for s in snapshots if s.market is Market.TOTAL and s.line is None
    )
    prices = _count_prices(snapshots)
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": league_name,
        "date": date_iso,
        "affiliate_id": affiliate_id,
        "sport_id": sport_id,
        "verification": DOCUMENTATION_ONLY,
        "events_checked": events_checked,
        "full_game_totals_present": len(presence.get(Period.FULL_GAME.value, set())),
        "quarters": {
            f"Q{i}": {"present": len(presence.get(f"Q{i}", set())), "total": events_checked}
            for i in range(1, 5)
        },
        "pre_match_vs_live": {
            "pre_match": pre_match_count,
            "not_pre_match": events_checked - pre_match_count,
        },
        "not_usable": not_usable,
        "prices_seen": prices,
        # Documented free-tier cost formula: 2 * events + prices.
        "expected_datapoints_cost": 2 * events_checked + prices,
        "quota": last_quota(),
        "gaps": len(_last_gaps),
    }
