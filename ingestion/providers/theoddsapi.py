"""The Odds API adapter — targeted full-game and Q1–Q4 total fallback.

Evidence status: DOCUMENTATION_ONLY (the-odds-api.com
/sports-odds-data/betting-markets.html and /liveapi/guides/v4/) until an
authenticated smoke test occurs; every coverage claim is DOCUMENTATION_ONLY
until then. Documented contract used here:

- ``GET https://api.the-odds-api.com/v4/sports/{sport_key}/odds/?apiKey=...&regions=us&markets=h2h,totals&bookmakers=draftkings&oddsFormat=decimal``
  for full-game moneyline/total discovery, plus a per-event call
  ``GET /v4/sports/{sport_key}/events/{event_id}/odds?apiKey=...&markets=totals_q1,totals_q2,totals_q3,totals_q4&...``
  for the period markets (quarters are ONLY served per event). Documented
  credit accounting is ~1 credit per market per region, so one event with
  h2h+totals plus four quarter markets ≈ 1 + 4 credits — verify at
  enablement. The key travels as the documented ``apiKey`` query parameter
  (from ``THE_ODDS_API_KEY``); ``http.redact`` scrubs it from any error
  text and this module never logs URLs.
- ``NBA_SPORT_KEY = "basketball_nba"`` is a constant that MUST be verified
  against GET /sports at enablement. NO EuroLeague sport key is assumed to
  exist — EuroLeague raises ``ValueError("not enabled for this league
  yet")`` (coverage unproven, never fabricated).
- ``oddsFormat=decimal`` is requested explicitly, so prices are ALREADY
  decimal and are stored AS-IS — ``books.american_to_decimal`` must NOT be
  applied (it would corrupt decimal prices; American conversion belongs to
  the TheRundown adapter).
- Market keys: ``h2h`` → MONEYLINE (home/away resolved by matching the
  outcome name to the event's ``home_team``/``away_team`` — never by array
  order; an outcome matching neither side is skipped with a gap, never
  guessed), ``totals`` → FULL_GAME TOTAL, ``totals_q1``..``totals_q4`` →
  exactly Q1..Q4. The total ``line`` is the outcome ``point`` (e.g. 232.5),
  never a price; a total without a point yields ``line=None`` and is NOT
  usable for total forecasts (smoke reports it under ``not_usable``).
- Quota: ``x-requests-remaining``/``x-requests-used`` headers via
  :func:`parse_quota`.
- Date narrowing is applied client-side on ``commence_time`` (UTC calendar
  date); events that already started keep ``pre_match=False`` (never marked
  pre-match) and are excluded downstream by the pre-match filters.

FetchFn convention: ``fetch_factory(provider)`` returns a ``FetchFn``
(``Callable[[str], Any]``) returning parsed JSON; an envelope dict
``{"json": ..., "headers": ...}`` is unwrapped so response headers can be
reported (see therundown module docstring).
"""

from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple
from urllib.parse import urlencode

from .. import identity
from ..models import BookSnapshot, CoverageGap, CoverageState, Market, Period, payload_hash
from ._common import bound_client_quota

PROVIDER = "theoddsapi"
ENV_KEY = "THE_ODDS_API_KEY"
BASE_URL = "https://api.the-odds-api.com/v4"

# NBA sport key. Verify against GET /sports at enablement. There is NO
# EuroLeague key here on purpose: EuroLeague coverage is unproven and the
# adapter raises for it instead of fabricating a key.
NBA_SPORT_KEY = "basketball_nba"
SPORT_KEYS = {
    "NBA": NBA_SPORT_KEY,
    # Documented WNBA key; authenticated coverage remains unverified until a
    # credentialed smoke run succeeds. Keyless ESPN markets remain independent.
    "WNBA": "basketball_wnba",
    "NCAA_M": "basketball_ncaab",
    "NCAA_W": "basketball_wncaab",
}

MONEYLINE_MARKET_KEYS: Dict[str, Period] = {
    "h2h": Period.FULL_GAME,
    "h2h_h1": Period.H1,
    "h2h_h2": Period.H2,
}
TOTALS_MARKET_KEYS: Dict[str, Period] = {
    "totals": Period.FULL_GAME,
    "totals_q1": Period.Q1,
    "totals_q2": Period.Q2,
    "totals_q3": Period.Q3,
    "totals_q4": Period.Q4,
    "totals_h1": Period.H1,
    "totals_h2": Period.H2,
}
SPREAD_MARKET_KEYS: Dict[str, Period] = {
    "spreads": Period.FULL_GAME,
    "spreads_q1": Period.Q1,
    "spreads_q2": Period.Q2,
    "spreads_q3": Period.Q3,
    "spreads_q4": Period.Q4,
    "spreads_h1": Period.H1,
    "spreads_h2": Period.H2,
}
FULL_GAME_MARKETS_PARAM = "h2h,totals,spreads"
QUARTER_MARKETS_PARAM = (
    "totals_q1,totals_q2,totals_q3,totals_q4,"
    "spreads_q1,spreads_q2,spreads_q3,spreads_q4,h2h_h1"
)
HALF_MARKETS_PARAM = (
    "totals_h1,totals_h2,spreads_h1,spreads_h2,h2h_h1,h2h_h2"
)
DEFAULT_REGIONS = "us"
DEFAULT_BOOKMAKERS = "draftkings"

DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"

_last_gaps: List[CoverageGap] = []
_last_quota: Dict[str, Any] = {}


def is_enabled() -> bool:
    """True only when THE_ODDS_API_KEY is configured (never raises)."""
    return bool(os.environ.get(ENV_KEY, "").strip())


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
# Quota / fetch conventions / league gate / URLs
# ---------------------------------------------------------------------------


def parse_quota(headers: Dict[str, str]) -> Dict[str, Any]:
    """Normalise ``x-requests-remaining``/``x-requests-used`` (and friends)
    response headers into a quota dict. Header names are matched
    case-insensitively; numeric values become ints."""
    out: Dict[str, Any] = {}
    for key, value in (headers or {}).items():
        name = str(key).strip().lower()
        if not name.startswith("x-requests"):
            continue
        text = str(value).strip()
        try:
            out[name] = int(text)
        except ValueError:
            out[name] = text
    return out


def _unwrap_fetch(response: Any) -> Tuple[Any, Dict[str, str]]:
    """Fetch envelope convention: ``{"json": ..., "headers": ...}`` unwraps;
    plain parsed JSON returns as (payload, {})."""
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
    """Return a verified The Odds API sport mapping, never a guessed key."""
    normalized = str(league or "").strip().upper().replace("-", "_")
    aliases = {"NCAAB": "NCAA_M", "WNCAAB": "NCAA_W"}
    normalized = aliases.get(normalized, normalized)
    if normalized == "EUROLEAGUE":
        raise ValueError("not enabled for this league yet")
    if normalized not in SPORT_KEYS:
        raise ValueError(
            f"league {league!r} is not enabled for the {PROVIDER} adapter; "
            f"verified mappings: {', '.join(SPORT_KEYS)}"
        )
    return normalized


def _regions() -> str:
    return os.environ.get("THE_ODDS_API_REGIONS", "").strip() or DEFAULT_REGIONS


def _bookmakers() -> str:
    return os.environ.get("THE_ODDS_API_BOOKMAKERS", "").strip() or DEFAULT_BOOKMAKERS


def _api_key() -> str:
    return os.environ.get(ENV_KEY, "")


def _full_game_odds_url(league: str = "NBA") -> str:
    """Documented list endpoint: full-game h2h + totals, decimal odds."""
    league = _require_supported_league(league)
    params = {
        "apiKey": _api_key(),
        "regions": _regions(),
        "markets": FULL_GAME_MARKETS_PARAM,
        "bookmakers": _bookmakers(),
        "oddsFormat": "decimal",
        "dateFormat": "iso",
    }
    return f"{BASE_URL}/sports/{SPORT_KEYS[league]}/odds/?{urlencode(params)}"


def _event_odds_url(event_id: str, league: str = "NBA") -> str:
    """Per-event native period markets (quarters for NBA, halves for NCAA)."""
    league = _require_supported_league(league)
    params = {
        "apiKey": _api_key(),
        "regions": _regions(),
        "markets": HALF_MARKETS_PARAM if league.startswith("NCAA_") else QUARTER_MARKETS_PARAM,
        "bookmakers": _bookmakers(),
        "oddsFormat": "decimal",
        "dateFormat": "iso",
    }
    return f"{BASE_URL}/sports/{SPORT_KEYS[league]}/events/{event_id}/odds?{urlencode(params)}"


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def _number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _team_side(name: str, home_team: str, away_team: str) -> Optional[str]:
    """Map an h2h outcome name onto home/away by NAME (never array order).

    Exact match first, then normalised comparison; ``None`` when the name
    matches neither team — the caller skips that outcome with a gap rather
    than guessing.
    """
    if name == home_team:
        return "home"
    if name == away_team:
        return "away"
    normalized = identity.normalize_team_name(name)
    if home_team and normalized == identity.normalize_team_name(home_team):
        return "home"
    if away_team and normalized == identity.normalize_team_name(away_team):
        return "away"
    return None


def _event_pre_match(event: Dict[str, Any], now_utc: str) -> bool:
    """Pre-match only when commence_time parses AND is in the future.

    Fail closed: a missing/unparseable commence_time is NOT pre-match —
    live data must never be marked pre-match.
    """
    commence = str(event.get("commence_time") or "").strip()
    if not commence:
        _record_gap(str(event.get("id", "")), "commence_time", "missing commence_time; pre_match=False")
        return False
    try:
        return identity.parse_utc(commence) > identity.parse_utc(now_utc)
    except ValueError:
        _record_gap(
            str(event.get("id", "")),
            "commence_time",
            f"unparseable commence_time {commence!r}; pre_match=False",
            CoverageState.PROVIDER_ERROR,
        )
        return False


def _moneyline_snapshot(
    market: Dict[str, Any],
    period: Period,
    game_id: str,
    book: str,
    observed_utc: str,
    pre_match: bool,
    home_team: str,
    away_team: str,
    event: Dict[str, Any],
) -> Optional[BookSnapshot]:
    home_price: Optional[float] = None
    away_price: Optional[float] = None
    draw_price: Optional[float] = None
    for outcome in market.get("outcomes") or []:
        if not isinstance(outcome, dict):
            _record_gap(game_id, "h2h_outcome", "non-dict outcome skipped", CoverageState.PROVIDER_ERROR)
            continue
        name = str(outcome.get("name") or "")
        if name.strip().lower() in {"draw", "x"}:
            draw_price = _number(outcome.get("price"))
            continue
        side = _team_side(name, home_team, away_team)
        if side is None:
            _record_gap(
                game_id,
                "h2h_outcome",
                f"outcome {name!r} matches neither {home_team!r} nor {away_team!r}; skipped (never guessed)",
            )
            continue
        # oddsFormat=decimal: stored AS-IS — no american_to_decimal here.
        price = _number(outcome.get("price"))
        if side == "home":
            home_price = price
        else:
            away_price = price
    if home_price is None and away_price is None and draw_price is None:
        return None
    return BookSnapshot(
        game_id=game_id,
        provider=PROVIDER,
        book=book,
        market=Market.MONEYLINE,
        period=period,
        observed_utc=observed_utc,
        source_modified_utc=str(market.get("last_update")) if market.get("last_update") else None,
        pre_match=pre_match,
        # Only the documented main market keys (h2h/totals*) are requested;
        # The Odds API alternates are separate markets not requested here.
        main_line=True,
        payload_hash=payload_hash(event),
        home_price_decimal=home_price,
        away_price_decimal=away_price,
        draw_price_decimal=draw_price,
    )


def _total_snapshot(
    market: Dict[str, Any],
    period: Period,
    game_id: str,
    book: str,
    observed_utc: str,
    pre_match: bool,
    event: Dict[str, Any],
) -> Optional[BookSnapshot]:
    over: Optional[Dict[str, Any]] = None
    under: Optional[Dict[str, Any]] = None
    for outcome in market.get("outcomes") or []:
        if not isinstance(outcome, dict):
            _record_gap(game_id, f"totals_{period.value.lower()}", "non-dict outcome skipped", CoverageState.PROVIDER_ERROR)
            continue
        name = str(outcome.get("name") or "").strip().lower()
        if name == "over" and over is None:
            over = outcome
        elif name == "under" and under is None:
            under = outcome
        else:
            _record_gap(
                game_id,
                f"totals_{period.value.lower()}",
                f"unexpected outcome {str(outcome.get('name'))!r} skipped",
            )
    # The LINE VALUE is the outcome "point" (e.g. 232.5) — never a price.
    # Missing point -> line=None: NOT usable for total forecasts (totals need
    # the line); smoke reports it under "not_usable".
    line = _number((over or {}).get("point"))
    if line is None:
        line = _number((under or {}).get("point"))
    over_price = _number((over or {}).get("price"))
    under_price = _number((under or {}).get("price"))
    if line is None and over_price is None and under_price is None:
        _record_gap(game_id, f"totals_{period.value.lower()}", "no usable point/price outcomes")
        return None
    return BookSnapshot(
        game_id=game_id,
        provider=PROVIDER,
        book=book,
        market=Market.TOTAL,
        period=period,
        observed_utc=observed_utc,
        source_modified_utc=str(market.get("last_update")) if market.get("last_update") else None,
        pre_match=pre_match,
        main_line=True,
        payload_hash=payload_hash(event),
        line=line,
        over_price_decimal=over_price,  # decimal AS-IS
        under_price_decimal=under_price,  # decimal AS-IS
    )


def _spread_snapshot(
    market: Dict[str, Any],
    period: Period,
    game_id: str,
    book: str,
    observed_utc: str,
    pre_match: bool,
    home_team: str,
    away_team: str,
    event: Dict[str, Any],
) -> Optional[BookSnapshot]:
    """Normalize a two-sided spread with a home-oriented line.

    The provider's home and away points must be exact opposites when both
    are present. Contradictory lines are rejected rather than averaged.
    """
    home: Optional[Dict[str, Any]] = None
    away: Optional[Dict[str, Any]] = None
    for outcome in market.get("outcomes") or []:
        if not isinstance(outcome, dict):
            _record_gap(
                game_id, "spread_outcome", "non-dict outcome skipped",
                CoverageState.PROVIDER_ERROR,
            )
            continue
        side = _team_side(str(outcome.get("name") or ""), home_team, away_team)
        if side == "home" and home is None:
            home = outcome
        elif side == "away" and away is None:
            away = outcome
        else:
            _record_gap(
                game_id,
                f"spreads_{period.value.lower()}",
                f"unexpected outcome {str(outcome.get('name'))!r} skipped",
            )
    home_line = _number((home or {}).get("point"))
    away_line = _number((away or {}).get("point"))
    if home_line is not None and away_line is not None:
        if abs(home_line + away_line) > 1e-9:
            _record_gap(
                game_id,
                f"spread_line_{period.value.lower()}",
                "home/away spread points are not exact opposites; market rejected",
                CoverageState.PROVIDER_ERROR,
            )
            return None
    elif home_line is None and away_line is not None:
        home_line = -away_line
    home_price = _number((home or {}).get("price"))
    away_price = _number((away or {}).get("price"))
    if home_line is None and home_price is None and away_price is None:
        return None
    return BookSnapshot(
        game_id=game_id,
        provider=PROVIDER,
        book=book,
        market=Market.SPREAD,
        period=period,
        observed_utc=observed_utc,
        source_modified_utc=(
            str(market.get("last_update")) if market.get("last_update") else None
        ),
        pre_match=pre_match,
        main_line=True,
        payload_hash=payload_hash(event),
        line=home_line,
        home_price_decimal=home_price,
        away_price_decimal=away_price,
    )


def parse_event(
    event: Dict[str, Any],
    observed_utc: str,
    now_utc: Optional[str] = None,
) -> List[BookSnapshot]:
    """Normalise one documented v4 event/odds payload into BookSnapshots.

    Works for both the list endpoint's events (h2h/totals) and the
    per-event odds payload (totals_q1..q4). ``book`` is the bookmaker KEY
    (title only as fallback). Malformed entries are skipped with gaps.
    """
    if not isinstance(event, dict):
        _record_gap(None, "event", "non-dict event skipped", CoverageState.PROVIDER_ERROR)
        return []
    event_id = event.get("id", event.get("eventID", event.get("eventId")))
    if event_id is None or str(event_id).strip() == "":
        _record_gap(None, "event_id", "event without id skipped", CoverageState.PROVIDER_ERROR)
        return []
    game_id = str(event_id)
    home_team = str(event.get("home_team") or "")
    away_team = str(event.get("away_team") or "")
    now = now_utc or identity.now_utc_iso()
    pre_match = _event_pre_match(event, now)

    bookmakers = event.get("bookmakers")
    if not isinstance(bookmakers, list):
        _record_gap(game_id, "bookmakers", "no bookmakers array present")
        return []

    snapshots: List[BookSnapshot] = []
    for bookmaker in bookmakers:
        if not isinstance(bookmaker, dict):
            _record_gap(game_id, "bookmaker", "non-dict bookmaker skipped", CoverageState.PROVIDER_ERROR)
            continue
        book = str(bookmaker.get("key") or bookmaker.get("title") or "unknown")
        markets = bookmaker.get("markets")
        if not isinstance(markets, list):
            _record_gap(game_id, "markets", f"bookmaker {book!r} has no markets array")
            continue
        for market in markets:
            if not isinstance(market, dict) or not market.get("key"):
                _record_gap(game_id, "market", "non-dict/keyless market skipped", CoverageState.PROVIDER_ERROR)
                continue
            key = str(market["key"])
            if key in MONEYLINE_MARKET_KEYS:
                snapshot = _moneyline_snapshot(
                    market, MONEYLINE_MARKET_KEYS[key], game_id, book,
                    observed_utc, pre_match, home_team, away_team, event,
                )
            elif key in TOTALS_MARKET_KEYS:
                snapshot = _total_snapshot(
                    market, TOTALS_MARKET_KEYS[key], game_id, book,
                    observed_utc, pre_match, event,
                )
            elif key in SPREAD_MARKET_KEYS:
                snapshot = _spread_snapshot(
                    market, SPREAD_MARKET_KEYS[key], game_id, book,
                    observed_utc, pre_match, home_team, away_team, event,
                )
            else:
                continue  # not requested; nothing to normalise
            if snapshot is not None:
                snapshots.append(snapshot)

    scheduled = str(event.get("commence_time") or "").strip()
    try:
        identity.parse_utc(scheduled)
    except (TypeError, ValueError):
        scheduled = ""
    for snapshot in snapshots:
        snapshot.provider_event_id = game_id
        snapshot.home_name = home_team or None
        snapshot.away_name = away_team or None
        snapshot.scheduled_utc = scheduled or None
    return snapshots


def parse_events(payload: Any, observed_utc: str, now_utc: Optional[str] = None) -> List[BookSnapshot]:
    """Normalise a documented v4 list payload (array of events) into snapshots."""
    events: List[Any] = []
    if isinstance(payload, list):
        events = payload
    elif isinstance(payload, dict) and isinstance(payload.get("data"), list):
        events = payload["data"]
    snapshots: List[BookSnapshot] = []
    for event in events:
        if isinstance(event, dict):
            snapshots.extend(parse_event(event, observed_utc, now_utc))
    return snapshots


def _on_utc_date(
    commence_time: str, date_yyyymmdd: str, tz_name: Optional[str] = None
) -> bool:
    """Match commence time to the slate's local (default UTC) calendar date."""
    from zoneinfo import ZoneInfo

    try:
        local = identity.parse_utc(commence_time).astimezone(ZoneInfo(tz_name or "UTC"))
    except (TypeError, ValueError):
        return False
    return local.strftime("%Y%m%d") == str(date_yyyymmdd).replace("-", "")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def fetch_odds(
    league: str,
    date_yyyymmdd: str,
    fetch_factory: Callable[[str], Callable[[str], Any]],
    now_utc: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> List[BookSnapshot]:
    """Fetch normalised pre-match odds for one league/UTC date.

    Full-game moneyline/totals come from the documented list endpoint;
    quarter totals require one per-event call per event (documented credit
    cost ≈ 1 per market per region — verify at enablement). Returns ``[]``
    when the provider is disabled (missing key) — the keyless pipeline must
    not fail. Raises ``ValueError`` for unsupported leagues (EuroLeague has
    no verified sport key).
    """
    _reset_gaps()
    if not is_enabled():
        return []
    league = _require_supported_league(league)
    fetch = fetch_factory(PROVIDER)
    now = now_utc or identity.now_utc_iso()
    observed_utc = now
    list_payload, _headers = _call(fetch, _full_game_odds_url(league))
    events = [e for e in (list_payload if isinstance(list_payload, list) else []) if isinstance(e, dict)]

    snapshots: List[BookSnapshot] = []
    seen: set = set()  # (game_id, book, market, period) — list endpoint wins

    def _extend(new: List[BookSnapshot]) -> None:
        for snapshot in new:
            key = (snapshot.game_id, snapshot.book, snapshot.market, snapshot.period)
            if key in seen:
                continue
            seen.add(key)
            snapshots.append(snapshot)

    for event in events:
        if not event.get("id"):
            _record_gap(None, "event_id", "event without id skipped", CoverageState.PROVIDER_ERROR)
            continue
        game_id = str(event["id"])
        commence = str(event.get("commence_time") or "")
        if not _on_utc_date(commence, date_yyyymmdd, tz_name):
            _record_gap(game_id, "date_window", f"commence_time {commence} outside {date_yyyymmdd} in {tz_name or 'UTC'}; skipped")
            continue
        _extend(parse_event(event, observed_utc, now))
        # Quarter markets exist only on the per-event endpoint.
        quarter_payload, _q_headers = _call(fetch, _event_odds_url(game_id, league))
        _extend(parse_event(quarter_payload if isinstance(quarter_payload, dict) else {}, observed_utc, now))
    return snapshots


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
    tz_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Coverage smoke for one league (DOCUMENTATION_ONLY until live-verified).

    Reports per-quarter coverage separately (partial quarters are allowed
    and never backfilled), pre-match vs started counts, the NOT-usable
    (line-less) total count, and the ``x-requests-*`` quota captured from
    the response headers via the fetch envelope convention.
    ``date_yyyymmdd`` defaults to today's UTC date.
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
    date_yyyymmdd = date_yyyymmdd or datetime.now(timezone.utc).strftime("%Y%m%d")
    fetch = fetch_factory(PROVIDER)
    now = now_utc or identity.now_utc_iso()
    observed_utc = now

    quota: Dict[str, Any] = {}

    def _fetch_and_track(url: str) -> Any:
        payload, headers = _call(fetch, url)
        if headers:
            quota.update(parse_quota(headers))
        return payload

    list_payload = _fetch_and_track(_full_game_odds_url(league_name))
    events = [e for e in (list_payload if isinstance(list_payload, list) else []) if isinstance(e, dict)]
    snapshots: List[BookSnapshot] = []
    seen: set = set()  # (game_id, book, market, period) — list endpoint wins

    def _extend(new: List[BookSnapshot]) -> None:
        for snapshot in new:
            key = (snapshot.game_id, snapshot.book, snapshot.market, snapshot.period)
            if key in seen:
                continue
            seen.add(key)
            snapshots.append(snapshot)

    events_checked = 0
    pre_match_count = 0
    for event in events:
        if not event.get("id"):
            _record_gap(None, "event_id", "event without id skipped", CoverageState.PROVIDER_ERROR)
            continue
        game_id = str(event["id"])
        commence = str(event.get("commence_time") or "")
        if not _on_utc_date(commence, date_yyyymmdd, tz_name):
            _record_gap(game_id, "date_window", f"commence_time {commence} outside {date_yyyymmdd} in {tz_name or 'UTC'}; skipped")
            continue
        events_checked += 1
        if _event_pre_match(event, now):
            pre_match_count += 1
        _extend(parse_event(event, observed_utc, now))
        quarter_payload = _fetch_and_track(_event_odds_url(game_id, league_name))
        _extend(parse_event(quarter_payload if isinstance(quarter_payload, dict) else {}, observed_utc, now))

    presence = _period_presence(snapshots)
    not_usable = sum(1 for s in snapshots if s.market is Market.TOTAL and s.line is None)
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": league_name,
        "date": date_yyyymmdd,
        "sport_key": SPORT_KEYS[league_name],
        "verification": DOCUMENTATION_ONLY,
        "events_checked": events_checked,
        "full_game_totals_present": len(presence.get(Period.FULL_GAME.value, set())),
        "native_periods": {
            period: {"present": len(presence.get(period, set())), "total": events_checked}
            for period in (
                ("H1", "H2") if league_name.startswith("NCAA_") else ("Q1", "Q2", "Q3", "Q4")
            )
        },
        **({
            "halves" if league_name.startswith("NCAA_") else "quarters": {
                period: {"present": len(presence.get(period, set())), "total": events_checked}
                for period in (
                    ("H1", "H2") if league_name.startswith("NCAA_") else ("Q1", "Q2", "Q3", "Q4")
                )
            }
        }),
        "pre_match_vs_live": {
            "pre_match": pre_match_count,
            "not_pre_match": events_checked - pre_match_count,
        },
        "not_usable": not_usable,
        "quota": quota,
        "gaps": len(_last_gaps),
    }
