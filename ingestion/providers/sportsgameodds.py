"""SportsGameOdds adapter — the quarter-capable bookmaker source (Q1–Q4 totals).

Evidence status: DOCUMENTATION_ONLY (sportsgameodds.com/faq and
/docs/data-types/markets, /docs/endpoints/getEvents) until an authenticated
smoke test occurs; every coverage claim is DOCUMENTATION_ONLY until then.
Documented contract used here:

- ``GET https://api.sportsgameodds.com/events?leagueID=NBA&dateFrom=...&dateTo=...&limit=...&oddsMarketType=prematch``
  narrowed by league/date (exact parameter names remain DOCUMENTATION_ONLY
  until an authenticated run verifies them). Auth uses the documented
  ``x-api-key`` header (key from ``SPORTSGAMEODDS_API_KEY``); the header is
  injected by the pipeline's auth-aware client — this module never places
  the key in a URL.
- Documented over/under total oddIDs: ``points-all-game-ou-over/-under``
  (full game) and ``points-all-1q-ou-over/-under`` .. ``points-all-4q-*``.
  The oddID pattern alone does NOT prove coverage — ``/markets`` must
  confirm the oddIDs before any coverage claim is upgraded; the parser here
  only maps oddIDs literally seen in a response, and only the documented
  ``points-…-ou-over/under`` totals grammar.
- Prices are American and are converted with ``books.american_to_decimal``.
  The total ``line`` is the LINE VALUE (e.g. 232.5) taken from a
  ``line``/``points`` field when present; when the payload lacks it,
  ``snapshot.line`` stays ``None`` and the snapshot is NOT usable for a
  total forecast — smoke reports it under
  ``not_usable``.
- Books come from the market's ``bookOdds`` keys; a market without any
  ``bookOdds`` yields no snapshot (book attribution is never guessed).
- ``status`` drives ``pre_match``: only clearly pre-match statuses
  ("upcoming"/"pre-match"/...) set ``pre_match=True``; live/in-progress or
  unreadable statuses fail closed. Quarters are ALLOWED for this provider,
  subject to measured coverage.
- Usage/quota: smoke() fetches ``GET /account/usage`` and reports it under
  ``quota["usage"]``.

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
from ..books import american_to_decimal
from ..models import BookSnapshot, CoverageGap, CoverageState, Market, Period, payload_hash
from ._common import bound_client_quota

PROVIDER = "sportsgameodds"
ENV_KEY = "SPORTSGAMEODDS_API_KEY"
BASE_URL = "https://api.sportsgameodds.com"
DEFAULT_LIMIT = 50

# periodID -> canonical period, from the documented oddID grammar
# ``{statID}-{statEntityID}-{periodID}-{betTypeID}-{sideID}``. Deliberately
# excludes halves (``1h``), OT and every other sub-period: they are never
# relabelled as quarters.
PERIOD_TOKENS: Tuple[Tuple[str, Period], ...] = (
    ("1q", Period.Q1),
    ("2q", Period.Q2),
    ("3q", Period.Q3),
    ("4q", Period.Q4),
    ("game", Period.FULL_GAME),
)

# Statuses that mean pre-match. Anything else (live, final, unknown) is NOT
# pre-match — fail closed, never mark live data as pre-match.
_PREGAME_STATUS_TOKENS = (
    "upcoming",
    "pre-match",
    "prematch",
    "pregame",
    "scheduled",
    "not started",
)

DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"

_last_gaps: List[CoverageGap] = []
_last_quota: Dict[str, Any] = {}


def is_enabled() -> bool:
    """True only when SPORTSGAMEODDS_API_KEY is configured (never raises)."""
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
# Fetch conventions / league gate / URLs
# ---------------------------------------------------------------------------


def _unwrap_fetch(response: Any) -> Tuple[Any, Dict[str, str]]:
    """Fetch envelope convention: ``{"json": ..., "headers": ...}`` unwraps;
    plain parsed JSON returns as (payload, {})."""
    if isinstance(response, dict) and "json" in response and isinstance(response.get("headers"), dict):
        return response["json"], response["headers"]
    return response, {}


def parse_quota(headers: Dict[str, str]) -> Dict[str, Any]:
    """Normalise ratelimit-style response headers (case-insensitive) into a
    quota dict; numeric values become ints. Used for envelope headers; the
    primary SGO usage signal is the ``/account/usage`` payload."""
    out: Dict[str, Any] = {}
    for key, value in (headers or {}).items():
        name = str(key).strip().lower()
        if not (name.startswith("x-ratelimit") or name.startswith("x-quota") or name.startswith("ratelimit")):
            continue
        text = str(value).strip()
        try:
            out[name] = int(text)
        except ValueError:
            out[name] = text
    return out


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


def _iso_date(date_yyyymmdd: str) -> str:
    digits = str(date_yyyymmdd).replace("-", "")
    return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"


def _limit() -> int:
    raw = os.environ.get("SPORTSGAMEODDS_LIMIT", "").strip()
    try:
        return int(raw) if raw else DEFAULT_LIMIT
    except ValueError:
        return DEFAULT_LIMIT


def _events_url(league: str, date_yyyymmdd: str) -> str:
    """Events URL narrowed by league and date (DOCUMENTATION_ONLY parameter
    names; event/oddID narrowing happens on targeted refreshes once game ids
    are mapped)."""
    iso = _iso_date(date_yyyymmdd)
    params = {
        "leagueID": league,
        "dateFrom": iso,
        "dateTo": iso,
        "limit": str(_limit()),
        "oddsMarketType": "prematch",
    }
    return f"{BASE_URL}/events?{urlencode(params)}"


def _usage_url() -> str:
    return f"{BASE_URL}/account/usage"


# ---------------------------------------------------------------------------
# oddID / status helpers
# ---------------------------------------------------------------------------


def parse_total_odd_id(odd_id: str) -> Optional[Tuple[Period, str]]:
    """Map an oddID literally seen in a response to ``(period, side)``.

    Only the documented points-total grammar is mapped:
    ``points-{entity}-{periodID}-ou-{over|under}`` where ``periodID`` is
    ``1q``/``2q``/``3q``/``4q``/``game`` (the documented ``-1q-`` …
    ``-game-`` suffixes). Halves, OT, other stats and other bet types return
    ``None`` and are skipped with a gap — never relabelled, and coverage is
    never inferred from the pattern alone.
    """
    text = str(odd_id or "").strip().lower()
    if not text.startswith("points-") or "-ou-" not in text:
        return None
    side: Optional[str] = None
    if text.endswith("-over"):
        side = "over"
    elif text.endswith("-under"):
        side = "under"
    if side is None:
        return None
    # core = "points-<entity>-<periodID>" (strip the "-ou-over"/"-ou-under" side)
    core = text[: -(len(side) + 4)]
    for token, period in PERIOD_TOKENS:
        if core.endswith(f"-{token}") or core == f"points-{token}":
            return period, side
    return None


def _status_text(event: Dict[str, Any]) -> str:
    status = event.get("status")
    if isinstance(status, dict):
        for key in ("displayed", "outcome", "details", "name", "type"):
            if status.get(key):
                return str(status[key])
        return ""
    return str(status or "")


def is_pregame_event(event: Dict[str, Any]) -> bool:
    """True only for clearly pre-match statuses; unknown fails closed."""
    text = _status_text(event).strip().lower()
    if not text:
        return False
    return any(token in text for token in _PREGAME_STATUS_TOKENS)


def _number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _market_line(*entries: Dict[str, Any]) -> Optional[float]:
    """Total LINE VALUE from a ``line``/``points`` field (never a price)."""
    for entry in entries:
        for key in ("line", "points"):
            value = _number(entry.get(key))
            if value is not None:
                return value
    return None


def _market_is_main(entry: Dict[str, Any]) -> bool:
    """``type: "main"`` when the payload says so; the documented
    ``points-all-*-ou-*`` oddIDs are the main lines (alternates carry
    distinct oddIDs), so absence of the field defaults to main."""
    if "type" in entry:
        return str(entry["type"]).strip().lower() == "main"
    return True


def _book_prices(entry: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    raw = entry.get("bookOdds")
    if not isinstance(raw, dict):
        return {}
    return {str(slug): book for slug, book in raw.items() if isinstance(book, dict)}


def _source_modified(entry: Dict[str, Any], event: Dict[str, Any]) -> Optional[str]:
    for source in (entry.get("timestampUpdated"), entry.get("lastModified"), event.get("lastModified"), event.get("updatedAt")):
        if source:
            return str(source)
    return None


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def _events_from_payload(payload: Any) -> List[Any]:
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for key in ("data", "events"):
            if isinstance(payload.get(key), list):
                return payload[key]
    return []


def _total_snapshots_for_group(
    base_id: str,
    period: Period,
    sides: Dict[str, Dict[str, Any]],
    game_id: str,
    observed_utc: str,
    pre_match: bool,
    event: Dict[str, Any],
) -> List[BookSnapshot]:
    """One over/under group x bookOdds keys -> per-book total snapshots.

    ``line`` comes from a ``line``/``points`` field; when the payload lacks
    it the snapshot is emitted with ``line=None`` and is NOT usable for
    total forecasts (the posted line is required) — smoke counts it under
    ``not_usable``. Books come from ``bookOdds`` keys only; a group without
    any bookOdds is skipped with a gap (book attribution is never guessed).
    """
    over = sides.get("over", {})
    under = sides.get("under", {})
    line = _market_line(over, under)
    over_by_book = _book_prices(over)
    under_by_book = _book_prices(under)
    slugs = sorted(set(over_by_book) | set(under_by_book))
    if not slugs:
        _record_gap(
            game_id,
            "bookodds",
            f"total market {base_id!r} has no bookOdds; skipped (book attribution never guessed)",
        )
        return []
    main_line = _market_is_main(over) if over else _market_is_main(under)
    if line is None:
        _record_gap(
            game_id,
            f"total_line_{period.value.lower()}",
            f"total market {base_id!r} has no line/points value; snapshot emitted with line=None (not usable)",
        )
    snapshots: List[BookSnapshot] = []
    for slug in slugs:
        over_entry = over_by_book.get(slug, {})
        under_entry = under_by_book.get(slug, {})
        over_price = _number(over_entry.get("price"))
        under_price = _number(under_entry.get("price"))
        if line is None and over_price is None and under_price is None:
            continue  # nothing usable for this book
        snapshots.append(
            BookSnapshot(
                game_id=game_id,
                provider=PROVIDER,
                book=slug,
                market=Market.TOTAL,
                period=period,
                observed_utc=observed_utc,
                source_modified_utc=_source_modified(over_entry or under_entry or over or under, event),
                pre_match=pre_match,
                main_line=main_line,
                payload_hash=payload_hash(event),
                line=line,
                over_price_decimal=american_to_decimal(over_price) if over_price is not None else None,
                under_price_decimal=american_to_decimal(under_price) if under_price is not None else None,
            )
        )
    return snapshots


def _event_teams(event: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    """Resolve home/away names only from explicit team-side flags."""
    roster = event.get("teams")
    if not isinstance(roster, list) or len(roster) != 2:
        return None, None
    home = away = None
    for team in roster:
        if not isinstance(team, dict) or not team.get("name"):
            return None, None
        is_home = team.get("home", team.get("is_home"))
        if is_home is True and home is None:
            home = str(team["name"])
        elif is_home is False and away is None:
            away = str(team["name"])
        else:
            return None, None
    return home, away


def _parse_event(event: Dict[str, Any], observed_utc: str) -> List[BookSnapshot]:
    event_id = event.get("eventID", event.get("id"))
    if event_id is None or str(event_id).strip() == "":
        _record_gap(None, "event_id", "event without eventID skipped", CoverageState.PROVIDER_ERROR)
        return []
    game_id = str(event_id)
    pre_match = is_pregame_event(event)
    markets = event.get("markets")
    if not isinstance(markets, dict) or not markets:
        _record_gap(game_id, "markets", "event without markets object skipped")
        return []

    # Group over/under sides by their exact base oddID (oddIDs literally seen
    # in the response — coverage is never inferred from the pattern).
    groups: Dict[str, Dict[str, Any]] = {}
    periods: Dict[str, Period] = {}
    for odd_id, entry in markets.items():
        if not isinstance(entry, dict):
            _record_gap(game_id, "market", f"non-dict market {odd_id!r} skipped", CoverageState.PROVIDER_ERROR)
            continue
        mapped = parse_total_odd_id(str(odd_id))
        if mapped is None:
            _record_gap(
                game_id,
                "oddid",
                f"oddID {odd_id!r} is not a documented points over/under total; skipped "
                "(halves/OT/other bet types are never relabelled)",
                CoverageState.UNSUPPORTED,
            )
            continue
        period, side = mapped
        base = str(odd_id).strip().lower()[: -(len(side) + 4)]
        groups.setdefault(base, {})[side] = entry
        periods[base] = period

    snapshots: List[BookSnapshot] = []
    for base, sides in groups.items():
        snapshots.extend(
            _total_snapshots_for_group(
                base, periods[base], sides, game_id, observed_utc, pre_match, event
            )
        )

    home, away = _event_teams(event)
    scheduled = str(event.get("startTime") or "").strip()
    try:
        identity.parse_utc(scheduled)
    except (TypeError, ValueError):
        scheduled = ""
    for snapshot in snapshots:
        snapshot.provider_event_id = game_id
        snapshot.home_name = home
        snapshot.away_name = away
        snapshot.scheduled_utc = scheduled or None
    return snapshots


def parse_events(payload: Any, observed_utc: str) -> List[BookSnapshot]:
    """Normalise an SGO events payload into per-book, per-period snapshots."""
    snapshots: List[BookSnapshot] = []
    for event in _events_from_payload(payload):
        if not isinstance(event, dict):
            _record_gap(None, "event", "non-dict event skipped", CoverageState.PROVIDER_ERROR)
            continue
        snapshots.extend(_parse_event(event, observed_utc))
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
    """Fetch normalised pre-match odds (full-game and Q1–Q4 totals).

    Returns ``[]`` when the provider is disabled (missing key) — the keyless
    pipeline must not fail. Raises ``ValueError`` for unsupported leagues
    (EuroLeague coverage is unproven and is never fabricated).
    """
    _reset_gaps()
    if not is_enabled():
        return []
    league_name = _require_supported_league(league)
    fetch = fetch_factory(PROVIDER)
    payload, _headers = _call(fetch, _events_url(league_name, date_yyyymmdd))
    observed_utc = now_utc or identity.now_utc_iso()
    return parse_events(payload, observed_utc)


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

    Reports per-quarter coverage separately (partial quarters are allowed
    and never backfilled), pre-match vs live status counts, the NOT-usable
    (line-less) total count, and the ``/account/usage`` payload under
    ``quota["usage"]``. ``date_yyyymmdd`` defaults to today's UTC date.
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
    payload, headers = _call(fetch, _events_url(league_name, date_yyyymmdd))
    observed_utc = now_utc or identity.now_utc_iso()
    snapshots = parse_events(payload, observed_utc)

    quota: Dict[str, Any] = parse_quota(headers)
    try:
        usage_payload, usage_headers = _call(fetch, _usage_url())
        quota["usage"] = usage_payload
        quota.update(parse_quota(usage_headers))
    except Exception as e:  # usage is observability; never fatal to smoke
        _record_gap(None, "usage", f"/account/usage fetch failed: {type(e).__name__}", CoverageState.PROVIDER_ERROR)
        quota["usage_error"] = type(e).__name__

    events = _events_from_payload(payload)
    events_checked = len(events)
    pre_match_count = sum(1 for e in events if isinstance(e, dict) and is_pregame_event(e))
    presence = _period_presence(snapshots)
    not_usable = sum(1 for s in snapshots if s.market is Market.TOTAL and s.line is None)
    return {
        "enabled": True,
        "provider": PROVIDER,
        "league": league_name,
        "date": date_yyyymmdd,
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
        "quota": quota,
        "gaps": len(_last_gaps),
    }
