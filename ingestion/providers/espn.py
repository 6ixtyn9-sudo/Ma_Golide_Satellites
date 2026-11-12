"""ESPN NBA provider: discovery, results, standings and BPI predictions.

Keyless, undocumented public endpoints (no SLA). Sampled responses were
LIVE_VERIFIED on 2026-11-06 (see docs/PROVIDERS.md); ordinary tests run
against small committed fixtures only, and smoke() labels its evidence
accordingly (LIVE_VERIFIED only when it is handed a real ``http.HttpClient``).

Discovery: the site scoreboard is the ONLY discovery source. The core
``events?dates=`` endpoint returned **0 events** during the 2026-11-06 live
sweep, while the site scoreboard returned the full slate. The scoreboard's
``dates=`` parameter selects the **US Eastern calendar day of the slate**,
while event datetimes inside the payload are UTC (``dates=20261105``
returned games stamped ``2026-11-06T02:00Z``). Everything canonical
produced here is UTC ISO 8601 with a ``Z`` suffix.

BPI orientation rule: the predictor attaches ``statistics`` to one team
side (observed on ``awayTeam`` in the golden NBA response, but placement is
not guaranteed). The side carrying statistics is identified by its team
``$ref``/id (pattern ``/teams/{id}``) and reconciled against the event's
competitor team ids — against the discovered ``Game`` when the caller
passes one (the recommended path used by ``smoke``), otherwise against the
payload's own ``homeTeam``/``awayTeam`` keys with a strict consistency
check (both refs must resolve and differ; identical refs raise
``IdentityError``). An unreconcilable side raises ``IdentityError`` —
home/away are never swapped to make data fit. Partial payloads
(WNBA/NCAA pattern) keep probabilities while leaving expected scores/total
``None`` (no silent imputation); callers emit the ``expected_scores`` gap.
"""

from __future__ import annotations

import math
import re
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from ..acquisition import AcquisitionClient
from ..books import american_to_decimal
from ..http import HttpClient
from ..identity import (
    canonical_team_name,
    normalize_team_name,
    now_utc_iso,
    parse_utc,
    to_utc_iso,
)
from ..leagues import period_kind
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
    Prediction,
    StandingsRow,
    payload_hash as make_payload_hash,
)
from . import CoverageAccumulator, FetchFn, IdentityError

SCOREBOARD_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/"
    "scoreboard?dates={yyyymmdd}&limit=100"
)
SUMMARY_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/summary?event={event_id}"
)
STANDINGS_URL = "https://site.api.espn.com/apis/v2/sports/basketball/nba/standings"
BPI_URL = (
    "https://sports.core.api.espn.com/v2/sports/basketball/leagues/nba/"
    "events/{event_id}/competitions/{event_id}/predictor"
)

LEAGUE_SLUGS = {
    League.NBA: "nba",
    League.WNBA: "wnba",
    League.NCAA_M: "mens-college-basketball",
    League.NCAA_W: "womens-college-basketball",
}

# Increment when stored period interpretation changes. NCAAW history written
# before v2 could label four quarters as two halves plus two overtime periods.
RESULT_PERIOD_PARSER_VERSION = 2

# Earliest non-empty core event-index years verified live on 2026-11-07.
# They are acquisition floors, not claims of complete season coverage.
HISTORY_FLOORS = {
    League.NBA: 1947,
    League.WNBA: 1997,
    League.NCAA_M: 2001,
    League.NCAA_W: 2001,
}
_EVENT_REF_RE = re.compile(r"/events/(\d+)(?:[?/]|$)")


def _league_slug(league: League) -> str:
    try:
        return LEAGUE_SLUGS[league]
    except KeyError:
        raise ValueError(f"ESPN acquisition is not enabled for {league.value}") from None


def _scoreboard_url(date_yyyymmdd: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return SCOREBOARD_URL.format(yyyymmdd=date_yyyymmdd)
    if league in {League.NCAA_M, League.NCAA_W}:
        # NCAA Division I slates exceed the default page size; groups=50
        # selects Division I. That group filter is not valid WNBA routing.
        return (
            f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
            f"scoreboard?dates={date_yyyymmdd}&groups=50&limit=500"
        )
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
        f"scoreboard?dates={date_yyyymmdd}&limit=100"
    )


def _summary_url(event_id: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return SUMMARY_URL.format(event_id=event_id)
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/summary?event={event_id}"
    )


def _bpi_url(event_id: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return BPI_URL.format(event_id=event_id)
    return (
        f"https://sports.core.api.espn.com/v2/sports/basketball/leagues/{slug}/"
        f"events/{event_id}/competitions/{event_id}/predictor"
    )


def _season_events_url(season: int, league: League, page: int) -> str:
    slug = _league_slug(league)
    return (
        f"https://sports.core.api.espn.com/v2/sports/basketball/leagues/{slug}/"
        f"events?dates={int(season)}&limit=1000&page={int(page)}"
    )


# Payload probabilities arrive as percentages (0-100); they are stored as
# Probabilities remain canonical 0-1 fractions in stored prediction facts.
PROB_SUM_TOLERANCE_PCT = 0.01  # |gameProjection + teamChanceLoss - 100| <= 0.01
MARGIN_TOLERANCE = 0.25  # |teamExpectedPts - oppExpectedPts - teamPredPtDiff| <= 0.25
# Standings sanity guard: a cumulative points-for total is always far above
# games-played x 3; a smaller value means a per-game/per-period average was
# written into a cumulative field.
PF_CUMULATIVE_MIN_FACTOR = 3

_STATE_TO_STATUS: Dict[str, GameStatus] = {
    "pre": GameStatus.SCHEDULED,
    "in": GameStatus.LIVE,
    "post": GameStatus.FINAL,
}

_TEAM_REF_RE = re.compile(r"/teams/(\d+)")


# ---------------------------------------------------------------------------
# Payload helpers
# ---------------------------------------------------------------------------


def _canonical_name(display_name: str) -> str:
    """Canonicalise a provider team name, keeping ESPN's display case.

    ``identity.canonical_team_name`` resolves alias variants to canonical
    names, but the shared alias table's keys are lowercased, so an
    already-canonical input (every ESPN ``displayName``) would come back
    lowercased. When the resolved canonical name equals the input modulo
    normalisation, the provider's display case is kept; genuinely different
    aliases still resolve through the shared table.
    """
    canonical = canonical_team_name(display_name)
    if normalize_team_name(canonical) == normalize_team_name(display_name):
        return display_name.strip()
    return canonical


def _map_status(state: Any) -> GameStatus:
    """pre/in/post -> SCHEDULED/LIVE/FINAL; anything else stays UNKNOWN."""
    return _STATE_TO_STATUS.get(str(state or ""), GameStatus.UNKNOWN)


def _team_ref_id(team: Dict[str, Any]) -> Optional[str]:
    """Team id from an explicit ``id`` or from a core-API ``$ref`` URL.

    Scoreboard/summary competitors carry ``team.id``; predictor sides only
    carry ``team.$ref`` URLs, so the id is extracted from the
    ``.../teams/{id}`` pattern.
    """
    if team.get("id") is not None:
        return str(team["id"])
    match = _TEAM_REF_RE.search(str(team.get("$ref") or ""))
    return match.group(1) if match else None


def _competitor_team_id(competitor: Dict[str, Any]) -> str:
    team = competitor.get("team") or {}
    team_id = _team_ref_id(team)
    if team_id is None:
        raise IdentityError(f"competitor team reference cannot be resolved: {team}")
    return team_id


def _competition(event: Dict[str, Any]) -> Dict[str, Any]:
    competitions = event.get("competitions") or []
    if not competitions:
        raise IdentityError(f"event {event.get('id')} carries no competitions")
    return competitions[0]


def _state(comp: Dict[str, Any]) -> Optional[str]:
    status = comp.get("status") or {}
    return (status.get("type") or {}).get("state")


def _home_away_competitors(comp: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Sides from the ``homeAway`` markers — never from array order."""
    competitors = comp.get("competitors") or []
    home = next((c for c in competitors if c.get("homeAway") == "home"), None)
    away = next((c for c in competitors if c.get("homeAway") == "away"), None)
    if home is None or away is None:
        raise IdentityError(
            "competitors missing homeAway markers; refusing to infer sides from array order"
        )
    return home, away


def _scheduled_iso(comp: Dict[str, Any], fallback_date: Any) -> str:
    raw = comp.get("date") or fallback_date
    if not raw:
        raise IdentityError("event carries no date")
    return to_utc_iso(parse_utc(raw))


def _game_from_event(event: Dict[str, Any], comp: Dict[str, Any], league: League) -> Game:
    """Canonical Game from one scoreboard event.

    ``game_id`` and ``provider_ids["espn"]`` both carry the ESPN event id;
    names are canonicalised through ``identity.canonical_team_name``; the
    scheduled time is normalised to second-precision UTC (``...Z``).
    """
    event_id = event.get("id")
    if not event_id:
        raise IdentityError("scoreboard event without id")
    home, away = _home_away_competitors(comp)
    season = event.get("season") or {}
    year = season.get("year")
    return Game(
        game_id=str(event_id),
        league=league,
        season=str(year) if year is not None else "",
        home_team_id=_competitor_team_id(home),
        away_team_id=_competitor_team_id(away),
        home_name=_canonical_name(str((home.get("team") or {}).get("displayName") or "")),
        away_name=_canonical_name(str((away.get("team") or {}).get("displayName") or "")),
        scheduled_utc=_scheduled_iso(comp, event.get("date")),
        status=_map_status(_state(comp)),
        phase="League",
        provider_ids={"espn": str(event_id)},
    )


_CONTEXT_DROP_KEYS = {
    "$ref",
    "categories",
    "footer",
    "headshot",
    "images",
    "link",
    "links",
    "logo",
    "logodark",
    "logos",
    "propbets",
    "tracking",
    "video",
    "videos",
}


def _public_context(value: Any) -> Any:
    """Retain public facts while dropping links, tracking, and bulky media."""

    if isinstance(value, dict):
        cleaned = {}
        for key, item in value.items():
            normalized = str(key).lower()
            if normalized in _CONTEXT_DROP_KEYS or "link" in normalized:
                continue
            cleaned[str(key)] = _public_context(item)
        return cleaned
    if isinstance(value, list):
        return [_public_context(item) for item in value[:100]]
    if isinstance(value, str):
        return value[:20000]
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return str(value)[:1000]


def _nested(mapping: Dict[str, Any], *path: str) -> Any:
    value: Any = mapping
    for key in path:
        if not isinstance(value, dict):
            return None
        value = value.get(key)
    return value


def _line_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    text = str(value).strip()
    if text[:1].lower() in {"o", "u"}:
        text = text[1:]
    try:
        number = float(text)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def _decimal_price(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    try:
        american = float(str(value).strip())
        if american == 0 or not math.isfinite(american):
            return None
        price = american_to_decimal(american)
    except (TypeError, ValueError):
        return None
    return price if math.isfinite(price) and 1.0 < price <= 1000.0 else None


def _odds_context(offer: Dict[str, Any]) -> Dict[str, Any]:
    """Extract open/current numbers without retaining affiliate/tracking links."""

    provider = offer.get("provider") or {}
    context: Dict[str, Any] = {
        "book": str(
            provider.get("displayName") or provider.get("name") or provider.get("id") or ""
        ),
        "provider_id": str(provider.get("id") or ""),
    }
    for market_name, sides in (
        ("moneyline", ("home", "away")),
        ("pointSpread", ("home", "away")),
        ("total", ("over", "under")),
    ):
        market = offer.get(market_name) or {}
        cleaned_market = {}
        for side in sides:
            phases = {}
            for phase in ("open", "close"):
                values = market.get(side, {}).get(phase) or {}
                phases[phase] = {
                    key: values.get(key) for key in ("line", "odds") if values.get(key) is not None
                }
            if any(phases.values()):
                cleaned_market[side] = phases
        if cleaned_market:
            context[market_name] = cleaned_market
    for key in ("details", "overUnder", "spread"):
        if offer.get(key) is not None:
            context[key] = offer[key]
    return context


def _book_snapshots_for_offer(
    event_id: str,
    game: Game,
    offer: Dict[str, Any],
    observed_utc: str,
    pre_match: bool,
) -> Tuple[List[BookSnapshot], List[str]]:
    provider = offer.get("provider") or {}
    raw_book = str(
        provider.get("displayName") or provider.get("name") or provider.get("id") or ""
    ).strip()
    if not raw_book:
        return [], ["odds offer without a provider/book identity was ignored"]
    # Normalize historical ESPN variant ``Draft Kings`` -> ``DraftKings``
    # so logical picks do not duplicate across book spellings.
    from ..identity import canonical_book_name

    book = canonical_book_name(raw_book)
    common = {
        "game_id": game.game_id,
        "provider": "espn",
        "book": book,
        "period": Period.FULL_GAME,
        "observed_utc": observed_utc,
        "pre_match": pre_match,
        "main_line": True,
        "provider_event_id": str(event_id),
        "home_name": game.home_name,
        "away_name": game.away_name,
        "scheduled_utc": game.scheduled_utc,
    }
    context_hash = make_payload_hash(_odds_context(offer))
    snapshots: List[BookSnapshot] = []
    notes: List[str] = []

    home_price = _decimal_price(_nested(offer, "moneyline", "home", "close", "odds"))
    away_price = _decimal_price(_nested(offer, "moneyline", "away", "close", "odds"))
    # Older scoreboard/summary variants expose only these flat American prices.
    if home_price is None:
        home_price = _decimal_price(_nested(offer, "homeTeamOdds", "moneyLine"))
    if away_price is None:
        away_price = _decimal_price(_nested(offer, "awayTeamOdds", "moneyLine"))
    # Free-tier fallback: if close is not yet published, use open (still pre-match, still timestamped at discovery).
    if home_price is None:
        home_price = _decimal_price(_nested(offer, "moneyline", "home", "open", "odds"))
        if home_price is None:
            home_price = _decimal_price(_nested(offer, "homeTeamOdds", "moneyLine"))
    if away_price is None:
        away_price = _decimal_price(_nested(offer, "moneyline", "away", "open", "odds"))
        if away_price is None:
            away_price = _decimal_price(_nested(offer, "awayTeamOdds", "moneyLine"))
    if home_price is not None or away_price is not None:
        snapshots.append(
            BookSnapshot(
                **common,
                market=Market.MONEYLINE,
                home_price_decimal=home_price,
                away_price_decimal=away_price,
                payload_hash=context_hash,
            )
        )

    spread_home = offer.get("pointSpread", {}).get("home", {})
    spread_away = offer.get("pointSpread", {}).get("away", {})
    home_line = _line_number(_nested(spread_home, "close", "line"))
    away_line = _line_number(_nested(spread_away, "close", "line"))
    if home_line is None:
        home_line = _line_number(_nested(spread_home, "open", "line"))
    if away_line is None:
        away_line = _line_number(_nested(spread_away, "open", "line"))
    if home_line is None and away_line is not None:
        home_line = -away_line
    if home_line is not None and away_line is not None:
        if abs(home_line + away_line) > 0.001:
            notes.append(f"{book}: contradictory home/away spread lines ignored")
            home_line = None
    if home_line is not None:
        # Prefer close prices, fall back to open when close not yet posted.
        spread_home_price = _decimal_price(_nested(spread_home, "close", "odds"))
        spread_away_price = _decimal_price(_nested(spread_away, "close", "odds"))
        if spread_home_price is None:
            spread_home_price = _decimal_price(_nested(spread_home, "open", "odds"))
        if spread_away_price is None:
            spread_away_price = _decimal_price(_nested(spread_away, "open", "odds"))
        snapshots.append(
            BookSnapshot(
                **common,
                market=Market.SPREAD,
                line=home_line,
                home_price_decimal=spread_home_price,
                away_price_decimal=spread_away_price,
                payload_hash=context_hash,
            )
        )

    over = offer.get("total", {}).get("over", {})
    under = offer.get("total", {}).get("under", {})
    over_line = _line_number(_nested(over, "close", "line"))
    under_line = _line_number(_nested(under, "close", "line"))
    total_line = over_line if over_line is not None else under_line
    if total_line is None:
        over_line = _line_number(_nested(over, "open", "line"))
        under_line = _line_number(_nested(under, "open", "line"))
        total_line = over_line if over_line is not None else under_line
    if total_line is None:
        total_line = _line_number(offer.get("overUnder"))
    if over_line is not None and under_line is not None:
        if abs(over_line - under_line) > 0.001:
            notes.append(f"{book}: contradictory over/under total lines ignored")
            total_line = None
    if total_line is not None:
        over_price = _decimal_price(_nested(over, "close", "odds"))
        under_price = _decimal_price(_nested(under, "close", "odds"))
        if over_price is None:
            over_price = _decimal_price(_nested(over, "open", "odds"))
        if under_price is None:
            under_price = _decimal_price(_nested(under, "open", "odds"))
        snapshots.append(
            BookSnapshot(
                **common,
                market=Market.TOTAL,
                line=total_line,
                over_price_decimal=over_price,
                under_price_decimal=under_price,
                payload_hash=context_hash,
            )
        )
    return snapshots, notes


def book_snapshots_from_event(
    event: Dict[str, Any], game: Game, observed_utc: Optional[str] = None
) -> List[BookSnapshot]:
    """Return full-game named-book markets embedded in a scoreboard event.

    Only current/close prices become timestamped snapshots. Open values are
    retained by scoreboard context but have no provider observation timestamp,
    so they are never misrepresented as opening snapshots or used for CLV.
    Spread lines are home-oriented; total lines are points, never prices.
    """

    observed = to_utc_iso(parse_utc(observed_utc or now_utc_iso()))
    comp = _competition(event)
    pre_match = _state(comp) == "pre" and parse_utc(observed) < parse_utc(game.scheduled_utc)
    snapshots: List[BookSnapshot] = []
    for offer in comp.get("odds") or []:
        if isinstance(offer, dict):
            parsed, _notes = _book_snapshots_for_offer(
                str(event.get("id") or game.game_id), game, offer, observed, pre_match
            )
            snapshots.extend(parsed)
    return snapshots


def _scoreboard_context(
    event: Dict[str, Any],
    comp: Dict[str, Any],
    observed_utc: str,
    book_parse_notes: List[str],
) -> Dict[str, Any]:
    home, away = _home_away_competitors(comp)
    teams = {}
    for side, competitor in (("home", home), ("away", away)):
        teams[side] = _public_context(
            {
                "team_id": _competitor_team_id(competitor),
                "team": competitor.get("team") or {},
                "records": competitor.get("records") or [],
                "statistics": competitor.get("statistics") or [],
                "leaders": competitor.get("leaders") or [],
            }
        )
    return {
        "schema_version": "magolide.espn.scoreboard-context.v1",
        "observed_utc": observed_utc,
        "source_payload_hash": make_payload_hash(event),
        "event_status": _public_context(comp.get("status") or event.get("status") or {}),
        "season": _public_context(event.get("season") or {}),
        "venue": _public_context(comp.get("venue") or {}),
        "neutral_site": comp.get("neutralSite"),
        "conference_competition": comp.get("conferenceCompetition"),
        "notes": _public_context(comp.get("notes") or []),
        "broadcasts": _public_context(comp.get("broadcasts") or []),
        "teams": teams,
        "odds": [
            _odds_context(offer) for offer in comp.get("odds") or [] if isinstance(offer, dict)
        ],
        "book_parse_notes": book_parse_notes,
    }


def _summary_context(payload: Dict[str, Any], game: Game, observed_utc: str) -> Dict[str, Any]:
    """Retain validated high-value summary evidence, not league-wide news noise."""

    header = payload.get("header") or {}
    competitions = header.get("competitions") or []
    comp = competitions[0] if competitions else {}
    if competitions:
        home, away = _home_away_competitors(comp)
        ids = {_competitor_team_id(home), _competitor_team_id(away)}
        if ids != {str(game.home_team_id), str(game.away_team_id)}:
            raise IdentityError(
                f"event {game.game_id}: summary competitors contradict scoreboard teams"
            )
    article = payload.get("article") or {}
    modified = article.get("lastModified") or article.get("published")
    try:
        source_modified = to_utc_iso(parse_utc(modified)) if modified else None
    except (TypeError, ValueError):
        source_modified = None
    boxscore = payload.get("boxscore") or {}
    player_boxscore = boxscore.get("players") or []
    player_ids = {
        str((entry.get("athlete") or {}).get("id"))
        for team in player_boxscore
        for category in team.get("statistics") or []
        for entry in category.get("athletes") or []
        if (entry.get("athlete") or {}).get("id") is not None
    }
    standings = payload.get("standings") or {}
    standings_groups = standings.get("groups") if isinstance(standings, dict) else []
    return {
        "schema_version": "magolide.espn.summary-context.v1",
        "observed_utc": observed_utc,
        "source_modified_utc": source_modified,
        "source_payload_hash": make_payload_hash(payload),
        "game_info": _public_context(payload.get("gameInfo") or {}),
        "format": _public_context(payload.get("format") or comp.get("format") or {}),
        "period_evidence": _public_context(
            {
                "status": comp.get("status") or {},
                "competitors": [
                    {
                        "homeAway": competitor.get("homeAway"),
                        "team_id": _competitor_team_id(competitor),
                        "score": competitor.get("score"),
                        "linescores": competitor.get("linescores") or [],
                    }
                    for competitor in comp.get("competitors") or []
                ],
            }
        ),
        "team_boxscore": _public_context(boxscore.get("teams") or []),
        "player_boxscore": _public_context(player_boxscore),
        "leaders": _public_context(boxscore.get("leaders") or payload.get("leaders") or []),
        # Full PBP/win-probability/shot-chart arrays are not copied into
        # canonical evidence (the enabled raw HTTP cache retains them). Counts
        # expose availability without truncating or duplicating thousands of
        # potentially in-play rows in every game record.
        "context_coverage": {
            "player_boxscore_teams": len(player_boxscore),
            "player_boxscore_athletes": len(player_ids),
            "play_by_play_events_in_raw_payload": len(payload.get("plays") or []),
            "win_probability_points_in_raw_payload": len(payload.get("winprobability") or []),
            "shot_chart_events_in_raw_payload": len(payload.get("shotChart") or []),
            "standings_groups_in_raw_payload": len(standings_groups or []),
        },
        "last_five_games": _public_context(payload.get("lastFiveGames") or []),
        "injuries": _public_context(payload.get("injuries") or []),
        "article": _public_context(article),
        "predictor": _public_context(payload.get("predictor") or {}),
        "against_the_spread": _public_context(payload.get("againstTheSpread") or []),
    }


def season_event_index(season: int, fetch: FetchFn, league: League) -> Dict[str, Any]:
    """Enumerate one core event-index year with bounded pagination.

    The index is an acquisition manifest, not a completeness guarantee. Older
    seasons can contain only playoff events or lack period/boxscore sources.
    """

    event_ids: List[str] = []
    count = 0
    page_count = 0
    metadata = []
    for page in range(1, 26):
        payload = fetch(_season_events_url(season, league, page)) or {}
        if not isinstance(payload, dict) or payload.get("code"):
            raise ValueError(f"season {season}: invalid core event index")
        if page == 1:
            count = int(payload.get("count") or 0)
            page_count = int(payload.get("pageCount") or 0)
        if payload.get("$meta"):
            metadata.append(_public_context(payload["$meta"]))
        for item in payload.get("items") or []:
            match = _EVENT_REF_RE.search(str(item.get("$ref") or ""))
            if match and match.group(1) not in event_ids:
                event_ids.append(match.group(1))
        if page >= page_count:
            break
    else:
        raise ValueError(f"season {season}: event index exceeds 25-page bound")
    if count and not event_ids:
        raise IdentityError(f"season {season}: event index has no resolvable event refs")
    return {
        "schema_version": "magolide.espn.season-index.v1",
        "index_year": int(season),
        "count": count,
        "page_count": page_count,
        "event_ids": event_ids,
        "metadata": metadata,
        "completeness": "provider_index_only_not_a_complete_season_claim",
    }


def fetch_game_context(
    event_id: str,
    fetch: FetchFn,
    game: Game,
    observed_utc: Optional[str] = None,
) -> Game:
    """Fetch one event summary and attach validated context to a copied game."""

    payload = fetch(_summary_url(event_id, game.league)) or {}
    if not isinstance(payload, dict) or not payload:
        raise ValueError(f"event {event_id}: empty summary payload")
    observed = to_utc_iso(parse_utc(observed_utc or now_utc_iso()))
    espn_evidence = dict(game.evidence.get("espn") or {})
    espn_evidence["summary"] = _summary_context(payload, game, observed)
    return replace(game, evidence={**game.evidence, "espn": espn_evidence})


def _linescore_values(competitor: Dict[str, Any]) -> List[float]:
    """Numeric linescore values, preferring ``value`` over ``displayValue``.

    The key distinction matters: ``dict.get(key, default)`` would NOT fall
    back when ``value`` exists but is ``None`` (summary payloads ship
    ``displayValue``-only elements), so the fallback is explicit.
    """
    values: List[float] = []
    for element in competitor.get("linescores") or []:
        raw = element.get("value")
        if raw is None:
            raw = element.get("displayValue")
        try:
            values.append(float(raw))
        except (TypeError, ValueError):
            raise IdentityError(f"unparseable linescore element: {element!r}") from None
    return values


def _final_score(competitor: Dict[str, Any]) -> float:
    try:
        return float(competitor["score"])
    except (KeyError, TypeError, ValueError):
        raise IdentityError(
            f"competitor {_competitor_team_id(competitor)} has no parseable final score"
        ) from None


def _actual_from_competitors(
    home: Dict[str, Any],
    away: Dict[str, Any],
    source: str,
    league: League = League.NBA,
    regulation_periods: Optional[int] = None,
) -> ActualResult:
    """Build an ActualResult from native competitor linescores.

    NBA and modern NCAA women use four quarters; NCAA men use two halves.
    An explicit two-period regulation declaration preserves older women's
    payloads. Every further element is one overtime period, summed into the
    canonical combined OT pair. Missing native periods, mismatched lengths,
    and unparseable values quarantine the result rather than relabelling or
    backfilling it.
    """
    home_values = _linescore_values(home)
    away_values = _linescore_values(away)
    if regulation_periods not in {None, 2, 4}:
        raise IdentityError(f"unsupported regulation period count: {regulation_periods!r}")
    native_count = regulation_periods or (2 if period_kind(league) is PeriodKind.HALVES else 4)
    native_kind = PeriodKind.HALVES if native_count == 2 else PeriodKind.QUARTERS
    period_label = "H1/H2" if native_count == 2 else "Q1..Q4"
    if len(home_values) < native_count or len(away_values) < native_count:
        raise IdentityError(f"linescores missing or incomplete (need at least {period_label})")
    if len(home_values) != len(away_values):
        raise IdentityError(
            f"mismatched linescore lengths: home {len(home_values)} vs away {len(away_values)}"
        )
    home_ot = sum(home_values[native_count:]) if len(home_values) > native_count else None
    away_ot = sum(away_values[native_count:]) if len(away_values) > native_count else None
    return ActualResult(
        period_kind=native_kind,
        home_quarters=home_values[:native_count],
        away_quarters=away_values[:native_count],
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=_final_score(home),
        away_final=_final_score(away),
        overtime_periods=max(len(home_values), len(away_values)) - native_count,
        source=source,
    )


def _regulation_periods(
    comp: Dict[str, Any], payload: Optional[Dict[str, Any]] = None
) -> Optional[int]:
    format_info = comp.get("format") or (payload or {}).get("format") or {}
    value = _nested(format_info, "regulation", "periods")
    try:
        periods = int(value)
    except (TypeError, ValueError):
        return None
    return periods if periods in {2, 4} else None


def _actual_for_history(
    home: Dict[str, Any],
    away: Dict[str, Any],
    comp: Dict[str, Any],
    league: League,
    payload: Optional[Dict[str, Any]] = None,
) -> Tuple[ActualResult, str]:
    """Preserve final scores even when native period evidence is absent/broken."""

    regulation = _regulation_periods(comp, payload)
    if not home.get("linescores") and not away.get("linescores"):
        return (
            ActualResult(
                period_kind=PeriodKind.UNKNOWN,
                home_quarters=[],
                away_quarters=[],
                home_final=_final_score(home),
                away_final=_final_score(away),
                source="espn",
            ),
            "final score present without native linescores",
        )
    try:
        return (
            _actual_from_competitors(
                home,
                away,
                source="espn",
                league=league,
                regulation_periods=regulation,
            ),
            "",
        )
    except IdentityError as exc:
        return (
            ActualResult(
                period_kind=PeriodKind.UNKNOWN,
                home_quarters=[],
                away_quarters=[],
                home_final=_final_score(home),
                away_final=_final_score(away),
                source="espn",
            ),
            str(exc),
        )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def discover_games(date_yyyymmdd: str, fetch: FetchFn, league: League = League.NBA) -> List[Game]:
    """Discover one slate via the site scoreboard (core ``events?dates=``
    returned 0 events in the 2026-11-06 live sweep — see module docstring).

    ``date_yyyymmdd`` is the US Eastern calendar day of the slate (the
    scoreboard ``dates=`` semantics); event datetimes are UTC and are
    normalised to UTC ISO 8601 here, so a ``20261106`` slate yields games
    stamped ``2026-11-07T00:00Z`` and later. Home/away come from the
    competitors' ``homeAway`` markers, never array order.
    """
    payload = fetch(_scoreboard_url(date_yyyymmdd, league)) or {}
    return [
        _game_from_event(event, _competition(event), league) for event in payload.get("events", [])
    ]


def discover_games_with_evidence(
    date_yyyymmdd: str,
    fetch: FetchFn,
    league: League = League.NBA,
    observed_utc: Optional[str] = None,
) -> Tuple[List[Game], Dict[str, List[BookSnapshot]]]:
    """Discover games plus same-response context and full-game book snapshots."""

    payload = fetch(_scoreboard_url(date_yyyymmdd, league)) or {}
    observed = to_utc_iso(parse_utc(observed_utc or now_utc_iso()))
    games: List[Game] = []
    books: Dict[str, List[BookSnapshot]] = {}
    for event in payload.get("events", []):
        comp = _competition(event)
        game = _game_from_event(event, comp, league)
        game_books: List[BookSnapshot] = []
        notes: List[str] = []
        for offer in comp.get("odds") or []:
            if not isinstance(offer, dict):
                continue
            pre_match = _state(comp) == "pre" and parse_utc(observed) < parse_utc(
                game.scheduled_utc
            )
            parsed, parse_notes = _book_snapshots_for_offer(
                str(event.get("id") or game.game_id),
                game,
                offer,
                observed,
                pre_match,
            )
            game_books.extend(parsed)
            notes.extend(parse_notes)
        espn_evidence = dict(game.evidence.get("espn") or {})
        espn_evidence["scoreboard"] = _scoreboard_context(event, comp, observed, notes)
        game.evidence = {**game.evidence, "espn": espn_evidence}
        games.append(game)
        books[game.game_id] = game_books
    return games, books


def _stats_by_name(statistics: List[Dict[str, Any]]) -> Dict[str, Any]:
    return {s.get("name"): s.get("value") for s in statistics or []}


def _dual_sided_bpi_prediction(
    event_id: str,
    payload: Dict[str, Any],
    home_side: Dict[str, Any],
    away_side: Dict[str, Any],
    game: Optional[Game],
    now_utc: Optional[str],
) -> Prediction:
    """Validate a predictor carrying independent statistics for both teams."""

    home_ref = _team_ref_id(home_side.get("team") or {})
    away_ref = _team_ref_id(away_side.get("team") or {})
    if not home_ref or not away_ref or home_ref == away_ref:
        raise IdentityError(
            f"event {event_id}: dual-sided predictor requires two distinct team refs"
        )
    if game is not None and (
        home_ref != str(game.home_team_id) or away_ref != str(game.away_team_id)
    ):
        raise IdentityError(
            f"event {event_id}: predictor refs {home_ref}/{away_ref} contradict "
            f"scoreboard home/away {game.home_team_id}/{game.away_team_id}"
        )
    home_stats = _stats_by_name(home_side["statistics"])
    away_stats = _stats_by_name(away_side["statistics"])

    def probabilities(stats: Dict[str, Any], label: str) -> Tuple[float, float]:
        projection = stats.get("gameProjection")
        chance_loss = stats.get("teamChanceLoss")
        if projection is None or chance_loss is None:
            raise ValueError(f"event {event_id}: {label} predictor missing win probabilities")
        for name, value in (
            ("gameProjection", projection),
            ("teamChanceLoss", chance_loss),
        ):
            if not 0.0 <= float(value) <= 100.0:
                raise ValueError(f"event {event_id}: {label} {name} {value} outside 0-100")
        if abs(float(projection) + float(chance_loss) - 100.0) > PROB_SUM_TOLERANCE_PCT:
            raise ValueError(f"event {event_id}: {label} probabilities do not sum to 100")
        return float(projection), float(chance_loss)

    home_projection, home_loss = probabilities(home_stats, "homeTeam")
    away_projection, away_loss = probabilities(away_stats, "awayTeam")
    if abs(home_projection + away_projection - 100.0) > PROB_SUM_TOLERANCE_PCT:
        raise ValueError(f"event {event_id}: home/away projections are not complementary")
    if (
        abs(home_loss - away_projection) > PROB_SUM_TOLERANCE_PCT
        or abs(away_loss - home_projection) > PROB_SUM_TOLERANCE_PCT
    ):
        raise ValueError(f"event {event_id}: reciprocal predictor probabilities contradict")

    def number(stats: Dict[str, Any], key: str) -> Optional[float]:
        value = stats.get(key)
        return float(value) if value is not None else None

    for stats, label in ((home_stats, "homeTeam"), (away_stats, "awayTeam")):
        side = number(stats, "teamExpectedPts")
        opponent = number(stats, "oppExpectedPts")
        difference = number(stats, "teamPredPtDiff")
        if (
            side is not None
            and opponent is not None
            and difference is not None
            and abs((side - opponent) - difference) > MARGIN_TOLERANCE
        ):
            raise ValueError(f"event {event_id}: {label} expected scores contradict its margin")

    home_scores = [
        value
        for value in (
            number(home_stats, "teamExpectedPts"),
            number(away_stats, "oppExpectedPts"),
        )
        if value is not None
    ]
    away_scores = [
        value
        for value in (
            number(away_stats, "teamExpectedPts"),
            number(home_stats, "oppExpectedPts"),
        )
        if value is not None
    ]
    if (len(home_scores) == 2 and abs(home_scores[0] - home_scores[1]) > MARGIN_TOLERANCE) or (
        len(away_scores) == 2 and abs(away_scores[0] - away_scores[1]) > MARGIN_TOLERANCE
    ):
        raise ValueError(f"event {event_id}: dual-sided expected scores contradict each other")
    home_margin = number(home_stats, "teamPredPtDiff")
    away_margin = number(away_stats, "teamPredPtDiff")
    margins = [
        value
        for value in (
            home_margin,
            -away_margin if away_margin is not None else None,
        )
        if value is not None
    ]
    if len(margins) == 2 and abs(margins[0] - margins[1]) > MARGIN_TOLERANCE:
        raise ValueError(f"event {event_id}: dual-sided margins contradict")

    expected_home = home_scores[0] if home_scores else None
    expected_away = away_scores[0] if away_scores else None
    last_modified = payload.get("lastModified")
    prediction = Prediction(
        game_id=str(event_id),
        provider="espn",
        model="espn-bpi",
        produced_by="ESPN BPI",
        observed_utc=to_utc_iso(parse_utc(now_utc)) if now_utc else now_utc_iso(),
        source_modified_utc=(to_utc_iso(parse_utc(last_modified)) if last_modified else None),
        source_revision=str(last_modified) if last_modified else None,
        prob_home=home_projection / 100.0,
        prob_away=away_projection / 100.0,
        expected_home_score=expected_home,
        expected_away_score=expected_away,
        expected_total=(
            expected_home + expected_away
            if expected_home is not None and expected_away is not None
            else None
        ),
        expected_margin=margins[0] if margins else None,
    )
    prediction.validate_orientation(tolerance=MARGIN_TOLERANCE)
    return prediction


def fetch_bpi(
    event_id: str,
    fetch: FetchFn,
    game: Optional[Game] = None,
    now_utc: Optional[str] = None,
    league: League = League.NBA,
) -> Prediction:
    """Fetch the ESPN BPI predictor and normalise it to a canonical Prediction.

    Orientation: single-sided payloads reconcile the populated side's team id
    against the event competitors. Dual-sided NCAA payloads instead require
    exact home/away team-ref matches and reciprocal probabilities; they are
    validated, not rejected merely because both sides carry statistics.
    Without a game the payload keys orient the sides (the ``name`` field,
    "Away at Home", documents the same order but carries no id→name table,
    so it cannot participate in id validation) and team refs must be distinct.
    A stats-side id matching neither competitor, or identical/missing refs,
    raises ``IdentityError`` — never a swap.

    Mapping: ``gameProjection`` = the stats side's win probability
    (payload 0-100 -> stored 0-1), ``teamChanceLoss`` = the opponent's,
    ``teamExpectedPts``/``oppExpectedPts`` = that side's/opponent's
    expected scores, ``teamPredPtDiff`` = side minus opponent (flipped to
    home-minus-away when the stats side is away). Dual-sided payloads map each
    side's own ``gameProjection`` and expected points directly to that team.

    Validation (quarantine via ``ValueError``): each probability in 0-100,
    ``|gameProjection + teamChanceLoss - 100| <= 0.01`` and
    ``|teamExpectedPts - oppExpectedPts - teamPredPtDiff| <= 0.25``.
    Partial payloads (WNBA/NCAA pattern: no expected points) keep
    probabilities and leave scores/total/margin ``None``; ``smoke`` emits
    the ``expected_scores`` gap for them.
    """
    if game is not None:
        event_id = str(game.provider_ids.get("espn") or event_id)
    league = game.league if game is not None else league
    payload = fetch(_bpi_url(event_id, league))
    if not isinstance(payload, dict) or not payload:
        raise ValueError(f"event {event_id}: empty predictor payload")

    home_side = payload.get("homeTeam") or {}
    away_side = payload.get("awayTeam") or {}
    stats_sides = [side for side in (home_side, away_side) if side.get("statistics")]
    if not stats_sides:
        raise ValueError(f"event {event_id}: predictor carries no team statistics")
    if len(stats_sides) == 2:
        return _dual_sided_bpi_prediction(event_id, payload, home_side, away_side, game, now_utc)
    stats_side = stats_sides[0]
    side_team_id = _competitor_team_id(stats_side)

    home_ref = _team_ref_id(home_side.get("team") or {})
    away_ref = _team_ref_id(away_side.get("team") or {})
    if home_ref is not None and away_ref is not None and home_ref == away_ref:
        raise IdentityError(
            f"event {event_id}: predictor homeTeam and awayTeam both reference "
            f"team {home_ref}; sides cannot be reconciled"
        )

    if game is not None and str(game.home_team_id) and str(game.away_team_id):
        home_id, away_id = str(game.home_team_id), str(game.away_team_id)
        if side_team_id == home_id:
            stats_is_home = True
        elif side_team_id == away_id:
            stats_is_home = False
        else:
            raise IdentityError(
                f"event {event_id}: predictor stats-side team id {side_team_id} matches "
                f"neither competitor (home {home_id}, away {away_id}); refusing to guess"
            )
    else:
        # No competitor ids available: orient by the payload's explicit
        # homeTeam/awayTeam keys (see docstring for the name-field caveat).
        stats_is_home = stats_side is home_side

    stats = _stats_by_name(stats_side["statistics"])
    projection = stats.get("gameProjection")
    chance_loss = stats.get("teamChanceLoss")
    if projection is None or chance_loss is None:
        raise ValueError(f"event {event_id}: predictor missing win probabilities")
    for name, value in (("gameProjection", projection), ("teamChanceLoss", chance_loss)):
        if not 0.0 <= float(value) <= 100.0:
            raise ValueError(f"event {event_id}: {name} {value} outside the 0-100 percent range")
    if abs(float(projection) + float(chance_loss) - 100.0) > PROB_SUM_TOLERANCE_PCT:
        raise ValueError(
            f"event {event_id}: probabilities sum to "
            f"{float(projection) + float(chance_loss):.4f}, not 100"
        )
    prob_side = float(projection) / 100.0
    prob_opp = float(chance_loss) / 100.0

    expected_side = stats.get("teamExpectedPts")
    expected_opp = stats.get("oppExpectedPts")
    pred_pt_diff = stats.get("teamPredPtDiff")
    if (
        expected_side is not None
        and expected_opp is not None
        and pred_pt_diff is not None
        and abs((float(expected_side) - float(expected_opp)) - float(pred_pt_diff))
        > MARGIN_TOLERANCE
    ):
        raise ValueError(
            f"event {event_id}: expected-score margin contradicts teamPredPtDiff "
            f"({expected_side} - {expected_opp} vs {pred_pt_diff})"
        )

    if stats_is_home:
        prob_home, prob_away = prob_side, prob_opp
        expected_home = float(expected_side) if expected_side is not None else None
        expected_away = float(expected_opp) if expected_opp is not None else None
        expected_margin = float(pred_pt_diff) if pred_pt_diff is not None else None
    else:
        prob_home, prob_away = prob_opp, prob_side
        expected_home = float(expected_opp) if expected_opp is not None else None
        expected_away = float(expected_side) if expected_side is not None else None
        expected_margin = -float(pred_pt_diff) if pred_pt_diff is not None else None
    expected_total = (
        float(expected_side) + float(expected_opp)
        if expected_side is not None and expected_opp is not None
        else None
    )

    last_modified = payload.get("lastModified")
    prediction = Prediction(
        game_id=str(event_id),
        provider="espn",
        model="espn-bpi",
        produced_by="ESPN BPI",
        observed_utc=(to_utc_iso(parse_utc(now_utc)) if now_utc else now_utc_iso()),
        source_modified_utc=(to_utc_iso(parse_utc(last_modified)) if last_modified else None),
        source_revision=str(last_modified) if last_modified else None,
        prob_home=prob_home,
        prob_away=prob_away,
        expected_home_score=expected_home,
        expected_away_score=expected_away,
        expected_total=expected_total,
        expected_margin=expected_margin,
    )
    # Belt-and-braces: the canonical validator re-checks orientation
    # (score-vs-margin and score-vs-total consistency) before release.
    prediction.validate_orientation(tolerance=MARGIN_TOLERANCE)
    return prediction


def fetch_game_result(
    event_id: str,
    fetch: FetchFn,
    game: Optional[Game] = None,
    league: League = League.NBA,
) -> Game:
    """Fetch the summary endpoint and fill ``actual`` from competitor linescores.

    Raises ``IdentityError`` while the game is not final (state != 'post')
    or its linescores are missing/inconsistent — a result row must never be
    built from live or partial data. When ``game`` is given it is copied
    (never mutated) with ``actual`` and FINAL status; otherwise a Game is
    built from the summary header.
    """
    league = game.league if game is not None else league
    payload = fetch(_summary_url(event_id, league)) or {}
    header = payload.get("header") or {}
    competitions = header.get("competitions") or []
    if not competitions:
        raise IdentityError(f"event {event_id}: summary carries no competition")
    comp = competitions[0]
    state = _state(comp)
    if state != "post":
        raise IdentityError(
            f"event {event_id}: summary status is not final (state={state!r}); result quarantined"
        )
    home, away = _home_away_competitors(comp)
    actual = _actual_from_competitors(
        home,
        away,
        source="espn",
        league=league,
        regulation_periods=_regulation_periods(comp, payload),
    )
    if game is not None:
        return replace(
            game,
            status=GameStatus.FINAL,
            actual=actual,
            provider_ids={**game.provider_ids, "espn": str(event_id)},
        )
    season = header.get("season") or {}
    year = season.get("year")
    return Game(
        game_id=str(event_id),
        league=league,
        season=str(year) if year is not None else "",
        home_team_id=_competitor_team_id(home),
        away_team_id=_competitor_team_id(away),
        home_name=_canonical_name(str((home.get("team") or {}).get("displayName") or "")),
        away_name=_canonical_name(str((away.get("team") or {}).get("displayName") or "")),
        scheduled_utc=_scheduled_iso(comp, None),
        status=GameStatus.FINAL,
        phase="League",
        provider_ids={"espn": str(event_id)},
        actual=actual,
    )


def fetch_historical_event(
    event_id: str,
    fetch: FetchFn,
    league: League,
    observed_utc: Optional[str] = None,
    index_year: Optional[int] = None,
) -> Game:
    """Acquire one indexed summary without requiring period-complete evidence."""

    payload = fetch(_summary_url(event_id, league)) or {}
    if not isinstance(payload, dict) or not payload:
        raise ValueError(f"event {event_id}: empty historical summary")
    header = payload.get("header") or {}
    competitions = header.get("competitions") or []
    if not competitions:
        raise IdentityError(f"event {event_id}: summary carries no competition")
    comp = competitions[0]
    home, away = _home_away_competitors(comp)
    season = header.get("season") or {}
    year = season.get("year")
    season_type = season.get("type")
    observed = to_utc_iso(parse_utc(observed_utc or now_utc_iso()))
    state = _state(comp)
    game = Game(
        game_id=str(event_id),
        league=league,
        season=str(year) if year is not None else "",
        home_team_id=_competitor_team_id(home),
        away_team_id=_competitor_team_id(away),
        home_name=_canonical_name(str((home.get("team") or {}).get("displayName") or "")),
        away_name=_canonical_name(str((away.get("team") or {}).get("displayName") or "")),
        scheduled_utc=_scheduled_iso(comp, header.get("date")),
        status=_map_status(state),
        phase="Playoffs" if season_type == 3 else "League",
        provider_ids={"espn": str(event_id)},
    )
    actual = None
    period_note = ""
    if state == "post":
        actual, period_note = _actual_for_history(home, away, comp, league, payload)
        game.status = GameStatus.FINAL
        game.actual = actual
        if actual.period_kind is PeriodKind.UNKNOWN:
            coverage = "final_score_only"
        elif actual.reconciles()[0]:
            coverage = "reconciled_periods"
        else:
            coverage = "unreconciled_periods"
    else:
        coverage = "non_final_provider_status"
        period_note = f"provider status is {state!r}, not post"
    espn_evidence = {
        "summary": _summary_context(payload, game, observed),
        "history": {
            "schema_version": "magolide.espn.history-event.v1",
            "result_period_parser_version": RESULT_PERIOD_PARSER_VERSION,
            "index_year": index_year,
            "coverage": coverage,
            "period_note": period_note,
            "provider_event_state": state,
            "regulation_periods": _regulation_periods(comp, payload),
        },
    }
    game.evidence = {"espn": espn_evidence}
    return game


def fetch_recent_results(
    date_yyyymmdd: str,
    fetch: FetchFn,
    league: League = League.NBA,
    observed_utc: Optional[Any] = None,
) -> List[Game]:
    """Completed games (state 'post') from a past-date scoreboard.

    Same ActualResult parsing as the summary path: regulation linescores are
    quarters or explicitly declared halves, and further elements aggregate
    into the single combined OT value with ``overtime_periods`` counting
    them. Records are returned even when period sums contradict the final
    score — the caller/smoke flags those with a ``reconciliation`` gap
    instead of folding a residual into OT.
    """
    payload = fetch(_scoreboard_url(date_yyyymmdd, league)) or {}
    observed_value = observed_utc() if callable(observed_utc) else observed_utc
    observed = to_utc_iso(parse_utc(observed_value or now_utc_iso()))
    results: List[Game] = []
    for event in payload.get("events", []):
        comp = _competition(event)
        if _state(comp) != "post":
            continue
        home, away = _home_away_competitors(comp)
        actual, period_note = _actual_for_history(home, away, comp, league)
        game = _game_from_event(event, comp, league)
        coverage = (
            "final_score_only"
            if actual.period_kind is PeriodKind.UNKNOWN
            else "reconciled_periods"
            if actual.reconciles()[0]
            else "unreconciled_periods"
        )
        game.evidence = {
            "espn": {
                "scoreboard": _scoreboard_context(event, comp, observed, []),
                "history": {
                    "coverage": coverage,
                    "period_note": period_note,
                    "result_period_parser_version": RESULT_PERIOD_PARSER_VERSION,
                },
            }
        }
        results.append(replace(game, actual=actual))
    return results


def fetch_standings(fetch: FetchFn) -> List[StandingsRow]:
    """Parse conference standings children into flat canonical rows.

    ``points_for``/``points_against`` are the CUMULATIVE season totals as
    reported by the payload — stored verbatim, never per-game averages and
    never divided by games played. ``position`` mirrors ESPN
    ``playoffSeed``; note that ESPN's playoffSeed is **not** guaranteed to
    equal the current conference rank (it is a playoff projection seed).
    ``games_played`` uses the provider stat when present, otherwise wins +
    losses (basketball has no draws). ``streak`` is signed: negative for a
    losing streak (ESPN ``L6`` -> -6). ``section`` keeps the conference
    child's name verbatim (e.g. ``"Eastern Conference"``). ``win_pct`` is
    the 0-1 ``winPercent`` value.

    Cumulative guard: ``points_for <= 3 * games_played`` is impossible for
    any real cumulative NBA total, so it raises ``ValueError`` — this
    catches non-cumulative values (single-game or per-period averages)
    being written into the cumulative field. The payload's ``pointsFor``
    stat is cumulative by schema, so the guard is a tripwire, not the
    primary defence.
    """
    payload = fetch(STANDINGS_URL) or {}
    rows: List[StandingsRow] = []
    for child in payload.get("children", []):
        entries = (child.get("standings") or {}).get("entries") or []
        for entry in entries:
            stats = {s.get("name"): s.get("value") for s in entry.get("stats") or []}
            team = entry.get("team") or {}
            if stats.get("wins") is None or stats.get("losses") is None:
                raise IdentityError(f"team {team.get('id')}: standings entry missing wins/losses")
            wins = int(stats["wins"])
            losses = int(stats["losses"])
            games_played = (
                int(stats["gamesPlayed"]) if stats.get("gamesPlayed") is not None else wins + losses
            )
            points_for = float(stats["pointsFor"]) if stats.get("pointsFor") is not None else None
            if (
                points_for is not None
                and games_played > 0
                and points_for <= PF_CUMULATIVE_MIN_FACTOR * games_played
            ):
                raise ValueError(
                    f"team {team.get('id')}: pointsFor {points_for} <= "
                    f"{PF_CUMULATIVE_MIN_FACTOR} x {games_played} games — value does "
                    "not look like a cumulative season total (per-game/period "
                    "average written into pointsFor?); quarantining standings"
                )
            seed = stats.get("playoffSeed")
            streak = stats.get("streak")
            rows.append(
                StandingsRow(
                    team_name=_canonical_name(str(team.get("displayName") or "")),
                    position=int(seed) if seed is not None else None,
                    games_played=games_played,
                    wins=wins,
                    losses=losses,
                    points_for=points_for,
                    points_against=(
                        float(stats["pointsAgainst"])
                        if stats.get("pointsAgainst") is not None
                        else None
                    ),
                    win_pct=(
                        float(stats["winPercent"]) if stats.get("winPercent") is not None else None
                    ),
                    streak=int(streak) if streak is not None else None,
                    section=str(child.get("name")) if child.get("name") else None,
                    source="espn",
                )
            )
    return rows


def smoke(date_yyyymmdd: str, fetch: FetchFn, league: League = League.NBA) -> Dict[str, Any]:
    """Coverage smoke over one slate date; reports ACTUAL coverage.

    Evidence label: ``LIVE_VERIFIED`` when ``fetch`` is a real
    ``http.HttpClient`` (an instance or a bound method such as
    ``HttpClient(...).get_json``, the documented production wiring —
    observed responses), otherwise ``FIXTURE_VERIFIED`` (fixture loader).
    Wrap HttpClient (rather than substituting an arbitrary callable) if a
    live caller needs its own transport, so the label stays honest.

    Per-event failures (missing predictor, malformed payload, identity
    error) are counted as incomplete and recorded as gaps instead of
    aborting the sweep; a final whose quarter sums contradict its final
    score keeps its record and is flagged with a ``reconciliation`` gap
    (no residual is folded into OT).
    """
    scoreboard_payload = fetch(_scoreboard_url(date_yyyymmdd, league)) or {}
    events = scoreboard_payload.get("events", [])
    games = [_game_from_event(event, _competition(event), league) for event in events]

    accumulator = CoverageAccumulator()
    bpi_complete = 0
    for game in games:
        try:
            prediction = fetch_bpi(game.game_id, fetch, game=game)
        except Exception as exc:  # keep sweeping; this event counts incomplete
            detail = f"{type(exc).__name__}: {exc}"
            for field in ("prob", "expected_scores", "total"):
                accumulator.observe(
                    field, False, detail=detail, game_id=game.game_id, provider="espn"
                )
            continue
        prob_ok = accumulator.observe(
            "prob",
            prediction.prob_home is not None and prediction.prob_away is not None,
            game_id=game.game_id,
            provider="espn",
        )
        scores_ok = accumulator.observe(
            "expected_scores",
            prediction.expected_home_score is not None
            and prediction.expected_away_score is not None,
            game_id=game.game_id,
            provider="espn",
        )
        total_ok = accumulator.observe(
            "total",
            prediction.expected_total is not None,
            game_id=game.game_id,
            provider="espn",
        )
        if prob_ok and scores_ok and total_ok:
            bpi_complete += 1

    finals = 0
    reconciled = 0
    for event in events:
        comp = _competition(event)
        if _state(comp) != "post":
            continue
        finals += 1
        game_id = str(event.get("id") or "")
        try:
            home, away = _home_away_competitors(comp)
            actual = _actual_from_competitors(
                home,
                away,
                source="espn",
                league=league,
                regulation_periods=_regulation_periods(comp),
            )
        except IdentityError as exc:
            accumulator.gaps.append(
                CoverageGap(
                    game_id=game_id,
                    required_field="actual_result",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="espn",
                    detail=str(exc),
                )
            )
            continue
        ok, reason = actual.reconciles()
        if ok:
            reconciled += 1
        else:
            accumulator.gaps.append(
                CoverageGap(
                    game_id=game_id,
                    required_field="reconciliation",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="espn",
                    detail=reason,
                )
            )

    book_markets = {market.value: 0 for market in (Market.MONEYLINE, Market.SPREAD, Market.TOTAL)}
    games_with_books = 0
    games_with_team_statistics = 0
    games_with_venue = 0
    for event, game in zip(events, games):
        snapshots = book_snapshots_from_event(event, game)
        if snapshots:
            games_with_books += 1
        for snapshot in snapshots:
            book_markets[snapshot.market.value] += 1
        comp = _competition(event)
        if any(c.get("statistics") for c in comp.get("competitors") or []):
            games_with_team_statistics += 1
        if comp.get("venue"):
            games_with_venue += 1

    fetch_target = getattr(fetch, "__self__", fetch)  # bound method -> client
    if isinstance(fetch_target, AcquisitionClient):
        wrapped = fetch_target.fetch
        fetch_target = getattr(wrapped, "__self__", wrapped)
    verification = "LIVE_VERIFIED" if isinstance(fetch_target, HttpClient) else "FIXTURE_VERIFIED"
    return {
        "fixtures_found": len(games),
        "bpi_complete": {"complete": bpi_complete, "total": len(games)},
        "results_reconciled": {
            "finals": finals,
            "reconciled": reconciled,
            "flagged": finals - reconciled,
        },
        "scoreboard_books": {
            "games_with_odds": games_with_books,
            "total_games": len(games),
            "market_snapshots": book_markets,
        },
        "scoreboard_context": {
            "games_with_team_statistics": games_with_team_statistics,
            "games_with_venue": games_with_venue,
            "total_games": len(games),
        },
        "verification": verification,
        "gaps": [
            {
                "game_id": gap.game_id,
                "required_field": gap.required_field,
                "reason": gap.reason.value,
                "detail": gap.detail,
            }
            for gap in accumulator.gaps
        ],
    }
