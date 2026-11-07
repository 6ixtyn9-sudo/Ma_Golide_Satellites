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

import re
from dataclasses import replace
from typing import Any, Dict, List, Optional, Tuple

from ..http import HttpClient
from ..identity import canonical_team_name, normalize_team_name, now_utc_iso, parse_utc, to_utc_iso
from ..leagues import period_kind
from ..models import (
    ActualResult,
    CoverageGap,
    CoverageState,
    Game,
    GameStatus,
    League,
    PeriodKind,
    Prediction,
    StandingsRow,
)
from . import CoverageAccumulator, FetchFn, IdentityError

SCOREBOARD_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/"
    "scoreboard?dates={yyyymmdd}&limit=100"
)
SUMMARY_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/"
    "summary?event={event_id}"
)
STANDINGS_URL = "https://site.api.espn.com/apis/v2/sports/basketball/nba/standings"
BPI_URL = (
    "https://sports.core.api.espn.com/v2/sports/basketball/leagues/nba/"
    "events/{event_id}/competitions/{event_id}/predictor"
)

LEAGUE_SLUGS = {
    League.NBA: "nba",
    League.WNBA: "wnba",  # existing predictor-only compatibility
    League.NCAA_M: "mens-college-basketball",
    League.NCAA_W: "womens-college-basketball",
}


def _league_slug(league: League) -> str:
    try:
        return LEAGUE_SLUGS[league]
    except KeyError:
        raise ValueError(f"ESPN acquisition is not enabled for {league.value}") from None


def _scoreboard_url(date_yyyymmdd: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return SCOREBOARD_URL.format(yyyymmdd=date_yyyymmdd)
    # NCAA Division I slates exceed the default page size; groups=50 selects
    # Division I and limit=500 prevents silent truncation.
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
        f"scoreboard?dates={date_yyyymmdd}&groups=50&limit=500"
    )


def _summary_url(event_id: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return SUMMARY_URL.format(event_id=event_id)
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
        f"summary?event={event_id}"
    )


def _bpi_url(event_id: str, league: League) -> str:
    slug = _league_slug(league)
    if league is League.NBA:
        return BPI_URL.format(event_id=event_id)
    return (
        f"https://sports.core.api.espn.com/v2/sports/basketball/leagues/{slug}/"
        f"events/{event_id}/competitions/{event_id}/predictor"
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
            "competitors missing homeAway markers; refusing to infer sides "
            "from array order"
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
) -> ActualResult:
    """Build an ActualResult from native competitor linescores.

    NBA uses four quarters; NCAA uses two halves. Every further element is
    one overtime period, summed into the canonical combined OT pair. Missing
    native periods, mismatched lengths, and unparseable values quarantine the
    result rather than relabelling or backfilling it.
    """
    home_values = _linescore_values(home)
    away_values = _linescore_values(away)
    native_count = 2 if period_kind(league) is PeriodKind.HALVES else 4
    period_label = "H1/H2" if native_count == 2 else "Q1..Q4"
    if len(home_values) < native_count or len(away_values) < native_count:
        raise IdentityError(
            f"linescores missing or incomplete (need at least {period_label})"
        )
    if len(home_values) != len(away_values):
        raise IdentityError(
            f"mismatched linescore lengths: home {len(home_values)} vs away {len(away_values)}"
        )
    home_ot = sum(home_values[native_count:]) if len(home_values) > native_count else None
    away_ot = sum(away_values[native_count:]) if len(away_values) > native_count else None
    return ActualResult(
        period_kind=period_kind(league),
        home_quarters=home_values[:native_count],
        away_quarters=away_values[:native_count],
        home_overtime=home_ot,
        away_overtime=away_ot,
        home_final=_final_score(home),
        away_final=_final_score(away),
        overtime_periods=max(len(home_values), len(away_values)) - native_count,
        source=source,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def discover_games(
    date_yyyymmdd: str, fetch: FetchFn, league: League = League.NBA
) -> List[Game]:
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
        _game_from_event(event, _competition(event), league)
        for event in payload.get("events", [])
    ]


def _stats_by_name(statistics: List[Dict[str, Any]]) -> Dict[str, Any]:
    return {s.get("name"): s.get("value") for s in statistics or []}


def fetch_bpi(
    event_id: str,
    fetch: FetchFn,
    game: Optional[Game] = None,
    now_utc: Optional[str] = None,
    league: League = League.NBA,
) -> Prediction:
    """Fetch the ESPN BPI predictor and normalise it to a canonical Prediction.

    Orientation: exactly one predictor side (``homeTeam``/``awayTeam``)
    carries ``statistics``; its team id (explicit or extracted from the
    ``$ref`` URL) is reconciled against the event competitors. When
    ``game`` is supplied its ``provider_ids["espn"]`` is the event id and
    its ``home_team_id``/``away_team_id`` (the scoreboard competitor ids)
    decide which canonical side owns the statistics — the payload's own
    ``homeTeam``/``awayTeam`` keys are NOT trusted for orientation there.
    Without a game the payload keys orient the sides (the ``name`` field,
    "Away at Home", documents the same order but carries no id→name table,
    so it cannot participate in id validation) and both team refs must
    resolve and be distinct. A stats-side id matching neither competitor,
    or identical/missing refs, raise ``IdentityError`` — never a swap.

    Mapping: ``gameProjection`` = the stats side's win probability
    (payload 0-100 -> stored 0-1), ``teamChanceLoss`` = the opponent's,
    ``teamExpectedPts``/``oppExpectedPts`` = that side's/opponent's
    expected scores, ``teamPredPtDiff`` = side minus opponent (flipped to
    home-minus-away when the stats side is away).

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
    if len(stats_sides) > 1:
        raise IdentityError(
            f"event {event_id}: statistics present on both predictor sides; cannot orient"
        )
    stats_side = stats_sides[0]
    side_team_id = _competitor_team_id(stats_side)

    home_ref = _team_ref_id(home_side.get("team") or {})
    away_ref = _team_ref_id(away_side.get("team") or {})
    if home_ref is not None and away_ref is not None and home_ref == away_ref:
        raise IdentityError(
            f"event {event_id}: predictor homeTeam and awayTeam both reference "
            f"team {home_ref}; sides cannot be reconciled"
        )

    if (
        game is not None
        and str(game.home_team_id)
        and str(game.away_team_id)
    ):
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
            raise ValueError(
                f"event {event_id}: {name} {value} outside the 0-100 percent range"
            )
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
        observed_utc=(
            to_utc_iso(parse_utc(now_utc)) if now_utc else now_utc_iso()
        ),
        source_modified_utc=(
            to_utc_iso(parse_utc(last_modified)) if last_modified else None
        ),
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
            f"event {event_id}: summary status is not final (state={state!r}); "
            "result quarantined"
        )
    home, away = _home_away_competitors(comp)
    actual = _actual_from_competitors(home, away, source="espn", league=league)
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


def fetch_recent_results(
    date_yyyymmdd: str, fetch: FetchFn, league: League = League.NBA
) -> List[Game]:
    """Completed games (state 'post') from a past-date scoreboard.

    Same ActualResult parsing as the summary path: exactly 4 linescore
    elements are Q1..Q4, further elements aggregate into the single
    combined OT value with ``overtime_periods`` counting them. Records are
    returned even when quarter sums contradict the final score — the
    caller/smoke flags those with a ``reconciliation`` gap instead of
    folding a residual into OT.
    """
    payload = fetch(_scoreboard_url(date_yyyymmdd, league)) or {}
    results: List[Game] = []
    for event in payload.get("events", []):
        comp = _competition(event)
        if _state(comp) != "post":
            continue
        home, away = _home_away_competitors(comp)
        actual = _actual_from_competitors(home, away, source="espn", league=league)
        results.append(replace(_game_from_event(event, comp, league), actual=actual))
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
                raise IdentityError(
                    f"team {team.get('id')}: standings entry missing wins/losses"
                )
            wins = int(stats["wins"])
            losses = int(stats["losses"])
            games_played = (
                int(stats["gamesPlayed"])
                if stats.get("gamesPlayed") is not None
                else wins + losses
            )
            points_for = (
                float(stats["pointsFor"]) if stats.get("pointsFor") is not None else None
            )
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
                        float(stats["winPercent"])
                        if stats.get("winPercent") is not None
                        else None
                    ),
                    streak=int(streak) if streak is not None else None,
                    section=str(child.get("name")) if child.get("name") else None,
                    source="espn",
                )
            )
    return rows


def smoke(
    date_yyyymmdd: str, fetch: FetchFn, league: League = League.NBA
) -> Dict[str, Any]:
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
            actual = _actual_from_competitors(home, away, source="espn", league=league)
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

    fetch_target = getattr(fetch, "__self__", fetch)  # bound method -> client
    verification = (
        "LIVE_VERIFIED" if isinstance(fetch_target, HttpClient) else "FIXTURE_VERIFIED"
    )
    return {
        "fixtures_found": len(games),
        "bpi_complete": {"complete": bpi_complete, "total": len(games)},
        "results_reconciled": {
            "finals": finals,
            "reconciled": reconciled,
            "flagged": finals - reconciled,
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
