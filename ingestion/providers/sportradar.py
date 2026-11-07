"""Sportradar Global Basketball CBA pre-match probabilities.

Evidence status: DOCUMENTATION_ONLY until an authenticated smoke test runs.
The base-package Competitions, Competition Seasons, and Season Probabilities
feeds are used; no undocumented endpoint or inferred probability is involved.
CBA is resolved as the unique men's ``CBA`` competition in category ``CHN``.
A prediction is joined to an iSportsAPI fixture only when home/away canonical
identities and the confirmed UTC start time match exactly and the join is
one-to-one. Unmatched or ambiguous events remain coverage gaps.

Probabilities are percentages in the documented ``2way`` market. Live or
removed markets are never relabelled as pre-match. Expected scores and totals
are absent from this feed and remain ``None``.
"""

from __future__ import annotations

from datetime import date as Date
import math
import os
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urlencode

from ..http import HttpClient
from ..identity import canonical_team_key, now_utc_iso, parse_utc, to_utc_iso
from ..models import Game, League, Prediction
from . import FetchFn, IdentityError

PROVIDER = "sportradar"
ENV_KEY = "SPORTRADAR_API_KEY"
ACCESS_ENV = "SPORTRADAR_ACCESS_LEVEL"
BASE_URL = "https://api.sportradar.com/basketball"
DOCUMENTATION_ONLY = "DOCUMENTATION_ONLY"
_PAGE_SIZE = 200
_MAX_PAGES = 100


def is_enabled() -> bool:
    return bool(os.environ.get(ENV_KEY, "").strip())


def _access_level() -> str:
    value = os.environ.get(ACCESS_ENV, "trial").strip().lower()
    if value not in {"trial", "production"}:
        raise ValueError(f"{ACCESS_ENV} must be 'trial' or 'production'")
    return value


def _root(access_level: Optional[str] = None) -> str:
    return f"{BASE_URL}/{access_level or _access_level()}/v2/en"


def _get(fetch: FetchFn, url: str) -> Dict[str, Any]:
    key = os.environ.get(ENV_KEY, "").strip()
    if not key:
        raise IdentityError(f"{ENV_KEY} is required for Sportradar CBA probabilities")
    payload = fetch(url, headers={"x-api-key": key, "Accept": "application/json"})
    if not isinstance(payload, dict):
        raise ValueError("Sportradar response must be a JSON object")
    return payload


def _cba_competition(fetch: FetchFn) -> Dict[str, Any]:
    payload = _get(fetch, f"{_root()}/competitions.json")
    rows = payload.get("competitions")
    if not isinstance(rows, list):
        raise ValueError("Sportradar competitions payload has no competitions list")
    matches = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        category = row.get("category") or {}
        if (
            str(row.get("name", "")).strip().casefold() == "cba"
            and str(category.get("country_code", "")).strip().upper() == "CHN"
            and str(row.get("gender", "")).strip().casefold() == "men"
            and row.get("id")
        ):
            matches.append(row)
    if len(matches) != 1:
        raise IdentityError(
            f"Sportradar CBA identity must be unique (found {len(matches)})"
        )
    return matches[0]


def _season_for_date(
    fetch: FetchFn, competition_id: str, target_date: str
) -> Dict[str, Any]:
    if len(target_date) == 8 and target_date.isdigit():
        target_date = f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"
    try:
        wanted = Date.fromisoformat(target_date)
    except ValueError as exc:
        raise ValueError("target_date must be YYYY-MM-DD") from exc
    url = f"{_root()}/competitions/{competition_id}/seasons.json"
    rows = _get(fetch, url).get("seasons")
    if not isinstance(rows, list):
        raise ValueError("Sportradar competition seasons payload has no seasons list")
    matches = []
    for row in rows:
        if not isinstance(row, dict) or row.get("disabled") is True:
            continue
        if row.get("competition_id") != competition_id or not row.get("id"):
            continue
        try:
            start = Date.fromisoformat(str(row["start_date"]))
            end = Date.fromisoformat(str(row["end_date"]))
        except (KeyError, ValueError):
            continue
        if start <= wanted <= end:
            matches.append(row)
    if len(matches) != 1:
        raise IdentityError(
            f"Sportradar CBA season for {target_date} must be unique "
            f"(found {len(matches)})"
        )
    return matches[0]


def _probability_rows(
    fetch: FetchFn, season_id: str
) -> Iterable[Tuple[Dict[str, Any], Optional[str]]]:
    start = 0
    for _ in range(_MAX_PAGES):
        query = urlencode({"start": start, "limit": _PAGE_SIZE})
        url = f"{_root()}/seasons/{season_id}/probabilities.json?{query}"
        payload = _get(fetch, url)
        rows = payload.get("sport_event_probabilities")
        if not isinstance(rows, list):
            raise ValueError(
                "Sportradar season probabilities payload has no "
                "sport_event_probabilities list"
            )
        for row in rows:
            if isinstance(row, dict):
                yield row, payload.get("generated_at")
        if len(rows) < _PAGE_SIZE:
            return
        start += len(rows)
    raise ValueError("Sportradar season probability pagination exceeded safety bound")


def _event_join(row: Dict[str, Any]) -> Optional[Tuple[str, Tuple[str, str, str]]]:
    event = row.get("sport_event") or {}
    event_id = str(event.get("id", "")).strip()
    if not event_id or event.get("start_time_confirmed") is not True:
        return None
    try:
        start = to_utc_iso(parse_utc(event["start_time"]))
    except (KeyError, TypeError, ValueError):
        return None
    teams: Dict[str, str] = {}
    for competitor in event.get("competitors") or []:
        if not isinstance(competitor, dict):
            continue
        qualifier = str(competitor.get("qualifier", "")).casefold()
        if qualifier in {"home", "away"} and competitor.get("name"):
            if qualifier in teams:
                return None
            teams[qualifier] = str(competitor["name"])
    if set(teams) != {"home", "away"}:
        return None
    key = (
        canonical_team_key(teams["home"]),
        canonical_team_key(teams["away"]),
        start,
    )
    return event_id, key


def _game_key(game: Game) -> Tuple[str, str, str]:
    return (
        canonical_team_key(game.home_name),
        canonical_team_key(game.away_name),
        to_utc_iso(parse_utc(game.scheduled_utc)),
    )


def _parse_probability(
    row: Dict[str, Any], game: Game, observed_utc: str, generated_at: Optional[str]
) -> Optional[Prediction]:
    event_id = str((row.get("sport_event") or {}).get("id", ""))
    for market in row.get("markets") or []:
        if not isinstance(market, dict):
            continue
        if str(market.get("name", "")).casefold() != "2way":
            continue
        if market.get("live") is True or market.get("removed") is True:
            continue
        outcomes: Dict[str, float] = {}
        valid = True
        for outcome in market.get("outcomes") or []:
            if not isinstance(outcome, dict):
                valid = False
                break
            name = str(outcome.get("name", ""))
            if name not in {"home_team_winner", "away_team_winner"} or name in outcomes:
                valid = False
                break
            try:
                value = float(outcome["probability"])
            except (KeyError, TypeError, ValueError):
                valid = False
                break
            if not math.isfinite(value) or not 0.0 <= value <= 100.0:
                valid = False
                break
            outcomes[name] = value / 100.0
        if not valid or set(outcomes) != {"home_team_winner", "away_team_winner"}:
            continue
        source_modified = market.get("last_updated") or generated_at
        if source_modified:
            try:
                source_modified = to_utc_iso(parse_utc(source_modified))
            except (TypeError, ValueError):
                source_modified = None
        return Prediction(
            game_id=game.game_id,
            provider=PROVIDER,
            observed_utc=observed_utc,
            source_modified_utc=source_modified,
            model="sportradar-pre-match-2way",
            model_version="global-basketball-v2",
            produced_by="Sportradar",
            prob_home=outcomes["home_team_winner"],
            prob_away=outcomes["away_team_winner"],
            source_revision=event_id,
            pre_tip=True,
        )
    return None


def fetch_cba_probabilities(
    target_date: str,
    games: Sequence[Game],
    fetch: Optional[FetchFn] = None,
    *,
    now_utc: Optional[str] = None,
) -> Dict[str, Prediction]:
    """Return one independent probability-only prediction per matched CBA game."""

    if not games:
        return {}
    if any(game.league is not League.CBA for game in games):
        raise ValueError("Sportradar adapter accepts CBA games only")
    fetcher = fetch or HttpClient().get_json
    observed = to_utc_iso(parse_utc(now_utc)) if now_utc else now_utc_iso()
    competition = _cba_competition(fetcher)
    season = _season_for_date(fetcher, str(competition["id"]), target_date)

    # Market last_updated is preferred; the page generated_at is the fallback.
    by_key: Dict[
        Tuple[str, str, str], List[Tuple[Dict[str, Any], Optional[str]]]
    ] = {}
    for row, generated_at in _probability_rows(fetcher, str(season["id"])):
        joined = _event_join(row)
        if joined is not None:
            _, key = joined
            by_key.setdefault(key, []).append((row, generated_at))

    matched: Dict[str, Prediction] = {}
    event_games: Dict[str, List[str]] = {}
    for game in games:
        rows = by_key.get(_game_key(game), [])
        event_ids = {
            str((row.get("sport_event") or {}).get("id", ""))
            for row, _ in rows
        }
        if len(rows) != 1 or len(event_ids) != 1:
            continue
        row, generated_at = rows[0]
        prediction = _parse_probability(row, game, observed, generated_at)
        if prediction is None:
            continue
        event_id = next(iter(event_ids))
        event_games.setdefault(event_id, []).append(game.game_id)
        matched[game.game_id] = prediction

    # A Sportradar event may never supply probabilities to two canonical games.
    return {
        game_id: prediction
        for game_id, prediction in matched.items()
        if len(event_games[prediction.source_revision or ""]) == 1
    }


def smoke(target_date: str, fetch: Optional[FetchFn] = None) -> Dict[str, Any]:
    """Read-only documentation-only provider smoke; no fixture join or writes."""

    if not is_enabled():
        return {"status": "UNAUTHORIZED", "reason": f"missing {ENV_KEY}"}
    fetcher = fetch or HttpClient().get_json
    competition = _cba_competition(fetcher)
    season = _season_for_date(fetcher, str(competition["id"]), target_date)
    count = sum(1 for _ in _probability_rows(fetcher, str(season["id"])))
    return {
        "status": "DOCUMENTATION_ONLY",
        "competition_id": competition["id"],
        "season_id": season["id"],
        "probability_events": count,
        "note": "Authenticated schema observed; fixture/market joins still require verification",
    }


__all__ = [
    "ACCESS_ENV", "DOCUMENTATION_ONLY", "ENV_KEY", "PROVIDER",
    "fetch_cba_probabilities", "is_enabled", "smoke",
]
