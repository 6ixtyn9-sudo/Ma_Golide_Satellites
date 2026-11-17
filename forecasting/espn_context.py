"""Point-in-time ESPN player/availability features for forecast diagnostics.

The returned values are candidate features only: they are not consumed by the
probability model until a separate walk-forward challenger validates them.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import timedelta
import math
from typing import Any, Dict, Iterable, List, Optional

from ingestion.identity import parse_utc
from ingestion.models import Game

SCHEMA_VERSION = "magolide.espn_player_context.v1"
MAX_PLAYER_MINUTES = 100.0  # permissive for multi-OT games; reject corrupted unit/scale values


def _field_presence(summary: Dict[str, Any], field: str) -> Optional[bool]:
    """Read raw payload presence, not canonical keys materialized as empty lists."""
    presence = summary.get("field_presence")
    if isinstance(presence, dict) and isinstance(presence.get(field), bool):
        return presence[field]
    if field not in summary:
        return False
    value = summary.get(field)
    if value not in (None, [], {}):
        return True
    # Legacy v1 summaries materialized empty containers whether the source
    # field existed or not; preserve that distinction as unknown.
    return None


def _summary(game: Game) -> Dict[str, Any]:
    evidence = game.evidence if isinstance(game.evidence, dict) else {}
    espn = evidence.get("espn") if isinstance(evidence.get("espn"), dict) else {}
    summary = espn.get("summary") if isinstance(espn.get("summary"), dict) else {}
    return summary


def _finite(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def _minutes(value: Any) -> Optional[float]:
    """Parse ESPN minutes such as ``34:12`` or a numeric minute value."""
    if isinstance(value, str) and ":" in value:
        parts = value.strip().split(":")
        if len(parts) != 2:
            return None
        minutes, seconds = (_finite(part) for part in parts)
        if minutes is None or seconds is None or not 0 <= seconds < 60:
            return None
        total = minutes + seconds / 60.0
        return total if 0 <= total <= MAX_PLAYER_MINUTES else None
    parsed = _finite(value)
    return parsed if parsed is not None and 0 <= parsed <= MAX_PLAYER_MINUTES else None


def _player_rows(game: Game, team_id: str) -> Dict[str, Dict[str, Any]]:
    """Extract minutes and common box-score stats from one ESPN game summary."""
    result: Dict[str, Dict[str, Any]] = {}
    for team in _summary(game).get("player_boxscore") or []:
        if not isinstance(team, dict):
            continue
        team_obj = team.get("team") if isinstance(team.get("team"), dict) else {}
        if str(team_obj.get("id") or "") != team_id:
            continue
        for category in team.get("statistics") or []:
            if not isinstance(category, dict):
                continue
            labels = category.get("labels") or category.get("names") or []
            if not isinstance(labels, list):
                continue
            positions = {str(label).strip().upper(): index for index, label in enumerate(labels)}
            for entry in category.get("athletes") or []:
                if not isinstance(entry, dict):
                    continue
                athlete = entry.get("athlete") if isinstance(entry.get("athlete"), dict) else {}
                athlete_id = str(athlete.get("id") or "").strip()
                values = entry.get("stats") or []
                if not athlete_id or not isinstance(values, list):
                    continue
                row = result.setdefault(
                    athlete_id,
                    {
                        "athlete_id": athlete_id,
                        "name": athlete.get("displayName"),
                        "minutes": None,
                        "points": None,
                        "rebounds": None,
                        "assists": None,
                    },
                )
                for output, aliases in (
                    ("minutes", ("MIN", "MINUTES")),
                    ("points", ("PTS", "POINTS")),
                    ("rebounds", ("REB", "REBOUNDS")),
                    ("assists", ("AST", "ASSISTS")),
                ):
                    position = next((positions[label] for label in aliases if label in positions), None)
                    if position is None or position >= len(values):
                        continue
                    parsed = _minutes(values[position]) if output == "minutes" else _finite(values[position])
                    if parsed is not None:
                        row[output] = parsed
    return result


def _team_history_rows(
    team_id: str,
    history: Iterable[Game],
    *,
    as_of_utc: str,
    result_delay_hours: float,
    window_games: int,
) -> List[tuple[Game, Dict[str, Dict[str, Any]]]]:
    cutoff = parse_utc(as_of_utc)
    available = []
    for game in history:
        if game.actual is None or not game.actual.reconciles()[0]:
            continue
        scheduled = parse_utc(game.scheduled_utc)
        if scheduled + timedelta(hours=result_delay_hours) > cutoff:
            continue
        observed = _summary(game).get("observed_utc")
        if not observed:
            continue
        try:
            if parse_utc(str(observed)) > cutoff:
                continue
        except (TypeError, ValueError):
            continue
        rows = _player_rows(game, team_id)
        if rows:
            available.append((game, rows))
    available.sort(key=lambda item: (parse_utc(item[0].scheduled_utc), item[0].game_id))
    return available[-window_games:]


def _team_form_context(game: Game, team_id: str, *, as_of_utc: str) -> Dict[str, Any]:
    """Extract timestamped ESPN form/ATS context for shadow research only."""
    summary = _summary(game)
    observed = summary.get("observed_utc")
    try:
        visible = bool(observed) and parse_utc(str(observed)) <= parse_utc(as_of_utc) and parse_utc(
            str(observed)
        ) < parse_utc(game.scheduled_utc)
    except (TypeError, ValueError):
        visible = False
    form_presence = _field_presence(summary, "last_five_games")
    ats_presence = _field_presence(summary, "against_the_spread")
    if not visible:
        unavailable_state = "summary_not_point_in_time" if observed else "summary_unavailable"
        return {
            "status": unavailable_state,
            "observed_utc": observed,
            "recent_results_status": unavailable_state,
            "recent_results": [],
            "against_the_spread_status": unavailable_state,
            "against_the_spread": None,
        }

    raw_form = summary.get("last_five_games")
    form_present = form_presence is True and isinstance(raw_form, list)
    recent_results: List[Dict[str, Any]] = []
    if form_present:
        for team in raw_form:
            if not isinstance(team, dict):
                continue
            team_obj = team.get("team") if isinstance(team.get("team"), dict) else {}
            if str(team_obj.get("id") or "") != team_id:
                continue
            for event in team.get("events") or []:
                if not isinstance(event, dict):
                    continue
                event_id = str(event.get("id") or "")
                event_date = event.get("gameDate")
                if not event_date or event_id == game.game_id:
                    continue
                try:
                    event_time = parse_utc(str(event_date))
                    if event_time >= parse_utc(game.scheduled_utc) or event_time > parse_utc(as_of_utc):
                        continue
                except (TypeError, ValueError):
                    continue
                opponent = event.get("opponent") if isinstance(event.get("opponent"), dict) else {}
                recent_results.append(
                    {
                        "event_id": event_id,
                        "game_date_utc": str(event_date),
                        "result": event.get("gameResult"),
                        "score": event.get("score"),
                        "opponent_team_id": str(opponent.get("id") or "") or None,
                        "opponent_name": opponent.get("displayName"),
                    }
                )
    recent_results.sort(key=lambda row: (parse_utc(row["game_date_utc"]), row["event_id"]))

    ats_value: Optional[List[Dict[str, Any]]] = None
    if ats_presence is True and isinstance(summary.get("against_the_spread"), list):
        ats_value = []
        for team in summary["against_the_spread"]:
            if not isinstance(team, dict):
                continue
            team_obj = team.get("team") if isinstance(team.get("team"), dict) else {}
            if str(team_obj.get("id") or "") == team_id:
                records = team.get("records")
                if isinstance(records, list):
                    ats_value.extend(
                        {"summary": row.get("summary"), "type": row.get("type")}
                        for row in records
                        if isinstance(row, dict)
                    )
    form_status = (
        "available" if recent_results else
        "no_entries_reported" if form_presence is True else
        "field_absent" if form_presence is False else "field_presence_unknown"
    )
    ats_status = (
        "available" if ats_value else
        "no_entries_reported" if ats_presence is True else
        "field_absent" if ats_presence is False else "field_presence_unknown"
    )
    overall_status = (
        "available" if "available" in {form_status, ats_status} else
        "field_absent" if form_status == ats_status == "field_absent" else
        "field_presence_unknown" if "field_presence_unknown" in {form_status, ats_status} else
        "no_entries_reported"
    )
    return {
        "status": overall_status,
        "observed_utc": str(observed),
        "source_payload_hash": summary.get("source_payload_hash"),
        "recent_results_status": form_status,
        "recent_results": recent_results,
        "against_the_spread_status": ats_status,
        "against_the_spread": ats_value,
    }


def _injury_entries(game: Game, *, as_of_utc: str) -> tuple[str, List[Dict[str, Any]], Optional[str]]:
    evidence = game.evidence if isinstance(game.evidence, dict) else {}
    espn = evidence.get("espn") if isinstance(evidence.get("espn"), dict) else {}
    summary = _summary(game)
    sources: List[tuple[str, str, List[Dict[str, Any]], bool]] = []

    def visible_timestamp(value: Any) -> bool:
        if not value:
            return False
        try:
            return (
                parse_utc(str(value)) <= parse_utc(as_of_utc)
                and parse_utc(str(value)) < parse_utc(game.scheduled_utc)
            )
        except (TypeError, ValueError):
            return False

    def entry_date_not_future(value: Any) -> bool:
        if not value:
            return True
        try:
            entry_time = parse_utc(str(value))
        except (TypeError, ValueError):
            return True  # keep verbatim but do not invent a timestamp interpretation
        return entry_time <= parse_utc(as_of_utc) and entry_time < parse_utc(game.scheduled_utc)

    summary_observed = summary.get("observed_utc")
    injury_field_presence = _field_presence(summary, "injuries")
    if visible_timestamp(summary_observed) and injury_field_presence is True:
        summary_rows: List[Dict[str, Any]] = []
        summary_team_ids = set()
        raw = summary.get("injuries")
        if isinstance(raw, list):
            for team in raw:
                if not isinstance(team, dict):
                    continue
                team_obj = team.get("team") if isinstance(team.get("team"), dict) else {}
                team_id = str(team_obj.get("id") or "").strip()
                if team_id:
                    summary_team_ids.add(team_id)
                for item in team.get("injuries") or []:
                    if not isinstance(item, dict):
                        continue
                    athlete = item.get("athlete") if isinstance(item.get("athlete"), dict) else {}
                    athlete_id = str(athlete.get("id") or "").strip()
                    if not team_id or not athlete_id or not entry_date_not_future(item.get("date")):
                        continue
                    details = item.get("details") if isinstance(item.get("details"), dict) else {}
                    summary_rows.append(
                        {
                            "team_id": team_id,
                            "athlete_id": athlete_id,
                            "name": athlete.get("displayName"),
                            "status": item.get("status"),
                            "type": details.get("type"),
                            "date": item.get("date"),
                        }
                    )
            target_team_ids = {str(game.home_team_id), str(game.away_team_id)}
            sources.append(
                (
                    "summary",
                    str(summary_observed),
                    summary_rows,
                    target_team_ids <= summary_team_ids,
                )
            )

    league_snapshot = espn.get("league_injuries")
    if isinstance(league_snapshot, dict):
        observed = league_snapshot.get("observed_utc")
        teams = league_snapshot.get("teams")
        if visible_timestamp(observed) and isinstance(teams, dict):
            snapshot_rows: List[Dict[str, Any]] = []
            complete = league_snapshot.get("team_coverage_complete") is True
            for team_id in {str(game.home_team_id), str(game.away_team_id)}:
                entries = teams.get(team_id)
                if not isinstance(entries, list):
                    complete = False
                    continue
                for item in entries:
                    if (
                        not isinstance(item, dict)
                        or not item.get("athlete_id")
                        or not entry_date_not_future(item.get("date"))
                    ):
                        continue
                    snapshot_rows.append(
                        {
                            "team_id": team_id,
                            "athlete_id": str(item["athlete_id"]),
                            "name": item.get("name"),
                            "status": item.get("status"),
                            "type": item.get("type"),
                            "date": item.get("date"),
                        }
                    )
            sources.append(("league_endpoint", str(observed), snapshot_rows, complete))

    if not sources:
        if summary_observed and not visible_timestamp(summary_observed):
            return "summary_not_point_in_time", [], str(summary_observed)
        if summary_observed and injury_field_presence is False:
            return "field_absent", [], str(summary_observed)
        if summary_observed and injury_field_presence is None:
            return "field_presence_unknown", [], str(summary_observed)
        if summary_observed:
            return "no_point_in_time_injury_snapshot", [], str(summary_observed)
        return "no_point_in_time_injury_snapshot", [], None

    rows_by_player: Dict[tuple[str, str], Dict[str, Any]] = {}
    for source, source_observed, source_rows, _ in sources:
        source_time = parse_utc(source_observed)
        for row in source_rows:
            key = (row["team_id"], row["athlete_id"])
            existing = rows_by_player.get(key)
            observation = {
                "source": source,
                "observed_utc": source_observed,
                "status": row.get("status"),
                "type": row.get("type"),
                "date": row.get("date"),
            }
            if existing is None:
                rows_by_player[key] = {
                    **row,
                    "sources": [source],
                    "source_observations": [observation],
                    "status_source": source,
                    "status_observed_utc": source_observed,
                }
                continue

            prior_time = parse_utc(existing["status_observed_utc"])
            all_sources = sorted(set(existing.get("sources", [])) | {source})
            all_observations = [*existing.get("source_observations", []), observation]
            if source_time > prior_time:
                rows_by_player[key] = {
                    **existing,
                    **row,
                    "sources": all_sources,
                    "source_observations": all_observations,
                    "status_source": source,
                    "status_observed_utc": source_observed,
                }
            elif source_time == prior_time and any(
                row.get(field) != existing.get(field) for field in ("status", "type", "date")
            ):
                rows_by_player[key] = {
                    **existing,
                    "status": None,
                    "status_conflict": True,
                    "status_source": "same_timestamp_conflict",
                    "sources": all_sources,
                    "source_observations": all_observations,
                }
            else:
                existing["sources"] = all_sources
                existing["source_observations"] = all_observations
                rows_by_player[key] = existing
    rows = sorted(rows_by_player.values(), key=lambda row: (row["team_id"], row["athlete_id"]))
    observed = max((source[1] for source in sources), key=parse_utc)
    complete = any(source[3] for source in sources)
    if rows:
        status = "snapshot_present"
    elif complete:
        status = "no_entries_reported"
    else:
        status = "partial_snapshot_no_entries"
    return status, rows, observed


def build_espn_player_context(
    game: Game,
    historical_games: Iterable[Game],
    *,
    as_of_utc: str,
    result_delay_hours: float = 4.0,
    window_games: int = 5,
) -> Dict[str, Any]:
    """Build an auditable, point-in-time candidate feature record for a game.

    Historical boxscores are admitted only when their reconciled game result
    was available by ``as_of_utc`` and the evidence itself had been observed by
    then. Current-game injuries and summary form/ATS context are admitted only
    from a timestamped summary observed before both the forecast cutoff and tip;
    last-five events are additionally filtered to dates before the target tip.
    Missing injury entries are never interpreted as proof that a team is healthy.
    """
    if window_games < 1 or result_delay_hours < 0:
        raise ValueError("window_games must be positive and result_delay_hours non-negative")
    injury_state, injuries, injury_observed = _injury_entries(game, as_of_utc=as_of_utc)
    history = list(historical_games)
    team_features: Dict[str, Any] = {}
    team_ids = {
        "home": str(game.home_team_id),
        "away": str(game.away_team_id),
    }
    for side, team_id in team_ids.items():
        recent = _team_history_rows(
            team_id,
            history,
            as_of_utc=as_of_utc,
            result_delay_hours=result_delay_hours,
            window_games=window_games,
        )
        players: Dict[str, Dict[str, Any]] = defaultdict(
            lambda: {
                "athlete_id": "",
                "name": None,
                "appearances": 0,
                "minutes_sum": 0.0,
                "points_sum": 0.0,
                "points_games": 0,
                "rebounds_sum": 0.0,
                "rebounds_games": 0,
                "assists_sum": 0.0,
                "assists_games": 0,
            }
        )
        for prior_game, rows in recent:
            for athlete_id, stats in rows.items():
                item = players[athlete_id]
                item["athlete_id"] = athlete_id
                item["name"] = stats.get("name") or item["name"]
                minutes = stats.get("minutes")
                if minutes is not None:
                    item["appearances"] += 1
                    item["minutes_sum"] += minutes
                for key in ("points", "rebounds", "assists"):
                    value = stats.get(key)
                    if value is not None:
                        item[f"{key}_sum"] += value
                        item[f"{key}_games"] += 1
        player_rows = []
        for item in players.values():
            appearances = item["appearances"]
            if not appearances:
                continue
            player_rows.append(
                {
                    "athlete_id": item["athlete_id"],
                    "name": item["name"],
                    "appearances": appearances,
                    "mean_minutes": round(item["minutes_sum"] / appearances, 3),
                    "mean_points": round(item["points_sum"] / item["points_games"], 3)
                    if item["points_games"]
                    else None,
                    "mean_rebounds": round(item["rebounds_sum"] / item["rebounds_games"], 3)
                    if item["rebounds_games"]
                    else None,
                    "mean_assists": round(item["assists_sum"] / item["assists_games"], 3)
                    if item["assists_games"]
                    else None,
                }
            )
        player_rows.sort(key=lambda row: (-row["mean_minutes"], row["athlete_id"]))
        out_ids = {
            row["athlete_id"]
            for row in injuries
            if row["team_id"] == team_id
            and str(row.get("status") or "").strip().lower() in {"out", "out for season", "inactive"}
        }
        rotation_minutes = {
            row["athlete_id"]: row["mean_minutes"]
            for row in player_rows
        }
        team_features[side] = {
            "team_id": team_id,
            "historical_boxscore_games": len(recent),
            "window_games": window_games,
            "players_with_minutes": len(player_rows),
            "recent_players": player_rows,
            "confirmed_out_player_ids": sorted(out_ids),
            "recent_minutes_for_reported_out_players": round(
                sum(rotation_minutes.get(athlete_id, 0.0) for athlete_id in out_ids), 3
            ),
            "minutes_feature_status": "partial" if recent else "unavailable",
            "recent_form_candidate": _team_form_context(game, team_id, as_of_utc=as_of_utc),
        }
    return {
        "schema_version": SCHEMA_VERSION,
        "game_id": game.game_id,
        "league": game.league.value,
        "as_of_utc": as_of_utc,
        "feature_role": "shadow_candidate_only_not_used_to_change_probability",
        "injury_snapshot_status": injury_state,
        "injury_snapshot_observed_utc": injury_observed,
        "reported_injury_entries": injuries,
        "teams": team_features,
    }
