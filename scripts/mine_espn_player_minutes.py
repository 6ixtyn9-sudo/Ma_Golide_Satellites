#!/usr/bin/env python3
"""Export an offline, provenance-aware ESPN player boxscore research CSV.

Reads existing DuckDB evidence only; it performs no network requests and does
not alter canonical storage. Rows are descriptive post-game observations. The
``available_no_earlier_than_utc`` column is a conservative lower bound for
when a boxscore can become historical input (max of summary observation and
scheduled tip + result delay), never a same-game pre-tip feature timestamp.
"""

from __future__ import annotations

import argparse
import csv
from datetime import timedelta
import json
import math
from pathlib import Path
import sys
from typing import Any, Dict, Iterable, List, Optional

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.identity import parse_utc, to_utc_iso  # noqa: E402
from ingestion.models import Game, GameStatus, League  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE  # noqa: E402

ESPN_LEAGUES = ("NBA", "WNBA", "NCAA_M", "NCAA_W")
MAX_PLAYER_MINUTES = 100.0  # allow multi-OT games while rejecting seconds/unit corruption
CSV_FIELDS = (
    "league", "game_id", "scheduled_utc", "team_id", "side", "player_id",
    "player_name", "position", "category", "minutes", "points", "rebounds",
    "assists", "raw_stats_json", "stat_labels_json", "evidence_observed_utc",
    "evidence_source_modified_utc", "evidence_payload_hash",
    "available_no_earlier_than_utc", "source_schema_version",
)


def _number(value: Any) -> Optional[float]:
    try:
        parsed = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return parsed if math.isfinite(parsed) else None


def _minutes(value: Any) -> Optional[float]:
    if isinstance(value, str) and ":" in value:
        parts = value.strip().split(":")
        if len(parts) != 2:
            return None
        minutes, seconds = (_number(part) for part in parts)
        if minutes is None or seconds is None or not 0 <= seconds < 60:
            return None
        total = minutes + seconds / 60.0
        return total if 0 <= total <= MAX_PLAYER_MINUTES else None
    parsed = _number(value)
    return parsed if parsed is not None and 0 <= parsed <= MAX_PLAYER_MINUTES else None


def _position_text(athlete: Dict[str, Any]) -> Optional[str]:
    position = athlete.get("position") or athlete.get("positionName")
    if isinstance(position, dict):
        position = position.get("abbreviation") or position.get("displayName") or position.get("name")
    return str(position).strip() if position is not None and str(position).strip() else None


def player_boxscore_rows(game: Game, *, result_delay_hours: float = 4.0) -> List[Dict[str, Any]]:
    """Return identity-checked player/category rows from one stored Game."""
    if result_delay_hours < 0:
        raise ValueError("result_delay_hours must be non-negative")
    evidence = game.evidence if isinstance(game.evidence, dict) else {}
    espn = evidence.get("espn") if isinstance(evidence.get("espn"), dict) else {}
    summary = espn.get("summary") if isinstance(espn.get("summary"), dict) else {}
    boxscore = summary.get("player_boxscore")
    if not isinstance(boxscore, list):
        return []
    observed = summary.get("observed_utc")
    try:
        observed_dt = parse_utc(str(observed)) if observed else None
    except (TypeError, ValueError):
        observed_dt = None
    earliest_result = parse_utc(game.scheduled_utc) + timedelta(hours=result_delay_hours)
    earliest_available = max(earliest_result, observed_dt) if observed_dt else None
    rows: List[Dict[str, Any]] = []
    side_by_team = {
        str(game.home_team_id): "home",
        str(game.away_team_id): "away",
    }
    for team in boxscore:
        if not isinstance(team, dict):
            continue
        team_obj = team.get("team") if isinstance(team.get("team"), dict) else {}
        team_id = str(team_obj.get("id") or "").strip()
        side = side_by_team.get(team_id)
        if side is None:
            continue
        for category in team.get("statistics") or []:
            if not isinstance(category, dict):
                continue
            labels = category.get("labels") or category.get("names") or []
            if not isinstance(labels, list):
                continue
            normalized_labels = [str(label).strip() for label in labels]
            positions = {label.upper(): index for index, label in enumerate(normalized_labels)}
            metric_positions = {
                key: next((positions[label] for label in aliases if label in positions), None)
                for key, aliases in {
                    "minutes": ("MIN", "MINUTES"),
                    "points": ("PTS", "POINTS"),
                    "rebounds": ("REB", "REBOUNDS"),
                    "assists": ("AST", "ASSISTS"),
                }.items()
            }
            for athlete_row in category.get("athletes") or []:
                if not isinstance(athlete_row, dict):
                    continue
                athlete = athlete_row.get("athlete") if isinstance(athlete_row.get("athlete"), dict) else {}
                player_id = str(athlete.get("id") or "").strip()
                stats = athlete_row.get("stats")
                if not player_id or not isinstance(stats, list):
                    continue
                raw = {
                    label: stats[index]
                    for index, label in enumerate(normalized_labels)
                    if index < len(stats)
                }
                parsed: Dict[str, Optional[float]] = {}
                for key, position in metric_positions.items():
                    value = stats[position] if position is not None and position < len(stats) else None
                    parsed[key] = _minutes(value) if key == "minutes" else _number(value)
                rows.append(
                    {
                        "league": game.league.value,
                        "game_id": game.game_id,
                        "scheduled_utc": game.scheduled_utc,
                        "team_id": team_id,
                        "side": side,
                        "player_id": player_id,
                        "player_name": athlete.get("displayName") or athlete.get("shortName"),
                        "position": _position_text(athlete),
                        "category": category.get("name"),
                        **parsed,
                        "raw_stats_json": json.dumps(raw, sort_keys=True, separators=(",", ":"), default=str),
                        "stat_labels_json": json.dumps(normalized_labels, separators=(",", ":")),
                        "evidence_observed_utc": observed,
                        "evidence_source_modified_utc": summary.get("source_modified_utc"),
                        "evidence_payload_hash": summary.get("source_payload_hash"),
                        "available_no_earlier_than_utc": (
                            to_utc_iso(earliest_available) if earliest_available else None
                        ),
                        "source_schema_version": summary.get("schema_version"),
                    }
                )
    return rows


def mine_games(
    games: Iterable[Game], *, result_delay_hours: float = 4.0
) -> tuple[List[Dict[str, Any]], Dict[str, int]]:
    rows: List[Dict[str, Any]] = []
    counts = {"games_scanned": 0, "reconciled_finals": 0, "games_with_player_rows": 0}
    for game in games:
        counts["games_scanned"] += 1
        if (
            game.status is not GameStatus.FINAL
            or game.actual is None
            or not game.actual.reconciles()[0]
        ):
            continue
        counts["reconciled_finals"] += 1
        game_rows = player_boxscore_rows(game, result_delay_hours=result_delay_hours)
        if game_rows:
            counts["games_with_player_rows"] += 1
            rows.extend(game_rows)
    return rows, counts


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", default=DEFAULT_WAREHOUSE)
    parser.add_argument("--league", choices=ESPN_LEAGUES, default="NBA")
    parser.add_argument("--season", default=None, help="optional exact canonical season filter")
    parser.add_argument("--after", default=None, help="optional inclusive aware ISO-8601 tip bound")
    parser.add_argument("--before", default=None, help="optional exclusive aware ISO-8601 tip bound")
    parser.add_argument("--output", required=True, help="derived CSV output path")
    parser.add_argument("--result-delay-hours", type=float, default=4.0)
    args = parser.parse_args()
    warehouse = Path(args.warehouse)
    if not warehouse.exists():
        parser.error(f"warehouse does not exist: {warehouse}")
    if args.result_delay_hours < 0:
        parser.error("--result-delay-hours must be non-negative")
    try:
        after = parse_utc(args.after) if args.after else None
        before = parse_utc(args.before) if args.before else None
    except (TypeError, ValueError) as exc:
        parser.error(f"date bounds must be timezone-aware ISO-8601: {exc}")
    if after is not None and before is not None and after >= before:
        parser.error("--after must be earlier than --before")

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    counts = {
        "games_in_warehouse": 0,
        "games_scanned_after_filters": 0,
        "reconciled_finals": 0,
        "games_with_player_rows": 0,
        "rows_written": 0,
    }
    store = DuckDBStore(str(warehouse), persist=False)
    try:
        league = League[args.league]
        games = store.list_games(league=league)
        counts["games_in_warehouse"] = len(games)
        with output.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS, extrasaction="ignore")
            writer.writeheader()
            for game in games:
                tip = parse_utc(game.scheduled_utc)
                if args.season is not None and str(game.season) != args.season:
                    continue
                if after is not None and tip < after:
                    continue
                if before is not None and tip >= before:
                    continue
                counts["games_scanned_after_filters"] += 1
                if (
                    game.status is not GameStatus.FINAL
                    or game.actual is None
                    or not game.actual.reconciles()[0]
                ):
                    continue
                counts["reconciled_finals"] += 1
                rows = player_boxscore_rows(game, result_delay_hours=args.result_delay_hours)
                if rows:
                    counts["games_with_player_rows"] += 1
                    writer.writerows(rows)
                    counts["rows_written"] += len(rows)
    finally:
        store.close()

    print(
        json.dumps(
            {
                "league": args.league,
                "season_filter": args.season,
                "after_inclusive": args.after,
                "before_exclusive": args.before,
                "output": str(output),
                **counts,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
