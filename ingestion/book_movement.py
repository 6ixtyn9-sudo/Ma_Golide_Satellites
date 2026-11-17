"""Descriptive observed book movement from immutable canonical slate ledgers.

This compares timestamped BookSnapshot rows only. Untimestamped scoreboard
open/close context is intentionally excluded; outputs are observations, not
closing-line value, forecast features, or evidence of value.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .identity import parse_utc

PRICE_FIELDS = (
    "home_price_decimal",
    "away_price_decimal",
    "draw_price_decimal",
    "over_price_decimal",
    "under_price_decimal",
)


def _load_ledgers(paths: Iterable[str | Path]) -> List[Dict[str, Any]]:
    payloads = []
    for path in paths:
        try:
            with Path(path).open(encoding="utf-8") as handle:
                payload = json.load(handle)
        except (OSError, ValueError):
            continue
        if isinstance(payload, dict) and isinstance(payload.get("book_snapshots"), list):
            payloads.append(payload)
    return payloads


def observed_book_movement(
    paths: Iterable[str | Path], *, league: Optional[str] = None, target_date: Optional[str] = None
) -> Dict[str, Any]:
    """Compare first/last admissible pre-match observations per market slot."""
    payloads = _load_ledgers(paths)
    eligible = [
        payload
        for payload in payloads
        if (league is None or str(payload.get("league", "")).upper() == league.upper())
        and (target_date is None or str(payload.get("target_date", "")).replace("-", "") == target_date.replace("-", ""))
        and payload.get("ledger_kind") in {"pregame", "freeze", "results"}
    ]
    observations: Dict[Tuple[str, str, str, str, str], Dict[str, Dict[str, Any]]] = {}
    for payload in eligible:
        games = {
            str(game.get("game_id")): game
            for game in payload.get("games", [])
            if isinstance(game, dict) and game.get("game_id")
        }
        for row in payload.get("book_snapshots", []):
            if not isinstance(row, dict) or row.get("pre_match") is not True:
                continue
            game_id = str(row.get("game_id") or "")
            game = games.get(game_id, {})
            scheduled = row.get("scheduled_utc") or game.get("scheduled_utc")
            observed = row.get("observed_utc")
            try:
                if not scheduled or not observed or parse_utc(observed) >= parse_utc(scheduled):
                    continue
            except (TypeError, ValueError):
                continue
            provider = str(row.get("provider") or "")
            book = str(row.get("book") or "")
            market = str(row.get("market") or "")
            period = str(row.get("period") or "")
            key = (game_id, provider, book, market, period)
            # Identical rows recur across immutable ledgers. Deduplicate by
            # full canonical row, not timestamp: two different prices can
            # share a provider timestamp and must not be silently overwritten.
            normalized_row = {field: value for field, value in row.items() if field != "frozen"}
            fingerprint = json.dumps(
                normalized_row, sort_keys=True, separators=(",", ":"), default=str
            )
            observations.setdefault(key, {})[fingerprint] = normalized_row

    movements = []
    same_timestamp_conflicts = 0
    for (game_id, provider, book, market, period), by_time in sorted(observations.items()):
        time_counts: Dict[str, int] = {}
        for row in by_time.values():
            stamp = str(row["observed_utc"])
            time_counts[stamp] = time_counts.get(stamp, 0) + 1
        if any(count > 1 for count in time_counts.values()):
            same_timestamp_conflicts += 1
            continue
        if len(time_counts) < 2:
            continue
        rows = sorted(by_time.values(), key=lambda row: parse_utc(str(row["observed_utc"])))
        first, last = rows[0], rows[-1]
        fields = ("line", *PRICE_FIELDS)
        deltas = {}
        for field in fields:
            old, new = first.get(field), last.get(field)
            if isinstance(old, (int, float)) and isinstance(new, (int, float)):
                deltas[field] = round(float(new) - float(old), 6)
        movements.append(
            {
                "game_id": game_id,
                "provider": provider,
                "book": book,
                "market": market,
                "period": period,
                "first_observed_utc": first["observed_utc"],
                "last_observed_utc": last["observed_utc"],
                "observations": len(by_time),
                "first": {field: first.get(field) for field in fields},
                "last": {field: last.get(field) for field in fields},
                "deltas": deltas,
            }
        )
    return {
        "schema_version": "magolide.observed_book_movement.v1",
        "league": league,
        "target_date": target_date,
        "ledgers_scanned": len(payloads),
        "ledgers_in_scope": len(eligible),
        "slots_with_multiple_timestamped_observations": len(movements),
        "slots_with_conflicting_values_at_same_timestamp": same_timestamp_conflicts,
        "movements": movements,
        "interpretation": "observed pre-match movement only; not CLV, a feature, or a value claim",
        "exclusions": ["untimestamped scoreboard open/close context", "post-match snapshots"],
    }
