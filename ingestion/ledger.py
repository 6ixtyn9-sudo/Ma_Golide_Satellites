"""Content-addressed, immutable acquisition and freeze ledgers."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any, Dict, Tuple
from zoneinfo import ZoneInfo

from .duckdb_store import DuckDBStore
from .models import League

_DEFAULT_TZ = {
    "NBA": "America/New_York",
    "EUROLEAGUE": "Europe/Paris",
    "CBA": "Asia/Shanghai",
    "WNBA": "America/New_York",
    "NCAA_M": "America/New_York",
    "NCAA_W": "America/New_York",
}


class LedgerConflictError(RuntimeError):
    """An immutable ledger path already contains different bytes."""


def _bytes(payload: Any) -> bytes:
    return (json.dumps(payload, indent=2, sort_keys=True, default=str) + "\n").encode()


def write_immutable_json(path: str, payload: Any) -> Tuple[str, bool]:
    """Create ``path`` once; an identical rerun is a no-op, a difference fails."""
    raw = _bytes(payload)
    target = Path(path)
    if target.exists():
        if target.read_bytes() != raw:
            raise LedgerConflictError(f"immutable ledger differs: {target}")
        return str(target), False
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        with target.open("xb") as handle:
            handle.write(raw)
    except FileExistsError:
        if target.read_bytes() != raw:
            raise LedgerConflictError(f"immutable ledger differs: {target}")
        return str(target), False
    return str(target), True


def write_versioned_ledger(directory: str, stem: str, payload: Any) -> Dict[str, Any]:
    """Write an immutable ledger whose content hash is part of its filename."""
    raw = _bytes(payload)
    digest = hashlib.sha256(raw).hexdigest()
    path = Path(directory) / f"{stem}_{digest[:16]}.json"
    written_path, created = write_immutable_json(str(path), payload)
    return {"path": written_path, "sha256": digest, "created": created}


def canonical_snapshot(
    store: DuckDBStore,
    league: League,
    target_date: str | None = None,
    tz_name: str | None = None,
) -> Dict[str, Any]:
    """Return one league-local slate's canonical facts in stable order.

    Limiting a dated ledger to its slate prevents every three-hour run from
    copying the league's entire history into another immutable artifact.
    """
    from datetime import datetime

    games = sorted(store.list_games(league=league), key=lambda game: game.game_id)
    normalized_date = target_date.replace("-", "") if target_date else None
    if normalized_date:
        if len(normalized_date) != 8 or not normalized_date.isdigit():
            raise ValueError(f"invalid target date: {target_date}")
        wanted = f"{normalized_date[:4]}-{normalized_date[4:6]}-{normalized_date[6:]}"
        timezone = ZoneInfo(tz_name or _DEFAULT_TZ[league.name])
        games = [
            game
            for game in games
            if datetime.fromisoformat(game.scheduled_utc.replace("Z", "+00:00"))
            .astimezone(timezone)
            .date()
            .isoformat()
            == wanted
        ]
    game_ids = {game.game_id for game in games}
    prediction_rows = sorted(
        (row for row in store.table_rows("predictions") if row["game_id"] in game_ids),
        key=lambda row: (row["game_id"], row["slot"]),
    )
    book_rows = sorted(
        (row for row in store.table_rows("book_snapshots") if row["game_id"] in game_ids),
        key=lambda row: (row["game_id"], row["slot"]),
    )
    return {
        "schema_version": "magolide.canonical.v1",
        "league": league.value,
        "games": [
            json.loads(row["payload_json"])
            for row in sorted(store.table_rows("games"), key=lambda row: row["game_id"])
            if row["game_id"] in game_ids
        ],
        "predictions": [
            json.loads(row["payload_json"])
            | {
                "frozen": bool(row["frozen"]),
            }
            for row in prediction_rows
        ],
        "book_snapshots": [
            json.loads(row["payload_json"])
            | {
                "frozen": bool(row["frozen"]),
            }
            for row in book_rows
        ],
    }


def write_canonical_ledger(
    store: DuckDBStore,
    league: League,
    target_date: str,
    kind: str,
    directory: str = "localdata/ledgers",
    tz_name: str | None = None,
) -> Dict[str, Any]:
    if kind not in {"pregame", "freeze", "results", "backfill"}:
        raise ValueError(f"unsupported ledger kind: {kind}")
    snapshot = canonical_snapshot(store, league, target_date, tz_name)
    snapshot["ledger_kind"] = kind
    snapshot["target_date"] = target_date
    stem = f"{kind}_{league.value}_{target_date}"
    return write_versioned_ledger(directory, stem, snapshot)
