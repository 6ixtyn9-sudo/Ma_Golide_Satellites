"""DuckDB/CSV-first canonical persistence.

DuckDB is the operational system of record.  Google Sheets is deliberately not
part of this implementation; the old Sheets store remains only as an explicit
legacy migration adapter.  Every mutation keeps the same natural keys and
freeze semantics as ``InMemoryStore`` and is mirrored transactionally to a
single-file DuckDB database.

Tables are intentionally analytics-friendly.  Large variable prediction/result
payloads use JSON, while fixture, book, timing and freeze fields are typed SQL
columns.  ``export_csv`` produces gzip CSV evidence for bounded Actions caches
and downstream research without an external application database.
"""

from __future__ import annotations

import csv
import gzip
import json
import os
from pathlib import Path
from typing import Any, Dict, Iterable, List

import duckdb

from .models import BookSnapshot, Game, League, Prediction, ProviderHealth
from .store import (
    InMemoryStore,
    _book_from_dict,
    _book_slot_key,
    _book_to_dict,
    _game_from_dict,
    _game_to_dict,
    _health_from_dict,
    _health_to_dict,
    _pred_slot_key,
    _prediction_from_dict,
    _prediction_to_dict,
)

SCHEMA_VERSION = "1"
EXPORT_TABLES = (
    "games",
    "predictions",
    "book_snapshots",
    "team_aliases",
    "provider_health",
    "pipeline_runs",
)

_SCHEMA = """
CREATE TABLE IF NOT EXISTS schema_meta (
    key VARCHAR PRIMARY KEY,
    value VARCHAR NOT NULL
);
CREATE TABLE IF NOT EXISTS games (
    game_id VARCHAR PRIMARY KEY,
    league VARCHAR NOT NULL,
    season VARCHAR,
    home_team_id VARCHAR NOT NULL,
    away_team_id VARCHAR NOT NULL,
    home_name VARCHAR NOT NULL,
    away_name VARCHAR NOT NULL,
    scheduled_utc VARCHAR NOT NULL,
    status VARCHAR NOT NULL,
    phase VARCHAR,
    provider_ids_json VARCHAR NOT NULL,
    actual_json VARCHAR,
    rescheduled_from VARCHAR,
    payload_json VARCHAR NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS predictions (
    game_id VARCHAR NOT NULL,
    provider VARCHAR NOT NULL,
    slot VARCHAR NOT NULL,
    captured_at VARCHAR NOT NULL,
    model_version VARCHAR,
    pre_tip BOOLEAN NOT NULL,
    frozen BOOLEAN NOT NULL DEFAULT FALSE,
    payload_json VARCHAR NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (game_id, slot)
);
CREATE TABLE IF NOT EXISTS book_snapshots (
    game_id VARCHAR NOT NULL,
    provider VARCHAR NOT NULL,
    book VARCHAR NOT NULL,
    market VARCHAR NOT NULL,
    period VARCHAR NOT NULL,
    slot VARCHAR NOT NULL,
    observed_utc VARCHAR NOT NULL,
    source_modified_utc VARCHAR,
    line DOUBLE,
    over_price_decimal DOUBLE,
    under_price_decimal DOUBLE,
    home_price_decimal DOUBLE,
    away_price_decimal DOUBLE,
    pre_match BOOLEAN NOT NULL,
    main_line BOOLEAN NOT NULL,
    payload_hash VARCHAR,
    frozen BOOLEAN NOT NULL DEFAULT FALSE,
    payload_json VARCHAR NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (game_id, slot)
);
CREATE TABLE IF NOT EXISTS team_aliases (
    canonical VARCHAR PRIMARY KEY,
    variants_json VARCHAR NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS provider_health (
    provider VARCHAR NOT NULL,
    league VARCHAR NOT NULL,
    last_attempt_utc VARCHAR,
    last_success_utc VARCHAR,
    state VARCHAR NOT NULL,
    detail VARCHAR,
    coverage_json VARCHAR NOT NULL,
    quota_json VARCHAR NOT NULL,
    payload_json VARCHAR NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (provider, league)
);
CREATE TABLE IF NOT EXISTS pipeline_runs (
    run_id VARCHAR NOT NULL,
    command VARCHAR NOT NULL,
    target_date VARCHAR,
    generated_utc VARCHAR NOT NULL,
    report_json VARCHAR NOT NULL,
    PRIMARY KEY (run_id, command)
);
CREATE INDEX IF NOT EXISTS idx_games_league_schedule
    ON games(league, scheduled_utc);
CREATE INDEX IF NOT EXISTS idx_predictions_game_provider
    ON predictions(game_id, provider, pre_tip);
CREATE INDEX IF NOT EXISTS idx_books_game_market_period
    ON book_snapshots(game_id, market, period, pre_match);
CREATE INDEX IF NOT EXISTS idx_runs_date_command
    ON pipeline_runs(target_date, command, generated_utc);
"""


class DuckDBStore(InMemoryStore):
    """Persistent canonical store backed by one DuckDB file.

    ``persist=False`` is a true dry run: an existing database is opened
    read-only, a missing database is replaced by an empty in-memory database,
    and no canonical mutation is written.  Reports/ledgers may still be written
    by the caller as run artifacts.
    """

    def __init__(self, path: str = "localdata/warehouse.duckdb", persist: bool = True):
        super().__init__()
        self.path = str(path)
        self.persist = bool(persist)
        self._connection_path = self.path
        read_only = False

        if self.path != ":memory:":
            db_path = Path(self.path).expanduser()
            self.path = str(db_path)
            if self.persist:
                db_path.parent.mkdir(parents=True, exist_ok=True)
            elif not db_path.exists():
                # DuckDB would otherwise create a file merely by opening it.
                self._connection_path = ":memory:"
            else:
                read_only = True

        self._con = duckdb.connect(self._connection_path, read_only=read_only)
        if not read_only:
            self._con.execute(_SCHEMA)
            self._con.execute(
                "INSERT OR REPLACE INTO schema_meta(key, value) "
                "VALUES ('schema_version', ?)",
                [SCHEMA_VERSION],
            )
        self._load()

    def close(self) -> None:
        self._con.close()

    def __enter__(self) -> "DuckDBStore":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    # -- loading ------------------------------------------------------------

    def _load(self) -> None:
        for row in self._con.execute(
            "SELECT game_id, payload_json FROM games"
        ).fetchall():
            self.games[row[0]] = _game_from_dict(json.loads(row[1]))

        for game_id, slot, frozen, payload in self._con.execute(
            "SELECT game_id, slot, frozen, payload_json FROM predictions"
        ).fetchall():
            prediction = _prediction_from_dict(json.loads(payload))
            self.predictions.setdefault(game_id, {})[slot] = prediction
            provider_freezes = self.frozen_predictions.setdefault(game_id, {})
            provider_freezes.setdefault(prediction.provider, False)
            if prediction.pre_tip and frozen:
                provider_freezes[prediction.provider] = True

        for game_id, slot, frozen, payload in self._con.execute(
            "SELECT game_id, slot, frozen, payload_json FROM book_snapshots"
        ).fetchall():
            snapshot = _book_from_dict(json.loads(payload))
            self.books.setdefault(game_id, {})[slot] = snapshot
            self.frozen_books.setdefault(game_id, {})[slot] = bool(frozen)

        for canonical, variants in self._con.execute(
            "SELECT canonical, variants_json FROM team_aliases"
        ).fetchall():
            self.team_aliases[canonical] = list(json.loads(variants))

        for provider, league, payload in self._con.execute(
            "SELECT provider, league, payload_json FROM provider_health"
        ).fetchall():
            league_enum = League(league)
            self.health[(provider, league_enum)] = _health_from_dict(
                json.loads(payload)
            )

    # -- small SQL helpers ---------------------------------------------------

    def _replace(
        self,
        table: str,
        where: str,
        key_values: Iterable[Any],
        columns: List[str],
        values: List[Any],
    ) -> None:
        if not self.persist:
            return
        placeholders = ", ".join("?" for _ in columns)
        self._con.execute("BEGIN TRANSACTION")
        try:
            self._con.execute(f"DELETE FROM {table} WHERE {where}", list(key_values))
            self._con.execute(
                f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})",
                values,
            )
            self._con.execute("COMMIT")
        except Exception:
            self._con.execute("ROLLBACK")
            raise

    @staticmethod
    def _json(value: Any) -> str:
        return json.dumps(value, sort_keys=True, default=str, separators=(",", ":"))

    # -- canonical mutations -------------------------------------------------

    def upsert_game(self, game: Game) -> None:
        with self._lock:
            super().upsert_game(game)
            payload = _game_to_dict(game)
            self._replace(
                "games",
                "game_id = ?",
                [game.game_id],
                [
                    "game_id",
                    "league",
                    "season",
                    "home_team_id",
                    "away_team_id",
                    "home_name",
                    "away_name",
                    "scheduled_utc",
                    "status",
                    "phase",
                    "provider_ids_json",
                    "actual_json",
                    "rescheduled_from",
                    "payload_json",
                ],
                [
                    game.game_id,
                    game.league.value,
                    game.season,
                    game.home_team_id,
                    game.away_team_id,
                    game.home_name,
                    game.away_name,
                    game.scheduled_utc,
                    game.status.value,
                    game.phase,
                    self._json(game.provider_ids),
                    self._json(payload["actual"]) if payload.get("actual") else None,
                    game.rescheduled_from,
                    self._json(payload),
                ],
            )

    def upsert_prediction(self, prediction: Prediction, frozen: bool = False) -> None:
        with self._lock:
            super().upsert_prediction(prediction, frozen)
            slot = _pred_slot_key(prediction)
            is_frozen = bool(
                prediction.pre_tip
                and self.is_prediction_frozen(prediction.game_id, prediction.provider)
            )
            self._replace(
                "predictions",
                "game_id = ? AND slot = ?",
                [prediction.game_id, slot],
                [
                    "game_id",
                    "provider",
                    "slot",
                    "captured_at",
                    "model_version",
                    "pre_tip",
                    "frozen",
                    "payload_json",
                ],
                [
                    prediction.game_id,
                    prediction.provider,
                    slot,
                    prediction.observed_utc,
                    prediction.model_version,
                    prediction.pre_tip,
                    is_frozen,
                    self._json(_prediction_to_dict(prediction)),
                ],
            )

    def upsert_book_snapshot(
        self, snapshot: BookSnapshot, frozen: bool = False
    ) -> None:
        with self._lock:
            super().upsert_book_snapshot(snapshot, frozen)
            slot = _book_slot_key(snapshot)
            is_frozen = self.is_book_snapshot_frozen(snapshot.game_id, snapshot)
            self._replace(
                "book_snapshots",
                "game_id = ? AND slot = ?",
                [snapshot.game_id, slot],
                [
                    "game_id",
                    "provider",
                    "book",
                    "market",
                    "period",
                    "slot",
                    "observed_utc",
                    "source_modified_utc",
                    "line",
                    "over_price_decimal",
                    "under_price_decimal",
                    "home_price_decimal",
                    "away_price_decimal",
                    "pre_match",
                    "main_line",
                    "payload_hash",
                    "frozen",
                    "payload_json",
                ],
                [
                    snapshot.game_id,
                    snapshot.provider,
                    snapshot.book,
                    snapshot.market.value,
                    snapshot.period.value,
                    slot,
                    snapshot.observed_utc,
                    snapshot.source_modified_utc,
                    snapshot.line,
                    snapshot.over_price_decimal,
                    snapshot.under_price_decimal,
                    snapshot.home_price_decimal,
                    snapshot.away_price_decimal,
                    snapshot.pre_match,
                    snapshot.main_line,
                    snapshot.payload_hash,
                    is_frozen,
                    self._json(_book_to_dict(snapshot)),
                ],
            )

    def upsert_team_alias(self, canonical: str, variants: List[str]) -> None:
        with self._lock:
            super().upsert_team_alias(canonical, variants)
            self._replace(
                "team_aliases",
                "canonical = ?",
                [canonical],
                ["canonical", "variants_json"],
                [canonical, self._json(variants)],
            )

    def upsert_provider_health(self, health: ProviderHealth) -> None:
        with self._lock:
            super().upsert_provider_health(health)
            payload = _health_to_dict(health)
            self._replace(
                "provider_health",
                "provider = ? AND league = ?",
                [health.provider, health.league.value],
                [
                    "provider",
                    "league",
                    "last_attempt_utc",
                    "last_success_utc",
                    "state",
                    "detail",
                    "coverage_json",
                    "quota_json",
                    "payload_json",
                ],
                [
                    health.provider,
                    health.league.value,
                    health.last_attempt_utc,
                    health.last_success_utc,
                    health.state.value,
                    health.detail,
                    self._json(health.coverage),
                    self._json(health.quota),
                    self._json(payload),
                ],
            )

    # -- operations / research ----------------------------------------------

    def record_run(
        self, run_id: str, command: str, target_date: str, report: dict
    ) -> None:
        """Persist a redacted run report; no provider credentials are accepted."""
        if not self.persist:
            return
        generated = str(report.get("generated_utc") or report.get("finished_utc") or "")
        with self._lock:
            self._replace(
                "pipeline_runs",
                "run_id = ? AND command = ?",
                [run_id, command],
                ["run_id", "command", "target_date", "generated_utc", "report_json"],
                [run_id, command, target_date, generated, self._json(report)],
            )

    def prune_pipeline_runs(self, older_than_utc: str) -> int:
        """Prune bounded operational telemetry; immutable ledgers are untouched."""
        if not self.persist:
            return 0
        with self._lock:
            count = self._con.execute(
                "SELECT count(*) FROM pipeline_runs "
                "WHERE generated_utc <> '' AND generated_utc < ?",
                [older_than_utc],
            ).fetchone()[0]
            if not count:
                return 0
            self._con.execute("BEGIN TRANSACTION")
            try:
                self._con.execute(
                    "DELETE FROM pipeline_runs "
                    "WHERE generated_utc <> '' AND generated_utc < ?",
                    [older_than_utc],
                )
                self._con.execute("COMMIT")
            except Exception:
                self._con.execute("ROLLBACK")
                raise
            return int(count)

    def table_rows(self, table: str) -> List[Dict[str, Any]]:
        """Return a whitelisted table for direct reads, exports, and ledgers."""
        if table not in EXPORT_TABLES:
            raise ValueError(f"unsupported canonical table: {table}")
        cursor = self._con.execute(f"SELECT * FROM {table}")
        names = [item[0] for item in cursor.description]
        return [dict(zip(names, row)) for row in cursor.fetchall()]

    def counts(self) -> Dict[str, int]:
        return {
            table: int(self._con.execute(f"SELECT count(*) FROM {table}").fetchone()[0])
            for table in EXPORT_TABLES
        }

    def restore_csv(self, directory: str = "localdata/csv") -> Dict[str, int]:
        """Restore an empty warehouse from canonical CSV.GZ materializations.

        Complete source payloads drive dataclass deserialization, so enum and
        nested-result types survive the round trip. Refusing a non-empty target
        prevents recovery from merging with or bypassing an active freeze.
        """
        if not self.persist:
            raise RuntimeError("CSV recovery requires a writable empty warehouse")
        if any(self.counts().values()):
            raise RuntimeError("CSV recovery requires an empty warehouse")
        source = Path(directory)

        def rows(table: str) -> List[Dict[str, str]]:
            with gzip.open(source / f"{table}.csv.gz", "rt", newline="") as handle:
                return list(csv.DictReader(handle))

        def boolean(value: Any) -> bool:
            return str(value).strip().lower() in {"1", "true", "t", "yes"}

        for row in rows("games"):
            self.upsert_game(_game_from_dict(json.loads(row["payload_json"])))
        for row in rows("predictions"):
            self.upsert_prediction(
                _prediction_from_dict(json.loads(row["payload_json"])),
                frozen=boolean(row["frozen"]),
            )
        for row in rows("book_snapshots"):
            self.upsert_book_snapshot(
                _book_from_dict(json.loads(row["payload_json"])),
                frozen=boolean(row["frozen"]),
            )
        for row in rows("team_aliases"):
            self.upsert_team_alias(row["canonical"], json.loads(row["variants_json"]))
        for row in rows("provider_health"):
            self.upsert_provider_health(
                _health_from_dict(json.loads(row["payload_json"]))
            )
        for row in rows("pipeline_runs"):
            self._replace(
                "pipeline_runs",
                "run_id = ? AND command = ?",
                [row["run_id"], row["command"]],
                ["run_id", "command", "target_date", "generated_utc", "report_json"],
                [
                    row["run_id"],
                    row["command"],
                    row["target_date"],
                    row["generated_utc"],
                    row["report_json"],
                ],
            )
        return self.counts()

    def export_csv(self, directory: str = "localdata/csv") -> Dict[str, str]:
        """Materialize every canonical table as a deterministic-path CSV.GZ file."""
        out = Path(directory)
        out.mkdir(parents=True, exist_ok=True)
        exported: Dict[str, str] = {}
        for table in EXPORT_TABLES:
            final = out / f"{table}.csv.gz"
            temporary = out / f".{table}.tmp.csv.gz"
            escaped = str(temporary).replace("'", "''")
            self._con.execute(
                f"COPY (SELECT * FROM {table}) TO '{escaped}' "
                "(FORMAT CSV, HEADER TRUE, COMPRESSION GZIP)"
            )
            os.replace(temporary, final)
            exported[table] = str(final)
        return exported
