"""Persistent data store backed by a dedicated Google Sheets spreadsheet.

LEGACY migration/compatibility adapter only. It is selected solely by the
explicit ``--data-store-id`` CLI flag. The canonical unattended system of
record is ``duckdb_store.DuckDBStore`` with cached ``localdata/`` state; this
Sheets store is never an implicit fallback.

Logical tables (one tab each): ``Games``, ``Predictions``, ``BookSnapshots``,
``TeamAliases``, ``ProviderHealth``.  Representation is bounded
latest+frozen keyed by natural identity; payload hashes / source revisions
are stored instead of raw provider JSON.

The store speaks the same :class:`ingestion.store.Store` interface and
round-trips through the same serialisers, so ``JsonFileStore`` and this
store are interchangeable in the pipeline.
"""

from __future__ import annotations

import json
import logging
import time
from typing import Any, Callable, Dict, List, Optional

from .models import (
    BookSnapshot,
    CoverageState,
    Game,
    GameStatus,
    League,
    Prediction,
    ProviderHealth,
)
from .store import (
    FrozenWriteError,
    Store,
    _pred_slot_key,
    _book_slot_key,
    _book_from_dict,
    _book_to_dict,
    _game_from_dict,
    _game_to_dict,
    _prediction_from_dict,
    _prediction_to_dict,
)

logger = logging.getLogger(__name__)

GAMES_TAB = "Games"
PREDICTIONS_TAB = "Predictions"
BOOKS_TAB = "BookSnapshots"
ALIASES_TAB = "TeamAliases"
HEALTH_TAB = "ProviderHealth"

GAMES_HEADERS = ["game_id", "league", "payload_json"]
PREDICTION_HEADERS = ["game_id", "provider", "frozen", "payload_json"]
BOOK_HEADERS = ["snapshot_key", "game_id", "provider", "frozen", "payload_json"]
ALIAS_HEADERS = ["canonical", "variants_json"]
HEALTH_HEADERS = ["provider", "league", "payload_json"]


class SheetsDataStore(Store):
    """Store implementation over five worksheet tabs.

    ``worksheets`` maps table name to a :class:`WorksheetLike` adapter; the
    pipeline builds these from a real spreadsheet via
    :func:`open_data_store` (gspread), tests from ``FakeWorksheet``.
    """

    def __init__(
        self,
        worksheets: Dict[str, Any],
        persist: bool = True,
        sleep: Optional[Callable[[float], None]] = None,
        rate_limit_delay: float = 1.1,
    ) -> None:
        self.persist = persist
        self.sleep = sleep or (lambda seconds: None)
        self.rate_limit_delay = rate_limit_delay
        self._ws: Dict[str, Any] = {}
        for name, headers in (
            (GAMES_TAB, GAMES_HEADERS),
            (PREDICTIONS_TAB, PREDICTION_HEADERS),
            (BOOKS_TAB, BOOK_HEADERS),
            (ALIASES_TAB, ALIAS_HEADERS),
            (HEALTH_TAB, HEALTH_HEADERS),
        ):
            ws = worksheets.get(name)
            if ws is None:
                from .sheets_writer import FakeWorksheet

                ws = FakeWorksheet(headers)
            self._ws[name] = ws
        self._load()

    # -- persistence helpers -------------------------------------------------

    def _rows(self, tab: str) -> List[Dict[str, str]]:
        values = self._ws[tab].get_all_values()
        if not values:
            return []
        headers = [str(h or "").strip() for h in values[0]]
        out = []
        for row in values[1:]:
            if not any(str(v or "").strip() for v in row):
                continue
            out.append({h: row[i] if i < len(row) else "" for i, h in enumerate(headers)})
        return out

    def _rewrite(self, tab: str, headers: List[str], rows: List[List[Any]]) -> None:
        """Replace the whole logical table with a bounded latest+frozen body.

        One batched range update covering the larger of the old/new extents
        pads leftover rows with "" so stale rows from a previous, larger
        table cannot survive. Tables are day-scale small by design.
        """
        if not self.persist:
            return
        ws = self._ws[tab]
        current = ws.get_all_values()
        old_extent = len(current)
        width = max(len(headers), *(len(r) for r in rows), 1) if rows or current else len(headers)
        n_rows = max(old_extent, len(rows) + 1)
        matrix: List[List[Any]] = [list(headers) + [""] * (width - len(headers))]
        for r in rows:
            row = list(r) + [""] * (width - len(r))
            matrix.append(row)
        while len(matrix) < n_rows:
            matrix.append([""] * width)
        from .sheets_writer import column_letter

        rng = f"A1:{column_letter(width)}{n_rows}"
        ws.batch_update_ranges([(rng, matrix)])
        self.sleep(self.rate_limit_delay)

    def _load(self) -> None:
        self._games: Dict[str, Game] = {}
        self._predictions: Dict[str, Dict[str, Prediction]] = {}
        self._frozen_predictions: Dict[str, Dict[str, bool]] = {}
        self._books: Dict[str, Dict[str, BookSnapshot]] = {}
        self._frozen_books: Dict[str, Dict[str, bool]] = {}
        self._aliases: Dict[str, List[str]] = {}
        self._health: Dict[Any, ProviderHealth] = {}

        for row in self._rows(GAMES_TAB):
            try:
                game = _game_from_dict(json.loads(row["payload_json"]))
                self._games[game.game_id] = game
            except Exception as e:  # quarantine bad rows, never crash reads
                logger.warning("data store: skipping bad Games row %s: %s", row.get("game_id"), e)

        for row in self._rows(PREDICTIONS_TAB):
            try:
                pred = _prediction_from_dict(json.loads(row["payload_json"]))
                slot = _pred_slot_key(pred)
                self._predictions.setdefault(pred.game_id, {})[slot] = pred
                if str(row.get("frozen", "")).lower() in ("true", "1", "yes"):
                    self._frozen_predictions.setdefault(pred.game_id, {})[pred.provider] = True
            except Exception as e:
                logger.warning("data store: skipping bad Predictions row: %s", e)

        for row in self._rows(BOOKS_TAB):
            try:
                snap = _book_from_dict(json.loads(row["payload_json"]))
                # Recompute the slotted key so legacy natural-key rows gain
                # pre/post separation on their next persistence pass.
                key = _book_slot_key(snap)
                self._books.setdefault(snap.game_id, {})[key] = snap
                if str(row.get("frozen", "")).lower() in ("true", "1", "yes"):
                    self._frozen_books.setdefault(snap.game_id, {})[key] = True
            except Exception as e:
                logger.warning("data store: skipping bad BookSnapshots row: %s", e)

        for row in self._rows(ALIASES_TAB):
            try:
                self._aliases[row["canonical"]] = json.loads(row.get("variants_json") or "[]")
            except Exception:
                continue

        for row in self._rows(HEALTH_TAB):
            try:
                payload = json.loads(row["payload_json"])
                payload["state"] = CoverageState(payload.get("state", "ok"))
                payload["league"] = League(payload.get("league"))
                health = ProviderHealth(**payload)
                self._health[(health.provider, health.league)] = health
            except Exception:
                continue

    # -- Store interface -------------------------------------------------

    def upsert_game(self, game: Game) -> None:
        existing = self._games.get(game.game_id)
        if existing and existing.rescheduled_from and not game.rescheduled_from:
            game.rescheduled_from = existing.rescheduled_from
        self._games[game.game_id] = game
        self._persist_games()

    def get_game(self, game_id: str) -> Optional[Game]:
        return self._games.get(game_id)

    def list_games(self, league=None, status=None, on_or_after_utc=None, before_utc=None):
        out = list(self._games.values())
        if league is not None:
            out = [g for g in out if g.league == league]
        if status is not None:
            out = [g for g in out if g.status == status]
        if on_or_after_utc:
            out = [g for g in out if g.scheduled_utc >= on_or_after_utc]
        if before_utc:
            out = [g for g in out if g.scheduled_utc < before_utc]
        return sorted(out, key=lambda g: (g.scheduled_utc, g.game_id))

    def upsert_prediction(self, prediction: Prediction, frozen: bool = False) -> None:
        slot = _pred_slot_key(prediction)
        per_game = self._predictions.setdefault(prediction.game_id, {})
        frozen_map = self._frozen_predictions.setdefault(prediction.game_id, {})
        if prediction.pre_tip and frozen_map.get(prediction.provider):
            existing = per_game.get(slot)
            if existing == prediction:
                return
            raise FrozenWriteError(
                f"prediction for {prediction.game_id}/{prediction.provider} is frozen"
            )
        per_game[slot] = prediction
        if frozen and prediction.pre_tip:
            frozen_map[prediction.provider] = True
        self._persist_predictions()

    def get_predictions(self, game_id: str) -> List[Prediction]:
        return list(self._predictions.get(game_id, {}).values())

    def is_prediction_frozen(self, game_id: str, provider: str) -> bool:
        return bool(self._frozen_predictions.get(game_id, {}).get(provider))

    def upsert_book_snapshot(self, snapshot: BookSnapshot, frozen: bool = False) -> None:
        key = _book_slot_key(snapshot)
        per_game = self._books.setdefault(snapshot.game_id, {})
        frozen_map = self._frozen_books.setdefault(snapshot.game_id, {})
        if frozen_map.get(key):
            existing = per_game.get(key)
            if existing == snapshot:
                return
            raise FrozenWriteError(f"book snapshot for {snapshot.game_id}/{key} is frozen")
        per_game[key] = snapshot
        if frozen:
            frozen_map[key] = True
        self._persist_books()

    def get_book_snapshots(self, game_id: str) -> List[BookSnapshot]:
        return list(self._books.get(game_id, {}).values())

    def is_book_frozen(self, game_id: str, provider: str) -> bool:
        return any(k.startswith(f"{provider}|") and v
                   for k, v in self._frozen_books.get(game_id, {}).items())

    def is_book_snapshot_frozen(self, game_id: str, snapshot: BookSnapshot) -> bool:
        return bool(
            self._frozen_books.get(game_id, {}).get(_book_slot_key(snapshot))
        )

    def get_team_aliases(self) -> Dict[str, List[str]]:
        return dict(self._aliases)

    def upsert_team_alias(self, canonical: str, variants: List[str]) -> None:
        self._aliases[canonical] = list(variants)
        self._persist_aliases()

    def upsert_provider_health(self, health: ProviderHealth) -> None:
        self._health[(health.provider, health.league)] = health
        self._persist_health()

    def get_provider_health(self, provider: str, league: League) -> Optional[ProviderHealth]:
        return self._health.get((provider, league))

    # -- table persistence -------------------------------------------------

    def _persist_games(self) -> None:
        rows = [
            [g.game_id, g.league.value, json.dumps(_game_to_dict(g), default=str)]
            for g in sorted(self._games.values(), key=lambda x: x.game_id)
        ]
        self._rewrite(GAMES_TAB, GAMES_HEADERS, rows)

    def _persist_predictions(self) -> None:
        rows = []
        for gid in sorted(self._predictions):
            for slot in sorted(self._predictions[gid]):
                pred = self._predictions[gid][slot]
                rows.append([
                    gid, slot,
                    "true" if self.is_prediction_frozen(gid, pred.provider) else "false",
                    json.dumps(_prediction_to_dict(pred), default=str),
                ])
        self._rewrite(PREDICTIONS_TAB, PREDICTION_HEADERS, rows)

    def _persist_books(self) -> None:
        rows = []
        for gid in sorted(self._books):
            for key in sorted(self._books[gid]):
                snap = self._books[gid][key]
                rows.append([
                    key, gid, snap.provider,
                    "true" if self._frozen_books.get(gid, {}).get(key) else "false",
                    json.dumps(_book_to_dict(snap), default=str),
                ])
        self._rewrite(BOOKS_TAB, BOOK_HEADERS, rows)

    def _persist_aliases(self) -> None:
        rows = [
            [canonical, json.dumps(variants)]
            for canonical, variants in sorted(self._aliases.items())
        ]
        self._rewrite(ALIASES_TAB, ALIAS_HEADERS, rows)

    def _persist_health(self) -> None:
        rows = []
        for (provider, league), health in sorted(
            self._health.items(), key=lambda kv: (kv[0][0], kv[0][1].value)
        ):
            payload = dict(health.__dict__)
            payload["league"] = health.league.value
            payload["state"] = health.state.value
            rows.append([provider, league.value, json.dumps(payload, default=str)])
        self._rewrite(HEALTH_TAB, HEALTH_HEADERS, rows)


def open_data_store(
    spreadsheet: Any, create: bool = True, persist: bool = True
) -> SheetsDataStore:
    """Build a Sheets store; only explicit write mode may create tabs.

    With ``persist=False`` existing tabs are read normally and mutations stay
    in memory, which supports genuine no-write dry runs against prior state.
    """
    worksheets: Dict[str, Any] = {}
    from .sheets_writer import GspreadWorksheetAdapter

    wanted = {
        GAMES_TAB: GAMES_HEADERS,
        PREDICTIONS_TAB: PREDICTION_HEADERS,
        BOOKS_TAB: BOOK_HEADERS,
        ALIASES_TAB: ALIAS_HEADERS,
        HEALTH_TAB: HEALTH_HEADERS,
    }
    existing = {ws.title: ws for ws in spreadsheet.worksheets()}
    for title, headers in wanted.items():
        if title in existing:
            worksheets[title] = GspreadWorksheetAdapter(existing[title])
        elif create:
            ws = spreadsheet.add_worksheet(
                title=title, rows=200, cols=max(8, len(headers) + 2)
            )
            ws.update(values=[headers], range_name="A1")
            time.sleep(1.1)
            worksheets[title] = GspreadWorksheetAdapter(ws)
    return SheetsDataStore(worksheets, persist=persist, sleep=time.sleep)
