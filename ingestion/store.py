"""Persistence interface for unattended runs.

All acquisition state (games, predictions, book observations, aliases, and
provider health) goes through this interface. ``DuckDBStore`` is the canonical
implementation; ``InMemoryStore`` supports isolated smoke runs and unit tests.

Freezing rule: once a prediction/book snapshot is stored with
``frozen=True`` (tip passed), later upserts for the same key are rejected
with ``FrozenWriteError`` unless they are the same value (idempotent).
"""

from __future__ import annotations

import threading
from abc import ABC, abstractmethod
from typing import Dict, List, Optional

from .models import (
    BookSnapshot,
    CoverageState,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    Prediction,
    ProviderHealth,
)


class FrozenWriteError(RuntimeError):
    """Attempted overwrite of frozen pre-tip state."""


def _pred_slot_key(prediction: Prediction) -> str:
    """Storage slot: provider + pre/post-tip (pre-tip bundle is protected)."""
    return f"{prediction.provider}|{'pre' if prediction.pre_tip else 'post'}"


def _book_natural_key(snapshot: BookSnapshot) -> str:
    """Book/market identity without its pre/post-match observation slot."""
    return (
        f"{snapshot.provider}|{snapshot.book}|{snapshot.market.value}|"
        f"{snapshot.period.value}"
    )


def _book_slot_key(snapshot: BookSnapshot) -> str:
    """Keep live/post-match observations from displacing pre-match state."""
    return f"{_book_natural_key(snapshot)}|{'pre' if snapshot.pre_match else 'post'}"


def _same(a, b) -> bool:
    return a == b


class Store(ABC):
    """Canonical persistence contract."""

    # -- games ---------------------------------------------------------------
    @abstractmethod
    def upsert_game(self, game: Game) -> None: ...

    @abstractmethod
    def get_game(self, game_id: str) -> Optional[Game]: ...

    @abstractmethod
    def list_games(
        self,
        league: Optional[League] = None,
        status: Optional[GameStatus] = None,
        on_or_after_utc: Optional[str] = None,
        before_utc: Optional[str] = None,
    ) -> List[Game]: ...

    # -- predictions ----------------------------------------------------------
    @abstractmethod
    def upsert_prediction(self, prediction: Prediction, frozen: bool = False) -> None: ...

    @abstractmethod
    def get_predictions(self, game_id: str) -> List[Prediction]: ...

    @abstractmethod
    def is_prediction_frozen(self, game_id: str, provider: str) -> bool: ...

    # -- book snapshots ---------------------------------------------------------
    @abstractmethod
    def upsert_book_snapshot(self, snapshot: BookSnapshot, frozen: bool = False) -> None: ...

    @abstractmethod
    def get_book_snapshots(self, game_id: str) -> List[BookSnapshot]: ...

    @abstractmethod
    def is_book_frozen(self, game_id: str, provider: str) -> bool: ...

    @abstractmethod
    def is_book_snapshot_frozen(self, game_id: str, snapshot: BookSnapshot) -> bool: ...

    # -- team aliases ---------------------------------------------------------
    @abstractmethod
    def get_team_aliases(self) -> Dict[str, List[str]]: ...

    @abstractmethod
    def upsert_team_alias(self, canonical: str, variants: List[str]) -> None: ...

    # -- provider health --------------------------------------------------------
    @abstractmethod
    def upsert_provider_health(self, health: ProviderHealth) -> None: ...

    @abstractmethod
    def get_provider_health(self, provider: str, league: League) -> Optional[ProviderHealth]: ...


# ---------------------------------------------------------------------------
# In-memory implementation (tests / dry runs)
# ---------------------------------------------------------------------------


class InMemoryStore(Store):
    def __init__(self) -> None:
        self._lock = threading.RLock()
        self.games: Dict[str, Game] = {}
        self.predictions: Dict[str, Dict[str, Prediction]] = {}
        self.frozen_predictions: Dict[str, Dict[str, bool]] = {}
        self.books: Dict[str, Dict[str, BookSnapshot]] = {}
        self.frozen_books: Dict[str, Dict[str, bool]] = {}
        self.team_aliases: Dict[str, List[str]] = {}
        self.health: Dict[tuple, ProviderHealth] = {}

    # games
    def upsert_game(self, game: Game) -> None:
        with self._lock:
            existing = self.games.get(game.game_id)
            if existing and existing.rescheduled_from and not game.rescheduled_from:
                game.rescheduled_from = existing.rescheduled_from
            if (
                existing is not None
                and existing.status is GameStatus.FINAL
                and existing.actual is not None
                and game.actual is None
                and game.scheduled_utc == existing.scheduled_utc
                and game.status in {GameStatus.SCHEDULED, GameStatus.LIVE, GameStatus.FINAL}
            ):
                # Schedule rediscovery is not an official result correction;
                # never erase a reconciled final with a metadata-only payload.
                game.actual = existing.actual
                game.status = GameStatus.FINAL
            self.games[game.game_id] = game

    def get_game(self, game_id: str) -> Optional[Game]:
        return self.games.get(game_id)

    def list_games(self, league=None, status=None, on_or_after_utc=None, before_utc=None) -> List[Game]:
        from .identity import parse_utc

        out = list(self.games.values())
        if league is not None:
            out = [g for g in out if g.league == league]
        if status is not None:
            out = [g for g in out if g.status == status]
        if on_or_after_utc:
            boundary = parse_utc(on_or_after_utc)
            out = [g for g in out if parse_utc(g.scheduled_utc) >= boundary]
        if before_utc:
            boundary = parse_utc(before_utc)
            out = [g for g in out if parse_utc(g.scheduled_utc) < boundary]
        return sorted(out, key=lambda g: (parse_utc(g.scheduled_utc), g.game_id))

    # predictions
    def upsert_prediction(self, prediction: Prediction, frozen: bool = False) -> None:
        """Latest prediction per (provider, pre/post-tip slot).

        The pre-tip slot is the frozen-state carrier: a post-tip observation
        lands in its own slot and can never displace the stored pre-tip
        bundle the results renderer depends on.
        """
        with self._lock:
            slot = _pred_slot_key(prediction)
            per_game = self.predictions.setdefault(prediction.game_id, {})
            frozen_map = self.frozen_predictions.setdefault(prediction.game_id, {})
            provider = prediction.provider
            if prediction.pre_tip and frozen_map.get(provider):
                existing = per_game.get(slot)
                if existing is not None and _same(existing, prediction):
                    return  # idempotent re-run
                raise FrozenWriteError(
                    f"prediction for {prediction.game_id}/{provider} is frozen"
                )
            per_game[slot] = prediction
            if frozen and prediction.pre_tip:
                frozen_map[provider] = True

    def get_predictions(self, game_id: str) -> List[Prediction]:
        return list(self.predictions.get(game_id, {}).values())

    def is_prediction_frozen(self, game_id: str, provider: str) -> bool:
        return bool(self.frozen_predictions.get(game_id, {}).get(provider))

    # books
    def upsert_book_snapshot(self, snapshot: BookSnapshot, frozen: bool = False) -> None:
        with self._lock:
            key = _book_slot_key(snapshot)
            per_game = self.books.setdefault(snapshot.game_id, {})
            frozen_map = self.frozen_books.setdefault(snapshot.game_id, {})
            if frozen_map.get(key):
                existing = per_game.get(key)
                if existing is not None and _same(existing, snapshot):
                    return
                raise FrozenWriteError(f"book snapshot for {snapshot.game_id}/{key} is frozen")
            per_game[key] = snapshot
            if frozen:
                frozen_map[key] = True

    def get_book_snapshots(self, game_id: str) -> List[BookSnapshot]:
        return list(self.books.get(game_id, {}).values())

    def is_book_frozen(self, game_id: str, provider: str) -> bool:
        return any(
            key.startswith(f"{provider}|") and val
            for key, val in self.frozen_books.get(game_id, {}).items()
        )

    def is_book_snapshot_frozen(self, game_id: str, snapshot: BookSnapshot) -> bool:
        return bool(
            self.frozen_books.get(game_id, {}).get(_book_slot_key(snapshot))
        )

    # aliases
    def get_team_aliases(self) -> Dict[str, List[str]]:
        return dict(self.team_aliases)

    def upsert_team_alias(self, canonical: str, variants: List[str]) -> None:
        with self._lock:
            self.team_aliases[canonical] = list(variants)

    # health
    def upsert_provider_health(self, health: ProviderHealth) -> None:
        with self._lock:
            self.health[(health.provider, health.league)] = health

    def get_provider_health(self, provider: str, league: League) -> Optional[ProviderHealth]:
        return self.health.get((provider, league))


# ---------------------------------------------------------------------------
# (de)serialisation helpers
# ---------------------------------------------------------------------------


def _game_to_dict(g: Game) -> dict:
    d = dict(
        game_id=g.game_id,
        league=g.league.value,
        season=g.season,
        home_team_id=g.home_team_id,
        away_team_id=g.away_team_id,
        home_name=g.home_name,
        away_name=g.away_name,
        scheduled_utc=g.scheduled_utc,
        status=g.status.value,
        phase=g.phase,
        provider_ids=g.provider_ids,
        rescheduled_from=g.rescheduled_from,
    )
    if g.actual is not None:
        d["actual"] = dict(
            period_kind=g.actual.period_kind.value,
            home_quarters=g.actual.home_quarters,
            away_quarters=g.actual.away_quarters,
            home_overtime=g.actual.home_overtime,
            away_overtime=g.actual.away_overtime,
            home_final=g.actual.home_final,
            away_final=g.actual.away_final,
            overtime_periods=g.actual.overtime_periods,
            source=g.actual.source,
        )
    return d


def _game_from_dict(d: dict) -> Game:
    actual = None
    if d.get("actual"):
        a = d["actual"]
        from .models import ActualResult, PeriodKind

        actual = ActualResult(
            period_kind=PeriodKind(a["period_kind"]),
            home_quarters=a["home_quarters"],
            away_quarters=a["away_quarters"],
            home_overtime=a.get("home_overtime"),
            away_overtime=a.get("away_overtime"),
            home_final=a.get("home_final"),
            away_final=a.get("away_final"),
            overtime_periods=a.get("overtime_periods", 0),
            source=a.get("source", ""),
        )
    return Game(
        game_id=d["game_id"],
        league=League(d["league"]),
        season=d["season"],
        home_team_id=d["home_team_id"],
        away_team_id=d["away_team_id"],
        home_name=d["home_name"],
        away_name=d["away_name"],
        scheduled_utc=d["scheduled_utc"],
        status=GameStatus(d.get("status", "SCHEDULED")),
        phase=d.get("phase", "League"),
        provider_ids=d.get("provider_ids", {}),
        actual=actual,
        rescheduled_from=d.get("rescheduled_from"),
    )


def _prediction_to_dict(p: Prediction) -> dict:
    return p.__dict__.copy()


def _prediction_from_dict(d: dict) -> Prediction:
    return Prediction(**d)


def _book_to_dict(s: BookSnapshot) -> dict:
    d = s.__dict__.copy()
    d["market"] = s.market.value
    d["period"] = s.period.value
    return d


def _book_from_dict(d: dict) -> BookSnapshot:
    d = dict(d)
    d["market"] = Market(d["market"])
    d["period"] = Period(d["period"])
    return BookSnapshot(**d)


def _health_to_dict(health: ProviderHealth) -> dict:
    d = health.__dict__.copy()
    d["league"] = health.league.value
    d["state"] = health.state.value
    return d


def _health_from_dict(d: dict) -> ProviderHealth:
    d = dict(d)
    d["league"] = League(d.get("league", League.NBA.value))
    d["state"] = CoverageState(d.get("state", "ok"))
    return ProviderHealth(**d)
