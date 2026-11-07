"""Warehouse-native walk-forward forecasts, calibration, and acca reports."""

from __future__ import annotations

from collections import Counter
from datetime import timedelta
import json
import os
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence
from zoneinfo import ZoneInfo

from ingestion.books import (
    BookConfig,
    select_moneyline_snapshots,
    select_spread_snapshots,
    select_total_snapshots,
)
from ingestion.identity import parse_utc
from ingestion.models import (
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Period,
    Prediction,
)
from ingestion.predictions import DEFAULT_PROVIDER_ORDER, is_complete

from .accas import build_accas
from .harness import (
    CalibrationConfig,
    apply_profiles,
    build_profiles,
    probability_metrics,
)
from .model import ForecastConfig, HistoryModel, QUARTERS
from .types import Forecast, MODEL_VERSION, grade_forecast


class MarketForecastService:
    def __init__(
        self,
        store,
        league: League,
        *,
        forecast_config: ForecastConfig | None = None,
        calibration_config: CalibrationConfig | None = None,
        book_config: BookConfig | None = None,
        provider_order: Sequence[str] = DEFAULT_PROVIDER_ORDER,
    ) -> None:
        self.store = store
        self.league = league
        self.forecast_config = forecast_config or ForecastConfig()
        self.calibration_config = calibration_config or CalibrationConfig()
        self.book_config = book_config or BookConfig()
        self.provider_order = tuple(provider_order)
        self.model = HistoryModel(self.forecast_config)

    def _prediction(
        self, game: Game, *, as_of_utc: str, historical: bool
    ) -> tuple[Optional[Prediction], bool]:
        tip = parse_utc(game.scheduled_utc)
        as_of = parse_utc(as_of_utc)
        by_provider: Dict[str, Prediction] = {}
        for prediction in self.store.get_predictions(game.game_id):
            if not prediction.pre_tip or not is_complete(prediction):
                continue
            observed = parse_utc(prediction.observed_utc)
            if observed >= tip or observed > as_of:
                continue
            frozen = self.store.is_prediction_frozen(
                game.game_id, prediction.provider
            )
            if historical and not frozen:
                continue
            try:
                prediction.validate_orientation()
            except ValueError:
                continue
            previous = by_provider.get(prediction.provider)
            if previous is None or observed > parse_utc(previous.observed_utc):
                by_provider[prediction.provider] = prediction
        ordered = [*self.provider_order, *sorted(by_provider)]
        for provider in ordered:
            if provider in by_provider:
                return by_provider[provider], self.store.is_prediction_frozen(
                    game.game_id, provider
                )
        return None, False

    def _books(
        self, game: Game, *, as_of_utc: str, historical: bool
    ) -> List[BookSnapshot]:
        tip = parse_utc(game.scheduled_utc)
        as_of = parse_utc(as_of_utc)
        visible = []
        for snapshot in self.store.get_book_snapshots(game.game_id):
            if not snapshot.pre_match or not self.book_config.matches_source(snapshot):
                continue
            if self.book_config.prefer_main_line and not snapshot.main_line:
                continue
            observed = parse_utc(snapshot.observed_utc)
            if observed >= tip or observed > as_of:
                continue
            if historical and not self.store.is_book_snapshot_frozen(
                game.game_id, snapshot
            ):
                continue
            visible.append(snapshot)
        return sorted(
            visible,
            key=lambda snapshot: (
                snapshot.market.value,
                snapshot.period.value,
                snapshot.provider,
                snapshot.book,
                parse_utc(snapshot.observed_utc),
            ),
        )

    def _generate(
        self,
        game: Game,
        finals: Sequence[Game],
        *,
        as_of_utc: str,
        historical: bool,
    ) -> List[Forecast]:
        history = self.model.available_history(
            finals, scheduled_utc=game.scheduled_utc, as_of_utc=as_of_utc
        )
        prediction, prediction_frozen = self._prediction(
            game, as_of_utc=as_of_utc, historical=historical
        )
        snapshots = self._books(game, as_of_utc=as_of_utc, historical=historical)
        frozen_books = {
            (
                f"{snapshot.provider}|{snapshot.book}|{snapshot.market.value}|"
                f"{snapshot.period.value}|pre"
            ): self.store.is_book_snapshot_frozen(game.game_id, snapshot)
            for snapshot in snapshots
        }
        forecasts: List[Forecast] = []
        if prediction is not None:
            forecasts.extend(
                self.model.moneyline(
                    game,
                    prediction,
                    select_moneyline_snapshots(snapshots, self.book_config),
                    frozen_books,
                    prediction_frozen,
                )
            )
        forecasts.extend(
            self.model.period_winner(
                game,
                history,
                select_moneyline_snapshots(
                    snapshots, self.book_config, Period.H1
                ),
                frozen_books,
            )
        )
        line_snapshots: List[BookSnapshot] = []
        for period in (Period.FULL_GAME, *QUARTERS):
            line_snapshots.extend(
                select_total_snapshots(snapshots, self.book_config, period)
            )
        for period in QUARTERS:
            line_snapshots.extend(
                select_spread_snapshots(snapshots, self.book_config, period)
            )
        forecasts.extend(
            self.model.lines(
                game,
                history,
                line_snapshots,
                frozen_books,
                prediction,
                prediction_frozen,
            )
        )
        forecasts.extend(self.model.highest_quarter(game, history))
        return sorted(
            forecasts,
            key=lambda forecast: (
                forecast.market,
                forecast.period,
                forecast.book.key if forecast.book else "",
                forecast.selection,
            ),
        )

    @staticmethod
    def _training_rows(forecasts: Iterable[Forecast]) -> List[Forecast]:
        """One book observation per game/model/market prevents book overweighting."""
        selected: Dict[tuple, Forecast] = {}
        for forecast in forecasts:
            key = (*forecast.market_key, forecast.source)
            current = selected.get(key)
            if current is None or (
                forecast.book.key if forecast.book else ""
            ) < (current.book.key if current.book else ""):
                selected[key] = forecast
        return sorted(
            selected.values(),
            key=lambda forecast: (
                forecast.market,
                forecast.period,
                forecast.source,
                forecast.book.key if forecast.book else "",
            ),
        )

    def _available_records(
        self, records: Iterable[Forecast], as_of_utc: str
    ) -> List[Forecast]:
        cutoff = parse_utc(as_of_utc) - timedelta(
            hours=self.forecast_config.result_available_after_hours
        )
        return [
            record
            for record in records
            if parse_utc(record.scheduled_utc) <= cutoff
        ]

    def build_report(
        self,
        *,
        as_of_utc: str,
        date_yyyymmdd: Optional[str] = None,
        timezone_name: str = "UTC",
        acca_legs: int = 3,
        acca_limit: int = 5,
    ) -> Dict[str, Any]:
        as_of = parse_utc(as_of_utc)
        games = self.store.list_games(league=self.league)
        finals = [
            game
            for game in games
            if game.status is GameStatus.FINAL
            and game.actual is not None
            and game.actual.reconciles()[0]
            and parse_utc(game.scheduled_utc) < as_of
        ]
        finals.sort(key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))

        records: List[Forecast] = []
        prequential: List[Forecast] = []
        for game in finals:
            # Simulate a forecast at tip: later finals are excluded by the
            # availability cutoff, and all prediction/book evidence is frozen.
            generated = self._training_rows(
                self._generate(
                    game,
                    finals,
                    as_of_utc=game.scheduled_utc,
                    historical=True,
                )
            )
            available = self._available_records(records, game.scheduled_utc)
            profiles = build_profiles(available, self.calibration_config)
            for forecast in generated:
                outcome = grade_forecast(forecast, game)
                if outcome is None:
                    continue
                forecast.outcome = outcome
                apply_profiles([forecast], profiles, self.calibration_config)
                prequential.append(forecast)
                records.append(forecast)

        current_profiles = build_profiles(
            self._available_records(records, as_of_utc), self.calibration_config
        )
        upcoming = []
        for game in games:
            if game.status is not GameStatus.SCHEDULED:
                continue
            if parse_utc(game.scheduled_utc) <= as_of:
                continue
            if date_yyyymmdd:
                local_date = parse_utc(game.scheduled_utc).astimezone(
                    ZoneInfo(timezone_name)
                ).strftime("%Y%m%d")
                if local_date != date_yyyymmdd.replace("-", ""):
                    continue
            upcoming.append(game)
        upcoming.sort(key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))

        current: List[Forecast] = []
        uncovered: List[Dict[str, str]] = []
        for game in upcoming:
            generated = self._generate(
                game, finals, as_of_utc=as_of_utc, historical=False
            )
            if not generated:
                uncovered.append(
                    {
                        "game_id": game.game_id,
                        "reason": "insufficient_history_predictions_or_exact_book_facts",
                    }
                )
            current.extend(generated)
        current = apply_profiles(current, current_profiles, self.calibration_config)
        accas = build_accas(current, legs=acca_legs, limit=acca_limit)
        market_counts = Counter(
            f"{forecast.period}:{forecast.market}" for forecast in current
        )
        eligible_games = {
            forecast.game_id
            for forecast in current
            if forecast.eligibility.get("acca_eligible")
        }
        return {
            "schema_version": "magolide.market_forecasts.v1",
            "status": "decision_support_only",
            "league": self.league.value,
            "generated_utc": as_of_utc,
            "target_date": date_yyyymmdd,
            "model_version": MODEL_VERSION,
            "execution": {
                "external_betting_writes": False,
                "guaranteed_winners": False,
                "certification": "none",
            },
            "capabilities": {
                "full_game_moneyline": "favorite_or_underdog_provider_probability",
                "quarter_totals": [period.value for period in QUARTERS],
                "quarter_spreads": [period.value for period in QUARTERS],
                "first_half_1x2": ["home", "draw", "away"],
                "full_game_total": "provider_baseline_plus_historical_residuals",
                "highest_quarter": (
                    "Q1-Q4 highest-or-tied forecast; acca requires an exact "
                    "priced highest-quarter book market and dead-heat rules"
                ),
            },
            "walk_forward": {
                "method": "prequential_by_scheduled_tip",
                "result_available_after_hours": (
                    self.forecast_config.result_available_after_hours
                ),
                "historical_finals": len(finals),
                "graded_forecasts": len(records),
                "raw_metrics": probability_metrics(prequential),
                "prequential_calibrated_metrics": probability_metrics(
                    prequential, calibrated=True
                ),
                "segments": [
                    current_profiles[key].to_dict() for key in sorted(current_profiles)
                ],
                "guards": [
                    "future_finals_excluded_per_forecast",
                    "historical_predictions_and_books_must_be_frozen_pre_tip",
                    "one_book_per_game_market_in_training",
                    "unreconciled_finals_excluded",
                    "pushes_excluded_from_win_loss_calibration",
                ],
            },
            "forecast_count": len(current),
            "forecasts_by_market": dict(sorted(market_counts.items())),
            "uncovered_games": uncovered,
            "forecasts": [forecast.to_dict() for forecast in current],
            "accas": {
                "requested_legs": acca_legs,
                "requested_limit": acca_limit,
                "eligible_games": len(eligible_games),
                "ranking": "highest_estimated_joint_win_probability",
                "results": accas,
            },
        }


def write_market_report(path: str, report: Dict[str, Any]) -> str:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    raw = (json.dumps(report, indent=2, sort_keys=True, default=str) + "\n").encode()
    if target.exists() and target.read_bytes() == raw:
        return str(target)
    temporary = target.with_suffix(".json.tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, target)
    return str(target)


__all__ = ["MarketForecastService", "write_market_report"]
