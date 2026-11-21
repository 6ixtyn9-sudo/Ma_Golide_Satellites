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
from ingestion.leagues import line_periods, supports_highest_quarter
from ingestion.models import (
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Period,
    Prediction,
)
from ingestion.predictions import DEFAULT_PROVIDER_ORDER

from .accas import (
    build_accas,
    filter_paper_candidate_picks,
    is_provisional_sample_only,
)
from .context_adjustment import (
    ContextAdjustmentConfig,
    build_context_adjustment,
)
from .espn_context import build_espn_player_context, has_confirmed_out_injury
from .feature_manifest import build_feature_use_manifest
from .harness import (
    CalibrationConfig,
    apply_profiles,
    build_profiles,
    calibration_population_diagnostics,
    probability_metrics,
)
from .model import ForecastConfig, HistoryModel, QUARTERS, league_forecast_config
from .shadow import (
    RecencyConfig,
    annotate_recency_profiles,
    build_recency_profiles,
    prequential_shadow_evaluation,
)
from .types import Forecast, MODEL_VERSION, grade_forecast


class MarketForecastService:
    # Leagues whose ESPN evidence path can support the availability surface.
    _ESPN_CONTEXT_LEAGUES = frozenset(
        {League.NBA, League.WNBA, League.NCAA_M, League.NCAA_W}
    )
    # These lanes can render transparent team-history forecasts after their
    # reconciled-history gate, but have no validated named-book execution lane.
    # Keep this explicit: an ESPN schedule/result source or unvalidated relayed
    # odds observation must not make a price/value/ticket claim.
    _UNPRICED_EMPIRICAL_LEAGUES = frozenset(
        {League.EUROLEAGUE_W, League.NBA_G_LEAGUE}
    )

    def __init__(
        self,
        store,
        league: League,
        *,
        forecast_config: ForecastConfig | None = None,
        calibration_config: CalibrationConfig | None = None,
        book_config: BookConfig | None = None,
        provider_order: Sequence[str] = DEFAULT_PROVIDER_ORDER,
        paper_pick_history: Iterable[Forecast] | None = None,
        context_config: ContextAdjustmentConfig | None = None,
    ) -> None:
        self.store = store
        self.league = league
        self.forecast_config = forecast_config or league_forecast_config(league)
        self.calibration_config = calibration_config or CalibrationConfig()
        self.book_config = book_config or BookConfig()
        self.provider_order = tuple(provider_order)
        self.paper_pick_history = (
            None if paper_pick_history is None else list(paper_pick_history)
        )
        self.context_config = context_config or ContextAdjustmentConfig()
        self.model = HistoryModel(self.forecast_config, self.context_config)
        self.shadow_config = RecencyConfig(
            result_delay_hours=self.forecast_config.result_available_after_hours
        )

    def _predictions(
        self, game: Game, *, as_of_utc: str, historical: bool
    ) -> List[tuple[Prediction, bool]]:
        """Return each provider's latest admissible record, in stable order.

        Provider order controls display/ranking stability only; it must not
        suppress otherwise valid provider evidence from forecast generation.
        """
        tip = parse_utc(game.scheduled_utc)
        as_of = parse_utc(as_of_utc)
        by_provider: Dict[str, tuple[Prediction, bool]] = {}
        for prediction in self.store.get_predictions(game.game_id):
            if not prediction.pre_tip:
                continue
            # A provider contributes if it has a coherent moneyline pair or
            # at least one usable full-game expectation. Some feeds publish
            # only score/total/margin forecasts, which remain useful to line
            # markets even without a moneyline probability.
            has_moneyline_probabilities = (
                prediction.prob_home is not None and prediction.prob_away is not None
            )
            has_score_expectation = any(
                value is not None
                for value in (
                    prediction.expected_home_score,
                    prediction.expected_away_score,
                    prediction.expected_total,
                    prediction.expected_margin,
                )
            )
            if not has_moneyline_probabilities and not has_score_expectation:
                continue
            try:
                observed = parse_utc(prediction.observed_utc)
            except (TypeError, ValueError):
                continue
            if observed >= tip or observed > as_of:
                continue
            frozen = self.store.is_prediction_frozen(game.game_id, prediction.provider)
            if historical and not frozen:
                continue
            try:
                prediction.validate_orientation()
            except ValueError:
                continue
            previous = by_provider.get(prediction.provider)
            if previous is None or observed > parse_utc(previous[0].observed_utc):
                by_provider[prediction.provider] = (prediction, frozen)
        ordered = list(dict.fromkeys([*self.provider_order, *sorted(by_provider)]))
        return [by_provider[provider] for provider in ordered if provider in by_provider]

    def _prediction(
        self, game: Game, *, as_of_utc: str, historical: bool
    ) -> tuple[Optional[Prediction], bool]:
        """Backward-compatible highest-priority prediction accessor."""
        predictions = self._predictions(game, as_of_utc=as_of_utc, historical=historical)
        return predictions[0] if predictions else (None, False)

    def _books(self, game: Game, *, as_of_utc: str, historical: bool) -> List[BookSnapshot]:
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
            if historical and not self.store.is_book_snapshot_frozen(game.game_id, snapshot):
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

    def _context_adjustment(
        self, game: Game, finals: Sequence[Game], *, as_of_utc: str
    ):
        """Availability adjustment for one game, or None when evidence is absent.

        Built from the same candidate record the report publishes, so forecast
        metadata and the report's evidence block can never disagree. The
        candidate record itself stays report-only; only its bounded scale
        reaches the model, and only on the separate context surface.
        """
        if self.league not in self._ESPN_CONTEXT_LEAGUES:
            return None
        if not has_confirmed_out_injury(game, as_of_utc=as_of_utc):
            return None
        record = build_espn_player_context(
            game,
            finals,
            as_of_utc=as_of_utc,
            result_delay_hours=self.forecast_config.result_available_after_hours,
        )
        return build_context_adjustment(record, config=self.context_config)

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
        predictions = self._predictions(game, as_of_utc=as_of_utc, historical=historical)
        snapshots = self._books(game, as_of_utc=as_of_utc, historical=historical)
        context = self._context_adjustment(game, finals, as_of_utc=as_of_utc)
        frozen_books = {
            (
                f"{snapshot.provider}|{snapshot.book}|{snapshot.market.value}|"
                f"{snapshot.period.value}|pre"
            ): self.store.is_book_snapshot_frozen(game.game_id, snapshot)
            for snapshot in snapshots
        }
        forecasts: List[Forecast] = []
        full_game_moneyline_books = select_moneyline_snapshots(
            snapshots, self.book_config, Period.FULL_GAME
        )
        for prediction, prediction_frozen in predictions:
            if prediction.prob_home is None or prediction.prob_away is None:
                continue
            forecasts.extend(
                self.model.moneyline(
                    game,
                    prediction,
                    full_game_moneyline_books,
                    frozen_books,
                    prediction_frozen,
                )
            )
        # The provider forecast and the empirical-history forecast are separate
        # model surfaces. The latter can be reconstructed for every prior final
        # using only information available at that tip, so it earns a real
        # walk-forward calibration profile even when historical provider
        # predictions were never captured. Current exact book prices attach to
        # both surfaces; neither inherits the other's calibration history.
        forecasts.extend(
            self.model.period_winner(
                game,
                history,
                full_game_moneyline_books,
                frozen_books,
                period=Period.FULL_GAME,
            )
        )
        forecasts.extend(
            self.model.period_winner(
                game,
                history,
                select_moneyline_snapshots(snapshots, self.book_config, Period.H1),
                frozen_books,
            )
        )
        line_snapshots: List[BookSnapshot] = []
        for period in line_periods(game.league):
            line_snapshots.extend(select_total_snapshots(snapshots, self.book_config, period))
            line_snapshots.extend(select_spread_snapshots(snapshots, self.book_config, period))
        # Keep an empirical line forecast and add a provider-hybrid forecast for
        # every source that supplies score/margin/total expectations. Provider
        # evidence is evaluated in its own source/calibration segment; no
        # provider wins merely because it appears first in provider_order.
        forecasts.extend(
            self.model.lines(
                game,
                history,
                line_snapshots,
                frozen_books,
            )
        )
        if context is not None:
            # Availability surface: full-game moneyline and lines only,
            # empirical history only, its own source and calibration segment.
            # It never modifies base or provider forecasts.
            forecasts.extend(
                self.model.period_winner(
                    game,
                    history,
                    full_game_moneyline_books,
                    frozen_books,
                    period=Period.FULL_GAME,
                    context=context,
                )
            )
            full_game_line_snapshots = [
                snapshot for snapshot in line_snapshots if snapshot.period is Period.FULL_GAME
            ]
            if full_game_line_snapshots:
                forecasts.extend(
                    self.model.lines(
                        game,
                        history,
                        full_game_line_snapshots,
                        frozen_books,
                        context=context,
                    )
                )
        for prediction, prediction_frozen in predictions:
            if not any(
                value is not None
                for value in (
                    prediction.expected_total,
                    prediction.expected_margin,
                    (
                        prediction.expected_home_score
                        if prediction.expected_away_score is not None
                        else None
                    ),
                )
            ):
                continue
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
        if supports_highest_quarter(game.league):
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
            if current is None or (forecast.book.key if forecast.book else "") < (
                current.book.key if current.book else ""
            ):
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

    def _available_records(self, records: Iterable[Forecast], as_of_utc: str) -> List[Forecast]:
        cutoff = parse_utc(as_of_utc) - timedelta(
            hours=self.forecast_config.result_available_after_hours
        )
        return [record for record in records if parse_utc(record.scheduled_utc) <= cutoff]

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
        final_status_games = [game for game in games if game.status is GameStatus.FINAL]
        finals_without_actual = sum(game.actual is None for game in final_status_games)
        finals_with_unreconciled_actual = sum(
            game.actual is not None and not game.actual.reconciles()[0]
            for game in final_status_games
        )
        reconciled_finals_not_yet_past = sum(
            game.actual is not None
            and game.actual.reconciles()[0]
            and parse_utc(game.scheduled_utc) >= as_of
            for game in final_status_games
        )

        records: List[Forecast] = []
        prequential: List[Forecast] = []
        replay_generated_by_segment: Counter = Counter()
        replay_ungraded_by_segment: Counter = Counter()
        replay_graded_by_segment: Counter = Counter()
        replay_outcomes: Counter = Counter()
        finals_with_generated_forecasts = 0
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
            if generated:
                finals_with_generated_forecasts += 1
            available = self._available_records(records, game.scheduled_utc)
            profiles = build_profiles(available, self.calibration_config)
            for forecast in generated:
                key = forecast.versioned_segment_key
                replay_generated_by_segment[key] += 1
                outcome = grade_forecast(forecast, game)
                if outcome is None:
                    replay_ungraded_by_segment[key] += 1
                    continue
                replay_graded_by_segment[key] += 1
                replay_outcomes[outcome] += 1
                forecast.outcome = outcome
                apply_profiles([forecast], profiles, self.calibration_config)
                prequential.append(forecast)
                records.append(forecast)

        shadow_profiles = None
        previous_prequential_shadow = [
            dict(forecast.shadow_calibration) for forecast in prequential
        ]
        try:
            shadow_summary = prequential_shadow_evaluation(
                prequential,
                self.shadow_config,
                champion_config=self.calibration_config,
            )
            shadow_profiles = build_recency_profiles(
                records,
                self.shadow_config,
                as_of_utc=as_of_utc,
                champion_config=self.calibration_config,
            )
        except Exception as exc:
            # Shadow diagnostics are optional and must not make champion
            # forecasts or the usual report unavailable. The evaluator annotates
            # historical rows in place, so roll back partial annotations too.
            for forecast, previous in zip(prequential, previous_prequential_shadow):
                forecast.shadow_calibration = previous
            shadow_summary = {
                "status": "unavailable",
                "error_type": type(exc).__name__,
            }
        available_records = self._available_records(records, as_of_utc)
        current_profiles = build_profiles(available_records, self.calibration_config)
        calibration_population = calibration_population_diagnostics(
            available_records, self.calibration_config
        )
        upcoming = []
        for game in games:
            if game.status is not GameStatus.SCHEDULED:
                continue
            if parse_utc(game.scheduled_utc) <= as_of:
                continue
            if date_yyyymmdd:
                local_date = (
                    parse_utc(game.scheduled_utc)
                    .astimezone(ZoneInfo(timezone_name))
                    .strftime("%Y%m%d")
                )
                if local_date != date_yyyymmdd.replace("-", ""):
                    continue
            upcoming.append(game)
        upcoming.sort(key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))

        current: List[Forecast] = []
        uncovered: List[Dict[str, str]] = []
        for game in upcoming:
            generated = self._generate(game, finals, as_of_utc=as_of_utc, historical=False)
            if not generated:
                uncovered.append(
                    {
                        "game_id": game.game_id,
                        "reason": "insufficient_history_predictions_or_exact_book_facts",
                    }
                )
            current.extend(generated)
        current = apply_profiles(current, current_profiles, self.calibration_config)
        espn_player_context_candidates: Dict[str, Dict[str, Any]] = {}
        if self.league in self._ESPN_CONTEXT_LEAGUES:
            for game in upcoming:
                espn_player_context_candidates[game.game_id] = build_espn_player_context(
                    game,
                    finals,
                    as_of_utc=as_of_utc,
                    result_delay_hours=self.forecast_config.result_available_after_hours,
                )
        if shadow_profiles is not None:
            previous_shadow = [dict(forecast.shadow_calibration) for forecast in current]
            try:
                current = annotate_recency_profiles(current, shadow_profiles, self.shadow_config)
            except Exception as exc:
                for forecast, previous in zip(current, previous_shadow):
                    forecast.shadow_calibration = previous
                shadow_summary = {
                    "status": "unavailable",
                    "error_type": type(exc).__name__,
                }

        current_by_segment = Counter(forecast.versioned_segment_key for forecast in current)
        current_no_history_by_segment = Counter(
            forecast.versioned_segment_key
            for forecast in current
            if forecast.calibration.get("status") == "no_history"
        )
        available_by_segment = Counter(
            forecast.versioned_segment_key for forecast in available_records
        )
        population_by_segment = {
            (
                row["league"],
                row["market"],
                row["period"],
                row["source"],
                row["calibration_version"],
            ): row
            for row in calibration_population["segments"]
        }
        diagnostic_segment_keys = (
            set(replay_generated_by_segment)
            | set(replay_graded_by_segment)
            | set(population_by_segment)
            | set(current_by_segment)
        )
        replay_diagnostic_segments = []
        for key in sorted(diagnostic_segment_keys):
            stage = population_by_segment.get(key, {})
            generated_count = replay_generated_by_segment[key]
            graded_count = replay_graded_by_segment[key]
            available_count = available_by_segment[key]
            deferred_count = max(0, graded_count - available_count)
            profile = current_profiles.get(key)
            profile_n = profile.n if profile is not None else 0
            profile_pool_n = stage.get("profile_pool_samples", 0)
            if profile_n:
                zero_n_reason = None
                sample_status = "sufficient" if profile.sufficient else "below_minimum"
            else:
                sample_status = "zero_n"
                if graded_count == 0 and generated_count:
                    zero_n_reason = "replayed_forecasts_not_gradeable"
                elif graded_count == 0:
                    zero_n_reason = "no_replayed_forecasts_for_segment"
                elif available_count == 0 and deferred_count:
                    zero_n_reason = "all_graded_forecasts_withheld_by_result_delay"
                elif (
                    available_count
                    and stage.get("excluded_before_epoch", 0) == available_count
                ):
                    zero_n_reason = "all_available_forecasts_precede_calibration_epoch"
                elif (
                    available_count
                    and stage.get("invalid_scheduled_times", 0) == available_count
                ):
                    zero_n_reason = "invalid_scheduled_times_excluded"
                elif (
                    stage.get("team_history_insufficient_records", 0)
                    and not stage.get("profile_pool_records", 0)
                ):
                    zero_n_reason = "team_history_insufficient_records_routed_to_diagnostics"
                elif not profile_pool_n and stage.get("profile_pool_pushes", 0):
                    zero_n_reason = "pushes_excluded_from_n"
                elif profile_pool_n:
                    zero_n_reason = "profile_missing_despite_eligible_samples"
                else:
                    zero_n_reason = "no_records_reached_calibration_profile_pool"

            replay_diagnostic_segments.append(
                {
                    "league": key[0],
                    "market": key[1],
                    "period": key[2],
                    "source": key[3],
                    "calibration_version": key[4],
                    "current_forecasts": current_by_segment[key],
                    "current_forecasts_no_history": current_no_history_by_segment[key],
                    "generated_replay_forecasts": generated_count,
                    "ungradeable_replay_forecasts": replay_ungraded_by_segment[key],
                    "graded_replay_forecasts": graded_count,
                    "withheld_by_result_delay": deferred_count,
                    "available_for_profile_filtering": available_count,
                    "invalid_scheduled_times": stage.get("invalid_scheduled_times", 0),
                    "excluded_before_epoch": stage.get("excluded_before_epoch", 0),
                    "epoch_eligible_records": stage.get("epoch_eligible_records", 0),
                    "team_history_insufficient_records": stage.get(
                        "team_history_insufficient_records", 0
                    ),
                    "team_history_insufficient_sample_n": stage.get(
                        "team_history_insufficient_samples", 0
                    ),
                    "team_history_insufficient_pushes": stage.get(
                        "team_history_insufficient_pushes", 0
                    ),
                    "profile_pool_records": stage.get("profile_pool_records", 0),
                    "profile_pool_pushes": stage.get("profile_pool_pushes", 0),
                    "profile_pool_n": profile_pool_n,
                    "profile_n": profile_n,
                    "diagnostic_profile_n": stage.get("team_history_insufficient_samples", 0),
                    "minimum_samples": self.calibration_config.min_samples,
                    "sample_status": sample_status,
                    "zero_n_reason": zero_n_reason,
                }
            )

        replay_diagnostics = {
            "finals": {
                "league_games": len(games),
                "final_status_games": len(final_status_games),
                "finals_without_actual": finals_without_actual,
                "finals_with_unreconciled_actual": finals_with_unreconciled_actual,
                "reconciled_finals_not_yet_past_as_of": reconciled_finals_not_yet_past,
                "historical_finals_replayed": len(finals),
            },
            "forecast_replay": {
                "historical_finals_with_generated_forecasts": finals_with_generated_forecasts,
                "historical_finals_without_generated_forecasts": (
                    len(finals) - finals_with_generated_forecasts
                ),
                "generated_forecasts": sum(replay_generated_by_segment.values()),
                "ungradeable_forecasts": sum(replay_ungraded_by_segment.values()),
                "graded_forecasts": len(records),
                "graded_outcomes": {
                    outcome: replay_outcomes[outcome] for outcome in ("win", "loss", "push")
                },
            },
            "result_availability": {
                "graded_forecasts": len(records),
                "withheld_by_result_delay": len(records) - len(available_records),
                "available_for_profile_filtering": len(available_records),
            },
            "calibration_population": {
                key: value
                for key, value in calibration_population.items()
                if key != "segments"
            },
            "segments": replay_diagnostic_segments,
        }
        strict_accas = build_accas(current, legs=acca_legs, limit=acca_limit)
        paper_history = (
            available_records
            if self.paper_pick_history is None
            else self.paper_pick_history
        )
        paper_history_source = (
            "prequential_settled_individual_forecasts"
            if self.paper_pick_history is None
            else "settled_pick_audit"
        )
        if self.league is League.EUROLEAGUE:
            paper_lane_options = {
                "minimum_historical_picks_for_band": 10,
                "require_positive_expected_value": True,
                "require_point_in_time_price": True,
            }
        else:
            # These source contracts have no validated named-book feed, so they
            # never inherit a priced-value paper lane from another league.
            paper_lane_options = {}
        paper_candidates, paper_pick_filter = filter_paper_candidate_picks(
            current,
            paper_history,
            as_of_utc=as_of_utc,
            history_source=paper_history_source,
            result_available_after_hours=self.forecast_config.result_available_after_hours,
            **paper_lane_options,
        )
        if (
            self.paper_pick_history is not None
            and paper_pick_filter["historical_unique_pick_count"] == 0
        ):
            # A valid but empty/unusable rolling audit can occur after a cold
            # cache start. Reuse the same past-only prequential forecasts that
            # calibration already replayed, rather than deadlocking the paper
            # ROI gate on an empty audit. The ROI filter still requires
            # settled outcomes and exact pre-tip prices, and current forecasts
            # retain all normal eligibility gates.
            paper_candidates, paper_pick_filter = filter_paper_candidate_picks(
                current,
                available_records,
                as_of_utc=as_of_utc,
                history_source="prequential_settled_individual_forecasts_cold_start_fallback",
                result_available_after_hours=self.forecast_config.result_available_after_hours,
                **paper_lane_options,
            )
            paper_pick_filter["history_fallback"] = {
                "used": True,
                "from": "settled_pick_audit",
                "reason": "no_usable_settled_individual_pick_history",
                "settled_audit_usable_pick_count": 0,
            }
        band_details_by_id = paper_pick_filter["candidate_band_details"]
        paper_single_picks = []
        for candidate in paper_candidates:
            single = candidate.to_dict()
            band = band_details_by_id.get(candidate.forecast_id)
            if band is not None:
                single["individual_pick_roi_band"] = dict(band)
            paper_single_picks.append(single)

        qualified_game_count = len({candidate.game_id for candidate in paper_candidates})
        effective_paper_legs = min(acca_legs, qualified_game_count)
        if effective_paper_legs < 2:
            effective_paper_legs = None
            paper_accas = []
        else:
            # Keep the requested size when enough qualifying games exist; if
            # fewer remain, make the largest viable smaller combo (minimum two).
            paper_accas = build_accas(
                paper_candidates,
                legs=effective_paper_legs,
                limit=acca_limit,
                allow_provisional=True,
                require_disjoint_legs=True,
            )
        for acca in paper_accas:
            acca["individual_pick_filter"] = dict(paper_pick_filter)
        market_counts = Counter(f"{forecast.period}:{forecast.market}" for forecast in current)
        eligible_games = {
            forecast.game_id for forecast in current if forecast.eligibility.get("acca_eligible")
        }
        provisional_rows = [
            forecast for forecast in current if is_provisional_sample_only(forecast)
        ]
        provisional_games = {forecast.game_id for forecast in provisional_rows}
        paper_candidate_games = eligible_games | provisional_games
        paper_quality_candidate_games = {forecast.game_id for forecast in paper_candidates}
        feature_use_manifest = build_feature_use_manifest(
            league=self.league,
            games=games,
            upcoming=upcoming,
            store=self.store,
            as_of_utc=as_of_utc,
            candidate_contexts=espn_player_context_candidates,
            result_delay_hours=self.forecast_config.result_available_after_hours,
        )
        return {
            "schema_version": "magolide.market_forecasts",
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
                "league_operator_policy": (
                    "euroleague_strict_named_book_ticket_lane"
                    if self.league is League.EUROLEAGUE
                    else "fiba_elw_empirical_history_only_no_validated_named_book_lane"
                    if self.league is League.EUROLEAGUE_W
                    else "nba_g_league_empirical_history_only_no_validated_named_book_lane"
                    if self.league is League.NBA_G_LEAGUE
                    else "standard_policy"
                ),
                "full_game_moneyline": [
                    "favorite_or_underdog_provider_probability",
                    "independent_empirical_history_distribution",
                ],
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
                "result_available_after_hours": self.forecast_config.result_available_after_hours,
                "calibration_epoch_utc": self.calibration_config.calibration_epoch_utc,
                "historical_finals": len(finals),
                "graded_forecasts": len(records),
                "raw_metrics": probability_metrics(prequential),
                "prequential_calibrated_metrics": probability_metrics(prequential, calibrated=True),
                "replay_diagnostics": replay_diagnostics,
                "shadow_challenger": shadow_summary,
                "segments": [current_profiles[key].to_dict() for key in sorted(current_profiles)],
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
            "feature_use_audit": {
                "team_history_empirical": [
                    "reconciled pre-tip historical team scores and period scores",
                    "team/league scoring baselines and historical score residuals",
                    "current posted market line",
                ],
                "provider_prediction": [
                    "all valid latest-per-provider pre-tip probabilities as separate forecasts",
                    "provider expected scores/margin/total as separate hybrid line forecasts",
                ],
                "espn_player_context": (
                    "confirmed-out availability drives the separate "
                    "context_availability surface (bounded multiplicative scale "
                    "on the empirical surface only); the full candidate record "
                    "is preserved below for audit"
                ),
                "context_availability_surface": (
                    "full-game moneyline and line forecasts scaled by "
                    "confirmed-out player scoring contribution; requires a "
                    "point-in-time injury snapshot and window evidence for both "
                    "teams; own calibration segment; never blended into base "
                    "or provider surfaces; zero adjustment emits no forecasts"
                ),
                "espn_bpi_auxiliary_metrics": (
                    "finite numeric non-core BPI values are retained for audit; "
                    "not applied to probabilities, calibration, or eligibility"
                ),
                "source_coverage": "see feature_use_manifest for observed counts and league-specific support",
            },
            "feature_use_manifest": feature_use_manifest,
            "espn_player_context_candidates": espn_player_context_candidates,
            "forecasts": [forecast.to_dict() for forecast in current],
            "accas": {
                "mode": "paper_only",
                "pricing_evidence_status": (
                    "unpriced_no_validated_named_book_lane"
                    if self.league in self._UNPRICED_EMPIRICAL_LEAGUES
                    else "priced_lane_policy_applies"
                ),
                "paper_policy": (
                    "admits a provisional segment only with at least one graded non-push "
                    "sample when insufficient samples are the sole blocker; exact price, "
                    "team-history, push-risk, and probability-floor gates remain enforced; "
                    "paper tickets additionally require a positive recency-weighted, "
                    "shrinkage-adjusted ROI band from past settled individual picks; "
                    "qualifying selections remain available as singles, and the operator "
                    "freezes a one-unit paper single only when exactly one selection qualifies; "
                    "combo size falls to available distinct qualifying games"
                    + (
                        "; EuroLeague additionally requires an exact pre-tip selected-side quote, "
                        "positive calibrated expected value, and at least ten prior settled picks "
                        "in the matching model/market/period/probability/odds band"
                        if self.league is League.EUROLEAGUE
                        else "; this league has no validated named-book lane: paper counts are unpriced "
                        "decision-support candidates, never value/EV claims"
                        if self.league in self._UNPRICED_EMPIRICAL_LEAGUES
                        else ""
                    )
                ),
                "results_policy": "strict",
                "paper_results_policy": (
                    "strict_plus_sample_only_provisional_and_positive_roi_bands"
                    + (
                        "_plus_EUROLEAGUE_exact_pretip_price_positive_EV_and_10_pick_band_gate"
                        if self.league is League.EUROLEAGUE
                        else ""
                    )
                ),
                "paper_lane": (
                    {
                        "name": "EUROLEAGUE_STRICT_NAMED_BOOK_VALUE_TICKET_LANE",
                        "role": "strict ticket-only lane; aggregate consensus quotes are excluded",
                        "requires_exact_selected_side_pre_tip_price": True,
                        "requires_named_book_quote": True,
                        "requires_positive_calibrated_expected_value": True,
                        "minimum_prior_settled_picks_in_matching_roi_band": 10,
                        "roi_band_scope": [
                            "league", "market", "period", "model", "calibrated_probability_band",
                            "individual_odds_band",
                        ],
                        "history_is_past_only_with_result_delay": True,
                        "status": paper_pick_filter["status"],
                        "candidate_pool_picks": paper_pick_filter["candidate_pool_pick_count"],
                        "point_in_time_priced_candidates": paper_pick_filter[
                            "candidate_with_point_in_time_price_count"
                        ],
                        "positive_expected_value_candidates": paper_pick_filter[
                            "candidate_positive_expected_value_count"
                        ],
                        "qualifying_picks": paper_pick_filter["qualifying_pick_count"],
                    }
                    if self.league is League.EUROLEAGUE
                    else None
                ),
                "paper_pick_filter": paper_pick_filter,
                "paper_single_picks": paper_single_picks,
                "paper_single_pick_count": len(paper_single_picks),
                "paper_ticket_leg_reuse_policy": "no_reused_wager_selections",
                "paper_single_ticket_policy": (
                    "freeze_one_unit_paper_single_only_when_exactly_one_"
                    "roi_qualified_selection_exists"
                ),
                "paper_ticket_leg_count_policy": (
                    "use_minimum_of_requested_legs_and_distinct_roi_qualified_games; "
                    "combination_requires_at_least_two_games"
                ),
                "requested_legs": acca_legs,
                "effective_paper_legs": effective_paper_legs,
                "roi_qualified_single_games": qualified_game_count,
                "requested_limit": acca_limit,
                "eligible_games": len(eligible_games),
                "strict_eligible_games": len(eligible_games),
                "provisional_sample_only_rows": len(provisional_rows),
                "provisional_sample_only_games": len(provisional_games),
                "paper_candidate_games": len(paper_candidate_games),
                "paper_quality_candidate_picks": paper_pick_filter["qualifying_pick_count"],
                "paper_quality_candidate_games": len(paper_quality_candidate_games),
                "ranking": "highest_estimated_joint_win_probability_after_pick_filter",
                "paper_result_count": len(paper_accas),
                "strict_result_count": len(strict_accas),
                "results": strict_accas,
                "strict_results": strict_accas,
                "paper_results": paper_accas,
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
