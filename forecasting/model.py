"""Leakage-conscious empirical basketball market forecasts.

Team scoring baselines use only supplied history. Residual distributions are
built from historical baselines that themselves used only earlier games. Book
lines and prices remain immutable evidence attached to each forecast.
"""

from __future__ import annotations

from bisect import bisect_right, bisect_left
from collections.abc import Iterator
from dataclasses import asdict, dataclass
from datetime import timedelta
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from ingestion.identity import canonical_team_key, parse_utc
from ingestion.models import BookSnapshot, Game, League, Market, Period, Prediction

from .context_adjustment import (
    CONTEXT_SOURCE,
    ContextAdjustment,
    ContextAdjustmentConfig,
    context_calibration_components,
)
from .harness import config_fingerprint
from .types import BookEvidence, Forecast, MODEL_VERSION, _period_scores, id_spelling

QUARTERS = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)
RECALIBRATION_MAX_SLOPE = 6.0
# Settings added after calibration identities were established. At their
# defaults they are omitted from the fingerprint so leagues that keep the
# original behaviour keep their accumulated calibration version.
_FINGERPRINT_OPTIONAL_DEFAULTS: Dict[str, Any] = {
    "recalibrate_expectations": False,
    "posted_line_model_weight": None,
}


@dataclass(frozen=True)
class ForecastConfig:
    min_history_games: int = 20
    min_residual_games: int = 20
    team_prior_games: float = 8.0
    recent_half_life_games: float = 20.0
    # None means use every available reconciled final; leagues with deep
    # histories must not silently lose earlier seasons to a league-wide cap.
    training_window: Optional[int] = None
    residual_window: int = 100
    result_available_after_hours: int = 4
    # Walk-forward linear recalibration of the empirical expected margin and
    # total. The raw baseline averages team for/against rates, which halves
    # rating differences, and early-season shrinkage compresses them further;
    # an OLS fit of actual outcomes on prior baselines (same residual window,
    # no look-ahead) restores the scale. Off by default for leagues whose
    # paper pipeline is frozen pending review.
    recalibrate_expectations: bool = False
    # Weight of the model mean against a posted bookmaker spread/total line.
    # ``None`` keeps pure-model line probabilities. With a weight w, the line
    # distribution is centred at w * model + (1 - w) * posted line.
    posted_line_model_weight: Optional[float] = None

    def __post_init__(self) -> None:
        if self.min_history_games < 1 or self.min_residual_games < 1:
            raise ValueError("history and residual minimums must be positive")
        if self.team_prior_games <= 0 or self.recent_half_life_games <= 0:
            raise ValueError("shrinkage and half-life must be positive")
        if self.training_window is not None and self.training_window < self.min_history_games:
            raise ValueError("training window is smaller than minimum history")
        if self.residual_window < self.min_residual_games:
            raise ValueError("residual window is smaller than minimum residual count")
        if self.result_available_after_hours < 0:
            raise ValueError("result availability lag cannot be negative")
        if self.posted_line_model_weight is not None and not (
            0.0 <= self.posted_line_model_weight <= 1.0
        ):
            raise ValueError("posted line model weight must be within [0, 1]")


# Per-league model settings validated by walk-forward backtests on the
# 2025-26 college seasons (ESPN finals, DraftKings closing lines):
#   NCAA_W (796 lined games): winner Brier 0.179 -> 0.148, margin MAE
#     15.1 -> 11.6; spread Brier 0.303 -> 0.250 with the line weight.
#   NCAA_M (1,139 lined games): winner Brier 0.216 -> 0.202, margin MAE
#     11.3 -> 9.8; spread Brier 0.278 -> 0.251 with the line weight.
# Pure-model spread probabilities were badly overconfident (claimed >=70%
# picks won 55% / 49%). Other leagues keep the original behaviour until
# they are backtested and their paper pipelines are open to change.
LEAGUE_FORECAST_OVERRIDES: Dict[League, Dict[str, Any]] = {
    League.NCAA_W: {"recalibrate_expectations": True, "posted_line_model_weight": 0.2},
    League.NCAA_M: {"recalibrate_expectations": True, "posted_line_model_weight": 0.2},
    # FIBA ELW starts on the transparent empirical/default surface.  It does
    # not inherit men's EuroLeague backtests, model weights, or price lane.
    League.EUROLEAGUE_W: {},
}


def league_forecast_config(league: League, **overrides: Any) -> ForecastConfig:
    """Default ``ForecastConfig`` for ``league``; explicit overrides win."""
    return ForecastConfig(**{**LEAGUE_FORECAST_OVERRIDES.get(league, {}), **overrides})


@dataclass(frozen=True)
class Baseline:
    home: float
    away: float
    games: int
    home_team_effective_games: float = 0.0
    away_team_effective_games: float = 0.0

    @property
    def margin(self) -> float:
        return self.home - self.away

    @property
    def total(self) -> float:
        return self.home + self.away


@dataclass(frozen=True)
class Distribution:
    baseline: Baseline
    margins: Tuple[float, ...]
    totals: Tuple[float, ...]

    @property
    def residuals(self) -> int:
        return len(self.margins)


@dataclass(frozen=True)
class _TeamRates:
    points_for: float
    points_against: float
    effective_games: float


@dataclass(frozen=True)
class _TeamPrefix:
    game_indices: Tuple[int, ...]
    points_for: Tuple[float, ...]
    points_against: Tuple[float, ...]
    weights: Tuple[float, ...]


def _game_neutral_site(game: Game) -> Optional[bool]:
    """Return tri-state neutral-site venue flag from provider evidence.

    ``True`` means explicit neutral site (zero home-court advantage, excluded
    from league home-advantage estimation). ``False`` and ``None`` preserve
    normal home-court advantage so lanes without a venue flag never silently
    lose home-court signal.
    """
    evidence = game.evidence if isinstance(game.evidence, dict) else {}
    espn = evidence.get("espn")
    if isinstance(espn, dict):
        scoreboard = espn.get("scoreboard")
        if isinstance(scoreboard, dict):
            raw = scoreboard.get("neutral_site")
            if isinstance(raw, bool):
                return raw
    raw_bal = evidence.get("neutral")
    if isinstance(raw_bal, bool):
        return raw_bal
    raw_direct = evidence.get("neutral_site")
    if isinstance(raw_direct, bool):
        return raw_direct
    return None


class _PeriodHistoryIndex:
    """Prefix sums for league and exponentially weighted team evidence."""

    def __init__(
        self,
        games: Sequence[Game],
        period: Period,
        half_life: float,
        *,
        home_keys: Optional[Tuple[str, ...]] = None,
        away_keys: Optional[Tuple[str, ...]] = None,
        neutral_sites: Optional[Tuple[bool, ...]] = None,
    ) -> None:
        league_home = [0.0]
        league_away = [0.0]
        league_count = [0]
        ha_diff = [0.0]
        ha_count = [0]
        scores_list: List[Optional[Tuple[float, float]]] = []
        team_games: Dict[str, List[int]] = {}
        team_for: Dict[str, List[float]] = {}
        team_against: Dict[str, List[float]] = {}
        team_weights: Dict[str, List[float]] = {}
        decay = 0.5 ** (1.0 / half_life)

        def add_team(team: str, game_index: int, points_for: float, points_against: float) -> None:
            indices = team_games.setdefault(team, [])
            for_values = team_for.setdefault(team, [0.0])
            against_values = team_against.setdefault(team, [0.0])
            weights = team_weights.setdefault(team, [0.0])
            indices.append(game_index)
            for_values.append(decay * for_values[-1] + points_for)
            against_values.append(decay * against_values[-1] + points_against)
            weights.append(decay * weights[-1] + 1.0)

        for index, game in enumerate(games):
            scores = _period_scores(game, period)
            scores_list.append(scores)
            if scores is None:
                league_home.append(league_home[-1])
                league_away.append(league_away[-1])
                league_count.append(league_count[-1])
                ha_diff.append(ha_diff[-1])
                ha_count.append(ha_count[-1])
                continue
            home, away = scores
            league_home.append(league_home[-1] + home)
            league_away.append(league_away[-1] + away)
            league_count.append(league_count[-1] + 1)
            is_neutral = (
                neutral_sites[index]
                if neutral_sites is not None
                else (_game_neutral_site(game) is True)
            )
            if is_neutral:
                ha_diff.append(ha_diff[-1])
                ha_count.append(ha_count[-1])
            else:
                ha_diff.append(ha_diff[-1] + (home - away))
                ha_count.append(ha_count[-1] + 1)
            home_team = (
                home_keys[index] if home_keys is not None else canonical_team_key(game.home_name)
            )
            away_team = (
                away_keys[index] if away_keys is not None else canonical_team_key(game.away_name)
            )
            add_team(home_team, index, home, away)
            if away_team != home_team:
                add_team(away_team, index, away, home)

        self.scores = tuple(scores_list)
        self.half_life = half_life
        self.decay = decay
        self.league_home = tuple(league_home)
        self.league_away = tuple(league_away)
        self.league_count = tuple(league_count)
        self.ha_diff = tuple(ha_diff)
        self.ha_count = tuple(ha_count)
        self.teams = {
            team: _TeamPrefix(
                tuple(team_games[team]),
                tuple(team_for[team]),
                tuple(team_against[team]),
                tuple(team_weights[team]),
            )
            for team in team_games
        }
        self.period = period

    def league_rates(self, start: int, end: int) -> Optional[tuple[float, float]]:
        count = self.league_count[end] - self.league_count[start]
        if not count:
            return None
        home = self.league_home[end] - self.league_home[start]
        away = self.league_away[end] - self.league_away[start]
        ha_n = self.ha_count[end] - self.ha_count[start]
        home_advantage = (self.ha_diff[end] - self.ha_diff[start]) / ha_n if ha_n else 0.0
        return (home + away) / (2 * count), home_advantage

    @staticmethod
    def _team_state(values: Tuple[float, ...], lo: int, hi: int, decay: float) -> float:
        return values[hi] - values[lo] * (decay ** (hi - lo))

    def team_rates_by_key(
        self, team_key: str, start: int, end: int, half_life: float
    ) -> Optional[_TeamRates]:
        evidence = self.teams.get(team_key)
        if evidence is None:
            return None
        lo = bisect_left(evidence.game_indices, start)
        hi = bisect_left(evidence.game_indices, end)
        if lo == hi:
            return None
        decay = self.decay if half_life == self.half_life else (0.5 ** (1.0 / half_life))
        weight = self._team_state(evidence.weights, lo, hi, decay)
        if weight <= 0:
            return None
        points_for = self._team_state(evidence.points_for, lo, hi, decay)
        points_against = self._team_state(evidence.points_against, lo, hi, decay)
        return _TeamRates(points_for / weight, points_against / weight, weight)

    def team_rates(
        self, team_name: str, start: int, end: int, half_life: float
    ) -> Optional[_TeamRates]:
        return self.team_rates_by_key(canonical_team_key(team_name), start, end, half_life)


class _IndexedHistory(Sequence[Game]):
    """Immutable chronological facts plus reusable prefix-stat indexes."""

    def __init__(self, games: Sequence[Game], half_life: float) -> None:
        self.games = tuple(games)
        self.times = tuple(parse_utc(game.scheduled_utc) for game in self.games)
        self.home_keys = tuple(canonical_team_key(game.home_name) for game in self.games)
        self.away_keys = tuple(canonical_team_key(game.away_name) for game in self.games)
        self.neutral_sites = tuple(_game_neutral_site(game) is True for game in self.games)
        self.half_life = half_life
        self._periods: Dict[Period, _PeriodHistoryIndex] = {}
        self._residual_entry_cache: Dict[
            Tuple[Any, ...], Optional[Tuple[float, float, float, float]]
        ] = {}
        self._zero_start_residuals: Dict[
            Tuple[Any, ...],
            List[Any],
        ] = {}
        self._quarter_vector_cache: Dict[
            Tuple[Any, ...], Optional[Tuple[float, float, float, float]]
        ] = {}
        self._zero_start_quarter_vectors: Dict[
            Tuple[Any, ...],
            List[Any],
        ] = {}

    def __len__(self) -> int:
        return len(self.games)

    def __getitem__(self, index):
        return self.games[index]

    def __iter__(self) -> Iterator[Game]:
        return iter(self.games)

    def period(self, period: Period) -> _PeriodHistoryIndex:
        result = self._periods.get(period)
        if result is None:
            result = _PeriodHistoryIndex(
                self.games,
                period,
                self.half_life,
                home_keys=self.home_keys,
                away_keys=self.away_keys,
                neutral_sites=self.neutral_sites,
            )
            self._periods[period] = result
        return result


class _HistoryView(Sequence[Game]):
    """A range into indexed history; avoids copying every replay prefix."""

    def __init__(self, index: _IndexedHistory, start: int, end: int) -> None:
        self.index = index
        self.start = start
        self.end = end
        self.residual_cache: Dict[Period, List[Tuple[float, float, float, float]]] = {}

    def __len__(self) -> int:
        return self.end - self.start

    def __getitem__(self, position):
        if isinstance(position, slice):
            start, stop, step = position.indices(len(self))
            return [self.index.games[self.start + i] for i in range(start, stop, step)]
        if position < 0:
            position += len(self)
        if position < 0 or position >= len(self):
            raise IndexError(position)
        return self.index.games[self.start + position]

    def __iter__(self) -> Iterator[Game]:
        return iter(self.index.games[self.start : self.end])

    def __reversed__(self) -> Iterator[Game]:
        return (self.index.games[i] for i in range(self.end - 1, self.start - 1, -1))


def _config_components(config: ForecastConfig) -> Dict[str, Any]:
    """Prediction-relevant config, omitting optional settings at defaults."""
    components = asdict(config)
    for key, default in _FINGERPRINT_OPTIONAL_DEFAULTS.items():
        if components.get(key) == default:
            components.pop(key, None)
    return components


class HistoryModel:
    def __init__(
        self,
        config: ForecastConfig | None = None,
        context_config: ContextAdjustmentConfig | None = None,
    ) -> None:
        self.config = config or ForecastConfig()
        self.context_config = context_config or ContextAdjustmentConfig()
        self._cached_source: Optional[Iterable[Game]] = None
        self._cached_source_size: Optional[int] = None
        self._cached_index: Optional[_IndexedHistory] = None
        self._context_calibration_version_cache: Optional[str] = None
        # Every setting in ForecastConfig can change the available evidence or
        # empirical probabilities. Source identity remains a separate segment
        # dimension; this tag isolates the algorithm/configuration itself.
        components = {"model_version": id_spelling(MODEL_VERSION), **_config_components(self.config)}
        self.calibration_version = f"{MODEL_VERSION}-{config_fingerprint(components)}"

    @property
    def context_calibration_version(self) -> str:
        """Version tag for the availability-adjusted surface only.

        Deliberately excludes the context knobs from the base version above:
        they cannot change base or provider probabilities, so base segments
        must keep their accumulated calibration identity.
        """
        if self._context_calibration_version_cache is None:
            components = context_calibration_components(
                self.context_config, base_config_asdict=_config_components(self.config)
            )
            self._context_calibration_version_cache = (
                f"{MODEL_VERSION}-context-{config_fingerprint(components)}"
            )
        return self._context_calibration_version_cache

    @staticmethod
    def _stable_prediction_model_phase(prediction: Prediction) -> Optional[str]:
        """Return a closed, normalized Predixsport model-phase label if known.

        Free-text ``produced_by`` remains provenance, not identity. The only
        provider-specific exception is Predixsport's documented earlyseason /
        mature distinction, recognized as an exact suffix of its model name;
        arbitrary labels are ignored rather than fragmenting calibration.
        """
        provider = (prediction.provider or "").strip().casefold()
        model = (prediction.model or "").strip().casefold()
        producer = (prediction.produced_by or "").strip().casefold()
        if provider != "predixsport" or not model.startswith("predixsport-"):
            return None
        prefix = f"{model}-"
        if not producer.startswith(prefix):
            return None
        phase = producer[len(prefix) :].replace("_", "-")
        if phase == "earlyseason":
            return "earlyseason"
        if phase in {"mature", "mature-model"}:
            return "mature"
        return None

    def _prediction_calibration_version(
        self, prediction: Prediction, *, hybrid: bool = False
    ) -> str:
        """Fingerprint stable model provenance, never per-event revisions.

        Provider artifacts may update their ``source_revision`` every game and
        payload-derived ``produced_by`` labels may vary by row. Calibration
        identity therefore uses provider/model/version metadata, plus only a
        closed Predixsport model-phase tag when it is recognized. Hybrid
        history-plus-provider lines also include this model's full
        configuration fingerprint.
        """
        components = {
            "provider": prediction.provider,
            "model": prediction.model,
            "model_version": id_spelling(prediction.model_version),
        }
        model_phase = self._stable_prediction_model_phase(prediction)
        if model_phase is not None:
            components["model_phase"] = model_phase
        prefix = "provider_prediction"
        if hybrid:
            components["history_model_calibration_version"] = id_spelling(
                self.calibration_version
            )
            prefix = "history_provider_hybrid"
        return f"{prefix}-{config_fingerprint(components)}"

    def available_history(
        self, games: Iterable[Game], *, scheduled_utc: str, as_of_utc: str
    ) -> Sequence[Game]:
        """Return all available reconciled finals (unless explicitly capped).

        The indexed source is reused across walk-forward targets. Use an
        immutable source snapshot during one replay (length changes invalidate
        the cache). A range view keeps per-target replay O(1) to create while
        preserving every game in the evidence window by default.
        """
        try:
            source_size = len(games)  # type: ignore[arg-type]
        except TypeError:
            source_size = None
        source_changed = (
            games is not self._cached_source
            or self._cached_index is None
            or (
                source_size is not None
                and self._cached_source_size is not None
                and source_size != self._cached_source_size
            )
        )
        if source_changed:
            eligible = [
                game
                for game in games
                if game.actual is not None
                and game.actual.reconciles()[0]
            ]
            eligible.sort(key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))
            self._cached_source = games
            self._cached_source_size = source_size
            self._cached_index = _IndexedHistory(eligible, self.config.recent_half_life_games)
        index = self._cached_index
        cutoff = min(parse_utc(scheduled_utc), parse_utc(as_of_utc)) - timedelta(
            hours=self.config.result_available_after_hours
        )
        end = bisect_right(index.times, cutoff)
        start = max(0, end - self.config.training_window) if self.config.training_window else 0
        return _HistoryView(index, start, end)

    @staticmethod
    def _team_key(name: str) -> str:
        return canonical_team_key(name)

    def _league_rates(
        self, history: Sequence[Game], period: Period
    ) -> Optional[tuple[float, float]]:
        home_scores: List[float] = []
        away_scores: List[float] = []
        non_neutral_diffs: List[float] = []
        for game in history:
            scores = _period_scores(game, period)
            if scores is not None:
                home_scores.append(scores[0])
                away_scores.append(scores[1])
                if _game_neutral_site(game) is not True:
                    non_neutral_diffs.append(scores[0] - scores[1])
        if not home_scores:
            return None
        per_team = (sum(home_scores) + sum(away_scores)) / (2 * len(home_scores))
        home_advantage = (
            sum(non_neutral_diffs) / len(non_neutral_diffs)
            if non_neutral_diffs
            else 0.0
        )
        return per_team, home_advantage

    def _team_rates(
        self, team_name: str, history: Sequence[Game], period: Period
    ) -> Optional[_TeamRates]:
        team = self._team_key(team_name)
        # Collect this team's games newest-first so age counts team games, not league games.
        # Bugfix: league-index age made 8.0 unreachable (spacing ~15 → ceiling 2.47); team-index restores documented intent "prior = 8 team games".
        team_games: list[tuple[float, float]] = []
        for game in reversed(history):
            scores = _period_scores(game, period)
            if scores is None:
                continue
            if self._team_key(game.home_name) == team:
                team_games.append((scores[0], scores[1]))
            elif self._team_key(game.away_name) == team:
                team_games.append((scores[1], scores[0]))
        weighted_for = weighted_against = weight_sum = 0.0
        for age, (points_for, points_against) in enumerate(team_games):
            weight = 0.5 ** (age / self.config.recent_half_life_games)
            weighted_for += weight * points_for
            weighted_against += weight * points_against
            weight_sum += weight
        if weight_sum == 0:
            return None
        return _TeamRates(
            weighted_for / weight_sum,
            weighted_against / weight_sum,
            weight_sum,
        )

    def _baseline_config_signature(self) -> Tuple[Any, ...]:
        cfg = self.config
        return (
            cfg.min_history_games,
            cfg.team_prior_games,
            cfg.recent_half_life_games,
            cfg.result_available_after_hours,
        )

    def _indexed_baseline_keys(
        self,
        home_key: str,
        away_key: str,
        is_neutral: bool,
        index: _IndexedHistory,
        start: int,
        end: int,
        period: Period,
    ) -> Optional[Baseline]:
        if end - start < self.config.min_history_games:
            return None
        stats = index.period(period)
        league = stats.league_rates(start, end)
        if league is None:
            return None
        mean, home_advantage = league
        prior = self.config.team_prior_games
        half_life = self.config.recent_half_life_games

        home_rates = stats.team_rates_by_key(home_key, start, end, half_life)
        away_rates = stats.team_rates_by_key(away_key, start, end, half_life)
        if home_rates is None:
            home_for = mean
            home_against = mean
            home_eff = 0.0
        else:
            home_eff = home_rates.effective_games
            denom = home_eff + prior
            home_for = (home_eff * home_rates.points_for + prior * mean) / denom
            home_against = (home_eff * home_rates.points_against + prior * mean) / denom
        if away_rates is None:
            away_for = mean
            away_against = mean
            away_eff = 0.0
        else:
            away_eff = away_rates.effective_games
            denom = away_eff + prior
            away_for = (away_eff * away_rates.points_for + prior * mean) / denom
            away_against = (away_eff * away_rates.points_against + prior * mean) / denom
        applied_home_advantage = 0.0 if is_neutral else home_advantage
        expected_home = (home_for + away_against) / 2 + applied_home_advantage / 2
        expected_away = (away_for + home_against) / 2 - applied_home_advantage / 2
        return Baseline(
            expected_home,
            expected_away,
            end - start,
            home_team_effective_games=home_eff,
            away_team_effective_games=away_eff,
        )

    def _indexed_baseline(
        self,
        game: Game,
        index: _IndexedHistory,
        start: int,
        end: int,
        period: Period,
    ) -> Optional[Baseline]:
        return self._indexed_baseline_keys(
            canonical_team_key(game.home_name),
            canonical_team_key(game.away_name),
            _game_neutral_site(game) is True,
            index,
            start,
            end,
            period,
        )

    def baseline(self, game: Game, history: Sequence[Game], period: Period) -> Optional[Baseline]:
        if isinstance(history, _HistoryView):
            return self._indexed_baseline(
                game, history.index, history.start, history.end, period
            )
        if len(history) < self.config.min_history_games:
            return None
        league = self._league_rates(history, period)
        if league is None:
            return None
        mean, home_advantage = league
        prior = self.config.team_prior_games

        def shrunk(rates: Optional[_TeamRates], field: str) -> float:
            if rates is None:
                return mean
            value = getattr(rates, field)
            weight = rates.effective_games
            return (weight * value + prior * mean) / (weight + prior)

        home_rates = self._team_rates(game.home_name, history, period)
        away_rates = self._team_rates(game.away_name, history, period)
        applied_home_advantage = 0.0 if _game_neutral_site(game) is True else home_advantage
        expected_home = (
            shrunk(home_rates, "points_for") + shrunk(away_rates, "points_against")
        ) / 2 + applied_home_advantage / 2
        expected_away = (
            shrunk(away_rates, "points_for") + shrunk(home_rates, "points_against")
        ) / 2 - applied_home_advantage / 2
        return Baseline(
            expected_home,
            expected_away,
            len(history),
            home_team_effective_games=(
                home_rates.effective_games if home_rates is not None else 0.0
            ),
            away_team_effective_games=(
                away_rates.effective_games if away_rates is not None else 0.0
            ),
        )

    def _team_history(self, baselines: Sequence[Baseline]) -> Dict[str, Any]:
        """Describe whether team evidence outweighs the league-only fallback.

        The prior already represents eight effective games by default. Requiring
        at least that much evidence for each team prevents a mostly league-wide
        early-season average from becoming an ACCA-quality team prediction.
        """
        required = self.config.team_prior_games
        home = min(row.home_team_effective_games for row in baselines)
        away = min(row.away_team_effective_games for row in baselines)
        return {
            "home_effective_games": round(home, 6),
            "away_effective_games": round(away, 6),
            "required_effective_games_per_team": required,
            "sufficient_for_accas": home >= required and away >= required,
        }

    def _compute_residual_entry(
        self,
        index: _IndexedHistory,
        stats: _PeriodHistoryIndex,
        start: int,
        target_index: int,
        period: Period,
        delay: timedelta,
    ) -> Optional[Tuple[float, float, float, float]]:
        actual = stats.scores[target_index]
        if actual is None:
            return None
        cutoff = index.times[target_index] - delay
        prior_end = min(
            target_index,
            bisect_right(index.times, cutoff, lo=start, hi=target_index),
        )
        baseline = self._indexed_baseline_keys(
            index.home_keys[target_index],
            index.away_keys[target_index],
            index.neutral_sites[target_index],
            index,
            start,
            prior_end,
            period,
        )
        if baseline is None:
            return None
        return (
            baseline.margin,
            baseline.total,
            actual[0] - actual[1],
            actual[0] + actual[1],
        )

    def _residual_samples(
        self, history: Sequence[Game], period: Period
    ) -> List[Tuple[float, float, float, float]]:
        """Walk-forward ``(baseline margin, baseline total, actual margin,
        actual total)`` samples, oldest first, for the configured window."""
        if isinstance(history, _HistoryView):
            cached = history.residual_cache.get(period)
            if cached is not None:
                return cached
            index, start, end = history.index, history.start, history.end
        else:
            index = _IndexedHistory(history, self.config.recent_half_life_games)
            start, end = 0, len(index)
        stats = index.period(period)
        delay = timedelta(hours=self.config.result_available_after_hours)
        cfg_sig = self._baseline_config_signature()
        window = self.config.residual_window

        if start == 0:
            stream_key = (period, cfg_sig)
            stream = index._zero_start_residuals.get(stream_key)
            if stream is not None or end <= window * 2:
                if stream is None:
                    stream = [0, [], []]  # [computed_upto, valid_indices, valid_samples]
                    index._zero_start_residuals[stream_key] = stream
                computed_upto, valid_indices, valid_samples = stream
                while computed_upto < end:
                    entry = self._compute_residual_entry(
                        index, stats, 0, computed_upto, period, delay
                    )
                    if entry is not None:
                        valid_indices.append(computed_upto)
                        valid_samples.append(entry)
                    computed_upto += 1
                stream[0] = computed_upto
                pos = (
                    len(valid_indices)
                    if end >= computed_upto
                    else bisect_left(valid_indices, end)
                )
                samples = valid_samples[max(0, pos - window) : pos]
                if isinstance(history, _HistoryView):
                    history.residual_cache[period] = samples
                return samples

        samples: List[Tuple[float, float, float, float]] = []
        entry_cache = index._residual_entry_cache
        # Residual distributions only use the most recent configured samples.
        # Walking backwards avoids recomputing old, discarded residuals.
        for target_index in range(end - 1, start - 1, -1):
            cache_key = (start, target_index, period, cfg_sig)
            if cache_key in entry_cache:
                entry = entry_cache[cache_key]
            else:
                entry = self._compute_residual_entry(
                    index, stats, start, target_index, period, delay
                )
                entry_cache[cache_key] = entry
            if entry is None:
                continue
            samples.append(entry)
            if len(samples) >= window:
                break
        samples.reverse()
        if isinstance(history, _HistoryView):
            history.residual_cache[period] = samples
        return samples

    def _residuals(
        self, history: Sequence[Game], period: Period
    ) -> tuple[List[float], List[float]]:
        samples = self._residual_samples(history, period)
        return (
            [actual_margin - margin for margin, _, actual_margin, _ in samples],
            [actual_total - total for _, total, _, actual_total in samples],
        )

    @staticmethod
    def _linear_recalibration(
        predicted: Sequence[float], actual: Sequence[float]
    ) -> Tuple[float, float]:
        """OLS ``actual ~ intercept + slope * predicted`` with safe fallbacks.

        A degenerate predictor (no spread, e.g. every team still at the league
        prior) keeps the identity slope; the slope is bounded to
        ``[0, RECALIBRATION_MAX_SLOPE]`` so a noisy window cannot invert or
        explode the rating scale.
        """
        n = len(predicted)
        mean_x = sum(predicted) / n
        mean_y = sum(actual) / n
        var_x = sum((x - mean_x) ** 2 for x in predicted)
        if var_x <= 1e-9 * n:
            slope = 1.0
        else:
            covariance = sum(
                (x - mean_x) * (y - mean_y) for x, y in zip(predicted, actual)
            )
            slope = min(RECALIBRATION_MAX_SLOPE, max(0.0, covariance / var_x))
        return mean_y - slope * mean_x, slope

    def distribution(
        self,
        game: Game,
        history: Sequence[Game],
        period: Period,
        prediction: Optional[Prediction] = None,
        context: Optional[ContextAdjustment] = None,
    ) -> Optional[Distribution]:
        if context is not None and period is not Period.FULL_GAME:
            raise ValueError("context adjustment is limited to full-game forecasts")
        if context is not None and prediction is not None:
            # Provider expectations may already internalize availability
            # information; scaling them again would double-count. The context
            # surface is strictly the empirical-history surface.
            raise ValueError("context adjustment applies only to the empirical surface")
        baseline = self.baseline(game, history, period)
        if baseline is None:
            return None
        samples = self._residual_samples(history, period)
        if len(samples) < self.config.min_residual_games:
            return None
        if self.config.recalibrate_expectations:
            margin_fit = self._linear_recalibration(
                [row[0] for row in samples], [row[2] for row in samples]
            )
            total_fit = self._linear_recalibration(
                [row[1] for row in samples], [row[3] for row in samples]
            )
            margin_residuals = [
                actual - (margin_fit[0] + margin_fit[1] * predicted)
                for predicted, _, actual, _ in samples
            ]
            total_residuals = [
                actual - (total_fit[0] + total_fit[1] * predicted)
                for _, predicted, _, actual in samples
            ]
            margin = margin_fit[0] + margin_fit[1] * baseline.margin
            total = total_fit[0] + total_fit[1] * baseline.total
            baseline = Baseline(
                (total + margin) / 2,
                (total - margin) / 2,
                baseline.games,
                home_team_effective_games=baseline.home_team_effective_games,
                away_team_effective_games=baseline.away_team_effective_games,
            )
        else:
            margin_residuals = [actual - predicted for predicted, _, actual, _ in samples]
            total_residuals = [actual - predicted for _, predicted, _, actual in samples]
        if period is Period.FULL_GAME and prediction is not None:
            prediction.validate_orientation()
            home = baseline.home
            away = baseline.away
            if (
                prediction.expected_home_score is not None
                and prediction.expected_away_score is not None
            ):
                home = prediction.expected_home_score
                away = prediction.expected_away_score
            else:
                margin = (
                    prediction.expected_margin
                    if prediction.expected_margin is not None
                    else baseline.margin
                )
                total = (
                    prediction.expected_total
                    if prediction.expected_total is not None
                    else baseline.total
                )
                home, away = (total + margin) / 2, (total - margin) / 2
            baseline = Baseline(
                home,
                away,
                baseline.games,
                home_team_effective_games=baseline.home_team_effective_games,
                away_team_effective_games=baseline.away_team_effective_games,
            )
        if context is not None:
            # Availability scaling moves the current-game mean only. Historical
            # residuals stay untouched: they describe how real outcomes
            # scattered around unadjusted baselines, and the empirical surface
            # is self-centering (unlike the provider hybrid, which replaces the
            # mean with an external expectation).
            baseline = Baseline(
                baseline.home * context.home_scale,
                baseline.away * context.away_scale,
                baseline.games,
                home_team_effective_games=baseline.home_team_effective_games,
                away_team_effective_games=baseline.away_team_effective_games,
            )
        if len(samples) < self.config.min_residual_games:
            return None
        # Hybrid path: baseline mean replaced by BPI, but raw residuals carry -2.34 mean → drag ~0.10 prob. Center them.
        if period is Period.FULL_GAME and prediction is not None:
            # Center only margins/totals when BPI replaced the mean; pure-empirical is self-correcting.
            m_mean = sum(margin_residuals) / len(margin_residuals) if margin_residuals else 0.0
            t_mean = sum(total_residuals) / len(total_residuals) if total_residuals else 0.0
            # Only center if prediction actually replaced baseline (check source: caller sets baseline already)
            # Heuristic: if prediction has any total/margin/score, it's hybrid.
            is_hybrid = (
                prediction.expected_total is not None
                or prediction.expected_margin is not None
                or prediction.expected_home_score is not None
            )
            if is_hybrid:
                margin_residuals = [v - m_mean for v in margin_residuals]
                total_residuals = [v - t_mean for v in total_residuals]
        return Distribution(
            baseline=baseline,
            margins=tuple(baseline.margin + value for value in margin_residuals),
            totals=tuple(baseline.total + value for value in total_residuals),
        )

    @staticmethod
    def _winner_probabilities(margins: Sequence[float]) -> Dict[str, float]:
        n = len(margins)
        return {
            "home": sum(value > 0 for value in margins) / n,
            "away": sum(value < 0 for value in margins) / n,
            "draw": sum(value == 0 for value in margins) / n,
        }

    @staticmethod
    def _select(probabilities: Dict[str, float], allowed: Iterable[str]) -> Optional[str]:
        choices = list(allowed)
        highest = max(probabilities[key] for key in choices)
        winners = [key for key in choices if probabilities[key] == highest]
        # Never turn an exact tie into an arbitrary favorite.
        return winners[0] if len(winners) == 1 else None

    def moneyline(
        self,
        game: Game,
        prediction: Prediction,
        books: Sequence[BookSnapshot],
        frozen_books: Dict[str, bool],
        prediction_frozen: bool,
    ) -> List[Forecast]:
        if prediction.prob_home is None or prediction.prob_away is None:
            return []
        prediction.validate_orientation()
        probabilities = {
            "home": prediction.prob_home,
            "away": prediction.prob_away,
        }
        selection = self._select(probabilities, ("home", "away"))
        if selection is None:
            return []
        snapshots: Sequence[Optional[BookSnapshot]] = books or (None,)
        forecasts = []
        for snapshot in snapshots:
            evidence = (
                BookEvidence.from_snapshot(
                    snapshot, frozen=frozen_books.get(_book_key(snapshot), False)
                )
                if snapshot is not None
                else None
            )
            role = None
            if evidence is not None:
                home_price = evidence.price_for("home")
                away_price = evidence.price_for("away")
                if home_price and away_price:
                    selected_price = evidence.price_for(selection)
                    other_price = away_price if selection == "home" else home_price
                    if selected_price and abs(selected_price - other_price) > 1e-9:
                        role = "underdog" if selected_price > other_price else "favorite"
            forecasts.append(
                Forecast(
                    game_id=game.game_id,
                    league=game.league.value,
                    scheduled_utc=game.scheduled_utc,
                    home_name=game.home_name,
                    away_name=game.away_name,
                    market=Market.MONEYLINE.value,
                    period=Period.FULL_GAME.value,
                    selection=selection,
                    raw_probability=probabilities[selection],
                    alternative_probabilities=probabilities,
                    source=f"provider:{prediction.provider}",
                    calibration_version=self._prediction_calibration_version(prediction),
                    book=evidence,
                    prediction_provider=prediction.provider,
                    prediction_observed_utc=prediction.observed_utc,
                    prediction_frozen=prediction_frozen,
                    prediction_model=prediction.model,
                    prediction_model_version=prediction.model_version,
                    prediction_produced_by=prediction.produced_by,
                    role=role,
                )
            )
        return forecasts

    def period_winner(
        self,
        game: Game,
        history: Sequence[Game],
        books: Sequence[BookSnapshot],
        frozen_books: Dict[str, bool],
        period: Period = Period.H1,
        context: Optional[ContextAdjustment] = None,
    ) -> List[Forecast]:
        effective_context = context if period is Period.FULL_GAME else None
        distribution = self.distribution(game, history, period, context=effective_context)
        if distribution is None:
            return []
        probabilities = self._winner_probabilities(distribution.margins)
        # Basketball full-game moneyline is two-way because overtime decides
        # the final. H1 retains its explicit draw selection and exact draw
        # price requirement; no draw may leak into a two-way full-game book.
        allowed = ("home", "away") if period is Period.FULL_GAME else ("home", "away", "draw")
        selection = self._select(probabilities, allowed)
        if selection is None:
            return []
        snapshots: Sequence[Optional[BookSnapshot]] = books or (None,)
        return [
            Forecast(
                game_id=game.game_id,
                league=game.league.value,
                scheduled_utc=game.scheduled_utc,
                home_name=game.home_name,
                away_name=game.away_name,
                market=Market.MONEYLINE.value,
                period=period.value,
                selection=selection,
                raw_probability=probabilities[selection],
                alternative_probabilities=probabilities,
                source=CONTEXT_SOURCE if effective_context is not None else MODEL_VERSION,
                calibration_version=(
                    self.context_calibration_version
                    if effective_context is not None
                    else self.calibration_version
                ),
                context=effective_context.to_metadata() if effective_context is not None else None,
                expected_margin=distribution.baseline.margin,
                expected_total=distribution.baseline.total,
                training_games=distribution.baseline.games,
                residual_games=distribution.residuals,
                team_history=self._team_history([distribution.baseline]),
                book=(
                    BookEvidence.from_snapshot(
                        snapshot, frozen=frozen_books.get(_book_key(snapshot), False)
                    )
                    if snapshot is not None
                    else None
                ),
            )
            for snapshot in snapshots
        ]

    @staticmethod
    def _anchor_lines(
        snapshots: Sequence[BookSnapshot],
    ) -> Dict[Tuple[Market, Period], float]:
        """Choose one posted market anchor line per ``(market, period)``.

        When ``posted_line_model_weight`` blends the model distribution toward
        the market expectation, anchoring to *each* snapshot's line treats every
        alternate line as if it were the 50/50 main line — stamping identical
        probabilities onto lines 1–2 points apart. Instead, compute a single
        deterministic anchor per ``(market, period)`` (preferring ``main_line``
        quotes, taking their median when multiple books/captures exist) and
        evaluate every alternate line against that one shifted distribution.
        """
        grouped: Dict[Tuple[Market, Period], List[BookSnapshot]] = {}
        for snapshot in snapshots:
            if snapshot.line is None or snapshot.market not in {
                Market.TOTAL,
                Market.SPREAD,
            }:
                continue
            grouped.setdefault((snapshot.market, snapshot.period), []).append(snapshot)
        anchors: Dict[Tuple[Market, Period], float] = {}
        for key, group in grouped.items():
            preferred = [snap for snap in group if snap.main_line]
            chosen = preferred if preferred else group
            values = sorted(float(snap.line) for snap in chosen if snap.line is not None)
            if not values:
                continue
            mid = len(values) // 2
            anchors[key] = (
                values[mid]
                if len(values) % 2 == 1
                else 0.5 * (values[mid - 1] + values[mid])
            )
        return anchors

    def lines(
        self,
        game: Game,
        history: Sequence[Game],
        snapshots: Iterable[BookSnapshot],
        frozen_books: Dict[str, bool],
        prediction: Optional[Prediction] = None,
        prediction_frozen: bool = False,
        context: Optional[ContextAdjustment] = None,
    ) -> List[Forecast]:
        if context is not None and prediction is not None:
            raise ValueError("context adjustment applies only to the empirical surface")
        snapshot_list = list(snapshots)
        weight = self.config.posted_line_model_weight
        anchor_lines = (
            self._anchor_lines(snapshot_list) if weight is not None else {}
        )
        forecasts: List[Forecast] = []
        distributions: Dict[Period, Optional[Distribution]] = {}
        for snapshot in snapshot_list:
            if snapshot.line is None or snapshot.market not in {
                Market.TOTAL,
                Market.SPREAD,
            }:
                continue
            # Availability scaling is defined for full-game expectations only;
            # period scoring shares of benched players are not observed, so
            # period lines stay unadjusted under a context call.
            effective_context = context if snapshot.period is Period.FULL_GAME else None
            if snapshot.period not in distributions:
                distributions[snapshot.period] = self.distribution(
                    game, history, snapshot.period, prediction, effective_context
                )
            distribution = distributions[snapshot.period]
            if distribution is None:
                continue
            if snapshot.market is Market.TOTAL:
                line = snapshot.line
                totals = distribution.totals
                if weight is not None:
                    # Centre once between the model total and the game's main
                    # posted total, then price every line against that single
                    # distribution so probabilities stay monotonic in line.
                    anchor = anchor_lines.get((snapshot.market, snapshot.period), line)
                    shift = (1.0 - weight) * (anchor - distribution.baseline.total)
                    totals = tuple(value + shift for value in totals)
                probabilities = {
                    "over": sum(value > line for value in totals)
                    / distribution.residuals,
                    "under": sum(value < line for value in totals)
                    / distribution.residuals,
                }
                push = sum(value == line for value in totals) / (
                    distribution.residuals
                )
            else:
                margins = distribution.margins
                if weight is not None:
                    # A home-oriented main line L implies a posted margin of -L.
                    anchor = anchor_lines.get(
                        (snapshot.market, snapshot.period), snapshot.line
                    )
                    shift = (1.0 - weight) * (-anchor - distribution.baseline.margin)
                    margins = tuple(value + shift for value in margins)
                adjusted = [value + snapshot.line for value in margins]
                probabilities = {
                    "home": sum(value > 0 for value in adjusted) / distribution.residuals,
                    "away": sum(value < 0 for value in adjusted) / distribution.residuals,
                }
                push = sum(value == 0 for value in adjusted) / distribution.residuals
            selection = self._select(probabilities, probabilities.keys())
            if selection is None:
                continue
            source = MODEL_VERSION
            if effective_context is not None:
                source = CONTEXT_SOURCE
            elif (
                snapshot.period is Period.FULL_GAME
                and prediction is not None
                and (
                    prediction.expected_total is not None
                    or prediction.expected_margin is not None
                    or prediction.expected_home_score is not None
                )
            ):
                source = f"provider:{prediction.provider}+{MODEL_VERSION}"
            forecasts.append(
                Forecast(
                    game_id=game.game_id,
                    league=game.league.value,
                    scheduled_utc=game.scheduled_utc,
                    home_name=game.home_name,
                    away_name=game.away_name,
                    market=snapshot.market.value,
                    period=snapshot.period.value,
                    selection=selection,
                    line=snapshot.line,
                    raw_probability=probabilities[selection],
                    alternative_probabilities=probabilities,
                    push_probability=push,
                    expected_margin=distribution.baseline.margin,
                    expected_total=distribution.baseline.total,
                    training_games=distribution.baseline.games,
                    residual_games=distribution.residuals,
                    team_history=self._team_history([distribution.baseline]),
                    source=source,
                    calibration_version=(
                        self.context_calibration_version
                        if effective_context is not None
                        else (
                            self._prediction_calibration_version(prediction, hybrid=True)
                            if source.startswith("provider:") and prediction is not None
                            else self.calibration_version
                        )
                    ),
                    context=(
                        effective_context.to_metadata()
                        if effective_context is not None
                        else None
                    ),
                    book=BookEvidence.from_snapshot(
                        snapshot, frozen=frozen_books.get(_book_key(snapshot), False)
                    ),
                    prediction_provider=(
                        prediction.provider if source.startswith("provider:") else None
                    ),
                    prediction_observed_utc=(
                        prediction.observed_utc if source.startswith("provider:") else None
                    ),
                    prediction_frozen=(
                        prediction_frozen if source.startswith("provider:") else False
                    ),
                    prediction_model=(
                        prediction.model if source.startswith("provider:") else None
                    ),
                    prediction_model_version=(
                        prediction.model_version if source.startswith("provider:") else None
                    ),
                    prediction_produced_by=(
                        prediction.produced_by if source.startswith("provider:") else None
                    ),
                )
            )
        return forecasts

    def _compute_quarter_vector(
        self,
        index: _IndexedHistory,
        quarter_stats: Tuple[_PeriodHistoryIndex, ...],
        start: int,
        target_index: int,
        delay: timedelta,
    ) -> Optional[Tuple[float, float, float, float]]:
        cutoff = index.times[target_index] - delay
        prior_end = min(
            target_index,
            bisect_right(index.times, cutoff, lo=start, hi=target_index),
        )
        if prior_end - start < self.config.min_history_games:
            return None
        home_key = index.home_keys[target_index]
        away_key = index.away_keys[target_index]
        is_neutral = index.neutral_sites[target_index]
        vector: List[float] = []
        for period, stats in zip(QUARTERS, quarter_stats):
            actual = stats.scores[target_index]
            if actual is None:
                return None
            baseline = self._indexed_baseline_keys(
                home_key, away_key, is_neutral, index, start, prior_end, period
            )
            if baseline is None:
                return None
            vector.append(sum(actual) - baseline.total)
        return (vector[0], vector[1], vector[2], vector[3])

    def highest_quarter(self, game: Game, history: Sequence[Game]) -> List[Forecast]:
        if len(history) < self.config.min_history_games:
            return []
        baselines = {period: self.baseline(game, history, period) for period in QUARTERS}
        if any(baseline is None for baseline in baselines.values()):
            return []
        if isinstance(history, _HistoryView):
            index, start, end = history.index, history.start, history.end
        else:
            index = _IndexedHistory(history, self.config.recent_half_life_games)
            start, end = 0, len(index)
        quarter_stats = tuple(index.period(period) for period in QUARTERS)
        delay = timedelta(hours=self.config.result_available_after_hours)
        cfg_sig = self._baseline_config_signature()
        window = self.config.residual_window
        vectors: List[tuple[float, float, float, float]] = []

        if start == 0:
            stream = index._zero_start_quarter_vectors.get(cfg_sig)
            if stream is not None or end <= window * 2:
                if stream is None:
                    stream = [0, [], []]
                    index._zero_start_quarter_vectors[cfg_sig] = stream
                computed_upto, valid_indices, valid_vectors = stream
                while computed_upto < end:
                    vec = self._compute_quarter_vector(
                        index, quarter_stats, 0, computed_upto, delay
                    )
                    if vec is not None:
                        valid_indices.append(computed_upto)
                        valid_vectors.append(vec)
                    computed_upto += 1
                stream[0] = computed_upto
                pos = (
                    len(valid_indices)
                    if end >= computed_upto
                    else bisect_left(valid_indices, end)
                )
                vectors = valid_vectors[max(0, pos - window) : pos]
            else:
                vec_cache = index._quarter_vector_cache
                for target_index in range(end - 1, start - 1, -1):
                    cache_key = (start, target_index, cfg_sig)
                    if cache_key in vec_cache:
                        vec = vec_cache[cache_key]
                    else:
                        vec = self._compute_quarter_vector(
                            index, quarter_stats, start, target_index, delay
                        )
                        vec_cache[cache_key] = vec
                    if vec is not None:
                        vectors.append(vec)
                        if len(vectors) >= window:
                            break
                vectors.reverse()
        else:
            vec_cache = index._quarter_vector_cache
            for target_index in range(end - 1, start - 1, -1):
                cache_key = (start, target_index, cfg_sig)
                if cache_key in vec_cache:
                    vec = vec_cache[cache_key]
                else:
                    vec = self._compute_quarter_vector(
                        index, quarter_stats, start, target_index, delay
                    )
                    vec_cache[cache_key] = vec
                if vec is not None:
                    vectors.append(vec)
                    if len(vectors) >= window:
                        break
            vectors.reverse()
        if len(vectors) < self.config.min_residual_games:
            return []
        wins = {period.value: 0 for period in QUARTERS}
        for vector in vectors:
            simulated = [
                baselines[period].total + vector[index] for index, period in enumerate(QUARTERS)
            ]
            highest = max(simulated)
            for index, period in enumerate(QUARTERS):
                if simulated[index] == highest:
                    wins[period.value] += 1
        probabilities = {key: value / len(vectors) for key, value in wins.items()}
        selection = self._select(probabilities, probabilities.keys())
        if selection is None:
            return []
        return [
            Forecast(
                game_id=game.game_id,
                league=game.league.value,
                scheduled_utc=game.scheduled_utc,
                home_name=game.home_name,
                away_name=game.away_name,
                market=Market.HIGHEST_QUARTER.value,
                period=Period.FULL_GAME.value,
                selection=selection,
                raw_probability=probabilities[selection],
                alternative_probabilities=probabilities,
                source=MODEL_VERSION,
                calibration_version=self.calibration_version,
                training_games=len(history),
                residual_games=len(vectors),
                team_history=self._team_history(
                    [baseline for baseline in baselines.values() if baseline is not None]
                ),
                expected_total=sum(baselines[period].total for period in QUARTERS),
            )
        ]


def _book_key(snapshot: BookSnapshot) -> str:
    return (
        f"{snapshot.provider}|{snapshot.book}|{snapshot.market.value}|"
        f"{snapshot.period.value}|{'pre' if snapshot.pre_match else 'post'}"
    )


__all__ = [
    "Distribution",
    "ForecastConfig",
    "HistoryModel",
    "MODEL_VERSION",
    "QUARTERS",
]
