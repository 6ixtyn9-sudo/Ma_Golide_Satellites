"""Leakage-conscious empirical basketball market forecasts.

Team scoring baselines use only supplied history. Residual distributions are
built from historical baselines that themselves used only earlier games. Book
lines and prices remain immutable evidence attached to each forecast.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import timedelta
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from ingestion.identity import canonical_team_key, parse_utc
from ingestion.models import BookSnapshot, Game, Market, Period, Prediction

from .types import BookEvidence, Forecast, MODEL_VERSION, _period_scores

QUARTERS = (Period.Q1, Period.Q2, Period.Q3, Period.Q4)


@dataclass(frozen=True)
class ForecastConfig:
    min_history_games: int = 20
    min_residual_games: int = 20
    team_prior_games: float = 8.0
    recent_half_life_games: float = 20.0
    training_window: int = 200
    residual_window: int = 100
    result_available_after_hours: int = 4

    def __post_init__(self) -> None:
        if self.min_history_games < 1 or self.min_residual_games < 1:
            raise ValueError("history and residual minimums must be positive")
        if self.team_prior_games <= 0 or self.recent_half_life_games <= 0:
            raise ValueError("shrinkage and half-life must be positive")
        if self.training_window < self.min_history_games:
            raise ValueError("training window is smaller than minimum history")
        if self.residual_window < self.min_residual_games:
            raise ValueError("residual window is smaller than minimum residual count")
        if self.result_available_after_hours < 0:
            raise ValueError("result availability lag cannot be negative")


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


class HistoryModel:
    def __init__(self, config: ForecastConfig | None = None) -> None:
        self.config = config or ForecastConfig()

    def available_history(
        self, games: Iterable[Game], *, scheduled_utc: str, as_of_utc: str
    ) -> List[Game]:
        """Return finals conservatively available before a forecast is made."""
        cutoff = min(parse_utc(scheduled_utc), parse_utc(as_of_utc)) - timedelta(
            hours=self.config.result_available_after_hours
        )
        eligible = [
            game
            for game in games
            if game.actual is not None
            and game.actual.reconciles()[0]
            and parse_utc(game.scheduled_utc) <= cutoff
        ]
        eligible.sort(key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))
        return eligible[-self.config.training_window :]

    @staticmethod
    def _team_key(name: str) -> str:
        return canonical_team_key(name)

    def _league_rates(
        self, history: Sequence[Game], period: Period
    ) -> Optional[tuple[float, float]]:
        home_scores: List[float] = []
        away_scores: List[float] = []
        for game in history:
            scores = _period_scores(game, period)
            if scores is not None:
                home_scores.append(scores[0])
                away_scores.append(scores[1])
        if not home_scores:
            return None
        per_team = (sum(home_scores) + sum(away_scores)) / (2 * len(home_scores))
        home_advantage = sum(home - away for home, away in zip(home_scores, away_scores)) / len(
            home_scores
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

    def baseline(self, game: Game, history: Sequence[Game], period: Period) -> Optional[Baseline]:
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
        expected_home = (
            shrunk(home_rates, "points_for") + shrunk(away_rates, "points_against")
        ) / 2 + home_advantage / 2
        expected_away = (
            shrunk(away_rates, "points_for") + shrunk(home_rates, "points_against")
        ) / 2 - home_advantage / 2
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

    def _residuals(
        self, history: Sequence[Game], period: Period
    ) -> tuple[List[float], List[float]]:
        margins: List[float] = []
        totals: List[float] = []
        for index, target in enumerate(history):
            cutoff = parse_utc(target.scheduled_utc) - timedelta(
                hours=self.config.result_available_after_hours
            )
            priors = [game for game in history[:index] if parse_utc(game.scheduled_utc) <= cutoff]
            baseline = self.baseline(target, priors, period)
            actual = _period_scores(target, period)
            if baseline is None or actual is None:
                continue
            margins.append(actual[0] - actual[1] - baseline.margin)
            totals.append(actual[0] + actual[1] - baseline.total)
        margins = margins[-self.config.residual_window :]
        totals = totals[-self.config.residual_window :]
        return margins, totals

    def distribution(
        self,
        game: Game,
        history: Sequence[Game],
        period: Period,
        prediction: Optional[Prediction] = None,
    ) -> Optional[Distribution]:
        baseline = self.baseline(game, history, period)
        if baseline is None:
            return None
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
        margin_residuals, total_residuals = self._residuals(history, period)
        if len(margin_residuals) < self.config.min_residual_games:
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
                    book=evidence,
                    prediction_provider=prediction.provider,
                    prediction_observed_utc=prediction.observed_utc,
                    prediction_frozen=prediction_frozen,
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
    ) -> List[Forecast]:
        distribution = self.distribution(game, history, period)
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
                source=MODEL_VERSION,
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

    def lines(
        self,
        game: Game,
        history: Sequence[Game],
        snapshots: Iterable[BookSnapshot],
        frozen_books: Dict[str, bool],
        prediction: Optional[Prediction] = None,
        prediction_frozen: bool = False,
    ) -> List[Forecast]:
        forecasts: List[Forecast] = []
        distributions: Dict[Period, Optional[Distribution]] = {}
        for snapshot in snapshots:
            if snapshot.line is None or snapshot.market not in {
                Market.TOTAL,
                Market.SPREAD,
            }:
                continue
            if snapshot.period not in distributions:
                distributions[snapshot.period] = self.distribution(
                    game, history, snapshot.period, prediction
                )
            distribution = distributions[snapshot.period]
            if distribution is None:
                continue
            if snapshot.market is Market.TOTAL:
                line = snapshot.line
                probabilities = {
                    "over": sum(value > line for value in distribution.totals)
                    / distribution.residuals,
                    "under": sum(value < line for value in distribution.totals)
                    / distribution.residuals,
                }
                push = sum(value == line for value in distribution.totals) / (
                    distribution.residuals
                )
            else:
                adjusted = [value + snapshot.line for value in distribution.margins]
                probabilities = {
                    "home": sum(value > 0 for value in adjusted) / distribution.residuals,
                    "away": sum(value < 0 for value in adjusted) / distribution.residuals,
                }
                push = sum(value == 0 for value in adjusted) / distribution.residuals
            selection = self._select(probabilities, probabilities.keys())
            if selection is None:
                continue
            source = MODEL_VERSION
            if (
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
                )
            )
        return forecasts

    def highest_quarter(self, game: Game, history: Sequence[Game]) -> List[Forecast]:
        if len(history) < self.config.min_history_games:
            return []
        baselines = {period: self.baseline(game, history, period) for period in QUARTERS}
        if any(baseline is None for baseline in baselines.values()):
            return []
        vectors: List[tuple[float, float, float, float]] = []
        for index, target in enumerate(history):
            cutoff = parse_utc(target.scheduled_utc) - timedelta(
                hours=self.config.result_available_after_hours
            )
            priors = [
                prior for prior in history[:index] if parse_utc(prior.scheduled_utc) <= cutoff
            ]
            vector: List[float] = []
            for period in QUARTERS:
                baseline = self.baseline(target, priors, period)
                actual = _period_scores(target, period)
                if baseline is None or actual is None:
                    break
                vector.append(sum(actual) - baseline.total)
            if len(vector) == 4:
                vectors.append(tuple(vector))
        vectors = vectors[-self.config.residual_window :]
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
