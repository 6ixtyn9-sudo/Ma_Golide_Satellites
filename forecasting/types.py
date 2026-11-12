"""Typed forecasting, evidence, grading, and calibration records."""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import math
from typing import Any, Dict, Optional

from ingestion.models import BookSnapshot, Game, Market, Period, PeriodKind

MODEL_VERSION = "team_history_empirical_v1"


@dataclass(frozen=True)
class BookEvidence:
    provider: str
    book: str
    observed_utc: str
    source_modified_utc: Optional[str]
    pre_match: bool
    main_line: bool
    frozen: bool
    prices: Dict[str, Optional[float]]
    game_id: Optional[str] = None
    market: Optional[str] = None
    period: Optional[str] = None
    line: Optional[float] = None

    @classmethod
    def from_snapshot(cls, snapshot: BookSnapshot, *, frozen: bool) -> "BookEvidence":
        return cls(
            provider=snapshot.provider,
            book=snapshot.book,
            observed_utc=snapshot.observed_utc,
            source_modified_utc=snapshot.source_modified_utc,
            pre_match=snapshot.pre_match,
            main_line=snapshot.main_line,
            frozen=frozen,
            prices={
                "home": snapshot.home_price_decimal,
                "away": snapshot.away_price_decimal,
                "draw": snapshot.draw_price_decimal,
                "over": snapshot.over_price_decimal,
                "under": snapshot.under_price_decimal,
            },
            game_id=snapshot.game_id,
            market=snapshot.market.value,
            period=snapshot.period.value,
            line=snapshot.line,
        )

    @property
    def key(self) -> str:
        return f"{self.provider}|{self.book}"

    def price_for(self, selection: str) -> Optional[float]:
        price = self.prices.get(selection)
        if price is None or not math.isfinite(price) or price <= 1.0:
            return None
        return float(price)


def opposite_selection(
    market: str | Market,
    period: str | Period,
    selection: str,
    *,
    has_draw: bool = False,
) -> Optional[str]:
    """Return a unique opposing selection, without deriving any price.

    First-half moneylines are three-way in this basketball model, even when a
    partial book snapshot has no draw quote. Full-game moneylines are two-way
    unless a draw quote explicitly identifies a three-way market.
    """
    market_value = market.value if isinstance(market, Market) else str(market)
    period_value = period.value if isinstance(period, Period) else str(period)
    if period_value not in {item.value for item in Period}:
        return None
    if market_value == Market.MONEYLINE.value:
        if period_value != Period.FULL_GAME.value or has_draw:
            return None
        return {"home": "away", "away": "home"}.get(selection)
    if market_value == Market.SPREAD.value:
        return {"home": "away", "away": "home"}.get(selection)
    if market_value == Market.TOTAL.value:
        return {"over": "under", "under": "over"}.get(selection)
    return None


@dataclass
class Forecast:
    game_id: str
    league: str
    scheduled_utc: str
    home_name: str
    away_name: str
    market: str
    period: str
    selection: str
    raw_probability: float
    alternative_probabilities: Dict[str, float]
    source: str
    line: Optional[float] = None
    push_probability: float = 0.0
    expected_margin: Optional[float] = None
    expected_total: Optional[float] = None
    training_games: int = 0
    residual_games: int = 0
    team_history: Optional[Dict[str, Any]] = None
    book: Optional[BookEvidence] = None
    prediction_provider: Optional[str] = None
    prediction_observed_utc: Optional[str] = None
    prediction_frozen: bool = False
    role: Optional[str] = None
    outcome: Optional[str] = None
    calibrated_probability: Optional[float] = None
    calibration: Dict[str, Any] = field(default_factory=dict)
    eligibility: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        for name, value in (
            ("raw_probability", self.raw_probability),
            ("push_probability", self.push_probability),
        ):
            if not math.isfinite(value) or not 0.0 <= value <= 1.0:
                raise ValueError(f"{name} must be a finite probability")
        if self.book is not None and not self.book.pre_match:
            raise ValueError("a forecast cannot use a post-tip book observation")

    @property
    def segment_key(self) -> tuple[str, str, str, str]:
        return (self.league, self.market, self.period, self.source)

    @property
    def market_key(self) -> tuple[str, str, str]:
        return (self.game_id, self.market, self.period)

    @property
    def forecast_id(self) -> str:
        raw = "|".join(
            [
                self.game_id,
                self.market,
                self.period,
                self.selection,
                "" if self.line is None else f"{self.line:g}",
                self.source,
                self.prediction_observed_utc or "",
                self.book.key if self.book else "no_book",
                self.book.observed_utc if self.book else "",
                MODEL_VERSION,
            ]
        )
        return hashlib.sha256(raw.encode()).hexdigest()[:20]

    @property
    def decimal_price(self) -> Optional[float]:
        return self.book.price_for(self.selection) if self.book else None

    @property
    def opposite_selection(self) -> Optional[str]:
        return opposite_selection(
            self.market,
            self.period,
            self.selection,
            has_draw=bool(self.book and self.book.prices.get("draw") is not None),
        )

    @property
    def opposite_decimal_price(self) -> Optional[float]:
        side = self.opposite_selection
        if (
            self.book is None
            or side is None
            or self.book.game_id != self.game_id
            or self.book.market != self.market
            or self.book.period != self.period
            or self.book.line != self.line
        ):
            return None
        return self.book.price_for(side)

    def to_dict(self) -> Dict[str, Any]:
        book = None
        if self.book is not None:
            book = {
                "provider": self.book.provider,
                "book": self.book.book,
                "observed_utc": self.book.observed_utc,
                "source_modified_utc": self.book.source_modified_utc,
                "pre_match": self.book.pre_match,
                "main_line": self.book.main_line,
                "frozen": self.book.frozen,
                "game_id": self.book.game_id,
                "market": self.book.market,
                "period": self.book.period,
                "line": self.book.line,
                "selected_decimal_price": self.decimal_price,
                "side_prices": dict(self.book.prices),
            }
        return {
            "forecast_id": self.forecast_id,
            "game_id": self.game_id,
            "league": self.league,
            "scheduled_utc": self.scheduled_utc,
            "home": self.home_name,
            "away": self.away_name,
            "market": self.market,
            "period": self.period,
            "selection": self.selection,
            "line": self.line,
            "role": self.role,
            "raw_probability": round(self.raw_probability, 6),
            "calibrated_probability": (
                round(self.calibrated_probability, 6)
                if self.calibrated_probability is not None
                else None
            ),
            "push_probability": round(self.push_probability, 6),
            "alternative_probabilities": {
                key: round(value, 6)
                for key, value in sorted(self.alternative_probabilities.items())
            },
            "expected_margin": self.expected_margin,
            "expected_total": self.expected_total,
            "model": self.source,
            "model_version": MODEL_VERSION,
            "training_games": self.training_games,
            "residual_games": self.residual_games,
            "team_history": (dict(self.team_history) if self.team_history is not None else None),
            "prediction": (
                {
                    "provider": self.prediction_provider,
                    "observed_utc": self.prediction_observed_utc,
                    "frozen": self.prediction_frozen,
                }
                if self.prediction_provider
                else None
            ),
            "book": book,
            "outcome": self.outcome,
            "calibration": self.calibration,
            "eligibility": self.eligibility,
        }


def _period_scores(game: Game, period: Period) -> Optional[tuple[float, float]]:
    if game.actual is None:
        return None
    actual = game.actual
    if period is Period.FULL_GAME:
        if actual.home_final is None or actual.away_final is None:
            return None
        return float(actual.home_final), float(actual.away_final)

    home: list[Optional[float]]
    away: list[Optional[float]]
    if period in {Period.H1, Period.H2}:
        if actual.period_kind is PeriodKind.HALVES:
            if len(actual.home_quarters) != 2 or len(actual.away_quarters) != 2:
                return None
            index = 0 if period is Period.H1 else 1
            home = [actual.home_quarters[index]]
            away = [actual.away_quarters[index]]
        elif actual.period_kind is PeriodKind.QUARTERS:
            if len(actual.home_quarters) != 4 or len(actual.away_quarters) != 4:
                return None
            bounds = (0, 2) if period is Period.H1 else (2, 4)
            home = actual.home_quarters[slice(*bounds)]
            away = actual.away_quarters[slice(*bounds)]
        else:
            return None
    elif period in {Period.Q1, Period.Q2, Period.Q3, Period.Q4}:
        if (
            actual.period_kind is not PeriodKind.QUARTERS
            or len(actual.home_quarters) != 4
            or len(actual.away_quarters) != 4
        ):
            return None
        index = int(period.value[1]) - 1
        home = [actual.home_quarters[index]]
        away = [actual.away_quarters[index]]
    else:
        return None
    if any(value is None for value in [*home, *away]):
        return None
    return float(sum(home)), float(sum(away))


def grade_forecast(forecast: Forecast, game: Game) -> Optional[str]:
    """Grade a frozen forecast against an already reconciled official final."""
    if game.actual is None or not game.actual.reconciles()[0]:
        return None
    market = Market(forecast.market)
    period = Period(forecast.period)
    scores = _period_scores(game, period)
    if market is Market.HIGHEST_QUARTER:
        quarters = []
        for quarter in (Period.Q1, Period.Q2, Period.Q3, Period.Q4):
            score = _period_scores(game, quarter)
            if score is None:
                return None
            quarters.append(sum(score))
        highest = max(quarters)
        index = int(forecast.selection[1]) - 1
        # This is explicitly a highest-or-tied research forecast. Book rules
        # can settle dead heats differently, so unpriced HQ is not acca eligible.
        return "win" if quarters[index] == highest else "loss"
    if scores is None:
        return None
    home, away = scores
    if market is Market.MONEYLINE:
        actual = "home" if home > away else "away" if away > home else "draw"
        return "win" if forecast.selection == actual else "loss"
    if forecast.line is None:
        return None
    if market is Market.TOTAL:
        difference = home + away - forecast.line
        desired = 1 if forecast.selection == "over" else -1
    elif market is Market.SPREAD:
        home_cover = home + forecast.line - away
        difference = home_cover if forecast.selection == "home" else -home_cover
        desired = 1
    else:
        return None
    if abs(difference) <= 1e-9:
        return "push"
    return "win" if difference * desired > 0 else "loss"


__all__ = [
    "BookEvidence",
    "Forecast",
    "MODEL_VERSION",
    "grade_forecast",
    "opposite_selection",
]
