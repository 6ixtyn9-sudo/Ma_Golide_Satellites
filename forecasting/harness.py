"""Prequential market calibration and acca-eligibility policy."""

from __future__ import annotations

from dataclasses import dataclass
import math
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from assayer.assayer_engine import wilson_lower_bound

from .types import Forecast

SegmentKey = Tuple[str, str, str, str]


@dataclass(frozen=True)
class CalibrationConfig:
    prior_strength: float = 30.0
    min_samples: int = 30
    min_probability: float = 0.50
    max_push_probability: float = 0.005
    probability_floor: float = 0.01
    probability_ceiling: float = 0.99

    def __post_init__(self) -> None:
        if self.prior_strength <= 0 or self.min_samples < 1:
            raise ValueError("calibration prior and sample minimum must be positive")
        if not 0 <= self.min_probability <= 1:
            raise ValueError("minimum probability must be between zero and one")
        if not 0 <= self.max_push_probability < 1:
            raise ValueError("push tolerance must be in [0, 1)")
        if not 0 < self.probability_floor < self.probability_ceiling < 1:
            raise ValueError("invalid calibrated probability bounds")


@dataclass(frozen=True)
class CalibrationProfile:
    key: SegmentKey
    wins: int
    losses: int
    pushes: int
    mean_raw_probability: float
    observed_win_rate: float
    adjustment: float
    wilson_lower_95: float
    brier_score: float
    log_loss: float
    min_samples: int

    @property
    def n(self) -> int:
        return self.wins + self.losses

    @property
    def sufficient(self) -> bool:
        return self.n >= self.min_samples

    def calibrate(self, probability: float, config: CalibrationConfig) -> float:
        adjusted = probability + self.adjustment
        return min(config.probability_ceiling, max(config.probability_floor, adjusted))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "league": self.key[0],
            "market": self.key[1],
            "period": self.key[2],
            "source": self.key[3],
            "n": self.n,
            "wins": self.wins,
            "losses": self.losses,
            "pushes": self.pushes,
            "mean_raw_probability": round(self.mean_raw_probability, 6),
            "observed_win_rate": round(self.observed_win_rate, 6),
            "probability_adjustment": round(self.adjustment, 6),
            "wilson_lower_95": round(self.wilson_lower_95, 6),
            "brier_score": round(self.brier_score, 6),
            "log_loss": round(self.log_loss, 6),
            "minimum_samples": self.min_samples,
            "sufficient_for_accas": self.sufficient,
        }


def _log_loss(probability: float, won: bool) -> float:
    bounded = min(1 - 1e-12, max(1e-12, probability))
    return -math.log(bounded if won else 1 - bounded)


def build_profiles(
    records: Iterable[Forecast], config: CalibrationConfig | None = None
) -> Dict[SegmentKey, CalibrationProfile]:
    """Fit additive reliability adjustments from graded past forecasts only."""
    cfg = config or CalibrationConfig()
    grouped: Dict[SegmentKey, List[Forecast]] = {}
    for record in records:
        if record.outcome in {"win", "loss", "push"}:
            grouped.setdefault(record.segment_key, []).append(record)
    profiles: Dict[SegmentKey, CalibrationProfile] = {}
    for key, rows in grouped.items():
        graded = [row for row in rows if row.outcome in {"win", "loss"}]
        pushes = sum(row.outcome == "push" for row in rows)
        n = len(graded)
        if not n:
            continue
        wins = sum(row.outcome == "win" for row in graded)
        raw_mean = sum(row.raw_probability for row in graded) / n
        observed = wins / n
        # Small samples stay near the model's own mean confidence. The
        # correction approaches observed-minus-claimed only as evidence grows.
        weight = n / (n + cfg.prior_strength)
        adjustment = weight * (observed - raw_mean)
        profiles[key] = CalibrationProfile(
            key=key,
            wins=wins,
            losses=n - wins,
            pushes=pushes,
            mean_raw_probability=raw_mean,
            observed_win_rate=observed,
            adjustment=adjustment,
            wilson_lower_95=wilson_lower_bound(wins, n),
            brier_score=sum(
                (row.raw_probability - float(row.outcome == "win")) ** 2
                for row in graded
            )
            / n,
            log_loss=sum(
                _log_loss(row.raw_probability, row.outcome == "win")
                for row in graded
            )
            / n,
            min_samples=cfg.min_samples,
        )
    return profiles


def apply_profiles(
    forecasts: Iterable[Forecast],
    profiles: Mapping[SegmentKey, CalibrationProfile],
    config: CalibrationConfig | None = None,
) -> List[Forecast]:
    cfg = config or CalibrationConfig()
    calibrated: List[Forecast] = []
    for forecast in forecasts:
        profile = profiles.get(forecast.segment_key)
        if profile is None:
            probability = forecast.raw_probability
            forecast.calibration = {
                "status": "no_history",
                "n": 0,
                "adjustment": 0.0,
                "sufficient_for_accas": False,
            }
        else:
            probability = profile.calibrate(forecast.raw_probability, cfg)
            forecast.calibration = profile.to_dict()
            forecast.calibration["status"] = (
                "calibrated" if profile.sufficient else "provisional"
            )
        forecast.calibrated_probability = probability
        reasons: List[str] = []
        if forecast.book is None:
            reasons.append("no_exact_bookmaker_market")
        elif forecast.decimal_price is None:
            reasons.append("missing_selected_decimal_price")
        if profile is None or not profile.sufficient:
            reasons.append("insufficient_walk_forward_samples")
        if forecast.push_probability > cfg.max_push_probability:
            reasons.append("push_risk_requires_leg_void_rules")
        if probability < cfg.min_probability:
            reasons.append("below_calibrated_probability_floor")
        forecast.eligibility = {
            "acca_eligible": not reasons,
            "reasons": reasons,
            "evidence_state": (
                "frozen_pre_tip"
                if forecast.book is not None and forecast.book.frozen
                else "current_pre_tip_unfrozen"
            ),
        }
        calibrated.append(forecast)
    return calibrated


def probability_metrics(
    records: Sequence[Forecast], *, calibrated: bool = False
) -> Dict[str, Optional[float] | int]:
    graded = [row for row in records if row.outcome in {"win", "loss"}]
    if not graded:
        return {"n": 0, "brier_score": None, "log_loss": None, "accuracy": None}
    probabilities = []
    for row in graded:
        value = row.calibrated_probability if calibrated else row.raw_probability
        if value is None:
            value = row.raw_probability
        probabilities.append(value)
    return {
        "n": len(graded),
        "brier_score": sum(
            (probability - float(row.outcome == "win")) ** 2
            for probability, row in zip(probabilities, graded)
        )
        / len(graded),
        "log_loss": sum(
            _log_loss(probability, row.outcome == "win")
            for probability, row in zip(probabilities, graded)
        )
        / len(graded),
        "accuracy": sum(row.outcome == "win" for row in graded) / len(graded),
    }


__all__ = [
    "CalibrationConfig",
    "CalibrationProfile",
    "SegmentKey",
    "apply_profiles",
    "build_profiles",
    "probability_metrics",
]
