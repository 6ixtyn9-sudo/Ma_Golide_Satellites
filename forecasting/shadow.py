"""Recency-decayed calibration challenger — shadow evaluation only.

The champion (``harness.build_profiles`` / ``apply_profiles``) is untouched:
production probabilities, acca selection, and eligibility gates never read
anything from this module. The challenger maintains per-segment exponentially
recency-decayed reliability adjustments and evaluates them prequentially: for
each graded prediction, its profile uses only outcomes whose results were
available by that prediction's tip time under the repository's availability
rule (``scheduled_utc + result_delay_hours <= tip``), mirroring
``ForecastConfig.result_available_after_hours`` and the chronological
walk-forward rules in ``service.MarketForecastService``.

Promotion policy: this module is diagnostic. Its parameters are fixed a priori
and were not tuned on any data a future claim would use. A future promotion of
a recency-decayed profile to champion requires untouched prospective evidence
accrued after the shadow shipped — the prequential numbers reported here
describe the past and confer no right to promote.

Weights are anchored to the event tip: a win/loss scheduled at ``s`` carries
weight ``0.5 ** ((t - s) / half_life)`` at evaluation time ``t``. It enters the
profile only once its result is available at ``s + result_delay_hours``. The
incremental evaluator compensates at ingestion so its weights match this
reference convention. Effective sample size uses Kish's formula
``(Σw)² / Σw²``, at most the raw count for weights in ``(0, 1]``. The
prequential pass is incremental (O(rows)) and cross-checked against the naive
rebuild in the test suite.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import timedelta
import math
from typing import Any, Dict, Iterable, List, Optional, Tuple

from ingestion.identity import parse_utc

from .harness import CALIBRATION_EPOCH_UTC, CalibrationConfig, _log_loss, config_fingerprint
from .types import Forecast

CHALLENGER_TAG = "recency_v1"

_DAY_SECONDS = 86400.0


@dataclass(frozen=True)
class RecencyConfig:
    """Challenger parameters, fixed a priori (never tuned on claimed evidence)."""

    half_life_days: float = 14.0
    prior_strength: float = 30.0
    min_samples: float = 30.0
    result_delay_hours: float = 4.0
    probability_floor: float = 0.01
    probability_ceiling: float = 0.99

    def __post_init__(self) -> None:
        values = (
            self.half_life_days,
            self.prior_strength,
            self.min_samples,
            self.result_delay_hours,
            self.probability_floor,
            self.probability_ceiling,
        )
        if any(not math.isfinite(value) for value in values):
            raise ValueError("recency configuration values must be finite")
        if self.half_life_days <= 0 or self.prior_strength <= 0:
            raise ValueError("half-life and prior strength must be positive")
        if self.min_samples <= 0:
            raise ValueError("minimum weighted sample must be positive")
        if self.result_delay_hours < 0:
            raise ValueError("result delay cannot be negative")
        if not 0 < self.probability_floor < self.probability_ceiling < 1:
            raise ValueError("invalid probability bounds")

    @property
    def version_tag(self) -> str:
        """Version-aware: any behaviour-relevant change yields a distinct tag."""
        fingerprint = config_fingerprint(
            {
                "half_life_days": self.half_life_days,
                "prior_strength": self.prior_strength,
                "min_samples": self.min_samples,
                "result_delay_hours": self.result_delay_hours,
                "probability_floor": self.probability_floor,
                "probability_ceiling": self.probability_ceiling,
            }
        )
        return f"{CHALLENGER_TAG}-{fingerprint}"


@dataclass
class RecencyProfile:
    key: Tuple[str, ...]
    n: int
    pushes: int
    weighted_n: float
    weighted_wins: float
    sum_weight_squared: float
    weighted_mean_raw: float
    adjustment: float
    brier_score: float
    log_loss: float
    min_samples: float
    half_life_days: float

    @property
    def ess(self) -> float:
        if self.sum_weight_squared <= 0:
            return 0.0
        return self.weighted_n * self.weighted_n / self.sum_weight_squared

    @property
    def sufficient(self) -> bool:
        return self.weighted_n >= self.min_samples

    def apply(self, probability: float, config: RecencyConfig) -> float:
        adjusted = probability + self.adjustment
        return min(config.probability_ceiling, max(config.probability_floor, adjusted))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "league": self.key[0],
            "market": self.key[1],
            "period": self.key[2],
            "source": self.key[3],
            "calibration_version": self.key[4] if len(self.key) > 4 else "legacy",
            "challenger": CHALLENGER_TAG,
            "challenger_version": self.key[-1],
            "n": self.n,
            "pushes": self.pushes,
            "weighted_n": round(self.weighted_n, 6),
            "effective_sample_size": round(self.ess, 6),
            "weighted_win_rate": (
                round(self.weighted_wins / self.weighted_n, 6) if self.weighted_n > 0 else None
            ),
            "weighted_mean_raw_probability": round(self.weighted_mean_raw, 6),
            "probability_adjustment": round(self.adjustment, 6),
            "brier_score": round(self.brier_score, 6),
            "log_loss": round(self.log_loss, 6),
            "minimum_weighted_samples": self.min_samples,
            "sufficient": self.sufficient,
            "half_life_days": self.half_life_days,
        }


def _tip(record: Forecast):
    try:
        return parse_utc(record.scheduled_utc)
    except Exception:
        return None


def _eligible_evidence(record: Forecast, epoch) -> bool:
    """Mirror the champion's evidence filters (epoch, team-history pool)."""
    if (record.team_history or {}).get("sufficient_for_accas") is False:
        return False
    tip = _tip(record)
    return tip is not None and tip >= epoch


def build_recency_profiles(
    records: Iterable[Forecast],
    config: RecencyConfig | None = None,
    *,
    as_of_utc: str,
) -> Dict[Tuple[str, ...], RecencyProfile]:
    """Naive reference builder: weights relative to ``as_of_utc``.

    Only graded outcomes whose results were available by ``as_of_utc`` under
    the repository's delay rule are used. O(rows) per call; the prequential
    pass uses an incremental equivalent (cross-checked in tests).
    """
    cfg = config or RecencyConfig()
    as_of = parse_utc(as_of_utc)
    epoch = parse_utc(CALIBRATION_EPOCH_UTC)
    delay = timedelta(hours=cfg.result_delay_hours)
    grouped: Dict[Tuple[str, ...], List[Forecast]] = {}
    for record in records:
        if record.outcome not in {"win", "loss", "push"}:
            continue
        tip = _tip(record)
        if not _eligible_evidence(record, epoch):
            continue
        if tip + delay > as_of:
            continue
        grouped.setdefault((*record.versioned_segment_key, cfg.version_tag), []).append(record)
    profiles: Dict[Tuple[str, ...], RecencyProfile] = {}
    for key, rows in grouped.items():
        graded = [row for row in rows if row.outcome in {"win", "loss"}]
        if not graded:
            continue
        weights = [
            0.5 ** ((as_of - _tip(row)).total_seconds() / _DAY_SECONDS / cfg.half_life_days)
            for row in graded
        ]
        weighted_n = sum(weights)
        weighted_wins = sum(w for w, row in zip(weights, graded) if row.outcome == "win")
        weighted_mean_raw = (
            sum(w * row.raw_probability for w, row in zip(weights, graded)) / weighted_n
        )
        observed = weighted_wins / weighted_n
        shrink = weighted_n / (weighted_n + cfg.prior_strength)
        n = len(graded)
        profiles[key] = RecencyProfile(
            key=key,
            n=n,
            pushes=sum(row.outcome == "push" for row in rows),
            weighted_n=weighted_n,
            weighted_wins=weighted_wins,
            sum_weight_squared=sum(w * w for w in weights),
            weighted_mean_raw=weighted_mean_raw,
            adjustment=shrink * (observed - weighted_mean_raw),
            brier_score=sum(
                (row.raw_probability - float(row.outcome == "win")) ** 2 for row in graded
            )
            / n,
            log_loss=sum(_log_loss(row.raw_probability, row.outcome == "win") for row in graded)
            / n,
            min_samples=cfg.min_samples,
            half_life_days=cfg.half_life_days,
        )
    return profiles


def prequential_shadow_evaluation(
    records: Iterable[Forecast], config: RecencyConfig | None = None
) -> Dict[str, Any]:
    """Shadow-only prequential evaluation of the recency-decayed challenger.

    Walks graded predictions in scheduled order; before scoring each one, the
    challenger (and a champion mirror) see only outcomes available under the
    repository's delay rule. Mutates nothing except the additive
    ``Forecast.shadow_calibration`` field — ``calibrated_probability``,
    ``calibration``, and ``eligibility`` are never touched, so acca inputs and
    selection cannot change.
    """
    cfg = config or RecencyConfig()
    champion_cfg = CalibrationConfig()
    epoch = parse_utc(CALIBRATION_EPOCH_UTC)
    delay = timedelta(hours=cfg.result_delay_hours)
    # Ingest at the weight that makes the incremental normalization equivalent
    # to the naive tip-anchored convention: a record ingested at its
    # availability time (tip + delay) carries w0 so that, after decaying to any
    # later read time t, its weight is 0.5 ** ((t - tip) / half_life).
    ingest_weight = (
        0.5 ** ((cfg.result_delay_hours * 3600.0) / _DAY_SECONDS / cfg.half_life_days)
        if cfg.half_life_days > 0
        else 1.0
    )

    pool: List[Tuple[Any, Forecast]] = []
    ungraded = 0
    for record in records:
        if record.outcome not in {"win", "loss", "push"}:
            ungraded += 1
            continue
        tip = _tip(record)
        if tip is None or not _eligible_evidence(record, epoch):
            ungraded += 1
            continue
        pool.append((tip, record))
    pool.sort(key=lambda item: (item[0], item[1].forecast_id))
    availability = sorted(
        (tip + delay, record.forecast_id, record) for tip, record in pool
    )
    evaluated = [(tip, record) for tip, record in pool if record.outcome in {"win", "loss"}]

    # Champion mirror: all available history, no decay (harness semantics).
    champion: Dict[Tuple[str, ...], List[float]] = {}
    # Challenger: recency-decayed sums normalized at `last_norm`.
    challenger: Dict[Tuple[str, ...], List[float]] = {}

    def _ingest(when, record: Forecast) -> None:
        ckey = record.versioned_segment_key
        stats = champion.setdefault(ckey, [0.0, 0.0, 0.0, 0.0])  # n, wins, sum_raw, pushes
        if record.outcome == "push":
            stats[3] += 1.0
            return
        stats[0] += 1.0
        if record.outcome == "win":
            stats[1] += 1.0
        stats[2] += record.raw_probability
        rkey = (*ckey, cfg.version_tag)
        state = challenger.get(rkey)
        if state is None:
            state = challenger[rkey] = [0.0, 0.0, 0.0, 0.0, 0.0, when]  # W, W2, Sx, Sp, n, norm
        else:
            norm = state[5]
            if when > norm:
                factor = 0.5 ** ((when - norm).total_seconds() / _DAY_SECONDS / cfg.half_life_days)
                state[0] *= factor
                state[1] *= factor
                state[2] *= factor
                state[3] *= factor
                state[5] = when
        state[0] += ingest_weight
        state[1] += ingest_weight * ingest_weight
        state[2] += ingest_weight if record.outcome == "win" else 0.0
        state[3] += ingest_weight * record.raw_probability
        state[4] += 1.0

    def _champion_state(ckey):
        stats = champion.get(ckey)
        if stats is None or stats[0] <= 0:
            return None
        n = stats[0]
        mean_raw = stats[2] / n
        observed = stats[1] / n
        adjustment = (n / (n + champion_cfg.prior_strength)) * (observed - mean_raw)
        return {
            "n": int(n),
            "adjustment": adjustment,
            "status": "sufficient" if n >= champion_cfg.min_samples else "provisional",
        }

    def _challenger_state(rkey, at):
        state = challenger.get(rkey)
        if state is None or state[0] <= 0:
            return None
        norm = state[5]
        # Layout: [Σw, Σw², Σw·y, Σw·raw, raw n, normalized_at].
        weighted_n = state[0]
        sum_weight_squared = state[1]
        weighted_wins = state[2]
        weighted_raw_sum = state[3]
        if at > norm:
            factor = 0.5 ** ((at - norm).total_seconds() / _DAY_SECONDS / cfg.half_life_days)
            weighted_n *= factor
            sum_weight_squared *= factor
            weighted_wins *= factor
            weighted_raw_sum *= factor
            state[0] = weighted_n
            state[1] = sum_weight_squared
            state[2] = weighted_wins
            state[3] = weighted_raw_sum
            state[5] = at
        if weighted_n <= 0:
            return None
        observed = weighted_wins / weighted_n
        mean_raw = weighted_raw_sum / weighted_n
        shrink = weighted_n / (weighted_n + cfg.prior_strength)
        ess = (
            weighted_n * weighted_n / sum_weight_squared if sum_weight_squared > 0 else 0.0
        )
        return {
            "n": int(state[4]),
            "weighted_n": weighted_n,
            "ess": ess,
            "adjustment": shrink * (observed - mean_raw),
            "status": "sufficient" if weighted_n >= cfg.min_samples else "provisional",
        }

    def _bounded(probability: float, adjustment: float) -> float:
        return min(cfg.probability_ceiling, max(cfg.probability_floor, probability + adjustment))

    index = 0
    raw_brier = raw_log = champ_brier = champ_log = chall_brier = chall_log = 0.0
    coverage = {
        "champion": {"no_history": 0, "provisional": 0, "sufficient": 0},
        "challenger": {"no_history": 0, "provisional": 0, "sufficient": 0},
    }
    ess_total = weighted_total = challenger_n_total = 0.0
    for tip, record in evaluated:
        while index < len(availability) and availability[index][0] <= tip:
            _ingest(availability[index][0], availability[index][2])
            index += 1
        won = record.outcome == "win"
        target = float(won)
        raw_brier += (record.raw_probability - target) ** 2
        raw_log += _log_loss(record.raw_probability, won)

        champ_state = _champion_state(record.versioned_segment_key)
        if champ_state is None:
            champ_probability = record.raw_probability
            champ_adjustment = 0.0
            coverage["champion"]["no_history"] += 1
        else:
            champ_probability = _bounded(
                record.raw_probability, champ_state["adjustment"]
            )
            champ_adjustment = champ_state["adjustment"]
            coverage["champion"][champ_state["status"]] += 1
        champ_brier += (champ_probability - target) ** 2
        champ_log += _log_loss(champ_probability, won)

        chall_state = _challenger_state(
            (*record.versioned_segment_key, cfg.version_tag), tip
        )
        if chall_state is None:
            chall_probability = record.raw_probability
            shadow = {
                "status": "no_history",
                "n": 0,
                "weighted_n": 0.0,
                "ess": 0.0,
                "adjustment": 0.0,
                "probability": record.raw_probability,
                "champion_status": (
                    "no_history" if champ_state is None else champ_state["status"]
                ),
                "champion_n": 0 if champ_state is None else champ_state["n"],
                "champion_adjustment": champ_adjustment,
            }
            coverage["challenger"]["no_history"] += 1
        else:
            chall_probability = _bounded(record.raw_probability, chall_state["adjustment"])
            shadow = {
                "status": chall_state["status"],
                "n": chall_state["n"],
                "weighted_n": round(chall_state["weighted_n"], 6),
                "ess": round(chall_state["ess"], 6),
                "adjustment": round(chall_state["adjustment"], 6),
                "probability": chall_probability,
                "champion_status": (
                    "no_history" if champ_state is None else champ_state["status"]
                ),
                "champion_n": 0 if champ_state is None else champ_state["n"],
                "champion_adjustment": champ_adjustment,
            }
            coverage["challenger"][chall_state["status"]] += 1
            ess_total += chall_state["ess"]
            weighted_total += chall_state["weighted_n"]
            challenger_n_total += chall_state["n"]
        record.shadow_calibration = shadow
        chall_brier += (chall_probability - target) ** 2
        chall_log += _log_loss(chall_probability, won)

    count = len(evaluated)

    def _mean(total: float) -> Optional[float]:
        return round(total / count, 6) if count else None

    def _metric(total_brier: float, total_log: float) -> Dict[str, Optional[float]]:
        return {"brier_score": _mean(total_brier), "log_loss": _mean(total_log)}

    return {
        "challenger_version": cfg.version_tag,
        "rows_input": len(pool) + ungraded,
        "rows_ungraded_or_ineligible_skipped": ungraded,
        "rows_push_skipped": len(pool) - count,
        "rows_evaluated": count,
        "raw_metrics": _metric(raw_brier, raw_log),
        "champion": {**_metric(champ_brier, champ_log), "coverage": coverage["champion"]},
        "challenger": {
            **_metric(chall_brier, chall_log),
            "coverage": coverage["challenger"],
            "mean_raw_n": _mean(challenger_n_total),
            "mean_weighted_n": _mean(weighted_total),
            "mean_effective_sample_size": _mean(ess_total),
        },
        "promotion_policy": (
            "shadow-only; parameters fixed a priori; promotion requires "
            "untouched prospective evidence accrued after deployment"
        ),
    }


__all__ = [
    "CHALLENGER_TAG",
    "RecencyConfig",
    "RecencyProfile",
    "build_recency_profiles",
    "prequential_shadow_evaluation",
]
