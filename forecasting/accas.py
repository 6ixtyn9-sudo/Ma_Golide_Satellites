"""Highest-estimated-probability acca construction from calibrated legs."""

from __future__ import annotations

import hashlib
import heapq
import math
from typing import Dict, Iterable, List, Tuple

from .types import Forecast


def is_provisional_sample_only(forecast: Forecast) -> bool:
    """Whether sample size is the forecast's sole remaining acca blocker.

    This is a narrow paper-exploration exception, not a general eligibility
    override. At least one graded non-push observation must exist, and the
    ordinary eligibility reasons must contain no other blocker.
    """
    eligibility = forecast.eligibility or {}
    if eligibility.get("acca_eligible") is not False:
        return False
    reasons = eligibility.get("reasons")
    if not isinstance(reasons, (list, tuple, set)):
        return False
    if set(reasons) != {"insufficient_walk_forward_samples"}:
        return False
    calibration = forecast.calibration
    if not isinstance(calibration, dict):
        return False
    if (
        calibration.get("status") != "provisional"
        or calibration.get("sufficient_for_accas") is not False
    ):
        return False
    n = calibration.get("n")
    if isinstance(n, bool) or not isinstance(n, int) or n < 1:
        return False
    minimum = calibration.get("minimum_samples")
    if minimum is not None and (
        isinstance(minimum, bool)
        or not isinstance(minimum, int)
        or minimum < 1
        or n >= minimum
    ):
        return False
    return True


def _best_leg_per_game(
    forecasts: Iterable[Forecast], *, allow_provisional: bool = False
) -> List[Forecast]:
    grouped: Dict[str, List[Forecast]] = {}
    for forecast in forecasts:
        if not forecast.eligibility.get("acca_eligible"):
            if not allow_provisional or not is_provisional_sample_only(forecast):
                continue
        probability = forecast.calibrated_probability
        price = forecast.decimal_price
        if (
            probability is None
            or not math.isfinite(probability)
            or not 0.0 < probability <= 1.0
            or price is None
            or not math.isfinite(price)
            or price <= 1.0
        ):
            continue
        grouped.setdefault(forecast.game_id, []).append(forecast)
    best = []
    for rows in grouped.values():
        rows.sort(
            key=lambda row: (
                -(row.calibrated_probability or 0.0),
                row.market,
                row.period,
                row.source,
                row.selection,
                row.book.key if row.book else "",
                row.forecast_id,
            )
        )
        best.append(rows[0])
    # Price is payout evidence, never a ranking signal. Stable identities are
    # the only tie-breakers after calibrated probability.
    best.sort(
        key=lambda row: (
            -(row.calibrated_probability or 0.0),
            row.game_id,
            row.market,
            row.period,
            row.source,
            row.forecast_id,
        )
    )
    return best


def _neighbors(state: Tuple[int, ...], population: int) -> Iterable[Tuple[int, ...]]:
    for position in range(len(state)):
        candidate = list(state)
        candidate[position] += 1
        if candidate[position] >= population:
            continue
        if position + 1 < len(state) and candidate[position] >= candidate[position + 1]:
            continue
        yield tuple(candidate)


def build_accas(
    forecasts: Iterable[Forecast],
    *,
    legs: int = 3,
    limit: int = 5,
    allow_provisional: bool = False,
) -> List[dict]:
    """Return exact top independent-probability combinations for the leg count.

    The optional paper path admits provisional forecasts only when insufficient
    sample size is their sole blocker and at least one graded observation exists.
    It never relaxes price, team-history, push-risk, or probability-floor gates.
    The default remains strict. The best leg from each game dominates every
    other leg from that game, so best-first search avoids combinatorial
    materialization under the cross-game conditional-independence assumption.
    """
    if legs < 2 or legs > 6:
        raise ValueError("acca legs must be between two and six")
    if limit < 1 or limit > 100:
        raise ValueError("acca limit must be between one and one hundred")
    candidates = _best_leg_per_game(forecasts, allow_provisional=allow_provisional)
    if len(candidates) < legs:
        return []
    probabilities = [float(row.calibrated_probability) for row in candidates]
    costs = [-math.log(probability) for probability in probabilities]
    initial = tuple(range(legs))
    heap: List[tuple[float, Tuple[int, ...]]] = [(sum(costs[index] for index in initial), initial)]
    visited = {initial}
    output: List[dict] = []
    while heap and len(output) < limit:
        cost, state = heapq.heappop(heap)
        selected = [candidates[index] for index in state]
        joint = math.exp(-cost)
        decimal_odds = math.prod(float(row.decimal_price) for row in selected)
        leg_ids = [row.forecast_id for row in selected]
        provisional_leg_count = sum(is_provisional_sample_only(row) for row in selected)
        provisional_gate_bypassed = allow_provisional and provisional_leg_count > 0
        output.append(
            {
                "acca_id": hashlib.sha256("|".join(leg_ids).encode()).hexdigest()[:20],
                "rank": len(output) + 1,
                "provisional_sample_gate_bypassed": provisional_gate_bypassed,
                "provisional_leg_count": provisional_leg_count,
                "legs": [row.to_dict() for row in selected],
                "leg_count": legs,
                "estimated_joint_win_probability": round(joint, 8),
                "combined_decimal_odds": round(decimal_odds, 4),
                "bookmaker_implied_probability": round(1 / decimal_odds, 8),
                "estimated_ev_per_unit": round(joint * decimal_odds - 1, 8),
                "ranking_metric": "calibrated_joint_win_probability",
                "ranking_status": (
                    "paper_only_provisional_sample_gate_bypassed"
                    if provisional_gate_bypassed
                    else "paper_only_strict_gates_passed"
                    if allow_provisional
                    else "highest_estimated_not_guaranteed"
                ),
                "dependence": {
                    "same_game_legs": "prohibited",
                    "cross_game_model": "conditional_independence",
                    "limitation": (
                        "Shared league, weather, scheduling, or other latent "
                        "dependence is not estimated; no same-game parlay is built."
                    ),
                },
            }
        )
        for neighbor in _neighbors(state, len(candidates)):
            if neighbor in visited:
                continue
            visited.add(neighbor)
            heapq.heappush(
                heap,
                (sum(costs[index] for index in neighbor), neighbor),
            )
    return output


__all__ = ["build_accas", "is_provisional_sample_only"]
