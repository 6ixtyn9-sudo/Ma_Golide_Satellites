"""Descriptive statistics for warehouse-native prediction research.

This module consumes normalized prediction outcomes, not provider payloads or
bookmaker bets. It deliberately contains no pick certification, stake sizing,
expected-total conversion, or profitability claims.
"""

from __future__ import annotations

from collections import defaultdict
import math
from typing import Any, Dict, Iterable, List

# 1.645 is the one-sided 95% normal quantile used for the Wilson lower bound.
WILSON_Z = 1.645
MIN_SEGMENT_N = 10
MIN_SAMPLE_N = 30


def wilson_lower_bound(wins: int, n: int, z: float = WILSON_Z) -> float:
    """Return the Wilson score lower bound for a binomial success rate."""
    if not isinstance(wins, int) or not isinstance(n, int):
        raise TypeError("wins and n must be integers")
    if n < 0 or wins < 0 or wins > n:
        raise ValueError("require 0 <= wins <= n")
    if z <= 0:
        raise ValueError("z must be positive")
    if n == 0:
        return 0.0
    p = wins / n
    denominator = 1 + z * z / n
    center = p + z * z / (2 * n)
    spread = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return max(0.0, (center - spread) / denominator)


def shrink_rate(wins: int, n: int, prior_alpha: float = 2, prior_beta: float = 2) -> float:
    """Return a Beta-binomial posterior mean with an explicit symmetric prior."""
    if n < 0 or wins < 0 or wins > n:
        raise ValueError("require 0 <= wins <= n")
    if prior_alpha <= 0 or prior_beta <= 0:
        raise ValueError("Beta prior parameters must be positive")
    return (wins + prior_alpha) / (n + prior_alpha + prior_beta)


def _confidence_bucket(confidence: Any) -> str:
    try:
        value = float(confidence)
    except (TypeError, ValueError) as exc:
        raise ValueError("prediction confidence must be numeric") from exc
    if not math.isfinite(value) or not 0 <= value <= 1:
        raise ValueError("prediction confidence must be between zero and one")
    if value >= 0.80:
        return "80+"
    if value >= 0.70:
        return "70-79"
    if value >= 0.60:
        return "60-69"
    if value >= 0.50:
        return "50-59"
    return "<50"


def assay_prediction_outcomes(
    rows: Iterable[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Aggregate frozen moneyline outcomes by provider and confidence segment.

    Rows must already contain official, reconciled outcomes. Invalid rows are
    rejected rather than silently changing the research sample. Aggregates
    below ``MIN_SEGMENT_N`` remain in the top-level sample but are not emitted
    as separate segments.
    """
    segments: Dict[tuple, Dict[str, int]] = defaultdict(lambda: {"wins": 0, "losses": 0})
    for row in rows:
        outcome = str(row.get("outcome", "")).strip().lower()
        if outcome not in {"win", "loss"}:
            raise ValueError(f"unsupported prediction outcome: {outcome!r}")
        key = (
            str(row.get("league") or "unknown"),
            str(row.get("period") or "full_game"),
            str(row.get("market") or "moneyline"),
            str(row.get("side") or "unknown"),
            _confidence_bucket(row.get("confidence")),
            str(row.get("provider") or "unknown"),
        )
        segments[key]["wins" if outcome == "win" else "losses"] += 1

    edges: List[Dict[str, Any]] = []
    for segment, counts in segments.items():
        wins = counts["wins"]
        losses = counts["losses"]
        n = wins + losses
        if n < MIN_SEGMENT_N:
            continue
        win_rate = wins / n
        edges.append(
            {
                "key": "|".join(segment),
                "segment": segment,
                "n": n,
                "wins": wins,
                "losses": losses,
                "win_rate": round(win_rate, 4),
                "shrunk_rate": round(shrink_rate(wins, n), 4),
                "wilson_lower_95": round(wilson_lower_bound(wins, n), 4),
                "observed_lift_over_50pct": round(win_rate - 0.5, 4),
                "minimum_sample_met": n >= MIN_SAMPLE_N,
            }
        )
    edges.sort(
        key=lambda edge: (
            edge["segment"][-1],
            -edge["win_rate"],
            -edge["n"],
            edge["key"],
        )
    )
    return edges
