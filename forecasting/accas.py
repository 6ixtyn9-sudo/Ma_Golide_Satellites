"""Highest-estimated-probability acca construction from calibrated legs."""

from __future__ import annotations

import hashlib
import heapq
import math
from datetime import timedelta
from typing import Dict, Iterable, List, Tuple

from ingestion.identity import parse_utc

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


def _candidate_rank(row: Forecast) -> tuple:
    return (
        -(row.calibrated_probability or 0.0),
        row.market,
        row.period,
        row.source,
        row.selection,
        row.book.key if row.book else "",
        row.forecast_id,
    )


def _eligible_candidate_picks(
    forecasts: Iterable[Forecast], *, allow_provisional: bool = False
) -> List[Forecast]:
    """Return one usable model row per exact wager selection."""
    by_wager: Dict[Tuple[str, str, str, str, str, str], Forecast] = {}
    for forecast in forecasts:
        if not (forecast.eligibility or {}).get("acca_eligible"):
            if not allow_provisional or not is_provisional_sample_only(forecast):
                continue
        probability = forecast.calibrated_probability
        price = forecast.decimal_price
        if (
            probability is None
            or isinstance(probability, bool)
            or not math.isfinite(probability)
            or not 0.0 < probability <= 1.0
            or price is None
            or not math.isfinite(price)
            or price <= 1.0
        ):
            continue
        wager_key = _wager_leg_key(
            forecast.league,
            forecast.game_id,
            forecast.market,
            forecast.period,
            forecast.selection,
            forecast.line,
        )
        previous = by_wager.get(wager_key)
        if previous is None or _candidate_rank(forecast) < _candidate_rank(previous):
            by_wager[wager_key] = forecast
    return sorted(by_wager.values(), key=_candidate_rank)


def _best_leg_per_game(
    forecasts: Iterable[Forecast], *, allow_provisional: bool = False
) -> List[Forecast]:
    grouped: Dict[str, List[Forecast]] = {}
    for forecast in _eligible_candidate_picks(
        forecasts, allow_provisional=allow_provisional
    ):
        grouped.setdefault(forecast.game_id, []).append(forecast)
    best = [rows[0] for rows in grouped.values()]
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


def _probability_band(probability: float) -> str:
    if probability < 0.50:
        return "<0.50"
    if probability < 0.60:
        return "0.50-<0.60"
    if probability < 0.70:
        return "0.60-<0.70"
    if probability < 0.80:
        return "0.70-<0.80"
    if probability < 0.90:
        return "0.80-<0.90"
    return "0.90-1.00"


def _odds_band(price: float) -> str:
    if price < 1.20:
        return "1.00-<1.20"
    if price < 1.50:
        return "1.20-<1.50"
    if price < 2.00:
        return "1.50-<2.00"
    return "2.00+"


def _roi_stats() -> dict:
    return {
        "count": 0,
        "wins": 0,
        "losses": 0,
        "pushes": 0,
        "raw_profit": 0.0,
        "raw_stake": 0.0,
        "weighted_profit": 0.0,
        "weighted_stake": 0.0,
    }


def _record_roi(stats: dict, outcome: str, price: float, recency_weight: float) -> None:
    profit = price - 1.0 if outcome == "win" else -1.0 if outcome == "loss" else 0.0
    stats["count"] += 1
    stats[{"win": "wins", "loss": "losses", "push": "pushes"}[outcome]] += 1
    stats["raw_profit"] += profit
    stats["raw_stake"] += 1.0
    stats["weighted_profit"] += profit * recency_weight
    stats["weighted_stake"] += recency_weight


def _roi(stats: dict) -> float | None:
    stake = stats["weighted_stake"]
    return stats["weighted_profit"] / stake if stake > 0 else None


def _observed_roi(stats: dict) -> float | None:
    stake = stats["raw_stake"]
    return stats["raw_profit"] / stake if stake > 0 else None


def _shrink_roi(stats: dict, prior_roi: float | None, prior_strength: float) -> float | None:
    if prior_roi is None:
        return _roi(stats)
    stake = stats["weighted_stake"]
    return (stats["weighted_profit"] + prior_strength * prior_roi) / (
        stake + prior_strength
    )


def _same_historical_scope(forecast: Forecast) -> bool:
    """Reject malformed or post-tip book evidence without requiring legacy scope fields."""
    book = forecast.book
    if book is None or not book.pre_match:
        return False
    try:
        scheduled = parse_utc(forecast.scheduled_utc)
        observed = parse_utc(book.observed_utc)
    except (TypeError, ValueError):
        return False
    if observed >= scheduled:
        return False
    if forecast.market in {"spread", "total"} and forecast.line is None:
        return False
    for observed_scope, expected_scope in (
        (book.game_id, forecast.game_id),
        (book.market, forecast.market),
        (book.period, forecast.period),
    ):
        if observed_scope not in (None, "") and observed_scope != expected_scope:
            return False
    if book.line is not None:
        if forecast.line is None:
            return False
        try:
            if not math.isclose(
                float(book.line), float(forecast.line), rel_tol=0.0, abs_tol=1e-9
            ):
                return False
        except (OverflowError, TypeError, ValueError):
            return False
    return True


def filter_paper_candidate_picks(
    forecasts: Iterable[Forecast],
    historical_forecasts: Iterable[Forecast] = (),
    *,
    as_of_utc: str,
    history_source: str = "prequential_settled_individual_forecasts",
    half_life_days: float = 14.0,
    prior_strength: float = 3.0,
    result_available_after_hours: float = 4.0,
) -> tuple[List[Forecast], dict]:
    """Route paper picks through adaptive, market-specific positive-ROI bands.

    Band performance uses prior settled *individual picks* at their preserved
    pre-tip prices; acca results are never inputs. Exact duplicate wagers count
    once. ROI is flat-stake profit (win: price minus one, loss: minus one,
    push: zero), weighted by recency, then shrunk toward model, market, and
    league baselines. A candidate band must itself have at least one settled
    pick, positive recency-weighted ROI, and positive adjusted ROI. There is no
    long sample-size qualification period and an unseen band cannot inherit
    eligibility merely from a positive parent baseline.
    """
    for name, value, allow_zero in (
        ("ROI-band half-life", half_life_days, False),
        ("ROI-band prior strength", prior_strength, False),
        ("result-availability delay", result_available_after_hours, True),
    ):
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
            or value < 0
            or (not allow_zero and value == 0)
        ):
            qualifier = "non-negative" if allow_zero else "positive"
            raise ValueError(f"{name} must be finite and {qualifier}")
    as_of = parse_utc(as_of_utc)
    history_cutoff = as_of - timedelta(hours=float(result_available_after_hours))
    candidates = _eligible_candidate_picks(forecasts, allow_provisional=True)

    historical_by_wager: Dict[Tuple[str, str, str, str, str, str], Forecast] = {}
    for forecast in historical_forecasts:
        outcome = forecast.outcome
        probability = forecast.calibrated_probability
        if (
            outcome not in {"win", "loss", "push"}
            or probability is None
            or isinstance(probability, bool)
            or not isinstance(probability, (int, float))
            or not math.isfinite(probability)
            or not 0.0 < probability <= 1.0
            or not _same_historical_scope(forecast)
        ):
            continue
        try:
            price = forecast.decimal_price
        except (OverflowError, TypeError, ValueError):
            continue
        if (
            price is None
            or isinstance(price, bool)
            or not isinstance(price, (int, float))
            or not math.isfinite(price)
            or price <= 1.0
        ):
            continue
        try:
            if parse_utc(forecast.scheduled_utc) > history_cutoff:
                continue
        except (TypeError, ValueError):
            continue
        wager_key = _wager_leg_key(
            forecast.league,
            forecast.game_id,
            forecast.market,
            forecast.period,
            forecast.selection,
            forecast.line,
        )
        previous = historical_by_wager.get(wager_key)
        if previous is None or _candidate_rank(forecast) < _candidate_rank(previous):
            historical_by_wager[wager_key] = forecast

    league_stats: Dict[str, dict] = {}
    market_stats: Dict[Tuple[str, str], dict] = {}
    model_stats: Dict[Tuple[str, str, str, str], dict] = {}
    band_stats: Dict[Tuple[str, str, str, str, str, str], dict] = {}
    historical_count = 0
    for forecast in historical_by_wager.values():
        probability = float(forecast.calibrated_probability)
        price = float(forecast.decimal_price)
        age_days = max(
            0.0,
            (as_of - parse_utc(forecast.scheduled_utc)).total_seconds() / 86400.0,
        )
        recency_weight = math.exp(-math.log(2.0) * age_days / float(half_life_days))
        if recency_weight <= 0.0:
            continue
        outcome = str(forecast.outcome)
        market_key = (forecast.league, forecast.market)
        model_key = (*market_key, forecast.period, forecast.source)
        band_key = (*model_key, _probability_band(probability), _odds_band(price))
        _record_roi(
            league_stats.setdefault(forecast.league, _roi_stats()),
            outcome,
            price,
            recency_weight,
        )
        _record_roi(market_stats.setdefault(market_key, _roi_stats()), outcome, price, recency_weight)
        _record_roi(model_stats.setdefault(model_key, _roi_stats()), outcome, price, recency_weight)
        _record_roi(band_stats.setdefault(band_key, _roi_stats()), outcome, price, recency_weight)
        historical_count += 1

    league_roi = {league: _roi(stats) for league, stats in league_stats.items()}
    market_roi: Dict[Tuple[str, str], float | None] = {}
    for key, stats in market_stats.items():
        market_roi[key] = _shrink_roi(stats, league_roi.get(key[0]), float(prior_strength))
    model_roi: Dict[Tuple[str, str, str, str], float | None] = {}
    for key, stats in model_stats.items():
        model_roi[key] = _shrink_roi(
            stats,
            market_roi.get((key[0], key[1]), league_roi.get(key[0])),
            float(prior_strength),
        )

    candidate_groups: Dict[Tuple[str, str, str, str, str, str], List[Forecast]] = {}
    for forecast in candidates:
        band_key = (
            forecast.league,
            forecast.market,
            forecast.period,
            forecast.source,
            _probability_band(float(forecast.calibrated_probability)),
            _odds_band(float(forecast.decimal_price)),
        )
        candidate_groups.setdefault(band_key, []).append(forecast)

    qualifying: List[Forecast] = []
    band_reports = []
    candidate_band_details = {}
    for band_key, rows in sorted(candidate_groups.items()):
        league, market, period, source, probability_band, price_band = band_key
        parent_roi = model_roi.get(
            (league, market, period, source),
            market_roi.get((league, market), league_roi.get(league)),
        )
        stats = band_stats.get(band_key, _roi_stats())
        weighted_roi = _roi(stats)
        raw_roi = _observed_roi(stats)
        adjusted_roi = (
            _shrink_roi(stats, parent_roi, float(prior_strength))
            if stats["count"]
            else None
        )
        qualifies = (
            stats["count"] > 0
            and weighted_roi is not None
            and weighted_roi > 0.0
            and adjusted_roi is not None
            and adjusted_roi > 0.0
        )
        if qualifies:
            qualifying.extend(rows)
        details = {
            "league": league,
            "market": market,
            "period": period,
            "model": source,
            "probability_band": probability_band,
            "individual_odds_band": price_band,
            "candidate_pick_count": len(rows),
            "historical_pick_count": stats["count"],
            "wins": stats["wins"],
            "losses": stats["losses"],
            "pushes": stats["pushes"],
            "raw_stake": round(stats["raw_stake"], 6),
            "raw_profit_units": round(stats["raw_profit"], 6),
            "weighted_history_stake": round(stats["weighted_stake"], 6),
            "recency_weighted_profit_units": round(stats["weighted_profit"], 6),
            "observed_roi": round(raw_roi, 6) if raw_roi is not None else None,
            "recency_weighted_roi": (
                round(weighted_roi, 6) if weighted_roi is not None else None
            ),
            "parent_roi": round(parent_roi, 6) if parent_roi is not None else None,
            "adjusted_roi": round(adjusted_roi, 6) if adjusted_roi is not None else None,
            "qualifies": qualifies,
        }
        band_reports.append(details)
        for row in rows:
            candidate_band_details[row.forecast_id] = dict(details)

    qualifying_band_count = sum(row["qualifies"] for row in band_reports)
    summary = {
        "policy": "adaptive_recency_weighted_positive_individual_pick_roi_bands",
        "decision_rule": (
            "at_least_one_settled_pick_and_positive_recency_weighted_band_roi "
            "and_positive_shrinkage_adjusted_roi"
        ),
        "history_source": history_source,
        "as_of_utc": as_of_utc,
        "prior_hierarchy": "model_source_period_then_market_then_league",
        "performance_basis": "settled_individual_picks_only_not_acca_outcomes",
        "stake_policy": "flat_one_unit_per_priced_pick_with_pushes_in_roi_denominator",
        "probability_bands": [
            "<0.50",
            "0.50-<0.60",
            "0.60-<0.70",
            "0.70-<0.80",
            "0.80-<0.90",
            "0.90-1.00",
        ],
        "individual_odds_bands": ["1.00-<1.20", "1.20-<1.50", "1.50-<2.00", "2.00+"],
        "half_life_days": float(half_life_days),
        "prior_strength_effective_picks": float(prior_strength),
        "result_available_after_hours": float(result_available_after_hours),
        "minimum_historical_pick_count_for_band": 1,
        "historical_unique_pick_count": historical_count,
        "league_recency_weighted_roi": {
            league: round(roi, 6) if roi is not None else None
            for league, roi in sorted(league_roi.items())
        },
        "candidate_pick_count": len(candidates),
        "candidate_band_count": len(candidate_groups),
        "qualifying_band_count": qualifying_band_count,
        "qualifying_pick_count": len(qualifying),
        "bands": band_reports,
        "candidate_band_details": candidate_band_details,
        "status": (
            "no_eligible_candidate_picks"
            if not candidates
            else "no_settled_historical_picks"
            if not historical_count
            else "applied"
            if qualifying
            else "no_positive_historical_roi_band"
        ),
    }
    return sorted(qualifying, key=_candidate_rank), summary

def _neighbors(state: Tuple[int, ...], population: int) -> Iterable[Tuple[int, ...]]:
    for position in range(len(state)):
        candidate = list(state)
        candidate[position] += 1
        if candidate[position] >= population:
            continue
        if position + 1 < len(state) and candidate[position] >= candidate[position + 1]:
            continue
        yield tuple(candidate)


def _wager_leg_key(
    league: str,
    game_id: str,
    market: str,
    period: str,
    selection: str,
    line: float | None,
) -> Tuple[str, str, str, str, str, str]:
    return (
        league,
        game_id,
        market,
        period,
        selection,
        "" if line is None else float(line).hex(),
    )


def build_accas(
    forecasts: Iterable[Forecast],
    *,
    legs: int = 3,
    limit: int = 5,
    allow_provisional: bool = False,
    require_disjoint_legs: bool = False,
) -> List[dict]:
    """Return highest-probability accas for the requested leg count.

    The optional paper path admits provisional forecasts only when insufficient
    sample size is their sole blocker and at least one graded observation exists.
    It never relaxes price, team-history, push-risk, or probability-floor gates.
    When ``require_disjoint_legs`` is true, greedily select top-ranked accas
    without reusing an exact wager selection (league, game, market, period,
    selection, and line). A different selection from the same game can be used
    in a later ticket. The default retains all top combinations for analytical
    ranking. Same-game legs remain prohibited within each acca.
    """
    if legs < 2 or legs > 6:
        raise ValueError("acca legs must be between two and six")
    if limit < 1 or limit > 100:
        raise ValueError("acca limit must be between one and one hundred")
    if require_disjoint_legs:
        all_forecasts = list(forecasts)
        used_leg_keys: set[Tuple[str, str, str, str, str, str]] = set()
        output: List[dict] = []
        while len(output) < limit:
            remaining = [
                forecast
                for forecast in all_forecasts
                if _wager_leg_key(
                    forecast.league,
                    forecast.game_id,
                    forecast.market,
                    forecast.period,
                    forecast.selection,
                    forecast.line,
                )
                not in used_leg_keys
            ]
            best = build_accas(
                remaining,
                legs=legs,
                limit=1,
                allow_provisional=allow_provisional,
            )
            if not best:
                break
            ticket = best[0]
            ticket["rank"] = len(output) + 1
            ticket["ticket_leg_reuse_policy"] = "no_reused_wager_selections"
            output.append(ticket)
            for leg in ticket["legs"]:
                used_leg_keys.add(
                    _wager_leg_key(
                        leg["league"],
                        leg["game_id"],
                        leg["market"],
                        leg["period"],
                        leg["selection"],
                        leg.get("line"),
                    )
                )
        return output

    candidates = _best_leg_per_game(forecasts, allow_provisional=allow_provisional)
    if len(candidates) < legs:
        return []
    probabilities = [float(row.calibrated_probability) for row in candidates]
    costs = [-math.log(probability) for probability in probabilities]
    initial = tuple(range(legs))
    heap: List[tuple[float, Tuple[int, ...]]] = [(sum(costs[index] for index in initial), initial)]
    visited = {initial}
    output = []
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


__all__ = ["build_accas", "filter_paper_candidate_picks", "is_provisional_sample_only"]
