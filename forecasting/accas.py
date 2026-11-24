"""Highest-estimated-probability acca construction from calibrated legs."""

from __future__ import annotations

import hashlib
import heapq
import math
from datetime import timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from ingestion.identity import parse_utc

from .bet_config import DEFAULTS as BET_CONFIG_DEFAULTS
from .types import Forecast

#: The cheapest price a paper ticket leg may carry, from the data-driven config
#: (``minTicketLegOdds``). A leg below it cannot pay for its own risk: on
#: 2026-11-24 the NCAA_M lane printed five "3-leg accas" whose best price was
#: @1.003 -- a full unit risked for 0.3% -- because ESPN reports -100000 as the
#: moneyline for a lopsided game and the builder ranks legs by probability alone.
#: The lane still had 376 priced, eligible picks at 1.20 or better, so the floor
#: removes the absurd tickets and nothing else.
MIN_TICKET_LEG_DECIMAL = float(BET_CONFIG_DEFAULTS["minTicketLegOdds"])

#: The most expensive price a paper ticket leg may carry (``maxTicketLegOdds``).
#: A ceiling keeps the price-aware ranking inside the range the local model can
#: actually speak about: on 2026-11-24 it rated Rice at p=72% against a 3.50
#: quote (+152% claimed edge), and a ranking free to reach that leg assembled a
#: 10.4 acca claiming +194% expected value -- a model overclaim dressed up as an
#: opportunity. With the ceiling the same slate's best ticket was @2.09.
MAX_TICKET_LEG_DECIMAL = float(BET_CONFIG_DEFAULTS["maxTicketLegOdds"])

#: How much of a ticket's ranking is price rather than probability, from the
#: data-driven config (``ticketRankingPriceWeight``). Ranking by probability
#: alone is not neutral: it always walks to the shortest prices on the board,
#: which is how five NCAA_M "3-leg accas" arrived at @1.003 on 2026-11-24 even
#: though the floor now blocks the individual 1.001 legs. A blend keeps
#: probability primary and lets a 1.90 leg outrank a 1.22 leg when the
#: probabilities are close, which is what a bill of bets actually wants.
TICKET_RANKING_PRICE_WEIGHT = float(BET_CONFIG_DEFAULTS["ticketRankingPriceWeight"])


def _ticket_ranking_score(
    row: Forecast, *, price_weight: Optional[float] = None
) -> float:
    """A ticket leg's ranking score: ``log p`` blended with ``log price``.

    ``price_weight=0.0`` reproduces the old pure-probability ranking;
    ``price_weight=0.5`` is the log of ``p * price`` (expected value per unit);
    ``price_weight=1.0`` chases price alone. The default sits at 0.35 because
    the model's edge over the market is small, so a full EV ranking would mostly
    rank noise, while ignoring price entirely produced tickets that could not
    pay for their own risk.

    The score is additive across legs, which is what keeps the k-best
    enumeration in :func:`build_accas` valid.
    """

    weight = TICKET_RANKING_PRICE_WEIGHT if price_weight is None else price_weight
    probability = float(row.calibrated_probability or 0.0)
    price = float(row.decimal_price or 0.0)
    if not (0.0 < probability <= 1.0) or price <= 1.0:
        return float("-inf")
    return (1.0 - weight) * math.log(probability) + weight * math.log(price)


def _ticket_rank(row: Forecast) -> tuple:
    """Sort key: blended ticket score first, then the stable identity ties."""

    return (-_ticket_ranking_score(row), row.game_id, *_candidate_rank(row)[1:])


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
    """Return one usable model row per exact wager selection.

    A row priced outside ``[MIN_TICKET_LEG_DECIMAL, MAX_TICKET_LEG_DECIMAL]`` is
    not usable for a paper ticket and is dropped here, before bands, singles or
    acca legs are built, so every downstream count agrees on what is on the
    table.
    """
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
            or not MIN_TICKET_LEG_DECIMAL <= price <= MAX_TICKET_LEG_DECIMAL
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
    # One best row per fixture: calibrated probability first, then stable
    # identities. Price enters the *ticket* ranking later
    # (``_ticket_ranking_score``), not the choice of which row represents a
    # fixture -- otherwise a long-priced row could stand in for a short-priced
    # one of the same game.
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


def ticket_leg_pool(
    forecasts: Iterable[Forecast], *, allow_provisional: bool = False
) -> List[Forecast]:
    """Legs a multi-leg ticket could actually be built from, one per fixture.

    Public wrapper over the builder's own pool so a caller can report *why* a
    ticket was not built instead of guessing. On 2026-11-24 the NCAA_W card said
    "9 of 14 candidate picks passed individual-pick ROI; no multi-leg acca was
    produced" -- true, and useless: nine picks qualified for a single, none was
    ``acca_eligible``, and only this pool can answer that.
    """

    return _best_leg_per_game(forecasts, allow_provisional=allow_provisional)


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
    minimum_historical_picks_for_band: int = 1,
    require_positive_expected_value: bool = False,
    require_point_in_time_price: bool = False,
    admit_cold_start_bands: bool = True,
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

    Cold start (operator decision 2026-11-22)
    -----------------------------------------
    A band is admitted *once* while the lane has settled nothing at all: if
    ``historical_count`` is zero there is no ROI to require, and refusing every
    band would deadlock a new lane permanently -- it can never settle the picks
    that would qualify it. That admission is narrow and self-terminating: it
    applies only when the whole scope has zero settled individual picks, only to
    bands with zero history, and it stops the moment one pick settles. Each
    admitted band is stamped ``qualification_basis=cold_start_no_settled_history``
    in ``bands`` and counted in ``cold_start`` so a reader can see it was an
    admission, not a measured ROI.
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
    if (
        isinstance(minimum_historical_picks_for_band, bool)
        or not isinstance(minimum_historical_picks_for_band, int)
        or minimum_historical_picks_for_band < 1
    ):
        raise ValueError("minimum historical picks for band must be a positive integer")
    if not isinstance(require_positive_expected_value, bool):
        raise ValueError("require_positive_expected_value must be a boolean")
    if not isinstance(require_point_in_time_price, bool):
        raise ValueError("require_point_in_time_price must be a boolean")
    if not isinstance(admit_cold_start_bands, bool):
        raise ValueError("admit_cold_start_bands must be a boolean")
    as_of = parse_utc(as_of_utc)
    history_cutoff = as_of - timedelta(hours=float(result_available_after_hours))
    candidate_pool = _eligible_candidate_picks(forecasts, allow_provisional=True)
    if require_point_in_time_price:
        scoped_candidates = []
        for row in candidate_pool:
            if not _same_historical_scope(row) or row.book is None:
                continue
            try:
                if parse_utc(row.book.observed_utc) > as_of:
                    continue
            except (TypeError, ValueError):
                continue
            scoped_candidates.append(row)
    else:
        scoped_candidates = candidate_pool
    candidates = (
        [
            row
            for row in scoped_candidates
            if float(row.calibrated_probability) * float(row.decimal_price) > 1.0
        ]
        if require_positive_expected_value
        else scoped_candidates
    )

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
        cold_start = (
            admit_cold_start_bands
            and historical_count == 0
            and stats["count"] == 0
        )
        qualifies = cold_start or (
            stats["count"] >= minimum_historical_picks_for_band
            and weighted_roi is not None
            and weighted_roi > 0.0
            and adjusted_roi is not None
            and adjusted_roi > 0.0
        )
        qualifying_rows = [
            row
            for row in rows
            if qualifies
            and (
                not require_positive_expected_value
                or float(row.calibrated_probability) * float(row.decimal_price) > 1.0
            )
        ]
        qualifying.extend(qualifying_rows)
        details = {
            "league": league,
            "market": market,
            "period": period,
            "model": source,
            "probability_band": probability_band,
            "individual_odds_band": price_band,
            "candidate_pick_count": len(rows),
            "qualified_candidate_pick_count": sum(
                not require_positive_expected_value
                or float(row.calibrated_probability) * float(row.decimal_price) > 1.0
                for row in rows
            ) if qualifies else 0,
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
            "qualification_basis": (
                "cold_start_no_settled_history"
                if cold_start
                else "measured_positive_recency_weighted_and_adjusted_roi"
                if qualifies
                else "rejected"
            ),
        }
        band_reports.append(details)
        for row in rows:
            row_details = dict(details)
            row_details["expected_value_per_unit"] = round(
                float(row.calibrated_probability) * float(row.decimal_price) - 1.0, 6
            )
            row_details["qualifies"] = row in qualifying_rows
            candidate_band_details[row.forecast_id] = row_details

    qualifying_band_count = sum(
        row["qualified_candidate_pick_count"] > 0 for row in band_reports
    )
    summary = {
        "policy": (
            "priced_positive_expected_value_positive_roi_bands"
            if require_positive_expected_value or require_point_in_time_price
            else "adaptive_recency_weighted_positive_individual_pick_roi_bands"
        ),
        "decision_rule": (
            "minimum_settled_picks_and_positive_recency_weighted_band_roi_and_"
            "positive_shrinkage_adjusted_roi_and_positive_calibrated_expected_value"
            if require_positive_expected_value
            else "minimum_settled_picks_and_positive_recency_weighted_band_roi_"
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
        "minimum_historical_pick_count_for_band": minimum_historical_picks_for_band,
        "require_positive_expected_value": require_positive_expected_value,
        "require_point_in_time_price": require_point_in_time_price,
        "cold_start_admission": {
            "enabled": admit_cold_start_bands,
            "applies_when_lane_settled_pick_count_is_zero": True,
            "admitted_band_count": sum(
                row.get("qualification_basis") == "cold_start_no_settled_history"
                for row in band_reports
            ),
            "applied_this_run": admit_cold_start_bands and historical_count == 0,
        },
        "candidate_pool_pick_count": len(candidate_pool),
        "candidate_with_point_in_time_price_count": len(scoped_candidates),
        "candidate_positive_expected_value_count": len(candidates),
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
            if not candidate_pool
            else "no_point_in_time_priced_candidates"
            if not scoped_candidates
            else "no_positive_expected_value_candidates"
            if not candidates
            else "applied_cold_start_no_settled_history"
            if not historical_count and qualifying
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
    """Return best-ranked accas for the requested leg count.

    Ranking is the blended probability/price score from
    :func:`_ticket_ranking_score`, not raw joint probability, so a ticket is
    assembled from legs that can pay rather than from whichever legs are most
    certain. Same-game legs stay prohibited and every eligibility gate is
    unchanged; ``estimated_joint_win_probability`` and
    ``estimated_ev_per_unit`` are still reported so the blend never hides the
    probability that was traded away.

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
    # Enumerate in ranked order: the k-best walk below assumes candidate index
    # order is ascending cost, so the blend has to be applied before it starts.
    candidates.sort(key=_ticket_rank)
    costs = [-_ticket_ranking_score(row) for row in candidates]
    initial = tuple(range(legs))
    heap: List[tuple[float, Tuple[int, ...]]] = [(sum(costs[index] for index in initial), initial)]
    visited = {initial}
    output = []
    while heap and len(output) < limit:
        cost, state = heapq.heappop(heap)
        selected = [candidates[index] for index in state]
        # The cost is now a blended score, so the joint probability is read from
        # the probabilities rather than assumed to be ``exp(-cost)``.
        joint = math.exp(
            sum(math.log(float(row.calibrated_probability)) for row in selected)
        )
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
                "ranking_metric": "probability_price_blend",
                "ranking_price_weight": TICKET_RANKING_PRICE_WEIGHT,
                "ranking_score": round(-cost, 8),
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


__all__ = [
    "build_accas",
    "filter_paper_candidate_picks",
    "is_provisional_sample_only",
    "ticket_leg_pool",
]
