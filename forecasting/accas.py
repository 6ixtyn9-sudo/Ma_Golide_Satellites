"""Highest-estimated-probability acca construction from calibrated legs."""

from __future__ import annotations

import bisect
import hashlib
import heapq
import math
from datetime import timedelta
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

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

#: Unified paper-ticket construction defaults (Option 1 + Option A):
#: - Variable 2–3 legs, limit 3, combined odds [3.30, 5.00] so any 1W-2L day
#:   nets >= +0.30u (strictly green, avoiding the breakeven/red 1W trap).
#: - Exact Kelly log-growth ranking g*(p, O) with stake-normalization term.
#: - Non-blocking audit flag at EV > 15% and hard sanity cap at EV > 40%.
PAPER_MIN_ACCA_LEGS = 2
PAPER_MAX_ACCA_LEGS = 3
PAPER_ACCA_LIMIT = 3
PAPER_MIN_COMBINED_ODDS = 3.30
PAPER_MAX_COMBINED_ODDS = 5.00
EDGE_FLAG_THRESHOLD = 0.15
MAX_CLAIMED_EDGE = 0.40
RANKING_METRIC_BLEND = "probability_price_blend"
RANKING_METRIC_KELLY = "kelly_log_growth"
VALID_RANKING_METRICS = frozenset({RANKING_METRIC_BLEND, RANKING_METRIC_KELLY})


def kelly_log_growth(probability: float, decimal_odds: float) -> float:
    """Exact full-Kelly expected log-wealth growth rate ``g*(p, O)``.

    For win probability ``p`` and decimal odds ``O`` (net odds ``O - 1``), the
    optimal Kelly stake fraction is ``f* = (p * O - 1) / (O - 1)`` and the
    resulting expected log growth per unit bankroll is:

        g*(p, O) = p * ln(p * O) + (1 - p) * ln(O * (1 - p) / (O - 1))

    At identical positive EV, shorter prices score higher than longshots
    because of lower variance drag (the ``O / (O - 1)`` stake-normalization
    term). Returns ``-inf`` whenever ``p * O <= 1.0`` or inputs are invalid.
    """
    if (
        isinstance(probability, bool)
        or isinstance(decimal_odds, bool)
        or not math.isfinite(probability)
        or not math.isfinite(decimal_odds)
        or decimal_odds <= 1.0
        or probability <= 0.0
        or probability > 1.0
    ):
        return float("-inf")
    if probability * decimal_odds <= 1.0:
        return float("-inf")
    if probability >= 1.0:
        return math.log(decimal_odds)
    return probability * math.log(probability * decimal_odds) + (
        1.0 - probability
    ) * math.log(decimal_odds * (1.0 - probability) / (decimal_odds - 1.0))


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


def _ticket_rank(row: Forecast, *, price_weight: Optional[float] = None) -> tuple:
    """Sort key: blended ticket score first, then the stable identity ties."""

    return (
        -_ticket_ranking_score(row, price_weight=price_weight),
        row.game_id,
        *_candidate_rank(row)[1:],
    )


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


def _history_canonical_rank(
    row: Forecast,
    empirical_segment_brier: Dict[Tuple[str, str, str, str], float] | None = None,
) -> tuple:
    """Canonical ordering when deduplicating historical settled wagers.

    Matches :func:`forecasting.harness._calibration_evidence_rank`: a
    walk-forward sufficient segment with lower ``brier_score`` owns the
    historical wager instead of letting an overconfident weak model erase
    the sharper model's settled ROI history. When older compacted audit rows
    omit ``calibration.brier_score``, falls back to the segment's past-only
    empirical Brier score computed from the same settled history pool.
    """

    calibration = row.calibration or {}
    sufficient = calibration.get("sufficient_for_accas") is True
    brier = calibration.get("brier_score")
    has_brier = (
        isinstance(brier, (int, float))
        and not isinstance(brier, bool)
        and math.isfinite(float(brier))
    )
    if not has_brier and empirical_segment_brier is not None:
        fallback = empirical_segment_brier.get(
            (row.league, row.market, row.period, row.source)
        )
        if fallback is not None and math.isfinite(fallback):
            brier = fallback
            has_brier = True
    use_brier = sufficient and has_brier
    is_main_line = bool(getattr(row.book, "main_line", True)) if row.book is not None else False
    return (
        0 if sufficient else 1,
        0 if use_brier else 1,
        float(brier) if use_brier else float("inf"),
        0 if has_brier else 1,
        float(brier) if has_brier else float("inf"),
        0 if is_main_line else 1,
        _candidate_rank(row),
    )


def _is_named_book_priced(forecast: Forecast) -> bool:
    """Whether a row's quote is a bookmaker's own, not an aggregate average.

    A lane that declares ``requires_named_book_quote`` may not be priced by a
    consensus average, however exact the number is. ``price_provenance`` is the
    only field that says which this is -- ``price_provenance_label`` exists so
    the answer is also legible on a card.
    """

    book = forecast.book
    if book is None:
        return False
    provenance = getattr(book, "price_provenance", None)
    return str(getattr(provenance, "value", provenance)) == "named_book"


def _eligible_candidate_picks(
    forecasts: Iterable[Forecast],
    *,
    allow_provisional: bool = False,
    min_leg_odds: Optional[float] = None,
    max_leg_odds: Optional[float] = None,
    min_leg_probability: Optional[float] = None,
) -> List[Forecast]:
    """Return one usable model row per exact wager selection.

    A row priced outside ``[MIN_TICKET_LEG_DECIMAL, MAX_TICKET_LEG_DECIMAL]`` is
    not usable for a paper ticket and is dropped here, before bands, singles or
    acca legs are built, so every downstream count agrees on what is on the
    table.
    """
    min_odds = MIN_TICKET_LEG_DECIMAL if min_leg_odds is None else float(min_leg_odds)
    max_odds = MAX_TICKET_LEG_DECIMAL if max_leg_odds is None else float(max_leg_odds)
    min_prob = None if min_leg_probability is None else float(min_leg_probability)
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
            or (min_prob is not None and probability < min_prob)
            or price is None
            or not math.isfinite(price)
            or not min_odds <= price <= max_odds
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
    forecasts: Iterable[Forecast],
    *,
    allow_provisional: bool = False,
    min_leg_odds: Optional[float] = None,
    max_leg_odds: Optional[float] = None,
    min_leg_probability: Optional[float] = None,
) -> List[Forecast]:
    grouped: Dict[Tuple[str, str], List[Forecast]] = {}
    for forecast in _eligible_candidate_picks(
        forecasts,
        allow_provisional=allow_provisional,
        min_leg_odds=min_leg_odds,
        max_leg_odds=max_leg_odds,
        min_leg_probability=min_leg_probability,
    ):
        grouped.setdefault((forecast.league, forecast.game_id), []).append(forecast)
    best = [rows[0] for rows in grouped.values()]
    # One best row per fixture: calibrated probability first, then stable
    # identities. Price enters the *ticket* ranking later
    # (``_ticket_ranking_score``), not the choice of which row represents a
    # fixture -- otherwise a long-priced row could stand in for a short-priced
    # one of the same game.
    best.sort(
        key=lambda row: (
            -(row.calibrated_probability or 0.0),
            row.league,
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
    require_named_book_quote: bool = False,
    require_calibrated_segment: bool = False,
    maximum_segment_brier_score: float | None = None,
    max_claimed_edge: float | None = None,
    edge_flag_threshold: float | None = None,
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

    Named-book lanes (strict contract)
    ----------------------------------
    ``require_named_book_quote`` makes a lane's *declared* contract operative.
    ``EUROLEAGUE_STRICT_NAMED_BOOK_VALUE_TICKET_LANE`` has always said
    "aggregate consensus quotes are excluded" and set
    ``requires_named_book_quote: true``, but nothing read it: the global policy
    accepts a consensus average so aggregate-only lanes can price paper tickets,
    and the strict lane inherited that as if it were its own. On 2026-11-25 that
    lane priced 5 of 11 eligible rows and all 4 strict acca candidates on quotes
    its own data labels "AGGREGATE CONSENSUS AVERAGE — NOT A NAMED-BOOK QUOTE".
    A lane that declares the stricter contract now gets it, and the excluded
    count is reported instead of the rows vanishing.

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
    if not isinstance(require_named_book_quote, bool):
        raise ValueError("require_named_book_quote must be a boolean")
    if not isinstance(require_calibrated_segment, bool):
        raise ValueError("require_calibrated_segment must be a boolean")
    if maximum_segment_brier_score is not None and (
        isinstance(maximum_segment_brier_score, bool)
        or not isinstance(maximum_segment_brier_score, (int, float))
        or not math.isfinite(float(maximum_segment_brier_score))
        or not 0.0 < float(maximum_segment_brier_score) <= 1.0
    ):
        raise ValueError("maximum_segment_brier_score must be in (0, 1]")
    if max_claimed_edge is not None and (
        isinstance(max_claimed_edge, bool)
        or not isinstance(max_claimed_edge, (int, float))
        or not math.isfinite(float(max_claimed_edge))
        or float(max_claimed_edge) <= 0.0
    ):
        raise ValueError("max_claimed_edge must be finite and positive")
    if edge_flag_threshold is not None and (
        isinstance(edge_flag_threshold, bool)
        or not isinstance(edge_flag_threshold, (int, float))
        or not math.isfinite(float(edge_flag_threshold))
        or float(edge_flag_threshold) <= 0.0
    ):
        raise ValueError("edge_flag_threshold must be finite and positive")
    if not isinstance(admit_cold_start_bands, bool):
        raise ValueError("admit_cold_start_bands must be a boolean")
    as_of = parse_utc(as_of_utc)
    history_cutoff = as_of - timedelta(hours=float(result_available_after_hours))
    candidate_pool = _eligible_candidate_picks(
        forecasts, allow_provisional=not require_calibrated_segment
    )
    excluded_uncalibrated_segment = 0
    if require_calibrated_segment:
        calibrated_only = [
            row
            for row in candidate_pool
            if (row.calibration or {}).get("sufficient_for_accas") is True
        ]
        excluded_uncalibrated_segment = len(candidate_pool) - len(calibrated_only)
        candidate_pool = calibrated_only
    excluded_segment_brier = 0
    if maximum_segment_brier_score is not None:
        max_brier = float(maximum_segment_brier_score)
        skilled_only = []
        for row in candidate_pool:
            brier = (row.calibration or {}).get("brier_score")
            if (
                isinstance(brier, (int, float))
                and not isinstance(brier, bool)
                and math.isfinite(float(brier))
                and float(brier) >= max_brier
            ):
                excluded_segment_brier += 1
                continue
            skilled_only.append(row)
        candidate_pool = skilled_only
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
    excluded_no_named_book_quote = 0
    if require_named_book_quote:
        named_book_only = [row for row in scoped_candidates if _is_named_book_priced(row)]
        excluded_no_named_book_quote = len(scoped_candidates) - len(named_book_only)
        scoped_candidates = named_book_only
    candidates = (
        [
            row
            for row in scoped_candidates
            if float(row.calibrated_probability) * float(row.decimal_price) > 1.0
        ]
        if require_positive_expected_value
        else scoped_candidates
    )
    positive_ev_count = len(candidates)
    excluded_implausible_edge = 0
    if max_claimed_edge is not None:
        edge_cap = float(max_claimed_edge)
        plausible_only = [
            row
            for row in candidates
            if float(row.calibrated_probability) * float(row.decimal_price) - 1.0 <= edge_cap
        ]
        excluded_implausible_edge = len(candidates) - len(plausible_only)
        candidates = plausible_only

    valid_history: List[Forecast] = []
    segment_brier_sums: Dict[Tuple[str, str, str, str], float] = {}
    segment_brier_counts: Dict[Tuple[str, str, str, str], int] = {}
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
        valid_history.append(forecast)
        if outcome in {"win", "loss"}:
            seg_key = (
                forecast.league,
                forecast.market,
                forecast.period,
                forecast.source,
            )
            y = 1.0 if outcome == "win" else 0.0
            segment_brier_sums[seg_key] = segment_brier_sums.get(seg_key, 0.0) + (
                float(probability) - y
            ) ** 2
            segment_brier_counts[seg_key] = segment_brier_counts.get(seg_key, 0) + 1

    empirical_segment_brier: Dict[Tuple[str, str, str, str], float] = {
        seg_key: segment_brier_sums[seg_key] / count
        for seg_key, count in segment_brier_counts.items()
        if count >= 5
    }

    historical_by_wager: Dict[Tuple[str, str, str, str, str], Forecast] = {}
    for forecast in valid_history:
        # Deduplicate across alternate lines on the same fixture/market/period/selection
        # so multi-line captures on one game cannot multiply correlated counts in band_stats.
        wager_key = _market_selection_leg_key(
            forecast.league,
            forecast.game_id,
            forecast.market,
            forecast.period,
            forecast.selection,
        )
        previous = historical_by_wager.get(wager_key)
        if previous is None or _history_canonical_rank(
            forecast, empirical_segment_brier
        ) < _history_canonical_rank(previous, empirical_segment_brier):
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
            ev = float(row.calibrated_probability) * float(row.decimal_price) - 1.0
            row_details["expected_value_per_unit"] = round(ev, 6)
            if edge_flag_threshold is not None:
                row_details["edge_flag_threshold"] = float(edge_flag_threshold)
                row_details["edge_implausible"] = ev > float(edge_flag_threshold)
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
        "require_named_book_quote": require_named_book_quote,
        "excluded_no_named_book_quote_count": excluded_no_named_book_quote,
        "require_calibrated_segment": require_calibrated_segment,
        "excluded_uncalibrated_segment_count": excluded_uncalibrated_segment,
        "maximum_segment_brier_score": (
            float(maximum_segment_brier_score)
            if maximum_segment_brier_score is not None
            else None
        ),
        "excluded_segment_brier_count": excluded_segment_brier,
        "max_claimed_edge": (
            float(max_claimed_edge) if max_claimed_edge is not None else None
        ),
        "excluded_implausible_edge_count": excluded_implausible_edge,
        "edge_flag_threshold": (
            float(edge_flag_threshold) if edge_flag_threshold is not None else None
        ),
        "flagged_implausible_edge_count": (
            sum(
                bool(detail.get("edge_implausible"))
                for detail in candidate_band_details.values()
            )
            if edge_flag_threshold is not None
            else 0
        ),
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
        "candidate_positive_expected_value_count": positive_ev_count,
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
            "no_calibrated_skill_candidates"
            if not candidate_pool and excluded_segment_brier > 0
            else "no_calibrated_segment_candidates"
            if not candidate_pool and excluded_uncalibrated_segment > 0
            else "no_eligible_candidate_picks"
            if not candidate_pool
            else "no_point_in_time_priced_candidates"
            if not scoped_candidates
            else "no_plausible_edge_candidates"
            if not candidates and excluded_implausible_edge > 0
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


def _market_selection_leg_key(
    league: str,
    game_id: str,
    market: str,
    period: str,
    selection: str,
) -> Tuple[str, str, str, str, str]:
    """Wager identity ignoring ``line`` so alternate lines collapse together."""
    return (league, game_id, market, period, selection)


def _conflicting_game_side_key(
    league: str, game_id: str, market: str, selection: str
) -> Tuple[str, str, str] | None:
    """Return the opposing directional commitment on the same fixture, if any."""

    sel = str(selection or "").lower()
    mkt = str(market or "").lower()
    if mkt in {"moneyline", "spread"}:
        if sel == "home":
            return (league, game_id, "side:away")
        if sel == "away":
            return (league, game_id, "side:home")
    elif mkt == "total":
        if sel == "over":
            return (league, game_id, "total:under")
        if sel == "under":
            return (league, game_id, "total:over")
    return None


def _committed_game_side_key(
    league: str, game_id: str, market: str, selection: str
) -> Tuple[str, str, str] | None:
    sel = str(selection or "").lower()
    mkt = str(market or "").lower()
    if mkt in {"moneyline", "spread"} and sel in {"home", "away"}:
        return (league, game_id, f"side:{sel}")
    if mkt == "total" and sel in {"over", "under"}:
        return (league, game_id, f"total:{sel}")
    return None


class _DescendingIdKey:
    """Wrapper reversing tuple comparison so smaller IDs win ties in a min-heap."""

    __slots__ = ("value",)

    def __init__(self, value: Tuple[str, ...]) -> None:
        self.value = value

    def __lt__(self, other: "_DescendingIdKey") -> bool:
        return self.value > other.value

    def __eq__(self, other: object) -> bool:
        return isinstance(other, _DescendingIdKey) and self.value == other.value


def _build_ticket_dict(
    selected: Sequence[Forecast],
    *,
    rank: int,
    ranking_metric: str,
    ranking_score: float,
    price_weight: Optional[float],
    allow_provisional: bool,
    edge_flag_threshold: Optional[float],
) -> dict:
    joint = math.exp(
        sum(math.log(float(row.calibrated_probability)) for row in selected)
    )
    decimal_odds = math.prod(float(row.decimal_price) for row in selected)
    leg_ids = [row.forecast_id for row in selected]
    provisional_leg_count = sum(is_provisional_sample_only(row) for row in selected)
    provisional_gate_bypassed = allow_provisional and provisional_leg_count > 0
    leg_dicts: List[dict] = []
    flagged_count = 0
    for row in selected:
        leg_dict = row.to_dict()
        if edge_flag_threshold is not None:
            leg_ev = (
                float(row.calibrated_probability) * float(row.decimal_price) - 1.0
            )
            implausible = leg_ev > float(edge_flag_threshold)
            leg_dict["claimed_edge"] = round(leg_ev, 6)
            leg_dict["edge_implausible"] = implausible
            if implausible:
                flagged_count += 1
        leg_dicts.append(leg_dict)
    ticket = {
        "acca_id": hashlib.sha256("|".join(leg_ids).encode()).hexdigest()[:20],
        "rank": rank,
        "provisional_sample_gate_bypassed": provisional_gate_bypassed,
        "provisional_leg_count": provisional_leg_count,
        "legs": leg_dicts,
        "leg_count": len(selected),
        "estimated_joint_win_probability": round(joint, 8),
        "combined_decimal_odds": round(decimal_odds, 4),
        "bookmaker_implied_probability": round(1 / decimal_odds, 8),
        "estimated_ev_per_unit": round(joint * decimal_odds - 1, 8),
        "ranking_metric": ranking_metric,
        "ranking_price_weight": (
            None if ranking_metric == RANKING_METRIC_KELLY else price_weight
        ),
        "ranking_score": round(ranking_score, 8),
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
    if ranking_metric == RANKING_METRIC_KELLY and decimal_odds > 1.0:
        ticket["kelly_fraction"] = round(
            max((joint * decimal_odds - 1.0) / (decimal_odds - 1.0), 0.0), 8
        )
    if edge_flag_threshold is not None:
        ticket["edge_flag_threshold"] = float(edge_flag_threshold)
        ticket["edge_implausible_leg_count"] = flagged_count
        ticket["edge_implausible"] = flagged_count > 0
    return ticket


def _exact_variable_accas(
    forecasts: Iterable[Forecast],
    *,
    min_legs: int,
    max_legs: int,
    limit: int,
    allow_provisional: bool,
    min_leg_odds: Optional[float],
    max_leg_odds: Optional[float],
    min_leg_probability: Optional[float],
    price_weight: Optional[float],
    min_combined_odds: Optional[float],
    max_combined_odds: Optional[float],
    ranking_metric: str,
    max_claimed_edge: Optional[float],
    edge_flag_threshold: Optional[float],
) -> List[dict]:
    """Exact sorted-price branch-and-bound acca search over ``min_legs..max_legs``.

    Keeps one best row per ``(league, game_id, market, period, selection)``
    (avoiding alternate-line combinatorics while preserving distinct markets
    per fixture so heavy moneyline favourites do not starve ``min_combined_odds``)
    and enforces strictly distinct ``(league, game_id)`` inside every ticket.
    """
    effective_price_weight = (
        TICKET_RANKING_PRICE_WEIGHT if price_weight is None else float(price_weight)
    )
    eligible = _eligible_candidate_picks(
        forecasts,
        allow_provisional=allow_provisional,
        min_leg_odds=min_leg_odds,
        max_leg_odds=max_leg_odds,
        min_leg_probability=min_leg_probability,
    )
    by_market_sel: Dict[Tuple[str, str, str, str, str], Tuple[tuple, Forecast]] = {}
    for row in eligible:
        prob = float(row.calibrated_probability)
        price = float(row.decimal_price)
        ev = prob * price - 1.0
        if ranking_metric == RANKING_METRIC_KELLY and ev <= 0.0:
            continue
        if max_claimed_edge is not None and ev > float(max_claimed_edge):
            continue
        single_score = (
            kelly_log_growth(prob, price)
            if ranking_metric == RANKING_METRIC_KELLY
            else _ticket_ranking_score(row, price_weight=effective_price_weight)
        )
        if not math.isfinite(single_score):
            continue
        is_main = bool(getattr(row.book, "main_line", True)) if row.book is not None else False
        order_key = (-single_score, 0 if is_main else 1, _candidate_rank(row))
        mkey = _market_selection_leg_key(
            row.league, row.game_id, row.market, row.period, row.selection
        )
        prev = by_market_sel.get(mkey)
        if prev is None or order_key < prev[0]:
            by_market_sel[mkey] = (order_key, row)

    if not by_market_sel:
        return []

    scored_rows: List[Tuple[float, float, Forecast]] = []
    for order_key, row in by_market_sel.values():
        scored_rows.append((float(row.decimal_price), -order_key[0], row))
    # Sort ascending by price for tight bisect/break bounds on combined odds.
    scored_rows.sort(
        key=lambda item: (
            item[0],
            -item[1],
            item[2].league,
            item[2].game_id,
            item[2].forecast_id,
        )
    )
    n = len(scored_rows)
    if n < min_legs:
        return []
    distinct_games = len({(item[2].league, item[2].game_id) for item in scored_rows})
    effective_max_legs = min(max_legs, distinct_games)
    if effective_max_legs < min_legs:
        return []

    prices = [item[0] for item in scored_rows]
    single_scores = [item[1] for item in scored_rows]
    rows = [item[2] for item in scored_rows]
    probs = [float(r.calibrated_probability) for r in rows]
    log_probs = [math.log(p) for p in probs]
    log_prices = [math.log(o) for o in prices]
    game_keys = [(r.league, r.game_id) for r in rows]
    min_o = None if min_combined_odds is None else float(min_combined_odds)
    max_o = None if max_combined_odds is None else float(max_combined_odds)
    eps = 1e-12

    top_heap: List[Tuple[float, float, _DescendingIdKey, Tuple[int, ...]]] = []

    def _consider(indices: Tuple[int, ...], joint_p: float, combined_o: float) -> None:
        if min_o is not None and combined_o < min_o:
            return
        if max_o is not None and combined_o > max_o:
            return
        if ranking_metric == RANKING_METRIC_KELLY:
            score = kelly_log_growth(joint_p, combined_o)
            if not math.isfinite(score):
                return
        else:
            score = sum(
                (1.0 - effective_price_weight) * log_probs[idx]
                + effective_price_weight * log_prices[idx]
                for idx in indices
            )
        ordered = tuple(
            sorted(
                indices,
                key=lambda idx: (
                    -single_scores[idx],
                    rows[idx].game_id,
                    rows[idx].forecast_id,
                ),
            )
        )
        id_key = _DescendingIdKey(tuple(rows[idx].forecast_id for idx in ordered))
        entry = (score, joint_p, id_key, ordered)
        if len(top_heap) < limit:
            heapq.heappush(top_heap, entry)
        elif entry[:3] > top_heap[0][:3]:
            heapq.heapreplace(top_heap, entry)

    def _last_range(start: int, prod_o: float) -> Tuple[int, int]:
        lo = start
        if min_o is not None:
            target_lo = (min_o / prod_o) * (1.0 - eps)
            lo = max(start, bisect.bisect_left(prices, target_lo, start, n))
        if max_o is not None:
            target_hi = (max_o / prod_o) * (1.0 + eps)
            hi = bisect.bisect_right(prices, target_hi, lo, n)
        else:
            hi = n
        return lo, hi

    for target_legs in range(min_legs, effective_max_legs + 1):
        if target_legs == 2:
            for i in range(n - 1):
                oi = prices[i]
                if max_o is not None and oi * prices[i + 1] > max_o * (1.0 + eps):
                    break
                pi = probs[i]
                gi = game_keys[i]
                lo, hi = _last_range(i + 1, oi)
                for j in range(lo, hi):
                    if game_keys[j] == gi:
                        continue
                    _consider((i, j), pi * probs[j], oi * prices[j])
        elif target_legs == 3:
            for i in range(n - 2):
                oi = prices[i]
                if (
                    max_o is not None
                    and oi * prices[i + 1] * prices[i + 2] > max_o * (1.0 + eps)
                ):
                    break
                pi = probs[i]
                gi = game_keys[i]
                for j in range(i + 1, n - 1):
                    gj = game_keys[j]
                    if gj == gi:
                        continue
                    oij = oi * prices[j]
                    if max_o is not None and oij * prices[j + 1] > max_o * (1.0 + eps):
                        break
                    pij = pi * probs[j]
                    lo, hi = _last_range(j + 1, oij)
                    for k in range(lo, hi):
                        gk = game_keys[k]
                        if gk == gi or gk == gj:
                            continue
                        _consider((i, j, k), pij * probs[k], oij * prices[k])
        else:
            # Recursive exact branch-and-bound for 4..6 legs.
            chosen_indices: List[int] = []
            used_games: set[Tuple[str, str]] = set()

            def _dfs(start: int, remaining: int, prod_p: float, prod_o: float) -> None:
                if remaining == 1:
                    lo, hi = _last_range(start, prod_o)
                    for idx in range(lo, hi):
                        if game_keys[idx] in used_games:
                            continue
                        _consider(
                            (*chosen_indices, idx),
                            prod_p * probs[idx],
                            prod_o * prices[idx],
                        )
                    return
                end = n - remaining + 1
                for idx in range(start, end):
                    if max_o is not None:
                        min_future = prod_o
                        for offset in range(remaining):
                            min_future *= prices[idx + offset]
                        if min_future > max_o * (1.0 + eps):
                            break
                    gkey = game_keys[idx]
                    if gkey in used_games:
                        continue
                    used_games.add(gkey)
                    chosen_indices.append(idx)
                    _dfs(idx + 1, remaining - 1, prod_p * probs[idx], prod_o * prices[idx])
                    chosen_indices.pop()
                    used_games.remove(gkey)

            _dfs(0, target_legs, 1.0, 1.0)

    ranked = sorted(top_heap, key=lambda item: item[:3], reverse=True)
    return [
        _build_ticket_dict(
            [rows[idx] for idx in ordered],
            rank=position,
            ranking_metric=ranking_metric,
            ranking_score=score,
            price_weight=effective_price_weight,
            allow_provisional=allow_provisional,
            edge_flag_threshold=edge_flag_threshold,
        )
        for position, (score, _joint, _id_key, ordered) in enumerate(ranked, 1)
    ]


def build_accas(
    forecasts: Iterable[Forecast],
    *,
    legs: int = 3,
    min_legs: Optional[int] = None,
    max_legs: Optional[int] = None,
    limit: int = 5,
    allow_provisional: bool = False,
    require_disjoint_legs: bool = False,
    forbid_conflicting_game_sides: bool = False,
    forbid_alt_line_reuse: bool = False,
    min_leg_odds: Optional[float] = None,
    max_leg_odds: Optional[float] = None,
    min_leg_probability: Optional[float] = None,
    price_weight: Optional[float] = None,
    min_combined_odds: Optional[float] = None,
    max_combined_odds: Optional[float] = None,
    ranking_metric: str = RANKING_METRIC_BLEND,
    max_claimed_edge: Optional[float] = None,
    edge_flag_threshold: Optional[float] = None,
) -> List[dict]:
    """Return best-ranked accas for the requested leg count or leg range.

    Supports two ranking objectives:
    - ``probability_price_blend`` (default for legacy callers & replay parity):
      additive ``(1 - w) * log p + w * log price`` over fixed ``legs``.
    - ``kelly_log_growth``: exact full-Kelly log-wealth growth ``g*(p, O)`` over
      variable ``min_legs..max_legs`` via sorted-price branch-and-bound,
      enforcing distinct fixtures per acca while allowing distinct markets on
      the same fixture to compete.

    Same-game legs remain strictly prohibited inside every acca. When
    ``require_disjoint_legs`` is true, tickets are selected greedily without
    reusing an exact wager selection; setting ``forbid_alt_line_reuse=True``
    additionally prevents reusing the same ``(league, game, market, period,
    selection)`` across accas via alternate lines.
    """
    lo_legs = legs if min_legs is None else int(min_legs)
    hi_legs = legs if max_legs is None else int(max_legs)
    if lo_legs < 2 or hi_legs > 6 or lo_legs > hi_legs:
        raise ValueError("acca legs must be between two and six with min_legs <= max_legs")
    if limit < 1 or limit > 100:
        raise ValueError("acca limit must be between one and one hundred")
    if ranking_metric not in VALID_RANKING_METRICS:
        raise ValueError(
            f"ranking_metric must be one of {sorted(VALID_RANKING_METRICS)}, got {ranking_metric!r}"
        )
    if max_claimed_edge is not None and (
        isinstance(max_claimed_edge, bool)
        or not isinstance(max_claimed_edge, (int, float))
        or not math.isfinite(float(max_claimed_edge))
        or float(max_claimed_edge) <= 0.0
    ):
        raise ValueError("max_claimed_edge must be finite and positive")
    if edge_flag_threshold is not None and (
        isinstance(edge_flag_threshold, bool)
        or not isinstance(edge_flag_threshold, (int, float))
        or not math.isfinite(float(edge_flag_threshold))
        or float(edge_flag_threshold) <= 0.0
    ):
        raise ValueError("edge_flag_threshold must be finite and positive")

    if require_disjoint_legs:
        all_forecasts = list(forecasts)
        used_leg_keys: set[Tuple[str, str, str, str, str, str]] = set()
        used_market_selection_keys: set[Tuple[str, str, str, str, str]] = set()
        used_game_sides: set[Tuple[str, str, str]] = set()
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
                and (
                    not forbid_alt_line_reuse
                    or _market_selection_leg_key(
                        forecast.league,
                        forecast.game_id,
                        forecast.market,
                        forecast.period,
                        forecast.selection,
                    )
                    not in used_market_selection_keys
                )
                and (
                    not forbid_conflicting_game_sides
                    or _conflicting_game_side_key(
                        forecast.league,
                        forecast.game_id,
                        forecast.market,
                        forecast.selection,
                    )
                    not in used_game_sides
                )
            ]
            best = build_accas(
                remaining,
                legs=hi_legs,
                min_legs=lo_legs,
                max_legs=hi_legs,
                limit=1,
                allow_provisional=allow_provisional,
                min_leg_odds=min_leg_odds,
                max_leg_odds=max_leg_odds,
                min_leg_probability=min_leg_probability,
                price_weight=price_weight,
                min_combined_odds=min_combined_odds,
                max_combined_odds=max_combined_odds,
                ranking_metric=ranking_metric,
                max_claimed_edge=max_claimed_edge,
                edge_flag_threshold=edge_flag_threshold,
            )
            if not best:
                break
            ticket = best[0]
            ticket["rank"] = len(output) + 1
            ticket["ticket_leg_reuse_policy"] = (
                "no_reused_wager_selections_or_opposing_game_sides"
                if forbid_conflicting_game_sides
                else "no_reused_wager_selections"
            )
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
                if forbid_alt_line_reuse:
                    used_market_selection_keys.add(
                        _market_selection_leg_key(
                            leg["league"],
                            leg["game_id"],
                            leg["market"],
                            leg["period"],
                            leg["selection"],
                        )
                    )
                if forbid_conflicting_game_sides:
                    committed = _committed_game_side_key(
                        leg["league"],
                        leg["game_id"],
                        leg["market"],
                        leg["selection"],
                    )
                    if committed is not None:
                        used_game_sides.add(committed)
        return output

    if ranking_metric == RANKING_METRIC_KELLY or lo_legs != hi_legs:
        return _exact_variable_accas(
            forecasts,
            min_legs=lo_legs,
            max_legs=hi_legs,
            limit=limit,
            allow_provisional=allow_provisional,
            min_leg_odds=min_leg_odds,
            max_leg_odds=max_leg_odds,
            min_leg_probability=min_leg_probability,
            price_weight=price_weight,
            min_combined_odds=min_combined_odds,
            max_combined_odds=max_combined_odds,
            ranking_metric=ranking_metric,
            max_claimed_edge=max_claimed_edge,
            edge_flag_threshold=edge_flag_threshold,
        )

    candidates = _best_leg_per_game(
        forecasts,
        allow_provisional=allow_provisional,
        min_leg_odds=min_leg_odds,
        max_leg_odds=max_leg_odds,
        min_leg_probability=min_leg_probability,
    )
    if max_claimed_edge is not None:
        edge_cap = float(max_claimed_edge)
        candidates = [
            row
            for row in candidates
            if float(row.calibrated_probability) * float(row.decimal_price) - 1.0
            <= edge_cap
        ]
    if len(candidates) < lo_legs:
        return []
    effective_price_weight = (
        TICKET_RANKING_PRICE_WEIGHT if price_weight is None else float(price_weight)
    )
    # Enumerate in ranked order: the k-best walk below assumes candidate index
    # order is ascending cost, so the blend has to be applied before it starts.
    candidates.sort(key=lambda row: _ticket_rank(row, price_weight=effective_price_weight))
    costs = [
        -_ticket_ranking_score(row, price_weight=effective_price_weight)
        for row in candidates
    ]
    initial = tuple(range(lo_legs))
    heap: List[tuple[float, Tuple[int, ...]]] = [(sum(costs[index] for index in initial), initial)]
    visited = {initial}
    output = []
    has_odds_band = min_combined_odds is not None or max_combined_odds is not None
    popped_states = 0
    max_states = 5000 if has_odds_band else None
    while heap and len(output) < limit:
        if max_states is not None and popped_states >= max_states:
            break
        cost, state = heapq.heappop(heap)
        popped_states += 1
        selected = [candidates[index] for index in state]
        decimal_odds = math.prod(float(row.decimal_price) for row in selected)
        if not (
            (min_combined_odds is not None and decimal_odds < float(min_combined_odds))
            or (max_combined_odds is not None and decimal_odds > float(max_combined_odds))
        ):
            output.append(
                _build_ticket_dict(
                    selected,
                    rank=len(output) + 1,
                    ranking_metric=RANKING_METRIC_BLEND,
                    ranking_score=-cost,
                    price_weight=effective_price_weight,
                    allow_provisional=allow_provisional,
                    edge_flag_threshold=edge_flag_threshold,
                )
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
    "EDGE_FLAG_THRESHOLD",
    "MAX_CLAIMED_EDGE",
    "PAPER_ACCA_LIMIT",
    "PAPER_MAX_ACCA_LEGS",
    "PAPER_MAX_COMBINED_ODDS",
    "PAPER_MIN_ACCA_LEGS",
    "PAPER_MIN_COMBINED_ODDS",
    "RANKING_METRIC_BLEND",
    "RANKING_METRIC_KELLY",
    "VALID_RANKING_METRICS",
    "build_accas",
    "filter_paper_candidate_picks",
    "is_provisional_sample_only",
    "kelly_log_growth",
    "ticket_leg_pool",
]
