"""Point-in-time model-versus-market evaluation lane.

Research only. This module compares Ma Golide's frozen pre-tip game-winner
probabilities with the market-implied probabilities inside named-book decimal
prices on exactly the same games, under a pre-registered entry policy.

It deliberately does NOT:

- place bets, recommend stakes for execution, or certify an edge;
- treat an observed quote as guaranteed available execution: the warehouse's
  ``book_snapshots`` contract describes observed quotes, not fills;
- use market-maximum prices, later closing prices, or any price quoted after
  the forecast decision cutoff;
- claim profitability. A probability-quality comparison on observed quotes is
  evidence about calibration relative to the market, never about real-money
  return.

Semantics follow the existing warehouse contract:

- ``book_snapshots.observed_utc`` is the quote time; ``pre_match`` marks
  pre-tip observations; ``frozen`` marks immutable history;
- ``predictions.captured_at`` (exposed by ``prediction_facts``) is the
  decision cutoff of a pre-tip forecast. Revisions resolve to the latest
  pre-tip observation per game, mirroring the paper picks' "latest pending
  pre-tip observation per logical pick";
- moneyline ties are pushes: reported separately and excluded from
  probability scores, mirroring the assay's push handling;
- decimal prices must satisfy the same plausibility bounds as
  ``ingestion.books.is_plausible_decimal``.

Every rejection has a named reason so coverage is auditable, in the same
spirit as the assay's independent exclusion counts.
"""

from __future__ import annotations

import math
import random
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

# --- exclusion reasons (each rejection is counted, never silently dropped) ---

EXCLUSION_NO_OUTCOME = "no_settled_outcome"
EXCLUSION_OUTCOME_NOT_AFTER_CUTOFF = "outcome_not_after_cutoff"
EXCLUSION_NO_QUOTE_IN_POLICY = "no_quote_within_entry_policy"
EXCLUSION_QUOTE_AFTER_DECISION = "quote_only_after_decision"
EXCLUSION_STALE_QUOTE = "quote_stale_beyond_max_age"
EXCLUSION_AMBIGUOUS_QUOTE = "ambiguous_quote_timestamp_tie"
EXCLUSION_AMBIGUOUS_FORECAST = "ambiguous_forecast_revision_tie"
EXCLUSION_INCOHERENT_PRICE_PAIR = "incoherent_two_way_price_pair"
EXCLUSION_INCOHERENT_PROBABILITY_PAIR = "incoherent_probability_pair"
EXCLUSION_MISSING_QUOTE_TIMESTAMP = "missing_quote_timestamp"
EXCLUSION_MISSING_FORECAST_TIMESTAMP = "missing_forecast_timestamp"
EXCLUSION_UNFROZEN_FORECAST = "forecast_not_frozen"
EXCLUSION_QUOTE_MISSING_PRICES = "quote_missing_side_prices"
EXCLUSION_NO_MODEL_SIDE = "no_model_side"

# --- plausibility bounds (mirror ingestion.books) ---

MIN_DECIMAL_PRICE = 1.01
MAX_DECIMAL_PRICE = 1000.0
# A two-way book implies a margin (sum of inverse prices). Below 1.0 is an
# arbitrage artifact, above 1.5 is not a coherent two-way market.
MIN_TWO_WAY_MARGIN = 1.0
MAX_TWO_WAY_MARGIN = 1.5

LOG_LOSS_CLAMP = 1e-12


class ProfitabilityNotEvaluable(RuntimeError):
    """Raised when profit cannot be evaluated without inventing evidence."""


@dataclass(frozen=True)
class MarketEvalConfig:
    """Pre-registered evaluation policy. Every field is part of the policy."""

    book: Optional[str] = None
    require_pre_match: bool = True
    # Quote freeze is optional: quote timestamps themselves establish when a
    # price was observable. Forecast freeze is mandatory by default because
    # the evaluated decision must be the immutable pre-tip record.
    require_frozen: bool = False
    require_frozen_forecast: bool = True
    quote_max_age_hours: float = 72.0
    decision_lag_minutes: float = 0.0
    stake_units: float = 1.0
    bank_units: float = 100.0
    # Research-only Kelly diagnostic: a fractional cap, never full Kelly.
    kelly_fraction: Optional[float] = 0.25
    max_stake_fraction: float = 0.02
    bootstrap_iterations: int = 2000
    bootstrap_seed: int = 20261117

    def __post_init__(self) -> None:
        if not math.isfinite(self.quote_max_age_hours) or self.quote_max_age_hours < 0:
            raise ValueError("quote_max_age_hours must be finite and non-negative")
        if not math.isfinite(self.decision_lag_minutes) or self.decision_lag_minutes < 0:
            raise ValueError("decision_lag_minutes must be finite and non-negative")
        if not math.isfinite(self.stake_units) or self.stake_units <= 0:
            raise ValueError("stake_units must be finite and positive")
        if not math.isfinite(self.bank_units) or self.bank_units <= 0:
            raise ValueError("bank_units must be finite and positive")
        if self.kelly_fraction is not None and not 0 <= self.kelly_fraction < 1:
            raise ValueError("kelly_fraction must be fractional (0 <= fraction < 1)")
        if not math.isfinite(self.max_stake_fraction) or not 0 < self.max_stake_fraction <= 1:
            raise ValueError("max_stake_fraction must be finite and in (0, 1]")
        if self.bootstrap_iterations < 0:
            raise ValueError("bootstrap_iterations cannot be negative")


@dataclass(frozen=True)
class ModelForecast:
    """One frozen pre-tip game-winner forecast (the decision record)."""

    game_id: str
    model: str
    prob_home: float
    prob_away: float
    captured_at: Optional[datetime]
    pre_tip: bool = True
    frozen: bool = False


@dataclass(frozen=True)
class MarketQuote:
    """One named-book two-way price pair at a stated quote time."""

    game_id: str
    book: str
    home_price_decimal: Optional[float]
    away_price_decimal: Optional[float]
    observed_utc: Optional[datetime]
    pre_match: bool = True
    frozen: bool = True


@dataclass(frozen=True)
class GameOutcome:
    """One reconciled final with its scheduled tip time."""

    game_id: str
    home_final: int
    away_final: int
    scheduled_utc: Optional[datetime]


@dataclass(frozen=True)
class JoinedPair:
    """One forecast/quote/outcome triple that survived every guard."""

    game_id: str
    model: str
    prob_home: float
    market_prob_home: float
    margin: float
    book: str
    model_side: str  # "home" / "away" / "none": the side the model backed
    side_price_decimal: float  # the model's side price at decision time; NaN when no side
    winner: str  # "home" / "away" / "push"
    event_date: str
    scheduled_utc: datetime
    captured_at: datetime


@dataclass(frozen=True)
class DevigResult:
    """Market-implied probabilities after multiplicative margin removal."""

    prob_home: float
    prob_away: float
    margin: float


def parse_utc(value: object) -> Optional[datetime]:
    """Parse ISO-8601 UTC timestamps ('Z' or offset); naive values are UTC.

    Warehouse timestamps are VARCHAR ISO strings; a missing value stays
    ``None`` so callers can count it as an explicit exclusion.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        moment = value
    else:
        text = str(value).strip()
        if not text:
            return None
        moment = datetime.fromisoformat(text.replace("Z", "+00:00"))
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(timezone.utc)


def _plausible_decimal(price: Optional[float]) -> bool:
    return (
        price is not None
        and isinstance(price, (int, float))
        and math.isfinite(float(price))
        and MIN_DECIMAL_PRICE <= float(price) <= MAX_DECIMAL_PRICE
    )


def devig_two_way(
    home_price_decimal: Optional[float], away_price_decimal: Optional[float]
) -> Optional[DevigResult]:
    """Remove the book margin from a two-way price pair (multiplicative).

    Returns ``None`` when the pair does not describe a coherent two-way book
    (a side is missing/implausible, or the implied margin falls outside
    ``[MIN_TWO_WAY_MARGIN, MAX_TWO_WAY_MARGIN]`` — for example both sides
    priced as favorites). Single implausible prices raise, mirroring the
    ingestion contract; the margin band is a data-driven exclusion.
    """
    if not _plausible_decimal(home_price_decimal) or not _plausible_decimal(away_price_decimal):
        raise ValueError("both decimal prices are required and must be plausible")
    raw_home = 1.0 / float(home_price_decimal)
    raw_away = 1.0 / float(away_price_decimal)
    margin = raw_home + raw_away
    if not (MIN_TWO_WAY_MARGIN <= margin <= MAX_TWO_WAY_MARGIN):
        return None
    return DevigResult(
        prob_home=raw_home / margin,
        prob_away=raw_away / margin,
        margin=margin,
    )


def _coherent_probability_pair(prob_home: Optional[float], prob_away: Optional[float]) -> bool:
    if prob_home is None or prob_away is None:
        return False
    if not math.isfinite(prob_home) or not math.isfinite(prob_away):
        return False
    if not (0.0 <= prob_home <= 1.0 and 0.0 <= prob_away <= 1.0):
        return False
    return 0.98 <= prob_home + prob_away <= 1.02


def _latest_record(
    records: List,
    timestamp_field: str,
    missing_key: str,
    tie_key: str,
    exclusions: Dict[str, int],
):
    """Select the latest record by timestamp.

    Records with a missing timestamp are counted under ``missing_key``; a tie
    at the newest timestamp is ambiguous evidence, counted under ``tie_key``.
    """
    stamped = [r for r in records if getattr(r, timestamp_field) is not None]
    if len(stamped) != len(records):
        exclusions[missing_key] = exclusions.get(missing_key, 0) + (
            len(records) - len(stamped)
        )
    if not stamped:
        return None
    newest = max(stamped, key=lambda r: getattr(r, timestamp_field))
    tied = [
        r for r in stamped if getattr(r, timestamp_field) == getattr(newest, timestamp_field)
    ]
    if len(tied) > 1:
        exclusions[tie_key] = exclusions.get(tie_key, 0) + 1
        return None
    return newest


def join_for_evaluation(
    forecasts: List[ModelForecast],
    quotes: List[MarketQuote],
    outcomes: List[GameOutcome],
    config: MarketEvalConfig,
) -> "EvalResult":
    """Join the three evidence streams under the pre-registered entry policy.

    Entry policy: for each game and provider/model version, the forecast is
    the latest frozen pre-tip revision; the price is the newest named-book
    quote observed at or before that forecast's decision cutoff (minus the
    declared lag), within the maximum quote age. Quotes after the decision
    cutoff are never used, even when they are still pre-tip. Every rejected
    game/model pair is counted under a named reason.
    """
    exclusions: Dict[str, int] = {}
    quotes_by_game: Dict[str, List[MarketQuote]] = {}
    for quote in quotes:
        quotes_by_game.setdefault(quote.game_id, []).append(quote)
    outcomes_by_game = {o.game_id: o for o in outcomes}

    # Keep each provider/model-version lane independent. Selecting only the
    # latest forecast across all providers would create a changing, opaque
    # ensemble and hide acquired forecast information from the comparison.
    forecasts_by_game_model: Dict[Tuple[str, str], List[ModelForecast]] = {}
    for forecast in forecasts:
        if not forecast.pre_tip:
            continue
        if config.require_frozen_forecast and not forecast.frozen:
            exclusions[EXCLUSION_UNFROZEN_FORECAST] = (
                exclusions.get(EXCLUSION_UNFROZEN_FORECAST, 0) + 1
            )
            continue
        forecasts_by_game_model.setdefault((forecast.game_id, forecast.model), []).append(
            forecast
        )

    pairs: List[JoinedPair] = []
    games_considered = len(forecasts_by_game_model)
    for (game_id, model), revisions in sorted(forecasts_by_game_model.items()):
        forecast = _latest_record(
            revisions,
            "captured_at",
            EXCLUSION_MISSING_FORECAST_TIMESTAMP,
            EXCLUSION_AMBIGUOUS_FORECAST,
            exclusions,
        )
        if forecast is None:
            continue
        if not _coherent_probability_pair(forecast.prob_home, forecast.prob_away):
            exclusions[EXCLUSION_INCOHERENT_PROBABILITY_PAIR] = (
                exclusions.get(EXCLUSION_INCOHERENT_PROBABILITY_PAIR, 0) + 1
            )
            continue
        outcome = outcomes_by_game.get(game_id)
        if outcome is None or outcome.home_final is None or outcome.away_final is None:
            exclusions[EXCLUSION_NO_OUTCOME] = exclusions.get(EXCLUSION_NO_OUTCOME, 0) + 1
            continue
        if outcome.scheduled_utc is None or forecast.captured_at is None:
            exclusions[EXCLUSION_MISSING_FORECAST_TIMESTAMP] = (
                exclusions.get(EXCLUSION_MISSING_FORECAST_TIMESTAMP, 0) + 1
            )
            continue
        if outcome.scheduled_utc <= forecast.captured_at:
            exclusions[EXCLUSION_OUTCOME_NOT_AFTER_CUTOFF] = (
                exclusions.get(EXCLUSION_OUTCOME_NOT_AFTER_CUTOFF, 0) + 1
            )
            continue
        if outcome.home_final == outcome.away_final:
            winner = "push"
        else:
            winner = "home" if outcome.home_final > outcome.away_final else "away"

        decision_cutoff = forecast.captured_at
        if config.decision_lag_minutes:
            decision_cutoff = decision_cutoff - timedelta(minutes=config.decision_lag_minutes)
        game_quotes = quotes_by_game.get(game_id, [])
        if config.book is not None:
            game_quotes = [q for q in game_quotes if q.book == config.book]
        if config.require_pre_match:
            game_quotes = [q for q in game_quotes if q.pre_match]
        if config.require_frozen:
            game_quotes = [q for q in game_quotes if q.frozen]
        if not game_quotes:
            exclusions[EXCLUSION_NO_QUOTE_IN_POLICY] = (
                exclusions.get(EXCLUSION_NO_QUOTE_IN_POLICY, 0) + 1
            )
            continue

        # Do not select the globally newest quote across different books: that
        # would silently mix execution venues and discard other acquired prices.
        # Produce one comparable model/book lane for every named book instead.
        quotes_by_book: Dict[str, List[MarketQuote]] = {}
        for quote in game_quotes:
            quotes_by_book.setdefault(quote.book, []).append(quote)
        for _, book_quotes in sorted(quotes_by_book.items()):
            unstamped = [q for q in book_quotes if q.observed_utc is None]
            if unstamped:
                exclusions[EXCLUSION_MISSING_QUOTE_TIMESTAMP] = (
                    exclusions.get(EXCLUSION_MISSING_QUOTE_TIMESTAMP, 0) + 1
                )
            stamped = [q for q in book_quotes if q.observed_utc is not None]
            if not stamped:
                continue
            before_decision = [q for q in stamped if q.observed_utc <= decision_cutoff]
            if not before_decision:
                exclusions[EXCLUSION_QUOTE_AFTER_DECISION] = (
                    exclusions.get(EXCLUSION_QUOTE_AFTER_DECISION, 0) + 1
                )
                continue
            max_age_seconds = config.quote_max_age_hours * 3600.0
            fresh = [
                q
                for q in before_decision
                if (decision_cutoff - q.observed_utc).total_seconds() <= max_age_seconds
            ]
            if not fresh:
                exclusions[EXCLUSION_STALE_QUOTE] = (
                    exclusions.get(EXCLUSION_STALE_QUOTE, 0) + 1
                )
                continue
            quote = _latest_record(
                fresh,
                "observed_utc",
                EXCLUSION_MISSING_QUOTE_TIMESTAMP,
                EXCLUSION_AMBIGUOUS_QUOTE,
                exclusions,
            )
            if quote is None:
                continue
            if quote.home_price_decimal is None or quote.away_price_decimal is None:
                exclusions[EXCLUSION_QUOTE_MISSING_PRICES] = (
                    exclusions.get(EXCLUSION_QUOTE_MISSING_PRICES, 0) + 1
                )
                continue
            try:
                devig = devig_two_way(quote.home_price_decimal, quote.away_price_decimal)
            except ValueError:
                devig = None
            if devig is None:
                exclusions[EXCLUSION_INCOHERENT_PRICE_PAIR] = (
                    exclusions.get(EXCLUSION_INCOHERENT_PRICE_PAIR, 0) + 1
                )
                continue
            if forecast.prob_home > 0.5:
                model_side = "home"
                side_price = quote.home_price_decimal
            elif forecast.prob_home < 0.5:
                model_side = "away"
                side_price = quote.away_price_decimal
            else:
                model_side = "none"
                side_price = math.nan
            pairs.append(
                JoinedPair(
                    game_id=game_id,
                    model=forecast.model,
                    prob_home=forecast.prob_home,
                    market_prob_home=devig.prob_home,
                    margin=devig.margin,
                    book=quote.book,
                    model_side=model_side,
                    side_price_decimal=side_price,
                    winner=winner,
                    event_date=outcome.scheduled_utc.date().isoformat(),
                    scheduled_utc=outcome.scheduled_utc,
                    captured_at=forecast.captured_at,
                )
            )
    return EvalResult(pairs=pairs, exclusions=exclusions, games_considered=games_considered)


@dataclass(frozen=True)
class EvalResult:
    """The join outcome: surviving pairs plus auditable exclusion counts."""

    pairs: List[JoinedPair]
    exclusions: Dict[str, int] = field(default_factory=dict)
    games_considered: int = 0


def _brier(probability: float, won: bool) -> float:
    target = 1.0 if won else 0.0
    return (probability - target) ** 2


def _log_loss(probability: float, won: bool) -> float:
    p = min(max(probability, LOG_LOSS_CLAMP), 1.0 - LOG_LOSS_CLAMP)
    target = 1.0 if won else 0.0
    return -(math.log(p) if target else math.log(1.0 - p))


def _date_blocked_groups(pairs: List[JoinedPair]) -> Dict[str, List[JoinedPair]]:
    """Group pairs by event date so a date never straddles a resample."""
    groups: Dict[str, List[JoinedPair]] = {}
    for pair in pairs:
        groups.setdefault(pair.event_date, []).append(pair)
    return groups


def date_blocked_folds(pairs: List[JoinedPair]) -> Dict[str, List[JoinedPair]]:
    """Return pairs grouped by event date; same-day games share one fold."""
    return _date_blocked_groups(pairs)


def _paired_delta_ci(
    deltas_by_date: Dict[str, List[float]],
    iterations: int,
    seed: int,
) -> Tuple[float, float, float]:
    """Date-block bootstrap CI for the mean paired delta.

    Whole event dates are resampled with replacement, so games sharing a date
    move together and within-date correlation cannot deflate the interval.
    """
    rng = random.Random(seed)
    dates = sorted(deltas_by_date)
    if not dates:
        return (float("nan"), float("nan"), float("nan"))
    all_deltas = [d for day in dates for d in deltas_by_date[day]]
    point = sum(all_deltas) / len(all_deltas)
    if iterations <= 0 or len(dates) < 2:
        return (point, float("nan"), float("nan"))
    means = []
    for _ in range(iterations):
        sample = []
        for _ in range(len(dates)):
            sample.extend(deltas_by_date[rng.choice(dates)])
        means.append(sum(sample) / len(sample))
    means.sort()
    low = means[max(0, int(0.025 * len(means)) - 1)]
    high = means[min(len(means) - 1, int(0.975 * len(means)))]
    return (point, low, high)


def probability_band(probability: float) -> str:
    """Band label mirroring the operator report's probability bands."""
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


def odds_band(price: float) -> str:
    """Band label mirroring the operator report's odds bands."""
    if price < 1.20:
        return "1.00-<1.20"
    if price < 1.50:
        return "1.20-<1.50"
    if price < 2.00:
        return "1.50-<2.00"
    if price < 3.00:
        return "2.00-<3.00"
    return "3.00+"


def evaluate_pairs(pairs: List[JoinedPair], config: MarketEvalConfig) -> Dict:
    """Probability-quality comparison on identical games (pushes excluded).

    Reports paired Brier/log-loss deltas between the frozen forecast and the
    devigged market probability, with a date-block bootstrap interval. This is
    a comparison, not an edge certificate.
    """
    decided = [p for p in pairs if p.winner != "push"]
    pushes = [p for p in pairs if p.winner == "push"]

    def _scores(rows: List[JoinedPair], probability_field: str) -> Tuple[float, float]:
        if not rows:
            return (float("nan"), float("nan"))
        briers, log_losses = [], []
        for row in rows:
            won = row.winner == "home"
            probability = getattr(row, probability_field)
            briers.append(_brier(probability, won))
            log_losses.append(_log_loss(probability, won))
        return (sum(briers) / len(briers), sum(log_losses) / len(log_losses))

    model_brier, model_logloss = _scores(decided, "prob_home")
    market_brier, market_logloss = _scores(decided, "market_prob_home")

    deltas_by_date: Dict[str, List[float]] = {}
    logloss_deltas_by_date: Dict[str, List[float]] = {}
    for pair in decided:
        won = pair.winner == "home"
        deltas_by_date.setdefault(pair.event_date, []).append(
            _brier(pair.prob_home, won) - _brier(pair.market_prob_home, won)
        )
        logloss_deltas_by_date.setdefault(pair.event_date, []).append(
            _log_loss(pair.prob_home, won) - _log_loss(pair.market_prob_home, won)
        )
    brier_point, brier_low, brier_high = _paired_delta_ci(
        deltas_by_date, config.bootstrap_iterations, config.bootstrap_seed
    )
    logloss_point, logloss_low, logloss_high = _paired_delta_ci(
        logloss_deltas_by_date,
        config.bootstrap_iterations,
        config.bootstrap_seed + 1,
    )

    margins = [p.margin for p in decided]
    band_counts: Dict[str, int] = {}
    for pair in decided:
        key = probability_band(pair.prob_home)
        band_counts[key] = band_counts.get(key, 0) + 1

    return {
        "n_pairs": len(pairs),
        "n_decided": len(decided),
        "n_push": len(pushes),
        "model_brier": model_brier,
        "market_brier": market_brier,
        "paired_brier_delta": {
            "mean": brier_point,
            "ci95_low": brier_low,
            "ci95_high": brier_high,
            "negative_is_model_better": True,
        },
        "model_log_loss": model_logloss,
        "market_log_loss": market_logloss,
        "paired_log_loss_delta": {
            "mean": logloss_point,
            "ci95_low": logloss_low,
            "ci95_high": logloss_high,
            "negative_is_model_better": True,
        },
        "mean_book_margin": (
            sum(margins) / len(margins) if margins else float("nan")
        ),
        "probability_bands": band_counts,
        "interpretation": (
            "paired probability-quality comparison on identical games under a "
            "pre-registered entry policy; not evidence of profitability"
        ),
    }


def _flat_stake_path(rows: List[JoinedPair], config: MarketEvalConfig) -> Dict:
    """Flat-stake replay of the model side at the decision-time book price."""
    bank = config.bank_units
    peak = bank
    max_drawdown = 0.0
    net = 0.0
    stakes = 0.0
    wins = 0
    for pair in rows:
        stakes += config.stake_units
        if pair.model_side == pair.winner:
            net += config.stake_units * (pair.side_price_decimal - 1.0)
            wins += 1
        else:
            net -= config.stake_units
        equity = bank + net
        peak = max(peak, equity)
        max_drawdown = max(max_drawdown, peak - equity)
    return {
        "bets": len(rows),
        "wins": wins,
        "stakes": stakes,
        "net_units": net,
        "yield_on_stake": (net / stakes) if stakes else float("nan"),
        "capital_return_on_stated_bank": (
            net / config.bank_units if config.bank_units else float("nan")
        ),
        "final_bank_units": bank + net,
        "max_drawdown_units": max_drawdown,
        "drawdown_basis": "not_a_settlement_time_equity_curve",
    }


def _fractional_kelly_diagnostic(rows: List[JoinedPair], config: MarketEvalConfig) -> Dict:
    """Capped fractional-Kelly stake fractions; full Kelly is refused."""
    if config.kelly_fraction is None or config.kelly_fraction >= 1.0:
        raise ValueError(
            "full Kelly is not permitted in the research lane; "
            "use a fractional cap (kelly_fraction < 1.0)"
        )
    fractions = []
    for pair in rows:
        if pair.winner == "push" or math.isnan(pair.side_price_decimal):
            continue
        price = pair.side_price_decimal
        b = price - 1.0
        if b <= 0:
            continue
        p = pair.prob_home if pair.model_side == "home" else 1.0 - pair.prob_home
        raw = (p * (b + 1.0) - 1.0) / b
        if raw <= 0:
            fractions.append(0.0)
            continue
        fractions.append(min(raw * config.kelly_fraction, config.max_stake_fraction))
    return {
        "kelly_fraction_cap": config.kelly_fraction,
        "max_stake_fraction": config.max_stake_fraction,
        "bets": len(fractions),
        "mean_stake_fraction": (
            sum(fractions) / len(fractions) if fractions else float("nan")
        ),
        "max_stake_fraction_used": max(fractions) if fractions else 0.0,
        "note": "diagnostic only; never an execution recommendation",
    }


def _is_value_pair(pair: JoinedPair) -> bool:
    """A value bet: the model's side is longer-priced than the market implies."""
    if pair.model_side == "home":
        return pair.prob_home > pair.market_prob_home
    if pair.model_side == "away":
        return pair.prob_home < pair.market_prob_home
    return False


def betting_diagnostics(pairs: List[JoinedPair], config: MarketEvalConfig) -> Dict:
    """Paper-only stake diagnostics at decision-time named-book prices.

    The selection policy is pre-registered: every joined pair with a model
    side is staked flat on that side at the quote captured at or before the
    decision cutoff. Yield on total stake is the primary rate; return on the
    stated bank is reported separately so the two are never conflated.
    """
    decided = [p for p in pairs if p.winner != "push" and not math.isnan(p.side_price_decimal)]
    no_side = [p for p in decided if p.model_side == "none"]
    bet_rows = [p for p in decided if p.model_side != "none"]
    if not bet_rows:
        return {
            "cannot_evaluate": "no timestamped pre-tip named-book price pairs",
            "bets": 0,
            "no_model_side": len(no_side),
        }
    ordered = sorted(bet_rows, key=lambda p: (p.scheduled_utc, p.game_id))
    value_rows = [p for p in ordered if _is_value_pair(p)]
    bands: Dict[str, Dict[str, float]] = {}
    for pair in ordered:
        key = f"p={probability_band(pair.prob_home)} odds={odds_band(pair.side_price_decimal)}"
        entry = bands.setdefault(key, {"n": 0, "wins": 0, "net_units": 0.0, "stakes": 0.0})
        entry["n"] += 1
        entry["stakes"] += config.stake_units
        if pair.model_side == pair.winner:
            entry["wins"] += 1
            entry["net_units"] += config.stake_units * (pair.side_price_decimal - 1.0)
        else:
            entry["net_units"] -= config.stake_units
    for entry in bands.values():
        entry["yield_on_stake"] = (
            entry["net_units"] / entry["stakes"] if entry["stakes"] else float("nan")
        )
    return {
        "selection_policy": "all_joined_pairs_flat_stake_on_model_side",
        "all_pairs": _flat_stake_path(ordered, config),
        "value_subset": _flat_stake_path(
            sorted(value_rows, key=lambda p: (p.scheduled_utc, p.game_id)), config
        ),
        "fractional_kelly_diagnostic": _fractional_kelly_diagnostic(ordered, config),
        "bands": bands,
        "note": (
            "observed quotes are not guaranteed execution; a positive paper "
            "replay is not evidence of profitability"
        ),
    }
