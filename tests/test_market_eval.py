"""Point-in-time market-evaluation lane guards: every rejection is explicit.

These tests inject exactly the failure modes the research lane must never
accept: quotes timestamped after the decision, stale quotes, duplicate
evidence, missing timestamps, incoherent books and probability pairs,
same-date straddles, pushes, unsettled games, and uncapped Kelly sizing.
"""

from __future__ import annotations

import math
from datetime import datetime, timezone

import pytest

from forecasting.market_eval import (
    EXCLUSION_AMBIGUOUS_FORECAST,
    EXCLUSION_AMBIGUOUS_QUOTE,
    EXCLUSION_INCOHERENT_PRICE_PAIR,
    EXCLUSION_INCOHERENT_PROBABILITY_PAIR,
    EXCLUSION_MISSING_FORECAST_TIMESTAMP,
    EXCLUSION_MISSING_QUOTE_TIMESTAMP,
    EXCLUSION_NO_OUTCOME,
    EXCLUSION_NO_QUOTE_IN_POLICY,
    EXCLUSION_OUTCOME_NOT_AFTER_CUTOFF,
    EXCLUSION_QUOTE_AFTER_DECISION,
    EXCLUSION_STALE_QUOTE,
    EXCLUSION_UNFROZEN_FORECAST,
    GameOutcome,
    JoinedPair,
    MarketEvalConfig,
    MarketQuote,
    ModelForecast,
    betting_diagnostics,
    date_blocked_folds,
    devig_two_way,
    evaluate_pairs,
    join_for_evaluation,
    parse_utc,
)


def utc(*parts) -> datetime:
    return datetime(*parts, tzinfo=timezone.utc)


def forecast(game="g1", p_home=0.65, captured=None, model="m1", frozen=True):
    return ModelForecast(
        game_id=game,
        model=model,
        prob_home=p_home,
        prob_away=1.0 - p_home,
        captured_at=captured if captured is not None else utc(2026, 11, 16, 10, 0),
        frozen=frozen,
    )


def quote(game="g1", home=2.10, away=1.80, observed=None, pre_match=True, stamped=True, book="Pinnacle"):
    return MarketQuote(
        game_id=game,
        book=book,
        home_price_decimal=home,
        away_price_decimal=away,
        observed_utc=(
            (observed if observed is not None else utc(2026, 11, 16, 9, 30))
            if stamped
            else None
        ),
        pre_match=pre_match,
    )


def outcome(game="g1", home=108, away=100, scheduled=None):
    return GameOutcome(
        game_id=game,
        home_final=home,
        away_final=away,
        scheduled_utc=scheduled if scheduled is not None else utc(2026, 11, 16, 18, 0),
    )


def join(forecasts, quotes, outcomes, config=None):
    return join_for_evaluation(
        forecasts, quotes, outcomes, config or MarketEvalConfig()
    )


# --- devig ---


def test_devig_removes_margin_and_preserves_normalisation():
    result = devig_two_way(2.10, 1.80)
    raw_home, raw_away = 1 / 2.10, 1 / 1.80
    assert result.margin == pytest.approx(raw_home + raw_away)
    assert result.prob_home == pytest.approx(raw_home / (raw_home + raw_away))
    assert result.prob_home + result.prob_away == pytest.approx(1.0)


def test_devig_rejects_an_impossible_book():
    # 1.20/1.20 implies a 66.7% overround: not a coherent two-way market.
    assert devig_two_way(1.20, 1.20) is None


def test_devig_rejects_implausible_prices():
    with pytest.raises(ValueError):
        devig_two_way(1.0, 2.0)
    with pytest.raises(ValueError):
        devig_two_way(None, 2.0)


def test_parse_utc_variants():
    assert parse_utc("2026-11-16T10:00:00Z") == utc(2026, 11, 16, 10, 0)
    assert parse_utc("2026-11-16T12:00:00+02:00") == utc(2026, 11, 16, 10, 0)
    assert parse_utc("2026-11-16T10:00:00").tzinfo == timezone.utc
    assert parse_utc(None) is None
    assert parse_utc("") is None


# --- join guards ---


def test_happy_path_joins_and_prices_the_model_side():
    result = join([forecast()], [quote()], [outcome()])
    assert len(result.pairs) == 1
    pair = result.pairs[0]
    assert pair.model_side == "home"
    assert pair.side_price_decimal == 2.10
    assert pair.winner == "home"
    assert pair.event_date == "2026-11-16"
    assert pair.market_prob_home == pytest.approx(devig_two_way(2.10, 1.80).prob_home)
    assert result.exclusions == {}


def test_quote_only_after_the_decision_is_rejected_even_if_pre_tip():
    # The quote exists before tip but after the forecast was frozen: it is
    # knowledge the decision did not have, so the pair must not exist.
    late = quote(observed=utc(2026, 11, 16, 10, 30))
    result = join([forecast()], [late], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_QUOTE_AFTER_DECISION] == 1


def test_newest_quote_at_or_before_the_decision_is_selected():
    older = quote(home=2.00, away=2.00, observed=utc(2026, 11, 16, 8, 0))
    newer = quote(home=2.10, away=1.80, observed=utc(2026, 11, 16, 9, 30))
    result = join([forecast()], [older, newer], [outcome()])
    assert len(result.pairs) == 1
    assert result.pairs[0].side_price_decimal == 2.10


def test_stale_quote_beyond_max_age_is_rejected():
    stale = quote(observed=utc(2026, 11, 13, 0, 0))  # 82h before the decision
    result = join([forecast()], [stale], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_STALE_QUOTE] == 1


def test_ambiguous_quote_timestamp_tie_within_one_book_is_rejected():
    first = quote(home=2.00, away=2.00)
    second = quote(home=2.10, away=1.80)
    result = join([forecast()], [first, second], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_AMBIGUOUS_QUOTE] == 1


def test_each_named_book_is_evaluated_as_a_separate_lane():
    pinnacle = quote(home=2.10, away=1.80, book="Pinnacle")
    other = quote(home=1.95, away=1.95, book="OtherBook")
    result = join([forecast()], [pinnacle, other], [outcome()])
    assert {pair.book for pair in result.pairs} == {"Pinnacle", "OtherBook"}
    assert len(result.pairs) == 2


def test_missing_quote_timestamp_is_rejected():
    result = join([forecast()], [quote(stamped=False)], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_MISSING_QUOTE_TIMESTAMP] == 1


def test_in_play_quote_is_filtered_by_the_pre_match_requirement():
    result = join([forecast()], [quote(pre_match=False)], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_NO_QUOTE_IN_POLICY] == 1


def test_incoherent_price_pair_is_rejected():
    result = join([forecast()], [quote(home=1.20, away=1.20)], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_INCOHERENT_PRICE_PAIR] == 1


def test_incoherent_probability_pair_is_rejected():
    bad = ModelForecast(
        game_id="g1", model="m1", prob_home=0.60, prob_away=0.60,
        captured_at=utc(2026, 11, 16, 10, 0), frozen=True,
    )
    result = join([bad], [quote()], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_INCOHERENT_PROBABILITY_PAIR] == 1


def test_unsettled_outcome_is_rejected():
    result = join([forecast()], [quote()], [outcome(home=None)])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_NO_OUTCOME] == 1


def test_outcome_at_or_before_the_cutoff_is_rejected():
    result = join([forecast()], [quote()], [outcome(scheduled=utc(2026, 11, 16, 10, 0))])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_OUTCOME_NOT_AFTER_CUTOFF] == 1


def test_push_is_joined_but_never_scored_or_staked():
    result = join([forecast()], [quote()], [outcome(home=100, away=100)])
    assert len(result.pairs) == 1
    assert result.pairs[0].winner == "push"
    report = evaluate_pairs(result.pairs, MarketEvalConfig())
    assert report["n_push"] == 1 and report["n_decided"] == 0
    stakes = betting_diagnostics(result.pairs, MarketEvalConfig())
    assert stakes["bets"] == 0 and stakes["cannot_evaluate"]


def test_latest_forecast_revision_wins_and_ties_are_ambiguous():
    older = forecast(p_home=0.70, captured=utc(2026, 11, 16, 9, 0))
    newer = forecast(p_home=0.65, captured=utc(2026, 11, 16, 10, 0))
    result = join([older, newer], [quote()], [outcome()])
    assert result.pairs[0].prob_home == 0.65

    tied_a = forecast(p_home=0.70)
    tied_b = forecast(p_home=0.65, model="m1")
    result = join([tied_a, tied_b], [quote()], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_AMBIGUOUS_FORECAST] == 1


def test_unfrozen_forecast_is_rejected_by_default():
    result = join([forecast(frozen=False)], [quote()], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_UNFROZEN_FORECAST] == 1


def test_distinct_provider_models_are_kept_as_separate_pairs():
    result = join(
        [forecast(model="provider-a::model-a"), forecast(model="provider-b::model-b", p_home=0.58)],
        [quote()],
        [outcome()],
    )
    assert {(pair.game_id, pair.model) for pair in result.pairs} == {
        ("g1", "provider-a::model-a"),
        ("g1", "provider-b::model-b"),
    }


def test_forecast_without_a_timestamp_is_rejected():
    stampless = ModelForecast(
        game_id="g1", model="m1", prob_home=0.65, prob_away=0.35,
        captured_at=None, frozen=True,
    )
    result = join([stampless], [quote()], [outcome()])
    assert result.pairs == []
    assert result.exclusions[EXCLUSION_MISSING_FORECAST_TIMESTAMP] == 1


# --- splits ---


def test_date_blocked_folds_never_straddle_a_date():
    pairs = [
        JoinedPair(
            game_id=f"g{i}", model="m1", prob_home=0.65, market_prob_home=0.5,
            margin=1.05, book="Pinnacle", model_side="home",
            side_price_decimal=2.0,
            winner="home" if i % 2 else "away",
            event_date=("2026-11-16" if i < 3 else "2026-11-17"),
            scheduled_utc=utc(2026, 11, 16 if i < 3 else 17, 18, 0),
            captured_at=utc(2026, 11, 16, 10, 0),
        )
        for i in range(5)
    ]
    folds = date_blocked_folds(pairs)
    assert sorted(folds) == ["2026-11-16", "2026-11-17"]
    for date, rows in folds.items():
        assert all(pair.event_date == date for pair in rows)


# --- evaluation ---


def _deterministic_pairs() -> list:
    """8 event dates x 5 games; model 0.70 vs market 0.50; 4 home wins per date."""
    pairs = []
    for day in range(1, 9):
        for game in range(5):
            home_wins = game < 4
            pairs.append(
                JoinedPair(
                    game_id=f"2026-11-{day:02d}-{game}", model="m1",
                    prob_home=0.70, market_prob_home=0.50, margin=1.05,
                    book="Pinnacle", model_side="home", side_price_decimal=2.00,
                    winner="home" if home_wins else "away",
                    event_date=f"2026-11-{day:02d}",
                    scheduled_utc=utc(2026, 11, day, 18, 0),
                    captured_at=utc(2026, 11, day, 10, 0),
                )
            )
    return pairs


def test_paired_delta_favours_the_better_model_with_reproducible_ci():
    config = MarketEvalConfig(bootstrap_iterations=500, bootstrap_seed=7)
    report = evaluate_pairs(_deterministic_pairs(), config)
    assert report["n_decided"] == 40 and report["n_push"] == 0
    # model: (4*0.09 + 1*0.49)/5 = 0.17 per date; market: 0.25 -> delta -0.08
    assert report["model_brier"] == pytest.approx(0.17)
    assert report["market_brier"] == pytest.approx(0.25)
    delta = report["paired_brier_delta"]
    assert delta["mean"] == pytest.approx(-0.08)
    assert delta["ci95_low"] <= delta["mean"] <= delta["ci95_high"]
    again = evaluate_pairs(_deterministic_pairs(), config)
    assert again["paired_brier_delta"] == delta


def test_betting_diagnostics_separate_yield_from_capital_return():
    pairs = []
    for i in range(10):
        home_wins = i >= 4  # first four lose, then six wins at 2.00
        pairs.append(
            JoinedPair(
                game_id=f"g{i}", model="m1", prob_home=0.60, market_prob_home=0.50,
                margin=1.05, book="Pinnacle", model_side="home",
                side_price_decimal=2.00,
                winner="home" if home_wins else "away",
                event_date=f"2026-11-{16 + i // 5:02d}",
                scheduled_utc=utc(2026, 11, 16, 18, i),
                captured_at=utc(2026, 11, 16, 10, 0),
            )
        )
    config = MarketEvalConfig(stake_units=1.0, bank_units=100.0)
    stakes = betting_diagnostics(pairs, config)
    all_pairs = stakes["all_pairs"]
    assert all_pairs["bets"] == 10 and all_pairs["wins"] == 6
    assert all_pairs["net_units"] == pytest.approx(2.0)
    assert all_pairs["yield_on_stake"] == pytest.approx(0.20)
    assert all_pairs["capital_return_on_stated_bank"] == pytest.approx(0.02)
    assert all_pairs["final_bank_units"] == pytest.approx(102.0)
    assert all_pairs["max_drawdown_units"] == pytest.approx(4.0)
    assert all_pairs["drawdown_basis"] == "not_a_settlement_time_equity_curve"


def test_value_subset_is_a_strict_subset_of_all_pairs():
    def make(game, p_home, winner):
        return JoinedPair(
            game_id=game, model="m1", prob_home=p_home, market_prob_home=0.70,
            margin=1.05, book="Pinnacle", model_side="home",
            side_price_decimal=2.00, winner=winner,
            event_date="2026-11-16", scheduled_utc=utc(2026, 11, 16, 18, 0),
            captured_at=utc(2026, 11, 16, 10, 0),
        )

    pairs = [make("a", 0.80, "home"), make("b", 0.55, "away")]
    stakes = betting_diagnostics(pairs, MarketEvalConfig())
    assert stakes["all_pairs"]["bets"] == 2
    assert stakes["value_subset"]["bets"] == 1


def test_fractional_kelly_is_capped_and_full_kelly_is_refused():
    pairs = _deterministic_pairs()[:5]
    config = MarketEvalConfig(kelly_fraction=0.25, max_stake_fraction=0.02)
    diagnostic = betting_diagnostics(pairs, config)["fractional_kelly_diagnostic"]
    # raw Kelly at p=0.7, price 2.0 is 0.40; 0.40 * 0.25 = 0.10 -> capped at 0.02
    assert diagnostic["max_stake_fraction_used"] == pytest.approx(0.02)
    assert diagnostic["note"].startswith("diagnostic only")
    with pytest.raises(ValueError):
        betting_diagnostics(pairs, MarketEvalConfig(kelly_fraction=1.0))
    with pytest.raises(ValueError):
        betting_diagnostics(pairs, MarketEvalConfig(kelly_fraction=None))


def test_diagnostics_without_priced_pairs_report_not_evaluable():
    stakes = betting_diagnostics([], MarketEvalConfig())
    assert stakes["cannot_evaluate"] and stakes["bets"] == 0
    report = evaluate_pairs([], MarketEvalConfig())
    assert report["n_decided"] == 0
    assert math.isnan(report["model_brier"])
