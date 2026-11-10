"""Canonical bookmaker filtering, chronology, and exact-period coverage."""

from __future__ import annotations

from pathlib import Path

import pytest

from ingestion import books
from ingestion.books import (
    BookConfig,
    PROVIDERS_BLOCKED_FROM_QUARTERS,
    american_to_decimal,
    newest_snapshot,
    quarter_coverage,
    select_moneyline_snapshots,
    select_spread_snapshots,
    select_total_snapshots,
)
from ingestion.models import BookSnapshot, Market, Period


def _snap(
    provider: str = "sportsgameodds",
    book: str = "DraftKings",
    market: Market = Market.TOTAL,
    period: Period = Period.Q1,
    line: float | None = 56.5,
    observed_utc: str = "2026-11-06T12:00:00Z",
    game_id: str = "g1",
    **extra,
) -> BookSnapshot:
    return BookSnapshot(
        game_id=game_id,
        provider=provider,
        book=book,
        market=market,
        period=period,
        observed_utc=observed_utc,
        line=line,
        **extra,
    )


def _config() -> BookConfig:
    return BookConfig(provider="sportsgameodds", book="DraftKings")


@pytest.mark.parametrize(
    ("american", "decimal"),
    [(-110, 1.9090909090909092), (150, 2.5), (100, 2.0), (-100, 2.0)],
)
def test_american_to_decimal(american, decimal):
    assert american_to_decimal(american) == pytest.approx(decimal, abs=1e-12)


@pytest.mark.parametrize("value", [0, None])
def test_american_zero_and_none_rejected(value):
    with pytest.raises(ValueError):
        american_to_decimal(value)


def test_total_filter_requires_exact_period_source_price_state_and_line():
    valid = _snap(pre_match=True, main_line=True)
    candidates = [
        valid,
        _snap(game_id="wrong-book", book="FanDuel"),
        _snap(game_id="post", pre_match=False),
        _snap(game_id="alternate", main_line=False),
        _snap(game_id="line-less", line=None),
        _snap(game_id="wrong-period", period=Period.Q2),
        _snap(game_id="wrong-market", market=Market.SPREAD, line=-1.5),
    ]
    assert select_total_snapshots(candidates, _config(), Period.Q1) == [valid]


def test_spread_filter_uses_real_home_line_and_exact_period():
    spread = _snap(market=Market.SPREAD, period=Period.Q1, line=-2.5)
    assert select_spread_snapshots([spread], _config(), Period.Q1) == [spread]
    assert select_spread_snapshots([spread], _config(), Period.Q2) == []
    assert (
        select_spread_snapshots([_snap(market=Market.SPREAD, line=None)], _config(), Period.Q1)
        == []
    )


def test_moneyline_filter_separates_full_game_and_first_half():
    full = _snap(
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        line=None,
        home_price_decimal=1.8,
        away_price_decimal=2.1,
    )
    half = _snap(
        market=Market.MONEYLINE,
        period=Period.H1,
        line=None,
        home_price_decimal=1.9,
        away_price_decimal=2.0,
        draw_price_decimal=11.0,
    )
    assert select_moneyline_snapshots([full, half], _config()) == [full]
    assert select_moneyline_snapshots([full, half], _config(), Period.H1) == [half]
    assert half.moneyline_for_side("draw") == 11.0


def test_therundown_is_blocked_only_from_non_full_game_markets():
    blocked = [_snap(provider="therundown", period=Period.Q1)]
    config = BookConfig(provider="therundown")
    assert select_total_snapshots(blocked, config, Period.Q1) == []
    full = [_snap(provider="therundown", line=232.5, period=Period.FULL_GAME)]
    assert select_total_snapshots(full, config, Period.FULL_GAME) == full
    assert "therundown" in PROVIDERS_BLOCKED_FROM_QUARTERS


def test_newest_snapshot_uses_instants_not_string_order():
    earlier = _snap(observed_utc="2026-11-06T23:30:00-02:00")  # 01:30Z next day
    later = _snap(observed_utc="2026-11-07T02:00:00Z")
    assert newest_snapshot([later, earlier]) is later


def test_quarter_coverage_counts_distinct_games_per_exact_period():
    snapshots = [
        _snap(game_id="g1", period=Period.Q1),
        _snap(game_id="g2", period=Period.Q1),
        _snap(game_id="g1", period=Period.Q2),
        _snap(game_id="g3", period=Period.Q4),
        _snap(game_id="g4", period=Period.Q1, pre_match=False),
        _snap(game_id="g5", period=Period.FULL_GAME, line=232.5),
    ]
    assert quarter_coverage(snapshots, BookConfig()) == {
        "Q1": 2,
        "Q2": 1,
        "Q3": 0,
        "Q4": 1,
    }


def test_book_module_contains_no_period_line_derivation():
    source = Path(books.__file__).read_text(encoding="utf-8")
    banned = [
        "/4",
        "/ 4",
        "//4",
        "// 4",
        "divmod(",
        "* 0.25",
        "0.25 *",
        "derive_quarter",
        "quarter_from",
        "from_full_game",
        "full_game_to_quarter",
    ]
    assert not [pattern for pattern in banned if pattern in source]
