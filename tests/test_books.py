"""Tests for ingestion/books.py — conversion, merge policy, quarter coverage.

Also proves the binding "no derivation" constraint: books.py contains no
arithmetic that could turn a full-game line/spread/price into a quarter line
(source scan below). books.py is NOT modified by these tests — coverage only.

Fixtures are synthetic/inline (DOCUMENTATION_ONLY); no network is used.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from ingestion import books
from ingestion.books import (
    BookConfig,
    ExistingCellValue,
    PROVIDERS_BLOCKED_FROM_QUARTERS,
    american_to_decimal,
    merge_book_cell,
    quarter_coverage,
    select_total_snapshots,
)
from ingestion.models import BookSnapshot, CoverageState, Market, Period


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# american_to_decimal
# ---------------------------------------------------------------------------


def test_american_negative_110_converts_to_1909090909():
    assert american_to_decimal(-110) == pytest.approx(1.9090909090909092, abs=1e-12)


def test_american_positive_150_converts_to_25():
    assert american_to_decimal(150) == pytest.approx(2.5, abs=1e-12)


def test_american_plus_100_is_2():
    assert american_to_decimal(100) == pytest.approx(2.0, abs=1e-12)


def test_american_minus_100_is_also_2_formula_pinned():
    # Documented formula behaviour: -100 is even money too, so the formula
    # yields 2.0 (1 + 100/100). Pinned so drift is noticed deliberately.
    assert american_to_decimal(-100) == pytest.approx(2.0, abs=1e-12)


def test_american_zero_and_none_rejected():
    with pytest.raises(ValueError):
        american_to_decimal(0)
    with pytest.raises(ValueError):
        american_to_decimal(None)  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# select_total_snapshots / merge_book_cell
# ---------------------------------------------------------------------------


def _config() -> BookConfig:
    return BookConfig(provider="sportsgameodds", book="DraftKings")


def test_merge_writes_fresh_incoming_q1_for_configured_book():
    incoming = [_snap(pre_match=True, main_line=True)]  # Q1 line 56.5, DraftKings
    decision = merge_book_cell(Period.Q1, incoming, None, _config())
    assert decision.action == "write"
    assert decision.value == pytest.approx(56.5)
    assert decision.source == "sportsgameodds"
    assert decision.observed_utc == "2026-11-06T12:00:00Z"


def test_merge_full_game_line_write():
    incoming = [
        _snap(market=Market.TOTAL, period=Period.FULL_GAME, line=232.5)
    ]
    decision = merge_book_cell(Period.FULL_GAME, incoming, None, _config())
    assert decision.action == "write"
    assert decision.value == pytest.approx(232.5)  # the LINE, never a price


def test_merge_keeps_manual_when_no_incoming():
    existing = ExistingCellValue(raw="55.5", source="manual", is_manual=True)
    decision = merge_book_cell(Period.Q1, [], existing, _config())
    assert decision.action == "keep_existing"
    assert decision.value == pytest.approx(55.5)
    assert decision.source == "manual"


def test_merge_blanks_stale_book_value_after_staleness_window():
    # Existing book value observed 48h ago; staleness window is 24h.
    existing = ExistingCellValue(
        raw="54.0", source="sportsgameodds", observed_utc="2026-11-04T00:00:00Z", is_manual=False
    )
    decision = merge_book_cell(
        Period.Q1, [], existing, _config(), now_utc="2026-11-06T00:00:00Z", staleness_hours=24.0
    )
    assert decision.action == "blank_with_gap"
    assert decision.value is None
    assert decision.gap is not None
    assert decision.gap.required_field == "book_q1_line"


def test_merge_keeps_manual_when_incoming_from_wrong_book():
    # Provider mismatch must not erase a valid existing manual line.
    incoming = [_snap(book="FanDuel")]  # not the configured book
    existing = ExistingCellValue(raw="56.5", source="manual", is_manual=True)
    decision = merge_book_cell(Period.Q1, incoming, existing, _config())
    assert decision.action == "keep_existing"
    assert decision.value == pytest.approx(56.5)
    assert decision.source == "manual"


def test_select_total_snapshots_blocks_therundown_from_quarters():
    blocked = [_snap(provider="therundown", line=56.5, period=Period.Q1)]
    assert select_total_snapshots(blocked, _config(), Period.Q1) == []
    # TheRundown remains allowed for the FULL-GAME total (its actual coverage).
    full = [_snap(provider="therundown", line=232.5, period=Period.FULL_GAME)]
    assert len(select_total_snapshots(full, _config(), Period.FULL_GAME)) == 1
    assert "therundown" in PROVIDERS_BLOCKED_FROM_QUARTERS


def test_merge_never_writes_therundown_quarter_snapshot():
    # Crafted therundown Q1 snapshot (defence in depth: adapter asserts too).
    incoming = [_snap(provider="therundown", period=Period.Q1, line=56.5)]
    decision = merge_book_cell(Period.Q1, incoming, None, _config())
    assert decision.action != "write"
    assert decision.value is None
    assert decision.action == "blank_with_gap"
    assert decision.gap is not None
    assert decision.gap.reason in (CoverageState.MISSING, CoverageState.UNSUPPORTED)


def test_merge_blanks_when_nothing_incoming_and_no_existing():
    decision = merge_book_cell(Period.Q3, [], None, _config(), game_id="g9")
    assert decision.action == "blank_with_gap"
    assert decision.gap is not None
    assert decision.gap.game_id == "g9"


# ---------------------------------------------------------------------------
# quarter_coverage
# ---------------------------------------------------------------------------


def test_quarter_coverage_counts_periods_separately():
    snapshots = [
        _snap(game_id="g1", period=Period.Q1),
        _snap(game_id="g2", period=Period.Q1),
        _snap(game_id="g1", period=Period.Q2),
        _snap(game_id="g3", period=Period.Q4),
        _snap(game_id="g4", period=Period.Q1, pre_match=False),  # not pre-match
        _snap(game_id="g5", period=Period.FULL_GAME, line=232.5),  # not a quarter
    ]
    counts = quarter_coverage(snapshots, _config())
    assert counts == {"Q1": 2, "Q2": 1, "Q3": 0, "Q4": 1}


# ---------------------------------------------------------------------------
# NO-DERIVATION proof (source scan)
# ---------------------------------------------------------------------------


def test_books_source_contains_no_quarter_derivation_patterns():
    """PROVE NO DERIVATION: books.py must have no code path that turns a
    full-game line/spread/price into a quarter line (no division by four,
    no quarter-fraction multipliers, no derivation helpers)."""
    source = Path(books.__file__).read_text(encoding="utf-8")
    banned_patterns = [
        "/4",        # integer style division by four
        "/ 4",       # spaced division by four
        "//4",       # floor division by four
        "// 4",      # spaced floor division by four
        "divmod(",   # any modular splitting
        "* 0.25",    # quarter-fraction multiplication
        "0.25 *",
        "* .25",
        ".25 *",
        "derive_quarter",
        "quarter_from",
        "from_full_game",
        "full_game_to_quarter",
        "full_to_quarter",
    ]
    hits = [pattern for pattern in banned_patterns if pattern in source]
    assert not hits, (
        "POSSIBLE QUARTER DERIVATION in ingestion/books.py — found banned pattern(s) "
        f"{hits}. A full-game total/spread/price must NEVER be divided or otherwise "
        "converted into a Q1-Q4 line (master prompt constraint 3, section 9)."
    )
