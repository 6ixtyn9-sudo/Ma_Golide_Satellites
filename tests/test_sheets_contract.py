"""Contract tests: sheets contract detection and row rendering (section 6)."""

import datetime as dt

import pytest

from ingestion.models import (
    ActualResult,
    FinalStatus,
    Game,
    GameStatus,
    League,
    PeriodKind,
    Prediction,
)
from ingestion.sheets_contract import (
    BASE_HEADERS,
    COL,
    OWNED_COLUMNS,
    ContractError,
    detect_contract,
    render_result_row,
    render_upcoming_row,
)

TZ = "America/New_York"


def _headers(extra=()):
    return list(BASE_HEADERS) + list(extra)


def _golden_pred():
    return Prediction(
        game_id="401909938",
        provider="espn",
        observed_utc="2026-11-06T21:51:00Z",
        prob_home=0.23062,
        prob_away=0.76938,
        expected_home_score=110.540,
        expected_away_score=120.046,
        expected_total=230.586,
        expected_margin=9.506,
    )


def _golden_game():
    return Game(
        game_id="401909938",
        league=League.NBA,
        season="2026",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
        status=GameStatus.SCHEDULED,
    )


# ---------------------------------------------------------------------------
# Header validation
# ---------------------------------------------------------------------------


def test_golden_headers_pass():
    contract = detect_contract(_headers(["t2-q1-conf"]), "upcoming")
    assert contract.base_ok
    assert contract.extra_headers() == ["t2-q1-conf"]


def test_missing_base_header_fails():
    headers = _headers()
    headers[10] = "OddsX"  # column K renamed
    with pytest.raises(ContractError) as e:
        detect_contract(headers, "upcoming")
    assert "drift" in str(e.value)


def test_shifted_base_header_fails():
    headers = _headers()
    headers[5], headers[6] = headers[6], headers[5]
    with pytest.raises(ContractError):
        detect_contract(headers, "upcoming")


def test_duplicate_unsuffixed_base_header_fails():
    headers = _headers(["ou-q1", "Q1"])  # a stray second Q1 after V
    with pytest.raises(ContractError) as e:
        detect_contract(headers, "upcoming")
    assert "duplicate" in str(e.value)


def test_short_header_row_fails():
    with pytest.raises(ContractError):
        detect_contract(BASE_HEADERS[:10], "results")


# ---------------------------------------------------------------------------
# Upcoming row semantics
# ---------------------------------------------------------------------------


def test_golden_upcoming_row_renders_exact_display():
    cells = render_upcoming_row(
        _golden_game(), _golden_pred(), 1.55,
        {"Q1": 58.5, "Q2": 57.0, "Q3": 57.5, "Q4": 59.0}, 232.0, TZ,
    )
    # Identity block.
    assert cells[COL["League"]] == "NBA"
    assert cells[COL["Game Type"]] == "League"
    assert cells[COL["Home"]] == "Indiana Pacers"
    assert cells[COL["Away"]] == "Denver Nuggets"
    local_date, local_time = cells[COL["Date"]], cells[COL["Time"]]
    assert isinstance(local_date, dt.date) and isinstance(local_time, dt.time)
    # 22:00Z -> 17:00 America/New_York in November (EST after DST end).
    assert local_time == dt.time(17, 0)
    # Coherent prediction bundle.
    assert cells[COL["Prob %"]] == "23.1 - 76.9"
    assert cells[COL["Pred"]] == 2
    assert cells[COL["Pred Score"]] == "111-120"
    assert cells[COL["Avg"]] == 230.6
    # Odds carries the predicted side's decimal price.
    assert cells[COL["Odds"]] == 1.55
    # Q1..Q4 and FT Score carry BOOK lines, numeric.
    assert cells[COL["Q1"]] == 58.5 and cells[COL["Q4"]] == 59.0
    assert cells[COL["FT Score"]] == 232.0
    # Upcoming OT/Status stay blank.
    assert cells[COL["OT"]] == "" and cells[COL["Status"]] == ""
    # No t2 columns in the rendered set at all.
    for name in ("t2-q1", "t2-q2", "t2-q3", "t2-q4"):
        assert COL[name] not in cells


def test_upcoming_row_blank_prediction_and_books():
    cells = render_upcoming_row(_golden_game(), None, None, {}, None, TZ)
    assert cells[COL["Prob %"]] == ""
    assert cells[COL["Pred"]] == ""
    assert cells[COL["Pred Score"]] == ""
    assert cells[COL["Avg"]] == ""
    assert cells[COL["Odds"]] == ""
    assert cells[COL["Q1"]] == "" and cells[COL["FT Score"]] == ""


def test_upcoming_avg_never_a_book_line():
    """Avg renders the MODEL total even when a book line also exists."""
    cells = render_upcoming_row(_golden_game(), _golden_pred(), None, {}, 233.5, TZ)
    assert cells[COL["Avg"]] == 230.6
    assert cells[COL["FT Score"]] == 233.5


def test_exact_tie_probability_prefers_away():
    pred = Prediction(
        game_id="x", provider="espn", observed_utc="2026-11-06T21:51:00Z",
        prob_home=0.5, prob_away=0.5,
    )
    cells = render_upcoming_row(_golden_game(), pred, None, {}, None, TZ)
    assert cells[COL["Pred"]] == 2


# ---------------------------------------------------------------------------
# Result row semantics
# ---------------------------------------------------------------------------


def _final_game(overtime=None):
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[25.0, 27.0, 25.0, 24.0],
        away_quarters=[24.0, 37.0, 25.0, 19.0],
        home_overtime=overtime[0] if overtime else None,
        away_overtime=overtime[1] if overtime else None,
        home_final=101.0 if not overtime else 101.0 + overtime[0],
        away_final=105.0 if not overtime else 105.0 + overtime[1],
        overtime_periods=1 if overtime else 0,
        source="espn",
    )
    return Game(
        game_id="401909936", league=League.NBA, season="2026",
        home_team_id="1", away_team_id="2",
        home_name="Team A", away_name="Team B",
        scheduled_utc="2026-11-06T02:00:00Z", status=GameStatus.FINAL,
        actual=actual,
    )


def test_result_row_quarter_pairs_and_final():
    cells = render_result_row(_final_game(), _golden_pred(), 1.34, TZ)
    assert cells[COL["Q1"]] == "25 - 24"
    assert cells[COL["Q2"]] == "27 - 37"
    assert cells[COL["OT"]] == ""
    assert cells[COL["Status"]] == "FT"
    assert cells[COL["FT Score"]] == "101 - 105"
    # Frozen pre-tip bundle carries over.
    assert cells[COL["Prob %"]] == "23.1 - 76.9"
    assert cells[COL["Odds"]] == 1.34


def test_result_row_aot_aggregated_ot():
    cells = render_result_row(_final_game((6.0, 4.0)), None, None, TZ)
    assert cells[COL["OT"]] == "6 - 4"
    assert cells[COL["Status"]] == "AOT"
    assert cells[COL["FT Score"]] == "107 - 109"


def test_result_row_without_actual_stays_blank():
    game = Game(
        game_id="x", league=League.NBA, season="2026",
        home_team_id="1", away_team_id="2",
        home_name="Team A", away_name="Team B",
        scheduled_utc="2026-11-06T02:00:00Z", status=GameStatus.FINAL,
        actual=None,
    )
    cells = render_result_row(game, None, None, TZ)
    assert cells[COL["FT Score"]] == ""
    assert cells[COL["Status"]] == ""
    assert cells[COL["Q1"]] == ""


def test_result_row_t2_columns_never_rendered():
    cells = render_result_row(_final_game(), _golden_pred(), None, TZ)
    assert set(cells) <= set(OWNED_COLUMNS)
    assert COL["t2-q1"] not in cells
