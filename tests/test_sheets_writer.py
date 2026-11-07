"""Writer safety tests: non-destructive merges, idempotency, protection."""

import datetime as dt

import pytest

from ingestion.sheets_contract import BASE_HEADERS, COL, ContractError, detect_contract
from ingestion.sheets_writer import (
    FakeWorksheet,
    UpdatePlan,
    apply_plan,
    column_letter,
    plan_update,
)

BASE = list(BASE_HEADERS) + [
    "t2-q1-conf", "ou-q1", "ou-best",  # extra model columns after V
]


def _contract():
    return detect_contract(BASE, "upcoming")


def _ws(rows=None):
    return FakeWorksheet(BASE, rows or [])


def _row(home="Indiana Pacers", away="Denver Nuggets", date="2026-11-07", **overrides):
    """Rendered cell dict; overrides accept base header names or column ints."""
    cells = {
        COL["Home"]: home, COL["Away"]: away, COL["Date"]: date,
        COL["League"]: "NBA", COL["Game Type"]: "League",
        COL["Prob %"]: "23.1 - 76.9", COL["Pred"]: 2,
        COL["FT Score"]: "",
    }
    for key, value in overrides.items():
        cells[COL[key] if isinstance(key, str) else key] = value
    return cells


# ---------------------------------------------------------------------------


def test_plan_appends_new_rows_only_in_owned_columns():
    ws = _ws()
    plan = plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")
    assert len(plan.appends) == 1 and not plan.changes
    assert all(v in ("",) or v is not None for v in plan.appends[0])
    # protected S:V + extras padded blank on append
    row = plan.appends[0]
    assert row[COL["t2-q1"] - 1] == ""
    assert row[22] == "" and row[23] == "" and row[24] == ""


def test_plan_is_idempotent_second_run_noop():
    ws = _ws()
    plan1 = plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")
    apply_plan(plan1, ws, allow_writes=True, sleep=lambda s: None)
    plan2 = plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")
    assert plan2.is_noop, plan2.to_dict()
    assert len(ws.rows()) == 1  # no duplicate rows


def test_plan_updates_only_changed_cells():
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    existing[COL["Q1"] - 1] = "58.5"  # manual book line already present
    existing[COL["t2-q1"] - 1] = "N/A"        # t2-q1 (protected)
    existing[23] = 0.5                        # t2-q1-conf (extra after V)
    ws = _ws([existing])
    plan = plan_update(
        "UpcomingClean", ws, _contract(),
        [_row(**{"Q1": 58.5, "FT Score": 232.0})], "upcoming",
    )
    changed_cols = {c.col for c in plan.changes}
    assert COL["FT Score"] in changed_cols
    assert COL["Q1"] not in changed_cols  # identical value -> no write
    assert COL["Prob %"] in changed_cols  # filled in
    # No protected column ever appears.
    assert not any(c.col >= 19 for c in plan.changes)


def test_plan_preserves_non_target_columns_after_write():
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    existing[18] = "N/A"      # t2-q1
    existing[19] = "N/A"      # t2-q2
    existing[22] = "EVEN"     # ou-q1-ish extra column W
    existing[24] = "MODEL"    # extra column Y
    ws = _ws([existing])
    plan = plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")
    apply_plan(plan, ws, allow_writes=True, sleep=lambda s: None)
    row = ws.rows()[0]
    assert row[18] == "N/A" and row[19] == "N/A"
    assert row[22] == "EVEN" and row[24] == "MODEL"


def test_extended_workbook_columns_through_bx_survive_update_and_append():
    headers = list(BASE_HEADERS) + [f"model_{col}" for col in range(23, 77)]
    existing = [""] * 76
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    for index in range(18, 76):
        existing[index] = f"=PRESERVE({index + 1})" if index % 2 else f"value-{index + 1}"
    protected_before = list(existing[18:])
    ws = FakeWorksheet(headers, [existing])
    contract = detect_contract(headers, "upcoming")

    plan = plan_update("UpcomingClean", ws, contract, [_row()], "upcoming")
    assert all(change.col <= 18 for change in plan.changes)
    apply_plan(plan, ws, allow_writes=True, sleep=lambda seconds: None)
    assert ws.rows()[0][18:] == protected_before

    appended = plan_update(
        "UpcomingClean", ws, contract, [_row(home="A", away="B")], "upcoming"
    )
    assert len(appended.appends[0]) == 76
    assert appended.appends[0][18:] == [""] * 58


def test_plan_refuses_protected_column_render():
    cells = _row()
    cells[COL["t2-q1"]] = "0.0"  # acquisition must never render S..V
    with pytest.raises(ContractError) as e:
        plan_update("UpcomingClean", _ws(), _contract(), [cells], "upcoming")
    assert "non-owned" in str(e.value)


def test_plan_refuses_formula_overwrite():
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    existing[COL["Odds"] - 1] = "=IF(1=1,1.5,2)"
    ws = _ws([existing])
    plan = plan_update("UpcomingClean", ws, _contract(),
                       [_row(**{"Odds": 1.55})], "upcoming")
    # The formula cell is skipped; unrelated identity cells may still fill.
    assert COL["Odds"] not in {c.col for c in plan.changes}
    assert plan.skipped_formulas
    assert "=IF(1=1,1.5,2)" in ws.rows()[0][COL["Odds"] - 1]


def test_plan_fails_on_duplicate_sheet_identity():
    rows = [[""] * 25 for _ in range(2)]
    for r in rows:
        r[COL["Home"] - 1] = "Indiana Pacers"
        r[COL["Away"] - 1] = "Denver Nuggets"
        r[COL["Date"] - 1] = "2026-11-07"
    with pytest.raises(ContractError) as e:
        plan_update("UpcomingClean", _ws(rows), _contract(), [_row()], "upcoming")
    assert "duplicate" in str(e.value)


def test_plan_fails_on_header_drift():
    ws = FakeWorksheet(["Wrong"] + BASE[1:], [])
    with pytest.raises(ContractError):
        plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")


def test_ambiguous_legacy_date_matches_under_both_readings():
    """``11/01/2026`` matches a 2026-01-11 rendered row (day-first legacy)."""
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "11/01/2026"
    ws = _ws([existing])
    plan = plan_update(
        "UpcomingClean", ws, _contract(),
        [_row(date=dt.date(2026, 1, 11))], "upcoming",
    )
    # Same fixture: existing blank identity cells get filled, no append.
    assert not plan.appends
    assert plan.changes


def test_one_ambiguous_legacy_row_cannot_be_claimed_by_two_fixtures():
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "11/01/2026"
    with pytest.raises(ContractError, match="claimed by both"):
        plan_update(
            "UpcomingClean", _ws([existing]), _contract(),
            [_row(date=dt.date(2026, 1, 11)), _row(date=dt.date(2026, 11, 1))],
            "upcoming",
        )


def test_alias_identity_matches_without_duplicate_or_manual_line_change():
    existing = [""] * 25
    existing[COL["Home"] - 1] = "LA Lakers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    existing[COL["Q1"] - 1] = "58.5"
    plan = plan_update(
        "UpcomingClean", _ws([existing]), _contract(),
        [_row(home="Los Angeles Lakers", **{"Q1": 58.5})], "upcoming",
    )
    assert not plan.appends
    assert COL["Q1"] not in {change.col for change in plan.changes}


def test_same_teams_different_dates_do_not_join():
    """Team-only matching across dates is forbidden."""
    existing = [""] * 25
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-20"
    ws = _ws([existing])
    plan = plan_update("UpcomingClean", ws, _contract(), [_row()], "upcoming")
    assert len(plan.appends) == 1  # new fixture appended, existing untouched
    assert not plan.changes


def test_apply_plan_requires_write_flag():
    plan = UpdatePlan(tab="UpcomingClean")
    with pytest.raises(PermissionError):
        apply_plan(plan, _ws(), allow_writes=False)


def test_apply_plan_batches_and_rate_limits():
    ws = _ws()
    plan = plan_update("UpcomingClean", ws, _contract(), [_row(), _row(home="A", away="B")], "upcoming")
    calls = []
    apply_plan(plan, ws, allow_writes=True, sleep=calls.append, rate_limit_delay=1.1)
    assert calls == [1.1]  # one batched append call, one rate-limit sleep


def test_rendered_key_dates_do_not_match_team_only():
    cells = _row()
    cells[COL["Date"]] = ""
    ws = _ws()
    plan = plan_update("UpcomingClean", ws, _contract(), [cells], "upcoming")
    assert len(plan.appends) == 1


def test_column_letter_helpers():
    assert column_letter(1) == "A"
    assert column_letter(19) == "S"
    assert column_letter(22) == "V"
    assert column_letter(27) == "AA"
