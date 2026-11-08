from __future__ import annotations

import json
from pathlib import Path

import pytest

from forecasting.operator import publish_operator_reports
from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind

DATE = "2026-11-07"
AS_OF = "2026-11-07T18:00:00Z"


class Store:
    def __init__(self, games):
        self.games = list(games)

    def list_games(self, league=None):
        return [game for game in self.games if game.league is league]


def game(number: int, *, home_score=None, away_score=None, total=None):
    if total is not None:
        home_score = total / 2
        away_score = total / 2
    scheduled = home_score is None and away_score is None
    actual = None
    if home_score is not None:
        actual = ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[home_score / 4] * 4,
            away_quarters=[away_score / 4] * 4,
            home_final=home_score,
            away_final=away_score,
        )
    return Game(
        game_id=f"game-{number}",
        league=League.NBA,
        season="2026-27",
        home_team_id=f"home-{number}",
        away_team_id=f"away-{number}",
        home_name=f"Home {number}",
        away_name=f"Away {number}",
        scheduled_utc=f"2026-11-{6 + number:02d}T20:00:00Z",
        status=GameStatus.SCHEDULED if scheduled else GameStatus.FINAL,
        actual=actual,
    )


def forecast(
    number: int,
    *,
    market="moneyline",
    selection="home",
    price=1.8,
    eligible=True,
    reasons=None,
    calibration_n=40,
):
    if reasons is None:
        reasons = [] if eligible else ["below_calibrated_probability_floor"]
    book = None
    if price is not None:
        book = {
            "provider": "test-books",
            "book": "Test Book",
            "observed_utc": "2026-11-07T17:00:00Z",
            "source_modified_utc": None,
            "pre_match": True,
            "main_line": True,
            "frozen": False,
            "selected_decimal_price": price,
            "selected_bookmaker_implied_probability": 1 / price,
            "bookmaker_bid_ask_probability_sum": None,
            "vig": None,
        }
    line = 200.0 if market == "total" else None
    return {
        "forecast_id": f"forecast-{number}-{market}-{selection}",
        "game_id": f"game-{number}",
        "league": "NBA",
        "scheduled_utc": f"2026-11-{6 + number:02d}T20:00:00Z",
        "home": f"Home {number}",
        "away": f"Away {number}",
        "market": market,
        "period": "full_game",
        "selection": selection,
        "line": line,
        "raw_probability": 0.60,
        "calibrated_probability": 0.62,
        "alternative_probabilities": {"away": 0.40} if market == "moneyline" else {"over": 0.40},
        "push_probability": 0.0 if market == "moneyline" else 0.02,
        "expected_margin": 4.0 if market == "moneyline" else None,
        "expected_total": None if market == "moneyline" else 200.0,
        "model": "provider-history",
        "model_version": "test-v1",
        "training_games": 30,
        "residual_games": 30,
        "book": book,
        "prediction": {
            "provider": "history",
            "observed_utc": AS_OF,
            "frozen": False,
            "model": "provider-history",
            "model_version": "test-v1",
        },
        "role": "basketball_native",
        "outcome": None,
        "calibration": {
            "status": "calibrated" if calibration_n >= 30 else "provisional",
            "n": calibration_n,
            "sufficient_for_accas": calibration_n >= 30,
        },
        "eligibility": {
            "acca_eligible": eligible,
            "reasons": reasons,
            "evidence_state": "current_pre_tip_unfrozen",
        },
    }


def acca(rows, *, acca_id="acca-original"):
    joint = 0.62 ** len(rows)
    odds = 1.8 ** len(rows)
    return {
        "acca_id": acca_id,
        "rank": 1,
        "legs": rows,
        "leg_count": len(rows),
        "estimated_joint_win_probability": joint,
        "combined_decimal_odds": odds,
        "bookmaker_implied_probability": 1 / odds,
        "estimated_ev_per_unit": joint * odds - 1,
    }


def report(rows, tickets=None):
    return {"forecasts": rows, "accas": {"results": tickets or []}}


def publish(tmp_path, store, market_report, run_id="run-1"):
    return publish_operator_reports(
        store,
        League.NBA,
        market_report,
        tmp_path,
        target_date=DATE,
        generated_utc=AS_OF,
        run_id=run_id,
    )


def test_ticket_is_frozen_empty_scans_do_not_cancel_and_known_leg_loss_settles(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    games = [game(number) for number in (1, 2, 3)]
    store = Store(games)
    first = publish(tmp_path, store, report(rows, [acca(rows)]))
    receipt_path = Path(first["ticket_receipt"])
    frozen_bytes = receipt_path.read_bytes()
    performance = json.loads(Path(first["performance"]).read_text())
    assert performance["paper_bank_units"] == 100.0
    assert performance["paper_committed_units"] == 1.0
    assert performance["pending_accas"] == 1
    assert "FROZEN AT 20:00 - FINAL" in Path(first["ticket_text"]).read_text()

    # A later empty scan and a different candidate set cannot rotate the slip.
    publish(tmp_path, store, report(rows), run_id="empty-scan")
    publish(
        tmp_path,
        store,
        report(rows, [acca(rows, acca_id="acca-rotated")]),
        run_id="later-scan",
    )
    assert receipt_path.read_bytes() == frozen_bytes
    assert json.loads(receipt_path.read_text())["decision_id"] == "run-1"

    # One known losing leg settles the whole acca even with two unresolved legs.
    games[0] = game(1, home_score=90, away_score=100)
    store.games = games
    publish(tmp_path, store, report(rows), run_id="grading-scan")
    performance = json.loads(Path(first["performance"]).read_text())
    assert performance["losses"] == 1
    assert performance["pending_accas"] == 0
    assert performance["unresolved_result_legs"] == 2
    assert performance["profit_units"] == -1.0
    assert performance["paper_bank_units"] == 99.0
    ticket_text = Path(first["ticket_text"]).read_text()
    assert "[PAPER ACCA #1]" in ticket_text
    assert "LOSS" in ticket_text
    assert receipt_path.read_bytes() == frozen_bytes

    # Missing immutable financial evidence fails closed; it never triggers a rebuild.
    receipt_path.unlink()
    with pytest.raises(FileNotFoundError, match="refusing to rotate"):
        publish(tmp_path, store, report(rows), run_id="missing-receipt")


def test_push_leg_is_removed_from_effective_acca_price(tmp_path):
    rows = [
        forecast(1, price=2.0),
        forecast(
            2,
            market="total",
            selection="under",
            price=1.5,
            reasons=["push_risk_requires_leg_void_rules"],
            eligible=False,
        ),
    ]
    ticket = acca(rows)
    ticket["combined_decimal_odds"] = 3.0
    games = [
        game(1, home_score=110, away_score=100),
        game(2, total=200),
    ]
    first = publish(tmp_path, Store(games), report(rows, [ticket]))
    performance = json.loads(Path(first["performance"]).read_text())
    settled = performance["history"][0]["accas"][0]
    assert settled["leg_outcomes"] == ["win", "push"]
    assert settled["grade_status"] == "win"
    assert settled["effective_decimal_odds"] == 2.0
    assert settled["profit_units"] == pytest.approx(1.0)
    picks = json.loads(Path(first["dated_picks"]).read_text())
    assert [row["grade_status"] for row in picks] == ["win", "push"]


def test_all_push_acca_returns_stake_without_a_win_loss_decision(tmp_path):
    rows = [
        forecast(1, market="total", selection="under", price=1.5, eligible=False),
        forecast(2, market="total", selection="under", price=1.5, eligible=False),
    ]
    ticket = acca(rows)
    ticket["combined_decimal_odds"] = 2.25
    result = publish(
        tmp_path,
        Store([game(1, total=200), game(2, total=200)]),
        report(rows, [ticket]),
    )
    performance = json.loads(Path(result["performance"]).read_text())
    assert performance["settled_accas"] == 1
    assert performance["pushes"] == 1
    assert performance["hit_rate"] is None
    assert performance["profit_units"] == 0.0
    assert performance["roi"] == 0.0
    assert performance["paper_bank_units"] == 100.0
    assert "no win/loss decisions yet" in Path(result["performance_text"]).read_text()


def test_pick_buckets_render_latest_logical_rows_and_audit_excludes_same_day(tmp_path):
    rows = [
        forecast(1),
        forecast(2, eligible=False, reasons=["insufficient_walk_forward_samples"], calibration_n=4),
        forecast(
            3,
            price=None,
            eligible=False,
            reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
            calibration_n=0,
        ),
        forecast(4, eligible=False),
    ]
    result = publish(tmp_path, Store([game(number) for number in range(1, 5)]), report(rows))
    text = Path(result["picks_text"]).read_text()
    assert "ELIGIBLE CLEAN PICKS" in text
    assert "CAUTION PICKS" in text
    assert "WATCHLIST — NO MATCHED ODDS" in text
    assert "SKIPPED — INELIGIBLE" in text
    assert "Active summary: ELIGIBLE_CLEAN=1" in text
    assert "Logical picks shown: 4 · superseded observations: 0" in text
    assert "Total archived pick evidence rows: 4" in text
    assert result["ticket_receipt"] is None
    ticket_text = Path(result["ticket_text"]).read_text()
    assert "NO PAPER BET" in ticket_text
    assert "CANDIDATE FUNNEL" in ticket_text
    assert "4 scan rows · 3 exact priced · 2 calibration-sufficient" in ticket_text
    assert "2 priced + calibrated · 1 eligible row · 1 eligible game" in ticket_text
    assert "Need 3 independent games for one acca" in ticket_text
    assert "2 insufficient_walk_forward_samples" in ticket_text
    assert result["candidate_funnel"] == {
        "scan_rows": 4,
        "exact_priced_rows": 3,
        "calibration_sufficient_rows": 2,
        "priced_and_calibrated_rows": 2,
        "eligible_rows": 1,
        "eligible_games": 1,
        "required_games": 3,
        "eligibility_blockers": {
            "below_calibrated_probability_floor": 1,
            "insufficient_walk_forward_samples": 2,
            "no_exact_bookmaker_market": 1,
        },
    }
    buckets = {row["bucket"] for row in json.loads(Path(result["dated_picks"]).read_text())}
    assert buckets == {"ELIGIBLE_CLEAN", "CAUTION", "WATCHLIST_NO_ODDS", "SKIPPED_INELIGIBLE"}
    audit = json.loads(Path(result["rolling_audit"]).read_text())
    assert audit["archived_pick_rows"] == 4
    assert audit["same_day_rows_excluded"] == 4
    assert audit["settled_picks"] == 0


def test_repeated_observations_render_and_grade_as_one_logical_pick(tmp_path):
    original = forecast(
        1,
        market="total",
        selection="under",
        price=1.8,
        eligible=False,
        reasons=["insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    original["prediction"]["observed_utc"] = "2026-11-07T17:00:00Z"
    first = publish(tmp_path, Store([game(1)]), report([original]), run_id="first")
    dated_path = Path(first["dated_picks"])

    later = json.loads(json.dumps(original))
    later["forecast_id"] = "forecast-1-total-under-later"
    later["book"]["observed_utc"] = "2026-11-07T17:30:00Z"
    later["book"]["selected_decimal_price"] = 1.95
    later["prediction"]["observed_utc"] = "2026-11-07T17:30:00Z"
    publish(tmp_path, Store([game(1)]), report([later]), run_id="second")

    archived = json.loads(dated_path.read_text())
    assert len(archived) == 2
    assert {row["forecast_id"] for row in archived} == {
        "forecast-1-total-under",
        "forecast-1-total-under-later",
    }
    text = Path(first["picks_text"]).read_text()
    assert text.count("TOTAL full_game 200 -> UNDER") == 1
    assert "@1.950 Test Book" in text
    assert "@1.800 Test Book" not in text
    assert "archived_revisions=2" in text
    assert "latest_evidence_utc=2026-11-07T17:30:00Z" in text
    assert "Logical picks shown: 1 · superseded observations: 1" in text

    # Once settled, repeated scans must not become extra paper wagers in the
    # rolling audit. The latest pre-tip price is used once; dated JSON keeps both.
    publish_operator_reports(
        Store([game(1, total=180)]),
        League.NBA,
        report([]),
        tmp_path,
        target_date="2026-11-08",
        generated_utc="2026-11-08T01:00:00Z",
        run_id="results-next-day",
    )
    audit = json.loads(Path(first["rolling_audit"]).read_text())
    assert audit["archived_pick_rows"] == 1
    assert audit["archived_evidence_rows"] == 2
    assert audit["superseded_observations"] == 1
    assert audit["eligible_prior_picks"] == 1
    assert audit["settled_picks"] == 1
    assert audit["wins"] == 1
    assert audit["priced_sample"] == 1
    assert audit["profit_units"] == pytest.approx(0.95)
    assert len(audit["ledger"]) == 1
    assert audit["ledger"][0]["forecast_id"] == "forecast-1-total-under-later"
    assert len(json.loads(dated_path.read_text())) == 2


def test_started_and_settled_games_leave_the_active_pick_sections(tmp_path):
    first = publish(tmp_path, Store([game(1)]), report([forecast(1)]))
    text_path = Path(first["picks_text"])

    publish_operator_reports(
        Store([game(1)]),
        League.NBA,
        report([]),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T21:00:00Z",
        run_id="started-no-final",
    )
    text = text_path.read_text()
    started_heading = "IN PROGRESS / AWAITING OFFICIAL RESULT — ARCHIVE ONLY"
    assert "Home 1 vs Away 1" not in text.split(started_heading)[0]
    assert "Home 1 vs Away 1" in text.split(started_heading)[1]
    assert "SETTLED=0" in text

    publish_operator_reports(
        Store([game(1, home_score=105, away_score=100)]),
        League.NBA,
        report([]),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T22:00:00Z",
        run_id="official-final",
    )
    text = text_path.read_text()
    settled_heading = "SETTLED PICKS — ARCHIVE ONLY"
    assert "Home 1 vs Away 1" not in text.split(settled_heading)[0]
    assert "Home 1 vs Away 1" in text.split(settled_heading)[1]
    assert "grade=WIN" in text.split(settled_heading)[1]
    assert "IN_PROGRESS_OR_AWAITING_RESULT=0 SETTLED=1" in text


def test_prior_dated_picks_are_graded_without_changing_frozen_evidence(tmp_path):
    row = forecast(1)
    first = publish(tmp_path, Store([game(1)]), report([row]))
    dated_path = Path(first["dated_picks"])
    original = json.loads(dated_path.read_text())[0]
    assert original["grade_status"] == "pending"
    frozen_probability = original["calibrated_probability"]
    frozen_price = original["book"]["selected_decimal_price"]

    publish(
        tmp_path,
        Store([game(1, home_score=105, away_score=100)]),
        report([]),
        run_id="results-only",
    )
    graded = json.loads(dated_path.read_text())[0]
    assert graded["grade_status"] == "win"
    assert graded["calibrated_probability"] == frozen_probability
    assert graded["book"]["selected_decimal_price"] == frozen_price
    audit = json.loads(Path(first["rolling_audit"]).read_text())
    assert audit["same_day_rows_excluded"] == 0
    assert audit["same_day_pending_picks"] == 0
    assert audit["same_day_settled_picks"] == 1
    assert audit["eligible_prior_picks"] == 0
    assert audit["settled_picks"] == 1
    assert audit["wins"] == 1
    assert audit["priced_sample"] == 1
    assert audit["profit_units"] == pytest.approx(0.8)
    assert len(audit["settled_ledger"]) == 1
    audit_text = Path(first["dated_audit"]).read_text()
    assert "## Settled picks — official-final audit" in audit_text
    assert "### 2026-11-07: Home 1 vs Away 1 (Final: **105–100**)" in audit_text
    assert "🟢 **WON — Full-game moneyline: Home 1 (home)**" in audit_text
    assert "Expected: **62.0%**" in audit_text
    assert "Price: **1.800 · Test Book**" in audit_text
    assert "Calibration: **n=40**" in audit_text
    assert "p=62.0% (raw 60.0%)" not in audit_text


def test_audit_uses_local_day_cutoff_for_previous_slate_pending_rows(tmp_path):
    rows = [forecast(1), forecast(2)]
    first = publish(
        tmp_path,
        Store([game(1, home_score=105, away_score=100), game(2)]),
        report(rows),
    )
    same_day = json.loads(Path(first["rolling_audit"]).read_text())
    assert same_day["same_day_cutoff"] == DATE
    assert same_day["include_same_day_settled"] is True
    assert same_day["eligible_prior_picks"] == 0
    assert same_day["settled_picks"] == 1
    assert same_day["same_day_settled_picks"] == 1
    assert same_day["same_day_pending_picks"] == 1
    assert same_day["same_day_rows_excluded"] == 1
    assert same_day["pending_picks"] == 0
    assert same_day["pending_picks_total"] == 1

    publish_operator_reports(
        Store([game(1, home_score=105, away_score=100), game(2)]),
        League.NBA,
        report([]),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-08T03:35:03Z",
        run_id="next-local-day-previous-slate",
    )
    next_day = json.loads(Path(first["rolling_audit"]).read_text())
    assert next_day["same_day_cutoff"] == "2026-11-08"
    assert next_day["eligible_prior_picks"] == 2
    assert next_day["settled_picks"] == 1
    assert next_day["same_day_settled_picks"] == 0
    assert next_day["same_day_pending_picks"] == 0
    assert next_day["same_day_rows_excluded"] == 0
    assert next_day["pending_picks"] == 1
    assert next_day["pending_picks_total"] == 1
    assert len(next_day["settled_ledger"]) == 1
    assert len(json.loads(Path(first["dated_picks"]).read_text())) == 2
    audit_text = Path(first["dated_audit"]).read_text()
    assert "🟢 **WON — Full-game moneyline: Home 1 (home)**" in audit_text
    settled_section = audit_text.split("## Settled picks — official-final audit", 1)[1]
    assert "Home 2 vs Away 2" not in settled_section


def test_audit_groups_markets_by_fixture_with_actual_period_scores(tmp_path):
    moneyline = forecast(1)
    h1 = forecast(
        1,
        selection="away",
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    h1.update(
        {
            "forecast_id": "forecast-1-moneyline-away-H1",
            "period": "H1",
            "model": "team_history_empirical_v1",
            "raw_probability": 0.28,
            "calibrated_probability": 0.31,
        }
    )
    total = forecast(
        1,
        market="total",
        selection="under",
        price=1.909,
        eligible=False,
        reasons=["insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    total.update(
        {
            "line": 200.0,
            "expected_total": 198.0,
            "raw_probability": 0.53,
            "calibrated_probability": 0.53,
        }
    )

    first = publish(
        tmp_path,
        Store([game(1, home_score=105, away_score=100)]),
        report([moneyline, h1, total]),
    )
    audit = json.loads(Path(first["rolling_audit"]).read_text())
    assert audit["settled_picks"] == 3
    assert audit["wins"] == 1
    assert audit["losses"] == 2

    text = Path(first["dated_audit"]).read_text()
    assert text.count("### 2026-11-07: Home 1 vs Away 1") == 1
    assert "(Final: **105–100**)" in text
    assert "🟢 **WON — Full-game moneyline: Home 1 (home)**" in text
    assert "🔴 **LOST — First-half moneyline: Away 1 (away)**" in text
    assert "Actual First-half score: **52.5–50**" in text
    assert "31.0%** calibrated (raw 28.0%)" in text
    assert "🔴 **LOST — Full-game total 200: Under**" in text
    assert "Actual Full-game total: **205**" in text
    assert "Price: **1.909 · Test Book**" in text
    assert "Price: **No exact bookmaker price**" in text
