from __future__ import annotations

import json
from pathlib import Path

import pytest

from forecasting.operator import (
    _audit_market_selection,
    _forecast_from_row,
    _load_state,
    _pick_row,
    _render_empty_ticket_scan,
    load_settled_pick_history,
    _selection_label,
    opposite_quote,
    publish_operator_reports,
)
from forecasting.types import BookEvidence, Forecast
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
)

DATE = "2026-11-07"
AS_OF = "2026-11-07T18:00:00Z"
# NBA dated cards freeze at 18:00 SAST (16:00Z). Fixture-level draft evidence
# remains independently frozen at each fixture's 60-minute pre-tip point.
TICKET_DECISION_AS_OF = "2026-11-07T16:00:00Z"
FIXTURE_FREEZE_AS_OF = "2026-11-07T19:00:00Z"


class Store:
    def __init__(self, games):
        self.games = list(games)

    def list_games(self, league=None):
        return [game for game in self.games if game.league is league]


def game(
    number: int,
    *,
    league=League.NBA,
    home_score=None,
    away_score=None,
    total=None,
    game_id=None,
    scheduled_utc=None,
    status=None,
):
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
        game_id=game_id or f"game-{number}",
        league=league,
        season="2026-27",
        home_team_id=f"home-{number}",
        away_team_id=f"away-{number}",
        home_name=f"Home {number}",
        away_name=f"Away {number}",
        scheduled_utc=scheduled_utc or f"2026-11-{6 + number:02d}T20:00:00Z",
        status=status or (GameStatus.SCHEDULED if scheduled else GameStatus.FINAL),
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
        "model_version": "test-model",
        "training_games": 30,
        "residual_games": 30,
        "book": book,
        "prediction": {
            "provider": "history",
            "observed_utc": AS_OF,
            "frozen": False,
            "model": "provider-history",
            "model_version": "test-model",
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


def with_quote_evidence(row, side_prices, **book_overrides):
    selected_price = (row.get("book") or {}).get("selected_decimal_price")
    book = {
        "provider": "espn",
        "book": "Test Book",
        "observed_utc": "2026-11-07T17:00:00Z",
        "source_modified_utc": None,
        "pre_match": True,
        "main_line": True,
        "frozen": False,
        "game_id": row["game_id"],
        "market": row["market"],
        "period": row["period"],
        "line": row["line"],
        "selected_decimal_price": selected_price,
        "side_prices": side_prices,
    }
    book.update(book_overrides)
    return {**row, "book": book}


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


def publish(
    tmp_path,
    store,
    market_report,
    run_id="run-1",
    *,
    target_date=DATE,
    generated_utc=TICKET_DECISION_AS_OF,
):
    return publish_operator_reports(
        store,
        League.NBA,
        market_report,
        tmp_path,
        target_date=target_date,
        generated_utc=generated_utc,
        run_id=run_id,
    )


def seed_historical_success_picks(tmp_path, market, *, wins, losses):
    """Persist settled, no-price history to exercise the win-rate gate."""
    rows = []
    games = []
    selection = "home" if market == "moneyline" else "under"
    for index in range(wins + losses):
        outcome = "win" if index < wins else "loss"
        game_id = f"history-{market}-{index}"
        scheduled_utc = f"2026-10-30T{index:02d}:00:00Z"
        forecast_row = forecast(
            1000 + index,
            market=market,
            selection=selection,
            price=None,
            eligible=False,
            reasons=["no_exact_bookmaker_market"],
            calibration_n=40,
        )
        forecast_row["game_id"] = game_id
        forecast_row["forecast_id"] = f"history-forecast-{market}-{index}"
        forecast_row["scheduled_utc"] = scheduled_utc
        forecast_row["prediction"]["observed_utc"] = "2026-10-29T12:00:00Z"
        history_row = _pick_row(forecast_row, "2026-10-30")
        rows.append(history_row)
        if index == 0:
            # A second book observation is still one historical wager decision.
            duplicate = {
                **history_row,
                "forecast_id": f"history-forecast-{market}-{index}-other-book",
                "pick_id": f"history-forecast-{market}-{index}-other-book",
                "book": {
                    "provider": "history-books",
                    "book": "Second Book",
                    "observed_utc": "2026-10-29T11:00:00Z",
                    "source_modified_utc": None,
                    "pre_match": True,
                    "main_line": True,
                    "frozen": False,
                    "game_id": game_id,
                    "market": market,
                    "period": "full_game",
                    "line": 200.0 if market == "total" else None,
                    "selected_decimal_price": None,
                    "side_prices": {selection: None},
                },
            }
            rows.append(duplicate)

        if market == "moneyline":
            home_score, away_score = (110, 100) if outcome == "win" else (100, 110)
        else:
            home_score, away_score = (90, 95) if outcome == "win" else (105, 110)
        games.append(
            game(
                1000 + index,
                home_score=home_score,
                away_score=away_score,
                game_id=game_id,
                scheduled_utc=scheduled_utc,
            )
        )

    ledger = tmp_path / "picks_2026-10-30.json"
    existing = json.loads(ledger.read_text()) if ledger.exists() else []
    ledger.write_text(json.dumps([*existing, *rows], indent=2) + "\n")
    return games


def test_load_state_compacts_duplicate_scan_rows_but_preserves_ticket_locks(tmp_path):
    state_path = tmp_path / "auto_tickets_state.json"
    lock = {"decision_id": "immutable-id", "withdrawn": False}
    state_path.write_text(
        json.dumps(
            {
                "schema_version": "magolide.paper_ticket_state",
                "frozen_by_date": {DATE: lock},
                "scans": [
                    {
                        "unpriced_audit": {
                            "count": 1,
                            "candidate_records": [{"candidate_id": "large-duplicate"}],
                            "forecast_ids": ["large-duplicate"],
                        },
                        "candidate_funnel": {
                            "scan_rows": 1,
                            "paper_pick_filter": {
                                "candidate_pick_count": 1,
                                "candidate_band_details": {"large-duplicate": {"x": 1}},
                                "bands": [{"x": 1}],
                            },
                        },
                    }
                ],
            }
        )
    )

    state = _load_state(tmp_path)

    assert state["frozen_by_date"][DATE] == lock
    assert state["scans"][0]["unpriced_audit"] == {"count": 1}
    assert state["scans"][0]["candidate_funnel"]["scan_rows"] == 1
    assert state["scans"][0]["candidate_funnel"]["paper_pick_filter"] == {
        "candidate_pick_count": 1
    }


def test_book_side_prices_survive_forecast_serialization_and_row_rehydration():
    snapshot = BookSnapshot(
        game_id="game-1",
        provider="espn",
        book="Test Book",
        market=Market.SPREAD,
        period=Period.FULL_GAME,
        observed_utc="2026-11-07T17:00:00Z",
        line=4.5,
        home_price_decimal=None,
        away_price_decimal=2.25,
    )
    forecast_row = Forecast(
        game_id="game-1",
        league="NBA",
        scheduled_utc="2026-11-07T20:00:00Z",
        home_name="Home 1",
        away_name="Away 1",
        market=Market.SPREAD.value,
        period=Period.FULL_GAME.value,
        selection="home",
        raw_probability=0.6,
        alternative_probabilities={"home": 0.6, "away": 0.4},
        source="test-model",
        line=4.5,
        calibration_version="model-config-a",
        prediction_provider="provider-a",
        prediction_observed_utc="2026-11-07T17:00:00Z",
        prediction_model="bpi-model",
        prediction_model_version="bpi-42",
        prediction_produced_by="ESPN BPI",
        shadow_calibration={
            "challenger_version": "recency-test",
            "status": "provisional",
            "n": 12,
            "weighted_n": 8.5,
            "ess": 10.0,
            "adjustment": 0.01,
            "champion_probability": 0.61,
            "champion_status": "provisional",
            "champion_n": 5,
            "champion_adjustment": -0.03,
            "probability": 0.62,
        },
        book=BookEvidence.from_snapshot(snapshot, frozen=False),
    ).to_dict()

    serialized_book = forecast_row["book"]
    assert serialized_book["selected_decimal_price"] is None
    assert serialized_book["side_prices"] == {
        "home": None,
        "away": 2.25,
        "draw": None,
        "over": None,
        "under": None,
    }
    assert serialized_book["game_id"] == "game-1"
    assert serialized_book["market"] == "spread"
    assert serialized_book["period"] == "full_game"
    assert serialized_book["line"] == 4.5

    rehydrated = _forecast_from_row(forecast_row)
    assert rehydrated.calibration_version == "model-config-a"
    assert rehydrated.forecast_id == forecast_row["forecast_id"]
    assert rehydrated.prediction_model == "bpi-model"
    assert rehydrated.prediction_model_version == "bpi-42"
    assert rehydrated.prediction_produced_by == "ESPN BPI"
    assert rehydrated.shadow_calibration["challenger_version"] == "recency-test"
    assert rehydrated.shadow_calibration["ess"] == 10.0
    assert rehydrated.shadow_calibration["champion_probability"] == 0.61
    assert rehydrated.shadow_calibration == {
        "challenger_version": "recency-test",
        "status": "provisional",
        "n": 12,
        "weighted_n": 8.5,
        "ess": 10.0,
        "adjustment": 0.01,
        "probability": 0.62,
        "champion_probability": 0.61,
        "champion_status": "provisional",
        "champion_n": 5,
        "champion_adjustment": -0.03,
    }
    assert rehydrated.book is not None
    assert rehydrated.book.prices == serialized_book["side_prices"]
    assert rehydrated.decimal_price is None
    assert rehydrated.opposite_decimal_price == 2.25
    assert opposite_quote(forecast_row) == (2.25, None)


@pytest.mark.parametrize(
    ("scope_key", "scope_value"),
    [
        ("game_id", "another-game"),
        ("market", "total"),
        ("period", "H1"),
        ("line", 5.5),
    ],
)
def test_opposite_quote_rejects_mismatched_quote_scope(scope_key, scope_value):
    row = forecast(1, market="spread", price=None)
    row["line"] = 4.5
    quote = with_quote_evidence(
        row,
        {"home": None, "away": 2.2, "draw": None, "over": None, "under": None},
        **{scope_key: scope_value},
    )
    assert opposite_quote(quote) == (None, "opposite_quote_scope_mismatch")


def test_opposite_quote_does_not_reuse_selected_price_or_accept_invalid_context():
    inconsistent = with_quote_evidence(
        forecast(1, price=1.8),
        {"home": 1.75, "away": 2.2, "draw": None, "over": None, "under": None},
    )
    assert opposite_quote(inconsistent) == (None, "selected_quote_inconsistent")

    one_sided = with_quote_evidence(
        forecast(1, price=1.8),
        {"home": 1.8, "away": None, "draw": None, "over": None, "under": None},
    )
    assert opposite_quote(one_sided) == (None, "opposite_quote_missing_in_payload")

    no_total_line = forecast(1, market="total", selection="over", price=None)
    no_total_line["line"] = None
    no_total_line = with_quote_evidence(
        no_total_line,
        {"home": None, "away": None, "draw": None, "over": 1.8, "under": 2.2},
    )
    assert opposite_quote(no_total_line) == (None, "market_line_missing")

    post_tip = with_quote_evidence(
        forecast(1, price=None),
        {"home": None, "away": 2.2, "draw": None, "over": None, "under": None},
        observed_utc="2026-11-07T20:00:00Z",
    )
    assert opposite_quote(post_tip) == (None, "not_pre_tip")


def test_opposite_quote_excludes_non_full_game_or_three_way_moneylines():
    first_half = forecast(1, price=None)
    first_half["period"] = "H1"
    first_half = with_quote_evidence(
        first_half,
        {"home": 1.8, "away": 2.2, "draw": None, "over": None, "under": None},
    )
    assert opposite_quote(first_half) == (None, "market_has_no_single_opposite_side")

    three_way = with_quote_evidence(
        forecast(1, price=None),
        {"home": 1.8, "away": 2.2, "draw": 25.0, "over": None, "under": None},
    )
    assert opposite_quote(three_way) == (None, "market_has_no_single_opposite_side")


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
    assert performance["awaiting_game_legs"] == 3
    assert performance["result_unresolvable_legs"] == 0
    assert performance["unresolved_result_legs"] == 3
    day = performance["history"][0]
    position = day["positions"][0]
    assert day["awaiting_game_legs"] == 3
    assert day["result_unresolvable_legs"] == 0
    assert day["unresolved_result_legs"] == 3
    assert position["awaiting_game_legs"] == 3
    assert position["result_unresolvable_legs"] == 0
    assert all(
        leg["result_resolution_status"] == "awaiting_game"
        for leg in position["legs"]
    )
    performance_text = Path(first["performance_text"]).read_text()
    pending_day_line = next(
        line for line in performance_text.splitlines() if line.startswith("  2026-11-07")
    )
    assert "pending 1.000" in pending_day_line
    assert not any(icon in pending_day_line for icon in ("✅", "❌", "➖"))
    ticket_text = Path(first["ticket_text"]).read_text()
    assert "DECISION FROZEN AT 18:00" in ticket_text
    assert "displayed grades refresh when reconciled results arrive" in ticket_text
    assert "TIP-OFF" in ticket_text
    assert "KICKOFF" not in ticket_text

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
    assert performance["awaiting_game_legs"] == 2
    assert performance["result_unresolvable_legs"] == 0
    assert performance["unresolved_result_legs"] == 2
    assert performance["profit_units"] == -1.0
    assert performance["paper_bank_units"] == 99.0
    ticket_text = Path(first["ticket_text"]).read_text()
    assert "[PAPER 3-LEG ACCA #1]" in ticket_text
    assert "LOSS" in ticket_text
    assert receipt_path.read_bytes() == frozen_bytes

    # Missing immutable financial evidence fails closed; it never triggers a rebuild.
    receipt_path.unlink()
    with pytest.raises(FileNotFoundError, match="refusing to rotate"):
        publish(tmp_path, store, report(rows), run_id="missing-receipt")


def test_ticket_waits_for_the_pre_tip_decision_window_and_uses_later_price(tmp_path):
    early_rows = [forecast(number) for number in (1, 2, 3)]
    early = publish(
        tmp_path,
        Store([game(number) for number in (1, 2, 3)]),
        report(early_rows, [acca(early_rows)]),
        run_id="before-decision-window",
        generated_utc="2026-11-07T15:30:00Z",
    )

    assert early["ticket_receipt"] is None
    assert early["ticket_decision"]["status"] == "pending_cutoff"
    assert early["ticket_decision"]["policy"] == "1800_sast_or_first_slate_tip_minus_60_minutes"
    assert early["ticket_decision"]["publication_local_time"] == "18:00"
    assert early["ticket_decision"]["planned_cutoff_utc"] == "2026-11-07T16:00:00Z"
    assert json.loads(Path(early["ticket_state"]).read_text())["frozen_by_date"] == {}
    assert "Ticket decision pending" in Path(early["ticket_text"]).read_text()

    later_rows = [forecast(number, price=1.95) for number in (1, 2, 3)]
    frozen = publish(
        tmp_path,
        Store([game(number) for number in (1, 2, 3)]),
        report(later_rows, [acca(later_rows, acca_id="later-priced-candidate")]),
        run_id="at-decision-window",
        generated_utc=TICKET_DECISION_AS_OF,
    )

    receipt = json.loads(Path(frozen["ticket_receipt"]).read_text())
    assert receipt["decision_id"] == "at-decision-window"
    assert receipt["ticket_decision"]["status"] == "due"
    assert receipt["ticket_decision"]["candidate_tickets_pre_tip"] == 1
    assert receipt["accas"][0]["legs"][0]["book"]["selected_decimal_price"] == 1.95


def _named_price_draft_rows():
    rows = []
    for number in (1, 2, 3):
        row = forecast(
            number,
            eligible=False,
            reasons=["insufficient_walk_forward_samples"],
            calibration_n=8,
        )
        row = with_quote_evidence(
            row,
            {"home": 1.8, "away": 2.0, "draw": None, "over": None, "under": None},
            price_provenance="named_book",
            source_permission="unresolved",
            operational_validation="unverified",
            ticket_price_eligible=True,
        )
        rows.append(row)
    return rows


def test_pre_cutoff_dated_card_is_the_only_named_price_draft_surface(tmp_path):
    rows = _named_price_draft_rows()

    result = publish(
        tmp_path,
        Store([game(number) for number in (1, 2, 3)]),
        report(rows),
        generated_utc="2026-11-07T06:30:00Z",
        run_id="morning-draft",
    )

    card = Path(result["ticket_text"]).read_text()
    assert "MA GOLIDE PAPER TICKETS — NBA 2026-11-07" in card
    assert "PRE-CUTOFF NAMED-PRICE DRAFT SLIPS — DRAFT ONLY / PRE-CUTOFF" in card
    assert "DRAFT ONLY: no paper stake, ticket receipt, settlement grade" in card
    assert "[DRAFT 2-LEG SLIP #1]" in card
    assert "[NAMED-BOOK PRICE]" in card
    assert "insufficient_walk_forward_samples" in card
    assert "price_source_permission_unresolved" not in card
    assert "price_operational_validation_incomplete" not in card
    assert "ticket_preview" not in result
    assert not list(tmp_path.glob("auto_tickets_*_preview.txt"))
    assert result["ticket_receipt"] is None
    assert json.loads(Path(result["ticket_state"]).read_text())["frozen_by_date"] == {}


def test_no_draft_content_is_promoted_to_a_terminal_no_bet_card(tmp_path):
    rows = _named_price_draft_rows()
    games = Store([game(number) for number in (1, 2, 3)])
    early = publish(
        tmp_path,
        games,
        report(rows),
        generated_utc="2026-11-07T06:30:00Z",
        run_id="pre-cutoff-draft",
    )
    assert "[NAMED-BOOK PRICE]" in Path(early["ticket_text"]).read_text()

    terminal = publish(
        tmp_path,
        games,
        report(rows),
        run_id="no-qualifying-ticket-at-cutoff",
        generated_utc=TICKET_DECISION_AS_OF,
    )

    card = Path(terminal["ticket_text"]).read_text()
    assert terminal["no_bet_decision"] is not None
    assert terminal["ticket_receipt"] is None
    assert "closed by an immutable no-bet decision" in card
    assert "PRE-CUTOFF NAMED-PRICE DRAFT SLIPS" not in card
    assert "[NAMED-BOOK PRICE]" not in card
    assert "[DRAFT" not in card
    assert not list(tmp_path.glob("auto_tickets_*_preview.txt"))


def test_strict_receipt_card_never_carries_draft_content(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    blocked = _named_price_draft_rows()
    games = Store([game(number) for number in (1, 2, 3)])
    early = publish(
        tmp_path,
        games,
        report(blocked),
        generated_utc="2026-11-07T06:30:00Z",
        run_id="pre-cutoff-draft",
    )
    assert "[NAMED-BOOK PRICE]" in Path(early["ticket_text"]).read_text()

    frozen = publish(
        tmp_path,
        games,
        report(rows, [acca(rows, acca_id="strict-at-cutoff")]),
        run_id="at-decision-window",
        generated_utc=TICKET_DECISION_AS_OF,
    )

    receipt = json.loads(Path(frozen["ticket_receipt"]).read_text())
    assert receipt["decision_id"] == "at-decision-window"
    card = Path(frozen["ticket_text"]).read_text()
    assert "PRE-CUTOFF NAMED-PRICE DRAFT SLIPS" not in card
    assert "[NAMED-BOOK PRICE]" not in card
    assert "[DRAFT" not in card
    assert not list(tmp_path.glob("auto_tickets_*_preview.txt"))


@pytest.mark.parametrize("league", list(League))
def test_fixture_pick_freeze_contract_is_global_for_every_supported_league(tmp_path, league):
    row = forecast(1)
    row["league"] = league.value
    fixture = game(1, league=league, scheduled_utc="2026-11-07T20:00:00Z")

    result = publish_operator_reports(
        Store([fixture]),
        league,
        report([row]),
        tmp_path / league.value.lower(),
        target_date=DATE,
        generated_utc=FIXTURE_FREEZE_AS_OF,
        run_id=f"{league.value.lower()}-fixture-freeze",
    )

    lifecycle_state = json.loads(Path(result["pick_lifecycle_state"]).read_text())
    freeze = lifecycle_state["frozen_by_date"][DATE][fixture.game_id]
    assert freeze["league"] == league.value
    assert freeze["game_id"] == fixture.game_id
    assert freeze["decision_cutoff_utc"] == FIXTURE_FREEZE_AS_OF
    assert freeze["rows"][0]["book"]["selected_decimal_price"] == 1.8
    assert not list((tmp_path / league.value.lower()).glob("pick_freeze_*.json"))
    assert result["pick_freezes"] == 1


def test_per_fixture_pick_lifecycle_publishes_drafts_freezes_once_and_settles(tmp_path):
    draft_rows = [forecast(1), forecast(2)]
    draft_rows[1]["scheduled_utc"] = "2026-11-07T22:00:00Z"
    games = Store(
        [
            game(1, scheduled_utc="2026-11-07T20:00:00Z"),
            game(2, scheduled_utc="2026-11-07T22:00:00Z"),
        ]
    )

    draft = publish(
        tmp_path,
        games,
        report(draft_rows),
        run_id="draft-before-cutoff",
        generated_utc="2026-11-07T18:00:00Z",
    )
    assert draft["pick_freezes"] == 0
    draft_lifecycle = json.loads(Path(draft["pick_lifecycle"]).read_text())
    assert {row["lifecycle_status"] for row in draft_lifecycle["rows"]} == {"DRAFT"}

    freeze_rows = [forecast(1, price=1.95), forecast(2)]
    freeze_rows[1]["scheduled_utc"] = "2026-11-07T22:00:00Z"
    frozen = publish(
        tmp_path,
        games,
        report(freeze_rows),
        run_id="fixture-one-cutoff",
        generated_utc="2026-11-07T19:00:00Z",
    )
    frozen_state = json.loads(Path(frozen["pick_lifecycle_state"]).read_text())
    frozen_evidence = frozen_state["frozen_by_date"][DATE]["game-1"]
    frozen_bytes = json.dumps(frozen_evidence, sort_keys=True).encode()
    assert not list(tmp_path.glob("pick_freeze_2026-11-07_*.json"))
    assert frozen_evidence["game_id"] == "game-1"
    assert frozen_evidence["rows"][0]["book"]["selected_decimal_price"] == 1.95
    assert frozen["pick_freezes"] == 1

    changed_rows = [forecast(1, price=2.10), forecast(2)]
    changed_rows[1]["scheduled_utc"] = "2026-11-07T22:00:00Z"
    publish(
        tmp_path,
        games,
        report(changed_rows),
        run_id="later-draft-must-not-rewrite-fixture-one",
        generated_utc="2026-11-07T19:30:00Z",
    )
    later_state = json.loads((tmp_path / "pick_lifecycle_state.json").read_text())
    assert json.dumps(
        later_state["frozen_by_date"][DATE]["game-1"], sort_keys=True
    ).encode() == frozen_bytes
    archived_after_freeze = json.loads((tmp_path / "picks_2026-11-07.json").read_text())
    fixture_one_prices = {
        row["book"]["selected_decimal_price"]
        for row in archived_after_freeze
        if row["game_id"] == "game-1"
    }
    assert fixture_one_prices == {1.95}
    lifecycle = json.loads((tmp_path / "picks_2026-11-07_lifecycle.json").read_text())
    frozen_row = next(row for row in lifecycle["rows"] if row["game_id"] == "game-1")
    assert frozen_row["book"]["selected_decimal_price"] == 1.95
    assert frozen_row["lifecycle_status"] == "FROZEN"

    games.games[0] = game(
        1,
        home_score=101,
        away_score=90,
        scheduled_utc="2026-11-07T20:00:00Z",
    )
    settled = publish(
        tmp_path,
        games,
        report(changed_rows),
        run_id="official-final-for-fixture-one",
        generated_utc="2026-11-07T20:30:00Z",
    )
    lifecycle = json.loads(Path(settled["pick_lifecycle"]).read_text())
    settled_row = next(row for row in lifecycle["rows"] if row["game_id"] == "game-1")
    draft_row = next(row for row in lifecycle["rows"] if row["game_id"] == "game-2")
    assert settled_row["grade_status"] == "win"
    assert settled_row["lifecycle_status"] == "SETTLED"
    assert draft_row["lifecycle_status"] == "DRAFT"
    assert "lifecycle=SETTLED" in Path(settled["picks_text"]).read_text()


def test_terminal_no_bet_at_cutoff_blocks_later_ticket_candidates(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    games = Store([game(number) for number in (1, 2, 3)])
    no_bet = publish(
        tmp_path,
        games,
        report(rows),
        run_id="no-qualifying-ticket-at-cutoff",
        generated_utc=TICKET_DECISION_AS_OF,
    )

    frozen_no_bet = no_bet["no_bet_decision"]
    assert no_bet["ticket_receipt"] is None
    assert frozen_no_bet["reason"] == "no_qualifying_pre_tip_ticket_at_decision_cutoff"
    state = json.loads(Path(no_bet["ticket_state"]).read_text())
    assert state["frozen_by_date"][DATE]["no_bet"] is True
    assert state["frozen_by_date"][DATE]["no_bet_decision"] == frozen_no_bet

    later = publish(
        tmp_path,
        games,
        report(rows, [acca(rows, acca_id="late-candidate-must-not-publish")]),
        run_id="later-price-cannot-reopen-date",
        generated_utc="2026-11-07T19:20:00Z",
    )

    assert later["ticket_receipt"] is None
    assert later["no_bet_decision"] == frozen_no_bet
    assert not list(tmp_path.glob("auto_tickets_*_no_bet.json"))
    assert "closed by an immutable no-bet decision" in Path(later["ticket_text"]).read_text()


def test_withdrawn_invalid_no_bet_cannot_be_recreated_from_later_candidates(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    games = Store([game(number) for number in (1, 2, 3)])
    first = publish(
        tmp_path,
        games,
        report(rows),
        generated_utc=TICKET_DECISION_AS_OF,
        run_id="invalid-no-bet-source",
    )
    decision = first["no_bet_decision"]
    withdrawal = {
        "decision_id": decision["decision_id"],
        "withdrawn": True,
        "target_date": DATE,
        "reason": "cross_league_source_identity_failure",
        "invalid_decision_type": "no_bet",
    }
    (tmp_path / f"auto_tickets_{DATE}_withdrawn.json").write_text(
        json.dumps(withdrawal, indent=2) + "\n"
    )
    state = json.loads(Path(first["ticket_state"]).read_text())
    state["frozen_by_date"][DATE] = {
        "decision_id": decision["decision_id"],
        "withdrawn": True,
        "final": True,
        "reason": withdrawal["reason"],
    }
    Path(first["ticket_state"]).write_text(json.dumps(state, indent=2) + "\n")

    later = publish(
        tmp_path,
        games,
        report(rows, [acca(rows)]),
        generated_utc="2026-11-07T07:20:00Z",
        run_id="later-candidate-must-not-recreate-invalid-no-bet",
    )

    assert later["ticket_receipt"] is None
    assert later["no_bet_decision"] is None
    assert not (tmp_path / f"auto_tickets_{DATE}_no_bet.json").exists()
    assert (tmp_path / f"auto_tickets_{DATE}_withdrawn.json").exists()


def test_terminal_no_bet_card_refreshes_reconciled_shadow_grades(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    store = Store([game(number) for number in (1, 2, 3, 4)] + history_games)
    no_bet = publish(
        tmp_path,
        store,
        report([*rows, shadow_pick]),
        run_id="no-bet-shadow-at-cutoff",
    )

    frozen_no_bet = no_bet["no_bet_decision"]
    ticket_text_path = Path(no_bet["ticket_text"])
    assert "SHADOW GROUP 1 — 1 audit-only pick (1 ⏳ PENDING)" in ticket_text_path.read_text()

    settled_games = [
        game(1, home_score=100, away_score=90),
        game(2, home_score=100, away_score=90),
        game(3, home_score=100, away_score=90),
        game(4, home_score=100, away_score=80),
        *history_games,
    ]
    later = publish(
        tmp_path,
        Store(settled_games),
        report([]),
        run_id="no-bet-shadow-result-refresh",
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
    )

    refreshed_card = ticket_text_path.read_text()
    assert later["ticket_receipt"] is None
    refreshed_state = json.loads((tmp_path / "auto_tickets_state.json").read_text())
    assert refreshed_state["frozen_by_date"][DATE]["no_bet_decision"] == frozen_no_bet
    assert not list(tmp_path.glob("auto_tickets_*_no_bet.json"))
    assert "closed by an immutable no-bet decision" in refreshed_card
    assert "SHADOW GROUP 1 — 1 audit-only pick (1 ✅ WIN)" in refreshed_card
    assert "[✅ WIN] Home 4 vs Away 4" in refreshed_card


def test_historical_unfrozen_scan_card_regrades_shadows_without_creating_a_decision(tmp_path):
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    first = publish(
        tmp_path,
        Store([game(4)] + history_games),
        report([shadow_pick]),
        generated_utc=AS_OF,
        run_id="historical-unfrozen-shadow",
    )
    card_path = Path(first["ticket_text"])
    ledger_path = Path(first["dated_picks"])
    original = json.loads(ledger_path.read_text())[0]
    assert first["ticket_receipt"] is None
    assert not (tmp_path / f"auto_tickets_{DATE}_no_bet.json").exists()
    assert "[⏳ PENDING] Home 4 vs Away 4" in card_path.read_text()

    later = publish(
        tmp_path,
        Store([game(4, home_score=100, away_score=80)] + history_games),
        report([]),
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
        run_id="historical-unfrozen-shadow-result-refresh",
    )

    refreshed = card_path.read_text()
    reconciled = json.loads(ledger_path.read_text())[0]
    assert later["historical_ticket_audit_cards_refreshed"] == [DATE]
    assert "HISTORICAL SCAN ARCHIVE" in refreshed
    assert "does not create, replace, or retrospectively freeze" in refreshed
    assert "[✅ WIN] Home 4 vs Away 4" in refreshed
    assert "⏳ PENDING" not in refreshed
    assert not (tmp_path / f"auto_tickets_{DATE}_decision.json").exists()
    assert not (tmp_path / f"auto_tickets_{DATE}_no_bet.json").exists()
    for key in ("pick_id", "forecast_id", "bucket", "scheduled_utc"):
        assert reconciled[key] == original[key]
    assert reconciled["grade_status"] == "win"


def test_historical_pending_cards_stay_byte_preserved_when_raw_ledger_is_unavailable(tmp_path):
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    first = publish(
        tmp_path,
        Store([game(4)] + history_games),
        report([shadow_pick]),
        generated_utc=AS_OF,
        run_id="historical-shadow-before-ledger-eviction",
    )
    pick_card = Path(first["picks_text"])
    ticket_card = Path(first["ticket_text"])
    pick_bytes = pick_card.read_bytes()
    ticket_bytes = ticket_card.read_bytes()
    Path(first["dated_picks"]).unlink()

    later = publish(
        tmp_path,
        Store([game(4, home_score=100, away_score=80)] + history_games),
        report([]),
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
        run_id="historical-shadow-without-retained-ledger",
    )

    assert later["historical_ticket_audit_cards_refreshed"] == []
    assert pick_card.read_bytes() == pick_bytes
    assert ticket_card.read_bytes() == ticket_bytes
    assert not (tmp_path / f"auto_tickets_{DATE}_decision.json").exists()
    assert not (tmp_path / f"auto_tickets_{DATE}_no_bet.json").exists()


def test_historical_ticket_audit_reconciliation_keeps_future_shadow_draft(tmp_path):
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    shadow_pick["scheduled_utc"] = "2026-11-20T20:00:00Z"
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    first = publish(
        tmp_path,
        Store([game(4, scheduled_utc="2026-11-20T20:00:00Z")] + history_games),
        report([shadow_pick]),
        generated_utc=AS_OF,
        run_id="historical-future-shadow",
    )
    card_path = Path(first["ticket_text"])

    later = publish(
        tmp_path,
        Store([game(4, scheduled_utc="2026-11-20T20:00:00Z")] + history_games),
        report([]),
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
        run_id="historical-future-shadow-refresh",
    )

    lifecycle = json.loads((tmp_path / f"picks_{DATE}_lifecycle.json").read_text())
    row = next(row for row in lifecycle["rows"] if row["game_id"] == "game-4")
    assert later["historical_ticket_audit_cards_refreshed"] == [DATE]
    assert row["grade_status"] == "pending"
    assert row["lifecycle_status"] == "DRAFT"
    assert "[⏳ PENDING] Home 4 vs Away 4" in card_path.read_text()
    assert not (tmp_path / f"auto_tickets_{DATE}_decision.json").exists()
    assert not (tmp_path / f"auto_tickets_{DATE}_no_bet.json").exists()


def test_persisted_performance_is_compact_and_points_to_immutable_receipt(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    rows[0]["team_history"] = {"large_debug_payload": "x" * 100_000}
    result = publish(
        tmp_path,
        Store([game(number) for number in (1, 2, 3)]),
        report(rows, [acca(rows)]),
    )

    stored = Path(result["performance"]).read_bytes()
    performance = json.loads(stored)
    day = performance["history"][0]
    leg = day["positions"][0]["legs"][0]
    assert "compact_settlement_summary" in performance["history_detail"]
    assert day["receipt_path"] == f"auto_tickets_{DATE}_decision.json"
    assert "accas" not in day and "singles" not in day
    assert "team_history" not in leg
    assert "prediction" not in leg
    assert len(stored) < 10_000


def test_later_scan_refreshes_historical_card_grades_without_rewriting_receipt(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    store = Store([game(number) for number in (1, 2, 3, 4)] + history_games)
    first = publish(
        tmp_path,
        store,
        report([*rows, shadow_pick], [acca(rows)]),
    )
    receipt_path = Path(first["ticket_receipt"])
    receipt_bytes = receipt_path.read_bytes()
    ticket_text_path = Path(first["ticket_text"])
    initial_card = ticket_text_path.read_text()
    assert "[PAPER 3-LEG ACCA #1]" in initial_card
    assert "UNPRICED AUDIT-ONLY SHADOW GROUPS" in initial_card
    assert "SHADOW GROUP 1 — 1 audit-only pick (1 ⏳ PENDING)" in initial_card

    settled_games = [
        game(1, home_score=90, away_score=100),
        game(2, home_score=110, away_score=90),
        game(3, home_score=100, away_score=90),
        game(4, home_score=100, away_score=80),
        *history_games,
    ]
    later = publish(
        tmp_path,
        Store(settled_games),
        report([]),
        run_id="next-day-result-refresh",
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
    )

    refreshed_card = ticket_text_path.read_text()
    assert "[PAPER 3-LEG ACCA #1] @5.832" in refreshed_card
    assert "· ❌ LOSS" in refreshed_card
    assert "Home 1 vs Away 1" in refreshed_card
    assert "· ELIGIBLE_CLEAN · ❌ LOSS" in refreshed_card
    assert "SHADOW GROUP 1 — 1 audit-only pick (1 ✅ WIN)" in refreshed_card
    assert "[✅ WIN] Home 4 vs Away 4" in refreshed_card
    assert receipt_path.read_bytes() == receipt_bytes

    performance = json.loads(Path(later["performance"]).read_text())
    assert performance["pending_positions"] == 0
    assert performance["unresolved_result_legs"] == 0
    assert performance["history"][0]["positions"][0]["grade_status"] == "loss"


def test_ncaam_historical_card_refreshes_priced_ineligible_shadow_grades(tmp_path):
    single = forecast(1)
    single["league"] = League.NCAA_M.value
    single["calibrated_probability"] = 0.82
    band = {
        "league": League.NCAA_M.value,
        "market": "moneyline",
        "period": "full_game",
        "model": "provider-history",
        "probability_band": "0.80-<0.90",
        "individual_odds_band": "1.50-<2.00",
        "candidate_pick_count": 1,
        "historical_pick_count": 14,
        "wins": 12,
        "losses": 2,
        "pushes": 0,
        "observed_roi": 0.16,
        "recency_weighted_roi": 0.16,
        "parent_roi": 0.08,
        "adjusted_roi": 0.11,
        "qualifies": True,
    }
    paper_filter = {
        "candidate_pick_count": 1,
        "candidate_band_count": 1,
        "qualifying_pick_count": 1,
        "qualifying_band_count": 1,
        "historical_unique_pick_count": 14,
        "candidate_band_details": {single["forecast_id"]: band},
        "bands": [band],
        "status": "applied",
    }
    shadow = forecast(
        2,
        price=1.8,
        eligible=False,
        reasons=["insufficient_team_specific_history"],
        calibration_n=10,
    )
    shadow["league"] = League.NCAA_M.value
    shadow["calibrated_probability"] = 0.85
    market_report = report([single, shadow])
    market_report["accas"].update(
        paper_results=[],
        paper_single_picks=[{**single, "individual_pick_roi_band": band}],
        paper_pick_filter=paper_filter,
    )

    first = publish_operator_reports(
        Store([game(1, league=League.NCAA_M), game(2, league=League.NCAA_M)]),
        League.NCAA_M,
        market_report,
        tmp_path,
        target_date=DATE,
        generated_utc=TICKET_DECISION_AS_OF,
        run_id="ncaa-single-and-shadow",
    )
    receipt_path = Path(first["ticket_receipt"])
    receipt_bytes = receipt_path.read_bytes()
    card_path = Path(first["ticket_text"])
    initial_card = card_path.read_text()
    assert "[PAPER SINGLE #1]" in initial_card
    assert "SHADOW GROUP 1 — 1 priced-ineligible audit-only pick (1 ⏳ PENDING)" in initial_card

    later = publish_operator_reports(
        Store(
            [
                game(1, league=League.NCAA_M, home_score=100, away_score=90),
                game(2, league=League.NCAA_M, home_score=100, away_score=90),
            ]
        ),
        League.NCAA_M,
        report([]),
        tmp_path,
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
        run_id="ncaa-results-refresh",
    )

    refreshed = card_path.read_text()
    assert "[PAPER SINGLE #1] @1.800" in refreshed
    assert "· ✅ WIN" in refreshed
    assert "SHADOW GROUP 1 — 1 priced-ineligible audit-only pick (1 ✅ WIN)" in refreshed
    assert "[✅ WIN] Home 2 vs Away 2" in refreshed
    assert receipt_path.read_bytes() == receipt_bytes
    assert json.loads(Path(later["performance"]).read_text())["pending_positions"] == 0


def test_withdrawn_ticket_survives_old_cache_without_becoming_a_live_position(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    store = Store([game(number) for number in (1, 2, 3)])
    first = publish(tmp_path, store, report(rows, [acca(rows)]))
    receipt_path = Path(first["ticket_receipt"])
    frozen_bytes = receipt_path.read_bytes()
    state_path = Path(first["ticket_state"])
    state = json.loads(state_path.read_text())
    state["frozen_by_date"][DATE].update(
        withdrawn=True,
        decision_id="run-1",
        withdrawal_reason="model_integrity",
    )
    state_path.write_text(json.dumps(state, indent=2) + "\n")

    # Even if an old cache restores the receipt and new candidates arrive, this
    # date stays closed and the withdrawn card contributes no stake or results.
    second = publish(
        tmp_path,
        store,
        report(rows, [acca(rows, acca_id="replacement")]),
        run_id="later-scan",
    )
    assert second["ticket_receipt"] is None
    assert receipt_path.read_bytes() == frozen_bytes
    performance = json.loads(Path(second["performance"]).read_text())
    assert performance["history"] == []
    assert performance["pending_accas"] == 0
    assert performance["paper_bank_units"] == 100.0
    text = Path(second["ticket_text"]).read_text()
    assert "withdrawn for model-integrity failure" in text
    assert "closed to replacement ticket publication" in text
    assert json.loads(state_path.read_text())["frozen_by_date"][DATE]["withdrawn"] is True


def test_historical_withdrawn_card_regrades_shadows_without_reviving_receipt(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    shadow_pick = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    first = publish(
        tmp_path,
        Store([*(game(number) for number in (1, 2, 3, 4)), *history_games]),
        report([*rows, shadow_pick], [acca(rows)]),
        run_id="withdrawn-shadow-receipt",
    )
    receipt_path = Path(first["ticket_receipt"])
    receipt_bytes = receipt_path.read_bytes()
    card_path = Path(first["ticket_text"])
    state_path = Path(first["ticket_state"])
    state = json.loads(state_path.read_text())
    state["frozen_by_date"][DATE].update(withdrawn=True, decision_id="withdrawn-shadow-receipt")
    state_path.write_text(json.dumps(state, indent=2) + "\n")
    # Materialise the Git-durable withdrawal marker while the original slate
    # is still the target date.
    publish(
        tmp_path,
        Store([*(game(number) for number in (1, 2, 3, 4)), *history_games]),
        report([]),
        run_id="materialise-withdrawal",
    )
    withdrawal_path = tmp_path / f"auto_tickets_{DATE}_withdrawn.json"
    assert withdrawal_path.exists()

    later = publish(
        tmp_path,
        Store(
            [
                *(game(number, home_score=100, away_score=90) for number in (1, 2, 3)),
                game(4, home_score=100, away_score=80),
                *history_games,
            ]
        ),
        report([]),
        target_date="2026-11-12",
        generated_utc="2026-11-12T18:00:00Z",
        run_id="withdrawn-shadow-result-refresh",
    )

    refreshed = card_path.read_text()
    assert later["historical_ticket_audit_cards_refreshed"] == [DATE]
    assert "withdrawn for model-integrity failure" in refreshed
    assert "HISTORICAL SCAN ARCHIVE" not in refreshed
    assert "[✅ WIN] Home 4 vs Away 4" in refreshed
    assert receipt_path.read_bytes() == receipt_bytes
    assert withdrawal_path.exists()
    assert json.loads(Path(later["performance"]).read_text())["history"] == []


def test_withdrawn_sidecar_written_and_blocks_adoption_without_state_lock(tmp_path):
    rows = [forecast(number) for number in (1, 2, 3)]
    store = Store([game(number) for number in (1, 2, 3)])
    first = publish(tmp_path, store, report(rows, [acca(rows)]))
    state_path = Path(first["ticket_state"])
    # Simulate manual withdrawal via state.json (as in previous test) — next run must materialize sidecar.
    state = json.loads(state_path.read_text())
    state["frozen_by_date"][DATE].update(withdrawn=True, decision_id="run-1")
    state_path.write_text(json.dumps(state, indent=2) + "\n")
    publish(tmp_path, store, report(rows, [acca(rows, acca_id="replacement")]), run_id="later-scan")
    sidecar = tmp_path / f"auto_tickets_{DATE}_withdrawn.json"
    assert sidecar.exists()
    assert json.loads(sidecar.read_text())["decision_id"] == "run-1"
    # State rollback + receipt still present — sidecar alone must keep date suppressed.
    state = json.loads(state_path.read_text())
    del state["frozen_by_date"][DATE]["withdrawn"]
    state_path.write_text(json.dumps(state, indent=2) + "\n")
    assert sidecar.exists()
    blocked = publish(
        tmp_path, store, report(rows, [acca(rows, acca_id="another")]), run_id="third-scan"
    )
    assert blocked["ticket_receipt"] is None
    assert "withdrawn for model-integrity failure" in Path(blocked["ticket_text"]).read_text()


def test_withdrawn_sidecar_drives_performance_exclusion_and_publish_skip(tmp_path):
    # Direct sidecar without state withdrawn flag — performance must still exclude.
    rows = [forecast(number) for number in (1, 2, 3)]
    store = Store([game(number) for number in (1, 2, 3)])
    first = publish(tmp_path, store, report(rows, [acca(rows)]))
    state_path = Path(first["ticket_state"])
    sidecar = tmp_path / f"auto_tickets_{DATE}_withdrawn.json"
    sidecar.write_text(
        json.dumps({"decision_id": "run-1", "withdrawn": True, "target_date": DATE}) + "\n"
    )
    # Remove withdrawn flag from state to prove sidecar drives exclusion.
    state = json.loads(state_path.read_text())
    if DATE in state["frozen_by_date"]:
        state["frozen_by_date"][DATE].pop("withdrawn", None)
    state_path.write_text(json.dumps(state, indent=2) + "\n")
    second = publish(
        tmp_path,
        store,
        report(rows, [acca(rows, acca_id="blocked-candidate")]),
        run_id="sidecar-only",
    )
    assert second["ticket_receipt"] is None
    assert "withdrawn for model-integrity failure" in Path(second["ticket_text"]).read_text()
    performance = json.loads(Path(second["performance"]).read_text())
    assert performance["history"] == []


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
    settled = performance["history"][0]["positions"][0]
    assert settled["leg_outcomes"] == ["win", "push"]
    assert settled["grade_status"] == "win"
    assert settled["effective_decimal_odds"] == 2.0
    assert settled["profit_units"] == pytest.approx(1.0)
    picks = json.loads(Path(first["dated_picks"]).read_text())
    assert [row["grade_status"] for row in picks] == ["win", "push"]


def test_acca_is_a_loss_as_soon_as_any_leg_loses_while_other_legs_are_pending(tmp_path):
    rows = [forecast(1), forecast(2)]
    result = publish(
        tmp_path,
        Store([game(1, home_score=90, away_score=100), game(2)]),
        report(rows, [acca(rows)]),
    )

    performance = json.loads(Path(result["performance"]).read_text())
    position = performance["history"][0]["positions"][0]
    assert position["leg_outcomes"] == ["loss", None]
    assert position["grade_status"] == "loss"
    assert position["unresolved_legs"] == 1
    assert position["profit_units"] == -position["stake_units"]
    assert performance["settled_accas"] == 1
    assert performance["pending_accas"] == 0
    assert "· ❌ LOSS" in Path(result["ticket_text"]).read_text()
    performance_text = Path(result["performance_text"]).read_text()
    record_line = next(line for line in performance_text.splitlines() if line.startswith("RECORD"))
    bet_day_line = next(
        line for line in performance_text.splitlines() if line.startswith("  2026-11-07")
    )
    assert "❌ LOSS 1L" in record_line
    assert "✅ WIN" not in record_line
    assert "➖ PUSH" not in record_line
    assert "❌ LOSS 1L" in bet_day_line
    assert "✅ WIN" not in bet_day_line
    assert "➖ PUSH" not in bet_day_line


@pytest.mark.parametrize(
    ("result_case", "expected_resolution"),
    [
        ("missing_game", "result_unresolvable"),
        ("unreconciled_final", "result_unresolvable"),
        (GameStatus.SCHEDULED, "awaiting_game"),
        (GameStatus.LIVE, "awaiting_game"),
        (GameStatus.POSTPONED, "awaiting_game"),
        (GameStatus.CANCELLED, "result_unresolvable"),
        (GameStatus.ABANDONED, "result_unresolvable"),
        (GameStatus.UNKNOWN, "result_unresolvable"),
    ],
)
def test_unresolvable_leg_results_are_distinguished_from_waiting_games(
    tmp_path, result_case, expected_resolution
):
    row = forecast(1)
    if result_case == "missing_game":
        store = Store([])
    elif result_case == "unreconciled_final":
        unreconciled = game(1, home_score=90, away_score=100)
        unreconciled.actual.home_quarters[0] += 1.0
        store = Store([unreconciled])
    else:
        store = Store([game(1, status=result_case)])

    result = publish(tmp_path, store, report([row], [acca([row])]))

    performance = json.loads(Path(result["performance"]).read_text())
    position = performance["history"][0]["positions"][0]
    leg = position["legs"][0]
    expected_awaiting = int(expected_resolution == "awaiting_game")
    expected_unresolvable = int(expected_resolution == "result_unresolvable")
    assert position["grade_status"] == "pending"
    assert position["unresolved_legs"] == 1
    assert position["awaiting_game_legs"] == expected_awaiting
    assert position["result_unresolvable_legs"] == expected_unresolvable
    assert leg["result_resolution_status"] == expected_resolution
    assert performance["awaiting_game_legs"] == expected_awaiting
    assert performance["result_unresolvable_legs"] == expected_unresolvable
    assert performance["unresolved_result_legs"] == 1
    assert (
        f"{expected_awaiting} awaiting games · "
        f"{expected_unresolvable} result unresolvable"
    ) in Path(result["performance_text"]).read_text()


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


def test_empty_ticket_card_summarizes_no_acca_without_printing_candidate_funnel():
    band = {
        "market": "moneyline",
        "period": "full_game",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "historical_pick_count": 12,
        "recency_weighted_roi": 0.12,
        "adjusted_roi": 0.08,
        "qualifies": True,
    }
    candidate_funnel = {
        "scan_rows": 50,
        "provisional_sample_only_rows": 7,
        "paper_candidate_rows": 7,
        "paper_single_picks": 2,
        "paper_single_games": 1,
        "paper_accas": 0,
        "strict_accas": 0,
        "eligibility_blockers": {"insufficient_walk_forward_samples": 7},
        "paper_pick_filter": {
            "candidate_pick_count": 7,
            "qualifying_pick_count": 2,
            "qualifying_band_count": 1,
            "historical_unique_pick_count": 12,
            "bands": [band],
            "status": "applied",
        },
    }
    text = _render_empty_ticket_scan(
        league=League.NBA,
        target_date=DATE,
        generated_utc=AS_OF,
        performance={
            "paper_bank_units": 100.0,
            "paper_free_units": 100.0,
            "paper_committed_units": 0.0,
        },
        timezone_name="UTC",
        candidate_funnel=candidate_funnel,
    )

    assert (
        "NO MULTI-LEG PAPER ACCA — 2 of 7 candidate picks passed individual-pick ROI; "
        "only 1 distinct game qualifies; 2 are required."
    ) in text
    assert text.index("NO MULTI-LEG PAPER ACCA") < text.index("CURRENT PAPER BANK")
    assert "CANDIDATE FUNNEL" not in text
    assert "50 scan rows" not in text
    assert "recency ROI" not in text
    assert "insufficient_walk_forward_samples" not in text


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
    assert (
        "NO PAPER BET — no ROI-qualified paper singles or multi-leg acca in this scan."
        in ticket_text
    )
    assert ticket_text.index("NO PAPER BET") < ticket_text.index("CURRENT PAPER BANK")
    assert (
        "UNPRICED AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY "
        "(1 audited; requires ≥70% historical success, n>10)"
    ) in ticket_text
    assert "forecast=forecast-3-moneyline-home" not in ticket_text
    assert "All candidates remain individually recorded in the audit ledger" in ticket_text
    assert "CANDIDATE FUNNEL" not in ticket_text
    assert "Individual-pick ROI filter" not in ticket_text
    assert result["candidate_funnel"] == {
        "scan_rows": 4,
        "exact_priced_rows": 3,
        "calibration_sufficient_rows": 2,
        "priced_and_calibrated_rows": 2,
        "eligible_rows": 1,
        "eligible_games": 1,
        "provisional_sample_only_rows": 1,
        "provisional_sample_only_games": 1,
        "paper_candidate_rows": 2,
        "paper_candidate_games": 2,
        "paper_single_picks": 0,
        "paper_single_games": 0,
        "paper_accas": 0,
        "strict_accas": 0,
        "required_games": 3,
        "effective_paper_legs": None,
        "eligibility_blockers": {
            "below_calibrated_probability_floor": 1,
            "insufficient_walk_forward_samples": 2,
            "no_exact_bookmaker_market": 1,
        },
    }
    ticket_state = json.loads(Path(result["ticket_state"]).read_text())
    assert ticket_state["scans"][-1]["candidate_funnel"] == result["candidate_funnel"]
    unpriced_scan = ticket_state["scans"][-1]["unpriced_audit"]
    assert unpriced_scan["audit_only"] is True
    assert unpriced_scan["ticketable"] is False
    assert unpriced_scan["stake_units"] == 0.0
    assert unpriced_scan["count"] == 1
    assert unpriced_scan["grade_counts"] == {"pending": 1}
    assert unpriced_scan["price_evidence_counts"] == {"UNMATCHED": 1}
    assert "candidate_records" not in unpriced_scan
    records = result["unpriced_audit"]["candidate_records"]
    assert len(records) == 1
    assert records[0]["candidate_id"] == "forecast-3-moneyline-home"
    assert records[0]["market"] == "moneyline"
    assert records[0]["selection"] == "home"
    assert records[0]["price"] is None
    assert records[0]["ticket_id"] is None
    assert records[0]["ticket_status"] == "audit_only_no_exact_selected_price"
    assert records[0]["ticketable"] is False
    assert records[0]["stake_units"] == 0.0
    assert records[0]["historical_success_rate"] is None
    assert records[0]["historical_success_decisions"] == 0
    assert records[0]["historical_success_qualifies"] is False
    assert records[0]["historical_success_status"] == "insufficient_history"
    assert unpriced_scan["historical_success_qualified_count"] == 0
    assert unpriced_scan["historical_success_policy"]["threshold_at_least"] == 0.70
    assert unpriced_scan["historical_success_policy"]["minimum_decisions_exclusive"] == 10
    assert result["counts"]["current_unpriced_picks"] == 1
    assert result["counts"]["dated_unpriced_audit_picks"] == 1
    buckets = {row["bucket"] for row in json.loads(Path(result["dated_picks"]).read_text())}
    assert buckets == {"ELIGIBLE_CLEAN", "CAUTION", "WATCHLIST_NO_ODDS", "SKIPPED_INELIGIBLE"}
    audit = json.loads(Path(result["rolling_audit"]).read_text())
    assert audit["archived_pick_rows"] == 4
    assert audit["same_day_rows_excluded"] == 4
    assert audit["settled_picks"] == 0


def test_roi_qualified_pick_is_published_as_a_paper_single_without_an_acca(tmp_path):
    row = forecast(1)
    band = {
        "league": "NBA",
        "market": "moneyline",
        "period": "full_game",
        "model": "provider-history",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "candidate_pick_count": 1,
        "historical_pick_count": 6,
        "wins": 6,
        "losses": 0,
        "pushes": 0,
        "observed_roi": 0.8,
        "recency_weighted_roi": 0.8,
        "parent_roi": 0.0,
        "adjusted_roi": 0.2,
        "qualifies": True,
    }
    paper_filter = {
        "candidate_pick_count": 1,
        "candidate_band_count": 1,
        "qualifying_pick_count": 1,
        "qualifying_band_count": 1,
        "historical_unique_pick_count": 6,
        "candidate_band_details": {row["forecast_id"]: band},
        "bands": [band],
        "status": "applied",
    }
    market_report = report([row])
    market_report["accas"].update(
        paper_results=[],
        strict_results=[],
        requested_legs=3,
        effective_paper_legs=None,
        paper_single_picks=[{**row, "individual_pick_roi_band": band}],
        paper_pick_filter=paper_filter,
    )

    result = publish(tmp_path, Store([game(1)]), market_report)

    today = json.loads(Path(result["picks_today"]).read_text())
    assert len(today) == 1
    assert today[0]["paper_roi_band_eligible"] is True
    assert today[0]["paper_roi_band"]["qualifies"] is True
    picks_text = Path(result["picks_text"]).read_text()
    assert "PAPER SINGLE · moneyline full_game" in picks_text
    assert "ROI-qualified individual paper singles: 1" in picks_text
    assert result["ticket_receipt"] is not None
    receipt = json.loads(Path(result["ticket_receipt"]).read_text())
    assert receipt["status"] == "frozen"
    assert receipt["paper_only"] is True
    assert receipt["accas"] == []
    assert len(receipt["singles"]) == 1
    single = receipt["singles"][0]
    assert single["ticket_type"] == "single"
    assert single["leg_count"] == 1
    assert single["stake_units"] == 1.0
    assert single["combined_decimal_odds"] == 1.8
    assert single["estimated_win_probability"] == 0.62
    assert single["estimated_joint_win_probability"] == 0.62
    assert single["individual_pick_filter"] == paper_filter
    assert single["legs"][0]["individual_pick_roi_band"] == band
    ticket_text = Path(result["ticket_text"]).read_text()
    assert "[PAPER SINGLE #1]" in ticket_text
    assert "1 paper single · 1 ticket publication" in ticket_text
    assert "ROI band: moneyline full_game" in ticket_text
    assert "NO MULTI-LEG PAPER ACCA" not in ticket_text
    assert "CANDIDATE FUNNEL" not in ticket_text
    assert "recency ROI" not in ticket_text
    assert result["candidate_funnel"]["paper_single_picks"] == 1
    assert result["candidate_funnel"]["paper_pick_filter"] == paper_filter
    ticket_state = json.loads(Path(result["ticket_state"]).read_text())
    saved_funnel = ticket_state["scans"][-1]["candidate_funnel"]
    assert saved_funnel.keys() == result["candidate_funnel"].keys()
    assert saved_funnel["paper_single_picks"] == 1
    expected_filter = {
        key: value
        for key, value in paper_filter.items()
        if key not in {"candidate_band_details", "bands"}
    }
    assert saved_funnel["paper_pick_filter"] == expected_filter
    assert ticket_state["scans"][-1]["fresh_single_tickets"] == 1
    assert result["counts"]["published_accas"] == 0
    assert result["counts"]["published_singles"] == 1
    assert result["counts"]["pending_positions"] == 1
    assert result["counts"]["pending_singles"] == 1
    assert result["counts"]["paper_single_picks"] == 1
    performance = json.loads(Path(result["performance"]).read_text())
    assert performance["settled_positions"] == 0
    assert performance["pending_positions"] == 1
    assert performance["pending_singles"] == 1
    assert performance["paper_committed_units"] == 1.0


def test_multiple_roi_qualified_selections_in_one_game_do_not_force_singles(tmp_path):
    rows = [
        forecast(1, market="moneyline", selection="home"),
        forecast(1, market="total", selection="under"),
    ]
    bands = {
        row["forecast_id"]: {
            "market": row["market"],
            "period": row["period"],
            "candidate_pick_count": 2,
            "historical_pick_count": 6,
            "wins": 6,
            "losses": 0,
            "pushes": 0,
            "observed_roi": 0.8,
            "adjusted_roi": 0.2,
            "qualifies": True,
        }
        for row in rows
    }
    for row in rows:
        row["individual_pick_roi_band"] = bands[row["forecast_id"]]
    paper_filter = {
        "candidate_pick_count": 2,
        "qualifying_pick_count": 2,
        "qualifying_band_count": 2,
        "historical_unique_pick_count": 12,
        "candidate_band_details": bands,
        "status": "applied",
    }
    market_report = report(rows)
    market_report["accas"].update(
        paper_results=[],
        paper_single_picks=rows,
        paper_pick_filter=paper_filter,
    )

    result = publish(tmp_path, Store([game(1)]), market_report)

    assert result["ticket_receipt"] is None
    assert result["counts"]["published_singles"] == 0
    ticket_text = Path(result["ticket_text"]).read_text()
    assert (
        "NO MULTI-LEG PAPER ACCA — 2 of 2 candidate picks passed individual-pick ROI; "
        "only 1 distinct game qualifies; 2 are required."
    ) in ticket_text


def test_single_ticket_does_not_bypass_non_sample_eligibility_gates(tmp_path):
    row = forecast(
        1,
        eligible=False,
        reasons=["below_calibrated_probability_floor"],
    )
    band = {
        "market": "moneyline",
        "period": "full_game",
        "candidate_pick_count": 1,
        "historical_pick_count": 6,
        "observed_roi": 0.8,
        "adjusted_roi": 0.2,
        "qualifies": True,
    }
    row["individual_pick_roi_band"] = band
    market_report = report([row])
    market_report["accas"].update(
        paper_results=[],
        paper_single_picks=[row],
        paper_pick_filter={
            "candidate_band_details": {row["forecast_id"]: band},
            "candidate_pick_count": 1,
            "qualifying_pick_count": 1,
        },
    )

    result = publish(tmp_path, Store([game(1)]), market_report)

    assert result["ticket_receipt"] is None
    assert result["counts"]["published_singles"] == 0
    assert "eligibility/price evidence could not be frozen as a single" in Path(
        result["ticket_text"]
    ).read_text()


def test_paper_single_settles_at_single_odds_and_keeps_write_once_receipt(tmp_path):
    first_pick = forecast(1)
    first_band = {
        "market": "moneyline",
        "period": "full_game",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "candidate_pick_count": 1,
        "historical_pick_count": 6,
        "wins": 6,
        "losses": 0,
        "pushes": 0,
        "observed_roi": 0.8,
        "adjusted_roi": 0.2,
        "qualifies": True,
    }
    first_filter = {
        "candidate_pick_count": 1,
        "qualifying_pick_count": 1,
        "qualifying_band_count": 1,
        "historical_unique_pick_count": 6,
        "candidate_band_details": {first_pick["forecast_id"]: first_band},
        "status": "applied",
    }
    first_pick["individual_pick_roi_band"] = first_band
    first_report = report([first_pick])
    first_report["accas"].update(
        paper_results=[],
        paper_single_picks=[first_pick],
        paper_pick_filter=first_filter,
    )

    first = publish(tmp_path, Store([game(1)]), first_report, run_id="single-first")
    receipt_path = Path(first["ticket_receipt"])
    frozen_bytes = receipt_path.read_bytes()
    pending = json.loads(Path(first["performance"]).read_text())
    assert pending["pending_singles"] == 1
    assert pending["pending_positions"] == 1
    assert pending["settled_accas"] == 0
    assert pending["paper_committed_units"] == 1.0

    replacement_pick = forecast(2)
    replacement_pick["individual_pick_roi_band"] = first_band
    replacement_report = report([replacement_pick])
    replacement_report["accas"].update(
        paper_results=[],
        paper_single_picks=[replacement_pick],
        paper_pick_filter=first_filter,
    )
    settled = publish_operator_reports(
        Store([game(1, home_score=110, away_score=100), game(2)]),
        League.NBA,
        replacement_report,
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T23:00:00Z",
        run_id="single-replacement-must-not-rotate",
    )

    assert receipt_path.read_bytes() == frozen_bytes
    performance = json.loads(Path(settled["performance"]).read_text())
    assert performance["settled_singles"] == 1
    assert performance["settled_positions"] == 1
    assert performance["settled_accas"] == 0
    assert performance["wins"] == 1
    assert performance["profit_units"] == pytest.approx(0.8)
    assert performance["roi"] == pytest.approx(0.8)
    assert performance["settled_stake_units"] == 1.0
    assert performance["paper_bank_units"] == pytest.approx(100.8)
    assert performance["paper_committed_units"] == 0.0
    assert performance["selected_leg_brier_score"] == pytest.approx((0.62 - 1) ** 2)
    daily = performance["history"][0]
    assert [position["ticket_type"] for position in daily["positions"]] == ["single"]
    assert daily["positions"][0]["grade_status"] == "win"
    assert daily["positions"][0]["profit_units"] == pytest.approx(0.8)
    ticket_text = Path(settled["ticket_text"]).read_text()
    assert "[PAPER SINGLE #1]" in ticket_text
    assert "· ✅ WIN" in ticket_text
    assert "Home 1 vs Away 1" in ticket_text
    performance_text = Path(settled["performance_text"]).read_text()
    record_line = next(line for line in performance_text.splitlines() if line.startswith("RECORD"))
    bet_day_line = next(
        line for line in performance_text.splitlines() if line.startswith("  2026-11-07")
    )
    assert "✅ WIN 1W" in record_line and "❌ LOSS" not in record_line
    assert "➖ PUSH" not in record_line
    assert "✅ WIN 1W" in bet_day_line and "❌ LOSS" not in bet_day_line
    assert "➖ PUSH" not in bet_day_line


def test_provisional_sample_only_acca_is_frozen_as_labeled_paper_evidence(tmp_path):
    provisional = forecast(
        1,
        eligible=False,
        reasons=["insufficient_walk_forward_samples"],
        calibration_n=4,
    )
    strict = forecast(2)
    ticket = acca([provisional, strict])
    ticket.update(
        provisional_sample_gate_bypassed=True,
        provisional_leg_count=1,
        ranking_status="paper_only_provisional_sample_gate_bypassed",
    )
    band_details = {
        "league": "NBA",
        "market": "moneyline",
        "period": "full_game",
        "model": "provider-history",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "candidate_pick_count": 2,
        "historical_pick_count": 30,
        "wins": 18,
        "losses": 12,
        "pushes": 0,
        "observed_roi": 0.12,
        "recency_weighted_roi": 0.10,
        "parent_roi": 0.03,
        "adjusted_roi": 0.09,
        "qualifies": True,
    }
    pick_filter = {
        "policy": "adaptive_recency_weighted_positive_individual_pick_roi_bands",
        "history_source": "settled_pick_audit",
        "performance_basis": "settled_individual_picks_only_not_acca_outcomes",
        "candidate_pick_count": 2,
        "candidate_band_count": 1,
        "qualifying_pick_count": 2,
        "qualifying_band_count": 1,
        "historical_unique_pick_count": 30,
        "candidate_band_details": {
            "forecast-1-moneyline-home": band_details,
            "forecast-2-moneyline-home": band_details,
        },
        "bands": [band_details],
        "status": "applied",
    }
    ticket["individual_pick_filter"] = pick_filter
    market_report = report([provisional, strict], [ticket])
    market_report["accas"].update(
        requested_legs=2,
        results=[],
        strict_results=[],
        paper_results=[ticket],
        paper_pick_filter=pick_filter,
    )

    result = publish(
        tmp_path,
        Store([game(1), game(2)]),
        market_report,
    )
    receipt = json.loads(Path(result["ticket_receipt"]).read_text())
    frozen_acca = receipt["accas"][0]
    assert frozen_acca["paper_only"] is True
    assert frozen_acca["provisional_sample_gate_bypassed"] is True
    assert frozen_acca["provisional_leg_count"] == 1
    assert frozen_acca["ranking_status"] == "paper_only_provisional_sample_gate_bypassed"
    assert frozen_acca["individual_pick_filter"] == pick_filter
    assert frozen_acca["legs"][0]["calibration"]["n"] == 4
    assert result["candidate_funnel"]["paper_pick_filter"] == pick_filter
    assert result["candidate_funnel"]["provisional_sample_only_rows"] == 1
    assert result["candidate_funnel"]["paper_candidate_games"] == 2
    assert result["candidate_funnel"]["paper_accas"] == 1
    assert result["candidate_funnel"]["strict_accas"] == 0

    ticket_text = Path(result["ticket_text"]).read_text()
    assert "PROVISIONAL SAMPLE GATE" in ticket_text
    assert (
        "Individual-pick ROI policy: 2/2 eligible candidates qualify across 1 positive "
        "bands; 30 unique settled picks · status=applied."
    ) in ticket_text
    assert (
        "ROI band: moneyline full_game · p 0.60-<0.70 · odds 1.50-<2.00 · "
        "n=30 W-L-P=18-12-0 · ROI +12.0% observed / +9.0% adjusted."
    ) in ticket_text
    assert "PROVISIONAL PAPER ONLY — sample-size gate bypassed" not in ticket_text
    assert "game-1 n=4/?" not in ticket_text


def test_roi_band_history_loads_settled_single_picks_not_acca_performance(tmp_path):
    settled = forecast(1)
    settled["calibrated_probability"] = 0.67
    settled["outcome"] = "win"
    settled["grade_status"] = "win"
    pending = json.loads(json.dumps(forecast(2)))
    pending["calibrated_probability"] = 0.85
    pending["outcome"] = None
    pending["grade_status"] = "pending"
    (tmp_path / "picks_audit_rolling.json").write_text(
        json.dumps({"settled_ledger": [settled, pending]})
    )
    (tmp_path / "auto_tickets_performance.json").write_text(
        json.dumps({"settled_accas": 999, "history": [{"status": "won"}]})
    )

    history = load_settled_pick_history(tmp_path)
    assert history is not None
    assert len(history) == 1
    assert history[0].outcome == "win"
    assert history[0].calibrated_probability == 0.67
    assert history[0].decimal_price == 1.8


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
    first = publish(
        tmp_path,
        Store([game(1)]),
        report([original]),
        run_id="first",
        generated_utc=AS_OF,
    )
    dated_path = Path(first["dated_picks"])

    later = json.loads(json.dumps(original))
    later["forecast_id"] = "forecast-1-total-under-later"
    later["book"]["observed_utc"] = "2026-11-07T17:30:00Z"
    later["book"]["selected_decimal_price"] = 1.95
    later["prediction"]["observed_utc"] = "2026-11-07T17:30:00Z"
    publish(
        tmp_path,
        Store([game(1)]),
        report([later]),
        run_id="second",
        generated_utc="2026-11-07T18:30:00Z",
    )

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


def test_paper_fade_audit_uses_exact_opposite_quotes_only(tmp_path):
    fade_loss = with_quote_evidence(
        forecast(1, price=1.8),
        {"home": 1.8, "away": 2.1, "draw": None, "over": None, "under": None},
    )
    fade_win = with_quote_evidence(
        forecast(2, price=None),
        {"home": None, "away": 2.2, "draw": None, "over": None, "under": None},
    )
    fade_push = with_quote_evidence(
        forecast(3, market="total", selection="over", price=None),
        {"home": None, "away": None, "draw": None, "over": 1.8, "under": 1.95},
    )
    legacy_one_sided = forecast(4, price=1.8)

    result = publish(
        tmp_path,
        Store(
            [
                game(1, home_score=105, away_score=100),
                game(2, home_score=100, away_score=105),
                game(3, total=200),
                game(4, home_score=100, away_score=105),
            ]
        ),
        report([fade_loss, fade_win, fade_push, legacy_one_sided]),
    )
    audit = json.loads(Path(result["rolling_audit"]).read_text())
    fade = audit["fade_audit"]

    assert fade["schema_version"] == "magolide.paper_fade_audit"
    assert fade["paper_only"] is True
    assert fade["selected_price_required"] is False
    assert fade["candidate_sample"] == 4
    assert fade["priced_sample"] == 3
    assert fade["unpriced_sample"] == 1
    assert fade["unpriced_reasons"] == {"opposite_quote_not_preserved": 1}
    assert (fade["wins"], fade["losses"], fade["pushes"]) == (1, 1, 1)
    assert fade["profit_units"] == pytest.approx(0.2)
    assert fade["roi"] == pytest.approx(0.2 / 3)
    assert fade["hit_rate"] == pytest.approx(0.5)
    assert fade["break_even_hit_rate"] == pytest.approx((1 / 2.1 + 1 / 2.2) / 2)
    assert fade["by_market"]["moneyline"]["candidate_sample"] == 3
    assert fade["by_market"]["moneyline"]["priced_sample"] == 2
    assert fade["by_market"]["total"]["pushes"] == 1
    assert fade["by_confidence_band"]["0.60-<0.70"]["priced_sample"] == 3

    # The selected-side ROI remains its existing calculation; the fade sample
    # can include quotes even when the selected-side price is missing.
    assert audit["priced_sample"] == 2
    assert audit["profit_units"] == pytest.approx(-0.2)
    assert audit["roi"] == pytest.approx(-0.1)
    dated_rows = json.loads(Path(result["dated_picks"]).read_text())
    legacy_saved = next(
        row for row in dated_rows if row["forecast_id"] == legacy_one_sided["forecast_id"]
    )
    assert "side_prices" not in legacy_saved["book"]
    legacy_rehydrated = _forecast_from_row(legacy_saved)
    assert legacy_rehydrated.book is not None
    assert legacy_rehydrated.book.prices == {"home": 1.8}

    markdown = Path(result["dated_audit"]).read_text()
    assert "## Paper-only opposite-side fade audit" in markdown
    assert "settled rows with exact opposite quotes: 3/4" in markdown
    assert "Counterfactual fade: **❌ LOSS @ 2.100 away · Test Book**" in markdown
    assert "Counterfactual fade: **not priced (opposite_quote_not_preserved)**" in markdown
    assert "No quote is inferred or historically backfilled" in markdown


def test_spread_displays_use_selected_side_handicap_everywhere(tmp_path):
    row = forecast(1, market="spread", selection="away", price=1.9090909090909092)
    row["line"] = -27.5  # canonical spread line is home-oriented
    row = with_quote_evidence(
        row,
        {
            "home": 1.9090909090909092,
            "away": 1.9090909090909092,
            "draw": None,
            "over": None,
            "under": None,
        },
    )

    assert _selection_label(row) == "SPREAD full_game +27.5 -> AWAY"
    assert _audit_market_selection(row) == "Full-game spread +27.5: Away 1 (away)"

    home_pick = {**row, "selection": "home"}
    assert _selection_label(home_pick) == "SPREAD full_game -27.5 -> HOME"
    assert _audit_market_selection(home_pick) == "Full-game spread -27.5: Home 1 (home)"

    result = publish(
        tmp_path,
        Store([game(1, home_score=76, away_score=44)]),
        report([row]),
    )
    markdown = Path(result["dated_audit"]).read_text()
    assert "❌ LOSS — Full-game spread +27.5: Away 1 (away)" in markdown
    assert "Counterfactual fade: **✅ WIN @ 1.909 Home 1 -27.5 (home) · Test Book**" in markdown


def test_later_exact_quote_hides_stale_unpriced_logical_pick_but_keeps_evidence(tmp_path):
    unpriced = forecast(
        1,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
        calibration_n=4,
    )
    unpriced["forecast_id"] = "forecast-1-before-quote"
    first = publish(
        tmp_path,
        Store([game(1)]),
        report([unpriced]),
        run_id="before-quote",
        generated_utc=AS_OF,
    )
    assert first["unpriced_audit"]["count"] == 1

    priced = forecast(
        1,
        price=1.91,
        eligible=False,
        reasons=["insufficient_walk_forward_samples"],
        calibration_n=4,
    )
    priced["forecast_id"] = "forecast-1-after-quote"
    result = publish(
        tmp_path,
        Store([game(1)]),
        report([priced]),
        run_id="after-quote",
        generated_utc="2026-11-07T18:30:00Z",
    )

    # The original no-price evidence remains immutable in the dated ledger,
    # but the card/audit reflects the later exact quote for the same selection.
    archived = json.loads((tmp_path / f"picks_{DATE}.json").read_text())
    assert {row["forecast_id"] for row in archived} == {
        "forecast-1-before-quote",
        "forecast-1-after-quote",
    }
    text = Path(result["picks_text"]).read_text()
    assert "WATCHLIST — NO MATCHED ODDS\n========================================================================\n  (none)" in text
    assert "@1.910 Test Book" in text
    assert "Logical picks shown: 1 · superseded observations: 1" in text
    assert result["unpriced_audit"]["count"] == 0
    state = json.loads(Path(result["ticket_state"]).read_text())
    assert state["scans"][-1]["unpriced_audit"]["count"] == 0


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
    assert "grade=✅ WIN" in text.split(settled_heading)[1]
    assert "IN_PROGRESS_OR_AWAITING_RESULT=0 SETTLED=1" in text


def test_unpriced_forecasts_are_audit_only_and_settle_without_paper_stake(tmp_path):
    unpriced = forecast(
        1,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    store = Store([game(1)])

    first = publish(tmp_path, store, report([unpriced]))

    assert first["ticket_receipt"] is None
    state = json.loads(Path(first["ticket_state"]).read_text())
    assert state["frozen_by_date"][DATE]["no_bet"] is True
    assert first["no_bet_decision"] is not None
    card = Path(first["ticket_text"]).read_text()
    assert "UNPRICED AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY" in card
    assert "forecast=forecast-1-moneyline-home" not in card
    assert "NO PAPER BET" in card
    audit = json.loads(Path(first["rolling_audit"]).read_text())
    assert audit["pending_picks_total"] == 1
    assert audit["priced_sample"] == 0
    assert audit["settled_ledger"] == []
    performance = json.loads(Path(first["performance"]).read_text())
    assert performance["pending_positions"] == 0
    assert performance["paper_bank_units"] == 100.0

    settled = publish_operator_reports(
        Store([game(1, home_score=105, away_score=100)]),
        League.NBA,
        report([]),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-08T03:35:03Z",
        run_id="official-final-no-price",
    )

    assert settled["ticket_receipt"] is None
    ticket_text = Path(settled["ticket_text"]).read_text()
    assert "UNPRICED AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY" in ticket_text
    assert "forecast=forecast-1-moneyline-home" not in ticket_text
    settled_records = settled["unpriced_audit"]["candidate_records"]
    assert len(settled_records) == 1
    assert settled_records[0]["candidate_id"] == "forecast-1-moneyline-home"
    assert settled_records[0]["grade_status"] == "win"
    assert settled_records[0]["outcome"] == "win"
    assert settled_records[0]["price"] is None
    assert settled_records[0]["ticket_id"] is None
    assert settled_records[0]["historical_success_decisions"] == 0
    assert settled_records[0]["historical_success_qualifies"] is False
    settled_audit = json.loads(Path(settled["rolling_audit"]).read_text())
    assert settled_audit["settled_picks"] == 1
    assert settled_audit["wins"] == 1
    assert settled_audit["hit_rate"] == 1.0
    assert settled_audit["priced_sample"] == 0
    assert settled_audit["profit_units"] == 0.0
    assert settled_audit["roi"] is None
    assert settled_audit["settled_ledger"][0]["grade_status"] == "win"
    assert settled_audit["settled_ledger"][0]["book"] is None
    performance = json.loads(Path(settled["performance"]).read_text())
    assert performance["settled_positions"] == 0
    assert performance["paper_bank_units"] == 100.0
    assert json.loads(Path(settled["ticket_state"]).read_text())["frozen_by_date"][DATE]["no_bet"] is True


def test_unpriced_matched_market_reports_missing_selected_side_quote(tmp_path):
    row = forecast(
        1,
        price=None,
        eligible=False,
        reasons=["missing_selected_decimal_price", "insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    row = with_quote_evidence(
        row,
        {
            "home": None,
            "away": 1.95,
            "draw": None,
            "over": None,
            "under": None,
        },
        selected_decimal_price=None,
    )

    result = publish(tmp_path, Store([game(1)]), report([row]))

    assert result["ticket_receipt"] is None
    card = Path(result["ticket_text"]).read_text()
    assert "UNPRICED AUDIT-ONLY SHADOW GROUPS — 0 QUALIFY" in card
    assert "Test Book matched; selected-side quote missing" not in card
    assert "1.950" not in card
    state = json.loads(Path(result["ticket_state"]).read_text())
    unpriced_scan = state["scans"][-1]["unpriced_audit"]
    assert unpriced_scan["price_evidence_counts"] == {
        "MATCHED_BUT_SELECTED_PRICE_MISSING": 1
    }
    assert "candidate_records" not in unpriced_scan
    assert (
        result["unpriced_audit"]["candidate_records"][0]["price_state"]
        == "selected_side_quote_missing"
    )
    saved = json.loads(Path(result["dated_picks"]).read_text())[0]
    assert saved["book"]["selected_decimal_price"] is None
    assert saved["book"]["side_prices"]["away"] == 1.95
    assert state["frozen_by_date"][DATE]["no_bet"] is True
    assert result["no_bet_decision"] is not None


def test_unpriced_audit_rows_do_not_enter_a_frozen_priced_ticket(tmp_path):
    priced = [forecast(number) for number in (1, 2, 3)]
    unpriced = forecast(
        4,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    # Same fixture, different selections and market: none may be collapsed
    # into a generic fixture-level shadow record.
    unpriced_away = forecast(
        4,
        selection="away",
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    unpriced_total = forecast(
        4,
        market="total",
        selection="under",
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market", "insufficient_walk_forward_samples"],
        calibration_n=0,
    )
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    ) + seed_historical_success_picks(tmp_path, "total", wins=8, losses=3)
    result = publish(
        tmp_path,
        Store([game(number) for number in (1, 2, 3, 4)] + history_games),
        report([*priced, unpriced, unpriced_away, unpriced_total], [acca(priced)]),
    )

    assert result["ticket_receipt"] is not None
    receipt = json.loads(Path(result["ticket_receipt"]).read_text())
    assert len(receipt["accas"]) == 1
    assert len(receipt["accas"][0]["legs"]) == 3
    assert all(
        leg["forecast_id"] not in {
            unpriced["forecast_id"],
            unpriced_away["forecast_id"],
            unpriced_total["forecast_id"],
        }
        for leg in receipt["accas"][0]["legs"]
    )
    ticket_text = Path(result["ticket_text"]).read_text()
    assert "[PAPER 3-LEG ACCA #1]" in ticket_text
    assert (
        "UNPRICED AUDIT-ONLY SHADOW GROUPS (≥70% HISTORICAL SUCCESS, n>10) — "
        "3 picks in 3 groups (3 ⏳ PENDING; 3 audited)"
    ) in ticket_text
    assert "historical success=72.7% (8/11)" in ticket_text
    assert "SHADOW GROUP 1 — 1 audit-only pick (1 ⏳ PENDING)" in ticket_text
    assert "SHADOW GROUP 2 — 1 audit-only pick (1 ⏳ PENDING)" in ticket_text
    assert "SHADOW GROUP 3 — 1 audit-only pick (1 ⏳ PENDING)" in ticket_text
    assert "forecast=forecast-4-moneyline-home" in ticket_text
    assert "forecast=forecast-4-moneyline-away" in ticket_text
    assert "forecast=forecast-4-total-under" in ticket_text
    state = json.loads(Path(result["ticket_state"]).read_text())
    # Full rows remain in the run result/artifacts, not duplicated into every
    # durable scan-state snapshot.
    records = result["unpriced_audit"]["candidate_records"]
    assert {record["candidate_id"] for record in records} == {
        unpriced["forecast_id"],
        unpriced_away["forecast_id"],
        unpriced_total["forecast_id"],
    }
    assert {record["market"] for record in records} == {"moneyline", "total"}
    assert {record["selection"] for record in records} == {"home", "away", "under"}
    assert all(record["historical_success_rate"] == 0.727273 for record in records)
    assert all(record["historical_success_decisions"] == 11 for record in records)
    assert all(record["historical_success_qualifies"] is True for record in records)
    scan = state["scans"][-1]
    assert "candidate_records" not in scan["unpriced_audit"]
    assert "forecast_ids" not in scan["unpriced_audit"]
    assert scan["unpriced_audit"]["historical_success_qualified_count"] == 3
    group_policy = state["scans"][-1]["unpriced_audit"]["shadow_group_policy"]
    assert group_policy["display_only"] is True
    assert group_policy["max_picks_per_group"] == 3
    assert group_policy["one_pick_per_game"] is True
    assert group_policy["joint_success_rate_calculated"] is False
    assert group_policy["ticket_id"] is None
    assert group_policy["receipt_created"] is False
    assert group_policy["stake_units"] == 0.0
    assert all(record["ticket_id"] is None for record in records)
    assert all(record["ticketable"] is False for record in records)
    assert all(record["stake_units"] == 0.0 for record in records)
    performance = json.loads(Path(result["performance"]).read_text())
    assert performance["pending_positions"] == 1
    assert performance["paper_committed_units"] == 1.0
    assert result["counts"]["current_unpriced_picks"] == 3
    assert result["counts"]["dated_unpriced_audit_picks"] == 3


def test_ncaam_shadow_uses_priced_ineligible_picks_with_odds_and_probability_floors(tmp_path):
    def candidate(number, *, price, probability, raw_probability=0.60, eligible=False, reasons=None):
        row = forecast(
            number,
            price=price,
            eligible=eligible,
            reasons=reasons,
            calibration_n=10 if reasons == ["insufficient_walk_forward_samples"] else 40,
        )
        row["league"] = "NCAA_M"
        row["calibrated_probability"] = probability
        row["raw_probability"] = raw_probability
        if price is not None:
            row = with_quote_evidence(
                row,
                {"home": price, "away": 2.0, "draw": None, "over": None, "under": None},
                selected_decimal_price=price,
            )
        return row

    threshold_candidate = candidate(
        1,
        price=1.30,
        probability=0.70,
        reasons=["insufficient_walk_forward_samples"],
    )
    higher_candidate = candidate(
        2,
        price=2.0,
        probability=0.80,
        reasons=["insufficient_team_specific_history"],
    )
    below_odds_floor = candidate(
        3,
        price=1.29,
        probability=0.90,
        reasons=["insufficient_team_specific_history"],
    )
    below_probability_floor = candidate(
        4,
        price=2.0,
        probability=0.699,
        raw_probability=0.99,
        reasons=["insufficient_team_specific_history"],
    )
    raw_probability_fallback = candidate(
        7,
        price=1.30,
        probability=None,
        raw_probability=0.70,
        reasons=["insufficient_team_specific_history"],
    )
    eligible = candidate(
        5,
        price=1.80,
        probability=0.85,
        eligible=True,
    )
    unpriced = candidate(
        6,
        price=None,
        probability=0.90,
        reasons=["no_exact_bookmaker_market"],
    )

    result = publish_operator_reports(
        Store([]),
        League.NCAA_M,
        report([
            threshold_candidate,
            higher_candidate,
            below_odds_floor,
            below_probability_floor,
            raw_probability_fallback,
            eligible,
            unpriced,
        ]),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T06:00:00Z",
        run_id="ncaam-priced-ineligible-shadow",
    )

    card = Path(result["ticket_text"]).read_text()
    assert (
        "PRICED-INELIGIBLE AUDIT-ONLY SHADOW GROUPS "
        "(decimal odds ≥1.30; p ≥70%) — 3 picks in 1 group "
        "(3 ⏳ PENDING; 5 exact-priced ineligible audited)"
    ) in card
    assert "selected decimal odds=1.300 Test Book" in card
    assert "p=70.0% (calibrated_probability)" in card
    assert "bucket=CAUTION" in card
    assert f"forecast={threshold_candidate['forecast_id']}" in card
    assert f"forecast={higher_candidate['forecast_id']}" in card
    assert f"forecast={raw_probability_fallback['forecast_id']}" in card
    assert "p=70.0% (raw_probability)" in card
    assert "UNPRICED AUDIT-ONLY SHADOW GROUPS" not in card
    for excluded in (below_odds_floor, below_probability_floor, eligible, unpriced):
        assert f"forecast={excluded['forecast_id']}" not in card

    assert result["ticket_receipt"] is None
    assert result["unpriced_audit"]["count"] == 1
    shadow = result["priced_ineligible_shadow"]
    assert shadow["priced_ineligible_audited_count"] == 5
    assert shadow["qualified_count"] == 3
    assert shadow["minimum_selected_decimal_odds"] == 1.30
    assert shadow["minimum_probability"] == 0.70
    assert {row["candidate_id"] for row in shadow["candidate_records"]} == {
        threshold_candidate["forecast_id"],
        higher_candidate["forecast_id"],
        raw_probability_fallback["forecast_id"],
    }
    assert all(row["acca_eligible"] is False for row in shadow["candidate_records"])
    fallback_record = next(
        row for row in shadow["candidate_records"]
        if row["candidate_id"] == raw_probability_fallback["forecast_id"]
    )
    assert fallback_record["probability"] == 0.70
    assert fallback_record["probability_basis"] == "raw_probability"
    assert all(row["ticketable"] is False for row in shadow["candidate_records"])
    assert all(row["stake_units"] == 0.0 for row in shadow["candidate_records"])

    state = json.loads(Path(result["ticket_state"]).read_text())
    scan_shadow = state["scans"][-1]["priced_ineligible_shadow"]
    assert scan_shadow["qualified_count"] == 3
    assert scan_shadow["shadow_groups"] == [[
        higher_candidate["forecast_id"],
        threshold_candidate["forecast_id"],
        raw_probability_fallback["forecast_id"],
    ]]
    assert state["scans"][-1]["unpriced_audit"]["count"] == 1
    assert state["frozen_by_date"] == {}


def test_qualified_no_price_picks_are_grouped_without_creating_a_ticket(tmp_path):
    history_games = seed_historical_success_picks(
        tmp_path, "moneyline", wins=8, losses=3
    )
    rows = [
        forecast(
            number,
            price=None,
            eligible=False,
            reasons=["no_exact_bookmaker_market"],
            calibration_n=0,
        )
        for number in (5, 6, 7)
    ]
    result = publish(
        tmp_path,
        Store([game(number) for number in (5, 6, 7)] + history_games),
        report(rows),
    )

    assert result["ticket_receipt"] is None
    card = Path(result["ticket_text"]).read_text()
    assert (
        "UNPRICED AUDIT-ONLY SHADOW GROUPS (≥70% HISTORICAL SUCCESS, n>10) — "
        "3 picks in 1 group (3 ⏳ PENDING; 3 audited)"
    ) in card
    assert "SHADOW GROUP 1 — 3 audit-only picks (3 ⏳ PENDING)" in card
    assert card.count("  SHADOW GROUP ") == 1
    assert "Display grouping only: no odds, combined success rate, ticket ID, receipt, or stake" in card
    for row in rows:
        assert f"forecast={row['forecast_id']}" in card

    audit = result["unpriced_audit"]
    assert audit["count"] == 3
    assert audit["historical_success_qualified_count"] == 3
    assert len(audit["candidate_records"]) == 3
    assert all(record["historical_success_qualifies"] is True for record in audit["candidate_records"])
    assert all(record["ticket_id"] is None for record in audit["candidate_records"])
    assert all(record["ticketable"] is False for record in audit["candidate_records"])
    assert all(record["stake_units"] == 0.0 for record in audit["candidate_records"])
    state = json.loads(Path(result["ticket_state"]).read_text())
    assert state["frozen_by_date"] == {}
    performance = json.loads(Path(result["performance"]).read_text())
    assert performance["pending_positions"] == 0
    assert performance["paper_bank_units"] == 100.0


def test_unpriced_success_rate_requires_70_percent_and_more_than_10_decisions(tmp_path):
    cases = [
        ("ten-decisions", 8, 2, 0.8, 10, "insufficient_history"),
        ("low-hit-rate", 7, 4, 0.636364, 11, "below_threshold"),
    ]
    for name, wins, losses, expected_rate, expected_n, expected_status in cases:
        localdata = tmp_path / name
        localdata.mkdir()
        history_games = seed_historical_success_picks(
            localdata, "moneyline", wins=wins, losses=losses
        )
        row = forecast(
            5,
            price=None,
            eligible=False,
            reasons=["no_exact_bookmaker_market"],
            calibration_n=40,
        )
        result = publish(
            localdata,
            Store([game(5), *history_games]),
            report([row]),
            run_id=name,
        )

        record = result["unpriced_audit"]["candidate_records"][0]
        assert record["historical_success_rate"] == expected_rate
        assert record["historical_success_decisions"] == expected_n
        assert record["historical_success_qualifies"] is False
        assert record["historical_success_status"] == expected_status
        assert "0 QUALIFY" in Path(result["ticket_text"]).read_text()
        assert row["forecast_id"] not in Path(result["ticket_text"]).read_text()


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
    overall = audit_text.split("## Overall", 1)[1].split("## ", 1)[0]
    assert "- settled picks: 1 (✅ WIN 1)" in overall
    assert "❌ LOSS" not in overall
    assert "➖ PUSH" not in overall
    assert "## Settled picks — official-final audit" in audit_text
    assert "### 2026-11-07: Home 1 vs Away 1 (Final: **105–100**)" in audit_text
    assert "**✅ WIN — Full-game moneyline: Home 1 (home)**" in audit_text
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
    assert "**✅ WIN — Full-game moneyline: Home 1 (home)**" in audit_text
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
            "model": "team_history_empirical",
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
    assert "**✅ WIN — Full-game moneyline: Home 1 (home)**" in text
    assert "**❌ LOSS — First-half moneyline: Away 1 (away)**" in text
    assert "Actual First-half score: **52.5–50**" in text
    assert "31.0%** calibrated (raw 28.0%)" in text
    assert "**❌ LOSS — Full-game total 200: Under**" in text
    assert "Actual Full-game total: **205**" in text
    assert "Price: **1.909 · Test Book**" in text
    assert "Price: **No exact bookmaker price**" in text


def test_euroleague_ticket_card_identifies_strict_named_book_ticket_lane():
    card = _render_empty_ticket_scan(
        league=League.EUROLEAGUE,
        target_date=DATE,
        generated_utc=AS_OF,
        performance={
            "paper_bank_units": 100.0,
            "paper_free_units": 100.0,
            "paper_committed_units": 0.0,
        },
        timezone_name="UTC",
        candidate_funnel={
            "scan_rows": 1,
            "paper_candidate_rows": 1,
            "paper_single_picks": 0,
            "paper_single_games": 0,
            "paper_pick_filter": {
                "candidate_pick_count": 1,
                "qualifying_pick_count": 0,
                "status": "no_positive_expected_value_candidates",
            },
        },
    )

    assert "EUROLEAGUE STRICT NAMED-BOOK TICKET LANE" in card
    assert "aggregate consensus prices are excluded" in card
    assert "calibrated p × odds > 1.00" in card
    assert "at least 10 past settled picks" in card


def test_retired_schema_tags_still_load_and_state_adopts_current_tag(tmp_path):
    from forecasting.operator import STATE_SCHEMA, TICKET_SCHEMA, _read_json
    from ingestion.models import schema_matches

    retired = ".v" + "1"  # spelling used by files written before the rename
    assert schema_matches(TICKET_SCHEMA + retired, TICKET_SCHEMA)
    assert schema_matches(TICKET_SCHEMA, TICKET_SCHEMA)
    assert not schema_matches(TICKET_SCHEMA + ".extra", TICKET_SCHEMA)
    assert not schema_matches(None, TICKET_SCHEMA)

    (tmp_path / "auto_tickets_state.json").write_text(
        json.dumps({"schema_version": STATE_SCHEMA + retired, "frozen_by_date": {}, "scans": []})
    )
    assert _load_state(tmp_path)["schema_version"] == STATE_SCHEMA

    # Pick history written before the rename is read back under current names.
    picks = tmp_path / "picks_2026-11-17.json"
    old_model = "provider:espn+team_history_empirical" + "_v" + "1"
    picks.write_text(json.dumps([{"model": old_model, "calibration_version": "provider_prediction"
                                  + "_v" + "2-7fb3e30d878d"}]))
    assert _read_json(picks, []) == [
        {"model": "provider:espn+team_history_empirical",
         "calibration_version": "provider_prediction-7fb3e30d878d"}
    ]


def _unpriced_moneyline(number, probability):
    row = forecast(
        number,
        price=None,
        eligible=False,
        reasons=["no_exact_bookmaker_market"],
        calibration_n=40,
    )
    row["raw_probability"] = probability
    row["calibrated_probability"] = probability
    return row


def _posted_spread(number, *, home_line, model_selection, side_prices):
    row = forecast(
        number,
        market="spread",
        selection=model_selection,
        price=side_prices[model_selection],
        eligible=False,
        reasons=["insufficient_team_specific_history"],
    )
    row["line"] = home_line
    row["calibrated_probability"] = 0.95
    return with_quote_evidence(row, side_prices)


def test_unpriced_moneyline_picks_are_shadow_tracked_against_posted_spreads(tmp_path):
    # Game 1: no moneyline price, but the book posts home -10.5. The model's own
    # spread pick is the underdog (+10.5); the shadow takes the moneyline side.
    rows = [
        _unpriced_moneyline(1, 0.97),
        _posted_spread(
            1,
            home_line=-10.5,
            model_selection="away",
            side_prices={"home": 1.909, "away": 1.952},
        ),
        # Game 2: strong no-price moneyline pick but no posted spread at all.
        _unpriced_moneyline(2, 0.90),
        # Game 3: below the probability floor, never re-expressed.
        _unpriced_moneyline(3, 0.65),
    ]
    first = publish(tmp_path, Store([game(1), game(2), game(3)]), report(rows))

    shadow = first["unpriced_spread_shadow"]
    assert shadow["audit_only"] is True and shadow["stake_units"] == 0.0
    assert shadow["policy"]["model_derived_lines"] is False
    assert shadow["today"]["tracked"] == 1
    assert shadow["today"]["untracked_no_posted_spread"] == 1
    tracked = next(r for r in shadow["candidate_records"] if r["posted_spread"])
    assert tracked["game_id"] == "game-1"
    assert tracked["selection"] == "home"
    assert tracked["selected_line"] == -10.5
    assert tracked["price"] == 1.909
    assert tracked["grade_status"] == "pending"
    untracked = next(r for r in shadow["candidate_records"] if not r["posted_spread"])
    assert untracked["game_id"] == "game-2"
    assert untracked["grade_status"] == "untracked_no_posted_spread"
    assert not any(r["game_id"] == "game-3" for r in shadow["candidate_records"])
    card = Path(first["ticket_text"]).read_text()
    assert "UNPRICED → POSTED-SPREAD SHADOW" in card
    assert "SPREAD full_game -10.5 -> HOME @1.909" in card
    assert "1 qualifying no-price moneyline pick had no posted spread to track." in card
    state = json.loads(Path(first["ticket_state"]).read_text())
    assert "candidate_records" not in state["scans"][-1]["unpriced_spread_shadow"]
    assert json.loads(Path(first["performance"]).read_text())["paper_bank_units"] == 100.0

    # Home wins by 20: the model's +10.5 underdog spread pick loses, so the
    # moneyline side at -10.5 covers and settles at the posted 1.909.
    settled = publish(
        tmp_path,
        Store([
            game(1, home_score=120, away_score=100),
            game(2, home_score=90, away_score=80),
            game(3),
        ]),
        report([]),
        run_id="results",
        generated_utc="2026-11-09T03:00:00Z",
    )
    shadow = settled["unpriced_spread_shadow"]
    tracked = next(r for r in shadow["candidate_records"] if r["posted_spread"])
    assert tracked["grade_status"] == "win"
    assert shadow["rolling"]["wins"] == 1
    assert shadow["rolling"]["losses"] == 0
    assert shadow["rolling"]["untracked_no_posted_spread"] == 1
    assert shadow["rolling"]["units_at_posted_price"] == pytest.approx(0.909)
    card = Path(settled["ticket_text"]).read_text()
    assert "1W-0L-0P (100.0% cover) · +0.91u at posted prices over 1 priced" in card
    assert "SPREAD full_game -10.5 -> HOME @1.909 Test Book · ML p 97.0% · ✅ WIN" in card
    assert json.loads(Path(settled["performance"]).read_text())["paper_bank_units"] == 100.0


def test_euroleague_consensus_price_paper_board_is_time_locked_and_has_no_ticket_bank_or_roi(tmp_path):
    rows = []
    for number, probability, price in ((1, 0.62, 1.80), (2, 0.59, 1.85), (3, 0.54, 1.70)):
        row = forecast(
            number,
            price=price,
            eligible=False,
            reasons=[
                "price_not_named_bookmaker_quote",
                "insufficient_walk_forward_samples",
            ],
            calibration_n=12,
        )
        row.update(
            {
                "league": League.EUROLEAGUE.value,
                "scheduled_utc": f"2026-11-07T{19 + number}:00:00Z",
                "calibrated_probability": probability,
            }
        )
        row["book"] = {
            **row["book"],
            "provider": "betexplorer",
            "book": "average",
            "observed_utc": "2026-11-07T06:00:00Z",
            "game_id": row["game_id"],
            "market": "moneyline",
            "period": "full_game",
            "line": None,
            "side_prices": {"home": price, "away": 2.0},
            "price_provenance": "aggregate_consensus",
            "source_permission": "unresolved",
            "operational_validation": "probation",
            "ticket_price_eligible": False,
        }
        rows.append(row)
    games = [
        game(
            number,
            league=League.EUROLEAGUE,
            scheduled_utc=row["scheduled_utc"],
        )
        for number, row in enumerate(rows, 1)
    ]

    first = publish_operator_reports(
        Store(games),
        League.EUROLEAGUE,
        report(rows),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T07:00:00Z",
        run_id="euroleague-consensus-freeze",
        slate_timezone_name="Europe/Paris",
    )

    board = first["euroleague_consensus_paper"]
    assert board is not None
    assert board["frozen_now"] is True
    assert board["candidate_count"] == 2
    assert board["no_stake_ticket_receipt_bank_roi_or_pnl"] is True
    roster_path = Path(board["roster"])
    roster = json.loads(roster_path.read_text())
    assert roster["schema_version"] == "magolide.euroleague_consensus_paper_roster"
    assert roster["decision_policy"] == "0900_sast_or_first_slate_tip_minus_60_minutes"
    assert len(roster["selections"]) == 2
    assert all("grade_status" not in row for row in roster["selections"])
    assert roster["selection_policy"]["no_ticket_receipt_bank_roi_or_pnl"] is True
    card = Path(board["card"]).read_text()
    assert "CONSENSUS-PRICE PAPER LEAGUE" in card
    assert "NO STAKE · NO TICKET · NO RECEIPT · NO PAPER BANK · NO ROI/P&L" in card
    assert "@1.800 average consensus" in card
    assert "strict-ticket blockers: price_not_named_bookmaker_quote" in card
    assert "TIME-LOCKED BOARD" in card
    assert first["ticket_receipt"] is None

    # A later scan cannot rotate the time-locked no-stake roster, even if its
    # aggregate prices or model probabilities now differ.
    changed = json.loads(json.dumps(rows))
    changed[0]["book"]["selected_decimal_price"] = 2.40
    changed[0]["book"]["side_prices"]["home"] = 2.40
    second = publish_operator_reports(
        Store(games),
        League.EUROLEAGUE,
        report(changed),
        tmp_path,
        target_date=DATE,
        generated_utc="2026-11-07T08:00:00Z",
        run_id="euroleague-consensus-rerun",
        slate_timezone_name="Europe/Paris",
    )
    assert second["euroleague_consensus_paper"]["frozen_now"] is False
    assert roster_path.read_text() == json.dumps(roster, indent=2, sort_keys=True) + "\n"
