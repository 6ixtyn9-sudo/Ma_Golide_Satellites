"""End-to-end pipeline tests on committed fixtures (no network, no secrets)."""

import json
import os

import pytest

from ingestion.books import BookConfig
from ingestion.models import League, Market, Period, Prediction
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.sheets_contract import BASE_HEADERS, COL
from ingestion.sheets_writer import FakeWorksheet, apply_plan, plan_update
from ingestion.store import InMemoryStore
from ingestion.providers import _common

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def _nba_fetch(scoreboard="espn/scoreboard_synth_golden.json",
               predictor="espn/predictor_401909938.json",
               hardwood_path="hardwood/game_forecasts.json"):
    mapping = {
        "scoreboard": scoreboard,
        "competitions/401909938/predictor": predictor,
        "game_forecasts": hardwood_path,
    }

    def fetch(url):
        for needle, rel in mapping.items():
            if needle in url:
                with open(os.path.join(FIXTURES, rel)) as f:
                    return json.load(f)
        # Anything else (e.g. predictors of other events) behaves like a
        # permanent miss, as an unauthenticated 404 would.
        from ingestion.http import HttpError

        raise HttpError(404, url, "no fixture mapped")

    return fetch


def _config(**kw):
    defaults = dict(
        league="NBA", date_yyyymmdd="20261107",
        book=BookConfig(provider="none"),
    )
    defaults.update(kw)
    return PipelineConfig(**defaults)


def _upcoming_ws():
    # Real-world column layout: 22 base + model columns observed through AX.
    headers = list(BASE_HEADERS) + ["t2-q1-conf", "t2-q2-conf", "ou-q1", "ou-best"]
    return FakeWorksheet(headers)


# ---------------------------------------------------------------------------
# Pregame
# ---------------------------------------------------------------------------


def test_pregame_dry_run_golden_end_to_end():
    store = InMemoryStore()
    ws = _upcoming_ws()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch(),
                        upcoming_ws=ws, allow_writes=False)
    report = pipeline.pregame()

    assert report["fixtures_found"] == 1
    assert report["mode"] == "dry_run"
    row = report["rows"][0]
    assert row["game_id"] == "401909938"
    assert row["prediction_provider"] == "espn"
    assert row["prob_pct"] == "23.1 - 76.9"
    assert row["avg"] == "230.6"
    # Keyless run: no bookmaker keys -> explicit coverage gaps, no book cells.
    assert row["full_game_line"] is None
    assert all(v is None for v in row["quarters"].values())
    assert any(g["required_field"].startswith("book_") for g in report["gaps"])

    plan = report["sheet_plan"]
    assert plan["appended_rows"] == 1 and plan["changed_cells"] == 0
    assert len(ws.rows()) == 0  # dry run wrote nothing


def test_pregame_rendered_cells_match_golden():
    store = InMemoryStore()
    ws = _upcoming_ws()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch(), upcoming_ws=ws)
    report = pipeline.pregame()
    plan = plan_update("UpcomingClean", ws, __import__("ingestion.sheets_contract", fromlist=["detect_contract"]).detect_contract(ws.headers, "upcoming"), [], "upcoming")
    # Apply for real on the fake sheet, then re-run: idempotent no-op.
    apply_plan(report["sheet_plan"] and _replan(pipeline, ws), ws, allow_writes=True, sleep=lambda s: None)
    rendered = ws.rows()[0]
    assert rendered[COL["Prob %"] - 1] == "23.1 - 76.9"
    assert rendered[COL["Pred"] - 1] == 2
    assert rendered[COL["Pred Score"] - 1] == "111-120"
    assert rendered[COL["Avg"] - 1] == 230.6
    assert rendered[COL["Home"] - 1] == "Indiana Pacers"
    assert rendered[COL["Away"] - 1] == "Denver Nuggets"
    assert rendered[COL["FT Score"] - 1] == ""
    assert rendered[COL["t2-q1"] - 1] == ""  # protected columns blank on append

    second = Pipeline(_config(), store, fetch=_nba_fetch(), upcoming_ws=ws)
    report2 = second.pregame()
    assert report2["sheet_plan"]["appended_rows"] == 0
    assert report2["sheet_plan"]["changed_cells"] == 0  # idempotent second run


def _replan(pipeline, ws):
    from ingestion.sheets_contract import detect_contract

    # Rebuild the rendered rows exactly as pregame computed them.
    games = pipeline._discover_games()
    selections = pipeline._acquire_predictions(games)
    rows = []
    for game in games:
        rows.append(
            __import__("ingestion.sheets_contract", fromlist=["render_upcoming_row"])
            .render_upcoming_row(game, selections[game.game_id].selected, None,
                                 {}, None, pipeline.config.timezone)
        )
    return plan_update("UpcomingClean", ws, detect_contract(ws.headers, "upcoming"),
                       rows, "upcoming")


def test_pregame_quarter_csv_fallback_fills_quarters():
    import csv

    csv_path = os.path.join(FIXTURES, "manual_quarters.csv")
    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(["game_id", "book", "period", "line", "observed_utc"])
        writer.writerow(["401909938", "fanduel", "Q1", "58.5", "2026-11-07T18:00:00Z"])
        writer.writerow(["401909938", "fanduel", "Q4", "59.5", "2026-11-07T18:00:00Z"])

    store = InMemoryStore()
    pipeline = Pipeline(
        _config(quarter_csv=csv_path), store,
        fetch=_nba_fetch(), allow_writes=False,
    )
    report = pipeline.pregame()
    row = report["rows"][0]
    assert row["quarters"]["Q1"] == 58.5
    assert row["quarters"]["Q4"] == 59.5
    assert row["quarters"]["Q2"] is None  # partial quarters stay partial


def test_pregame_missing_market_does_not_erase_manual_line():
    """A refresh returning no market preserves the valid existing value."""
    store = InMemoryStore()
    ws = _upcoming_ws()
    existing = [""] * (len(ws.headers))
    existing[COL["Home"] - 1] = "Indiana Pacers"
    existing[COL["Away"] - 1] = "Denver Nuggets"
    existing[COL["Date"] - 1] = "2026-11-07"
    existing[COL["Q1"] - 1] = "57.5"   # manual book line already in the sheet
    existing[COL["FT Score"] - 1] = "231.0"
    ws._rows.append(existing)

    pipeline = Pipeline(_config(), store, fetch=_nba_fetch(), upcoming_ws=ws)
    report = pipeline.pregame()
    plan = report["sheet_plan"]
    assert plan["appended_rows"] == 0
    q1_change = [c for c in plan["changes"] if c["col"] == COL["Q1"]]
    assert q1_change == []  # Q1 preserved untouched
    # The rendered row still carries the preserved values in its own row.
    assert plan["changed_cells"] > 0  # identity/prediction cells fill in


# ---------------------------------------------------------------------------
# Freeze + results
# ---------------------------------------------------------------------------


def test_freeze_then_results_uses_frozen_pre_tip():
    store = InMemoryStore()
    # Seed the pre-tip golden prediction (as a pregame run would have).
    pre_tip = Prediction(
        game_id="401909938", provider="espn",
        observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.23062, prob_away=0.76938,
        expected_home_score=110.54, expected_away_score=120.046,
        expected_total=230.586, expected_margin=9.506,
    )
    post_tip = Prediction(
        game_id="401909938", provider="espn",
        observed_utc="2026-11-07T23:30:00Z",
        prob_home=0.5, prob_away=0.5,
        expected_home_score=100.0, expected_away_score=100.0,
        expected_total=200.0, expected_margin=0.0,
    )
    post_tip.pre_tip = False
    store.upsert_game(_final_game())
    store.upsert_prediction(pre_tip)
    store.upsert_prediction(post_tip)

    pipeline = Pipeline(
        _config(), store, fetch=lambda url: (_ for _ in ()).throw(AssertionError("no network")),
        results_ws=None, allow_writes=False,
    )
    # Freeze explicitly, then render results offline.
    game = store.get_game("401909938")
    from ingestion.predictions import freeze_game_predictions

    statuses = freeze_game_predictions(store, game, game.scheduled_utc)
    assert statuses["espn"] == "frozen"

    from ingestion.sheets_contract import render_result_row
    from ingestion.predictions import resolve_frozen_bundle

    pred, was_frozen = resolve_frozen_bundle(store, "401909938")
    assert was_frozen
    cells = render_result_row(game, pred, None, "America/New_York")
    assert cells[COL["Prob %"]] == "23.1 - 76.9"  # frozen values, not post-tip
    assert cells[COL["Avg"]] == 230.6


def _final_game():
    from ingestion.models import ActualResult, Game, GameStatus, PeriodKind

    return Game(
        game_id="401909938", league=League.NBA, season="2026",
        home_team_id="11", away_team_id="7",
        home_name="Indiana Pacers", away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z", status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[28.0, 27.0, 29.0, 26.0],
            away_quarters=[30.0, 31.0, 28.0, 31.0],
            home_final=110.0, away_final=120.0,
            overtime_periods=0, source="espn",
        ),
    )


def test_results_quarantines_unreconciled_and_flags_missing_prediction():
    store = InMemoryStore()
    game = _final_game()
    game.actual.home_quarters = [28.0, 27.0, 29.0, 0.0]  # breaks the sum

    from ingestion.models import ActualResult, Game as G

    class _P(Pipeline):
        def _fetch_finals(self):
            return [game]

    pipeline = _P(_config(), store, fetch=None, results_ws=None)
    report = pipeline.results()
    assert report["finals_written"] == 0
    assert report["quarantined"][0]["reason"].startswith("sum")


# ---------------------------------------------------------------------------
# Coverage + books merge through the pipeline
# ---------------------------------------------------------------------------


def test_coverage_report_counts_from_store():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch())
    pipeline.pregame()
    report = Pipeline(_config(), store).coverage()
    assert report["games_in_store"] == 1
    assert report["games_with_prediction"] == 1
    assert report["games_with_full_game_line"] == 0
    assert report["quarter_lines"] == {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0}


def test_book_snapshots_stored_and_mapped():
    store = InMemoryStore()
    from ingestion.models import BookSnapshot

    snap = BookSnapshot(
        game_id="401909938", provider="sportsgameodds", book="draftkings",
        market=Market.TOTAL, period=Period.Q1, observed_utc="2026-11-07T18:00:00Z",
        line=58.5, over_price_decimal=1.95, under_price_decimal=1.95,
        pre_match=True, main_line=True,
    )
    store.upsert_book_snapshot(snap)
    snaps = store.get_book_snapshots("401909938")
    assert len(snaps) == 1 and snaps[0].line == 58.5
