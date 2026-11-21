"""End-to-end headless pipeline tests on committed fixtures (no network)."""

from __future__ import annotations

import json
import os

import pytest

from ingestion.books import BookConfig
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    Prediction,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import InMemoryStore

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def _nba_fetch(
    scoreboard="espn/scoreboard_synth_golden.json",
    predictor="espn/predictor_401909938.json",
    hardwood_path="hardwood/game_forecasts.json",
):
    mapping = {
        "scoreboard": scoreboard,
        "competitions/401909938/predictor": predictor,
        "summary?event=401909938": "espn/summary_401909938_context.json",
        "nba/injuries": "espn/injuries_sample.json",
        "game_forecasts": hardwood_path,
    }

    def fetch(url):
        for needle, relative in mapping.items():
            if needle in url:
                with open(os.path.join(FIXTURES, relative)) as handle:
                    return json.load(handle)
        from ingestion.http import HttpError

        raise HttpError(404, url, "no fixture mapped")

    return fetch


def _config(**changes):
    values = dict(
        league="NBA",
        date_yyyymmdd="20261107",
        book=BookConfig(provider="none"),
        fixture_now_utc="2026-11-07T18:00:00Z",
    )
    values.update(changes)
    return PipelineConfig(**values)


def _final_game() -> Game:
    return Game(
        game_id="401909938",
        league=League.NBA,
        season="2026",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[28.0, 27.0, 29.0, 26.0],
            away_quarters=[30.0, 31.0, 28.0, 31.0],
            home_final=110.0,
            away_final=120.0,
            overtime_periods=0,
            source="espn",
        ),
    )


def test_pregame_dry_run_returns_canonical_prediction_and_book_gaps():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch(), allow_writes=False)
    report = pipeline.pregame()

    assert report["fixtures_found"] == 1
    assert report["mode"] == "dry_run"
    assert report["persistence"]["system_of_record"] == "ephemeral"
    row = report["rows"][0]
    assert row["game_id"] == "401909938"
    assert row["prediction_provider"] == "espn"
    assert row["prediction"]["prob_home"] == pytest.approx(0.23062)
    assert row["prediction"]["prob_away"] == pytest.approx(0.76938)
    assert row["prediction"]["expected_total"] == pytest.approx(230.586)
    assert "prob_pct" not in row
    assert row["game_id"] == "401909938"
    assert row["full_game_line"] is None
    assert all(value is None for value in row["quarters"].values())
    assert row["book_evidence"]["full_game"] is None
    assert any(gap["required_field"].startswith("book_") for gap in report["gaps"])
    assert len(store.get_predictions("401909938")) == 1
    assert report["espn_context"]["nba_injuries"]["status"] == "ok"
    attached = store.get_game("401909938").evidence["espn"]["league_injuries"]
    assert attached["team_coverage_complete"] is True
    assert attached["teams"]["11"][0]["athlete_id"] == "player-1"


def test_pregame_defers_paid_books_but_attempts_the_direct_fallback_outside_ticket_window():
    store = InMemoryStore()
    pipeline = Pipeline(
        _config(acquire_book_lines=False), store, fetch=_nba_fetch(), allow_writes=False
    )
    report = pipeline.pregame()

    assert report["book_provider_status"]["policy"]["status"] == (
        "paid_sources_deferred_direct_fallback_attempted"
    )
    assert report["book_provider_status"]["betexplorer_direct"]["enabled"] is True
    assert any(
        gap["provider"] == "price_budget" and gap["required_field"] == "paid_named_book_price"
        for gap in report["gaps"]
    )


def test_pregame_uses_only_fresh_exact_book_evidence():
    store = InMemoryStore()
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="theoddsapi",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q1,
            observed_utc="2026-11-07T17:00:00Z",
            line=58.5,
            over_price_decimal=1.95,
            under_price_decimal=1.95,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="theoddsapi",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q2,
            observed_utc="2026-11-04T17:00:00Z",  # stale beyond 72h boundary
            line=999.0,
        )
    )
    config = _config(book=BookConfig(provider="theoddsapi", book="draftkings"))
    report = Pipeline(config, store, fetch=_nba_fetch()).pregame()
    row = report["rows"][0]
    assert row["quarters"]["Q1"] == 58.5
    assert row["book_evidence"]["Q1"]["provider"] == "theoddsapi"
    assert row["quarters"]["Q2"] is None
    assert row["book_evidence"]["Q2"] is None


def test_bounded_backfill_acquires_reconciled_history_without_learning_claim():
    store = InMemoryStore()
    calls = []

    def fetch(url):
        calls.append(url)
        assert "dates=20261106" in url
        with open(os.path.join(FIXTURES, "espn/scoreboard_20261105.json")) as handle:
            return json.load(handle)

    pipeline = Pipeline(_config(history_days=1, history_request_budget=1), store, fetch=fetch)
    report = pipeline.backfill()
    assert report["backfill"]["requested_days"] == 1
    assert report["backfill"]["learning_or_calibration_claimed"] is False
    assert report["backfill"]["dates"] == [
        {
            "date": "20261106",
            "status": "acquired",
            "finals_found": 2,
            "finals_reconciled": 2,
            "canonical_persistence": "dry_run_not_persisted",
        }
    ]
    assert report["finals_written"] == 2
    assert len(store.list_games(league=League.NBA)) == 2
    assert all(game.status is GameStatus.FINAL for game in store.list_games())
    assert len(calls) == 1
    assert report["acquisition"]["network_requests"] == 1


def test_backfill_budget_is_explicitly_bounded():
    store = InMemoryStore()
    calls = []

    def fetch(_url):
        calls.append(1)
        return {"events": []}

    report = Pipeline(
        _config(history_days=5, history_request_budget=2), store, fetch=fetch
    ).backfill()
    assert len(calls) == 2
    assert len(report["backfill"]["dates"]) == 2
    assert any(gap["required_field"] == "historical_backfill_budget" for gap in report["gaps"])


def test_pipeline_stage_gaps_do_not_accumulate_between_commands():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch())
    first = pipeline.pregame()
    first_gaps = len(first["gaps"])
    second = pipeline.pregame()
    assert first_gaps > 0
    assert len(second["gaps"]) == first_gaps
    coverage = pipeline.coverage()
    assert coverage["gaps"] == []


def test_freeze_then_results_uses_frozen_pre_tip_prediction_and_price():
    store = InMemoryStore()
    scheduled = _final_game()
    scheduled.status = GameStatus.SCHEDULED
    scheduled.actual = None
    store.upsert_game(scheduled)
    store.upsert_prediction(
        Prediction(
            game_id=scheduled.game_id,
            provider="espn",
            observed_utc="2026-11-07T18:00:00Z",
            prob_home=0.23062,
            prob_away=0.76938,
            expected_home_score=110.54,
            expected_away_score=120.046,
            expected_total=230.586,
            expected_margin=-9.506,
        )
    )
    store.upsert_prediction(
        Prediction(
            game_id=scheduled.game_id,
            provider="espn",
            observed_utc="2026-11-07T23:30:00Z",
            prob_home=0.5,
            prob_away=0.5,
            expected_home_score=100,
            expected_away_score=100,
            expected_total=200,
            expected_margin=0,
            pre_tip=False,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id=scheduled.game_id,
            provider="theoddsapi",
            book="draftkings",
            market=Market.MONEYLINE,
            period=Period.FULL_GAME,
            observed_utc="2026-11-07T18:00:00Z",
            home_price_decimal=4.0,
            away_price_decimal=1.25,
        )
    )

    class OfflinePipeline(Pipeline):
        def _fetch_finals(self):
            return [_final_game()]

    config = _config(
        fixture_now_utc="2026-11-08T00:00:00Z",
        book=BookConfig(provider="theoddsapi", book="draftkings"),
    )
    pipeline = OfflinePipeline(config, store)
    frozen = pipeline.freeze()
    assert frozen["book_snapshots_frozen"] == 1
    report = pipeline.results()
    assert report["finals_written"] == 1
    assert report["frozen_predictions_used"] == 1
    row = report["rows"][0]
    assert row["prediction"]["prob_away"] == 0.76938
    assert row["prediction_frozen"] is True
    assert row["home_final"] == 110
    assert row["away_final"] == 120
    assert row["odds"] == 1.25
    assert row["odds_evidence"]["book"] == "draftkings"


def test_results_quarantines_unreconciled_final():
    game = _final_game()
    game.actual.home_quarters = [28.0, 27.0, 29.0, 0.0]

    class OfflinePipeline(Pipeline):
        def _fetch_finals(self):
            return [game]

    report = OfflinePipeline(_config(), InMemoryStore()).results()
    assert report["finals_written"] == 0
    assert report["rows"] == []
    assert report["quarantined"][0]["reason"].startswith("sum")


def test_coverage_counts_only_reconciled_finals_and_real_quarter_lines():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch())
    pipeline.pregame()
    unreconciled = _final_game()
    unreconciled.game_id = "bad-final"
    unreconciled.actual.away_final = 999
    store.upsert_game(unreconciled)
    store.upsert_game(_final_game())
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="theoddsapi",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q1,
            observed_utc="2026-11-07T17:00:00Z",
            line=58.5,
        )
    )
    report = pipeline.coverage()
    assert report["games_in_store"] == 2
    assert report["games_with_prediction"] == 1
    assert report["games_with_full_game_line"] == 0
    assert report["quarter_lines"] == {"Q1": 1, "Q2": 0, "Q3": 0, "Q4": 0}
    assert report["finals_reconciled"] == 1


def _scoreboard_with_live_game():
    with open(os.path.join(FIXTURES, "espn/scoreboard_20261105.json")) as handle:
        final = json.load(handle)
    live = json.loads(json.dumps(final))
    status = live["events"][-1]["competitions"][0]["status"]
    status["type"] = {**status.get("type", {}), "state": "in", "completed": False}
    return live, final


def test_backfill_does_not_cache_yesterdays_scoreboard_while_a_game_is_live(tmp_path):
    """A late game still live at the first next-day run must settle on the next run."""
    from ingestion.providers.espn import scoreboard_is_settled

    live, final = _scoreboard_with_live_game()
    assert not scoreboard_is_settled(live)
    assert scoreboard_is_settled(final)
    served = [live, final]
    calls = []

    def fetch(url):
        assert "dates=20261106" in url
        calls.append(url)
        return served[min(len(calls), len(served)) - 1]

    def run(now_utc):
        store = InMemoryStore()
        pipeline = Pipeline(
            _config(
                history_days=1,
                history_request_budget=1,
                deep_history_event_budget=0,
                fixture_now_utc=now_utc,
            ),
            store,
            fetch=fetch,
        )
        # Production enables the durable cache only for warehouse-backed live
        # runs; point it at a temp dir to reproduce the cross-run behaviour.
        pipeline.acquisition.cache_dir = tmp_path / "http-cache" / "espn"
        report = pipeline.backfill()
        return store, report

    store, report = run("2026-11-07T06:11:00Z")
    assert report["backfill"]["dates"][0]["finals_found"] == 1
    assert report["acquisition"]["unsettled_cache_skips"] == 1
    assert not list((tmp_path / "http-cache" / "espn").glob("*.json"))

    store, report = run("2026-11-07T09:11:00Z")
    assert len(calls) == 2, "second run must re-read the scoreboard, not a cached live copy"
    assert report["backfill"]["dates"][0]["finals_found"] == 2
    assert all(game.status is GameStatus.FINAL for game in store.list_games())

    _, report = run("2026-11-07T12:11:00Z")
    assert len(calls) == 2, "a fully-final scoreboard is cached as before"
    assert report["acquisition"]["cache_hits"] == 1


# -- late finals for earlier days (every league) ------------------------------


def _late_game(league, game_id, tip_utc, status=GameStatus.SCHEDULED, final=False):
    return Game(
        game_id=game_id,
        league=league,
        season="2026",
        home_team_id=f"{game_id}-h",
        away_team_id=f"{game_id}-a",
        home_name=f"Home {game_id}",
        away_name=f"Away {game_id}",
        scheduled_utc=tip_utc,
        status=GameStatus.FINAL if final else status,
        actual=(
            ActualResult(
                period_kind=PeriodKind.QUARTERS,
                home_quarters=[18.0, 18.0, 18.0, 28.0],
                away_quarters=[28.0, 21.0, 12.0, 21.0],
                home_overtime=13.0,
                away_overtime=15.0,
                home_final=95.0,
                away_final=97.0,
                overtime_periods=1,
                source="official",
            )
            if final
            else None
        ),
    )


def _settle(game):
    return _late_game(game.league, game.game_id, game.scheduled_utc, final=True)


@pytest.mark.parametrize(
    "league, run_date, now_utc, late_tip_utc, local_date",
    [
        # EuroLeague (Paris day): tipped 19:30Z, still live at the last run of the day.
        ("EUROLEAGUE", "20261118", "2026-11-18T06:00:00Z", "2026-11-17T19:30:00Z", "20261117"),
        # NBA (US Eastern day): a 22:30 ET tip is 03:30Z the next UTC day.
        ("NBA", "20261118", "2026-11-18T16:00:00Z", "2026-11-18T03:30:00Z", "20261117"),
        # CBA (Beijing day).
        ("CBA", "20261118", "2026-11-18T06:00:00Z", "2026-11-17T11:35:00Z", "20261117"),
    ],
)
def test_results_settle_earlier_day_games_left_pending_for_every_league(
    monkeypatch, league, run_date, now_utc, late_tip_utc, local_date
):
    """A game not final at the last run of its day settles on a later run."""
    from ingestion.providers import cbaleague, espn, euroleague

    enum = League(league)
    store = InMemoryStore()
    late = _late_game(enum, "late-1", late_tip_utc)
    already_final = _late_game(enum, "done-1", late_tip_utc, final=True)
    store.upsert_game(late)
    store.upsert_game(already_final)
    calls = []

    def by_date(date, *args, **kwargs):
        calls.append(str(date))
        return [] if str(date) == run_date else [_settle(late), _settle(already_final)]

    def el_discover(date_iso, fetch, *args, **kwargs):
        calls.append(date_iso)
        if date_iso is not None:
            return []
        return [_late_game(enum, g.game_id, g.scheduled_utc, GameStatus.FINAL) for g in (late,)]

    monkeypatch.setattr(espn, "fetch_recent_results", by_date)
    monkeypatch.setattr(cbaleague, "fetch_recent_results", by_date)
    monkeypatch.setattr(euroleague, "discover_games", el_discover)
    monkeypatch.setattr(euroleague, "fetch_game_result", lambda game, fetch: _settle(game))

    config = _config(league=league, date_yyyymmdd=run_date, fixture_now_utc=now_utc)
    report = Pipeline(config, store, fetch=lambda url: {}).results()

    stored = store.get_game("late-1")
    assert stored.status is GameStatus.FINAL and stored.actual.home_final == 95.0
    assert [row["game_id"] for row in report["rows"]] == ["late-1"]
    iso = f"{local_date[:4]}-{local_date[4:6]}-{local_date[6:]}"
    assert report["rechecked_dates"] == [iso]
    if league == "EUROLEAGUE":
        assert calls == ["2026-11-18", None]  # today's slate, then one season request
    else:
        assert calls == [run_date, local_date]  # only the day with a pending game


def test_pending_recheck_skips_settled_closed_and_out_of_window_games(monkeypatch):
    from ingestion.providers import espn

    store = InMemoryStore()
    store.upsert_game(_late_game(League.NBA, "final", "2026-11-17T00:00:00Z", final=True))
    store.upsert_game(
        _late_game(League.NBA, "postponed", "2026-11-16T00:30:00Z", GameStatus.POSTPONED)
    )
    store.upsert_game(_late_game(League.NBA, "ancient", "2026-10-01T00:00:00Z"))
    store.upsert_game(_late_game(League.NBA, "not-yet", "2026-11-19T00:00:00Z"))
    calls = []
    monkeypatch.setattr(
        espn, "fetch_recent_results", lambda date, *a, **k: calls.append(date) or []
    )
    config = _config(date_yyyymmdd="20261118", fixture_now_utc="2026-11-18T16:00:00Z")
    report = Pipeline(config, store, fetch=lambda url: {}).results()
    assert calls == ["20261118"]
    assert report["rechecked_dates"] == []

    # A pending game is ignored entirely when the lookback is disabled.
    store.upsert_game(_late_game(League.NBA, "late", "2026-11-17T01:00:00Z"))
    calls.clear()
    Pipeline(_config(**{**config.__dict__, "history_days": 0}), store).results()
    assert calls == ["20261118"]


def test_pending_recheck_failure_on_one_day_does_not_block_other_days(monkeypatch):
    from ingestion.providers import espn

    store = InMemoryStore()
    early = _late_game(League.NBA, "early", "2026-11-16T01:00:00Z")  # ET 2026-11-15
    late = _late_game(League.NBA, "late", "2026-11-17T01:00:00Z")  # ET 2026-11-16
    store.upsert_game(early)
    store.upsert_game(late)

    def fetch_recent_results(date, *args, **kwargs):
        if date == "20261115":
            raise RuntimeError("provider down")
        return [_settle(late)] if date == "20261116" else []

    monkeypatch.setattr(espn, "fetch_recent_results", fetch_recent_results)
    config = _config(date_yyyymmdd="20261118", fixture_now_utc="2026-11-18T16:00:00Z")
    report = Pipeline(config, store, fetch=lambda url: {}).results()
    assert [row["game_id"] for row in report["rows"]] == ["late"]
    assert report["rechecked_dates"] == ["2026-11-16"]
    assert store.get_game("early").status is GameStatus.SCHEDULED
    assert any(gap["required_field"] == "pending_results_recheck" for gap in report["gaps"])
