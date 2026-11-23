"""A late West-Coast game settles on the next run, not on its ticket day.

Golden State 88-87 Boston (ESPN 401910035) tipped at 2026-11-23T01:30Z, i.e.
20:30 on 2026-11-22 in the league's own timezone. A ticket frozen for 11-22
therefore holds a leg whose *ledger* date is 11-22 while its clock time is on
the 23rd. Result acquisition is keyed by the ledger date, so the run that
frozen the ticket cannot have the final; the next run has to look back.

These tests pin that contract: the recheck queries the game's own ledger date
(not the run's), and a same-day run leaves the game to the current-day results
path instead of silently double-fetching it.
"""

from __future__ import annotations

from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind
from ingestion.pipeline import Pipeline, PipelineConfig, espn
from ingestion.store import InMemoryStore

TIP_UTC = "2026-11-23T01:30:00Z"


def _game(status: GameStatus = GameStatus.SCHEDULED, actual: ActualResult | None = None) -> Game:
    return Game(
        game_id="401910035",
        league=League.NBA,
        season="2026",
        home_team_id="9",
        away_team_id="2",
        home_name="Golden State Warriors",
        away_name="Boston Celtics",
        scheduled_utc=TIP_UTC,
        status=status,
        actual=actual,
        provider_ids={"espn": "401910035"},
    )


def _final() -> Game:
    return _game(
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[27.0, 17.0, 25.0, 19.0],
            away_quarters=[20.0, 22.0, 33.0, 12.0],
            home_final=88.0,
            away_final=87.0,
            source="espn summary",
        ),
    )


def _pipeline(store: InMemoryStore, date_yyyymmdd: str, now_utc: str) -> Pipeline:
    return Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd=date_yyyymmdd,
            fixture_now_utc=now_utc,
        ),
        store,
    )


def test_next_day_run_requeries_the_games_own_ledger_date(monkeypatch):
    store = InMemoryStore()
    store.upsert_game(_game())
    asked: list[str] = []

    def fetch_recent_results(date, fetch, league, observed_utc=None):
        asked.append(date)
        return [_final()]

    monkeypatch.setattr(espn, "fetch_recent_results", fetch_recent_results)
    pipeline = _pipeline(store, "20261123", "2026-11-23T08:00:00Z")

    finals = pipeline._recheck_pending_finals(set())

    # 2026-11-23T01:30Z is 20:30 on 11-22 in New York, so the lookback must ask
    # for 11-22 — asking for the run's own date would miss the game forever.
    assert asked == ["20261122"]
    assert [game.game_id for game in finals] == ["401910035"]
    assert finals[0].status is GameStatus.FINAL
    assert finals[0].actual is not None
    assert finals[0].actual.reconciles() == (True, "")


def test_same_day_run_leaves_the_game_to_the_current_day_results_path(monkeypatch):
    store = InMemoryStore()
    store.upsert_game(_game())
    called = False

    def fetch_recent_results(*args, **kwargs):  # pragma: no cover - must not run
        nonlocal called
        called = True
        return []

    monkeypatch.setattr(espn, "fetch_recent_results", fetch_recent_results)
    pipeline = _pipeline(store, "20261122", "2026-11-22T18:00:00Z")

    assert pipeline._recheck_pending_finals(set()) == []
    assert called is False


def test_a_failed_recheck_records_a_gap_instead_of_inventing_a_result(monkeypatch):
    """A stored FINAL with no usable periods stays pending when the fetch fails."""

    store = InMemoryStore()
    store.upsert_game(_game(status=GameStatus.FINAL, actual=None))

    def boom(*args, **kwargs):
        raise RuntimeError("provider down")

    monkeypatch.setattr(espn, "fetch_recent_results", boom)
    pipeline = _pipeline(store, "20261123", "2026-11-23T08:00:00Z")

    assert pipeline._recheck_pending_finals(set()) == []
    gaps = [gap for gap in pipeline.gaps if gap.required_field == "pending_results_recheck"]
    assert gaps and "20261122" in gaps[0].detail
    # The stored game keeps its unreconciled state; nothing was graded from air.
    stored = store.list_games(league=League.NBA)[0]
    assert stored.status is GameStatus.FINAL and stored.actual is None
