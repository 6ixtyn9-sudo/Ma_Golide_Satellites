"""Settled totals-only finals must be repairable, or they starve the model.

KBL's list route settles a game without a period split, so the stored final is
``PeriodKind.UNKNOWN`` and ``ActualResult.reconciles()`` rejects it. The on-date
results stage repairs *today's* slate only; every earlier settled row therefore
stays unreconcilable forever, the forecast corpus never reaches
``min_history_games = 20``, and the lane emits no forecast at all -- which reads
on the card exactly like a league with no games.

These tests cover the bounded repair pass that closes that gap. The detail
payload is the committed first-party capture
``tests/fixtures/kbl/captures/match__gmkey__S47G01N49.json`` (gmkey
``S47G01N49``, `isEnded: 1`, real quarter split and real team codes
``tcodeH=66`` / ``tcodeA=50``). One captured body is re-served under several
gmkeys; only that key is substituted, no score is invented.
"""

import copy
import json
from pathlib import Path

from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import InMemoryStore

CAPTURE = json.loads(
    (
        Path(__file__).parent / "fixtures" / "kbl" / "captures" / "match__gmkey__S47G01N49.json"
    ).read_text()
)
HOME_CODE, AWAY_CODE = "66", "50"


def _detail(gmkey: str) -> dict:
    body = copy.deepcopy(CAPTURE)
    body["game"]["gmkey"] = gmkey
    return body


def _totals_only(game_id: str, gmkey: str, scheduled_utc: str) -> Game:
    """A settled row as discovery stores it: final score, no period split."""

    return Game(
        game_id=game_id,
        league=League.SOUTH_KOREA_KBL,
        season="2026",
        home_team_id=HOME_CODE,
        away_team_id=AWAY_CODE,
        home_name="Goyang Sono",
        away_name="Changwon LG",
        scheduled_utc=scheduled_utc,
        status=GameStatus.FINAL,
        provider_ids={"kbl": gmkey},
        actual=ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=74.0,
            away_final=81.0,
            source="kbl match list",
        ),
    )


def _pipeline(store, fetch, *, budget=25, allow_writes=True, date="20261124"):
    config = PipelineConfig(
        league="SOUTH_KOREA_KBL",
        date_yyyymmdd=date,
        history_repair_budget=budget,
    )
    return Pipeline(config, store, fetch=fetch, allow_writes=allow_writes)


def _counting_fetch(calls):
    def fetch(url: str):
        calls.append(url)
        gmkey = url.rstrip("/").rsplit("/", 1)[-1]
        return _detail(gmkey)

    return fetch


def test_a_totals_only_final_is_repaired_and_enters_the_corpus():
    store = InMemoryStore()
    store.upsert_game(_totals_only("kbl-1", "S47G01N49", "2026-11-20T10:00:00Z"))
    calls = []

    summary = _pipeline(store, _counting_fetch(calls))._repair_unreconciled_history()

    assert summary["status"] == "ok"
    assert summary["candidates"] == 1
    assert summary["requests"] == 1
    assert summary["repaired"] == 1
    repaired = store.get_game("kbl-1")
    ok, reason = repaired.actual.reconciles()
    assert ok, reason
    assert repaired.actual.period_kind is PeriodKind.QUARTERS
    # The corpus now sees the game: the exact gate the model applies.
    assert len([g for g in store.list_games(league=League.SOUTH_KOREA_KBL) if g.actual.reconciles()[0]]) == 1
    assert calls and calls[0].endswith("S47G01N49")


def test_repair_is_bounded_and_newest_first():
    store = InMemoryStore()
    for index in range(5):
        store.upsert_game(
            _totals_only(
                f"kbl-{index}",
                f"S47G01N{index:02d}",
                f"2026-11-{15 + index:02d}T10:00:00Z",
            )
        )
    calls = []

    summary = _pipeline(store, _counting_fetch(calls), budget=2)._repair_unreconciled_history()

    assert summary["candidates"] == 5
    assert summary["requests"] == 2, "budget is a request cap, not a target"
    assert len(calls) == 2
    newest = sorted(
        (
            g for g in store.list_games(league=League.SOUTH_KOREA_KBL) if g.actual.reconciles()[0]
        ),
        key=lambda g: g.scheduled_utc,
    )
    assert newest[-1].game_id == "kbl-4", "the two newest rows were the ones repaired"


def test_already_reconcilable_rows_are_never_refetched():
    store = InMemoryStore()
    good = _totals_only("kbl-ok", "S47G01N49", "2026-11-20T10:00:00Z")
    good.actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[21.0, 24.0, 18.0, 11.0],
        away_quarters=[25.0, 20.0, 20.0, 16.0],
        home_final=74.0,
        away_final=81.0,
        source="kbl match",
    )
    store.upsert_game(good)
    calls = []

    summary = _pipeline(store, _counting_fetch(calls))._repair_unreconciled_history()

    assert summary["candidates"] == 0
    assert summary["requests"] == 0
    assert calls == []


def test_a_failed_detail_read_is_a_named_gap_not_an_upgrade():
    store = InMemoryStore()
    store.upsert_game(_totals_only("kbl-1", "S47G01N49", "2026-11-20T10:00:00Z"))

    def broken(url: str):
        raise ValueError("KBL match S47G01N49 is not ended; refusing a result")

    pipeline = _pipeline(store, broken)
    summary = pipeline._repair_unreconciled_history()

    assert summary["failed"] == 1
    assert summary["repaired"] == 0
    gap = next(g for g in pipeline.gaps if g.required_field == "historical_reconciliation")
    assert gap.provider == "kbl"
    assert gap.reason is not None and gap.reason.value == "provider_error"
    # The row is untouched: a totals-only final is not silently promoted.
    assert store.get_game("kbl-1").actual.period_kind is PeriodKind.UNKNOWN


def test_a_body_for_another_game_is_refused():
    store = InMemoryStore()
    # Stored as gmkey S47G01N49; the detail body answers that key, so a game
    # whose stored key differs must be refused by the adapter's identity check.
    game = _totals_only("kbl-1", "S47G01N49", "2026-11-20T10:00:00Z")
    game.home_team_id = "99"
    store.upsert_game(game)

    pipeline = _pipeline(store, _counting_fetch([]))
    summary = pipeline._repair_unreconciled_history()

    assert summary["failed"] == 1
    assert store.get_game("kbl-1").actual.period_kind is PeriodKind.UNKNOWN


def test_a_dry_run_spends_no_requests():
    store = InMemoryStore()
    store.upsert_game(_totals_only("kbl-1", "S47G01N49", "2026-11-20T10:00:00Z"))
    calls = []

    summary = _pipeline(
        store, _counting_fetch(calls), allow_writes=False
    )._repair_unreconciled_history()

    assert summary["status"] == "disabled_canonical_write_required"
    assert calls == []
    assert store.get_game("kbl-1").actual.period_kind is PeriodKind.UNKNOWN


def test_a_lane_with_reconcilable_results_does_no_work():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd="20261124"),
        store,
        fetch=lambda url: {},
        allow_writes=True,
    )

    summary = pipeline._repair_unreconciled_history()

    assert summary["status"] == "unsupported"
    assert summary["requests"] == 0


def test_pregame_runs_the_repair_and_reports_it(monkeypatch):
    """The pass is wired into the run, not just available to call."""

    store = InMemoryStore()
    store.upsert_game(_totals_only("kbl-1", "S47G01N49", "2026-11-20T10:00:00Z"))
    pipeline = _pipeline(store, _counting_fetch([]))
    # Discovery is not what is under test here; the repair runs before it.
    monkeypatch.setattr(Pipeline, "_discover_games", lambda self: [])

    report = pipeline.pregame()

    assert report["history_repair"]["status"] == "ok"
    assert report["history_repair"]["repaired"] == 1
    assert store.get_game("kbl-1").actual.reconciles()[0]
