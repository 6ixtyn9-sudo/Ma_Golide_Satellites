"""EuroCup lane: the validated EuroLeague adapter with competition code U.

EuroCup and EuroLeague are served by one host with an identical route shape;
the competition letter is carried by the season code (``U2026`` vs ``E2026``).
Two independent public clients document the ``U`` code, so this lane is a
*referenced* route contract — Ma Golide retains no first-party EuroCup capture.

The committed fixture ``tests/fixtures/eurocup/schedule_scores.json`` is a
structural transform of the live-captured EuroLeague schedule fixture
(2026-11-16): only the competition identity is rewritten (season code, season
name, competition code, identifier). Scores, periods, and game codes are the
captured ones, so reconciliation is exercised against real provider shapes.
"""

from __future__ import annotations

import json
from typing import Any, Callable, Dict

import pytest

from ingestion.models import GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import euroleague
from ingestion.store import InMemoryStore

FIXTURES = "tests/fixtures/"


def _fixture_json(relative: str) -> Any:
    with open(FIXTURES + relative, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _mapping_fetch(mapping: Dict[str, Callable[[], Any]]) -> Callable[[str], Any]:
    def fetch(url: str) -> Any:
        for needle, loader in mapping.items():
            if needle in url:
                return loader()
        raise AssertionError(f"no fixture mapped for URL: {url}")

    return fetch


def _eurocup_fetch() -> Callable[[str], Any]:
    return _mapping_fetch(
        {"seasons/U2026/games": lambda: _fixture_json("eurocup/schedule_scores.json")}
    )


def test_competition_codes_are_closed_and_derived_from_the_season_code():
    assert euroleague.competition_of("U2026") == "U"
    assert euroleague.competition_of("E2026") == "E"
    assert euroleague.season_code("U", 2026) == "U2026"
    for bad in ("2026", "X2026", "", "U26"):
        with pytest.raises(ValueError):
            euroleague.competition_of(bad)
    with pytest.raises(ValueError):
        euroleague.season_code("X", 2026)


def test_eurocup_discovery_uses_the_u_competition_path_and_u_game_ids():
    urls = []

    def fetch(url: str):
        urls.append(url)
        return _fixture_json("eurocup/schedule_scores.json")

    games = euroleague.discover_games(None, fetch, "U2026")
    assert urls and all("/competitions/U/seasons/U2026/games" in url for url in urls)
    assert games
    assert all(game.game_id.startswith("U2026-") for game in games)
    assert {game.season for game in games} == {"2026"}
    assert {game.provider_ids["euroleague_competition"] for game in games} == {"U"}


def test_eurocup_season_results_reconcile_the_captured_finals():
    failures: list = []
    results = euroleague.season_results("U2026", _eurocup_fetch(), failures)
    assert results
    assert failures == []
    for game in results:
        assert game.status is GameStatus.FINAL
        assert game.actual is not None
        ok, reason = game.actual.reconciles()
        assert ok, f"{game.game_id}: {reason}"
        assert game.game_id.startswith("U2026-")


def test_eurocup_result_fetch_targets_the_u_season_code():
    seen = []

    def fetch(url: str):
        seen.append(url)
        if "/competitions/U/seasons/U2026/games" in url:
            return _fixture_json("eurocup/schedule_scores.json")
        # Boxscore responses are not asserted on here beyond the season code.
        raise AssertionError(f"unexpected URL: {url}")

    games = [
        game
        for game in euroleague.discover_games(None, fetch, "U2026")
        if game.status is GameStatus.FINAL
    ]
    assert games
    with pytest.raises(Exception):
        euroleague.fetch_game_result(games[0], fetch)
    boxscore_urls = [url for url in seen if "Boxscore" in url]
    assert boxscore_urls
    assert all("seasoncode=U2026" in url for url in boxscore_urls)


def test_pipeline_eurocup_lane_discovers_the_target_date_only():
    store = InMemoryStore()
    config = PipelineConfig(league="EUROCUP", date_yyyymmdd="20261118")
    pipeline = Pipeline(config, store, fetch=_eurocup_fetch(), allow_writes=False)
    assert pipeline._euro_seasoncode() == "U2026"
    games = pipeline._discover_games()
    assert games
    assert all(game.game_id.startswith("U2026-") for game in games)
    assert all(game.scheduled_utc.startswith("2026-11-18") for game in games)


def test_pipeline_euro_season_code_follows_the_july_boundary():
    store = InMemoryStore()
    for league, date, expected in (
        ("EUROCUP", "20261118", "U2026"),
        ("EUROCUP", "20260520", "U2025"),
        ("EUROLEAGUE", "20261118", "E2026"),
        ("EUROLEAGUE", "20260520", "E2025"),
    ):
        pipeline = Pipeline(
            PipelineConfig(league=league, date_yyyymmdd=date), store, fetch=lambda url: {}
        )
        assert pipeline._euro_seasoncode() == expected


def test_eurocup_backfill_is_wired_to_the_shared_history_engine():
    store = InMemoryStore()
    config = PipelineConfig(league="EUROCUP", date_yyyymmdd="20261118")
    pipeline = Pipeline(config, store, fetch=_eurocup_fetch(), allow_writes=False)
    pipeline.config.history_request_budget = 1
    report = pipeline.backfill()
    deep = report["backfill"]["deep_history"]
    # A dry run must never advance a cursor over unpersisted history.
    assert deep["status"] == "disabled_canonical_write_required"


def test_eurocup_registry_contract_declares_only_the_fallback_price_lane():
    from ingestion.league_registry import league_spec

    spec = league_spec(League.EUROCUP)
    assert spec.schedule_source == "euroleague"
    assert spec.results_source == "euroleague"
    assert spec.automated_schedule_supported is True
    assert spec.automated_results_supported is True
    assert spec.book_providers == ("betexplorer_direct",)


def test_captured_u2026_overtime_row_reconciles():
    """One captured EuroCup row: Maxima Roma 86-84 London Lions after OT.

    ``partials1..4`` are points-in-period and ``extraPeriods`` carries the
    overtime by index; 76+10 and 76+8 reproduce the 86-84 finals exactly, which
    is why the row is safe to reconcile rather than merely display.
    """

    def fetch(url: str) -> Any:
        assert "competitions/U/seasons/U2026/games" in url, url
        return _fixture_json("eurocup/U2026_ot_game.json")

    games = euroleague.season_results("U2026", fetch)
    assert [game.game_id for game in games] == ["U2026-87"]
    actual = games[0].actual
    assert actual is not None
    assert actual.home_quarters == [16.0, 26.0, 21.0, 13.0]
    assert actual.away_quarters == [22.0, 22.0, 11.0, 21.0]
    assert (actual.home_overtime, actual.away_overtime) == (10.0, 8.0)
    assert (actual.home_final, actual.away_final) == (86, 84)
    assert actual.reconciles() == (True, "")
