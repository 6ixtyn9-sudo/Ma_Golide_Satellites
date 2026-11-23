"""ACB (Liga Endesa) lane: referenced JSON route contract, fixture-tested.

The ACB's JSON API is referenced from SportsDataverse's generated ``acb``
module (``https://api2.acb.com/api``). Ma Golide retains no first-party capture,
so the committed fixtures are *shaped* from the documented field contract rather
than being recorded responses: the calendar carries ids/dates/scores without
team names, and the per-game header carries names, status, final score and
``quarterScores[]`` (points in each period).

Cases covered: a clean regulation final, an overtime final, an incomplete
period body that must refuse reconciliation, a not-yet-played game, identity
failures, and the UTC/local timestamp trap.
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Callable, Dict

import pytest

from ingestion.models import GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import acb
from ingestion.store import InMemoryStore

FIXTURES = "tests/fixtures/"


def _fixture_json(relative: str) -> Any:
    with open(FIXTURES + relative, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _headers() -> Dict[str, Any]:
    return _fixture_json("acb/match_headers_2027.json")


def _acb_fetch(headers: Dict[str, Any] | None = None) -> Callable[[str], Any]:
    table = headers if headers is not None else _headers()

    def fetch(url: str) -> Any:
        if "season-calendar" in url:
            return _fixture_json("acb/calendar_2027.json")
        if "match-header" in url:
            match_id = url.rsplit("matchId=", 1)[-1].split("&")[0]
            if match_id not in table:
                raise AssertionError(f"no header fixture for matchId={match_id}")
            return table[match_id]
        raise AssertionError(f"unexpected URL: {url}")

    return fetch


def test_season_end_year_and_edition_id_follow_the_documented_naming():
    assert acb.season_end_year(datetime(2026, 11, 18)) == 2027
    assert acb.season_end_year(datetime(2027, 5, 20)) == 2027
    assert acb.season_id_for(2027) == 91
    with pytest.raises(ValueError):
        acb.season_id_for(1983)


def test_discover_games_resolves_identity_from_the_header_for_one_utc_date():
    games = acb.discover_games("2026-11-18", _acb_fetch(), 2027)
    assert acb.last_gaps() == []
    assert {game.game_id for game in games} == {"ACB-105401", "ACB-105402"}
    assert all(game.league is League.SPAIN_ACB for game in games)
    assert all(game.season == "2027" for game in games)
    by_id = {game.game_id: game for game in games}
    assert by_id["ACB-105401"].home_name == "Real Madrid"
    assert by_id["ACB-105401"].away_name == "FC Barcelona"
    assert by_id["ACB-105401"].status is GameStatus.FINAL
    assert all(game.scheduled_utc.startswith("2026-11-18") for game in games)


def test_regulation_final_reconciles_from_quarter_scores():
    game = {
        game.game_id: game for game in acb.discover_games("2026-11-18", _acb_fetch(), 2027)
    }["ACB-105401"]
    actual = game.actual
    assert actual is not None
    assert actual.home_quarters == [22.0, 21.0, 20.0, 25.0]
    assert actual.away_quarters == [20.0, 22.0, 18.0, 20.0]
    assert actual.overtime_periods == 0
    ok, reason = actual.reconciles()
    assert ok, reason


def test_overtime_final_sums_extra_periods_and_still_reconciles():
    game = {
        game.game_id: game for game in acb.discover_games("2026-11-18", _acb_fetch(), 2027)
    }["ACB-105402"]
    actual = game.actual
    assert actual is not None
    assert actual.overtime_periods == 1
    assert (actual.home_overtime, actual.away_overtime) == (11.0, 9.0)
    ok, reason = actual.reconciles()
    assert ok, reason


def test_incomplete_period_body_is_recorded_not_invented():
    failures: list = []
    results = acb.season_results(2027, _acb_fetch(), failures)
    assert {game.game_id for game in results} == {"ACB-105401", "ACB-105402"}
    assert [row["game_id"] for row in failures] == ["ACB-105403"]
    assert "regulation quarters" in failures[0]["error"]


def test_identity_failure_skips_a_candidate_and_records_a_gap():
    headers = _headers()
    broken = json.loads(json.dumps(headers))
    broken["105402"]["teams"]["away"]["fullName"] = broken["105402"]["teams"]["home"]["fullName"]
    games = acb.discover_games("2026-11-18", _acb_fetch(broken), 2027)
    assert {game.game_id for game in games} == {"ACB-105401"}
    assert [gap["game_id"] for gap in acb.last_gaps()] == ["105402"]


def test_naive_provider_timestamp_is_refused():
    headers = _headers()
    naive = json.loads(json.dumps(headers))
    naive["105401"]["start"] = "2026-11-18T19:30:00"
    games = acb.discover_games("2026-11-18", _acb_fetch(naive), 2027)
    assert {game.game_id for game in games} == {"ACB-105402"}
    assert any("no UTC offset" in gap["detail"] for gap in acb.last_gaps())


def test_empty_calendar_is_refused_rather_than_called_a_scoreless_season():
    def fetch(url: str):
        return {"rounds": []}

    with pytest.raises(ValueError, match="no rounds"):
        acb.fetch_calendar(2027, fetch)


def test_fetch_game_result_refreshes_from_the_header_and_checks_identity():
    games = acb.discover_games("2026-11-18", _acb_fetch(), 2027)
    refreshed = acb.fetch_game_result(games[0], _acb_fetch())
    assert refreshed.game_id == games[0].game_id
    assert refreshed.actual is not None

    headers = _headers()
    renamed = json.loads(json.dumps(headers))
    renamed["105401"]["teams"]["home"]["fullName"] = "Some Other Club"
    with pytest.raises(Exception):
        acb.fetch_game_result(games[0], _acb_fetch(renamed))


def test_smoke_reports_reconciled_counts_without_raising():
    report = acb.smoke(_acb_fetch(), 2027)
    assert report["games_found"] == 4
    assert report["results_reconciled"] == {"complete": 2, "total": 3}
    assert [row["game_id"] for row in report["failures"]] == ["ACB-105403"]

    def broken(url: str):
        raise RuntimeError("boom")

    failed = acb.smoke(broken, 2027)
    assert failed["status"] == "provider_error"


def test_pipeline_acb_lane_discovers_and_settles_a_date():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="SPAIN_ACB", date_yyyymmdd="20261118"),
        store,
        fetch=_acb_fetch(),
        allow_writes=False,
    )
    games = pipeline._discover_games()
    assert {game.game_id for game in games} == {"ACB-105401", "ACB-105402"}
    finals = pipeline._acb_finals(games)
    assert {game.game_id for game in finals} == {"ACB-105401", "ACB-105402"}
    assert all(game.actual is not None for game in finals)


def test_acb_registry_contract_declares_the_fallback_price_lane():
    from ingestion.league_registry import league_spec

    spec = league_spec(League.SPAIN_ACB)
    assert spec.schedule_source == "acb"
    assert spec.results_source == "acb"
    assert spec.automated_schedule_supported is True
    # The direct comparison page remains the named-book fallback; the
    # competition's own fixtures page now also feeds aggregate, research-only
    # observations (verified 2026-11-23). Neither is a cleared API lane.
    assert spec.book_providers == ("betexplorer", "betexplorer_direct")


def test_gated_match_header_is_reported_as_auth_required_not_an_empty_season():
    from ingestion.http import HttpError

    def gated(url: str):
        if "season-calendar" in url:
            return _fixture_json("acb/calendar_2027.json")
        raise HttpError(401, url, "Unauthorized")

    games = acb.discover_games("2026-11-18", gated, 2027)
    assert games == []
    gaps = acb.last_gaps()
    assert len(gaps) == 1  # one clear reason, not one per candidate
    assert "requires authorisation" in gaps[0]["detail"]

    report = acb.smoke(gated, 2027)
    assert report["status"] == "auth_required"
    assert "401/403" in report["detail"]
    assert len(report["failures"]) == 1


def test_acb_sends_the_public_bundle_key_and_honours_the_env_override(monkeypatch):
    """The host gates the match header; the header is a published site value.

    The lane read zero games for 728 runs because the calendar was public while
    the per-game header — the only route carrying team names and period scores —
    answered 401/403 to an unauthenticated client.
    """

    monkeypatch.delenv(acb.KEY_ENV, raising=False)
    assert acb.request_headers()["X-APIKEY"] == acb.DEFAULT_KEY
    monkeypatch.setenv(acb.KEY_ENV, "operator-supplied")
    assert acb.request_headers()["X-APIKEY"] == "operator-supplied"


def test_acb_falls_back_to_its_own_keyed_transport(monkeypatch):
    """With no injected fetch the adapter must not fall back to a bare client."""

    monkeypatch.setattr(acb, "default_fetch", lambda: _acb_fetch())

    # One date only: the fixture calendar deliberately carries a match whose
    # header is absent (105404), and that case is covered separately.
    games = acb.discover_games("2026-11-18", None, 2027)

    assert games
    assert all(game.league is League.SPAIN_ACB for game in games)
    probe = acb.last_probe()
    assert "season-calendar" in probe["url"]
    assert probe["rows"] >= len(games)
    assert len(probe["sha256"]) == 64
