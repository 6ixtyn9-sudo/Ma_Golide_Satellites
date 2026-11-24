"""An unreachable schedule host must cost a lane, not the whole day.

Live evidence (run 45357690260, 2026-11-24T08:39Z): the EuroLeague tick died at
discovery with

    HTTP 0 for https://api-live.euroleague.net/v2/competitions/E/seasons/E2026/games:
    invalid JSON: Expecting value: line 1 column 1 (char 0)

``HTTP 0`` means no response at all -- there was no payload to parse.  The
adapter raised, ``_discover_games`` had no guard on that branch, and the
exception unwound to ``scripts/daily.py``, which can only print
``{"status": "failed"}`` and exit 1.  A warm warehouse (6,737 games) and six live
predictions were therefore reported as a failed, empty day.

The contract asserted here: a dead provider becomes a *named* gap plus a
``provider_error`` health row, the run continues on the canonical stored slate,
and an unreadable source is never confused with "no games today".
"""

from pathlib import Path

import pytest

from ingestion.http import HttpError
from ingestion.models import CoverageState, Game, GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import euroleague
from ingestion.store import InMemoryStore

FIXTURES = Path(__file__).parent / "fixtures" / "euroleague"
SEASONCODE = "E2026"


def _dead_fetch(url: str):
    """Reproduce the incident's failure class: no response, so no body."""

    raise HttpError(0, url, "invalid JSON: Expecting value: line 1 column 1 (char 0)")


def _healthy_fetch(url: str):
    import json

    if "/seasons/" in url and url.endswith("/games"):
        return json.loads((FIXTURES / "schedule.json").read_text())
    raise KeyError(f"no fixture mapped for URL: {url}")


def _pipeline(fetch=_dead_fetch, date="20261123", store=None, league="EUROLEAGUE"):
    return Pipeline(
        PipelineConfig(league=league, date_yyyymmdd=date),
        store or InMemoryStore(),
        fetch=fetch,
        allow_writes=False,
    )


def _schedule_gap(pipeline: Pipeline):
    return next(
        (
            gap
            for gap in pipeline.gaps
            if gap.required_field == "official_schedule"
            and gap.reason is CoverageState.PROVIDER_ERROR
        ),
        None,
    )


def _void_gap(pipeline: Pipeline):
    return next((g for g in pipeline.gaps if "unknown, not empty" in str(g.detail)), None)


def _stored_game(game_id="E2026_379", scheduled_utc="2026-11-23T12:00:00Z") -> Game:
    # Midday UTC so the game stays on the config date in any European zone.
    return Game(
        game_id=game_id,
        league=League.EUROLEAGUE,
        season=SEASONCODE,
        home_team_id="1",
        away_team_id="2",
        home_name="Home",
        away_name="Away",
        scheduled_utc=scheduled_utc,
        status=GameStatus.SCHEDULED,
        provider_ids={"euroleague": game_id},
    )


def test_dead_provider_is_a_named_gap_not_an_empty_slate():
    pipeline = _pipeline()
    games = pipeline._discover_games()

    assert games == []
    gap = _schedule_gap(pipeline)
    assert gap is not None, "an unreadable schedule must be named"
    assert gap.provider == "euroleague"
    assert gap.game_id is None
    assert SEASONCODE in gap.detail
    assert "0 stored game(s)" in gap.detail
    assert "HttpError" in gap.detail
    # The void guard speaks for lanes that read nothing *and said nothing*.
    # This lane said why, so it must not also report the generic void.
    assert _void_gap(pipeline) is None

    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    assert row["state"] == CoverageState.PROVIDER_ERROR.value
    assert row["coverage"]["stored_slate_fallback"] is True
    assert row["coverage"]["seasoncode"] == SEASONCODE
    assert row["last_success_utc"] is None, "an unreadable fetch is not a success"


def test_day_continues_on_the_stored_slate_when_the_host_is_dead():
    store = InMemoryStore()
    stored = _stored_game()
    store.upsert_game(stored)

    pipeline = _pipeline(store=store)
    games = pipeline._discover_games()

    assert [game.game_id for game in games] == [stored.game_id]
    assert "1 stored game(s)" in _schedule_gap(pipeline).detail
    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    assert row["coverage"]["games"] == 1


def test_results_stage_survives_a_dead_schedule_provider():
    # The incident also lost settlement: ``_fetch_finals`` re-enters
    # ``_discover_games`` on this branch, so the guard has to hold there too.
    pipeline = _pipeline()
    report = pipeline.results()

    assert isinstance(report, dict)
    assert _schedule_gap(pipeline) is not None


@pytest.mark.parametrize("league", ["EUROLEAGUE", "EUROCUP"])
def test_both_euro_basketball_lanes_degrade(league):
    pipeline = _pipeline(league=league)
    assert pipeline._discover_games() == []
    assert _schedule_gap(pipeline) is not None


def test_a_healthy_read_is_unchanged():
    pipeline = _pipeline(fetch=_healthy_fetch)
    pipeline._discover_games()

    assert _schedule_gap(pipeline) is None
    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    assert row["state"] in {CoverageState.OK.value, CoverageState.MISSING.value}
    assert "stored_slate_fallback" not in row["coverage"]


def test_adapter_still_fails_closed_for_its_caller():
    """The pipeline is the guard; the adapter keeps raising and records nothing."""

    euroleague.reset_probe()
    with pytest.raises(HttpError):
        euroleague.discover_games("2026-11-23", _dead_fetch, SEASONCODE)
    assert not euroleague.last_probe(), "a failed read must not leave a receipt"
