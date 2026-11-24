"""A refused fixtures page must cost the slate, not the whole tick.

Live evidence (run 45376480832, 2026-11-24T10:40:47Z, lane ISRAEL_LIGA_LEUMIT):

    {"status": "failed",
     "error": "HTTP 429 for https://www.betexplorer.com/basketball/israel/liga-leumit/fixtures/: "}

The transport retried twice (1.5s, 2.25s), the host kept refusing, ``HttpError``
left the adapter, and ``_discover_games`` had no guard on the
``betexplorer_fixtures`` branch -- so the exception unwound to
``scripts/daily.py``, which can only print ``{"status": "failed"}`` and exit 1.
The job went red having written no card, no coverage gap and no lifecycle state:
a *refused* read was rendered as a failed day, which is the same class of error
as an empty slate.

The EuroLeague branch already answers this (see
``tests/test_euroleague_provider_degrade.py``). The contract asserted here is the
same one, against the fixtures-table branch:

* an HTTP refusal becomes a *named* ``official_schedule`` gap with the status it
  was given, plus a ``provider_error`` health row;
* the tick continues on the canonical stored slate, so settlement is not lost;
* the void guard stays silent -- the lane said why, so it is not "unknown";
* a page that arrives without its table (``IdentityError``) still raises: a
  source we can read but no longer parse is a code failure, not the host's.
"""

from pathlib import Path

import pytest

from ingestion.http import HttpError
from ingestion.models import CoverageState, Game, GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import IdentityError, betexplorer_fixtures
from ingestion.store import InMemoryStore

LEAGUE = "ISRAEL_LIGA_LEUMIT"
DATE = "20261124"
FIXTURES_URL = "https://www.betexplorer.com/basketball/israel/liga-leumit/fixtures/"
_CAPTURED_PAGE = (
    Path(__file__).resolve().parent / "fixtures" / "betexplorer" / "el_fixtures_table.html"
)


def _refused_429(url: str):
    """Reproduce the incident: the host answers 429 and keeps refusing."""

    raise HttpError(429, url, "")


def _tableless_page(url: str):
    """A response did arrive, but it is not a fixtures table."""

    return "<html><body><h1>Too many requests</h1></body></html>"


def _healthy_page():
    # The served table shape is identical across competitions (see
    # tests/test_betexplorer_fixtures_lane.py), so a real captured page is the
    # right stand-in for "the host answered with a table".
    return _CAPTURED_PAGE.read_text(encoding="utf-8")


def _pipeline(fetch=_refused_429, store=None, date=DATE, league=LEAGUE) -> Pipeline:
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


def _stored_game(game_id="betexplorer:israel-1", scheduled_utc="2026-11-24T12:00:00Z") -> Game:
    # Midday UTC so the game stays on the config date in any acquisition zone.
    return Game(
        game_id=game_id,
        league=League.ISRAEL_LIGA_LEUMIT,
        season="2026-2027",
        home_team_id="",
        away_team_id="",
        home_name="Maccabi Rehovot",
        away_name="Elitzur Netanya",
        scheduled_utc=scheduled_utc,
        status=GameStatus.SCHEDULED,
        provider_ids={"betexplorer": game_id},
    )


def test_a_refused_fixtures_page_is_a_named_gap_not_a_failed_tick():
    pipeline = _pipeline()
    games = pipeline._discover_games()

    assert games == []
    gap = _schedule_gap(pipeline)
    assert gap is not None, "an unreadable schedule must be named, never silent"
    assert gap.provider == betexplorer_fixtures.SCHEDULE_SOURCE
    assert gap.game_id is None
    assert "HttpError" in gap.detail
    assert "429" in gap.detail

    row = next(
        h for h in pipeline._health_written if h["provider"] == betexplorer_fixtures.SCHEDULE_SOURCE
    )
    assert row["state"] == CoverageState.PROVIDER_ERROR.value
    assert row["coverage"]["http_status"] == 429
    assert row["coverage"]["stored_slate_fallback"] is True
    assert row["last_success_utc"] is None, "a refused read is not a success"


def test_the_void_guard_does_not_also_fire():
    # One named reason is the contract; a second generic gap would say the lane
    # both explained itself and said nothing.
    pipeline = _pipeline()
    pipeline._discover_games()

    assert _void_gap(pipeline) is None


def test_the_tick_continues_on_the_stored_slate():
    store = InMemoryStore()
    stored = _stored_game()
    store.upsert_game(stored)

    pipeline = _pipeline(store=store)
    games = pipeline._discover_games()

    assert [game.game_id for game in games] == [stored.game_id]
    assert "1 stored game(s)" in _schedule_gap(pipeline).detail
    row = next(
        h for h in pipeline._health_written if h["provider"] == betexplorer_fixtures.SCHEDULE_SOURCE
    )
    assert row["coverage"]["games"] == 1


def test_the_run_is_not_unwound_by_a_refused_read():
    """The incident's real cost: the exception left no stage a chance to run.

    This lane's results stage does not re-enter discovery (unlike EuroLeague's),
    so the guard that matters here is simply that the schedule stage *returns*
    and the next stage still executes.
    """

    pipeline = _pipeline(store=InMemoryStore())

    assert pipeline._discover_games() == []  # returns with a gap, does not raise
    assert _schedule_gap(pipeline) is not None
    assert isinstance(pipeline.results(), dict)


def test_a_page_without_its_table_still_fails_loudly():
    """The guard is for refusals, not for a parser that no longer fits."""

    pipeline = _pipeline(fetch=_tableless_page)
    with pytest.raises(IdentityError):
        pipeline._discover_games()

    assert _schedule_gap(pipeline) is None


def test_a_healthy_read_is_unchanged():
    pipeline = _pipeline(fetch=lambda url: _healthy_page())
    games = pipeline._discover_games()

    # Whatever the captured page yields for this date, the read happened: no
    # refusal gap, no generic void, and a health row that is not a fallback.
    assert isinstance(games, list)
    assert _schedule_gap(pipeline) is None
    assert _void_gap(pipeline) is None
    row = next(
        h for h in pipeline._health_written if h["provider"] == betexplorer_fixtures.SCHEDULE_SOURCE
    )
    assert row["state"] in {CoverageState.OK.value, CoverageState.MISSING.value}
    assert "stored_slate_fallback" not in row["coverage"]
