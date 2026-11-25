"""Settlement from the aggregator's match page, for the lanes promoted on it.

The reader's own rules are tested in ``test_betexplorer_results_reader``. This
file tests the *lane*: a stored fixture settles from the code it was stored
under, a page that does not add up settles nothing, and the three promoted lanes
are wired into the pending-final re-check rather than merely relabelled.

The page shape and the numbers are the ones served for NKL code ``McCqxqs3`` on
2026-11-25; the wrapper tags are not assumed anywhere, because they were never
captured and the reader works from the text.
"""

from __future__ import annotations

from ingestion.models import Game, GameStatus, League, PeriodKind
from ingestion.pipeline import AGGREGATOR_MATCH_LEAGUES, Pipeline, PipelineConfig
from ingestion.providers import betexplorer_results
from ingestion.store import InMemoryStore

MATCH_PAGE = (
    "<html><body>"
    "<p>24.11.2026 - 12:00</p><p>93:107</p><p>(15:25, 30:28, 17:27, 31:27)</p>"
    "</body></html>"
)
NO_SPLIT_PAGE = "<html><body><p>24.11.2026 - 12:00</p><p>93:107</p></body></html>"
WRONG_SUM_PAGE = (
    "<html><body><p>94:107</p><p>(15:25, 30:28, 17:27, 31:27)</p></body></html>"
)


def _game(
    *,
    league: League = League.GREECE_BASKET_LEAGUE,
    game_id: str = "betexplorer:McCqxqs3",
    slug: str = "jurbarkas-vytis-vdu",
    code: str = "McCqxqs3",
    scheduled_utc: str = "2026-11-24T12:00:00Z",
) -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026-2027",
        home_team_id="",
        away_team_id="",
        home_name="Jurbarkas",
        away_name="Vytis VDU",
        scheduled_utc=scheduled_utc,
        status=GameStatus.SCHEDULED,
        provider_ids={"betexplorer": code, "betexplorer_slug": slug},
    )


def _pipeline(page=MATCH_PAGE, *, league="GREECE_BASKET_LEAGUE", store=None) -> Pipeline:
    pipeline = Pipeline(
        PipelineConfig(league=league, date_yyyymmdd="20261125"),
        store or InMemoryStore(),
        fetch=lambda url: page,
        allow_writes=False,
    )
    # The host's own floor is real, but a test that sleeps is a test nobody runs.
    pipeline._pause = lambda seconds: None
    return pipeline


def _result_gap(pipeline: Pipeline):
    return next(
        (gap for gap in pipeline.gaps if gap.required_field == "aggregator_match_result"),
        None,
    )


def test_a_stored_fixture_settles_from_its_own_provider_code():
    pipeline = _pipeline()
    game = _game()

    settled = pipeline._betexplorer_match_settle({game.game_id: game})

    assert [g.game_id for g in settled] == [game.game_id]
    result = settled[0].actual
    assert settled[0].status is GameStatus.FINAL
    assert result is not None and result.period_kind is PeriodKind.QUARTERS
    assert (result.home_final, result.away_final) == (93.0, 107.0)
    assert result.home_quarters == [15.0, 30.0, 17.0, 31.0]
    assert result.away_quarters == [25.0, 28.0, 27.0, 27.0]
    # The whole point of reading the page rather than the index beside it.
    assert result.reconciles()[0] is True
    # An aggregator number, never a federation final, and it says so.
    assert result.source == betexplorer_results.SOURCE


def test_a_page_without_a_split_settles_nothing_and_says_why():
    pipeline = _pipeline(NO_SPLIT_PAGE)
    game = _game()

    assert pipeline._betexplorer_match_settle({game.game_id: game}) == []
    gap = _result_gap(pipeline)
    assert gap is not None
    assert gap.game_id == game.game_id
    assert "not a settlement record" in gap.detail


def test_a_split_that_does_not_add_up_is_refused_not_written():
    pipeline = _pipeline(WRONG_SUM_PAGE)

    assert pipeline._betexplorer_match_settle({_game().game_id: _game()}) == []
    assert _result_gap(pipeline) is not None


def test_a_fixture_without_a_match_code_is_a_named_gap():
    pipeline = _pipeline()
    game = _game()
    game.provider_ids.pop("betexplorer")

    assert pipeline._betexplorer_match_settle({game.game_id: game}) == []
    gap = next(g for g in pipeline.gaps if g.required_field == "aggregator_match_identity")
    assert gap.game_id == game.game_id


def test_the_three_promoted_lanes_are_wired_into_the_pending_recheck():
    assert AGGREGATOR_MATCH_LEAGUES == {
        League.GREECE_BASKET_LEAGUE,
        League.CHINA_WCBA_W,
        League.AUSTRALIA_WNBL_W,
    }
    for league in AGGREGATOR_MATCH_LEAGUES:
        store = InMemoryStore()
        store.upsert_game(_game(league=league))
        pipeline = _pipeline(league=league.value, store=store)

        finals = pipeline._recheck_pending_finals(set())

        assert [game.game_id for game in finals] == ["betexplorer:McCqxqs3"], league


def test_a_promoted_lane_whose_page_is_unreadable_is_a_gap_not_a_failure():
    store = InMemoryStore()
    store.upsert_game(_game())

    def refused(url):
        raise OSError("connection reset")

    pipeline = Pipeline(
        PipelineConfig(league="GREECE_BASKET_LEAGUE", date_yyyymmdd="20261125"),
        store,
        fetch=refused,
        allow_writes=False,
    )
    pipeline._pause = lambda seconds: None

    assert pipeline._recheck_pending_finals(set()) == []
    gap = next(g for g in pipeline.gaps if g.required_field == "aggregator_match_result")
    assert gap.game_id == "betexplorer:McCqxqs3"
