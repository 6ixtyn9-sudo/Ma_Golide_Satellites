"""A lane that cannot predict must say so, not return an empty dict.

Run 45250804751, head `eb45dd2`: Italy LBA logged ``fixtures_found: 6`` and
Israel ``fixtures_found: 3``, both with a fresh named full-game moneyline on
every slate game, and both reported ``games_with_any_pre_tip_prediction: 0``
and ``predictions_selected: {}``.

The cause is structural, not transient: ``_predixsport_candidates_for_games``
returns before contacting anyone unless the league is NBA or EUROLEAGUE, so the
only external forecast feed is unreachable for every other lane, and none of
them is an ESPN BPI lane. No number of fixtures and no price can turn into a
pick without a probability.

Before this change the report said nothing about that. These tests pin the
attribution.
"""

from __future__ import annotations

from ingestion.models import Game, GameStatus, League
from ingestion.pipeline import Pipeline
from ingestion.pipeline import PipelineConfig


def _game(league: League, game_id: str = "G1") -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026",
        scheduled_utc="2026-11-23T19:00:00Z",
        home_team_id="H",
        away_team_id="A",
        home_name="Home",
        away_name="Away",
        status=GameStatus.SCHEDULED,
        provider_ids={},
    )


def _pipeline(league: str) -> Pipeline:
    from ingestion.store import InMemoryStore

    return Pipeline(PipelineConfig(league=league, date_yyyymmdd="20261123"), InMemoryStore())


def test_a_lane_with_no_wired_prediction_source_files_an_unsupported_gap():
    pipeline = _pipeline("ITALY_LBA")
    games = [_game(League.ITALY_LBA, "LBA-1"), _game(League.ITALY_LBA, "LBA-2")]

    selections = pipeline._acquire_predictions(games)

    # Every game is declined, and now the reason is named at slate level too:
    # the per-game 'pregame_prediction' gap says "no fresh candidate", which is
    # true of a model that declined and of a lane that has no model.
    assert all(selection.selected is None for selection in selections.values())
    gap = next(g for g in pipeline.gaps if g.required_field == "pre_tip_prediction")
    assert gap.reason.value == "unsupported"
    assert "ITALY_LBA" in gap.detail
    assert "2 fixtures" in gap.detail
    assert "EUROLEAGUE, NBA" in gap.detail
    status = pipeline.prediction_provider_status["capability"]
    assert status["league"] == "ITALY_LBA"
    assert status["games"] == 2


def test_the_capability_receipt_names_the_sources_each_league_can_use():
    lba = _pipeline("ITALY_LBA")
    lba._note_prediction_capability([_game(League.ITALY_LBA)])
    assert lba.prediction_provider_status["capability"]["sources"] == ["stored_pre_tip"]

    nba = _pipeline("NBA")
    nba._note_prediction_capability([_game(League.NBA)])
    sources = nba.prediction_provider_status["capability"]["sources"]
    assert {"espn_bpi", "hardwood", "predixsport", "stored_pre_tip"} <= set(sources)
    # An ESPN lane is capable, so it must not be told it can never pick.
    assert [g for g in nba.gaps if g.required_field == "pre_tip_prediction"] == []


def test_an_espn_lane_that_is_not_nba_is_capable_but_not_predixsport_capable():
    pipeline = _pipeline("NCAA_M")
    pipeline._note_prediction_capability([_game(League.NCAA_M)])

    sources = pipeline.prediction_provider_status["capability"]["sources"]
    assert "espn_bpi" in sources
    assert "predixsport" not in sources
    assert [g for g in pipeline.gaps if g.required_field == "pre_tip_prediction"] == []


def test_an_empty_slate_is_not_called_incapable():
    pipeline = _pipeline("ITALY_LBA")
    pipeline._note_prediction_capability([])

    assert [g for g in pipeline.gaps if g.required_field == "pre_tip_prediction"] == []
    assert pipeline.prediction_provider_status["capability"]["games"] == 0
