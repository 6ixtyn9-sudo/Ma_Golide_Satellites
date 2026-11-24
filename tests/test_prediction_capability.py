"""A lane that cannot predict must say so, not return an empty dict.

Run 45250804751, head `eb45dd2`: Italy LBA logged ``fixtures_found: 6`` and
Israel ``fixtures_found: 3``, both with a fresh named full-game moneyline on
every slate game, and both reported ``games_with_any_pre_tip_prediction: 0``
and ``predictions_selected: {}``.

The cause was structural, not transient: ``_predixsport_candidates_for_games``
returns before contacting anyone unless the league is NBA or EUROLEAGUE, so the
only external forecast feed is unreachable for every other lane, and none of
them is an ESPN BPI lane. No number of fixtures and no price can turn into a
pick without a probability.

Two of these tests have since been re-pinned to the *next* contract. The wall
was broken by giving every non-ESPN lane a lane-local model
(``MarketForecastService.history_prediction``), so a slate-level "no source is
wired, it can never pick" is no longer true for LBA: the lane has a source and
declines per game, naming the missing evidence. The old assertions pinned the
silence that change removes. The slate-level gap still exists for a league
where ``stored_pre_tip`` is the only source -- it is now genuinely unreachable
rather than merely unfired, which is why the guard is stated as a set test.
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


def test_a_lane_with_an_empty_store_names_the_evidence_it_is_missing():
    """The lane now HAS a source; the decline must say what it lacks.

    Before the local model was wired, this case filed a slate-level
    ``unsupported`` gap saying the lane could never pick. That is no longer
    true, and saying it would hide the actionable fact: the store holds no
    reconciled finals yet.
    """

    pipeline = _pipeline("ITALY_LBA")
    games = [_game(League.ITALY_LBA, "LBA-1"), _game(League.ITALY_LBA, "LBA-2")]

    selections = pipeline._acquire_predictions(games)

    assert all(selection.selected is None for selection in selections.values())
    declined = [g for g in pipeline.gaps if g.required_field == "pre_tip_prediction"]
    assert {g.game_id for g in declined} == {"LBA-1", "LBA-2"}
    assert all(g.provider == "history_model" for g in declined)
    assert all(g.reason.value == "missing" for g in declined)
    assert all("no_reconciled_finals_in_store" in g.detail for g in declined)
    assert [g for g in pipeline.gaps if g.reason.value == "unsupported"] == []
    status = pipeline.prediction_provider_status["capability"]
    assert status["league"] == "ITALY_LBA"
    assert status["games"] == 2
    assert "history_model" in status["sources"]


def test_the_capability_receipt_names_the_sources_each_league_can_use():
    lba = _pipeline("ITALY_LBA")
    lba._note_prediction_capability([_game(League.ITALY_LBA)])
    assert lba.prediction_provider_status["capability"]["sources"] == [
        "stored_pre_tip",
        "history_model",
    ]
    # A wired lane-local model must not also be told it can never pick.
    assert [g for g in lba.gaps if g.required_field == "pre_tip_prediction"] == []

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
