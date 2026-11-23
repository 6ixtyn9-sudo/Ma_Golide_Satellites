"""A frozen leg settles on the team it was priced on, not on a side label.

ESPN 401928998 (Pensacola Invitational, 2026-11-21) is the regression case.  The
ticket was frozen at 2026-11-22T00:18:31Z with ``home`` = Valparaiso Beacons and
the price taken on that side; the game came back with Troy Trojans designated
home and Valparaiso away.  Valparaiso won 77-72 (24-37, 53-35), yet grading the
frozen ``home`` selection against the *result's* home side settled the winner as
a loss, which turned a winning acca into a loss.

The receipt always records the team names each leg was priced on, so the frozen
orientation can be restored: mirror the result when the two agree only
cross-wise, and leave everything else exactly as it was.
"""

from __future__ import annotations

from dataclasses import replace

from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind
from ingestion.pipeline import Pipeline, PipelineConfig, espn
from ingestion.store import InMemoryStore
from forecasting.operator import (
    SIDE_IDENTITY_MIRRORED,
    SIDE_IDENTITY_UNVERIFIED,
    _frozen_side_identity,
    _grade_row,
    _set_grade,
)

TIP_UTC = "2026-11-22T01:00:00Z"


def _game(
    *,
    home_name: str = "Troy Trojans",
    away_name: str = "Valparaiso Beacons",
    home_quarters: list[float] | None = None,
    away_quarters: list[float] | None = None,
) -> Game:
    """The stored result: ESPN designated Troy home, Valparaiso won 77-72."""

    home_quarters = home_quarters or [37.0, 35.0]
    away_quarters = away_quarters or [24.0, 53.0]
    actual = ActualResult(
        period_kind=PeriodKind.HALVES,
        home_quarters=list(home_quarters),
        away_quarters=list(away_quarters),
        home_final=sum(home_quarters),
        away_final=sum(away_quarters),
        source="espn",
    )
    return Game(
        game_id="401928998",
        league=League.NCAA_M,
        season="2026",
        home_team_id="2653",
        away_team_id="2674",
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=TIP_UTC,
        status=GameStatus.FINAL,
        actual=actual,
        provider_ids={"espn": "401928998"},
    )


def _leg(**overrides) -> dict:
    row = {
        "game_id": "401928998",
        "league": "NCAA_M",
        "scheduled_utc": TIP_UTC,
        "home": "Valparaiso Beacons",
        "away": "Troy Trojans",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "raw_probability": 0.595,
        "model": "team_history_empirical",
        "line": None,
    }
    row.update(overrides)
    return row


def test_a_flipped_result_settles_the_team_the_ticket_was_priced_on():
    game = _game()
    row = _leg()

    assert _frozen_side_identity(row, game) == SIDE_IDENTITY_MIRRORED
    assert _grade_row(row, {game.game_id: game}) == "win"

    _set_grade(row, {game.game_id: game})
    assert row["outcome"] == "win"
    assert row["result_side_identity"] == SIDE_IDENTITY_MIRRORED


def test_the_flipped_result_also_settles_the_other_frozen_side():
    game = _game()

    assert _grade_row(_leg(selection="away"), {game.game_id: game}) == "loss"


def test_the_flip_only_needs_the_teams_not_the_score_agreement():
    # Reversed scorelines stay mirrored too: the orientation check reads names.
    game = _game(home_quarters=[37.0, 35.0], away_quarters=[24.0, 53.0])
    row = _leg(selection="away")

    assert _grade_row(row, {game.game_id: game}) == "loss"


def test_the_frozen_spread_line_keeps_its_own_orientation():
    # Frozen: home = Valparaiso +4.5, so the away side (Troy) gave 4.5 and lost
    # by 5. Under the result's orientation the same line would have won.
    game = _game()
    row = _leg(market="spread", selection="away", line=4.5, raw_probability=0.455)

    assert _grade_row(row, {game.game_id: game}) == "loss"


def test_aligned_rows_are_not_marked_or_mirrored():
    aligned = _game(
        home_name="Valparaiso Beacons",
        away_name="Troy Trojans",
        home_quarters=[24.0, 53.0],
        away_quarters=[37.0, 35.0],
    )
    row = _leg()

    assert _frozen_side_identity(row, aligned) != SIDE_IDENTITY_MIRRORED
    assert _grade_row(row, {aligned.game_id: aligned}) == "win"
    _set_grade(row, {aligned.game_id: aligned})
    assert "result_side_identity" not in row


def test_unrecognised_names_keep_the_previous_side_label_behaviour():
    game = _game()
    row = _leg(home="Some Other Beacons", away="Another Trojans")

    assert _frozen_side_identity(row, game) == SIDE_IDENTITY_UNVERIFIED
    # Legacy path: the result's home side (Troy, 72) lost, so it grades a loss.
    assert _grade_row(row, {game.game_id: game}) == "loss"


def test_acquisition_flags_a_side_flip_for_the_operator(monkeypatch):
    """The recheck keeps the result but tells the operator the sides moved."""

    store = InMemoryStore()
    # Frozen orientation: Valparaiso home, still awaiting its result.
    frozen = _game(home_name="Valparaiso Beacons", away_name="Troy Trojans")
    store.upsert_game(
        replace(frozen, status=GameStatus.SCHEDULED, actual=None)
    )
    fetched = _game()  # ESPN now reports Troy home.

    monkeypatch.setattr(
        espn, "fetch_recent_results", lambda *args, **kwargs: [fetched]
    )
    pipeline = Pipeline(
        PipelineConfig(
            league="NCAA_M",
            date_yyyymmdd="20261122",
            fixture_now_utc="2026-11-22T12:00:00Z",
        ),
        store,
    )

    finals = pipeline._recheck_pending_finals(set())

    assert [game.game_id for game in finals] == ["401928998"]
    flips = [gap for gap in pipeline.gaps if gap.required_field == "result_side_identity"]
    assert len(flips) == 1
    assert flips[0].game_id == "401928998"
    assert "other way round" in flips[0].detail
