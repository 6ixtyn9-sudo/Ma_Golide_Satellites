"""The bridge from a published table to the numbers the standings scorer reads.

Values here are the served NKL table captured on 2026-11-25 -- Telsiai 16 15 1
1449:1186, Neptunas 2 at rank 2 with pct 0.7647, BC Olimpas Palanga at rank 3 with
pct 0.8125 -- because those are the rows that proved the route and they carry the
awkward case (a rank order that is not win percentage) that the bridge must not
"fix".
"""

from __future__ import annotations

from forecasting.magolide_score import (
    MaGolideConfig,
    build_standings_features,
    score_features,
)
from forecasting.standings import PUBLISHED_SOURCE, SOURCE, from_published
from ingestion.models import StandingsRow as PublishedRow

TEAM_IDS = {
    "Telsiai": "IHmmS3rt",
    "Neptunas 2": "NEPT02",
    "BC Olimpas Palanga": "OLIM01",
}


def _row(
    team: str,
    *,
    position: int | None = 1,
    played: int | None = 16,
    wins: int | None = 15,
    losses: int | None = 1,
    pf: float | None = 1449.0,
    pa: float | None = 1186.0,
    win_pct: float | None = None,
    home_record: str | None = None,
    away_record: str | None = None,
) -> PublishedRow:
    return PublishedRow(
        team_name=team,
        position=position,
        games_played=played,
        wins=wins,
        losses=losses,
        points_for=pf,
        points_against=pa,
        win_pct=win_pct,
        home_record=home_record,
        away_record=away_record,
        source=PUBLISHED_SOURCE,
    )


def test_a_served_row_becomes_a_scorer_row_by_arithmetic_only():
    rows = from_published([_row("Telsiai")])

    assert len(rows) == 1
    tel = rows[0]
    assert tel.team == "Telsiai"
    assert tel.rank == 1.0
    assert tel.pct == round(15 / 16, 4)
    # (1449 - 1186) / 16, a points-per-game difference and not a rating.
    assert tel.net_rtg == 16.44
    assert tel.home_pct is None
    assert tel.away_pct is None
    assert tel.streak is None


def test_rank_is_quoted_even_where_it_disagrees_with_win_percentage():
    """NKL ranked a 0.7647 club second and a 0.8571 club fifth, as served."""

    published = [
        _row("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
        _row("Neptunas 2", position=2, played=17, wins=13, losses=4, pf=1354.0, pa=1278.0),
        _row("BC Olimpas Palanga", position=3, played=16, wins=13, losses=3, pf=1553.0, pa=1350.0),
        _row("Jurbarkas", position=5, played=14, wins=12, losses=2, pf=1256.0, pa=1111.0),
    ]

    rows = from_published(published)

    assert [row.team for row in rows] == [r.team_name for r in published]
    assert [row.rank for row in rows] == [1.0, 2.0, 3.0, 5.0]
    # The published order is not pct order, and the bridge does not re-sort it.
    assert rows[1].pct < rows[2].pct
    assert rows[3].pct > rows[1].pct


def test_split_records_become_percentages():
    rows = from_published([_row("Telsiai", home_record="8-1", away_record="7-0")])

    assert rows[0].home_pct == round(8 / 9, 4)
    assert rows[0].away_pct == 1.0


def test_an_empty_or_unreadable_split_is_absent_not_zero():
    rows = from_published(
        [
            _row("Telsiai", home_record="0-0", away_record="not-a-record"),
            _row("Jurbarkas", position=5, home_record="", away_record=None),
        ]
    )

    assert rows[0].home_pct is None
    assert rows[0].away_pct is None
    assert rows[1].home_pct is None
    assert rows[1].away_pct is None


def test_derived_history_fills_the_streak_without_overwriting_published_numbers():
    tel = _row("Telsiai")
    graded = from_published([tel])[0].__class__(
        team="Telsiai",
        rank=9.0,
        pct=0.5,
        net_rtg=-3.0,
        home_pct=0.25,
        away_pct=0.25,
        streak=4,
        l10_wins=7.0,
        l10_losses=3.0,
    )

    rows = from_published([tel], derived={"Telsiai": graded})

    assert rows[0].pct == round(15 / 16, 4)  # published wins
    assert rows[0].net_rtg == 16.44  # published wins
    assert rows[0].rank == 1.0  # published wins
    assert rows[0].streak == 4  # only the finals can say this
    assert (rows[0].l10_wins, rows[0].l10_losses) == (7.0, 3.0)


def test_a_club_without_derived_history_keeps_an_absent_streak():
    rows = from_published([_row("Telsiai")], derived={})

    assert rows[0].streak is None
    assert rows[0].l10_wins is None


def test_a_club_that_has_not_played_is_absent_rather_than_zero():
    rows = from_published(
        [
            _row("Telsiai"),
            _row("No Games", position=18, played=0, wins=0, losses=0),
        ]
    )

    assert [row.team for row in rows] == ["Telsiai"]


def test_a_row_without_games_played_is_absent():
    rows = from_published([_row("Telsiai"), _row("Blank", played=None, wins=None)])

    assert [row.team for row in rows] == ["Telsiai"]


def test_a_repeated_club_is_listed_once():
    rows = from_published([_row("Telsiai"), _row(" Telsiai ", position=9)])

    assert len(rows) == 1
    assert rows[0].rank == 1.0


def test_the_provider_pct_is_used_when_it_is_present():
    rows = from_published([_row("Telsiai", win_pct=0.9375)])

    assert rows[0].pct == 0.9375


def test_missing_totals_leave_the_difference_absent():
    rows = from_published([_row("Telsiai", pf=None, pa=None)])

    assert rows[0].net_rtg is None
    assert rows[0].pct is not None


def test_the_bridge_actually_feeds_the_scorer():
    """The point of the bridge: a published table yields a real verdict."""

    published = [
        _row("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0,
             home_record="8-1", away_record="7-0"),
        _row("Rytas 2", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0,
             home_record="1-7", away_record="0-8"),
    ]
    rows = {row.team: row for row in from_published(published)}

    features = build_standings_features(rows["Telsiai"], rows["Rytas 2"])
    verdict = score_features(features, MaGolideConfig())

    assert features.rank_diff == 17.0
    assert features.net_rtg_diff > 0
    assert verdict.prediction == "HOME"
    assert verdict.meta["mode"] == "NEW"
    assert verdict.meta["weights_source"] == "validator_defaults_not_a_fitted_config"


def test_the_two_sources_stay_distinguishable():
    assert SOURCE != PUBLISHED_SOURCE
    assert PUBLISHED_SOURCE == "betexplorer_published_table"


def test_a_winning_record_does_not_imply_a_streak():
    """A 15-1 table says nothing about the order of those fifteen wins."""

    rows = from_published([_row("Telsiai", played=16, wins=15, losses=1)])

    assert rows[0].pct == round(15 / 16, 4)
    assert rows[0].streak is None
