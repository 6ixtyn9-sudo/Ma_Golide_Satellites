"""Research rows and paper slips for lanes with a table and no ticket route.

Standings values are the served NKL table of 2026-11-25 (Telsiai 16 15 1
1449:1186 first, Rytas 2 16 1 15 1268:1428 last), and prices are typed as
``aggregate_consensus`` because that is the only provenance this lane's source can
offer. Nothing here reaches a stake, a receipt or a bank balance, and several tests
exist only to pin that.
"""

from __future__ import annotations

import pytest

from forecasting.magolide_score import MaGolideConfig, MaGolideVerdict
from forecasting.research_slate import (
    PRICE_LABEL,
    RESEARCH_BUCKET,
    RESEARCH_LABEL,
    TICKET_LABEL,
    render_research_section,
    research_rows,
    research_tickets,
)
from ingestion.models import (
    BookSnapshot,
    Game,
    League,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
)
from ingestion.models import StandingsRow as PublishedRow

LEAGUE = "LITHUANIA_NKL"


def _published(team: str, *, position: int, played: int, wins: int, losses: int,
               pf: float, pa: float, section: str | None = None) -> PublishedRow:
    return PublishedRow(
        team_name=team,
        position=position,
        games_played=played,
        wins=wins,
        losses=losses,
        points_for=pf,
        points_against=pa,
        win_pct=round(wins / played, 4),
        home_record=None,
        away_record=None,
        section=section,
        source="betexplorer_published_table",
    )


TABLE = [
    _published("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
    _published("Rytas 2", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
]


def _game(game_id: str = "bes:McCqxqs3", home: str = "Telsiai", away: str = "Rytas 2") -> Game:
    return Game(
        game_id=game_id,
        league=League.LITHUANIA_NKL,
        season="2026",
        home_team_id="x",
        away_team_id="y",
        home_name=home,
        away_name=away,
        scheduled_utc="2026-12-01T17:00:00Z",
    )


def _snapshot(
    *,
    game_id: str = "bes:McCqxqs3",
    home: str = "Telsiai",
    away: str = "Rytas 2",
    home_price: float | None = 1.15,
    away_price: float | None = 6.50,
    provenance: PriceProvenance = PriceProvenance.AGGREGATE_CONSENSUS,
) -> BookSnapshot:
    return BookSnapshot(
        game_id=game_id,
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-25T18:00:00Z",
        home_price_decimal=home_price,
        away_price_decimal=away_price,
        price_provenance=provenance,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
        home_name=home,
        away_name=away,
        scheduled_utc="2026-12-01T17:00:00Z",
    )


# -- scoring ----------------------------------------------------------------


def test_a_served_table_produces_a_row_with_a_side_and_a_price():
    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    assert len(rows) == 1
    row = rows[0]
    assert isinstance(row.verdict, MaGolideVerdict)
    assert row.verdict.prediction == "HOME"
    assert row.selection == "HOME"
    assert row.decimal_price == 1.15
    assert row.price_book == "average"
    assert row.price_provenance == "aggregate_consensus"
    assert row.standings_evidence["home_standings_present"] is True
    assert row.standings_evidence["away_standings_present"] is True
    assert row.standings_evidence["neutral_defaults_applied"] is False


def test_two_absent_clubs_still_get_a_row_and_read_risky():
    rows = research_rows(
        LEAGUE,
        [_game(home="Unknown BC", away="Also Unknown")],
        TABLE,
        snapshots=[_snapshot(home="Unknown BC", away="Also Unknown")],
    )

    assert len(rows) == 1
    assert rows[0].standings_evidence["neutral_defaults_applied"] is True
    assert rows[0].standings_evidence["home_standings_present"] is False
    assert rows[0].verdict.prediction == "RISKY"
    assert rows[0].selection is None


def test_one_absent_club_is_scored_against_the_known_row_not_filtered_out():
    """A thin table is not a reason to hide a fixture — only to label it."""

    rows = research_rows(LEAGUE, [_game(home="Unknown BC")], TABLE, snapshots=[_snapshot()])

    assert rows[0].standings_evidence["neutral_defaults_applied"] is True
    assert rows[0].selection == "HOME"  # against a 1-15 club the neutral row still leans home
    assert rows[0].verdict.prediction == "HOME"


def test_a_priced_row_can_still_have_no_selection():
    """RISKY means no side. The row is visible; nothing is proposed."""

    rows = research_rows(
        LEAGUE,
        [_game(home="Unknown BC", away="Also Unknown")],
        TABLE,
        snapshots=[_snapshot(home="Unknown BC", away="Also Unknown")],
    )

    assert rows[0].observed is True  # the market was there
    assert rows[0].priced is False  # but there is no side to price
    assert rows[0].decimal_price is None
    assert rows[0].has_selection is False
    assert research_tickets(rows) == []


def test_the_away_side_takes_the_away_price():
    table = [
        _published("Rytas 2", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
        _published("Telsiai", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
    ]
    rows = research_rows(LEAGUE, [_game()], table, snapshots=[_snapshot()])

    assert rows[0].selection == "AWAY"
    assert rows[0].decimal_price == 6.50


def test_a_fixture_without_a_snapshot_is_unpriced_not_estimated():
    rows = research_rows(LEAGUE, [_game()], TABLE)

    assert rows[0].priced is False
    assert rows[0].price_book is None
    assert rows[0].selection == "HOME"  # the read still exists
    assert research_tickets(rows) == []


def test_a_snapshot_matches_by_club_pair_when_the_ids_differ():
    rows = research_rows(
        LEAGUE,
        [_game(game_id="store:fixture-1")],
        TABLE,
        snapshots=[_snapshot(game_id="betexplorer:McCqxqs3")],
    )

    assert rows[0].decimal_price == 1.15


def test_a_split_market_row_is_not_used_as_a_full_game_price():
    half = BookSnapshot(
        game_id="bes:McCqxqs3",
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.H1,
        observed_utc="2026-11-25T18:00:00Z",
        home_price_decimal=1.80,
        away_price_decimal=1.90,
        price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
    )

    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[half])

    assert rows[0].priced is False


# -- what a row may never claim ---------------------------------------------


def test_no_row_is_ever_acca_eligible_or_a_bank_position():
    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    for row in rows:
        payload = row.as_dict()
        assert payload["acca_eligible"] is False
        assert payload["bank_impact"] == "none"
        assert payload["bucket"] == RESEARCH_BUCKET
        assert payload["price_provenance"] == "aggregate_consensus"


def test_the_row_names_both_provenances_and_the_default_weights():
    line = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0].as_line()

    assert "weights=validator_defaults_not_a_fitted_config" in line
    assert "standings=betexplorer_published_table" in line
    assert PRICE_LABEL in line
    assert "acca_eligible=False" in line


def test_the_uncalibrated_number_is_reported_as_a_number_not_a_probability():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0]

    assert 0.0 <= row.verdict.uncalibrated_logistic_pct <= 100.0
    assert "uncalibrated_logistic" in row.as_line()
    assert "probability" not in row.as_line().lower().replace(
        "uncalibrated_logistic", ""
    ).replace("no observed price", "")


# -- paper slips -------------------------------------------------------------


def _two_games() -> list[Game]:
    return [
        _game(game_id="g1", home="Telsiai", away="Rytas 2"),
        _game(game_id="g2", home="Telsiai", away="Rytas 2"),
    ]


def test_one_priced_selection_makes_one_single():
    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    tickets = research_tickets(rows)

    assert len(tickets) == 1
    assert len(tickets[0].legs) == 1
    assert tickets[0].combined_decimal == 1.15
    assert "SINGLE" in tickets[0].label


def test_two_priced_fixtures_make_singles_and_a_two_leg_slip():
    snapshots = [
        _snapshot(game_id="g1"),
        _snapshot(game_id="g2", home_price=1.40, away_price=4.00),
    ]
    rows = research_rows(LEAGUE, _two_games(), TABLE, snapshots=snapshots)

    tickets = research_tickets(rows)

    assert len(tickets) == 3  # two singles and one slip
    slip = [ticket for ticket in tickets if len(ticket.legs) == 2][0]
    assert slip.combined_decimal == round(1.15 * 1.40, 6)
    assert "2-LEG SLIP" in slip.label
    assert TICKET_LABEL in slip.label


def test_a_slip_never_repeats_one_fixture():
    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    for ticket in research_tickets(rows):
        ids = [leg.game_id for leg in ticket.legs]
        assert len(ids) == len(set(ids))


def test_a_slip_does_not_estimate_a_combined_probability():
    snapshots = [_snapshot(game_id="g1"), _snapshot(game_id="g2")]
    rows = research_rows(LEAGUE, _two_games(), TABLE, snapshots=snapshots)

    lines = "\n".join(
        line for ticket in research_tickets(rows) for line in ticket.as_lines()
    )

    assert "combined probability: NOT ESTIMATED" in lines


# -- the section ------------------------------------------------------------


def test_the_section_is_fenced_and_counts_what_it_holds():
    snapshots = [_snapshot(game_id="g1"), _snapshot(game_id="g2", home_price=None,
                                                    away_price=4.00)]
    rows = research_rows(LEAGUE, _two_games(), TABLE, snapshots=snapshots)

    lines = render_research_section(rows, date_label="2026-12-01", league=LEAGUE)
    text = "\n".join(lines)

    assert lines[0] == ""
    assert lines[1].startswith(RESEARCH_LABEL)
    assert "2 fixture(s), 2 with a side, 1 priced" in lines[1]
    assert "No named-book quote, no stake, no receipt, no bank or ROI effect." in text
    assert "acca_eligible=False on every row" in text


def test_the_section_carries_no_pending_marker():
    """The cards' readers treat ``PENDING`` as "this card needs rebuilding"."""

    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    text = "\n".join(render_research_section(rows, date_label="2026-12-01", league=LEAGUE))

    assert "PENDING" not in text.upper()
    assert "grade=NOT_SETTLEABLE_HERE" in text


def test_an_empty_slate_renders_nothing_at_all():
    assert render_research_section([], date_label="2026-12-01", league=LEAGUE) == []


def test_the_section_reports_when_neutral_defaults_were_used():
    rows = research_rows(LEAGUE, [_game(home="Unknown BC")], TABLE)

    text = "\n".join(render_research_section(rows, date_label="2026-12-01", league=LEAGUE))

    assert "scored on the scorer's neutral defaults" in text


# -- ranks belong to a group -------------------------------------------------


WINNERS = "Promotion - NKL (Winners stage: )"
LOSERS = "NKL (Losers stage: )"


def _staged(team: str, *, position: int, wins: int, losses: int, pf: float, pa: float,
            section: str) -> PublishedRow:
    return _published(
        team, position=position, played=16, wins=wins, losses=losses,
        pf=pf, pa=pa, section=section,
    )


def test_a_rank_inside_one_group_is_not_a_rank_inside_another():
    """NKL lists both stages as one 18-row table; the ranks are not one ladder."""

    table = [
        _staged("Alpha", position=2, wins=13, losses=3, pf=1300.0, pa=1200.0, section=WINNERS),
        _staged("Beta", position=16, wins=3, losses=13, pf=1150.0, pa=1350.0, section=LOSERS),
    ]

    row = research_rows(LEAGUE, [_game(home="Alpha", away="Beta")], table)[0]

    assert row.rank_suppressed is True
    assert row.standings_evidence["home_section"] == WINNERS
    assert row.standings_evidence["away_section"] == LOSERS
    assert row.standings_evidence["home_rank"] is None  # dropped, not hidden


def test_clubs_in_the_same_group_keep_their_ranks():
    table = [
        _staged("Alpha", position=2, wins=13, losses=3, pf=1300.0, pa=1200.0, section=WINNERS),
        _staged("Beta", position=7, wins=8, losses=8, pf=1290.0, pa=1295.0, section=WINNERS),
    ]

    row = research_rows(LEAGUE, [_game(home="Alpha", away="Beta")], table)[0]

    assert row.rank_suppressed is False
    assert row.standings_evidence["home_rank"] == 2.0
    assert row.standings_evidence["away_rank"] == 7.0


def test_the_default_weights_give_rank_no_weight_at_all():
    """So rank is provenance here: the score rides on pct, net_rtg and home court."""

    assert MaGolideConfig().rank == 0.0
    assert MaGolideConfig().pct > 0.0


def test_the_suppression_removes_the_gap_once_rank_is_weighted():
    """Under the defaults this is preventative; a weighted rank would otherwise bite."""

    weighted = MaGolideConfig(rank=1.0)
    cross = [
        _staged("Alpha", position=2, wins=13, losses=3, pf=1300.0, pa=1200.0, section=WINNERS),
        _staged("Beta", position=16, wins=3, losses=13, pf=1150.0, pa=1350.0, section=LOSERS),
    ]
    same = [
        _staged("Alpha", position=2, wins=13, losses=3, pf=1300.0, pa=1200.0, section=WINNERS),
        _staged("Beta", position=16, wins=3, losses=13, pf=1150.0, pa=1350.0, section=WINNERS),
    ]

    a = research_rows(LEAGUE, [_game(home="Alpha", away="Beta")], cross, config=weighted)[0]
    b = research_rows(LEAGUE, [_game(home="Alpha", away="Beta")], same, config=weighted)[0]

    # 16 - 2 = 14 ranks, at weight 1.0, present only where the ladder is one ladder.
    assert b.verdict.score - a.verdict.score == pytest.approx(14.0)
    assert a.rank_suppressed and not b.rank_suppressed


def test_the_section_says_when_ranks_were_not_compared():
    table = [
        _staged("Alpha", position=2, wins=13, losses=3, pf=1300.0, pa=1200.0, section=WINNERS),
        _staged("Beta", position=16, wins=3, losses=13, pf=1150.0, pa=1350.0, section=LOSERS),
    ]
    rows = research_rows(LEAGUE, [_game(home="Alpha", away="Beta")], table)

    text = "\n".join(render_research_section(rows, date_label="2026-11-26", league=LEAGUE))

    assert "different stages/groups" in text
    assert "a rank inside one group is not a rank inside another" in text
