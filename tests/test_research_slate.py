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
    confidence_tier,
    decide_bet,
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
    # Ma Golide still made a bet: with a market read, the dog is the bet.
    assert rows[0].kind == "ROBBER"
    assert rows[0].selection == "AWAY"  # 6.50 is the longer price


def test_one_absent_club_is_scored_against_the_known_row_not_filtered_out():
    """A thin table is not a reason to hide a fixture — only to label it."""

    rows = research_rows(LEAGUE, [_game(home="Unknown BC")], TABLE, snapshots=[_snapshot()])

    assert rows[0].standings_evidence["neutral_defaults_applied"] is True
    assert rows[0].selection == "HOME"  # against a 1-15 club the neutral row still leans home
    assert rows[0].verdict.prediction == "HOME"


def test_a_risky_fixture_takes_the_dog_rather_than_no_bet():
    """The operator's rule, 2026-11-26: no model edge still means a bet."""

    rows = research_rows(
        LEAGUE,
        [_game(home="Unknown BC", away="Also Unknown")],
        TABLE,
        snapshots=[_snapshot(home="Unknown BC", away="Also Unknown")],
    )

    assert rows[0].observed is True
    assert rows[0].priced is True
    assert rows[0].has_selection is True
    assert rows[0].verdict.prediction == "RISKY"
    assert rows[0].kind == "ROBBER"
    assert rows[0].basis == "risky_tier_market_dog"
    assert rows[0].decimal_price == 6.50
    assert len(research_tickets(rows)) == 1


def test_the_one_case_with_no_bet_names_itself():
    """No table row for either club and no market: nothing to be a banker about."""

    rows = research_rows(
        LEAGUE, [_game(home="Unknown BC", away="Also Unknown")], TABLE
    )

    assert rows[0].kind == "NO_EVIDENCE"
    assert rows[0].selection is None
    assert rows[0].basis == "risky_tier_no_readable_table_and_no_price"
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
    assert "2 fixture(s), 2 banker(s), 0 robber(s), 1 priced" in lines[1]
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


# -- banker and robber -------------------------------------------------------


def test_a_scored_fixture_is_a_banker():
    rows = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])

    assert rows[0].verdict.prediction == "HOME"
    assert rows[0].tier == "STRONG"
    assert rows[0].kind == "BANKER"
    assert rows[0].basis == "strong_tier_model_pick"
    assert rows[0].selection == "HOME"


def test_a_risky_unpriced_fixture_backs_the_table_side():
    table = [
        _published("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
        _published("Rytas 2", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
    ]
    # Force RISKY by making the two clubs equal on every input the score uses.
    equal = [
        _published("Telsiai", position=1, played=16, wins=8, losses=8, pf=1300.0, pa=1300.0),
        _published("Rytas 2", position=1, played=16, wins=8, losses=8, pf=1300.0, pa=1300.0),
    ]

    ties = research_rows(LEAGUE, [_game()], equal)
    real = research_rows(LEAGUE, [_game()], table)

    assert ties[0].verdict.prediction == "RISKY"
    # Level on every field this reads, with no market to break the tie: a measured
    # draw, named as such rather than dressed up as a bet.
    assert ties[0].kind == "NO_EVIDENCE"
    assert ties[0].basis == "risky_tier_table_level_and_no_price"
    assert real[0].kind == "BANKER"


def test_the_dog_is_the_longer_price_even_when_the_table_favours_the_home_side():
    rows = research_rows(
        LEAGUE,
        [_game(home="Unknown BC", away="Also Unknown")],
        TABLE,
        snapshots=[_snapshot(home="Unknown BC", away="Also Unknown",
                             home_price=4.00, away_price=1.25)],
    )

    assert rows[0].kind == "ROBBER"
    assert rows[0].selection == "HOME"  # 4.00 pays for it
    assert rows[0].decimal_price == 4.00
    assert rows[0].basis == "risky_tier_market_dog"


def test_a_lone_price_cannot_name_a_dog():
    close = [
        _published("Telsiai", position=1, played=16, wins=8, losses=8, pf=1300.0, pa=1300.0),
        _published("Rytas 2", position=2, played=17, wins=8, losses=9, pf=1310.0, pa=1320.0),
    ]
    lone = _snapshot()
    lone.away_price_decimal = None

    rows = research_rows(LEAGUE, [_game()], close, snapshots=[lone])

    assert rows[0].verdict.prediction == "RISKY"
    # One price cannot say which side is longer, so the dog comes off the table:
    # Rytas 2 at 8-9 is the weaker row, and a RISKY fixture is a robber's fixture.
    assert rows[0].kind == "ROBBER"
    assert rows[0].selection == "AWAY"
    assert rows[0].basis == "risky_tier_table_weak_side"


def test_every_fixture_with_a_readable_table_gets_a_bet():
    table = [
        _published("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
        _published("Ritas", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
    ]
    games = [
        _game(game_id="a"),
        _game(game_id="b", home="Telsiai", away="Nobody"),
        _game(game_id="c", home="Nobody", away="Nobody Else"),
    ]
    snapshots = [
        _snapshot(game_id="a"),
        _snapshot(game_id="b", home="Telsiai", away="Nobody"),
        _snapshot(game_id="c", home="Nobody", away="Nobody Else"),
    ]

    rows = research_rows(LEAGUE, games, table, snapshots=snapshots)

    assert len(rows) == 3
    assert all(row.kind in {"BANKER", "ROBBER"} for row in rows)


def test_the_card_states_the_rule_and_counts_the_kinds():
    rows = research_rows(
        LEAGUE,
        [_game(), _game(game_id="b", home="Unknown BC", away="Also Unknown")],
        TABLE,
        snapshots=[_snapshot(), _snapshot(game_id="b", home="Unknown BC",
                                          away="Also Unknown")],
    )
    text = "\n".join(render_research_section(rows, date_label="2026-11-26", league=LEAGUE))

    assert "1 banker(s), 1 robber(s)" in text
    assert "Every fixture gets a bet" in text
    assert "BANKER = STRONG on either side, or MEDIUM on the home side" in text
    assert "ROBBER = an AWAY pick that only reached MEDIUM" in text
    assert "kind=" not in text  # the kind leads each line instead
    assert "BANKER · Telsiai vs Rytas 2" in text


# -- the tiers the retired scorer used --------------------------------------
#
# Thresholds and the banker/robber rule are re-read from source, not inferred:
# ``docs/Forecaster.gs`` (``getConfidenceTier``, ``evaluateConfig`` and the
# ``tier_*_min_score`` keys) and ``docs/Contract_Enforcement.gs`` (the BANKERS
# section built from ``bankerThreshold || 60``), both recovered from ``de966196^``.


def _verdict(score: float, prediction: str, *, threshold: float = 25.0) -> MaGolideVerdict:
    return MaGolideVerdict(
        score=score,
        prediction=prediction,
        confidence=70.0,
        uncalibrated_logistic_pct=60.0,
        factor_breakdown={},
        meta={"threshold": threshold},
    )


@pytest.mark.parametrize(
    "score,expected",
    [
        (75.0, "STRONG"),
        (74.99, "MEDIUM"),
        (60.0, "MEDIUM"),
        (59.99, "WEAK"),
        (50.0, "WEAK"),
        (49.99, "LOW"),
        (25.0, "LOW"),
        (24.99, "RISKY"),
        (-75.0, "STRONG"),
        (-24.99, "RISKY"),
    ],
)
def test_the_tier_ladder_is_the_retired_scripts(score, expected):
    assert confidence_tier(score) == expected


def test_the_risky_floor_is_the_threshold_that_scored_the_game():
    """``config.threshold`` is the RISKY boundary, so a lane's override moves it."""

    assert confidence_tier(40.0) == "LOW"
    assert confidence_tier(40.0, risky_below=45.0) == "RISKY"


def test_a_strong_away_pick_is_a_banker_too():
    """``Bankers = STRONG-tier picks`` -- the script does not care which side."""

    selection, kind, basis = decide_bet(_verdict(-80.0, "AWAY"), None, None, None)

    assert (selection, kind, basis) == ("AWAY", "BANKER", "strong_tier_model_pick")


def test_a_medium_home_pick_clears_the_banker_gate():
    selection, kind, basis = decide_bet(_verdict(65.0, "HOME"), None, None, None)

    assert (selection, kind, basis) == ("HOME", "BANKER", "medium_tier_model_pick")


def test_a_medium_away_pick_is_the_scripts_robber():
    """``Robbers = AWAY (upset) picks that cleared MEDIUM threshold``."""

    selection, kind, basis = decide_bet(_verdict(-65.0, "AWAY"), None, None, None)

    assert (selection, kind, basis) == ("AWAY", "ROBBER", "medium_tier_away_upset")


def test_a_weak_home_pick_takes_the_market_dog():
    """Below the gate there is no edge to bank, so the robber buys the price."""

    selection, kind, basis = decide_bet(
        _verdict(55.0, "HOME"), None, None, _snapshot(home_price=1.20, away_price=5.00)
    )

    assert (selection, kind, basis) == ("AWAY", "ROBBER", "weak_tier_market_dog")


def test_a_low_tier_fixture_with_nothing_to_read_takes_the_opposite_side():
    """Last resort: a named side means the other side is the dog."""

    selection, kind, basis = decide_bet(_verdict(35.0, "HOME"), None, None, None)

    assert (selection, kind, basis) == ("AWAY", "ROBBER", "low_tier_opposite_the_model_pick")


def test_the_line_prints_the_tier_beside_the_prediction():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0]

    assert "tier=STRONG" in row.as_line()
    assert row.tier == "STRONG"


# -- the rest of the retired inventory --------------------------------------
#
# Seven bet types existed (BANKER, ROBBER, FIRST_HALF_1X2, FT_OU, SNIPER_MARGIN,
# SNIPER_OU, HIGHEST_QUARTER). Three need per-quarter splits a standings table
# cannot give, so they are not offered; the other two appear here, and an
# unavailable one names the missing input rather than printing "no bet".


def _totals(
    *,
    game_id: str = "bes:McCqxqs3",
    home: str = "Telsiai",
    away: str = "Rytas 2",
    line: float = 160.0,
    over: float | None = 1.90,
    under: float | None = 1.92,
) -> BookSnapshot:
    return BookSnapshot(
        game_id=game_id,
        provider="betexplorer",
        book="average",
        market=Market.TOTAL,
        period=Period.FULL_GAME,
        observed_utc="2026-11-25T18:00:00Z",
        line=line,
        over_price_decimal=over,
        under_price_decimal=under,
        price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
        home_name=home,
        away_name=away,
        scheduled_utc="2026-12-01T17:00:00Z",
    )


def _proposal(row, bet_type: str):
    return next(p for p in row.bets if p.bet_type == bet_type)


def test_a_fixture_with_no_total_line_names_the_missing_input():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0]

    totals = _proposal(row, "FT_OU")

    assert totals.available is False
    assert totals.reason == "no_totals_line"
    assert "ft_ou=UNAVAILABLE(no_totals_line)" in row.as_line()


def test_an_observed_total_line_produces_an_over_under_bet():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals()])[0]

    totals = _proposal(row, "FT_OU")

    assert totals.available is True
    assert totals.selection == "OVER"  # both clubs average ~166 total against a 160 line
    assert totals.line == 160.0
    assert totals.price == 1.90
    assert totals.confidence > 55.0
    assert totals.basis == "model_total_shrunk_to_observed_line"


def test_a_total_line_without_an_under_price_is_not_a_bet():
    row = research_rows(
        LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals(under=None)]
    )[0]

    assert _proposal(row, "FT_OU").reason == "no_over_under_prices"


def test_a_sharp_line_on_a_strong_model_still_has_to_clear_the_edge_gate():
    """A line priced so that the model's edge vanishes is not a bet."""

    # 1.01 over / 30.0 under: implied 0.99, no model total clears that edge.
    row = research_rows(
        LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals(over=1.01, under=30.0)]
    )[0]

    totals = _proposal(row, "FT_OU")

    assert totals.available is False
    assert totals.reason == "edge_below_ftou_gate"


def test_a_first_half_bet_is_derived_from_the_score():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0]

    first_half = _proposal(row, "FIRST_HALF_1X2")

    assert first_half.available is True
    assert first_half.selection == "HOME"
    assert first_half.basis == "h1_signal_from_magolide_score"
    assert 51.0 <= first_half.confidence <= 85.0
    assert "first_half_1x2=HOME" in row.as_line()


def test_the_first_half_side_follows_the_score_when_the_dog_is_strong():
    table = [
        _published("Telsiai", position=1, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
        _published("Rytas 2", position=18, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
    ]

    row = research_rows(LEAGUE, [_game()], table)[0]
    first_half = _proposal(row, "FIRST_HALF_1X2")

    assert row.verdict.prediction == "AWAY"
    assert first_half.available is True
    assert first_half.selection == "AWAY"


def test_a_coin_flip_has_no_first_half_bet():
    level = [
        _published("Telsiai", position=1, played=16, wins=8, losses=8, pf=1300.0, pa=1300.0),
        _published("Rytas 2", position=2, played=16, wins=8, losses=8, pf=1300.0, pa=1300.0),
    ]

    row = research_rows(LEAGUE, [_game()], level)[0]

    first_half = _proposal(row, "FIRST_HALF_1X2")
    assert first_half.available is False
    assert first_half.reason == "confidence_below_first_half_gate"


def test_a_disabled_bet_type_names_its_flag():
    from forecasting.bet_config import default_config

    off = default_config().with_overrides(
        {"enableFirstHalf": False, "enableFTOU": False}
    )

    row = research_rows(
        LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals()], bet_config=off
    )[0]

    assert _proposal(row, "FIRST_HALF_1X2").reason == "enableFirstHalf_off"
    assert _proposal(row, "FT_OU").reason == "enableFTOU_off"


def test_every_row_carries_the_config_stamp_that_produced_it():
    from forecasting.bet_config import default_config

    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot()])[0]

    assert row.config_stamp == default_config().stamp_id
    assert f"config_stamp={row.config_stamp}" in row.as_line()
    assert row.as_dict()["config_stamp"] == row.config_stamp


def test_an_override_gate_changes_the_stamp_on_the_row():
    from forecasting.bet_config import default_config

    strict = default_config().with_overrides({"firstHalfMinConf": 99.0})
    row = research_rows(LEAGUE, [_game()], TABLE, bet_config=strict)[0]

    assert row.config_stamp == strict.stamp_id
    assert _proposal(row, "FIRST_HALF_1X2").available is False


def test_the_card_counts_the_bet_inventory_it_holds():
    rows = research_rows(
        LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals()]
    )
    text = "\n".join(render_research_section(rows, date_label="2026-12-01", league=LEAGUE))

    assert "1 first-half, 1 totals" in text
    assert "Bet inventory per fixture" in text
    assert "Quarter SNIPER and HIGHEST_QUARTER need" in text
    assert "Bet config: CFG_" in text


def test_a_slip_leg_names_its_bet_type_and_each_fixture_travels_once():
    """One bet per match in the acca, and the leg says which bet it is."""

    snapshots = [_snapshot(game_id="g1"), _snapshot(game_id="g2", home_price=1.40, away_price=4.00)]
    rows = research_rows(LEAGUE, _two_games(), TABLE, snapshots=snapshots)

    slip = [ticket for ticket in research_tickets(rows) if len(ticket.legs) == 2][0]
    text = "\n".join(slip.as_lines())

    assert "leg: BANKER" in text
    assert len(slip.legs) == len({leg.game_id for leg in slip.legs})


def test_a_fixture_carrying_extra_bet_types_still_lends_one_leg():
    row = research_rows(LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals()])[0]

    assert len(row.bets) == 3  # moneyline + first half + totals
    assert len(research_tickets([row])[0].legs) == 1


# -- the tuner's league discount has to bite --------------------------------
#
# Config_Tuner.gs advised ``league_weight_<LEAGUE> = 0.7`` for a league under 50%.
# A weight the card ignores would be a decorative number, so it multiplies the
# confidence the gates compare -- and shows as ``lw=`` on the line.


def test_a_discounted_league_lowers_the_first_half_confidence():
    from forecasting.bet_config import default_config

    discounted = default_config().with_leagues({LEAGUE: 0.9})
    plain = research_rows(LEAGUE, [_game()], TABLE)[0]
    cut = research_rows(LEAGUE, [_game()], TABLE, bet_config=discounted)[0]

    plain_bet = _proposal(plain, "FIRST_HALF_1X2")
    cut_bet = _proposal(cut, "FIRST_HALF_1X2")

    assert cut_bet.available is True  # 79.2 * 0.9 still clears the 58 gate
    assert cut_bet.league_weight == 0.9
    # Each is rounded to 0.1 for printing, so compare within one print step.
    assert cut_bet.confidence == pytest.approx(plain_bet.confidence * 0.9, abs=0.1)
    assert "lw=0.90" in cut.as_line()


def test_a_discount_can_take_a_borderline_first_half_bet_below_its_gate():
    from forecasting.bet_config import default_config

    heavy = default_config().with_leagues({LEAGUE: 0.6})
    row = research_rows(LEAGUE, [_game()], TABLE, bet_config=heavy)[0]

    first_half = _proposal(row, "FIRST_HALF_1X2")
    assert first_half.available is False
    assert first_half.reason == "confidence_below_first_half_gate"
    # The line still names the discount that took it under, so the cause is visible.
    assert "lw=0.60" in row.as_line()


def test_a_discount_can_take_a_totals_bet_below_its_gate():
    from forecasting.bet_config import default_config

    heavy = default_config().with_leagues({LEAGUE: 0.6})
    row = research_rows(
        LEAGUE, [_game()], TABLE, snapshots=[_snapshot(), _totals()], bet_config=heavy
    )[0]

    assert _proposal(row, "FT_OU").reason == "confidence_below_ftou_gate"


def test_an_undiscounted_league_carries_no_weight_token():
    row = research_rows(LEAGUE, [_game()], TABLE)[0]

    assert _proposal(row, "FIRST_HALF_1X2").league_weight == 1.0
    assert "lw=" not in row.as_line()


def test_the_scorer_runs_under_the_bet_configs_own_numbers():
    """Otherwise the tuner's ``confidence_scale`` would move a stamp and nothing else."""

    from forecasting.bet_config import default_config

    # A close table, so the probability is not pinned at the confMax ceiling where
    # every scale produces the same clamped answer.
    close = [
        _published("Telsiai", position=1, played=16, wins=9, losses=7, pf=1300.0, pa=1290.0),
        _published("Rytas 2", position=6, played=16, wins=8, losses=8, pf=1295.0, pa=1296.0),
    ]
    shrunk = default_config().with_overrides({"confidence_scale": 60.0, "threshold": 30.0})

    plain = research_rows(LEAGUE, [_game()], close)[0]
    tuned = research_rows(LEAGUE, [_game()], close, bet_config=shrunk)[0]

    assert tuned.verdict.meta["scale_factor"] == 60.0
    assert tuned.verdict.meta["threshold"] == 30.0
    # A wider logistic scale means the same score implies a softer probability.
    assert tuned.verdict.uncalibrated_logistic_pct < plain.verdict.uncalibrated_logistic_pct


def test_a_pinned_scorer_config_wins_over_the_bet_config():
    from forecasting.bet_config import default_config

    pinned = MaGolideConfig(confidence_scale=30.0)
    shrunk = default_config().with_overrides({"confidence_scale": 60.0})

    row = research_rows(
        LEAGUE, [_game()], TABLE, config=pinned, bet_config=shrunk
    )[0]

    assert row.verdict.meta["scale_factor"] == 30.0
