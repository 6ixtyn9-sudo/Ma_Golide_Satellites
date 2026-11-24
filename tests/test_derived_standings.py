"""Standings derived from our own finals: arithmetic, with the arithmetic checked.

The scorer in ``forecasting/magolide_score.py`` rates a fixture from standings
because it has no results history to learn from. These tests pin the source of
those standings: wins, losses, percentages, differential, home/away splits, streak
and last-ten, each computed from reconciled finals and each refused when the final
does not reconcile.
"""

from __future__ import annotations


from forecasting.standings import standings_from_finals
from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind


def _final(
    home: str,
    away: str,
    home_points: int,
    away_points: int,
    *,
    tip: str,
    game_id: str = "g",
    quarters: bool = True,
) -> Game:
    """A reconciling final by default; ``quarters=False`` gives a total only."""

    if quarters:
        half_home = home_points // 2
        half_away = away_points // 2
        actual = ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[half_home, home_points - half_home, 0.0, 0.0],
            away_quarters=[half_away, away_points - half_away, 0.0, 0.0],
            home_final=float(home_points),
            away_final=float(away_points),
            source="test",
        )
    else:
        actual = ActualResult(
            period_kind=PeriodKind.UNKNOWN,
            home_quarters=[],
            away_quarters=[],
            home_final=float(home_points),
            away_final=float(away_points),
            source="test",
        )
    return Game(
        game_id=game_id,
        league=League.LITHUANIA_NKL,
        season="2026",
        home_team_id=home.lower(),
        away_team_id=away.lower(),
        home_name=home,
        away_name=away,
        scheduled_utc=tip,
        status=GameStatus.FINAL,
        actual=actual,
    )


def test_wins_losses_percentages_and_the_differential_come_from_the_finals():
    games = [
        _final("Kretinga", "Neptunas 2", 76, 91, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final("Kretinga", "Vytis", 90, 80, tip="2026-10-08T16:00:00Z", game_id="2"),
    ]

    rows = {row.team: row for row in standings_from_finals(games)}

    kretinga = rows["Kretinga"]
    assert kretinga.pct == 0.5
    assert kretinga.rank == 2.0  # 0.5 behind a 1.000 side
    # Points per game differ by (76+90-91-80)/2 = -2.5
    assert kretinga.net_rtg == -2.5
    assert rows["Neptunas 2"].pct == 1.0
    assert rows["Neptunas 2"].rank == 1.0
    assert rows["Vytis"].pct == 0.0


def test_home_and_away_splits_are_counted_separately():
    games = [
        _final("Kretinga", "Vytis", 90, 80, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final("Vytis", "Kretinga", 70, 99, tip="2026-10-08T16:00:00Z", game_id="2"),
    ]

    kretinga = {row.team: row for row in standings_from_finals(games)}["Kretinga"]

    assert kretinga.home_pct == 1.0   # won the home game
    assert kretinga.away_pct == 1.0   # and the away one
    assert kretinga.pct == 1.0


def test_the_streak_is_signed_and_reads_backwards_from_the_latest_game():
    games = [
        _final("Kretinga", "Vytis", 70, 80, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final("Kretinga", "Vytis", 71, 80, tip="2026-10-08T16:00:00Z", game_id="2"),
        _final("Kretinga", "Vytis", 99, 80, tip="2026-10-15T16:00:00Z", game_id="3"),
    ]

    kretinga = {row.team: row for row in standings_from_finals(games)}["Kretinga"]

    # Two losses then a win: the current streak is one win, not two losses.
    assert kretinga.streak == 1


def test_last_ten_is_a_window_not_a_total():
    games = [
        _final(
            "Kretinga", "Vytis", 90 if i % 2 else 70, 80,
            tip=f"2026-10-{i + 1:02d}T16:00:00Z", game_id=f"g{i}",
        )
        for i in range(12)
    ]

    kretinga = {row.team: row for row in standings_from_finals(games)}["Kretinga"]

    assert kretinga.l10_wins is not None and kretinga.l10_losses is not None
    assert kretinga.l10_wins + kretinga.l10_losses == 10


def test_a_final_that_does_not_reconcile_is_not_counted():
    """A totals-only final cannot be graded, so it cannot enter a table either."""

    games = [
        _final("Kretinga", "Vytis", 90, 80, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final(
            "Kretinga", "Vytis", 90, 80,
            tip="2026-10-08T16:00:00Z", game_id="2", quarters=False,
        ),
    ]

    kretinga = {row.team: row for row in standings_from_finals(games)}["Kretinga"]

    assert kretinga.pct == 1.0
    assert kretinga.l10_wins == 1.0  # one countable game, not two


def test_a_club_with_no_finals_is_absent_rather_than_present_at_zero():
    """0-0 would read as a played record; absence is the truth and is sayable."""

    rows = standings_from_finals([])

    assert rows == []


def test_as_of_rebuilds_the_table_as_it_stood_before_a_game():
    games = [
        _final("Kretinga", "Vytis", 90, 80, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final("Kretinga", "Vytis", 60, 80, tip="2026-10-20T16:00:00Z", game_id="2"),
    ]

    before = {row.team: row for row in standings_from_finals(games, as_of_utc="2026-10-10T00:00:00Z")}

    assert before["Kretinga"].pct == 1.0
    assert before["Vytis"].pct == 0.0


def test_rank_breaks_ties_on_points_difference_then_name():
    games = [
        _final("Alpha", "Beta", 90, 80, tip="2026-10-01T16:00:00Z", game_id="1"),
        _final("Gamma", "Delta", 120, 80, tip="2026-10-02T16:00:00Z", game_id="2"),
    ]

    rows = standings_from_finals(games)

    # Both winners are 1.000; Gamma's +40 beats Alpha's +10.
    assert [row.team for row in rows][:2] == ["Gamma", "Alpha"]
    assert [row.rank for row in rows] == [1.0, 2.0, 3.0, 4.0]


def test_the_table_says_which_table_it_is():
    """A derived table must not be mistaken for the competition's own."""

    from forecasting.standings import SOURCE

    rows = standings_from_finals(
        [_final("Kretinga", "Vytis", 90, 80, tip="2026-10-01T16:00:00Z")]
    )

    assert SOURCE == "derived_from_reconciled_finals"
    assert rows[0].team == "Kretinga"
    assert rows[0].pct == 1.0
