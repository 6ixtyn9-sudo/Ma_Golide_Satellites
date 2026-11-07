"""Tests for H2H / recent-form / standings derivations (section 6)."""

import pytest

from ingestion.models import (
    ActualResult,
    Game,
    GameStatus,
    League,
    PeriodKind,
    StandingsRow,
)
from ingestion.derive import (
    STANDINGS_HEADERS,
    TIER2_HEADERS,
    bundle_rows,
    derive_h2h,
    derive_recent,
    standings_layout,
)

UPCOMING_TIP = "2026-01-11T22:30:00Z"


def _game(gid, home, away, when, qh, qa, final=None, ot=None):
    ot_h = ot[0] if ot else 0
    ot_a = ot[1] if ot else 0
    hf = float(sum(qh) + ot_h) if final is None else final[0]
    af = float(sum(qa) + ot_a) if final is None else final[1]
    return Game(
        game_id=gid, league=League.NBA, season="2026",
        home_team_id=home[:2], away_team_id=away[:2],
        home_name=home, away_name=away,
        scheduled_utc=when, status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[float(q) for q in qh],
            away_quarters=[float(q) for q in qa],
            home_overtime=ot[0] if ot else None,
            away_overtime=ot[1] if ot else None,
            home_final=hf, away_final=af,
            overtime_periods=1 if ot else 0,
            source="espn",
        ),
    )


HISTORY = [
    # Target team: Memphis Grizzlies (upcoming home) and Brooklyn Nets (away).
    _game("g1", "Memphis Grizzlies", "Brooklyn Nets", "2026-01-05T22:00:00Z",
          [30, 28, 25, 30], [24, 26, 25, 24]),            # meeting, home win
    _game("g2", "Brooklyn Nets", "Memphis Grizzlies", "2026-01-02T22:00:00Z",
          [22, 30, 28, 24], [28, 26, 30, 34]),            # meeting, away win
    _game("g3", "Memphis Grizzlies", "Miami Heat", "2026-01-08T22:00:00Z",
          [31, 25, 24, 27], [26, 24, 30, 22]),            # recent home venue
    _game("g4", "Charlotte Hornets", "Memphis Grizzlies", "2026-01-09T22:00:00Z",
          [25, 27, 23, 24], [28, 25, 26, 22]),            # recent away venue
    _game("g5", "Brooklyn Nets", "Toronto Raptors", "2026-01-07T22:00:00Z",
          [26, 25, 24, 27], [28, 22, 25, 24]),            # away team's other game
    _game("g6", "New York Knicks", "Boston Celtics", "2026-01-06T22:00:00Z",
          [24, 26, 28, 25], [26, 24, 27, 30]),            # unrelated
    # Future / current game must be excluded from history.
    _game("g7", "Memphis Grizzlies", "Brooklyn Nets", UPCOMING_TIP,
          [25, 25, 25, 25], [24, 24, 24, 24]),
    # Unreconciled game quarantined out of history.
    _game("g8", "Memphis Grizzlies", "Atlanta Hawks", "2026-01-04T22:00:00Z",
          [20, 20, 20, 20], [30, 30, 30, 30], final=(80, 130)),
]


def test_h2h_only_meetings_newest_first():
    rows = derive_h2h(HISTORY, "Memphis Grizzlies", "Brooklyn Nets", UPCOMING_TIP)
    assert [r[0] for r in rows] == ["2026-01-05", "2026-01-02"]  # newest first
    assert rows[0][1] == "Memphis Grizzlies" and rows[0][2] == "Brooklyn Nets"
    assert rows[1][1] == "Brooklyn Nets" and rows[1][2] == "Memphis Grizzlies"


def test_h2h_winner_relative_to_row_orientation():
    rows = derive_h2h(HISTORY, "Memphis Grizzlies", "Brooklyn Nets", UPCOMING_TIP)
    assert rows[0][13] == "H"   # g1: home Memphis won
    assert rows[1][13] == "A"   # g2: away Memphis won
    assert rows[0][16] == "113-99"
    assert rows[1][16] == "104-118"


def test_h2h_excludes_upcoming_and_future_games():
    rows = derive_h2h(HISTORY, "Memphis Grizzlies", "Brooklyn Nets", UPCOMING_TIP)
    assert all(r[0] < "2026-01-11" for r in rows)
    assert len(rows) == 2


def test_h2h_excludes_unreconciled_games():
    rows = derive_h2h(HISTORY, "Memphis Grizzlies", "Atlanta Hawks", UPCOMING_TIP)
    assert rows == []  # g8 quarters sum 80-120 but final 80-130: quarantined


def test_recent_includes_both_venues_and_no_duplicates():
    rows = derive_recent(HISTORY, "Memphis Grizzlies", UPCOMING_TIP)
    dates = [r[0] for r in rows]
    assert dates == ["2026-01-09", "2026-01-08", "2026-01-05", "2026-01-02"]
    venues = {(r[1], r[2]) for r in rows}
    assert ("Charlotte Hornets", "Memphis Grizzlies") in venues  # away venue included
    assert ("Memphis Grizzlies", "Miami Heat") in venues         # home venue included


def test_recent_for_upcoming_away_team():
    rows = derive_recent(HISTORY, "Brooklyn Nets", UPCOMING_TIP)
    matchups = {(r[1], r[2]) for r in rows}
    assert ("Brooklyn Nets", "Toronto Raptors") in matchups
    assert ("Memphis Grizzlies", "Brooklyn Nets") in matchups


def test_bundle_rows_shape_and_headers():
    bundles = bundle_rows(HISTORY, "Memphis Grizzlies", "Brooklyn Nets", UPCOMING_TIP)
    assert set(bundles) == {"H2H", "RecentHome", "RecentAway"}
    for rows in bundles.values():
        for r in rows:
            assert len(r) == len(TIER2_HEADERS)


def test_t2_row_ot_columns_and_sum_mismatch():
    ot_game = _game("ot", "Team X", "Team Y", "2026-01-03T22:00:00Z",
                    [28, 26, 31, 25], [29, 28, 26, 28], ot=(12.0, 7.0))
    rows = derive_h2h([ot_game], "Team X", "Team Y", "2026-01-10T00:00:00Z")
    r = rows[0]
    assert r[11] == 12 and r[12] == 7          # OTH, OTA
    assert r[13] == "H"
    assert r[16] == "122-118"
    assert r[17] == ""                          # no mismatch


def test_depth_cap():
    history = [
        _game(f"m{i}", "Memphis Grizzlies", f"Opponent {i}",
              f"2026-01-{i+1:02d}T22:00:00Z", [30, 30, 30, 30], [25, 25, 25, 25])
        for i in range(1, 10)
    ]
    rows = derive_recent(history, "Memphis Grizzlies", "2026-01-20T00:00:00Z", depth=4)
    assert len(rows) == 4
    assert rows[0][0] == "2026-01-10"  # newest first


# ---------------------------------------------------------------------------
# Standings layout
# ---------------------------------------------------------------------------


def test_standings_layout_sections_and_cumulative_pf():
    rows = [
        StandingsRow(team_name="Chicago Bulls", position=14, games_played=8,
                     wins=1, losses=7, points_for=893.0, points_against=1037.0,
                     win_pct=0.125, streak=-6, section="East"),
        StandingsRow(team_name="Indiana Pacers", position=1, games_played=9,
                     wins=6, losses=3, points_for=1100.0, points_against=1080.0,
                     win_pct=0.667, streak=3, section="East"),
        StandingsRow(team_name="Denver Nuggets", position=2, games_played=9,
                     wins=6, losses=3, points_for=1150.0, points_against=1100.0,
                     win_pct=0.667, streak=2, section="West"),
    ]
    layout = standings_layout(rows)
    assert layout[0] == ["East"]
    assert layout[1] == list(STANDINGS_HEADERS)
    assert layout[2][1] == "Indiana Pacers"
    assert layout[2][5] == 1100  # cumulative PF, NOT a per-game average
    assert layout[4] == ["West"]
    # PF/PA sanity: cumulative values must exceed plausible per-game means.
    for row in layout:
        if len(row) == len(STANDINGS_HEADERS) and isinstance(row[5], (int, float)):
            assert row[5] > 3 * row[2]
