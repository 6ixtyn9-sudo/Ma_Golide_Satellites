from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind
from scripts.mine_espn_player_minutes import mine_games, player_boxscore_rows


def _final_game():
    return Game(
        game_id="historic-1",
        league=League.NBA,
        season="2025",
        home_team_id="h",
        away_team_id="a",
        home_name="Home",
        away_name="Away",
        scheduled_utc="2025-01-01T00:00:00Z",
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[25, 25, 25, 25],
            away_quarters=[20, 20, 20, 20],
            home_final=100,
            away_final=80,
            source="test",
        ),
        evidence={
            "espn": {
                "summary": {
                    "schema_version": "summary-v1",
                    "observed_utc": "2025-01-01T04:30:00Z",
                    "source_modified_utc": None,
                    "source_payload_hash": "abc123",
                    "player_boxscore": [
                        {
                            "team": {"id": "h"},
                            "statistics": [
                                {
                                    "name": "starters",
                                    "labels": ["MIN", "PTS", "REB", "AST"],
                                    "athletes": [
                                        {
                                            "athlete": {"id": "p1", "displayName": "Player One"},
                                            "stats": ["34:30", "24", "8", "6"],
                                        }
                                    ],
                                }
                            ],
                        }
                    ],
                }
            }
        },
    )


def test_offline_player_minutes_mining_preserves_raw_and_conservative_availability():
    game = _final_game()
    rows = player_boxscore_rows(game)
    assert len(rows) == 1
    assert rows[0]["player_id"] == "p1"
    assert rows[0]["side"] == "home"
    assert rows[0]["minutes"] == 34.5
    assert rows[0]["points"] == 24.0
    assert rows[0]["evidence_observed_utc"] == "2025-01-01T04:30:00Z"
    assert rows[0]["evidence_payload_hash"] == "abc123"
    assert rows[0]["available_no_earlier_than_utc"] == "2025-01-01T04:30:00Z"
    assert '"PTS":"24"' in rows[0]["raw_stats_json"]

    mined, coverage = mine_games([game])
    assert len(mined) == 1
    assert coverage == {
        "games_scanned": 1,
        "reconciled_finals": 1,
        "games_with_player_rows": 1,
    }
