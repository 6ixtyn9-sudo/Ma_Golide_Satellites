from forecasting.feature_manifest import build_feature_use_manifest
from ingestion.models import BookSnapshot, Game, League, Market, Period, Prediction
from ingestion.store import InMemoryStore


def _game(game_id, *, evidence=None):
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2026",
        home_team_id="h",
        away_team_id="a",
        home_name="Home",
        away_name="Away",
        scheduled_utc="2026-11-20T00:00:00Z",
        evidence=evidence or {},
    )


def test_feature_manifest_reports_visible_coverage_and_excludes_late_records():
    store = InMemoryStore()
    game = _game(
        "game-1",
        evidence={
            "espn": {
                "summary": {
                    "observed_utc": "2026-11-19T10:00:00Z",
                    "injuries": [
                        {"team": {"id": "h"}, "injuries": [{"athlete": {"id": "p1"}}]}
                    ],
                    "leaders": [],
                    "context_coverage": {
                        "player_boxscore_teams": 2,
                        "player_boxscore_athletes": 10,
                    },
                }
            }
        },
    )
    store.upsert_game(game)
    store.upsert_prediction(
        Prediction(
            game_id=game.game_id,
            provider="espn",
            model="espn-bpi",
            observed_utc="2026-11-19T11:00:00Z",
            prob_home=0.6,
            prob_away=0.4,
            auxiliary_metrics={"home": {"matchupQuality": 70.0}},
        )
    )
    store.upsert_prediction(
        Prediction(
            game_id=game.game_id,
            provider="hardwood",
            observed_utc="2026-11-20T00:01:00Z",
            prob_home=0.7,
            prob_away=0.3,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id=game.game_id,
            provider="testbook",
            book="Test",
            market=Market.MONEYLINE,
            period=Period.FULL_GAME,
            observed_utc="2026-11-19T12:00:00Z",
            home_price_decimal=1.8,
            away_price_decimal=2.1,
        )
    )

    manifest = build_feature_use_manifest(
        league=League.NBA,
        games=[game],
        upcoming=[game],
        store=store,
        as_of_utc="2026-11-19T23:00:00Z",
        candidate_contexts={
            game.game_id: {
                "injury_snapshot_status": "no_entries_reported",
                "reported_injury_entries": [],
                "teams": {"home": {"historical_boxscore_games": 1}},
            }
        },
    )

    coverage = manifest["visible_upcoming_coverage"]
    assert coverage["predictions"]["games_with_any_pre_tip_prediction"] == 1
    assert coverage["predictions"]["pre_tip_prediction_rows_by_provider"] == {"espn": 1}
    assert coverage["predictions"]["rows_with_auxiliary_metrics"] == 1
    assert coverage["bookmakers"]["games_with_pre_tip_snapshot"] == 1
    assert coverage["espn_player_and_injury_candidates"][
        "games_with_point_in_time_injury_observation"
    ] == 1
    assert manifest["safety"]["candidate_features_used_by_probability_model"] is False
    inventory = manifest["evidence_resource_inventory"]
    assert inventory["provider_resources"]["espn"]["summary"]["games_with_resource"] == 1
    assert inventory["espn_summary_context_families"]["games_with_field"] == {
        "injuries": 1,
        "leaders": 1,
    }
    assert inventory["espn_summary_context_families"]["entry_counts"]["injury_teams"] == 1
    assert inventory["espn_summary_context_families"]["entry_counts"]["injury_entries"] == 1
