from forecasting.feature_manifest import build_feature_use_manifest, _resource_inventory
from ingestion.models import BookSnapshot, Game, League, Market, Period, Prediction
from ingestion.providers.espn import _summary_context
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
                    "field_presence": {
                        "injuries": True,
                        "leaders": True,
                        "last_five_games": False,
                        "against_the_spread": False,
                        "game_info": False,
                        "predictor": False,
                        "team_boxscore": False,
                        "player_boxscore": False,
                    },
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
    assert inventory["espn_summary_context_families"]["games_with_field"]["injuries"] == 1
    assert inventory["espn_summary_context_families"]["games_with_field"]["leaders"] == 1
    assert inventory["espn_summary_context_families"]["field_presence_counts"]["injuries"] == {
        "present": 1,
        "absent": 0,
        "unknown": 0,
    }
    assert inventory["espn_summary_context_families"]["entry_counts"]["injury_teams"] == 1
    assert inventory["espn_summary_context_families"]["entry_counts"]["injury_entries"] == 1


def test_manifest_uses_raw_summary_field_presence_for_materialized_empty_fields(load_fixture):
    payload = load_fixture("espn/summary_ncaaw_401918841_quarters.json")
    game = Game(
        game_id="401918841",
        league=League.NCAA_W,
        season="2026",
        home_team_id="251",
        away_team_id="2458",
        home_name="Texas Longhorns",
        away_name="Northern Colorado Bears",
        scheduled_utc="2026-11-07T01:00:00Z",
    )
    game.evidence = {
        "espn": {
            "summary": _summary_context(
                payload, game, observed_utc="2026-11-07T05:00:00Z"
            )
        }
    }

    inventory = _resource_inventory([game])["espn_summary_context_families"]

    assert inventory["games_with_field"].get("injuries", 0) == 0
    assert inventory["field_presence_counts"]["injuries"] == {
        "present": 0,
        "absent": 1,
        "unknown": 0,
    }
    assert inventory["field_presence_counts"]["last_five_games"] == {
        "present": 0,
        "absent": 1,
        "unknown": 0,
    }


def test_manifest_does_not_claim_unverified_wnba_injury_field_coverage():
    manifest = build_feature_use_manifest(
        league=League.WNBA,
        games=[],
        upcoming=[],
        store=InMemoryStore(),
        as_of_utc="2026-11-20T00:00:00Z",
        candidate_contexts={},
    )

    injury_catalog = manifest["catalog"]["current_injury_status"]
    assert "unverified" in injury_catalog["league_support"]
    assert injury_catalog["league_endpoint_support"] == "not_supported"
