from forecasting.espn_context import build_espn_player_context
from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind


def _game(game_id, scheduled_utc, *, evidence=None, actual=None):
    return Game(
        game_id=game_id,
        league=League.NCAA_M,
        season="2026",
        home_team_id="home-id",
        away_team_id="away-id",
        home_name="Home Team",
        away_name="Away Team",
        scheduled_utc=scheduled_utc,
        status=GameStatus.FINAL if actual else GameStatus.SCHEDULED,
        actual=actual,
        evidence=evidence or {},
    )


def _final():
    return ActualResult(
        period_kind=PeriodKind.HALVES,
        home_quarters=[40, 40],
        away_quarters=[35, 35],
        home_final=80,
        away_final=70,
        source="test",
    )


def _summary(observed_utc, *, players=None, injuries=None):
    result = {"observed_utc": observed_utc}
    if players is not None:
        result["player_boxscore"] = players
    if injuries is not None:
        result["injuries"] = injuries
    return {"espn": {"summary": result}}


def test_player_context_uses_only_observed_prior_boxscores_and_pretip_injuries():
    previous = _game(
        "old-game",
        "2026-11-01T00:00:00Z",
        actual=_final(),
        evidence=_summary(
            "2026-11-01T05:00:00Z",
            players=[
                {
                    "team": {"id": "home-id"},
                    "statistics": [
                        {
                            "labels": ["MIN", "PTS", "REB", "AST"],
                            "athletes": [
                                {
                                    "athlete": {"id": "player-1", "displayName": "Player One"},
                                    "stats": ["30:30", "18", "5", "4"],
                                }
                            ],
                        }
                    ],
                }
            ],
        ),
    )
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence=_summary(
            "2026-11-02T10:00:00Z",
            injuries=[
                {
                    "team": {"id": "home-id"},
                    "injuries": [
                        {
                            "athlete": {"id": "player-1", "displayName": "Player One"},
                            "status": "Out",
                        }
                    ],
                }
            ],
        ),
    )

    features = build_espn_player_context(
        current,
        [previous],
        as_of_utc="2026-11-02T12:00:00Z",
    )

    assert features["feature_role"] == "shadow_candidate_only_not_used_to_change_probability"
    assert features["injury_snapshot_status"] == "snapshot_present"
    assert features["teams"]["home"]["historical_boxscore_games"] == 1
    assert features["teams"]["home"]["recent_players"][0]["mean_minutes"] == 30.5
    assert features["teams"]["home"]["recent_minutes_for_reported_out_players"] == 30.5
    assert features["teams"]["away"]["historical_boxscore_games"] == 0


def test_player_context_rejects_late_evidence_and_never_treats_missing_injuries_as_healthy():
    late_previous = _game(
        "late-old-game",
        "2026-11-01T00:00:00Z",
        actual=_final(),
        evidence=_summary(
            "2026-11-03T00:01:00Z",
            players=[
                {
                    "team": {"id": "home-id"},
                    "statistics": [
                        {
                            "labels": ["MIN", "PTS"],
                            "athletes": [
                                {
                                    "athlete": {"id": "player-1"},
                                    "stats": ["30", "18"],
                                }
                            ],
                        }
                    ],
                }
            ],
        ),
    )
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence=_summary("2026-11-02T10:00:00Z"),
    )

    features = build_espn_player_context(
        current,
        [late_previous],
        as_of_utc="2026-11-02T12:00:00Z",
    )

    assert features["injury_snapshot_status"] == "field_absent"
    assert features["reported_injury_entries"] == []
    assert features["teams"]["home"]["historical_boxscore_games"] == 0
    assert features["teams"]["home"]["minutes_feature_status"] == "unavailable"


def test_player_context_keeps_only_point_in_time_form_and_ats_candidates():
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence={
            "espn": {
                "summary": {
                    "observed_utc": "2026-11-02T10:00:00Z",
                    "source_payload_hash": "summary-hash",
                    "last_five_games": [
                        {
                            "team": {"id": "home-id"},
                            "events": [
                                {
                                    "id": "older-game",
                                    "gameDate": "2026-11-01T00:00:00Z",
                                    "gameResult": "W",
                                    "score": "100-90",
                                    "opponent": {"id": "opponent-1", "displayName": "Other"},
                                },
                                {
                                    "id": "current-game",
                                    "gameDate": "2026-11-03T00:00:00Z",
                                    "gameResult": "W",
                                },
                                {
                                    "id": "future-game",
                                    "gameDate": "2026-11-04T00:00:00Z",
                                    "gameResult": "L",
                                },
                            ],
                        }
                    ],
                    "against_the_spread": [
                        {
                            "team": {"id": "home-id"},
                            "records": [{"summary": "4-5", "type": "overall"}],
                        }
                    ],
                }
            }
        },
    )

    features = build_espn_player_context(
        current,
        [],
        as_of_utc="2026-11-02T12:00:00Z",
    )

    candidate = features["teams"]["home"]["recent_form_candidate"]
    assert candidate["status"] == "available"
    assert [row["event_id"] for row in candidate["recent_results"]] == ["older-game"]
    assert candidate["against_the_spread"] == [{"summary": "4-5", "type": "overall"}]
    assert candidate["source_payload_hash"] == "summary-hash"
    assert features["teams"]["away"]["recent_form_candidate"]["recent_results"] == []


def test_future_dated_injury_rows_are_not_used_even_in_an_earlier_snapshot():
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence=_summary(
            "2026-11-02T10:00:00Z",
            injuries=[
                {
                    "team": {"id": "home-id"},
                    "injuries": [
                        {
                            "athlete": {"id": "player-1"},
                            "status": "Out",
                            "date": "2026-11-04T00:00:00Z",
                        }
                    ],
                }
            ],
        ),
    )

    features = build_espn_player_context(current, [], as_of_utc="2026-11-02T12:00:00Z")

    assert features["injury_snapshot_status"] == "partial_snapshot_no_entries"
    assert features["reported_injury_entries"] == []


def test_context_minutes_parser_rejects_seconds_scale_corruption():
    from forecasting.espn_context import _minutes

    assert _minutes("2052") is None
    assert _minutes(2052) is None
    assert _minutes("100:01") is None
    assert _minutes("99:59") == 99 + 59 / 60


def test_injury_status_merge_uses_latest_observation_not_fixed_source_order():
    def evidence(summary_time, endpoint_time, summary_status, endpoint_status):
        injury_summary = _summary(
            summary_time,
            injuries=[
                {"team": {"id": "home-id"}, "injuries": [{
                    "athlete": {"id": "player-1"}, "status": summary_status
                }]},
                {"team": {"id": "away-id"}, "injuries": []},
            ],
        )
        espn = injury_summary["espn"]
        espn["league_injuries"] = {
            "observed_utc": endpoint_time,
            "team_coverage_complete": True,
            "teams": {
                "home-id": [{"athlete_id": "player-1", "status": endpoint_status}],
                "away-id": [],
            },
        }
        return injury_summary

    older_endpoint = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence=evidence(
            "2026-11-02T11:00:00Z",
            "2026-11-02T10:00:00Z",
            "Out",
            "Questionable",
        ),
    )
    newer_endpoint = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence=evidence(
            "2026-11-02T10:00:00Z",
            "2026-11-02T11:00:00Z",
            "Questionable",
            "Out",
        ),
    )

    summary_wins = build_espn_player_context(
        older_endpoint, [], as_of_utc="2026-11-02T12:00:00Z"
    )["reported_injury_entries"][0]
    endpoint_wins = build_espn_player_context(
        newer_endpoint, [], as_of_utc="2026-11-02T12:00:00Z"
    )["reported_injury_entries"][0]

    assert summary_wins["status"] == "Out"
    assert summary_wins["status_source"] == "summary"
    assert summary_wins["status_observed_utc"] == "2026-11-02T11:00:00Z"
    assert endpoint_wins["status"] == "Out"
    assert endpoint_wins["status_source"] == "league_endpoint"
    assert endpoint_wins["status_observed_utc"] == "2026-11-02T11:00:00Z"
    assert len(endpoint_wins["source_observations"]) == 2


def test_normalized_empty_keys_respect_raw_field_presence_flags():
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence={
            "espn": {
                "summary": {
                    "observed_utc": "2026-11-02T10:00:00Z",
                    "field_presence": {
                        "injuries": False,
                        "last_five_games": False,
                        "against_the_spread": False,
                    },
                    "injuries": [],
                    "last_five_games": [],
                    "against_the_spread": [],
                }
            }
        },
    )

    features = build_espn_player_context(current, [], as_of_utc="2026-11-02T12:00:00Z")

    assert features["injury_snapshot_status"] == "field_absent"
    assert features["teams"]["home"]["recent_form_candidate"]["recent_results_status"] == "field_absent"
    assert features["teams"]["home"]["recent_form_candidate"]["against_the_spread_status"] == "field_absent"


def test_same_timestamp_injury_status_conflict_is_reported_without_arbitrary_winner():
    summary = _summary(
        "2026-11-02T11:00:00Z",
        injuries=[
            {"team": {"id": "home-id"}, "injuries": [{
                "athlete": {"id": "player-1"}, "status": "Out"
            }]},
            {"team": {"id": "away-id"}, "injuries": []},
        ],
    )["espn"]["summary"]
    current = _game(
        "current-game",
        "2026-11-03T00:00:00Z",
        evidence={
            "espn": {
                "summary": summary,
                "league_injuries": {
                    "observed_utc": "2026-11-02T11:00:00Z",
                    "team_coverage_complete": True,
                    "teams": {
                        "home-id": [{"athlete_id": "player-1", "status": "Questionable"}],
                        "away-id": [],
                    },
                },
            }
        },
    )

    row = build_espn_player_context(
        current, [], as_of_utc="2026-11-02T12:00:00Z"
    )["reported_injury_entries"][0]

    assert row["status"] is None
    assert row["status_conflict"] is True
    assert row["status_source"] == "same_timestamp_conflict"
    assert {item["status"] for item in row["source_observations"]} == {"Out", "Questionable"}
