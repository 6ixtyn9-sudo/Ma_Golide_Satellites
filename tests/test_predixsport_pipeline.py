"""Predixsport slate acquisition is bounded and identity-safe."""

from __future__ import annotations

from ingestion.models import Game, GameStatus, League, Prediction
from ingestion.pipeline import Pipeline, PipelineConfig, _match_predixsport_prediction
from ingestion.providers import predixsport
from ingestion.store import InMemoryStore


FAKE_KEY = "px-dummy-test-key-not-a-real-credential"


def _game(game_id: str, home: str, away: str, tip: str) -> Game:
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2026",
        home_team_id=f"{game_id}-home",
        away_team_id=f"{game_id}-away",
        home_name=home,
        away_name=away,
        scheduled_utc=tip,
        status=GameStatus.SCHEDULED,
    )


def _factory(payload: dict, calls: list[str]):
    def factory(provider: str):
        assert provider == "predixsport"

        def fetch(url: str):
            calls.append(url)
            return payload

        return fetch

    return factory


def test_predixsport_slate_is_fetched_once_and_joined_by_teams_and_utc_tip(monkeypatch):
    monkeypatch.setenv(predixsport.ENV_KEY, FAKE_KEY)
    games = [
        _game("espn-1", "Atlanta Hawks", "Washington Wizards", "2026-11-07T00:00:00Z"),
        _game("espn-2", "Indiana Pacers", "Denver Nuggets", "2026-11-07T22:00:00Z"),
    ]
    payload = {
        "matches": [
            {
                "match_id": "predix-1",
                "home": "Atlanta Hawks",
                "away": "Washington Wizards",
                "tipoff_utc": "2026-11-07T00:00:00Z",
                "probabilities": {"home": 0.70, "away": 0.30},
            },
            {
                "match_id": "predix-2",
                "home": "Indiana Pacers",
                "away": "Denver Nuggets",
                "tipoff_utc": "2026-11-07T22:00:00Z",
                "probabilities": {"home": 0.20, "away": 0.80},
            },
        ]
    }
    calls: list[str] = []
    pipeline = Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd="20261107"),
        InMemoryStore(),
        fetch_factory=_factory(payload, calls),
    )

    candidates = pipeline._predixsport_candidates_for_games(games)

    assert calls == ["https://api.predixsport.com/v1/predictions/nba/upcoming?days=7"]
    assert [prediction.game_id for prediction in candidates["espn-1"]] == ["espn-1"]
    assert [prediction.game_id for prediction in candidates["espn-2"]] == ["espn-2"]
    assert candidates["espn-1"][0].auxiliary_metrics["source_match_id"] == "predix-1"
    assert pipeline.prediction_provider_status["predixsport"] == {
        "enabled": True,
        "status": "fetched",
        "records_returned": 2,
        "records_matched": 2,
        "records_unmatched": 0,
        "records_rejected": 0,
    }


def test_predixsport_never_joins_a_team_only_or_wrong_time_prediction():
    game = _game("espn-1", "Atlanta Hawks", "Washington Wizards", "2026-11-07T00:00:00Z")
    prediction = Prediction(
        game_id="predix-1",
        provider="predixsport",
        observed_utc="2026-11-06T17:00:00Z",
        prob_home=0.70,
        prob_away=0.30,
        auxiliary_metrics={
            "source_match_id": "predix-1",
            "source_home_name": "Atlanta Hawks",
            "source_away_name": "Washington Wizards",
            # Beyond the strict cross-provider 30-minute tolerance.
            "source_tipoff_utc": "2026-11-07T01:00:01Z",
        },
    )

    matched, reason = _match_predixsport_prediction(prediction, [game])

    assert matched is None
    assert reason == "Predixsport prediction 'predix-1' has no exact home/away/UTC-tip match"


def test_nba_prediction_acquisition_admits_predixsport_candidate(monkeypatch):
    game = _game("espn-1", "Atlanta Hawks", "Washington Wizards", "2026-11-07T00:00:00Z")
    predix = Prediction(
        game_id=game.game_id,
        provider="predixsport",
        observed_utc="2026-11-06T17:00:00Z",
        prob_home=0.70,
        prob_away=0.30,
    )
    pipeline = Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd="20261107",
            fixture_now_utc="2026-11-06T18:00:00Z",
        ),
        InMemoryStore(),
        # Keep ESPN/Hardwood unavailable: this focused test asserts that the
        # NBA branch no longer skips the independently supplied candidate.
        fetch=lambda _url: {},
    )
    seen: list[str] = []

    def candidates(games):
        seen.extend(candidate.game_id for candidate in games)
        return {game.game_id: [predix]}

    monkeypatch.setattr(pipeline, "_predixsport_candidates_for_games", candidates)
    selection = pipeline._acquire_predictions([game])[game.game_id]

    assert seen == [game.game_id]
    assert selection.selected is not None
    assert selection.selected.provider == "predixsport"
    assert pipeline.store.get_predictions(game.game_id)[0].provider == "predixsport"
