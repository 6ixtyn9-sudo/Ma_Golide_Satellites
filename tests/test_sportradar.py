"""Hermetic Sportradar independent CBA probability contract tests."""

from __future__ import annotations

from copy import deepcopy
from urllib.parse import parse_qs, urlparse

import pytest

from ingestion.models import Game, League
from ingestion.providers import IdentityError, sportradar

OBSERVED = "2026-11-07T08:00:00Z"
TIP = "2026-11-07T11:35:00Z"


def _game(game_id="101"):
    return Game(
        game_id=game_id,
        league=League.CBA,
        season="26-27",
        home_team_id="10",
        away_team_id="20",
        home_name="Beijing Ducks",
        away_name="Shanghai Sharks",
        scheduled_utc=TIP,
        provider_ids={"isportsapi": game_id},
    )


def _probability_row(event_id="sr:sport_event:1"):
    return {
        "sport_event": {
            "id": event_id,
            "start_time": TIP,
            "start_time_confirmed": True,
            "competitors": [
                {"qualifier": "home", "name": "Beijing Ducks"},
                {"qualifier": "away", "name": "Shanghai Sharks"},
            ],
        },
        "markets": [{
            "name": "2way",
            "live": False,
            "removed": False,
            "last_updated": "2026-11-07T07:55:00+00:00",
            "outcomes": [
                {"name": "home_team_winner", "probability": 61.5},
                {"name": "away_team_winner", "probability": 38.5},
            ],
        }],
    }


def _payloads():
    return {
        "competitions": {
            "generated_at": "2026-11-07T07:00:00Z",
            "competitions": [
                {"id": "sr:competition:999", "name": "CBA", "gender": "men",
                 "category": {"country_code": "AUS", "name": "Australia"}},
                {"id": "sr:competition:321", "name": "CBA", "gender": "men",
                 "category": {"country_code": "CHN", "name": "China"}},
            ],
        },
        "seasons": {
            "seasons": [{
                "id": "sr:season:26",
                "competition_id": "sr:competition:321",
                "name": "CBA 26/27",
                "start_date": "2026-10-01",
                "end_date": "2027-05-01",
                "disabled": False,
            }],
        },
        "probabilities": {
            "generated_at": "2026-11-07T07:56:00Z",
            "sport_event_probabilities": [_probability_row()],
        },
    }


def _fetch(payloads, calls=None):
    def fetch(url, headers=None):
        if calls is not None:
            calls.append((url, headers))
        if url.endswith("/competitions.json"):
            return deepcopy(payloads["competitions"])
        if "/competitions/sr:competition:321/seasons.json" in url:
            return deepcopy(payloads["seasons"])
        if "/seasons/sr:season:26/probabilities.json" in url:
            return deepcopy(payloads["probabilities"])
        raise AssertionError(f"unexpected URL {url}")
    return fetch


def test_fetches_independent_probability_without_inventing_scores(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "secret-sportradar")
    monkeypatch.setenv("SPORTRADAR_ACCESS_LEVEL", "trial")
    calls = []
    predictions = sportradar.fetch_cba_probabilities(
        "20261107", [_game()], _fetch(_payloads(), calls), now_utc=OBSERVED
    )
    prediction = predictions["101"]
    assert prediction.provider == "sportradar"
    assert prediction.prob_home == pytest.approx(0.615)
    assert prediction.prob_away == pytest.approx(0.385)
    assert prediction.expected_home_score is None
    assert prediction.expected_total is None
    assert prediction.source_revision == "sr:sport_event:1"
    assert prediction.source_modified_utc == "2026-11-07T07:55:00Z"
    assert prediction.observed_utc == OBSERVED
    assert all("secret-sportradar" not in url for url, _ in calls)
    assert all(headers["x-api-key"] == "secret-sportradar" for _, headers in calls)


def test_wrong_country_or_ambiguous_competition_fails_closed(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "key")
    payloads = _payloads()
    payloads["competitions"]["competitions"].append(
        {"id": "sr:competition:322", "name": "CBA", "gender": "men",
         "category": {"country_code": "CHN", "name": "China"}}
    )
    with pytest.raises(IdentityError, match="must be unique"):
        sportradar.fetch_cba_probabilities(
            "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
        )


def test_cross_provider_join_requires_confirmed_exact_fixture(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "key")
    for mutation in (
        lambda row: row["sport_event"].update(start_time="2026-11-07T11:36:00Z"),
        lambda row: row["sport_event"].update(start_time_confirmed=False),
        lambda row: row["sport_event"]["competitors"][0].update(name="Other Team"),
    ):
        payloads = _payloads()
        mutation(payloads["probabilities"]["sport_event_probabilities"][0])
        assert sportradar.fetch_cba_probabilities(
            "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
        ) == {}


def test_duplicate_and_live_probabilities_remain_gaps(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "key")
    payloads = _payloads()
    row = payloads["probabilities"]["sport_event_probabilities"][0]
    payloads["probabilities"]["sport_event_probabilities"].append(deepcopy(row))
    assert sportradar.fetch_cba_probabilities(
        "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
    ) == {}

    payloads = _payloads()
    payloads["probabilities"]["sport_event_probabilities"][0]["markets"][0]["live"] = True
    assert sportradar.fetch_cba_probabilities(
        "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
    ) == {}


def test_generated_at_is_fallback_provenance_and_percentages_are_validated(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "key")
    payloads = _payloads()
    market = payloads["probabilities"]["sport_event_probabilities"][0]["markets"][0]
    del market["last_updated"]
    prediction = sportradar.fetch_cba_probabilities(
        "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
    )["101"]
    assert prediction.source_modified_utc == "2026-11-07T07:56:00Z"

    market["outcomes"][0]["probability"] = 101
    assert sportradar.fetch_cba_probabilities(
        "2026-11-07", [_game()], _fetch(payloads), now_utc=OBSERVED
    ) == {}


def test_season_probability_pagination_is_bounded_and_complete(monkeypatch):
    monkeypatch.setenv("SPORTRADAR_API_KEY", "key")
    monkeypatch.setattr(sportradar, "_PAGE_SIZE", 2)
    starts = []

    def fetch(url, headers=None):
        query = parse_qs(urlparse(url).query)
        start = int(query["start"][0])
        starts.append(start)
        rows = [_probability_row(f"sr:sport_event:{index}") for index in range(5)]
        return {
            "generated_at": "2026-11-07T07:56:00Z",
            "sport_event_probabilities": rows[start:start + 2],
        }

    rows = list(sportradar._probability_rows(fetch, "sr:season:26"))
    assert len(rows) == 5
    assert starts == [0, 2, 4]


def test_smoke_does_not_fetch_without_credentials(monkeypatch):
    monkeypatch.delenv("SPORTRADAR_API_KEY", raising=False)
    result = sportradar.smoke("20261107", lambda *_args, **_kwargs: pytest.fail("network"))
    assert result["status"] == "UNAUTHORIZED"
