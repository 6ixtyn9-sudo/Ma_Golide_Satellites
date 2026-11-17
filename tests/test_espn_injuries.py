import pytest

from ingestion.providers.espn import INJURIES_URL, fetch_nba_injuries


def test_fetch_nba_injuries_normalizes_ids_and_preserves_time_provenance():
    payload = {
        "timestamp": "2026-11-17T12:00:00Z",
        "injuries": [
            {
                "team": {"id": "1", "displayName": "Boston Celtics"},
                "injuries": [
                    {
                        "athlete": {"id": "p1", "displayName": "Player One"},
                        "status": "Out",
                        "details": {"type": "Ankle", "detail": "sprain"},
                        "date": "2026-11-16T18:00:00Z",
                    }
                ],
            },
            {"team": {"id": "2"}, "injuries": []},
        ],
    }
    urls = []

    def fetch(url):
        urls.append(url)
        return payload

    result = fetch_nba_injuries(fetch, observed_utc="2026-11-17T12:01:00Z")

    assert urls == [INJURIES_URL]
    assert result["teams_seen"] == 2
    assert result["injury_entries_seen"] == 1
    assert result["source_modified_utc"] == "2026-11-17T12:00:00Z"
    assert result["teams"]["1"][0] == {
        "athlete_id": "p1",
        "name": "Player One",
        "status": "Out",
        "type": "Ankle",
        "detail": "sprain",
        "date": "2026-11-16T18:00:00Z",
    }
    assert result["teams"]["2"] == []


def test_fetch_nba_injuries_rejects_unusable_payloads():
    with pytest.raises(ValueError, match="no team injury list"):
        fetch_nba_injuries(lambda _url: {"timestamp": "2026-11-17T12:00:00Z"})

    with pytest.raises(ValueError, match="no identifiable teams"):
        fetch_nba_injuries(lambda _url: {"injuries": [{"injuries": []}]})
