import json
from pathlib import Path

from ingestion.models import League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import InMemoryStore

ROOT = Path(__file__).parent / "fixtures"


def _fixture(path):
    return json.loads(path.read_text(encoding="utf-8"))


def test_nba_standings_smoke_reports_hash_and_does_not_persist_or_change_model_records():
    standings = _fixture(ROOT / "espn" / "standings.json")
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd="20261120",
            fixture_now_utc="2026-11-19T12:00:00Z",
        ),
        store,
        fetch=lambda _url: standings,
    )

    report = pipeline._standings_smoke()

    assert report["status"] == "ok"
    assert report["provider"] == "espn"
    assert report["teams_seen"] > 0
    assert report["payload_hash"]
    assert report["observed_utc"] == "2026-11-19T12:00:00Z"
    assert report["role"] == "evidence_coverage_only_not_a_model_input"
    assert store.list_games(league=League.NBA) == []


def test_euroleague_standings_smoke_uses_prefixed_season_and_reports_coverage():
    schedule = _fixture(ROOT / "euroleague" / "schedule.json")
    standings = _fixture(ROOT / "euroleague" / "standings_v3.json")
    seen = []

    def fetch(url):
        seen.append(url)
        if "/games" in url:
            assert "/seasons/E2026/" in url
            return schedule
        assert "/rounds/10/basicstandings" in url
        return standings

    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE",
            date_yyyymmdd="20261120",
            fixture_now_utc="2026-11-19T12:00:00Z",
        ),
        InMemoryStore(),
        fetch=fetch,
    )

    report = pipeline._standings_smoke()

    assert report["status"] == "ok"
    assert report["provider"] == "euroleague"
    assert report["teams_seen"] == len(standings["teams"])
    assert report["payload_hash"]
    assert len(seen) == 2
