"""EuroLeague / EuroCup must hold a read receipt, like every other lane.

Live evidence (lane-job bundle, run 45250804751): the EuroLeague and EuroCup
jobs were the only lanes still printing the void guard

    adapter returned no fixtures and recorded no coverage gap; treat the
    slate as unknown, not empty

on a night with no on-date games.  The guard is right to exist -- a lane that
read nothing must not be rendered as an empty slate -- but the euroleague
adapter read a full season schedule (6 rows in the committed fixture) and the
pipeline threw that fact away.  `_fetch_season_rows` now records what it read
so the guard speaks only for lanes that genuinely read nothing.
"""

import hashlib
import json
from pathlib import Path

import pytest

from ingestion.models import CoverageState
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import euroleague
from ingestion.providers import euroleague as euro
from ingestion.store import InMemoryStore

FIXTURES = Path(__file__).parent / "fixtures" / "euroleague"


def _payload() -> dict:
    return json.loads((FIXTURES / "schedule.json").read_text())


def _fetch(url: str) -> dict:
    assert "/seasons/" in url
    return _payload()


def _off_date_pipeline(league: str = "EUROLEAGUE") -> Pipeline:
    # 2026-11-23 is not in the fixture's slate (2026-10-22 … 2027-04-16).
    config = PipelineConfig(league=league, date_yyyymmdd="20261123")
    return Pipeline(config, InMemoryStore(), fetch=_fetch, allow_writes=False)


def _void_gap(pipeline: Pipeline):
    return next(
        (g for g in pipeline.gaps if "unknown, not empty" in str(g.detail)),
        None,
    )


def test_off_date_slate_records_the_read_instead_of_a_void():
    pipeline = _off_date_pipeline()
    games = pipeline._discover_games()

    assert games == []
    assert _void_gap(pipeline) is None, "a lane that read 6 rows is not a void"
    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    assert row["coverage"]["games"] == 0
    assert row["coverage"]["rows"] == len(_payload()["data"])
    assert row["coverage"]["seasoncode"] == "E2026"
    assert row["state"] == CoverageState.MISSING.value
    assert len(row["coverage"]) == 3


def test_read_receipt_carries_payload_digest_and_does_not_call_it_a_capture():
    pipeline = _off_date_pipeline()
    pipeline._discover_games()

    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    digest = row["detail"].split("payload_sha256=")[-1]
    assert len(digest) == 64
    assert hashlib.sha256(bytes.fromhex(digest)).hexdigest()  # hex, decodable
    # Named `payload_*`, never `bytes`/`sha256`: this is the decoded payload,
    # not a wire capture, and the label must not imply otherwise.
    tokens = {part.split("=")[0] for part in row["detail"].split()}
    assert {"payload_bytes", "payload_sha256"} <= tokens
    assert not tokens & {"bytes", "sha256"}


def test_on_date_slate_is_ok_and_still_quiet():
    config = PipelineConfig(league="EUROLEAGUE", date_yyyymmdd="20261106")
    pipeline = Pipeline(config, InMemoryStore(), fetch=_fetch, allow_writes=False)
    games = pipeline._discover_games()

    assert games and all(g.game_id.startswith("E2026-") for g in games)
    assert _void_gap(pipeline) is None
    row = next(h for h in pipeline._health_written if h["provider"] == "euroleague")
    assert row["state"] == CoverageState.OK.value
    assert row["last_success_utc"]


def test_verified_empty_season_payload_is_still_refused_but_recorded():
    """Fail closed on a bad season code -- and still show the read happened."""

    euro.reset_probe()
    with pytest.raises(ValueError, match="empty"):
        euro.discover_games(None, lambda url: {"data": [], "total": 0}, "E2026")
    probe = euro.last_probe()
    assert probe["rows"] == 0
    assert probe["seasoncode"] == "E2026"


def test_probe_is_empty_when_nothing_was_read():
    euro.reset_probe()

    def boom(url):
        raise RuntimeError("socket closed")

    with pytest.raises(RuntimeError):
        euro.discover_games(None, boom, "E2026")
    assert euro.last_probe() == {}


def test_guard_still_fires_when_the_lane_reads_nothing(monkeypatch):
    """The guard keeps its teeth: silence is still unknown, not empty."""

    monkeypatch.setattr(euroleague, "discover_games", lambda *a, **k: [])
    monkeypatch.setattr(euroleague, "last_probe", lambda: {})
    monkeypatch.setattr(euroleague, "reset_probe", lambda: None)

    pipeline = _off_date_pipeline()
    pipeline._discover_games()

    gap = _void_gap(pipeline)
    assert gap is not None
    assert gap.reason.value == "unsupported"
    assert gap.required_field == "official_schedule"


def test_eurocup_shares_the_receipt():
    euro.reset_probe()
    rows = json.loads((FIXTURES / "schedule_scores.json").read_text())
    euro.discover_games(None, lambda url: rows, "U2026")
    probe = euro.last_probe()
    assert probe["seasoncode"] == "U2026"
    assert probe["rows"] > 0
    assert "/competitions/U/" in probe["url"]
