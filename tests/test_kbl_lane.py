"""Korea KBL lane: documented date-window route, fail-closed identity.

The fixtures in ``tests/fixtures/kbl`` are **shape fixtures**, not captures:
they are written to the column contract documented by the SportsDataverse
``kbl`` module (MIT, read 2026-11-23), so they prove the parser handles the
documented shape and nothing more. Per the 2026-11-23 source policy a lane is
promoted on a live receipt (URL, UTC, bytes, sha256 - see
``kbl.last_probe``), not on a shape fixture.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest

from ingestion import adaptive_gate
from ingestion.adaptive_gate import GateUncertain, fetch_schedule
from ingestion.league_registry import league_spec
from ingestion.models import GameStatus, League
from ingestion.providers import kbl

FIXTURES = Path(__file__).parent / "fixtures" / "kbl"


def _load(name: str):
    return json.loads((FIXTURES / name).read_text())


def _transport(rows=None, teams=None, seen=None):
    rows = _load("match_list_shape.json") if rows is None else rows
    teams = _load("team_list_shape.json") if teams is None else teams

    def fetch(url: str):
        if seen is not None:
            seen.append(url)
        if "/common/teamList" in url:
            return teams
        if "/match/list" in url:
            return rows
        raise AssertionError(f"unexpected KBL url {url}")

    return fetch


def test_registry_wires_the_korean_lane_to_its_own_adapter():
    spec = league_spec(League.SOUTH_KOREA_KBL)
    assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
        "kbl",
        "kbl",
        "kbl",
    )


def test_window_row_becomes_a_canonical_game_in_utc():
    seen: list[str] = []
    games = kbl.discover_games("2026-11-23", _transport(seen=seen), end_year=2026)

    assert [game.game_id for game in games] == ["KBL-S47G01N49", "KBL-S47G01N50", "KBL-S47G01N51"]
    first = games[0]
    # 19:00 KST is 10:00Z; KST carries no daylight saving.
    assert first.scheduled_utc == "2026-11-23T10:00:00Z"
    assert first.league is League.SOUTH_KOREA_KBL
    assert first.season == "2026"
    assert first.home_team_id == "16" and first.away_team_id == "35"
    # English names come from /common/teamList, not the Korean short name.
    assert "WONJU" in first.home_name.upper()
    assert "SAMSUNG" in first.away_name.upper()
    assert first.status is GameStatus.FINAL
    assert first.actual is not None
    assert (first.actual.home_final, first.actual.away_final) == (88.0, 81.0)
    assert first.actual.period_kind.value == "unknown"  # totals only, never quarters
    assert first.provider_ids["kbl"] == "S47G01N49"
    assert any("/match/list?season=2026" in url for url in seen)


def test_unstarted_row_is_scheduled_and_live_row_is_live():
    games = kbl.discover_games("2026-11-23", _transport(), end_year=2026)
    by_id = {game.game_id: game for game in games}
    assert by_id["KBL-S47G01N50"].status is GameStatus.SCHEDULED
    assert by_id["KBL-S47G01N50"].actual is None
    assert by_id["KBL-S47G01N51"].status is GameStatus.LIVE
    assert by_id["KBL-S47G01N51"].actual is None


def test_ended_row_without_scores_is_refused_rather_than_settled():
    rows = [dict(_load("match_list_shape.json")["data"][0])]
    rows[0].update({"is_ended": 1, "home_score": None, "away_score": None})

    games = kbl.discover_games("2026-11-23", _transport(rows=rows), end_year=2026)

    assert [game.status for game in games] == [GameStatus.SCHEDULED]
    assert games[0].actual is None
    assert "is_ended=1 without both scores" in kbl.last_gaps()[0]["detail"]


def test_an_unrecognisable_envelope_raises_instead_of_reading_as_an_empty_slate():
    with pytest.raises(ValueError, match="no recognisable row list"):
        kbl.fetch_window(2026, "20261122", "20261124", _transport(rows={"unexpected": []}))


def test_probe_records_url_bytes_digest_and_row_keys():
    kbl.discover_games("2026-11-23", _transport(), end_year=2026)
    probe = kbl.last_probe()

    assert probe["rows"] == 3
    # The documented envelope is not asserted, so the probe reports which
    # shape was actually served: this fixture carries a {"data": [...]} envelope.
    assert probe["shape"] == "object"
    assert "/match/list?season=2026" in probe["url"]
    assert probe["bytes"] > 0 and len(probe["sha256"]) == 64
    assert "gmkey" in probe["row_keys"]


def test_gate_reads_the_same_date_window_and_never_a_season(monkeypatch):
    monkeypatch.setattr(
        adaptive_gate,
        "_read_json_response",
        lambda url, headers=None: _load("match_list_shape.json")["data"],
    )
    config = adaptive_gate.LeagueGateConfig(
        label="South Korea KBL",
        league="SOUTH_KOREA_KBL",
        league_key="SOUTH_KOREA_KBL",
        state_key="south_korea_kbl",
        acquisition_timezone="Asia/Seoul",
        schedule_source="kbl",
    )

    games = fetch_schedule(config, "2026-11-23")

    assert len(games) == 3
    assert games[0].scheduled_utc == datetime(2026, 11, 23, 10, 0, tzinfo=timezone.utc)
    assert games[0].game_id == "KBL-S47G01N49"


def test_gate_fails_open_when_the_body_has_no_row_list(monkeypatch):
    monkeypatch.setattr(
        adaptive_gate, "_read_json_response", lambda url, headers=None: {"unexpected": 1}
    )
    config = adaptive_gate.LeagueGateConfig(
        label="South Korea KBL",
        league="SOUTH_KOREA_KBL",
        league_key="SOUTH_KOREA_KBL",
        state_key="south_korea_kbl",
        acquisition_timezone="Asia/Seoul",
        schedule_source="kbl",
    )

    with pytest.raises(GateUncertain, match="no recognisable row list"):
        fetch_schedule(config, "2026-11-23")
