"""Turkey TBF (BSL) lane: checked against a real federation response body.

Unlike the KBL lane, this one has an actual capture: the 2026-10-11 daily body
the federation served, vendored at
``tests/fixtures/tbf/captures/get-daily-matches-2026-10-11.json`` and attributed
in ``tests/fixtures/CAPTURE-PROVENANCE.md``. That body carries three
competitions in one envelope (kbsl, tbl, tkbl), which is the whole reason one
adapter can serve BSL, TBL, KBSL and TKBL.

Cases covered: the competition filter, Turkey-local -> UTC conversion,
played/not-played mapping, a played row with no scores being refused as a final,
totals-only settlement, a failure envelope, the probe receipt, and the gate's
date window.
"""

from __future__ import annotations

import json
from typing import Any, Callable, Dict, List

import pytest

from ingestion.adaptive_gate import LeagueGateConfig
from ingestion.league_registry import league_spec
from ingestion.models import GameStatus, League, PeriodKind
from ingestion.providers import tbf

CAPTURE = "tests/fixtures/tbf/captures/get-daily-matches-2026-10-11.json"


def _capture() -> Dict[str, Any]:
    with open(CAPTURE, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _rows() -> List[Dict[str, Any]]:
    return [
        match
        for group in _capture()["data"]
        for match in group.get("matches") or []
        if isinstance(match, dict)
    ]


def _fetch(payload: Any | None = None, calls: List[str] | None = None) -> Callable[[str], Any]:
    body = _capture() if payload is None else payload

    def fetch(url: str) -> Any:
        if calls is not None:
            calls.append(url)
        if "get-daily-matches" not in url:
            raise AssertionError(f"unexpected URL: {url}")
        return body

    return fetch


def test_registry_points_the_lane_at_the_federation_route():
    spec = league_spec(League.TURKEY_BSL)
    assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
        "tbf",
        "tbf",
        "tbf",
    )
    assert "get-daily-matches" in spec.source_detail


def test_capture_carries_three_competitions_in_one_envelope():
    codes = {str(row.get("activityCode")).lower() for row in _rows()}
    assert {"kbsl", "tbl", "tkbl"} <= codes


def test_turkey_local_tip_is_converted_to_utc():
    games = tbf.discover_games("2026-10-11", _fetch(), competition="kbsl")
    assert games
    # The capture's first row tips at 14:00 Turkey time = 11:00Z.
    assert any(game.scheduled_utc.endswith("T11:00:00Z") for game in games)


def test_competition_filter_selects_one_league_from_the_shared_body():
    kbsl = tbf.discover_games("2026-10-11", _fetch(), competition="kbsl")
    tbl = tbf.discover_games("2026-10-11", _fetch(), competition="tbl")
    everything = tbf.discover_games("2026-10-11", _fetch(), competition=None)

    assert kbsl and tbl
    assert {g.evidence["activity_code"] for g in kbsl} == {"kbsl"}
    assert {g.evidence["activity_code"] for g in tbl} == {"tbl"}
    assert len(everything) >= len(kbsl) + len(tbl)


def test_played_row_settles_totals_only_and_never_invents_quarters():
    games = tbf.discover_games("2026-10-11", _fetch(), competition="kbsl")
    settled = [game for game in games if game.status is GameStatus.FINAL]
    assert settled
    for game in settled:
        assert game.actual is not None
        assert game.actual.period_kind is PeriodKind.UNKNOWN
        assert game.actual.home_quarters == []
        assert game.actual.away_quarters == []
        assert game.actual.home_final is not None
        assert game.actual.away_final is not None


def test_played_row_without_scores_is_refused_as_a_final():
    body = _capture()
    first = body["data"][0]["matches"][0]
    first["homeTeam"]["score"] = ""
    first["awayTeam"]["score"] = None

    games = tbf.discover_games("2026-10-11", _fetch(body), competition="kbsl")

    match_id = str(int(first["matchId"]))
    refused = [game for game in games if game.provider_ids["tbf"] == match_id]
    assert refused and refused[0].status is not GameStatus.FINAL
    assert refused[0].actual is None
    assert any("refused as a final" in gap["detail"] for gap in tbf.last_gaps())


def test_unplayed_row_stays_scheduled():
    body = _capture()
    first = body["data"][0]["matches"][0]
    first["matchStatusId"] = 0.0
    first["homeTeam"]["score"] = ""
    first["awayTeam"]["score"] = ""

    games = tbf.discover_games("2026-10-11", _fetch(body), competition="kbsl")

    match_id = str(int(first["matchId"]))
    row = [game for game in games if game.provider_ids["tbf"] == match_id][0]
    assert row.status is GameStatus.SCHEDULED
    assert row.actual is None


def test_failure_envelope_raises_instead_of_reading_as_an_empty_slate():
    with pytest.raises(ValueError):
        tbf.discover_games(
            "2026-10-11",
            _fetch({"data": [], "isSuccess": False, "message": "problemDetails"}),
            competition="bsl",
        )


def test_body_without_a_data_list_raises():
    with pytest.raises(ValueError):
        tbf.discover_games("2026-10-11", _fetch({"isSuccess": True}), competition="bsl")


def test_probe_records_what_was_read_not_what_was_expected():
    calls: List[str] = []
    tbf.discover_games("2026-10-11", _fetch(calls=calls), competition="tbl")

    probe = tbf.last_probe()
    assert probe["url"].startswith(tbf.DAILY_MATCHES_URL + "?match_date=2026-10-")
    assert probe["bytes"] > 0
    assert len(probe["sha256"]) == 64
    assert probe["is_success"] is True
    assert probe["rows"] >= 1
    assert {"kbsl", "tbl", "tkbl"} <= set(probe["competitions"])
    # The UTC target day is read through the two Turkey local dates it spans.
    assert len(calls) == 2


def test_gate_reads_the_same_route_for_its_target_date(monkeypatch):
    from ingestion import adaptive_gate

    seen: List[str] = []

    def fake_read(url: str, headers=None, **kwargs):  # noqa: ANN001, ANN003
        seen.append(url)
        return _capture()

    monkeypatch.setattr(adaptive_gate, "_read_json_response", fake_read)
    config = LeagueGateConfig(
        label="Turkey BSL",
        league="turkey_bsl",
        league_key="TURKEY_BSL",
        state_key="turkey_bsl",
        acquisition_timezone="Europe/Istanbul",
        schedule_source="tbf",
    )
    slate = adaptive_gate._tbf_schedule(config, "2026-10-11")

    assert seen and all("get-daily-matches" in url for url in seen)
    assert all(game.game_id.startswith("TBF-") for game in slate)
    # The capture holds no BSL row, so the gate must report an empty BSL slate
    # rather than borrowing another competition's games.
    assert slate == []
