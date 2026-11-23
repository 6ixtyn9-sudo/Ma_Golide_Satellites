"""Basketball Champions League lane on the FIBA GDAP platform.

The body here is one **verbatim row** from a captured response
(``getgdapgamesbycompetitionid?gdapCompetitionId=209123``, HTTP 200, 23,789 B,
2026-11-23T12:11:54Z, out-of-sandbox probe — this sandbox has no egress to
fiba.basketball). The other states are derived by mutating that row inside each
test, which is why the mutations are explicit rather than hidden in a fixture.

Cases covered: registry wiring, the shared-contract reuse, a settled row mapping
to totals-only, the UTC tip, a scheduled row, an unassigned bracket slot, an
unclassifiable status, a foreign competition id, a final without scores, the
season-index probe, the date filter, and the gate.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List

import pytest

from ingestion.league_registry import league_spec
from ingestion.models import GameStatus, League, PeriodKind
from ingestion.providers import bcl
from ingestion.providers import IdentityError
from ingestion.providers.fiba_euroleague_w_contract import (
    PUBLIC_FALLBACK_SUBSCRIPTION_KEY,
)

ROW = "tests/fixtures/bcl/captures/gdap-games-209123-row.json"


def _row() -> Dict[str, Any]:
    with open(ROW, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _fetch(payload: Any | None = None) -> Callable[[str], Any]:
    body = [_row()] if payload is None else payload

    def fetch(url: str) -> Any:
        if "getgdapgamesbycompetitionid" not in url:
            raise AssertionError(f"unexpected URL: {url}")
        return body

    return fetch


def test_registry_points_the_lane_at_gdap():
    spec = league_spec(League.BASKETBALL_CHAMPIONS_LEAGUE)
    assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
        "bcl",
        "bcl",
        "bcl",
    )
    assert "209123" in spec.source_detail


def test_competition_id_differs_from_the_euroleague_women_lane():
    from ingestion.providers.fiba_euroleague_w_contract import COMPETITION_ID as ELW_ID

    assert bcl.COMPETITION_ID != ELW_ID


def test_request_carries_the_site_shipped_apim_key():
    headers = bcl.request_headers()
    assert headers["Ocp-Apim-Subscription-Key"] == PUBLIC_FALLBACK_SUBSCRIPTION_KEY


def test_settled_row_maps_to_totals_only_with_a_utc_tip():
    game = bcl.game_from_row(_row())

    assert game.game_id == "BCL-136223"
    assert game.league is League.BASKETBALL_CHAMPIONS_LEAGUE
    assert game.home_name == "BC Rytas"
    assert game.away_name == "Sabah BC"
    # gameDateTimeUTC 2026-10-06T16:30:00 (offset-less UTC field) = 18:30 in Vilnius.
    assert game.scheduled_utc == "2026-10-06T16:30:00Z"
    assert game.status is GameStatus.FINAL
    assert game.season == "2027"
    assert game.phase == "Regular Season"
    assert game.actual.period_kind is PeriodKind.UNKNOWN
    assert (game.actual.home_final, game.actual.away_final) == (82.0, 73.0)
    assert game.actual.home_quarters == []


def test_team_identity_is_namespaced_away_from_the_womens_competition():
    game = bcl.game_from_row(_row())

    assert game.home_team_id.startswith("bcl:")
    assert game.home_team_id != f"fiba-elw:{bcl.COMPETITION_ID}:284917"


def test_unplayed_row_stays_scheduled_and_carries_no_result():
    row = _row()
    row["teamAScore"] = None
    row["teamBScore"] = None
    row["gameResultStatusCode"] = None
    row["liveGameStatus"] = 0
    row["statusCode"] = "VALID"
    row["isLive"] = False

    game = bcl.game_from_row(row)

    assert game.status is GameStatus.SCHEDULED
    assert game.actual is None


def test_unassigned_bracket_slot_is_skipped_and_recorded():
    row = _row()
    row["teamA"] = None

    games = bcl.discover_games(None, _fetch([row]))

    assert games == []
    assert any("unassigned bracket slot" in gap["detail"] for gap in bcl.last_gaps())


def test_unclassifiable_status_is_never_guessed():
    row = _row()
    row["statusCode"] = "WEIRD"
    row["liveGameStatus"] = 0
    row["gameResultStatusCode"] = None
    row["isLive"] = False

    games = bcl.discover_games(None, _fetch([row]))

    assert games == []
    assert any("refusing to guess" in gap["detail"] for gap in bcl.last_gaps())


def test_row_from_another_competition_is_refused():
    row = _row()
    row["competition"]["competitionId"] = 209127

    with pytest.raises(IdentityError):
        bcl.game_from_row(row)


def test_final_without_scores_is_refused():
    row = _row()
    row["teamAScore"] = None

    with pytest.raises(ValueError):
        bcl.game_from_row(row)


def test_date_filter_keeps_only_the_target_utc_day():
    assert [g.game_id for g in bcl.discover_games("2026-10-06", _fetch())] == [
        "BCL-136223"
    ]
    assert bcl.discover_games("2026-10-07", _fetch()) == []


def test_probe_records_url_size_and_digest():
    bcl.discover_games(None, _fetch())

    probe = bcl.last_probe()
    assert probe["url"].endswith("gdapCompetitionId=209123")
    assert probe["rows"] == 1
    assert probe["games"] == 1
    assert probe["bytes"] > 0
    assert len(probe["sha256"]) == 64


def test_envelope_shape_is_accepted_and_junk_is_refused():
    assert bcl.discover_games(None, _fetch({"value": [_row()]}))
    with pytest.raises(ValueError):
        bcl.discover_games(None, _fetch({"data": []}))


def test_gate_reads_the_index_and_filters_by_utc_date(monkeypatch):
    from ingestion import adaptive_gate

    seen: List[str] = []

    def fake_read(url: str, headers=None, **kwargs):  # noqa: ANN001, ANN003
        seen.append(url)
        return [_row()]

    monkeypatch.setattr(adaptive_gate, "_read_json_response", fake_read)
    config = adaptive_gate.LeagueGateConfig(
        label="Basketball Champions League",
        league="basketball_champions_league",
        league_key="BASKETBALL_CHAMPIONS_LEAGUE",
        state_key="basketball_champions_league",
        acquisition_timezone="Europe/Paris",
        schedule_source="bcl",
    )
    slate = adaptive_gate._bcl_schedule(config, "2026-10-06")

    assert len(seen) == 1
    assert slate and slate[0].game_id == "BCL-136223"
    assert slate[0].status == "final"
    assert slate[0].scheduled_utc == datetime(2026, 10, 6, 16, 30, tzinfo=timezone.utc)
    assert adaptive_gate._bcl_schedule(config, "2026-10-08") == []
