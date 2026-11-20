"""Hermetic ABA League official-contract proposal tests.

These are synthetic contract examples written for Ma Golide's fail-closed
adapter. They are not retained ABA League page bodies and do not imply live
source-use permission or live verification.
"""

from __future__ import annotations

import pytest

from ingestion.duckdb_store import DuckDBStore
from ingestion.models import GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import IdentityError
from ingestion.providers import aba_liga as aba
from ingestion.store import InMemoryStore


CALENDAR = """
<html><body>
<select id="season"><option value="26/1/">Season 2026/27</option>
<option value="25/1/" selected>Season 2025/26</option></select>
<div class="panel"><h4 class="panel-title">Round 1</h4>
<table class="league_calendar_table"><thead><tr>
<th>Game</th><th>Result</th><th>Date and Time</th><th>Info</th>
</tr></thead><tbody>
<tr>
<td><p class="hidden-xs"><a href="/match/6/25/1/Overview/q1/1/home/buducnost-spartak/">Budućnost VOLI <span>:</span> Spartak Office Shoes</a></p></td>
<td>93 : 97</td><td>Sunday, 05.10.2025 19:30 CET</td><td>Overview</td>
</tr>
<tr>
<td><p class="hidden-xs"><a href="/match/7/25/1/Overview/q1/1/home/buducnost-fmp/">Budućnost VOLI <span>:</span> FMP</a></p></td>
<td></td><td>Friday, 10.10.2025 20:00 CET</td><td>Preview</td>
</tr>
</tbody></table></div>
</body></html>
"""

MATCH = """
<html><body>
<input id="match_data_id" value="6"><input id="match_data_sez" value="25"><input id="match_data_lea" value="1">
<div id="basic_match_info"><div class="dateAndVenue_container">Sunday, 05.10.2025 19:30 CET<br>Venue: Morača</div></div>
<div id="match_clubs_and_results_info_table"><table><tr>
<td class="hidden-xs"><span class="club"><a href="/team/12/25/1/0/buducnost/">Budućnost VOLI</a></span></td>
<td class="gameScore">93 : 97</td>
<td class="hidden-xs"><span class="club"><a href="/team/91/25/1/0/spartak/">Spartak Office Shoes</a></span></td>
</tr></table></div>
<table id="match_cetrtine_rezultat"><tbody>
<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td></tr>
<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>
</tbody></table>
</body></html>
"""


def _fetch(url: str) -> str:
    if "/calendar/25/1/" in url:
        return CALENDAR
    if "/match/6/25/1/Overview" in url:
        return MATCH
    raise AssertionError(f"unexpected ABA contract URL: {url}")


def test_official_calendar_contract_uses_belgrade_dst_not_fixed_cet():
    games = aba.parse_calendar(CALENDAR, 2026)

    assert [(game.game_no, game.has_final) for game in games] == [(6, True), (7, False)]
    # The site labels the wall time CET, but October 2025 is actually CEST in
    # Belgrade. A fixed CET parser would incorrectly produce 18:30Z.
    assert games[0].scheduled_utc == "2025-10-05T17:30:00Z"
    assert games[0].home_name == "Budućnost VOLI"
    assert games[0].away_name == "Spartak Office Shoes"
    assert aba.season_key(2026) == 25
    assert aba.calendar_url(2026).endswith("/calendar/25/1/")
    assert aba.match_url(6, 2026).endswith("/match/6/25/1/Overview")


def test_discovery_requires_detail_team_ids_and_marks_calendar_finality():
    games = aba.discover_games("20251005", text_fetch=_fetch)

    assert len(games) == 1
    game = games[0]
    assert game.game_id == "aba_liga:25:1:6"
    assert game.league is League.ABA_LEAGUE
    assert game.status is GameStatus.FINAL
    assert game.actual is None  # periods are only accepted through result parsing
    assert game.home_team_id == "aba_liga:team:12"
    assert game.away_team_id == "aba_liga:team:91"
    assert game.provider_ids == {"aba_liga": "25:1:6"}
    assert game.evidence["aba_liga"]["evidence"] == aba.EVIDENCE


def test_pipeline_uses_aba_adapter_for_discovery_and_results_without_fallback_gap():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="ABA_LEAGUE", date_yyyymmdd="20251005"),
        store,
        fetch=_fetch,
    )

    games = pipeline._discover_games()
    assert [game.game_id for game in games] == ["aba_liga:25:1:6"]
    assert store.get_game("aba_liga:25:1:6") is not None
    assert not any(gap.required_field == "official_schedule" for gap in pipeline.gaps)

    finals = pipeline._fetch_finals()
    assert [game.game_id for game in finals] == ["aba_liga:25:1:6"]
    assert finals[0].actual is not None


def test_result_contract_reconciles_all_quarters_before_canonical_final():
    finals = aba.fetch_recent_results("2025-10-05", text_fetch=_fetch)

    assert len(finals) == 1
    actual = finals[0].actual
    assert actual is not None
    assert actual.home_quarters == [21.0, 25.0, 20.0, 27.0]
    assert actual.away_quarters == [24.0, 24.0, 23.0, 26.0]
    assert actual.home_final == 93.0 and actual.away_final == 97.0
    assert actual.reconciles() == (True, "")


def test_result_refuses_scope_drift_and_does_not_emit_a_final():
    out_of_scope = MATCH.replace('id="match_data_sez" value="25"', 'id="match_data_sez" value="24"')

    def fetch(url: str) -> str:
        return CALENDAR if "/calendar/" in url else out_of_scope

    assert aba.fetch_recent_results("2025-10-05", text_fetch=fetch) == []
    gaps = aba.last_gaps()
    assert len(gaps) == 1
    assert gaps[0].required_field == "official_result"
    assert "does not match" in gaps[0].detail


def test_result_refuses_period_total_mismatch_and_timestamp_drift():
    bad_periods = MATCH.replace("27 : 26", "26 : 26")

    def bad_period_fetch(url: str) -> str:
        return CALENDAR if "/calendar/" in url else bad_periods

    assert aba.fetch_recent_results("2025-10-05", text_fetch=bad_period_fetch) == []
    assert "does not reconcile" in aba.last_gaps()[0].detail

    wrong_tip = MATCH.replace("19:30 CET", "20:30 CET")

    def wrong_tip_fetch(url: str) -> str:
        return CALENDAR if "/calendar/" in url else wrong_tip

    assert aba.discover_games("20251005", text_fetch=wrong_tip_fetch) == []
    assert "scheduled time contradicts calendar" in aba.last_gaps()[0].detail


def test_history_batch_stops_at_bad_game_and_does_not_claim_it_was_written():
    calendar = CALENDAR.replace(
        "</tbody></table></div>",
        """<tr>
<td><p class=\"hidden-xs\"><a href=\"/match/8/25/1/Overview/q1/1/home/fmp-buducnost/\">FMP <span>:</span> Budućnost VOLI</a></p></td>
<td>80 : 70</td><td>Sunday, 12.10.2025 18:00 CET</td><td>Overview</td>
</tr></tbody></table></div>""",
    )
    broken = MATCH.replace('value="6"', 'value="8"', 1).replace("93 : 97", "80 : 70", 1)

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        if "/match/6/" in url:
            return MATCH
        if "/match/8/" in url:
            return broken  # scope passes, header teams/time do not
        raise AssertionError(url)

    games, attempted, exhausted = aba.season_finals(2026, text_fetch=fetch)
    assert [game.game_id for game in games] == ["aba_liga:25:1:6"]
    assert attempted == [6, 8]
    assert exhausted is False
    assert aba.last_gaps()[0].required_field == "historical_official_result"


def test_pipeline_history_writes_reconciled_final_and_advances_game_cursor(tmp_path):
    store = DuckDBStore(str(tmp_path / "aba" / "warehouse.duckdb"), persist=True)
    try:
        pipeline = Pipeline(
            PipelineConfig(
                league="ABA_LEAGUE",
                date_yyyymmdd="20251006",
                history_days=0,
                history_request_budget=1,
                deep_history_event_budget=5,
                fixture_now_utc="2025-10-06T12:00:00Z",
            ),
            store,
            fetch=_fetch,
            allow_writes=True,
        )
        report = pipeline.backfill()
        deep = report["backfill"]["deep_history"]
        assert deep["status"] == "acquired"
        assert deep["events_acquired"] == 1
        assert deep["final_scores_written"] == 1
        assert deep["cursor"] == {
            "floor_season": 2002,
            "last_game_no": 0,
            "next_season": 2025,
        }
        stored = store.get_game("aba_liga:25:1:6")
        assert stored is not None and stored.actual is not None
        assert stored.actual.reconciles() == (True, "")
        health = store.get_provider_health("aba-liga-history", League.ABA_LEAGUE)
        assert health is not None and health.coverage == deep["cursor"]
    finally:
        store.close()


def test_adapter_is_enabled_and_smoke_replays_the_calendar_contract():
    assert aba.is_enabled() is True
    report = aba.smoke(text_fetch=_fetch, season=2026)
    assert report["calendar_games"] == 2
    assert report["calendar_finals"] == 1
    assert report["verification"] == aba.EVIDENCE


def test_parser_rejects_a_calendar_scope_or_malformed_score():
    with pytest.raises(IdentityError, match="calendar scope"):
        aba.parse_calendar(CALENDAR.replace('value="25/1/" selected', 'value="25/1/"'), 2026)
    with pytest.raises(IdentityError, match="calendar result"):
        aba.parse_calendar(CALENDAR.replace("93 : 97", "FINAL"), 2026)
