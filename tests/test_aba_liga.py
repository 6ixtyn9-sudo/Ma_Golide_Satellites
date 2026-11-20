"""Hermetic ABA League official-contract proposal tests.

These are synthetic contract examples written for Ma Golide's fail-closed
adapter. They are not retained ABA League page bodies and do not imply live
source-use permission or live verification.
"""

from __future__ import annotations

import pytest

from ingestion.duckdb_store import DuckDBStore
from ingestion.models import FinalStatus, GameStatus, League, ProviderHealth
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

UNPINNED_DATE_RANGE = "Tuesday, 01.02.2033 - Friday, 04.02.2033"
UNPINNED_DATE = "Saturday, 05.02.2033"

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

ADMINISTRATIVE_CALENDAR = CALENDAR.replace("93 : 97", "20 : 0")
EMPTY_PERIOD_CONTAINER = (
    '<table id="match_cetrtine_rezultat"><tbody><tr></tr><tr></tr></tbody></table>'
)
ADMINISTRATIVE_MATCH = MATCH.replace("93 : 97", "20 : 0").replace(
    '<table id="match_cetrtine_rezultat"><tbody>\n'
    '<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td></tr>\n'
    '<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>\n'
    "</tbody></table>",
    EMPTY_PERIOD_CONTAINER,
)
MATCH_CONTENT = MATCH.replace(
    "<html><body>", "<html><body><div id=\"match_content\">"
).replace("</body></html>", "</div></body></html>")
RETROACTIVELY_REGISTERED_MATCH = (
    MATCH_CONTENT.replace("93 : 97", "20 : 0")
    .replace(
        '<table id="match_cetrtine_rezultat"><tbody>\n'
        '<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td></tr>\n'
        '<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>\n'
        "</tbody></table>",
        EMPTY_PERIOD_CONTAINER,
    )
    .replace(
        "</div></body></html>",
        "<div><i>This match took place, but has been later registered with official result. "
        "Statistics of the players from this match have been annulled.</i></div>"
        "</div></body></html>",
    )
)
PLAYED_TWENTY_TO_ZERO_MATCH = (
    MATCH.replace("93 : 97", "20 : 0")
    .replace(
        "<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>",
        "<tr><td>5 : 0</td><td>5 : 0</td><td>5 : 0</td><td>5 : 0</td></tr>",
    )
)


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


def test_calendar_skips_unpinned_future_date_range_with_visible_coverage_gap():
    calendar = CALENDAR.replace("Friday, 10.10.2025 20:00 CET", UNPINNED_DATE_RANGE)

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        if "/match/6/" in url:
            return MATCH
        raise AssertionError(url)

    games, attempted, exhausted, calendar_final_count = aba.season_finals(2026, text_fetch=fetch)
    assert [game.game_id for game in games] == ["aba_liga:25:1:6"]
    assert attempted == [6]
    assert exhausted is True
    assert calendar_final_count == 1
    gaps = aba.last_gaps()
    assert len(gaps) == 1
    assert gaps[0].game_id == "aba_liga:25:1:7"
    assert gaps[0].required_field == "official_schedule_time"
    assert gaps[0].reason.value == "not_applicable"
    assert "no canonical game emitted" in gaps[0].detail


def test_calendar_accepts_cest_with_informational_local_time_annotation():
    calendar = CALENDAR.replace(
        "Friday, 10.10.2025 20:00 CET",
        "Friday, 10.10.2025 20:00 CEST (21:00 Example City local time)",
    )

    games = aba.parse_calendar(calendar, 2026)
    assert games[1].scheduled_utc == "2025-10-10T18:00:00Z"


def test_calendar_skips_date_only_nameless_placeholder_before_team_parsing():
    calendar = (
        CALENDAR.replace("Friday, 10.10.2025 20:00 CET", UNPINNED_DATE)
        .replace("Budućnost VOLI <span>:</span> FMP", "<span>:</span>")
    )

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        if "/match/6/" in url:
            return MATCH
        raise AssertionError(url)

    games = aba.discover_games("20251005", text_fetch=fetch, season=2026)
    assert [game.game_id for game in games] == ["aba_liga:25:1:6"]
    gaps = aba.last_gaps()
    assert len(gaps) == 1
    assert gaps[0].game_id == "aba_liga:25:1:7"
    assert gaps[0].reason.value == "not_applicable"


def test_calendar_skips_unscored_undated_placeholder_before_team_parsing():
    placeholder = """
<tr><td><p class="hidden-xs"><a href="/match/8/25/1/Overview">SCD <span>:</span> KRK</a></p></td></tr>
"""
    calendar = CALENDAR.replace("</tbody>", f"{placeholder}</tbody>")

    games = aba.parse_calendar(calendar, 2026)

    assert [game.game_no for game in games] == [6, 7]
    assert any(
        gap.game_id == "aba_liga:25:1:8"
        and gap.required_field == "official_schedule_time"
        and gap.reason.value == "not_applicable"
        and "unscored undated placeholder" in gap.detail
        for gap in aba.last_gaps()
    )


@pytest.mark.parametrize(
    "cell_text",
    (
        "SCD : KRK 20 : 0",
        "SCD : KRK 05.02.2033",
        "SCD : KRK 20:00 CET",
    ),
)
def test_calendar_rejects_scored_or_timed_undated_single_cell_row(cell_text):
    placeholder = f"""
<tr><td><p class="hidden-xs"><a href="/match/8/25/1/Overview">{cell_text}</a></p></td></tr>
"""
    calendar = CALENDAR.replace("</tbody>", f"{placeholder}</tbody>")

    with pytest.raises(IdentityError, match="incomplete scored or timed row"):
        aba.parse_calendar(calendar, 2026)


def test_calendar_rejects_scored_three_cell_row_without_a_timestamp():
    malformed = """
<tr>
<td><p class="hidden-xs"><a href="/match/8/25/1/Overview">FMP <span>:</span> OLI</a></p></td>
<td>64 : 63</td><td></td>
</tr>
"""
    calendar = CALENDAR.replace("</tbody>", f"{malformed}</tbody>")

    with pytest.raises(IdentityError, match="calendar row width 3 != 4"):
        aba.parse_calendar(calendar, 2026)


def test_calendar_with_only_unpinned_future_rows_is_a_visible_empty_slate():
    calendar = (
        CALENDAR.replace("93 : 97", "")
        .replace("Sunday, 05.10.2025 19:30 CET", "Tuesday, 01.02.2033 - Friday, 04.02.2033")
        .replace("Friday, 10.10.2025 20:00 CET", UNPINNED_DATE)
        .replace("Budućnost VOLI <span>:</span> FMP", "<span>:</span>")
    )

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        raise AssertionError(url)

    assert aba.discover_games("20251005", text_fetch=fetch, season=2026) == []
    gaps = aba.last_gaps()
    assert len(gaps) == 2
    assert all(gap.reason.value == "not_applicable" for gap in gaps)
    assert all(gap.required_field == "official_schedule_time" for gap in gaps)


def test_result_contract_reconciles_all_quarters_before_canonical_final():
    finals = aba.fetch_recent_results("2025-10-05", text_fetch=_fetch)

    assert len(finals) == 1
    actual = finals[0].actual
    assert actual is not None
    assert actual.home_quarters == [21.0, 25.0, 20.0, 27.0]
    assert actual.away_quarters == [24.0, 24.0, 23.0, 26.0]
    assert actual.home_final == 93.0 and actual.away_final == 97.0
    assert actual.reconciles() == (True, "")


def test_ot_final_with_two_overtime_periods_aggregates_and_reconciles():
    overtime_match = (
        MATCH.replace("93 : 97", "103 : 108", 1)
        .replace(
            "<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td></tr>",
            "<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td><td>OT1</td><td>OT2</td></tr>",
        )
        .replace(
            "<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>",
            "<tr><td>20 : 20</td><td>20 : 20</td><td>20 : 20</td><td>20 : 20</td><td>10 : 10</td><td>13 : 18</td></tr>",
        )
    )

    def fetch(url: str) -> str:
        return CALENDAR if "/calendar/" in url else overtime_match

    final = aba.fetch_recent_results("2025-10-05", text_fetch=fetch)[0]
    assert final.actual is not None
    assert final.actual.home_overtime == 23.0
    assert final.actual.away_overtime == 28.0
    assert final.actual.overtime_periods == 2
    assert final.actual.reconciles() == (True, "")
    assert final.final_status() is FinalStatus.AOT


def test_match_detail_accepts_cest_wall_clock_suffix():
    cesty_match = MATCH.replace("19:30 CET", "19:30 CEST (20:30 Example City local time)")

    def fetch(url: str) -> str:
        return CALENDAR if "/calendar/" in url else cesty_match

    games = aba.discover_games("20251005", text_fetch=fetch)
    assert [game.scheduled_utc for game in games] == ["2025-10-05T17:30:00Z"]


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


def test_default_forfeit_score_requires_and_records_generic_administrative_disposition():
    def fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else ADMINISTRATIVE_MATCH

    games, attempted, exhausted, calendar_final_count = aba.season_finals(2026, text_fetch=fetch)

    assert games == []
    assert attempted == [6]
    assert exhausted is True
    assert calendar_final_count == 1
    assert aba.last_administrative_dispositions() == [
        aba.AdministrativeDisposition(
            season_end_year=2026,
            game_no=6,
            game_id="aba_liga:25:1:6",
            calendar_score=(20, 0),
            disposition_kind="official_default_award_without_periods",
            calendar_has_exact_tip=True,
        )
    ]
    assert aba.last_history_progress_game_no() == 6
    gap = aba.last_gaps()[0]
    assert gap.required_field == "official_administrative_disposition"
    assert gap.reason.value == "not_applicable"


def test_administrative_tip_verification_accepts_local_time_but_not_a_different_clock():
    local_time_match = ADMINISTRATIVE_MATCH.replace("19:30 CET", "19:30 Local time")

    def fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else local_time_match

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert exhausted is True
    assert aba.last_administrative_dispositions()[0].calendar_has_exact_tip is True

    wrong_clock = local_time_match.replace("19:30 Local time", "20:30 Local time")

    def wrong_clock_fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else wrong_clock

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=wrong_clock_fetch)
    assert games == []
    assert exhausted is False
    assert aba.last_administrative_dispositions() == []
    assert "scheduled time contradicts calendar" in aba.last_gaps()[0].detail


def test_administrative_identity_uses_scoped_team_ids_not_city_or_sponsor_display_suffix():
    calendar = ADMINISTRATIVE_CALENDAR.replace(
        "Budućnost VOLI <span>:</span> Spartak Office Shoes",
        "Koper Primorska <span>:</span> Partizan Mozzart Bet",
    )
    suffixed_detail = (
        ADMINISTRATIVE_MATCH.replace("Budućnost VOLI", "Koper Primorska")
        .replace("Spartak Office Shoes", "Partizan Mozzart Bet Beograd")
    )

    def fetch(url: str) -> str:
        return calendar if "/calendar/" in url else suffixed_detail

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert exhausted is True
    assert len(aba.last_administrative_dispositions()) == 1

    out_of_scope = suffixed_detail.replace("/team/91/25/1/", "/team/91/24/1/")

    def out_of_scope_fetch(url: str) -> str:
        return calendar if "/calendar/" in url else out_of_scope

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=out_of_scope_fetch)
    assert games == []
    assert exhausted is False
    assert aba.last_administrative_dispositions() == []
    assert "outside requested scope" in aba.last_gaps()[0].detail


def test_nonempty_but_unreadable_period_table_still_blocks_default_score():
    malformed_periods = ADMINISTRATIVE_MATCH.replace(
        EMPTY_PERIOD_CONTAINER,
        "<table id=\"match_cetrtine_rezultat\"><tbody>"
        "<tr><td>Q1</td><td>Q2</td><td>Q3</td><td>Q4</td></tr>"
        "<tr><td></td><td></td><td></td><td></td></tr>"
        "</tbody></table>",
    )

    def fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else malformed_periods

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert exhausted is False
    assert aba.last_administrative_dispositions() == []
    assert "malformed" in aba.last_gaps()[0].detail


def test_retroactive_marker_must_be_in_match_content_subtree():
    marker_outside_content = MATCH_CONTENT.replace("93 : 97", "20 : 0").replace(
        "</div></body></html>",
        "</div><div><i>This match took place, but has been later registered with official result."
        "</i></div></body></html>",
    )

    def fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else marker_outside_content

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert exhausted is False
    assert aba.last_administrative_dispositions() == []
    assert "does not reconcile" in aba.last_gaps()[0].detail


def test_retroactively_registered_default_is_administrative_but_a_played_20_to_0_is_not():
    def retroactive_fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else RETROACTIVELY_REGISTERED_MATCH

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=retroactive_fetch)
    assert games == []
    assert exhausted is True
    assert aba.last_administrative_dispositions()[0].disposition_kind == (
        "retroactively_registered_official_result"
    )

    def played_fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else PLAYED_TWENTY_TO_ZERO_MATCH

    games, _attempted, exhausted, _count = aba.season_finals(2026, text_fetch=played_fetch)
    assert exhausted is True
    assert len(games) == 1
    assert games[0].actual is not None and games[0].actual.reconciles() == (True, "")
    assert aba.last_administrative_dispositions() == []


def test_forfeit_score_alone_or_malformed_periods_do_not_bypass_integrity_checks():
    malformed_detail = MATCH.replace("93 : 97", "20 : 0")

    def fetch(url: str) -> str:
        return ADMINISTRATIVE_CALENDAR if "/calendar/" in url else malformed_detail

    games, attempted, exhausted, _count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert attempted == [6]
    assert exhausted is False
    assert aba.last_administrative_dispositions() == []
    assert aba.last_gaps()[0].required_field == "historical_official_result"
    assert "does not reconcile" in aba.last_gaps()[0].detail


def test_undated_default_award_is_not_assigned_an_invented_tip():
    undated_calendar = ADMINISTRATIVE_CALENDAR.replace(
        "Sunday, 05.10.2025 19:30 CET", ""
    )

    def fetch(url: str) -> str:
        return undated_calendar if "/calendar/" in url else ADMINISTRATIVE_MATCH

    games, attempted, exhausted, calendar_final_count = aba.season_finals(2026, text_fetch=fetch)
    assert games == []
    assert attempted == [6]
    assert exhausted is True
    assert calendar_final_count == 1
    disposition = aba.last_administrative_dispositions()[0]
    assert disposition.calendar_has_exact_tip is False
    assert disposition.disposition_kind == "official_default_award_without_periods"


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

    games, attempted, exhausted, calendar_final_count = aba.season_finals(2026, text_fetch=fetch)
    assert [game.game_id for game in games] == ["aba_liga:25:1:6"]
    assert attempted == [6, 8]
    assert exhausted is False
    assert calendar_final_count == 2
    assert aba.last_gaps()[0].required_field == "historical_official_result"


def test_pipeline_history_writes_reconciled_final_and_advances_game_cursor(tmp_path):
    calendar = CALENDAR.replace("Friday, 10.10.2025 20:00 CET", UNPINNED_DATE_RANGE)

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        if "/match/6/" in url:
            return MATCH
        raise AssertionError(url)

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
            fetch=fetch,
            allow_writes=True,
        )
        report = pipeline.backfill()
        deep = report["backfill"]["deep_history"]
        assert deep["status"] == "acquired"
        assert deep["events_acquired"] == 1
        assert deep["final_scores_written"] == 1
        assert deep["quarantined_rows"] == 1
        assert any(
            gap.required_field == "official_schedule_time" and gap.reason.value == "not_applicable"
            for gap in pipeline.gaps
        )
        receipt = {
            "season_end_year": 2026,
            "calendar_final_count": 1,
            "reconciled_final_count": 1,
            "administrative_disposition_count": 0,
            # The unpinned placeholder is visible but not part of the
            # final-score denominator.
            "not_applicable_rows": 1,
            "blocking_gap_count": 0,
            "status": "calendar_terminal_count_reconciled",
            "source": "aba_liga",
            "evidence": aba.EVIDENCE,
            "completed_at_utc": "2025-10-06T12:00:00Z",
        }
        assert deep["season_receipts"] == [receipt]
        assert deep["cursor"] == {
            "floor_season": 2002,
            "last_game_no": 0,
            "last_completed_season": 2026,
            "last_completed_final_count": 1,
            "next_season": 2025,
            "season_receipts": {"2026": receipt},
        }
        stored = store.get_game("aba_liga:25:1:6")
        assert stored is not None and stored.actual is not None
        assert stored.actual.reconciles() == (True, "")
        health = store.get_provider_health("aba-liga-history", League.ABA_LEAGUE)
        assert health is not None and health.coverage == deep["cursor"]
    finally:
        store.close()


def test_history_receipt_separates_verified_administrative_dispositions(tmp_path):
    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return ADMINISTRATIVE_CALENDAR
        if "/match/6/" in url:
            return ADMINISTRATIVE_MATCH
        raise AssertionError(url)

    store = DuckDBStore(str(tmp_path / "aba-admin" / "warehouse.duckdb"), persist=True)
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
            fetch=fetch,
            allow_writes=True,
        )
        deep = pipeline.backfill()["backfill"]["deep_history"]

        assert deep["status"] == "acquired"
        assert deep["events_acquired"] == 0
        assert deep["final_scores_written"] == 0
        assert deep["administrative_disposition_count"] == 1
        assert deep["season_receipts"] == [
            {
                "season_end_year": 2026,
                "calendar_final_count": 1,
                "reconciled_final_count": 0,
                "administrative_disposition_count": 1,
                "not_applicable_rows": 0,
                "blocking_gap_count": 0,
                "status": "calendar_terminal_count_reconciled",
                "source": "aba_liga",
                "evidence": aba.EVIDENCE,
                "completed_at_utc": "2025-10-06T12:00:00Z",
            }
        ]
        assert store.get_game("aba_liga:25:1:6") is None
        assert any(
            gap.required_field == "official_administrative_disposition"
            and gap.reason.value == "not_applicable"
            for gap in pipeline.gaps
        )
    finally:
        store.close()


def test_bounded_history_cursor_carries_generic_administrative_count_across_batches(tmp_path):
    calendar = ADMINISTRATIVE_CALENDAR.replace(
        "<td></td><td>Friday, 10.10.2025 20:00 CET</td>",
        "<td>80 : 60</td><td>Friday, 10.10.2025 20:00 CET</td>",
    )
    later_played_match = (
        MATCH.replace('id="match_data_id" value="6"', 'id="match_data_id" value="7"')
        .replace("Sunday, 05.10.2025 19:30 CET", "Friday, 10.10.2025 20:00 CET")
        .replace("Spartak Office Shoes", "FMP")
        .replace("93 : 97", "80 : 60")
        .replace(
            "<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>",
            "<tr><td>20 : 15</td><td>20 : 15</td><td>20 : 15</td><td>20 : 15</td></tr>",
        )
    )

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        if "/match/6/" in url:
            return ADMINISTRATIVE_MATCH
        if "/match/7/" in url:
            return later_played_match
        raise AssertionError(url)

    store = DuckDBStore(str(tmp_path / "aba-bounded-admin" / "warehouse.duckdb"), persist=True)
    config = {
        "league": "ABA_LEAGUE",
        "date_yyyymmdd": "20251006",
        "history_days": 0,
        "history_request_budget": 1,
        "fixture_now_utc": "2025-10-06T12:00:00Z",
    }
    try:
        first = Pipeline(
            PipelineConfig(**config, deep_history_event_budget=1), store, fetch=fetch, allow_writes=True
        ).backfill()["backfill"]["deep_history"]
        assert first["events_acquired"] == 0
        assert first["cursor"] == {
            "administrative_disposition_count": 1,
            "floor_season": 2002,
            "last_game_no": 6,
            "next_season": 2026,
        }

        second = Pipeline(
            PipelineConfig(**config, deep_history_event_budget=5), store, fetch=fetch, allow_writes=True
        ).backfill()["backfill"]["deep_history"]
        receipt = second["season_receipts"][0]
        assert receipt["calendar_final_count"] == 2
        assert receipt["reconciled_final_count"] == 1
        assert receipt["administrative_disposition_count"] == 1
        assert store.get_game("aba_liga:25:1:7") is not None
    finally:
        store.close()


def test_history_withholds_receipt_and_rewinds_when_a_final_was_not_persisted():
    class DroppingStore(InMemoryStore):
        # Exercise the same write-required ABA path while simulating a storage
        # failure after a fully reconciled adapter batch.
        persist = True

        def upsert_game(self, game):
            return None

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return CALENDAR
        if "/match/6/" in url:
            return MATCH
        raise AssertionError(url)

    store = DroppingStore()
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
        fetch=fetch,
        allow_writes=True,
    )

    deep = pipeline.backfill()["backfill"]["deep_history"]

    assert deep["status"] == "provider_error_cursor_retained"
    assert deep["season_receipts"] == []
    assert deep["cursor"] == {
        "floor_season": 2002,
        "last_game_no": 0,
        "next_season": 2026,
    }
    assert any(
        gap.required_field == "aba_liga_season_receipt"
        and gap.reason.value == "missing"
        and "receipt withheld" in gap.detail
        for gap in pipeline.gaps
    )


def test_history_withholds_receipt_for_an_all_placeholder_calendar(tmp_path):
    calendar = (
        CALENDAR.replace("93 : 97", "")
        .replace("Sunday, 05.10.2025 19:30 CET", UNPINNED_DATE_RANGE)
        .replace("Friday, 10.10.2025 20:00 CET", UNPINNED_DATE)
    )

    def fetch(url: str) -> str:
        if "/calendar/" in url:
            return calendar
        raise AssertionError(url)

    store = DuckDBStore(str(tmp_path / "aba-empty" / "warehouse.duckdb"), persist=True)
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
            fetch=fetch,
            allow_writes=True,
        )

        deep = pipeline.backfill()["backfill"]["deep_history"]

        assert deep["status"] == "provider_error_cursor_retained"
        assert deep["season_receipts"] == []
        assert deep["cursor"] == {
            "floor_season": 2002,
            "last_game_no": 0,
            "next_season": 2026,
        }
        assert any(
            gap.required_field == "aba_liga_season_receipt"
            and gap.reason.value == "missing"
            and "no final rows" in gap.detail
            for gap in pipeline.gaps
        )
    finally:
        store.close()


@pytest.mark.parametrize(
    ("season", "next_season", "expected_text"),
    (
        (2009, 2008, "2008-09"),
        (2020, 2019, "2019-20"),
    ),
)
def test_history_explicitly_excludes_approved_season_without_fetching_or_silent_skip(
    season, next_season, expected_text
):
    store = InMemoryStore()
    store.persist = True
    store.upsert_provider_health(
        ProviderHealth(
            provider="aba-liga-history",
            league=League.ABA_LEAGUE,
            coverage={
                "floor_season": 2002,
                "last_game_no": 0,
                "next_season": season,
            },
        )
    )

    pipeline = Pipeline(
        PipelineConfig(
            league="ABA_LEAGUE",
            date_yyyymmdd="20261120",
            history_days=0,
            history_request_budget=1,
            deep_history_event_budget=5,
            fixture_now_utc="2026-11-20T12:00:00Z",
        ),
        store,
        fetch=lambda url: (_ for _ in ()).throw(AssertionError(url)),
        allow_writes=True,
    )

    deep = pipeline.backfill()["backfill"]["deep_history"]

    exclusion = {
        "season_end_year": season,
        "status": "excluded_by_operator_policy",
        "detail": aba.HISTORY_EXCLUDED_SEASONS[season],
    }
    assert deep["excluded_seasons"] == [exclusion]
    assert deep["seasons"] == [exclusion]
    assert deep["cursor"] == {
        "floor_season": 2002,
        "last_game_no": 0,
        "next_season": next_season,
    }
    assert any(
        gap.required_field == "aba_liga_history_policy"
        and gap.reason.value == "not_applicable"
        and expected_text in gap.detail
        for gap in pipeline.gaps
    )


def test_history_replays_current_season_when_a_late_final_appears_below_cursor(tmp_path):
    late_calendar = CALENDAR.replace(
        "<td></td><td>Friday, 10.10.2025 20:00 CET</td>",
        "<td>80 : 60</td><td>Friday, 10.10.2025 20:00 CET</td>",
    )
    late_match = (
        MATCH.replace('id="match_data_id" value="6"', 'id="match_data_id" value="7"')
        .replace("Sunday, 05.10.2025 19:30 CET", "Friday, 10.10.2025 20:00 CET")
        .replace("Spartak Office Shoes", "FMP")
        .replace("93 : 97", "80 : 60")
        .replace(
            "<tr><td>21 : 24</td><td>25 : 24</td><td>20 : 23</td><td>27 : 26</td></tr>",
            "<tr><td>20 : 15</td><td>20 : 15</td><td>20 : 15</td><td>20 : 15</td></tr>",
        )
    )

    def fetch(calendar: str):
        def _fetch(url: str) -> str:
            if "/calendar/" in url:
                return calendar
            if "/match/6/" in url:
                return MATCH
            if "/match/7/" in url:
                return late_match
            raise AssertionError(url)

        return _fetch

    config = PipelineConfig(
        league="ABA_LEAGUE",
        date_yyyymmdd="20251006",
        history_days=0,
        history_request_budget=1,
        deep_history_event_budget=5,
        fixture_now_utc="2025-10-06T12:00:00Z",
    )
    store = DuckDBStore(str(tmp_path / "aba-repair" / "warehouse.duckdb"), persist=True)
    try:
        first = Pipeline(config, store, fetch=fetch(CALENDAR), allow_writes=True).backfill()
        first_cursor = first["backfill"]["deep_history"]["cursor"]
        assert first_cursor["last_completed_season"] == 2026
        assert first_cursor["last_completed_final_count"] == 1

        second = Pipeline(config, store, fetch=fetch(late_calendar), allow_writes=True).backfill()
        deep = second["backfill"]["deep_history"]
        assert deep["completeness_repair"] == {
            "season_end_year": 2026,
            "calendar_final_count": 2,
            "stored_reconciled_final_count": 1,
        }
        repaired = store.get_game("aba_liga:25:1:7")
        assert repaired is not None and repaired.actual is not None
        assert repaired.actual.reconciles() == (True, "")
        assert deep["cursor"]["last_completed_final_count"] == 2
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
    with pytest.raises(IdentityError, match="no readable CET/CEST/Local time timestamp"):
        aba.parse_calendar(CALENDAR.replace("Friday, 10.10.2025 20:00 CET", "TBD"), 2026)
    with pytest.raises(IdentityError, match="completed calendar game 7 has an unpinned date range"):
        aba.parse_calendar(
            CALENDAR.replace(
                "<td></td><td>Friday, 10.10.2025 20:00 CET</td>",
                f"<td>80 : 70</td><td>{UNPINNED_DATE_RANGE}</td>",
            ),
            2026,
        )
    with pytest.raises(IdentityError, match="completed calendar game 7 has an unpinned date"):
        aba.parse_calendar(
            CALENDAR.replace(
                "<td></td><td>Friday, 10.10.2025 20:00 CET</td>",
                f"<td>80 : 70</td><td>{UNPINNED_DATE}</td>",
            ),
            2026,
        )
    with pytest.raises(IdentityError, match="calendar matchup has an unresolved team"):
        aba.parse_calendar(CALENDAR.replace("Budućnost VOLI <span>:</span> FMP", "<span>:</span>"), 2026)
