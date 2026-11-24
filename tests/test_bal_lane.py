"""Basketball Africa League lane: the season blob behind bal.nba.com/schedule.

The game object here is **verbatim** from a captured ``__NEXT_DATA__`` payload
(HTTP 200, 379,284 B, 23 game dates, 2026-11-23T12:12:51Z, out-of-sandbox probe —
this sandbox has no egress to bal.nba.com). The surrounding
``props.pageProps.gameDates`` envelope is synthesised and labelled as such.

Cases covered: registry wiring, blob extraction, the UTC tip, the settled-score
rule (a 0-0 row is never a final), postponement, live, the date filter, a missing
blob, a missing gameDates list, identity failures, the probe receipt, and the
gate.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List

import pytest

from ingestion.league_registry import league_spec
from ingestion.models import GameStatus, League, PeriodKind
from ingestion.providers import bal
from ingestion.providers import IdentityError

FIXTURE = "tests/fixtures/bal/captures/schedule-next-data-2026-03-27.json"


def _payload() -> Dict[str, Any]:
    with open(FIXTURE, encoding="utf-8") as handle:
        return json.loads(handle.read())


def _row() -> Dict[str, Any]:
    return _payload()["props"]["pageProps"]["gameDates"][0]["games"][0]


def _html(payload: Any | None = None) -> str:
    body = json.dumps(_payload() if payload is None else payload, ensure_ascii=False)
    return (
        "<html><head><script id=\"__NEXT_DATA__\" type=\"application/json\">"
        + body
        + "</script></head><body></body></html>"
    )


def _fetch(payload: Any | None = None) -> Callable[[str], str]:
    page = _html(payload)

    def fetch(url: str) -> str:
        if "bal.nba.com/schedule" not in url:
            raise AssertionError(f"unexpected URL: {url}")
        return page

    return fetch


def test_registry_points_the_lane_at_the_nba_host():
    spec = league_spec(League.BASKETBALL_AFRICA_LEAGUE)
    assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
        "bal",
        "bal",
        "bal",
    )
    assert "__NEXT_DATA__" in spec.source_detail


def test_blob_is_extracted_from_the_served_html():
    payload = bal.extract_schedule(_html())

    assert payload["props"]["pageProps"]["gameDates"][0]["seasonYear"] == 2026


def test_page_without_the_blob_raises():
    with pytest.raises(ValueError):
        bal.extract_schedule("<html><body>nothing here</body></html>")


def test_payload_without_game_dates_raises():
    with pytest.raises(ValueError):
        bal._game_dates({"props": {"pageProps": {}}})


def test_unplayed_row_stays_scheduled_and_carries_no_result():
    game = bal.game_from_row(_row(), 2026)

    # The captured row is gameStatus 1 with 0-0 and a future tip.
    assert game.status is GameStatus.SCHEDULED
    assert game.actual is None
    assert game.game_id == "BAL-1122600001"
    assert game.league is League.BASKETBALL_AFRICA_LEAGUE
    assert game.home_name == "RSSB Tigers"
    assert game.away_name == "Libya Al Ahly Ly"
    assert game.scheduled_utc == "2026-03-27T14:00:00Z"
    assert game.season == "2026"


def test_a_zero_zero_scoreline_is_never_a_final():
    row = _row()
    row["gameStatus"] = 3  # a code this adapter does not trust on its own
    row["gameStatusText"] = "Final"

    game = bal.game_from_row(row, 2026)

    assert game.status is GameStatus.SCHEDULED
    assert game.actual is None


def test_settled_scores_make_a_totals_only_final():
    row = _row()
    row["homeTeam"]["score"] = 88
    row["awayTeam"]["score"] = 79

    game = bal.game_from_row(row, 2026)

    assert game.status is GameStatus.FINAL
    assert game.actual.period_kind is PeriodKind.UNKNOWN
    assert (game.actual.home_final, game.actual.away_final) == (88.0, 79.0)
    assert game.actual.home_quarters == []


def test_postponed_row_is_marked_postponed():
    row = _row()
    row["postponedStatus"] = "P"

    assert bal.game_from_row(row, 2026).status is GameStatus.POSTPONED


def test_live_row_is_live_only_without_a_final_score():
    row = _row()
    row["gameStatus"] = 2

    assert bal.game_from_row(row, 2026).status is GameStatus.LIVE


def test_identical_team_ids_are_refused():
    row = _row()
    row["awayTeam"]["teamId"] = row["homeTeam"]["teamId"]

    with pytest.raises(IdentityError):
        bal.game_from_row(row, 2026)


def test_date_filter_keeps_only_the_target_utc_day():
    assert [g.game_id for g in bal.discover_games("2026-03-27", _fetch())] == [
        "BAL-1122600001"
    ]
    assert bal.discover_games("2026-03-28", _fetch()) == []


def test_probe_records_the_season_read():
    bal.discover_games("2026-03-27", _fetch())

    probe = bal.last_probe()
    assert probe["url"] == bal.SCHEDULE_URL
    assert probe["game_dates"] == 1
    assert probe["rows"] == 1
    assert probe["games"] == 1
    assert len(probe["sha256"]) == 64


def test_gate_reads_the_blob_and_derives_state_from_scores(monkeypatch):
    from ingestion import adaptive_gate

    seen: List[str] = []

    def fake_read(url: str, headers=None, **kwargs):  # noqa: ANN001, ANN003
        seen.append(url)
        return _html()

    monkeypatch.setattr(adaptive_gate, "_read_limited_response", fake_read)
    config = adaptive_gate.LeagueGateConfig(
        label="Basketball Africa League",
        league="basketball_africa_league",
        league_key="BASKETBALL_AFRICA_LEAGUE",
        state_key="basketball_africa_league",
        acquisition_timezone="Africa/Johannesburg",
        schedule_source="bal",
    )
    slate = adaptive_gate._bal_schedule(config, "2026-03-27")

    assert len(seen) == 1
    assert len(slate) == 1
    assert slate[0].game_id == "BAL-1122600001"
    assert slate[0].status == "scheduled"
    assert slate[0].scheduled_utc == datetime(2026, 3, 27, 14, 0, tzinfo=timezone.utc)
    assert adaptive_gate._bal_schedule(config, "2026-03-28") == []


def _payload_with(rows: List[Any], season_year: int = 2026) -> Dict[str, Any]:
    """A one-game-date envelope carrying the supplied rows verbatim."""

    return {
        "props": {
            "pageProps": {
                "gameDates": [{"seasonYear": season_year, "games": rows}],
            }
        }
    }


def test_off_date_rows_are_counted_not_silently_dropped():
    """Run 45250804751: ``rows: 42``, ``games: 0``, ``off_date_dropped: 0``.

    The bare ``continue`` upstream of every counter made those three numbers
    consistent with a lane that read 42 rows and explained none of them.
    """

    row = _row()
    off_date = json.loads(json.dumps(row))
    off_date["gameDateTimeUTC"] = "2026-03-28T18:00:00Z"
    payload = _payload_with([row, off_date, json.loads(json.dumps(off_date))])

    games = bal.discover_games("2026-03-27", _fetch(payload))

    assert [g.game_id for g in games] == ["BAL-1122600001"]
    probe = bal.last_probe()
    assert probe["rows"] == 3
    assert probe["games"] == 1
    assert probe["off_date_dropped"] == 2
    assert probe["missing_stamp_dropped"] == 0
    assert probe["non_dict_rows"] == 0


def test_a_row_without_a_stamp_is_counted_under_its_own_reason():
    row = _row()
    stampless = json.loads(json.dumps(row))
    stampless["gameDateTimeUTC"] = None
    empty = json.loads(json.dumps(row))
    empty["gameDateTimeUTC"] = ""

    games = bal.discover_games("2026-03-27", _fetch(_payload_with([row, stampless, empty])))

    assert [g.game_id for g in games] == ["BAL-1122600001"]
    probe = bal.last_probe()
    assert probe["missing_stamp_dropped"] == 2
    assert probe["off_date_dropped"] == 0


def test_non_dict_rows_are_counted_rather_than_skipped():
    payload = _payload_with([_row(), "not-a-row", None, ["also", "not", "a", "row"]])

    games = bal.discover_games("2026-03-27", _fetch(payload))

    assert [g.game_id for g in games] == ["BAL-1122600001"]
    assert bal.last_probe()["non_dict_rows"] == 3
    assert bal.last_probe()["rows"] == 4


def test_unfiltered_read_counts_nothing_as_dropped():
    """With no date filter there is no drop decision to count."""

    bal.discover_games(None, _fetch())

    probe = bal.last_probe()
    assert probe["games"] == 1
    assert probe["off_date_dropped"] == 0
    assert probe["missing_stamp_dropped"] == 0
    assert probe["non_dict_rows"] == 0


def test_pipeline_receipt_prints_the_bal_counters():
    from ingestion.pipeline import Pipeline, PipelineConfig
    from ingestion.store import InMemoryStore

    row = _row()
    off_date = json.loads(json.dumps(row))
    off_date["gameDateTimeUTC"] = "2026-03-28T18:00:00Z"
    page = _html(_payload_with([row, off_date]))
    payload = PipelineConfig(league="BASKETBALL_AFRICA_LEAGUE", date_yyyymmdd="20260327")
    pipeline = Pipeline(payload, InMemoryStore(), fetch=lambda url: page, allow_writes=False)

    assert pipeline._discover_games()

    row_out = next(h for h in pipeline._health_written if h["provider"] == "bal")
    assert row_out["coverage"]["off_date_dropped"] == 1
    assert row_out["coverage"]["missing_stamp_dropped"] == 0
    assert row_out["coverage"]["non_dict_rows"] == 0
    assert "off_date_dropped=1" in row_out["detail"]
