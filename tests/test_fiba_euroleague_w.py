"""Hermetic official-FIBA EuroLeague Women routing and result regressions."""

from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

import ingestion.adaptive_gate as adaptive_gate
from forecasting.model import league_forecast_config
from forecasting.service import MarketForecastService
from ingestion.cli import _warehouse
from ingestion.duckdb_store import DuckDBStore
from ingestion.models import GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import fiba_euroleague_w as fiba
from ingestion.store import InMemoryStore

FIXTURES = Path(__file__).parent / "fixtures" / "fiba_euroleague_w"
SCHEDULE = json.loads((FIXTURES / "schedule_209127_20261118.json").read_text())
DETAIL = json.loads((FIXTURES / "game_detail_135150_20261118.json").read_text())


def _fiba_fetch(url: str):
    if "getgdapgamesbycompetitionid" in url:
        return deepcopy(SCHEDULE)
    if "/games/135150-ZARAG-BLA" in url:
        return deepcopy(DETAIL)
    raise AssertionError(f"unexpected FIBA URL: {url}")


def test_fiba_elw_schedule_uses_competition_209127_canonical_utc_and_isolated_ids():
    urls: list[str] = []

    def fetch(url: str):
        urls.append(url)
        return deepcopy(SCHEDULE)

    games = fiba.discover_games("2026-10-14", fetch)

    assert "gdapCompetitionId=209127" in urls[0]
    assert games[0].league is League.EUROLEAGUE_W
    assert games[0].game_id == "fiba-elw-209127-135150"
    assert games[0].home_team_id == "fiba-elw:209127:285048"
    assert games[0].away_team_id == "fiba-elw:209127:285044"
    assert games[0].scheduled_utc == "2026-10-14T17:00:00Z"
    assert games[0].provider_ids["fiba_elw_competition"] == "209127"
    assert games[0].provider_ids["fiba_elw_event_slug"] == "euroleague-women-26-27"
    assert games[0].status is GameStatus.FINAL


def test_fiba_elw_skips_future_teamless_slots_and_reports_bounded_coverage():
    games = fiba.discover_games("2026-10-14", _fiba_fetch)

    assert [game.game_id for game in games] == ["fiba-elw-209127-135150"]
    gaps = fiba.last_gaps()
    assert len(gaps) == 1
    assert gaps[0].required_field == "fiba_elw_unassigned_schedule_slots"
    assert "skipped 1" in gaps[0].detail
    assert "0 fall on requested Europe/Paris slate 2026-10-14" in gaps[0].detail


def test_fiba_elw_all_teamless_requested_slate_is_an_explicit_empty_slate_not_an_error():
    schedule = deepcopy(SCHEDULE)
    schedule[0]["teamA"] = None
    schedule[0]["teamB"] = None

    games = fiba.discover_games("2026-10-14", lambda _url: schedule)

    assert games == []
    gap = fiba.last_gaps()[0]
    assert "skipped 2" in gap.detail
    assert "1 fall on requested Europe/Paris slate 2026-10-14" in gap.detail


def test_fiba_elw_teamless_rows_outside_a_slate_never_block_discovery():
    schedule = deepcopy(SCHEDULE)
    schedule[1]["gameDateTimeUTC"] = "2026-12-16T17:00:00"
    # A null team object is normal. Even an otherwise unusable placeholder
    # outside the requested day must not map through _clean_name.
    schedule[1]["gameId"] = None

    games = fiba.discover_games("2026-10-14", lambda _url: schedule)

    assert [game.game_id for game in games] == ["fiba-elw-209127-135150"]


def test_fiba_elw_skips_teamless_slots_with_missing_timestamps_and_reports_them():
    schedule = deepcopy(SCHEDULE)
    schedule[1]["gameDateTimeUTC"] = None

    games = fiba.discover_games("2026-10-14", lambda _url: schedule)

    assert [game.game_id for game in games] == ["fiba-elw-209127-135150"]
    assert "no parseable timestamp" in fiba.last_gaps()[0].detail


def test_fiba_elw_skips_bad_team_bearing_timestamps_without_blocking_slate_or_history():
    schedule = deepcopy(SCHEDULE)
    malformed = deepcopy(SCHEDULE[0])
    malformed["gameId"] = 135212
    malformed["gameDateTimeUTC"] = "not-an-iso-timestamp"
    schedule.append(malformed)

    slate = fiba.discover_games("2026-10-14", lambda _url: schedule)
    history = fiba.discover_games(None, lambda _url: schedule)

    assert [game.game_id for game in slate] == ["fiba-elw-209127-135150"]
    assert [game.game_id for game in history] == ["fiba-elw-209127-135150"]
    assert any(gap.required_field == "fiba_elw_schedule_timestamp" for gap in fiba.last_gaps())

    report = fiba.smoke(lambda _url: schedule)
    assert report["unassigned_schedule_slot_gaps"] == 1
    assert report["schedule_timestamp_gaps"] == 1


def test_fiba_elw_parses_explicit_offsets_once_in_provider_and_gate(monkeypatch):
    row = deepcopy(SCHEDULE[0])
    row["gameDateTimeUTC"] = "2026-10-14T19:00:00+02:00"

    game = fiba.discover_games("2026-10-14", lambda _url: [row])[0]
    assert game.scheduled_utc == "2026-10-14T17:00:00Z"

    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")
    gate_games = adaptive_gate._fiba_elw_schedule(config, "2026-10-14")
    assert gate_games[0].scheduled_utc.isoformat() == "2026-10-14T17:00:00+00:00"


def test_fiba_elw_gate_ignores_teamless_bracket_slots_before_timestamp_parsing(monkeypatch):
    row = deepcopy(SCHEDULE[0])
    row["teamA"] = None
    row["teamB"] = None
    row["gameDateTimeUTC"] = None
    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")

    assert adaptive_gate._fiba_elw_schedule(config, "2026-10-14") == []


def test_fiba_schedule_url_never_exposes_the_apim_key():
    assert "subscription-key" not in fiba.schedule_url()


def test_live_fiba_client_refreshes_page_shipped_key_and_keeps_it_out_of_schedule_url():
    class FakeHttp:
        def __init__(self):
            self.calls = []

        def get_text(self, _url, **_kwargs):
            return 'NEXT_CLIENT_APIM_SUBSCRIPTION_KEY\\\\":\\\\"0123456789abcdef0123456789abcdef\\\\"'

        def get_json(self, url, *, headers):
            self.calls.append((url, headers))
            return deepcopy(SCHEDULE)

    http = FakeHttp()
    payload = fiba.FibaClient(http).get_json(fiba.schedule_url())

    assert payload == SCHEDULE
    url, headers = http.calls[0]
    assert "subscription-key" not in url
    assert headers["Ocp-Apim-Subscription-Key"] == "0123456789abcdef0123456789abcdef"


def test_live_fiba_client_refuses_a_schedule_url_for_another_competition():
    with pytest.raises(fiba.VerificationError, match="non-ELW"):
        fiba.FibaClient().get_json(fiba.schedule_url(123456))


def test_fiba_game_page_404_names_the_event_route_risk():
    class FakeHttp:
        def get_text(self, url, **_kwargs):
            from ingestion.http import HttpError

            raise HttpError(404, url)

    with pytest.raises(fiba.VerificationError, match="event slug not found"):
        fiba.FibaClient(FakeHttp()).game_detail(
            "https://www.fiba.basketball/en/events/euroleague-women-26-27/games/135150-ZARAG-BLA"
        )


def test_elw_default_warehouse_isolated_from_the_shared_mens_path(monkeypatch):
    monkeypatch.delenv("MAGOLIDE_WAREHOUSE", raising=False)

    assert _warehouse(SimpleNamespace(warehouse=None, league="EUROLEAGUE_W")) == (
        "localdata/euroleague_w/warehouse.duckdb"
    )


def test_fiba_elw_game_detail_reconciles_official_quarters_and_overtime():
    game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    final = fiba.fetch_game_result(game, _fiba_fetch)

    assert final.actual is not None
    assert final.actual.home_quarters == [25, 13, 19, 11]
    assert final.actual.away_quarters == [14, 20, 20, 14]
    assert final.actual.home_overtime == 8
    assert final.actual.away_overtime == 7
    assert final.actual.overtime_periods == 1
    assert final.actual.reconciles() == (True, "")
    assert final.actual.source == "fiba-game-page"
    assert final.evidence["fiba"]["game_detail"]["team_detail_available"] is True
    assert final.evidence["fiba"]["game_detail"]["periods"] == ["Q1", "Q2", "Q3", "Q4", "OT1"]


def test_fiba_elw_refuses_game_detail_that_contradicts_gdap_final_score():
    game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    contradictory = deepcopy(DETAIL)
    contradictory["game"]["teamAScore"] = 77

    with pytest.raises(fiba.VerificationError, match="contradicts"):
        fiba.result_from_game_detail(game, contradictory)


def test_fiba_elw_rsc_extractor_selects_the_requested_game_bound_pbp_after_decoys():
    decoy = deepcopy(DETAIL["playByPlay"])
    decoy["gameTechnicalName"] = "999999"
    correct = deepcopy(DETAIL["playByPlay"])
    # A Flight-style escaped push carries both a valid-shape decoy and the
    # requested PBP. Exact game binding—not score coincidence or order—wins.
    wrapped = json.dumps(
        json.dumps({"playByPlay": decoy, "game": DETAIL["game"], "also": {"playByPlay": correct}})
    )
    flight = f"self.__next_f.push([1,{wrapped}])"

    extracted = fiba.extract_game_page(flight, "135150")

    assert extracted["game"]["competition"]["competitionId"] == 209127
    assert extracted["playByPlay"]["gameTechnicalName"] == "135150"
    assert extracted["playByPlay"]["items"]["OT1"]["scoreA"] == 76


def test_fiba_elw_refuses_game_detail_with_play_by_play_bound_to_another_game():
    game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    wrong_pbp = deepcopy(DETAIL)
    wrong_pbp["playByPlay"]["gameTechnicalName"] = "999999"

    with pytest.raises(fiba.VerificationError, match="play-by-play belongs to another game"):
        fiba.result_from_game_detail(game, wrong_pbp)


def test_fiba_elw_refuses_rsc_play_by_play_without_an_explicit_game_binding():
    flight = json.dumps({"game": DETAIL["game"], "playByPlay": {"items": DETAIL["playByPlay"]["items"]}})

    with pytest.raises(fiba.VerificationError, match="requested-game play-by-play"):
        fiba.extract_game_page(flight, "135150")


def test_fiba_elw_refuses_game_detail_with_reversed_fiba_team_identity():
    game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    reversed_teams = deepcopy(DETAIL)
    reversed_teams["game"]["teamA"], reversed_teams["game"]["teamB"] = (
        reversed_teams["game"]["teamB"],
        reversed_teams["game"]["teamA"],
    )

    with pytest.raises(fiba.IdentityError, match="team identity"):
        fiba.result_from_game_detail(game, reversed_teams)


def test_fiba_elw_refuses_incomplete_period_detail_instead_of_inventing_q4():
    game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    incomplete = deepcopy(DETAIL)
    del incomplete["playByPlay"]["items"]["Q4"]

    with pytest.raises(fiba.VerificationError, match="complete Q1-Q4"):
        fiba.result_from_game_detail(game, incomplete)


@pytest.mark.parametrize(
    ("status_code", "expected"),
    [
        ("CANCELLED", GameStatus.CANCELLED),
        ("CANCELED", GameStatus.CANCELLED),
        ("ABANDONED", GameStatus.ABANDONED),
        ("VOID", GameStatus.ABANDONED),
    ],
)
def test_fiba_terminal_status_contract_is_shared_with_the_fail_open_gate(
    monkeypatch, status_code, expected
):
    row = deepcopy(SCHEDULE[0])
    row["statusCode"] = status_code
    row["isLive"] = False
    row["gameResultStatusCode"] = "EMPTY"
    row["liveGameStatus"] = 0
    row["teamAScore"] = None
    row["teamBScore"] = None

    assert fiba.game_from_schedule_row(row).status is expected
    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")
    with pytest.raises(adaptive_gate.GateUncertain, match=expected.value.lower()):
        adaptive_gate._fiba_elw_schedule(config, "2026-10-14")


@pytest.mark.parametrize(
    ("overrides", "expected", "gate_accepts"),
    [
        (
            {"isPostponed": True, "statusCode": "CANCELLED", "isLive": True},
            GameStatus.POSTPONED,
            False,
        ),
        (
            {"isPostponed": False, "statusCode": "CANCELLED", "isLive": True},
            GameStatus.CANCELLED,
            False,
        ),
        (
            {"isPostponed": False, "statusCode": "VALID", "isLive": True},
            GameStatus.LIVE,
            True,
        ),
    ],
)
def test_fiba_status_precedence_is_shared_with_the_gate(monkeypatch, overrides, expected, gate_accepts):
    row = deepcopy(SCHEDULE[0])
    row.update(overrides)

    assert fiba.game_from_schedule_row(row).status is expected
    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")
    if gate_accepts:
        assert [game.status for game in adaptive_gate._fiba_elw_schedule(config, "2026-10-14")] == [
            expected.value.lower()
        ]
    else:
        with pytest.raises(adaptive_gate.GateUncertain, match=expected.value.lower()):
            adaptive_gate._fiba_elw_schedule(config, "2026-10-14")


@pytest.mark.parametrize(
    ("result_status", "scores"),
    [
        ("VALID", (None, None)),
        ("VALID", (True, 75)),
        ("EMPTY", (76, 75)),
    ],
)
def test_fiba_final_like_status_without_a_complete_valid_signature_fails_open(
    monkeypatch, result_status, scores
):
    row = deepcopy(SCHEDULE[0])
    row["statusCode"] = "VALID"
    row["gameResultStatusCode"] = result_status
    row["liveGameStatus"] = 999
    row["teamAScore"], row["teamBScore"] = scores
    row["isLive"] = False

    assert fiba.game_from_schedule_row(row).status is GameStatus.UNKNOWN
    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")
    with pytest.raises(adaptive_gate.GateUncertain, match="unknown status"):
        adaptive_gate._fiba_elw_schedule(config, "2026-10-14")


def test_fiba_unknown_status_is_not_silently_scheduled_by_the_gate(monkeypatch):
    row = deepcopy(SCHEDULE[0])
    row["statusCode"] = "NEW_UNDOCUMENTED_STATUS"
    row["gameResultStatusCode"] = "EMPTY"
    row["liveGameStatus"] = 0
    row["teamAScore"] = None
    row["teamBScore"] = None

    assert fiba.game_from_schedule_row(row).status is GameStatus.UNKNOWN
    monkeypatch.setattr(adaptive_gate, "_fiba_schedule_payload", lambda _url: [row])
    config = next(config for config in adaptive_gate.LEAGUES if config.league == "EUROLEAGUE_W")
    with pytest.raises(adaptive_gate.GateUncertain, match="unknown status"):
        adaptive_gate._fiba_elw_schedule(config, "2026-10-14")


def test_fiba_event_slug_tracks_the_gdap_edition_instead_of_a_frozen_season_template():
    row = deepcopy(SCHEDULE[0])
    row["competition"]["season"] = 2028

    game = fiba.game_from_schedule_row(row)

    assert game.provider_ids["fiba_elw_event_slug"] == "euroleague-women-27-28"
    assert "/events/euroleague-women-27-28/games/" in fiba.game_page_url(
        game.provider_ids["fiba_elw_game"],
        game.provider_ids["fiba_elw_team_a_code"],
        game.provider_ids["fiba_elw_team_b_code"],
        event_slug=game.provider_ids["fiba_elw_event_slug"],
    )


def test_fiba_elw_backfill_retains_cursor_when_every_detail_request_fails(tmp_path, monkeypatch):
    store = DuckDBStore(str(tmp_path / "warehouse.duckdb"), persist=True)
    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE_W",
            date_yyyymmdd="20261015",
            fixture_now_utc="2026-10-15T12:00:00Z",
            deep_history_event_budget=1,
        ),
        store,
        fetch=_fiba_fetch,
        allow_writes=True,
    )
    monkeypatch.setattr(
        fiba,
        "fetch_game_result",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(ValueError("FIBA detail unavailable")),
    )

    report = pipeline._backfill_fiba_euroleague_w_history()

    assert report["status"] == "provider_error_cursor_retained"
    assert report["events_acquired"] == 0
    assert report["quarantined_rows"] == 1
    assert store.get_provider_health("fiba-elw-history", League.EUROLEAGUE_W) is None
    store.close()


def test_fiba_elw_backfill_rechecks_a_late_gdap_score_correction(tmp_path):
    store = DuckDBStore(str(tmp_path / "warehouse.duckdb"), persist=True)
    original_game = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    store.upsert_game(fiba.fetch_game_result(original_game, _fiba_fetch))
    changed_schedule = deepcopy(SCHEDULE)
    changed_schedule[0]["teamAScore"] = 77
    changed_detail = deepcopy(DETAIL)
    changed_detail["game"]["teamAScore"] = 77
    changed_detail["playByPlay"]["items"]["OT1"]["scoreA"] = 77

    def changed_fetch(url: str):
        if "getgdapgamesbycompetitionid" in url:
            return deepcopy(changed_schedule)
        if "/games/135150-ZARAG-BLA" in url:
            return deepcopy(changed_detail)
        raise AssertionError(f"unexpected FIBA URL: {url}")

    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE_W",
            date_yyyymmdd="20261015",
            fixture_now_utc="2026-10-15T12:00:00Z",
            deep_history_event_budget=1,
        ),
        store,
        fetch=changed_fetch,
        allow_writes=True,
    )

    report = pipeline._backfill_fiba_euroleague_w_history()

    assert report["score_corrections_detected"] == 1
    assert report["score_corrections_written"] == 1
    corrected = store.list_games(league=League.EUROLEAGUE_W)[0]
    assert corrected.actual is not None and corrected.actual.home_final == 77
    store.close()


def test_pending_elw_final_aged_out_of_recheck_window_is_reported_as_unresolvable():
    store = InMemoryStore()
    old_pending = fiba.discover_games("2026-10-14", _fiba_fetch)[0]
    old_pending.status = GameStatus.SCHEDULED
    store.upsert_game(old_pending)
    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE_W",
            date_yyyymmdd="20261118",
            fixture_now_utc="2026-11-18T12:00:00Z",
            history_days=21,
        ),
        store,
        fetch=_fiba_fetch,
    )

    assert pipeline._recheck_pending_finals(set()) == []
    assert any(
        gap.required_field == "pending_final_expired" and "result_unresolvable" in gap.detail
        for gap in pipeline.gaps
    )


def test_pipeline_routes_elw_discovery_results_and_direct_fallback_without_mens_feed():
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE_W",
            date_yyyymmdd="20261014",
            fixture_now_utc="2026-10-15T12:00:00Z",
            history_days=0,
        ),
        store,
        fetch=_fiba_fetch,
    )

    pregame = pipeline.pregame()
    results = pipeline.results()

    assert pregame["league"] == "EUROLEAGUE_W"
    assert "policy" not in pregame["book_provider_status"]
    assert pregame["book_provider_status"]["betexplorer_direct"]["enabled"] is True
    assert any(gap["required_field"] == "fiba_elw_unassigned_schedule_slots" for gap in pregame["gaps"])
    assert pregame["rows"][0]["prediction_provider"] is None
    assert results["finals_written"] == 1
    final = store.list_games(league=League.EUROLEAGUE_W)[0]
    assert final.actual is not None and final.actual.reconciles()[0]
    assert not store.list_games(league=League.EUROLEAGUE)


def test_elw_forecast_policy_is_explicit_and_never_inherits_mens_price_lane():
    assert league_forecast_config(League.EUROLEAGUE_W).posted_line_model_weight is None
    service = MarketForecastService(InMemoryStore(), League.EUROLEAGUE_W)
    report = service.build_report(
        as_of_utc="2026-10-14T12:00:00Z",
        date_yyyymmdd="20261014",
        timezone_name="Europe/Paris",
    )
    assert report["capabilities"]["league_operator_policy"] == (
        "fiba_elw_empirical_history_with_betexplorer_direct_fallback"
    )
    assert report["accas"]["paper_lane"] is None
    assert report["accas"]["pricing_evidence_status"] == "priced_lane_policy_applies"
