"""Hermetic official-FIBA EuroLeague Women routing and result regressions."""

from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from forecasting.model import league_forecast_config
from forecasting.service import MarketForecastService
from ingestion.cli import _warehouse
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
    assert games[0].status is GameStatus.FINAL


def test_live_fiba_client_refreshes_page_shipped_key_and_keeps_it_out_of_schedule_url():
    class FakeHttp:
        def __init__(self):
            self.calls = []

        def get_text(self, _url, **_kwargs):
            return 'NEXT_CLIENT_APIM_SUBSCRIPTION_KEY\\":\\"0123456789abcdef0123456789abcdef\\"'

        def get_json(self, url, *, headers):
            self.calls.append((url, headers))
            return deepcopy(SCHEDULE)

    http = FakeHttp()
    payload = fiba.FibaClient(http).get_json(fiba.schedule_url())

    assert payload == SCHEDULE
    url, headers = http.calls[0]
    assert "subscription-key" not in url
    assert headers["Ocp-Apim-Subscription-Key"] == "0123456789abcdef0123456789abcdef"


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


def test_fiba_elw_rsc_extractor_keeps_official_detail_objects_for_result_parser():
    # FIBA Flight streams can expose a UI timeout configuration before the
    # actual PBP object; extraction must keep scanning for ``items``.
    flight = (
        '0:{"playByPlay":{"timeout":5}}\n'
        + "1:"
        + json.dumps({"game": DETAIL["game"], "playByPlay": DETAIL["playByPlay"]})
    )

    extracted = fiba.extract_game_page(flight, "135150")

    assert extracted["game"]["competition"]["competitionId"] == 209127
    assert extracted["playByPlay"]["items"]["OT1"]["scoreA"] == 76


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


def test_pipeline_routes_elw_discovery_results_and_unpriced_policy_without_mens_feed():
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
    assert pregame["book_provider_status"]["policy"]["status"] == "not_configured_for_league"
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
        "fiba_elw_empirical_history_only_no_validated_named_book_lane"
    )
    assert report["accas"]["paper_lane"] is None
