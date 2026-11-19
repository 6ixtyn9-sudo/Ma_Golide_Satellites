"""Contracts for the registry-based final basketball expansion."""

from __future__ import annotations

import pytest

from forecasting.harness import CalibrationConfig, apply_profiles
from forecasting.types import BookEvidence, Forecast
from ingestion.adaptive_gate import GateUncertain, LEAGUES, fetch_schedule
from ingestion.duckdb_store import DuckDBStore
from ingestion.league_registry import SPECS, league_spec, parse_league, price_evidence_policy
from ingestion.models import (
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
)
from ingestion.pipeline import ACQUISITION_TZ, DEFAULT_TZ, Pipeline, PipelineConfig
from ingestion.providers import theoddsapi
from ingestion.store import InMemoryStore


EXPANSION = {
    League.EUROCUP,
    League.BASKETBALL_CHAMPIONS_LEAGUE,
    League.NBA_G_LEAGUE,
    League.BASKETBALL_AFRICA_LEAGUE,
    League.AUSTRALIA_NBL,
    League.SPAIN_ACB,
    League.FRANCE_LNB_PRO_A,
    League.GERMANY_BBL,
    League.ITALY_LBA,
    League.TURKEY_BSL,
    League.GREECE_BASKET_LEAGUE,
    League.ABA_LEAGUE,
    League.ISRAEL_PREMIER_LEAGUE,
    League.JAPAN_B_LEAGUE,
    League.SOUTH_KOREA_KBL,
    League.BRAZIL_NBB,
    League.ARGENTINA_LIGA_NACIONAL,
}


def _game(league: League) -> Game:
    return Game(
        game_id=f"manual-{league.value.lower()}",
        league=league,
        season="2026",
        home_team_id="h",
        away_team_id="a",
        home_name="Home",
        away_name="Away",
        scheduled_utc="2026-11-19T18:00:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={"manual_official_import": "fixture-1"},
    )


def test_final_expansion_is_typed_and_registry_backed_everywhere():
    assert len(EXPANSION) == 17
    assert EXPANSION <= set(League)
    assert set(League) == {spec.league for spec in SPECS}
    for league in EXPANSION:
        spec = league_spec(league)
        assert spec.state_key
        assert DEFAULT_TZ[league.value] == spec.timezone
        assert ACQUISITION_TZ[league.value] == spec.acquisition_timezone
        assert parse_league(league.value) is league
        # No unverified endpoint is represented as a functional adapter.
        assert spec.schedule_source == "source_unavailable"
        assert spec.results_source == "source_unavailable"


def test_adaptive_gate_schedules_every_registry_lane_and_fails_open_for_unavailable_sources():
    gate_by_league = {config.league: config for config in LEAGUES}
    for spec in SPECS:
        gate_league = "EuroLeague" if spec.league is League.EUROLEAGUE else spec.league.value
        config = gate_by_league[gate_league]
        assert config.state_key == spec.state_key
        assert config.acquisition_timezone == spec.acquisition_timezone
    unsupported = gate_by_league[League.SPAIN_ACB.value]
    with pytest.raises(GateUncertain, match="no gate schedule source"):
        list(fetch_schedule(unsupported, "2026-11-19"))


def test_legacy_and_short_boundary_aliases_normalise_to_canonical_leagues():
    assert parse_league("EuroLeague") is League.EUROLEAGUE
    assert parse_league("nbl") is League.AUSTRALIA_NBL
    assert League.NBL is League.AUSTRALIA_NBL
    assert League.BCL is League.BASKETBALL_CHAMPIONS_LEAGUE
    assert parse_league("BCL") is League.BASKETBALL_CHAMPIONS_LEAGUE
    assert parse_league("Japan B.League") is League.JAPAN_B_LEAGUE
    with pytest.raises(ValueError, match="unknown league"):
        parse_league("invented-super-league")


def test_unavailable_adapter_reports_gap_but_preserves_manually_canonical_lifecycle():
    store = InMemoryStore()
    game = _game(League.SPAIN_ACB)
    store.upsert_game(game)
    pipeline = Pipeline(
        PipelineConfig(
            league="SPAIN_ACB",
            date_yyyymmdd="20261119",
            fixture_now_utc="2026-11-19T12:00:00Z",
        ),
        store,
        fetch=lambda _url: {},
    )

    report = pipeline.pregame()

    assert report["fixtures_found"] == 1
    assert any(
        gap["required_field"] == "official_schedule" and gap["reason"] == "unsupported"
        for gap in report["gaps"]
    )
    assert report["book_provider_status"]["policy"]["status"] == "not_configured_for_league"
    assert any(gap["required_field"] == "named_book_price" for gap in report["gaps"])

    results = pipeline.results()
    assert results["finals_written"] == 0
    assert any(gap["required_field"] == "official_results" for gap in results["gaps"])


def test_registry_price_policy_is_provider_specific_and_fail_closed():
    nba_espn = price_evidence_policy(League.NBA, "espn")
    assert nba_espn.price_provenance is PriceProvenance.NAMED_BOOK
    assert nba_espn.source_permission is SourcePermission.UNRESOLVED
    assert nba_espn.operational_validation is OperationalValidation.UNVERIFIED

    cba_isports = price_evidence_policy(League.CBA, "isportsapi")
    assert cba_isports.source_permission is SourcePermission.UNRESOLVED
    assert cba_isports.operational_validation is OperationalValidation.PROBATION

    nba_direct = price_evidence_policy(League.NBA, "theoddsapi")
    assert nba_direct.source_permission is SourcePermission.CLEARED
    assert nba_direct.operational_validation is OperationalValidation.VALIDATED

    unknown = price_evidence_policy(League.NBA, "future-unreviewed-provider")
    assert unknown.source_permission is SourcePermission.UNRESOLVED
    assert unknown.operational_validation is OperationalValidation.UNVERIFIED


def test_aggregate_price_is_visible_but_never_ticket_eligible():
    evidence = BookEvidence(
        provider="betexplorer",
        book="average",
        observed_utc="2026-11-19T12:00:00Z",
        source_modified_utc=None,
        pre_match=True,
        main_line=True,
        frozen=False,
        prices={"home": 1.8, "away": 2.1},
        price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
        game_id="g",
        market="moneyline",
        period="full_game",
    )
    forecast = Forecast(
        game_id="g",
        league="EUROLEAGUE",
        scheduled_utc="2026-11-19T18:00:00Z",
        home_name="Home",
        away_name="Away",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=0.6,
        alternative_probabilities={"home": 0.6, "away": 0.4},
        source="fixture",
        book=evidence,
    )

    calibrated = apply_profiles([forecast], {}, CalibrationConfig(min_samples=1))[0]
    assert calibrated.decimal_price == 1.8
    assert calibrated.book.ticket_price_eligible is False
    assert calibrated.eligibility["acca_eligible"] is False
    assert {
        "price_not_named_bookmaker_quote",
        "price_source_permission_unresolved",
        "price_operational_validation_incomplete",
    } <= set(calibrated.eligibility["reasons"])
    row = calibrated.to_dict()
    assert row["book"]["price_provenance"] == "aggregate_consensus"
    assert row["book"]["ticket_price_eligible"] is False


def test_no_price_provenance_rejects_a_fabricated_decimal_quote():
    with pytest.raises(ValueError, match="no_price evidence"):
        BookSnapshot(
            game_id="g",
            provider="fixture",
            book="unpriced",
            market=Market.MONEYLINE,
            period=Period.FULL_GAME,
            observed_utc="2026-11-19T12:00:00Z",
            home_price_decimal=1.8,
            price_provenance=PriceProvenance.NO_PRICE,
        )


def test_legacy_book_payload_without_axes_is_not_resurrected_as_ticket_eligible():
    from ingestion.store import _book_from_dict

    loaded = _book_from_dict(
        {
            "game_id": "legacy",
            "provider": "espn",
            "book": "DraftKings",
            "market": "moneyline",
            "period": "full_game",
            "observed_utc": "2026-11-19T12:00:00Z",
            "home_price_decimal": 1.8,
            "away_price_decimal": 2.1,
        }
    )
    assert loaded.price_provenance is PriceProvenance.NAMED_BOOK
    assert loaded.source_permission is SourcePermission.UNRESOLVED
    assert loaded.operational_validation is OperationalValidation.UNVERIFIED


def test_evidence_axes_round_trip_through_duckdb(tmp_path):
    snapshot = BookSnapshot(
        game_id="g",
        provider="fixture",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-19T12:00:00Z",
        home_price_decimal=1.8,
        away_price_decimal=2.1,
        price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
    )
    path = tmp_path / "warehouse.duckdb"
    with DuckDBStore(str(path), persist=True) as store:
        store.upsert_book_snapshot(snapshot)
    with DuckDBStore(str(path), persist=False) as store:
        loaded = store.get_book_snapshots("g")[0]
        assert loaded.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
        assert loaded.source_permission is SourcePermission.UNRESOLVED
        assert loaded.operational_validation is OperationalValidation.PROBATION


def test_nbl_odds_api_mapping_is_probationary_not_a_validated_ticket_lane(monkeypatch):
    monkeypatch.setenv("THE_ODDS_API_KEY", "fixture-key")
    event = {
        "id": "nbl-1",
        "home_team": "Home",
        "away_team": "Away",
        "commence_time": "2026-11-19T18:00:00Z",
        "bookmakers": [
            {
                "key": "fixturebook",
                "markets": [
                    {
                        "key": "h2h",
                        "outcomes": [
                            {"name": "Home", "price": 1.8},
                            {"name": "Away", "price": 2.1},
                        ],
                    }
                ],
            }
        ],
    }
    urls: list[str] = []

    def fetch(url: str):
        urls.append(url)
        return event if "/events/nbl-1/odds" in url else [event]

    snapshots = theoddsapi.fetch_odds(
        "AUSTRALIA_NBL",
        "20261120",  # Sydney local date for 18:00Z during daylight saving
        lambda _provider: fetch,
        now_utc="2026-11-19T12:00:00Z",
        tz_name="Australia/Sydney",
    )
    assert snapshots
    assert "basketball_nbl" in urls[0]
    assert all(snapshot.price_provenance is PriceProvenance.NAMED_BOOK for snapshot in snapshots)
    assert all(snapshot.operational_validation is OperationalValidation.PROBATION for snapshot in snapshots)
