"""Contracts for the registry-based final basketball expansion."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from forecasting.harness import CalibrationConfig, apply_profiles
from forecasting.service import MarketForecastService
from forecasting.types import BookEvidence, Forecast
from ingestion import adaptive_gate
from ingestion.adaptive_gate import GateUncertain, LEAGUES, fetch_schedule
from ingestion.duckdb_store import DuckDBStore
from ingestion.league_registry import SPECS, league_spec, parse_league, price_evidence_policy
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    OperationalValidation,
    Period,
    PeriodKind,
    PriceProvenance,
    SourcePermission,
)
from ingestion.pipeline import (
    ACQUISITION_TZ,
    DEFAULT_TZ,
    ESPN_LEAGUES,
    Pipeline,
    PipelineConfig,
)
from ingestion.providers import espn, theoddsapi
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
        if league is League.NBA_G_LEAGUE:
            # Captured ESPN NBA G League scoreboards and summaries establish
            # canonical schedule/result handling, but not an external BPI or
            # named-book price contract.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "espn",
                "espn",
                "espn",
            )
        elif league is League.ABA_LEAGUE:
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "aba_liga",
                "aba_liga",
                "aba_liga",
            )
        elif league is League.ITALY_LBA:
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "lba",
                "lba",
                "lba",
            )
        elif league is League.ISRAEL_PREMIER_LEAGUE:
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "israel_ibl",
                "israel_ibl",
                "israel_ibl",
            )
        elif league is League.SPAIN_ACB:
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "acb",
                "acb",
                "acb",
            )
        elif league is League.AUSTRALIA_NBL:
            # ESPN's own NBL competition route (basketball/nbl, league UID 55),
            # verified live on 2026-11-23; pricing stays on the probationary
            # The Odds API lane and is unchanged by the schedule route.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "espn",
                "espn",
                "espn",
            )
        elif league is League.SOUTH_KOREA_KBL:
            # The KBL lane reads the league's own date-window data API; the
            # route is documented by a third party, so the adapter records the
            # URL/bytes/sha256 it actually saw on every run.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "kbl",
                "kbl",
                "kbl",
            )
        elif league is League.TURKEY_BSL:
            # No longer a live lane. The TBF daily-matches route answers 200 but
            # ignores match_date (byte-identical bodies for 2026-11-21/22/23)
            # and carries only TKBL rows; the vendored 2026-10-11 body has TBL,
            # KBSL and TKBL and no BSL row either. An external probe on
            # 2026-11-23 recorded this in lane-probe-2026-11-23.md. The adapter
            # and its parser tests stay, but the registry must not claim a
            # source that cannot yield this league's fixtures.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "source_unavailable",
                "source_unavailable",
                "source_unavailable",
            )
        elif league is League.JAPAN_B_LEAGUE:
            # The B.League lane reads the site's server-rendered slate plus a
            # keyless S3 snapshot; both were captured out-of-sandbox and are
            # vendored, so the parser is checked against served bytes.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "bleague",
                "bleague",
                "bleague",
            )
        elif league is League.BASKETBALL_CHAMPIONS_LEAGUE:
            # BCL rides the FIBA GDAP contract module the EuroLeague Women lane
            # already validates; only the competition id and identity namespace
            # are its own.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "bcl",
                "bcl",
                "bcl",
            )
        elif league is League.BASKETBALL_AFRICA_LEAGUE:
            # BAL is served by the NBA site stack (__NEXT_DATA__), not FIBA GDAP.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "bal",
                "bal",
                "bal",
            )
        elif league is League.EUROCUP:
            # EuroCup rides the validated EuroLeague Basketball adapter with
            # competition code U; it is a referenced route contract, not a
            # first-party capture, and pricing stays on the direct fallback.
            assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
                "euroleague",
                "euroleague",
                "euroleague",
            )
        else:
            assert spec.schedule_source == "source_unavailable"
            assert spec.results_source == "source_unavailable"


def test_adaptive_gate_schedules_every_registry_lane_and_fails_open_for_unavailable_sources():
    gate_by_league = {config.league: config for config in LEAGUES}
    for spec in SPECS:
        gate_league = "EuroLeague" if spec.league is League.EUROLEAGUE else spec.league.value
        config = gate_by_league[gate_league]
        assert config.state_key == spec.state_key
        assert config.acquisition_timezone == spec.acquisition_timezone
    unsupported = gate_by_league[League.FRANCE_LNB_PRO_A.value]
    with pytest.raises(GateUncertain, match="no gate schedule source"):
        list(fetch_schedule(unsupported, "2026-11-19"))
    aba = gate_by_league[League.ABA_LEAGUE.value]
    with pytest.raises(GateUncertain, match="ABA League has no independent date-bounded gate source"):
        list(fetch_schedule(aba, "2026-11-19"))


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
    game = _game(League.FRANCE_LNB_PRO_A)
    store.upsert_game(game)
    pipeline = Pipeline(
        PipelineConfig(
            league="FRANCE_LNB_PRO_A",
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
    assert "policy" not in report["book_provider_status"]
    assert report["book_provider_status"]["betexplorer_direct"]["enabled"] is True

    results = pipeline.results()
    assert results["finals_written"] == 0
    assert any(gap["required_field"] == "official_results" for gap in results["gaps"])


def test_g_league_espn_pipeline_routes_canonical_results_but_never_fabricates_bpi(
    fixture_fetch,
):
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(
            league="NBA_G_LEAGUE",
            date_yyyymmdd="20260410",
            fixture_now_utc="2026-04-12T12:00:00Z",
            history_days=0,
            deep_history_event_budget=1,
        ),
        store,
        fetch=fixture_fetch({"dates=20260410": "espn/scoreboard_gleague_20260410.json"}),
    )

    backfill = pipeline.backfill()
    pregame = pipeline.pregame()
    results = pipeline.results()
    smoke = pipeline.smoke()

    assert backfill["backfill"]["status"] == "disabled"
    assert backfill["backfill"]["deep_history"]["status"] == "not_validated_for_league"
    assert pregame["fixtures_found"] == 1
    assert not any(gap["required_field"] == "official_schedule" for gap in pregame["gaps"])
    # This lane has no validated external pre-tip predictor; it must not
    # manufacture an ESPN BPI provider failure or a substitute probability.
    assert not any(gap["required_field"] == "espn_bpi" for gap in pregame["gaps"])
    assert pregame["rows"][0]["prediction"] is None
    assert results["finals_written"] == 1
    game = store.get_game("401867657")
    assert game.league is League.NBA_G_LEAGUE
    assert game.actual is not None and game.actual.reconciles()[0]
    assert smoke["predictions"] == {
        "status": "no_external_prediction_provider_validated",
        "policy": "empirical_history_only_after_reconciled_history_threshold",
    }
    assert smoke["bpi_complete"] == {"complete": 0, "total": 0}


def test_g_league_emits_empirical_forecast_cards_only_after_reconciled_history():
    """No provider probability or price is needed for an unpriced forecast card."""

    store = InMemoryStore()
    first_tip = datetime(2026, 9, 1, 20, tzinfo=timezone.utc)
    # The model requires 20 baseline games and 20 independently generated
    # residual samples, so 40 reconciled finals demonstrate both evidence
    # gates rather than weakening them for a new league.
    for index in range(40):
        home = 95 + index % 5
        away = 88 + index % 4
        tip = first_tip + timedelta(days=index)
        store.upsert_game(
            Game(
                game_id=f"gleague-final-{index}",
                league=League.NBA_G_LEAGUE,
                season="2027",
                home_team_id="23",
                away_team_id="9",
                home_name="Stockton Kings",
                away_name="Greensboro Swarm",
                scheduled_utc=tip.strftime("%Y-%m-%dT%H:%M:%SZ"),
                status=GameStatus.FINAL,
                provider_ids={"espn": f"historical-{index}"},
                actual=ActualResult(
                    period_kind=PeriodKind.QUARTERS,
                    home_quarters=[24.0, 23.0, 25.0, float(home - 72)],
                    away_quarters=[22.0, 21.0, 23.0, float(away - 66)],
                    home_final=float(home),
                    away_final=float(away),
                    source="espn",
                ),
            )
        )
    store.upsert_game(
        Game(
            game_id="gleague-upcoming",
            league=League.NBA_G_LEAGUE,
            season="2027",
            home_team_id="23",
            away_team_id="9",
            home_name="Stockton Kings",
            away_name="Greensboro Swarm",
            scheduled_utc="2026-11-19T20:00:00Z",
            status=GameStatus.SCHEDULED,
            provider_ids={"espn": "upcoming"},
        )
    )

    report = MarketForecastService(store, League.NBA_G_LEAGUE).build_report(
        as_of_utc="2026-11-19T12:00:00Z",
        date_yyyymmdd="20261119",
        timezone_name="America/New_York",
    )

    cards = [card for card in report["forecasts"] if card["game_id"] == "gleague-upcoming"]
    assert cards
    assert all(card["prediction"] is None for card in cards)
    assert all(card["model"] == "team_history_empirical" for card in cards)
    assert all(card["book"] is None for card in cards)
    assert report["capabilities"]["league_operator_policy"] == (
        "nba_g_league_empirical_history_with_betexplorer_direct_fallback"
    )
    assert report["accas"]["pricing_evidence_status"] == "priced_lane_policy_applies"
    assert report["accas"]["strict_result_count"] == 0
    readiness = report["empirical_lane_readiness"]
    assert readiness["status"] == "monitoring_only"
    assert readiness["purpose"] == "empirical_evidence_readiness_not_accuracy_or_value_claim"
    assert readiness["source_contract"] == "espn_date_bounded_schedule_and_reconciled_results_only"
    assert readiness["model_change"] == (
        "none_no_g_league_specific_override_without_walk_forward_evidence"
    )
    assert readiness["named_book_or_ticket_lane"] == (
        "betexplorer_direct_fallback_when_response_has_named_book"
    )
    assert readiness["reconciled_finals_as_of"] == 40
    assert readiness["minimum_reconciled_finals_for_first_empirical_forecast"] == 40
    assert readiness["current_empirical_forecasts"] == len(cards)
    assert readiness["current_empirical_forecasts_with_sufficient_team_history"] == len(cards)
    assert readiness["graded_prequential_empirical_forecasts"] == 0
    assert "no_graded_prequential_empirical_forecasts_yet" in readiness["blockers"]


def test_registry_price_policy_is_provider_specific_and_fail_closed():
    nba_espn = price_evidence_policy(League.NBA, "espn")
    assert nba_espn.price_provenance is PriceProvenance.NAMED_BOOK
    assert nba_espn.source_permission is SourcePermission.UNRESOLVED
    assert nba_espn.operational_validation is OperationalValidation.UNVERIFIED

    nba_direct = price_evidence_policy(League.NBA, "theoddsapi")
    assert nba_direct.source_permission is SourcePermission.CLEARED
    assert nba_direct.operational_validation is OperationalValidation.VALIDATED

    euroleague_aggregate = price_evidence_policy(League.EUROLEAGUE, "betexplorer")
    assert euroleague_aggregate.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert euroleague_aggregate.source_permission is SourcePermission.UNRESOLVED
    assert euroleague_aggregate.operational_validation is OperationalValidation.PROBATION
    assert "pick reporting only" in euroleague_aggregate.detail

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
    assert "price_not_named_bookmaker_quote" in calibrated.eligibility["reasons"]
    assert "price_source_permission_unresolved" not in calibrated.eligibility["reasons"]
    assert "price_operational_validation_incomplete" not in calibrated.eligibility["reasons"]
    row = calibrated.to_dict()
    assert row["book"]["price_provenance"] == "aggregate_consensus"
    assert row["book"]["ticket_price_eligible"] is False


def test_named_book_price_is_not_blocked_by_permission_or_validation_labels():
    evidence = BookEvidence(
        provider="espn",
        book="DraftKings",
        observed_utc="2026-11-19T12:00:00Z",
        source_modified_utc=None,
        pre_match=True,
        main_line=True,
        frozen=False,
        prices={"home": 1.8, "away": 2.1},
        price_provenance=PriceProvenance.NAMED_BOOK,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.UNVERIFIED,
        game_id="g",
        market="moneyline",
        period="full_game",
    )
    forecast = Forecast(
        game_id="g",
        league="NBA",
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

    assert calibrated.book.ticket_price_eligible is True
    assert "price_source_permission_unresolved" not in calibrated.eligibility["reasons"]
    assert "price_operational_validation_incomplete" not in calibrated.eligibility["reasons"]
    assert calibrated.eligibility["reasons"] == ["insufficient_walk_forward_samples"]


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


def test_australia_nbl_rides_the_espn_competition_route_not_a_country_slug():
    """NBL is ESPN's own competition route, not a ``basketball/xx.1`` guess.

    The core registry rejects ``basketball/esp.1``-style country slugs with
    "Invalid sport/league combination"; ``basketball/nbl`` resolves to league
    UID 55 and served three games for 2026-11-21. The identity guard must carry
    that same UID or every event would be quarantined (the NCAA_W ``50`` bug).
    """

    assert espn.LEAGUE_SLUGS[League.AUSTRALIA_NBL] == "nbl"
    assert espn.ESPN_LEAGUE_IDS[League.AUSTRALIA_NBL] == "55"
    assert League.AUSTRALIA_NBL in ESPN_LEAGUES

    url = espn._scoreboard_url("20261121", League.AUSTRALIA_NBL)
    assert "basketball/nbl/scoreboard" in url
    assert "dates=20261121" in url

    gate = adaptive_gate._espn_scoreboard_url(
        type("C", (), {"league": "AUSTRALIA_NBL"})(), "2026-11-21"
    )
    assert "basketball/nbl/scoreboard" in gate
    assert "dates=20261121" in gate
