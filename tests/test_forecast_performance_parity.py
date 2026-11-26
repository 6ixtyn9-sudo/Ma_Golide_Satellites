from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json
import tarfile

from forecasting.harness import (
    CalibrationConfig,
    IncrementalProfileBuilder,
    build_profiles,
)
from forecasting.model import ForecastConfig, HistoryModel
from forecasting.service import MarketForecastService
from forecasting.types import Forecast
from ingestion.identity import (
    DEFAULT_TEAM_ALIASES,
    canonical_team_key,
    canonical_team_name,
    parse_utc,
)
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    Prediction,
)
from ingestion.store import InMemoryStore
from scripts.daily import _prune_report_files
from scripts.release_archive import build_release_archives


def _make_sample_games(league: League, count: int, *, quarters: bool) -> list[Game]:
    base = datetime(2025, 11, 1, 18, 0, tzinfo=timezone.utc)
    teams = [
        "Real Madrid",
        "FC Barcelona",
        "Fenerbahçe Beko",
        "Žalgiris Kaunas",
        "Olympiacos Piraeus",
        "Panathinaikos AKTOR",
        "AS Monaco",
        "Anadolu Efes",
    ]
    games: list[Game] = []
    for i in range(count):
        home = teams[(i * 3) % len(teams)]
        away = teams[(i * 3 + 1 + (i % (len(teams) - 1))) % len(teams)]
        if away == home:
            away = teams[(i * 3 + 2) % len(teams)]
        iso = (base + timedelta(hours=6 * i)).strftime("%Y-%m-%dT%H:%M:%SZ")
        hs = 70 + ((i * 7) % 25)
        as_ = 68 + ((i * 11) % 23)
        if quarters:
            qh = [hs // 4, hs // 4 + 1, hs // 4 - 1, hs - 3 * (hs // 4)]
            qa = [as_ // 4 + (i % 2), as_ // 4, as_ // 4, as_ - 3 * (as_ // 4) - (i % 2)]
            actual = ActualResult(
                period_kind=PeriodKind.QUARTERS,
                home_quarters=qh,
                away_quarters=qa,
                home_final=hs,
                away_final=as_,
                source="euroleague",
            )
        else:
            h1h = hs // 2
            h1a = as_ // 2
            actual = ActualResult(
                period_kind=PeriodKind.HALVES,
                home_quarters=[h1h, hs - h1h],
                away_quarters=[h1a, as_ - h1a],
                home_final=hs,
                away_final=as_,
                source="espn",
            )
        games.append(
            Game(
                game_id=f"{league.value.lower()}-{i}",
                league=league,
                season="2025-26",
                home_team_id=home,
                away_team_id=away,
                home_name=home,
                away_name=away,
                scheduled_utc=iso,
                status=GameStatus.FINAL,
                actual=actual,
                evidence={"neutral_site": (i % 7 == 0)},
            )
        )
    return games


def test_canonical_team_name_cache_matches_explicit_table_and_custom_aliases():
    for canonical, variants in DEFAULT_TEAM_ALIASES.items():
        assert canonical_team_name(canonical) == canonical_team_name(
            canonical, aliases=DEFAULT_TEAM_ALIASES
        )
        for variant in variants:
            assert canonical_team_name(variant) == canonical_team_name(
                variant, aliases=DEFAULT_TEAM_ALIASES
            )
            assert canonical_team_key(variant) == canonical_team_key(canonical)
    assert canonical_team_name("  Fenerbahçe  Beko ") == canonical_team_name(
        "  Fenerbahçe  Beko ", aliases=DEFAULT_TEAM_ALIASES
    )
    assert canonical_team_name("Custom Club", aliases={"My Club": ["Custom Club"]}) == "My Club"
    assert parse_utc("2026-11-25T18:00:00Z") == parse_utc(
        datetime(2026, 11, 25, 18, 0, tzinfo=timezone.utc)
    )


def test_incremental_profile_builder_matches_build_profiles_exactly():
    cfg = CalibrationConfig(calibration_epoch_utc="2025-11-05T00:00:00Z", min_samples=5)
    tracker = IncrementalProfileBuilder(cfg)
    records: list[Forecast] = []
    base = datetime(2025, 11, 1, 12, 0, tzinfo=timezone.utc)
    outcomes = ["win", "loss", "win", "push", "loss", "win", None]
    for i in range(42):
        iso = (base + timedelta(hours=12 * i)).strftime("%Y-%m-%dT%H:%M:%SZ")
        row = Forecast(
            game_id=f"g-{i}",
            league="NCAA_M",
            scheduled_utc=iso,
            home_name="Team A",
            away_name="Team B",
            market="total" if i % 2 == 0 else "moneyline",
            period="full_game",
            selection="over" if i % 2 == 0 else "home",
            raw_probability=0.52 + (i % 7) * 0.03,
            alternative_probabilities={},
            source="history_v1" if i % 3 != 0 else "provider:espn",
            calibration_version="v2",
            team_history={"sufficient_for_accas": (i % 5 != 0)},
            outcome=outcomes[i % len(outcomes)],
        )
        records.append(row)
        tracker.add(row)
        expected = build_profiles(records, cfg)
        actual = tracker.profiles()
        assert set(actual.keys()) == set(expected.keys())
        for key, exp_prof in expected.items():
            assert actual[key].to_dict() == exp_prof.to_dict()
            assert tracker.get(key) == exp_prof


def test_indexed_residuals_and_highest_quarter_match_unindexed_history():
    import pytest

    games = _make_sample_games(League.EUROLEAGUE, 65, quarters=True)
    for cfg in (
        ForecastConfig(),
        ForecastConfig(min_residual_games=10, residual_window=10),
        ForecastConfig(training_window=35, residual_window=25),
    ):
        model = HistoryModel(cfg)
        for target_idx in (64, 30, 45):
            target = games[target_idx]
            view = model.available_history(
                games, scheduled_utc=target.scheduled_utc, as_of_utc=target.scheduled_utc
            )
            fresh_model = HistoryModel(cfg)
            fresh_view = fresh_model.available_history(
                games, scheduled_utc=target.scheduled_utc, as_of_utc=target.scheduled_utc
            )
            plain = list(view)
            for period in (Period.FULL_GAME, Period.Q1, Period.Q2, Period.Q3, Period.Q4):
                indexed_samples = model._residual_samples(view, period)
                assert indexed_samples == fresh_model._residual_samples(fresh_view, period)
                plain_samples = model._residual_samples(plain, period)
                assert len(indexed_samples) == len(plain_samples)
                for left_s, right_s in zip(indexed_samples, plain_samples):
                    assert left_s == pytest.approx(right_s, abs=1e-12)
            indexed_hq = [f.to_dict() for f in model.highest_quarter(target, view)]
            fresh_hq = [f.to_dict() for f in fresh_model.highest_quarter(target, fresh_view)]
            assert indexed_hq == fresh_hq
            plain_hq = [f.to_dict() for f in model.highest_quarter(target, plain)]
            assert len(indexed_hq) == len(plain_hq)
            for left, right in zip(indexed_hq, plain_hq):
                assert left["expected_total"] == pytest.approx(right["expected_total"], abs=1e-12)
                assert {k: v for k, v in left.items() if k != "expected_total"} == {
                    k: v for k, v in right.items() if k != "expected_total"
                }


def test_build_report_walk_forward_matches_non_incremental_reference():
    store = InMemoryStore()
    games = _make_sample_games(League.EUROLEAGUE, 45, quarters=True)
    for idx, game in enumerate(games):
        store.upsert_game(game)
        if idx % 2 == 0:
            obs = (parse_utc(game.scheduled_utc) - timedelta(hours=1)).strftime(
                "%Y-%m-%dT%H:%M:%SZ"
            )
            store.upsert_prediction(
                Prediction(
                    game_id=game.game_id,
                    provider="espn",
                    observed_utc=obs,
                    pre_tip=True,
                    prob_home=0.62,
                    prob_away=0.38,
                    expected_total=156.5,
                    expected_margin=3.5,
                ),
                frozen=True,
            )
            store.upsert_book_snapshot(
                BookSnapshot(
                    game_id=game.game_id,
                    provider="espn",
                    book="DraftKings",
                    market=Market.TOTAL,
                    period=Period.FULL_GAME,
                    observed_utc=obs,
                    pre_match=True,
                    line=155.5,
                    over_price_decimal=1.91,
                    under_price_decimal=1.91,
                ),
                frozen=True,
            )
    target_dt = parse_utc(games[-1].scheduled_utc) + timedelta(hours=24)
    target_iso = target_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    store.upsert_game(
        Game(
            game_id="euro-target",
            league=League.EUROLEAGUE,
            season="2025-26",
            home_team_id="Real Madrid",
            away_team_id="FC Barcelona",
            home_name="Real Madrid",
            away_name="FC Barcelona",
            scheduled_utc=target_iso,
            status=GameStatus.SCHEDULED,
            actual=None,
        )
    )
    as_of = (target_dt - timedelta(hours=2)).strftime("%Y-%m-%dT%H:%M:%SZ")
    report = MarketForecastService(
        store,
        League.EUROLEAGUE,
        calibration_config=CalibrationConfig(calibration_epoch_utc="2025-11-01T00:00:00Z"),
    ).build_report(date_yyyymmdd=target_dt.strftime("%Y%m%d"), as_of_utc=as_of)
    assert report["walk_forward"]["historical_finals"] == 45
    assert report["walk_forward"]["graded_forecasts"] > 0
    assert report["forecast_count"] > 0


def test_release_archive_excludes_transient_daily_reports_and_pruner_bounds_them(tmp_path):
    reports = tmp_path / "localdata" / "reports"
    reports.mkdir(parents=True)
    (reports / "daily_old.json").write_text(json.dumps({"finished_utc": "2026-11-20T00:00:00Z"}))
    (reports / "daily_recent.json").write_text(
        json.dumps({"finished_utc": "2026-11-25T10:00:00Z"})
    )
    (reports / "markets_mid.json").write_text(
        json.dumps({"generated_utc": "2026-11-20T00:00:00Z"})
    )

    pruned = _prune_report_files(
        tmp_path / "localdata",
        cutoff_utc="2026-10-26T00:00:00Z",
        daily_cutoff_utc="2026-11-23T12:00:00Z",
    )
    assert pruned == 1
    assert not (reports / "daily_old.json").exists()
    assert (reports / "daily_recent.json").exists()
    assert (reports / "markets_mid.json").exists()

    artifacts = tmp_path / "artifacts"
    state_dir = artifacts / "ncaa-success" / "localdata" / "ncaa_m"
    (state_dir / "csv").mkdir(parents=True)
    (state_dir / "reports").mkdir(parents=True)
    import gzip

    (state_dir / "csv" / "games.csv.gz").write_bytes(
        gzip.compress(b"game_id,payload_json\ng-1,{}\n", mtime=0)
    )
    (state_dir / "reports" / "daily_run_123.json").write_text('{"huge":"telemetry"}\n')
    (state_dir / "reports" / "markets_NCAA_M_2026-11-25.json").write_text("{}\n")

    repo = tmp_path / "repo"
    output = tmp_path / "out"
    build_release_archives(artifacts=artifacts, repo=repo, output=output)
    with tarfile.open(output / "magolide-state-ncaa_m.tar.gz", "r:gz") as bundle:
        names = bundle.getnames()
    assert "artifact/ncaa_m/reports/markets_NCAA_M_2026-11-25.json" in names
    assert not any("daily_run_123.json" in name for name in names)
