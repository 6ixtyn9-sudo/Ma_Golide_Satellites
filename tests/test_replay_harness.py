from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List

import pytest

from forecasting.accas import build_accas
from forecasting.operator import (
    _compact_performance_report,
    _performance,
    publish_operator_reports,
)
from forecasting.replay import (
    ReplaySpec,
    check_parity,
    compare_variants,
    kelly_sweep,
    load_localdata_evidence,
    parse_variant_spec,
    replay,
)
from forecasting.types import BookEvidence, Forecast
from ingestion.models import (
    ActualResult,
    Game,
    GameStatus,
    League,
    PeriodKind,
    PriceProvenance,
)
from scripts.publish_localdata import _compact_ticket_receipt
from scripts.replay_harness import main as cli_main


def _make_forecast(
    *,
    league: League,
    game_id: str,
    home: str,
    away: str,
    scheduled_utc: str,
    market: str = "moneyline",
    selection: str = "home",
    line: float | None = None,
    prob: float = 0.70,
    price: float = 1.65,
    source: str = "team_history_empirical",
    brier: float = 0.18,
) -> Forecast:
    opp = {"home": "away", "away": "home", "over": "under", "under": "over"}[selection]
    return Forecast(
        game_id=game_id,
        league=league.value,
        scheduled_utc=scheduled_utc,
        home_name=home,
        away_name=away,
        market=market,
        period="full_game",
        selection=selection,
        raw_probability=prob,
        alternative_probabilities={selection: prob, opp: round(1.0 - prob, 6)},
        source=source,
        line=line,
        book=BookEvidence(
            provider="espn",
            book="DraftKings",
            observed_utc="2026-11-27T09:00:00Z",
            source_modified_utc=None,
            pre_match=True,
            main_line=True,
            frozen=True,
            prices={selection: price, opp: 2.15},
            price_provenance=PriceProvenance.NAMED_BOOK,
            game_id=game_id,
            market=market,
            period="full_game",
            line=line,
        ),
        calibrated_probability=prob,
        calibration={
            "status": "calibrated",
            "n": 80,
            "minimum_samples": 30,
            "sufficient_for_accas": True,
            "brier_score": brier,
        },
        eligibility={"acca_eligible": True, "reasons": [], "status": "eligible"},
    )


def _final_game(
    *,
    league: League,
    game_id: str,
    home: str,
    away: str,
    scheduled_utc: str,
    home_score: int,
    away_score: int,
) -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026-27",
        home_team_id=f"home-{game_id}",
        away_team_id=f"away-{game_id}",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled_utc,
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[home_score / 4] * 4,
            away_quarters=[away_score / 4] * 4,
            home_final=home_score,
            away_final=away_score,
        ),
    )


def _scheduled_game(
    *,
    league: League,
    game_id: str,
    home: str,
    away: str,
    scheduled_utc: str,
) -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026-27",
        home_team_id=f"home-{game_id}",
        away_team_id=f"away-{game_id}",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled_utc,
        status=GameStatus.SCHEDULED,
    )


class _DummyStore:
    def __init__(self, games: Dict[str, Game]) -> None:
        self._games = list(games.values())

    def list_games(self, *, league: League) -> List[Game]:
        return [g for g in self._games if g.league is league]


def _publish_lane_day(
    root: Path,
    *,
    league: League,
    target_date: str,
    generated_utc: str,
    forecasts: List[Forecast],
    games: Dict[str, Game],
    forbid_opposing: bool = False,
    extra_single_rows: List[Forecast] | None = None,
) -> Path:
    league_dir = root / league.value.lower()
    league_dir.mkdir(parents=True, exist_ok=True)
    paper_accas = build_accas(
        forecasts,
        legs=min(3, len({fc.game_id for fc in forecasts})),
        limit=5,
        allow_provisional=False,
        require_disjoint_legs=True,
        forbid_conflicting_game_sides=forbid_opposing,
    )
    single_pool = list(forecasts) + list(extra_single_rows or [])
    market_report = {
        "league": league.value,
        "generated_utc": generated_utc,
        "target_date": target_date.replace("-", ""),
        "forecasts": [fc.to_dict() for fc in single_pool],
        "accas": {
            "requested_legs": 3,
            "effective_paper_legs": min(3, len({fc.game_id for fc in forecasts})),
            "requested_limit": 5,
            "paper_lane": {"requires_calibrated_segment": True},
            "paper_ticket_leg_reuse_policy": (
                "no_reused_wager_selections_or_opposing_game_sides"
                if forbid_opposing
                else "no_reused_wager_selections"
            ),
            "paper_single_picks": [fc.to_dict() for fc in single_pool],
            "paper_results": paper_accas,
        },
    }
    publish_operator_reports(
        _DummyStore(games),
        league,
        market_report,
        league_dir,
        target_date=target_date,
        generated_utc=generated_utc,
        run_id=f"run_{league.value}_{target_date}",
        slate_timezone_name="America/New_York",
    )
    return league_dir


def test_build_accas_combined_odds_band_and_cross_league_game_id_isolation() -> None:
    # Two leagues deliberately share game_id "g1" — must not collide in cross-league pool.
    nba_g1 = _make_forecast(
        league=League.NBA,
        game_id="g1",
        home="Boston Celtics",
        away="Miami Heat",
        scheduled_utc="2026-11-27T23:00:00Z",
        prob=0.68,
        price=1.80,
    )
    ncaaw_g1 = _make_forecast(
        league=League.NCAA_W,
        game_id="g1",
        home="UConn Huskies",
        away="Duke Blue Devils",
        scheduled_utc="2026-11-27T23:30:00Z",
        prob=0.72,
        price=1.85,
    )
    euro_g2 = _make_forecast(
        league=League.EUROLEAGUE,
        game_id="g2",
        home="Real Madrid",
        away="Olympiacos",
        scheduled_utc="2026-11-27T19:00:00Z",
        prob=0.80,
        price=1.25,
    )

    # Without combined odds floor, 2-leg picks the highest-probability pair (euro_g2 @1.25 + ncaaw_g1 @1.85 = @2.3125).
    default_accas = build_accas([nba_g1, ncaaw_g1, euro_g2], legs=2, limit=3)
    assert default_accas
    assert default_accas[0]["combined_decimal_odds"] == pytest.approx(1.25 * 1.85, abs=1e-4)

    # With min_combined_odds=3.0, live build_accas skips @2.31 and picks NBA g1 (@1.80) + NCAA_W g1 (@1.85) = @3.33!
    targeted = build_accas(
        [nba_g1, ncaaw_g1, euro_g2],
        legs=2,
        limit=3,
        min_combined_odds=3.0,
        max_combined_odds=4.0,
    )
    assert len(targeted) == 1
    assert targeted[0]["combined_decimal_odds"] == pytest.approx(1.80 * 1.85, abs=1e-4)
    assert {leg["league"] for leg in targeted[0]["legs"]} == {"NBA", "NCAA_W"}


def test_receipt_captures_pool_compacts_cleanly_grades_counterfactuals_and_passes_parity(
    tmp_path: Path,
) -> None:
    target_date = "2026-11-27"
    generated_utc = "2026-11-27T10:15:00Z"  # 12:15 SAST (post-cutoff for NBA)
    # Four NBA games: 3 high-prob short prices (@1.30, @1.35, @1.40 -> live 3-leg @2.457)
    # plus 1 mid-price spread (@1.95) that a @3.0+ target-odds variant can use.
    f1 = _make_forecast(
        league=League.NBA,
        game_id="nba1",
        home="Celtics",
        away="Knicks",
        scheduled_utc="2026-11-27T23:00:00Z",
        prob=0.78,
        price=1.30,
    )
    f2 = _make_forecast(
        league=League.NBA,
        game_id="nba2",
        home="Nuggets",
        away="Suns",
        scheduled_utc="2026-11-27T23:30:00Z",
        prob=0.75,
        price=1.35,
    )
    f3 = _make_forecast(
        league=League.NBA,
        game_id="nba3",
        home="Thunder",
        away="Clippers",
        scheduled_utc="2026-11-28T00:00:00Z",
        prob=0.72,
        price=1.40,
    )
    f4 = _make_forecast(
        league=League.NBA,
        game_id="nba4",
        home="Lakers",
        away="Warriors",
        scheduled_utc="2026-11-28T01:00:00Z",
        market="total",
        selection="over",
        line=224.5,
        prob=0.62,
        price=1.95,
    )
    # Stale started-game forecast row (from prior slate) must be excluded from candidate_pool!
    started = _make_forecast(
        league=League.NBA,
        game_id="nba_old",
        home="Heat",
        away="Hawks",
        scheduled_utc="2026-11-26T22:00:00Z",
        prob=0.82,
        price=1.50,
    )

    sched_games = {
        fc.game_id: _scheduled_game(
            league=League.NBA,
            game_id=fc.game_id,
            home=fc.home_name,
            away=fc.away_name,
            scheduled_utc=fc.scheduled_utc,
        )
        for fc in (f1, f2, f3, f4)
    }

    league_dir = _publish_lane_day(
        tmp_path,
        league=League.NBA,
        target_date=target_date,
        generated_utc=generated_utc,
        forecasts=[f1, f2, f3, f4],
        extra_single_rows=[started],
        games=sched_games,
        forbid_opposing=True,
    )

    receipt_path = league_dir / f"auto_tickets_{target_date}_decision.json"
    raw_receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    assert "candidate_pool" in raw_receipt
    assert "selection_policy" in raw_receipt
    pool_gids = {row["game_id"] for row in raw_receipt["candidate_pool"]}
    assert pool_gids == {"nba1", "nba2", "nba3", "nba4"}  # started game excluded!

    # Simulate Git publication compaction and verify idempotency + parity survival.
    compacted = _compact_ticket_receipt(raw_receipt)
    assert _compact_ticket_receipt(compacted) == compacted
    receipt_path.write_text(json.dumps(compacted, indent=2) + "\n", encoding="utf-8")

    # Zero settled days before finals arrive -> parity still runs cleanly, replay reports 0 settled days.
    ev_pending = load_localdata_evidence(tmp_path)
    parity = check_parity(ev_pending)
    assert parity["checked"] == 1
    assert parity["matched"] == 1
    assert parity["mismatched"] == 0
    res_pending = replay(ev_pending, ReplaySpec())
    assert res_pending["settled_days"] == 0
    assert res_pending["pending_days"] == 1

    # Now settle all 4 games: nba1 W, nba2 W, nba3 L, nba4 (over 224.5 -> 120+115=235) W!
    final_games = {
        "nba1": _final_game(
            league=League.NBA, game_id="nba1", home="Celtics", away="Knicks",
            scheduled_utc=f1.scheduled_utc, home_score=112, away_score=104,
        ),
        "nba2": _final_game(
            league=League.NBA, game_id="nba2", home="Nuggets", away="Suns",
            scheduled_utc=f2.scheduled_utc, home_score=118, away_score=109,
        ),
        "nba3": _final_game(
            league=League.NBA, game_id="nba3", home="Thunder", away="Clippers",
            scheduled_utc=f3.scheduled_utc, home_score=99, away_score=105,
        ),
        "nba4": _final_game(
            league=League.NBA, game_id="nba4", home="Lakers", away="Warriors",
            scheduled_utc=f4.scheduled_utc, home_score=120, away_score=115,
        ),
    }
    perf = _compact_performance_report(_performance([compacted], final_games))
    (league_dir / "auto_tickets_performance.json").write_text(
        json.dumps(perf, indent=2) + "\n", encoding="utf-8"
    )
    # Candidate pool outcomes must include nba4 even though live 3-leg acca used nba1+nba2+nba3.
    day_outcomes = perf["history"][0]["candidate_pool_outcomes"]
    assert day_outcomes[f4.forecast_id] == "win"
    assert day_outcomes[f3.forecast_id] == "loss"

    ev_settled = load_localdata_evidence(tmp_path)
    # Live 3-leg (nba1+nba2+nba3) lost because nba3 lost.
    live_res = replay(ev_settled, ReplaySpec())
    assert live_res["settled_days"] == 1
    assert live_res["losses"] == 1

    # Counterfactual: 3-leg targeting @3.0+ forces nba4 (@1.95) into the slip alongside nba1 (@1.30) + nba2 (@1.35) = @3.4223 -> WIN!
    target_spec = parse_variant_spec("legs=3,min_odds=3.0,max_odds=4.5,stake_mode=daily_pct,stake_frac=0.25,skim=0.10")
    cf_res = replay(ev_settled, target_spec)
    assert cf_res["settled_days"] == 1
    assert cf_res["wins"] == 1
    assert cf_res["losses"] == 0
    assert cf_res["total_wealth"] > 100.0
    assert cf_res["skimmed_vault"] > 0.0


def test_cross_league_replay_ab_noop_and_cli_smoke(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    target_date = "2026-11-28"
    # Publish NBA (12:15 SAST) AND NCAA_W (16:15 SAST) on the same calendar day!
    nba_f1 = _make_forecast(
        league=League.NBA, game_id="n1", home="Bucks", away="Bulls",
        scheduled_utc="2026-11-28T23:00:00Z", prob=0.72, price=1.80,
    )
    nba_f2 = _make_forecast(
        league=League.NBA, game_id="n2", home="Mavs", away="Spurs",
        scheduled_utc="2026-11-28T23:30:00Z", prob=0.68, price=1.55,
    )
    ncaaw_f1 = _make_forecast(
        league=League.NCAA_W, game_id="w1", home="South Carolina", away="LSU",
        scheduled_utc="2026-11-28T22:00:00Z", prob=0.74, price=1.85,
    )
    ncaaw_f2 = _make_forecast(
        league=League.NCAA_W, game_id="w2", home="UCLA", away="USC",
        scheduled_utc="2026-11-28T23:00:00Z", prob=0.66, price=1.40,
    )
    nba_sched = {
        fc.game_id: _scheduled_game(league=League.NBA, game_id=fc.game_id, home=fc.home_name, away=fc.away_name, scheduled_utc=fc.scheduled_utc)
        for fc in (nba_f1, nba_f2)
    }
    ncaaw_sched = {
        fc.game_id: _scheduled_game(league=League.NCAA_W, game_id=fc.game_id, home=fc.home_name, away=fc.away_name, scheduled_utc=fc.scheduled_utc)
        for fc in (ncaaw_f1, ncaaw_f2)
    }
    nba_finals = {
        "n1": _final_game(league=League.NBA, game_id="n1", home="Bucks", away="Bulls", scheduled_utc=nba_f1.scheduled_utc, home_score=115, away_score=100),
        "n2": _final_game(league=League.NBA, game_id="n2", home="Mavs", away="Spurs", scheduled_utc=nba_f2.scheduled_utc, home_score=100, away_score=110),
    }
    ncaaw_finals = {
        "w1": _final_game(league=League.NCAA_W, game_id="w1", home="South Carolina", away="LSU", scheduled_utc=ncaaw_f1.scheduled_utc, home_score=82, away_score=68),
        "w2": _final_game(league=League.NCAA_W, game_id="w2", home="UCLA", away="USC", scheduled_utc=ncaaw_f2.scheduled_utc, home_score=65, away_score=72),
    }

    nba_dir = _publish_lane_day(
        tmp_path, league=League.NBA, target_date=target_date,
        generated_utc="2026-11-28T10:15:00Z", forecasts=[nba_f1, nba_f2],
        games=nba_sched, forbid_opposing=True,
    )
    ncaaw_dir = _publish_lane_day(
        tmp_path, league=League.NCAA_W, target_date=target_date,
        generated_utc="2026-11-28T14:15:00Z", forecasts=[ncaaw_f1, ncaaw_f2],
        games=ncaaw_sched, forbid_opposing=True,
    )
    for d_dir, finals in ((nba_dir, nba_finals), (ncaaw_dir, ncaaw_finals)):
        rcpt = json.loads((d_dir / f"auto_tickets_{target_date}_decision.json").read_text(encoding="utf-8"))
        perf = _compact_performance_report(_performance([rcpt], finals))
        (d_dir / "auto_tickets_performance.json").write_text(json.dumps(perf, indent=2) + "\n", encoding="utf-8")

    days = load_localdata_evidence(tmp_path)
    assert len(days) == 2
    assert {d.league for d in days} == {"NBA", "NCAA_W"}

    # Cross-league 2-leg @3.0+ pairs NBA n1 (@1.80, W) + NCAA_W w1 (@1.85, W) = @3.33 -> WIN!
    # Whereas per-league 2-leg had to pair n1+n2 (L) and w1+w2 (L)!
    cross_spec = parse_variant_spec("legs=2,limit=1,min_odds=3.0,max_odds=4.0,stake_mode=daily_pct,stake_frac=0.25")
    cross_res = replay(days, cross_spec, mode="cross-league")
    assert cross_res["wins"] == 1
    assert cross_res["losses"] == 0
    assert cross_res["total_wealth"] > 150.0

    # Paired A/B NO-OP check refuses bootstrap when identical specs are compared.
    noop = compare_variants(days, cross_spec, cross_spec, mode="cross-league")
    assert noop["status"] == "no_op"

    # Sizing-only A/B runs bootstrap cleanly.
    sizing_cmp = compare_variants(
        days,
        parse_variant_spec("legs=2,limit=1,min_odds=3.0,max_odds=4.0,stake_mode=flat"),
        cross_spec,
        mode="cross-league",
        bootstraps=100,
    )
    assert sizing_cmp["status"] == "sizing_only"
    assert sizing_cmp["delta_mean_log_growth"] > 0

    # Kelly sweep runs across daily % curve.
    sweep = kelly_sweep(days, parse_variant_spec("legs=2,limit=1,min_odds=3.0,max_odds=4.0"), mode="cross-league")
    assert len(sweep) == 8
    assert sweep[4]["daily_pct"] == 0.25

    # CLI smoke: --parity, --kelly, --slots, --since future date (0 days).
    assert cli_main(["--localdata-root", str(tmp_path), "--parity"]) == 0
    assert cli_main(["--localdata-root", str(tmp_path), "--mode", "cross-league", "--variant", "legs=2,limit=1,min_odds=3.0", "--kelly", "--slots"]) == 0
    assert cli_main(["--localdata-root", str(tmp_path), "--since", "2026-12-01"]) == 0
    out = capsys.readouterr().out
    assert "0 settled replayable days" in out


def test_legacy_receipt_without_pool_is_excluded_from_selection_replays(tmp_path: Path) -> None:
    league_dir = tmp_path / "euroleague"
    league_dir.mkdir()
    legacy_receipt = {
        "schema_version": "magolide.paper_ticket_receipt",
        "date": "2026-11-20",
        "league": "EUROLEAGUE",
        "decision_id": "legacy_1",
        "status": "frozen",
        "frozen_at": "2026-11-20T10:00:00Z",
        "accas": [
            {
                "ticket_id": "t1",
                "ticket_type": "acca",
                "leg_count": 2,
                "combined_decimal_odds": 3.24,
                "estimated_joint_win_probability": 0.45,
                "legs": [
                    {"forecast_id": "f1", "book": {"selected_decimal_price": 1.8}},
                    {"forecast_id": "f2", "book": {"selected_decimal_price": 1.8}},
                ],
            }
        ],
        "singles": [],
    }
    (league_dir / "auto_tickets_2026-11-20_decision.json").write_text(json.dumps(legacy_receipt))
    (league_dir / "auto_tickets_performance.json").write_text(
        json.dumps(
            {
                "history": [
                    {
                        "date": "2026-11-20",
                        "decision_id": "legacy_1",
                        "positions": [
                            {
                                "legs": [
                                    {"forecast_id": "f1", "outcome": "win"},
                                    {"forecast_id": "f2", "outcome": "win"},
                                ]
                            }
                        ],
                    }
                ]
            }
        )
    )
    days = load_localdata_evidence(tmp_path)
    assert len(days) == 1
    assert not days[0].has_candidate_pool

    # Default live-faithful mode refuses to pretend a legacy receipt is a full pool.
    faithful = replay(days, ReplaySpec())
    assert faithful["settled_days"] == 0
    assert faithful["skipped_legacy_no_pool"] == 1

    # Explicit opt-in sizing-only mode replays staking on the frozen published tickets.
    sizing = replay(
        days,
        parse_variant_spec("stake_mode=daily_pct,stake_frac=0.25"),
        allow_legacy_sizing_only=True,
    )
    assert sizing["settled_days"] == 1
    assert sizing["legacy_sizing_days"] == 1
    assert sizing["wins"] == 1
    assert sizing["total_wealth"] == pytest.approx(100.0 * (0.75 + 0.25 * 3.24), rel=1e-4)


def test_derived_spread_survives_receipt_compaction_and_pool_grading() -> None:
    row = {
        "forecast_id": "ds_leg_1",
        "pick_id": "ds_leg_1",
        "game_id": "w_ds_1",
        "league": "NCAA_W",
        "date": "2026-11-28",
        "scheduled_utc": "2026-11-28T22:00:00Z",
        "home": "UConn",
        "away": "Duke",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": -5.5,
        "raw_probability": 0.72,
        "calibrated_probability": 0.70,
        "model": "provider:espn",
        "book": {
            "provider": "espn",
            "book": "DraftKings",
            "observed_utc": "2026-11-28T13:00:00Z",
            "pre_match": True,
            "main_line": True,
            "frozen": True,
            "selected_decimal_price": 1.91,
        },
        "derived_spread": {
            "contract": "unpriced_moneyline_spread",
            "selected_spread_line": -5.5,
        },
    }
    receipt = {
        "schema_version": "magolide.paper_ticket_receipt",
        "date": "2026-11-28",
        "league": "NCAA_W",
        "decision_id": "ds_run_1",
        "status": "frozen",
        "frozen_at": "2026-11-28T14:05:00Z",
        "accas": [],
        "singles": [
            {
                "ticket_id": "ds_leg_1",
                "ticket_type": "single",
                "publication_index": 1,
                "leg_count": 1,
                "combined_decimal_odds": 1.91,
                "stake_units": 1.0,
                "estimated_joint_win_probability": 0.70,
                "legs": [row],
            }
        ],
        "candidate_pool": [row],
    }
    compacted = _compact_ticket_receipt(receipt)
    assert compacted["candidate_pool"][0]["derived_spread"]["selected_spread_line"] == -5.5
    assert compacted["singles"][0]["legs"][0]["derived_spread"]["contract"] == "unpriced_moneyline_spread"

    # Home wins by 8 (> 5.5) -> derived spread cover WIN; home wins by 3 (< 5.5) -> LOSS!
    win_game = _final_game(
        league=League.NCAA_W, game_id="w_ds_1", home="UConn", away="Duke",
        scheduled_utc="2026-11-28T22:00:00Z", home_score=78, away_score=70,
    )
    loss_game = _final_game(
        league=League.NCAA_W, game_id="w_ds_1", home="UConn", away="Duke",
        scheduled_utc="2026-11-28T22:00:00Z", home_score=73, away_score=70,
    )
    perf_win = _compact_performance_report(_performance([compacted], {"w_ds_1": win_game}))
    perf_loss = _compact_performance_report(_performance([compacted], {"w_ds_1": loss_game}))
    assert perf_win["history"][0]["candidate_pool_outcomes"] == {"ds_leg_1": "win"}
    assert perf_loss["history"][0]["candidate_pool_outcomes"] == {"ds_leg_1": "loss"}


def test_replay_parity_and_ab_across_legacy_blend_and_kelly_option_1a(tmp_path: Path) -> None:
    """Both legacy probability_price_blend receipts AND new Option 1+A Kelly
    receipts must pass check_parity() 100%, and replay_harness A/B can compare
    the two policies over identical frozen pools."""
    from forecasting.accas import (
        EDGE_FLAG_THRESHOLD,
        MAX_CLAIMED_EDGE,
        PAPER_ACCA_LIMIT,
        PAPER_MAX_ACCA_LEGS,
        PAPER_MAX_COMBINED_ODDS,
        PAPER_MIN_ACCA_LEGS,
        PAPER_MIN_COMBINED_ODDS,
        RANKING_METRIC_KELLY,
    )
    from forecasting.operator import (
        _grade_receipt,
        _new_receipt,
        _render_ticket_card,
        _selection_policy_snapshot,
    )

    sched = "2026-11-28T20:00:00Z"
    pool_rows = [
        _make_forecast(league=League.NCAA_M, game_id="m1", home="Duke", away="UNC", scheduled_utc=sched, prob=0.78, price=1.40),
        _make_forecast(league=League.NCAA_M, game_id="m2", home="Kansas", away="Baylor", scheduled_utc=sched, prob=0.72, price=1.50),
        _make_forecast(league=League.NCAA_M, game_id="m3", home="UConn", away="Creighton", scheduled_utc=sched, prob=0.61, price=1.95),  # +19.0% EV -> flagged!
        _make_forecast(league=League.NCAA_M, game_id="m4", home="Gonzaga", away="SMC", scheduled_utc=sched, prob=0.56, price=1.92),
        _make_forecast(league=League.NCAA_M, game_id="m5", home="Houston", away="Texas", scheduled_utc=sched, prob=0.80, price=1.95),  # +56.0% EV -> hard-capped!
    ]
    # 1. Freeze a NEW Option 1+A Kelly receipt on 2026-11-28.
    kelly_accas = build_accas(
        pool_rows,
        legs=PAPER_MAX_ACCA_LEGS,
        min_legs=PAPER_MIN_ACCA_LEGS,
        max_legs=PAPER_MAX_ACCA_LEGS,
        limit=PAPER_ACCA_LIMIT,
        require_disjoint_legs=True,
        forbid_conflicting_game_sides=True,
        forbid_alt_line_reuse=True,
        min_combined_odds=PAPER_MIN_COMBINED_ODDS,
        max_combined_odds=PAPER_MAX_COMBINED_ODDS,
        ranking_metric=RANKING_METRIC_KELLY,
        max_claimed_edge=MAX_CLAIMED_EDGE,
        edge_flag_threshold=EDGE_FLAG_THRESHOLD,
    )
    assert len(kelly_accas) == 1
    assert kelly_accas[0]["leg_count"] == 3
    assert kelly_accas[0]["ranking_metric"] == RANKING_METRIC_KELLY
    assert kelly_accas[0]["edge_implausible"] is True
    assert all(leg["game_id"] != "m5" for leg in kelly_accas[0]["legs"])

    sel_policy = _selection_policy_snapshot(
        {
            "requested_legs": PAPER_MAX_ACCA_LEGS,
            "effective_paper_legs": PAPER_MAX_ACCA_LEGS,
            "requested_limit": PAPER_ACCA_LIMIT,
            "paper_lane": {"requires_calibrated_segment": True},
            "paper_ticket_leg_reuse_policy": "no_reused_wager_selections_or_opposing_game_sides",
            "paper_ticket_policy": {
                "requested_legs": PAPER_MAX_ACCA_LEGS,
                "min_legs": PAPER_MIN_ACCA_LEGS,
                "max_legs": PAPER_MAX_ACCA_LEGS,
                "effective_paper_legs": PAPER_MAX_ACCA_LEGS,
                "limit": PAPER_ACCA_LIMIT,
                "ranking_metric": RANKING_METRIC_KELLY,
                "min_combined_odds": PAPER_MIN_COMBINED_ODDS,
                "max_combined_odds": PAPER_MAX_COMBINED_ODDS,
                "forbid_conflicting_game_sides": True,
                "forbid_alt_line_reuse": True,
                "max_claimed_edge": MAX_CLAIMED_EDGE,
                "edge_flag_threshold": EDGE_FLAG_THRESHOLD,
            },
        },
        League.NCAA_M,
    )
    league_dir = tmp_path / "ncaa_m"
    league_dir.mkdir(parents=True, exist_ok=True)
    receipt = _compact_ticket_receipt(
        _new_receipt(
            league=League.NCAA_M,
            target_date="2026-11-28",
            captured_utc="2026-11-28T10:15:00Z",
            run_id="kelly_run_1",
            accas=kelly_accas,
            ticket_decision={"status": "frozen"},
            candidate_pool=[r.to_dict() for r in pool_rows],
            selection_policy=sel_policy,
        )
    )
    (league_dir / "auto_tickets_2026-11-28_decision.json").write_text(
        json.dumps(receipt, indent=2) + "\n", encoding="utf-8"
    )
    # Render card and confirm [EDGE>15% AUDIT] appears on the +19% EV leg
    # even after Git-durable receipt compaction!
    graded = _grade_receipt(receipt, {})
    perf = _performance([receipt], {})
    card_text = _render_ticket_card(graded, perf, timezone_name="Africa/Johannesburg")
    assert "[EDGE>15% AUDIT]" in card_text

    # 2. Freeze a LEGACY blend receipt on 2026-11-27 (without ranking_metric / min_legs / min_combined_odds).
    legacy_accas = build_accas(
        pool_rows[:4],
        legs=3,
        limit=5,
        require_disjoint_legs=True,
        forbid_conflicting_game_sides=False,
    )
    legacy_receipt = _compact_ticket_receipt(
        _new_receipt(
            league=League.NCAA_M,
            target_date="2026-11-27",
            captured_utc="2026-11-27T10:15:00Z",
            run_id="legacy_run_1",
            accas=legacy_accas,
            ticket_decision={"status": "frozen"},
            candidate_pool=[r.to_dict() for r in pool_rows[:4]],
            selection_policy={
                "requested_legs": 3,
                "effective_paper_legs": 3,
                "limit": 5,
                "allow_provisional": True,
                "require_disjoint_legs": True,
                "forbid_conflicting_game_sides": False,
                "min_leg_odds": 1.2,
                "max_leg_odds": 3.0,
                "price_weight": 0.15,
            },
        )
    )
    (league_dir / "auto_tickets_2026-11-27_decision.json").write_text(
        json.dumps(legacy_receipt, indent=2) + "\n", encoding="utf-8"
    )

    days = load_localdata_evidence(tmp_path)
    parity = check_parity(days)
    assert parity["checked"] == 2
    assert parity["matched"] == 2
    assert parity["mismatched"] == 0

    # Verify variant parsing supports the new Option 1+A keys cleanly.
    spec = parse_variant_spec(
        "ranking_metric=kelly_log_growth,min_legs=2,max_legs=3,limit=3,min_odds=3.3,max_odds=5.0,max_claimed_edge=0.40,edge_flag_threshold=0.15,forbid_alt_line_reuse=1"
    )
    assert spec.ranking_metric == "kelly_log_growth"
    assert spec.min_legs == 2 and spec.max_legs == 3
    assert spec.min_odds == 3.3 and spec.max_odds == 5.0
    assert spec.max_claimed_edge == 0.40 and spec.edge_flag_threshold == 0.15
    assert spec.forbid_alt_line_reuse is True


