"""Tests for league auditability on ROI band lines and compacted JSON (Item B)."""

from __future__ import annotations

from pathlib import Path

from forecasting.operator import publish_operator_reports
from ingestion.models import League
from scripts.publish_localdata import _compact_roi_band
from tests.test_operator import Store, forecast, game, report


def test_ticket_and_picks_cards_print_league_on_roi_band_lines(tmp_path: Path):
    row = forecast(1, price=1.80)
    band = {
        "league": "NCAA_M",
        "market": "moneyline",
        "period": "full_game",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "historical_pick_count": 28,
        "wins": 17,
        "losses": 11,
        "pushes": 0,
        "observed_roi": 0.049,
        "recency_weighted_roi": 0.034,
        "adjusted_roi": 0.023,
        "qualifies": True,
    }
    paper_filter = {
        "status": "applied",
        "candidate_pick_count": 1,
        "qualifying_pick_count": 1,
        "qualifying_band_count": 1,
        "historical_unique_pick_count": 1907,
        "history_source": "settled_pick_audit",
        "candidate_band_details": {row["forecast_id"]: band},
        "bands": [band],
    }
    row["league"] = "NCAA_M"
    row["scheduled_utc"] = "2026-11-25T20:00:00Z"
    market = report([row])
    market["accas"].update(
        paper_results=[],
        strict_results=[],
        requested_legs=3,
        effective_paper_legs=None,
        paper_single_picks=[{**row, "individual_pick_roi_band": band}],
        paper_pick_filter=paper_filter,
    )
    res = publish_operator_reports(
        Store([game(1, scheduled_utc="2026-11-25T20:00:00Z")]),
        League.NCAA_M,
        market,
        tmp_path,
        target_date="2026-11-25",
        generated_utc="2026-11-25T10:00:00Z",
        run_id="item-b-audit",
    )
    picks_card = Path(res["picks_text"]).read_text(encoding="utf-8")
    ticket_card = Path(res["ticket_text"]).read_text(encoding="utf-8")
    assert (
        "PAPER SINGLE · NCAA_M moneyline full_game · p 0.60-<0.70 · odds 1.50-<2.00 · n=28"
        in picks_card
    )
    assert (
        "ROI band: NCAA_M moneyline full_game · p 0.60-<0.70 · odds 1.50-<2.00 · "
        "n=28 W-L-P=17-11-0 · ROI +4.9% observed / +2.3% adjusted."
        in ticket_card
    )


def test_compact_roi_band_retains_league_without_changing_other_keys():
    raw = {
        "league": "NBA",
        "market": "moneyline",
        "period": "full_game",
        "probability_band": "0.60-<0.70",
        "individual_odds_band": "1.50-<2.00",
        "historical_pick_count": 7,
        "wins": 4,
        "losses": 3,
        "pushes": 0,
        "observed_roi": 0.071,
        "adjusted_roi": 0.118,
        "recency_weighted_roi": 0.070,
        "qualifies": True,
        "extra_untracked_field": "drop_me",
    }
    compacted = _compact_roi_band(raw)
    assert compacted is not None
    assert list(compacted.keys()) == [
        "league",
        "market",
        "period",
        "probability_band",
        "individual_odds_band",
        "historical_pick_count",
        "wins",
        "losses",
        "pushes",
        "observed_roi",
        "adjusted_roi",
        "recency_weighted_roi",
        "qualifies",
    ]
    assert compacted["league"] == "NBA"
    assert "extra_untracked_field" not in compacted
