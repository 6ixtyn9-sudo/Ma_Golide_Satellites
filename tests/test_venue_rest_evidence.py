"""Tests for observational venue and rest evidence at pick time (Item A)."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Dict

from forecasting.operator import _render_picks, _venue_rest_evidence_line
from forecasting.service import _extract_neutral_site, _venue_rest_evidence
from forecasting.types import Forecast
from ingestion.models import Game, GameStatus, League
from scripts.publish_localdata import (
    PENDING_GRADE_RE,
    _compact_frozen_forecast,
    _would_regress_settled_card,
)


def _game(
    game_id: str,
    home: str,
    away: str,
    scheduled_utc: str,
    *,
    status: GameStatus = GameStatus.FINAL,
    league: League = League.NCAA_M,
    evidence: Dict | None = None,
    provider_ids: Dict[str, str] | None = None,
) -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026-27",
        home_team_id=f"team:{home}",
        away_team_id=f"team:{away}",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled_utc,
        status=status,
        provider_ids=provider_ids or {"espn": game_id},
        evidence=evidence or {},
    )


def test_neutral_site_tri_state_across_espn_bal_and_unknown():
    absent = _game("g1", "Alpha", "Beta", "2026-11-25T20:00:00Z", status=GameStatus.SCHEDULED)
    assert _extract_neutral_site(absent) == (None, None)

    espn_true = _game(
        "g2",
        "Alpha",
        "Beta",
        "2026-11-25T20:00:00Z",
        status=GameStatus.SCHEDULED,
        evidence={"espn": {"scoreboard": {"neutral_site": True}}},
    )
    assert _extract_neutral_site(espn_true) == (True, "espn.scoreboard")

    espn_false = _game(
        "g3",
        "Alpha",
        "Beta",
        "2026-11-25T20:00:00Z",
        status=GameStatus.SCHEDULED,
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    assert _extract_neutral_site(espn_false) == (False, "espn.scoreboard")

    bal_true = _game(
        "g4",
        "Petro",
        "Al Ahly",
        "2026-11-25T20:00:00Z",
        status=GameStatus.SCHEDULED,
        league=League.BASKETBALL_AFRICA_LEAGUE,
        evidence={"provider": "bal", "neutral": True},
    )
    assert _extract_neutral_site(bal_true) == (True, "bal")

    bal_none = _game(
        "g5",
        "Petro",
        "Al Ahly",
        "2026-11-25T20:00:00Z",
        status=GameStatus.SCHEDULED,
        league=League.BASKETBALL_AFRICA_LEAGUE,
        evidence={"provider": "bal", "neutral": None},
    )
    assert _extract_neutral_site(bal_none) == (None, None)


def test_exact_rest_hours_and_signed_diff():
    target = _game(
        "target",
        "Home Club",
        "Away Club",
        "2026-11-25T20:00:00Z",
        status=GameStatus.SCHEDULED,
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    prev_home = _game("h-prev", "Other A", "Home Club", "2026-11-22T20:00:00Z")  # 72.0h
    prev_away = _game("a-prev", "Away Club", "Other B", "2026-11-23T20:00:00Z")  # 48.0h

    ev = _venue_rest_evidence(target, [prev_home, prev_away, target], as_of_utc="2026-11-25T18:00:00Z")
    assert ev["neutral_site"] is False
    assert ev["neutral_site_source"] == "espn.scoreboard"
    assert ev["home_rest_hours"] == 72.0
    assert ev["away_rest_hours"] == 48.0
    assert ev["rest_diff_hours"] == 24.0
    assert ev["home_back_to_back"] is False
    assert ev["away_back_to_back"] is False
    assert ev["any_back_to_back"] is False
    assert ev["home_prev_fixture_id"] == "h-prev"
    assert ev["home_prev_fixture_date"] == "2026-11-22"
    assert ev["away_prev_fixture_id"] == "a-prev"
    assert ev["away_prev_fixture_date"] == "2026-11-23"


def test_fractional_rest_and_tri_state_back_to_back():
    target = _game("target", "Alpha", "Beta", "2026-11-25T18:30:00Z", status=GameStatus.SCHEDULED)
    # 42.5 hours before target
    prev_home = _game("h1", "Alpha", "Gamma", "2026-11-24T00:00:00Z")
    ev_partial = _venue_rest_evidence(target, [prev_home, target], as_of_utc="2026-11-25T16:00:00Z")
    assert ev_partial["home_rest_hours"] == 42.5
    assert ev_partial["home_back_to_back"] is False
    assert ev_partial["away_rest_hours"] is None
    assert ev_partial["away_back_to_back"] is None
    assert ev_partial["rest_diff_hours"] is None
    # One side False and other unknown -> any_back_to_back stays None (tri-state)
    assert ev_partial["any_back_to_back"] is None

    # Now add a short-rest away game (22.0h < 26.0h) -> any_back_to_back becomes True
    prev_away = _game("a1", "Delta", "Beta", "2026-11-24T20:30:00Z")
    ev_b2b = _venue_rest_evidence(
        target, [prev_home, prev_away, target], as_of_utc="2026-11-25T16:00:00Z"
    )
    assert ev_b2b["away_rest_hours"] == 22.0
    assert ev_b2b["away_back_to_back"] is True
    assert ev_b2b["home_back_to_back"] is False
    assert ev_b2b["any_back_to_back"] is True
    assert ev_b2b["rest_diff_hours"] == 20.5


def test_no_prior_game_is_none_never_zero():
    target = _game("target", "Fresh Home", "Fresh Away", "2026-11-25T20:00:00Z", status=GameStatus.SCHEDULED)
    ev = _venue_rest_evidence(target, [target], as_of_utc="2026-11-25T18:00:00Z")
    assert ev["neutral_site"] is None
    assert ev["home_rest_hours"] is None
    assert ev["away_rest_hours"] is None
    assert ev["rest_diff_hours"] is None
    assert ev["home_back_to_back"] is None
    assert ev["away_back_to_back"] is None
    assert ev["any_back_to_back"] is None


def test_relist_within_6h_and_cancelled_postponed_abandoned_and_post_as_of_excluded():
    target = _game("target", "Alpha", "Beta", "2026-11-25T20:00:00Z", status=GameStatus.SCHEDULED)
    # Valid earlier game for Alpha: 50h before target
    valid_home = _game("h-real", "Alpha", "Other", "2026-11-23T18:00:00Z", status=GameStatus.FINAL)
    # Re-listed duplicate 4h before target -> must be skipped
    relist_home = _game("h-relist", "Alpha", "Beta", "2026-11-25T16:00:00Z", status=GameStatus.SCHEDULED)
    # Excluded statuses 20h before target -> must be skipped
    postponed = _game("a-post", "Beta", "X", "2026-11-24T22:00:00Z", status=GameStatus.POSTPONED)
    cancelled = _game("a-canc", "Beta", "Y", "2026-11-24T21:00:00Z", status=GameStatus.CANCELLED)
    abandoned = _game("a-aband", "Beta", "Z", "2026-11-24T20:00:00Z", status=GameStatus.ABANDONED)
    # Game scheduled after as_of_utc (even though before target) -> must be excluded
    after_as_of = _game("a-future", "Beta", "W", "2026-11-25T10:00:00Z", status=GameStatus.SCHEDULED)
    # Game scheduled after target -> must be excluded
    after_target = _game("h-after", "Alpha", "Q", "2026-11-26T20:00:00Z", status=GameStatus.SCHEDULED)

    ev = _venue_rest_evidence(
        target,
        [target, valid_home, relist_home, postponed, cancelled, abandoned, after_as_of, after_target],
        as_of_utc="2026-11-25T08:00:00Z",
    )
    assert ev["home_prev_fixture_id"] == "h-real"
    assert ev["home_rest_hours"] == 50.0
    assert ev["away_prev_fixture_id"] is None
    assert ev["away_rest_hours"] is None


def test_forecast_identity_hash_unchanged_by_venue_rest_evidence():
    base = Forecast(
        game_id="g1",
        league="NBA",
        scheduled_utc="2026-11-25T20:00:00Z",
        home_name="Alpha",
        away_name="Beta",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=0.62,
        alternative_probabilities={"home": 0.62, "away": 0.38},
        source="team_history_empirical",
    )
    with_ev = Forecast(
        game_id="g1",
        league="NBA",
        scheduled_utc="2026-11-25T20:00:00Z",
        home_name="Alpha",
        away_name="Beta",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=0.62,
        alternative_probabilities={"home": 0.62, "away": 0.38},
        source="team_history_empirical",
        venue_rest_evidence={"neutral_site": True, "home_rest_hours": 24.0},
    )
    assert base.forecast_id == with_ev.forecast_id


def test_pick_card_renders_unknowns_honestly_and_never_triggers_pending_or_card_regression(
    tmp_path: Path,
):
    row = {
        "pick_id": "abc123",
        "forecast_id": "abc123",
        "game_id": "g1",
        "league": "NCAA_M",
        "date": "2026-11-25",
        "scheduled_utc": "2026-11-25T20:00:00Z",
        "home": "Alpha",
        "away": "Beta",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": None,
        "raw_probability": 0.65,
        "calibrated_probability": 0.65,
        "model": "team_history_empirical",
        "model_version": "team_history_empirical",
        "bucket": "WATCHLIST_NO_ODDS",
        "price_evidence": "UNMATCHED",
        "grade_status": "win",
        "outcome": "win",
        "lifecycle_status": "SETTLED",
        "eligibility": {"acca_eligible": False, "reasons": ["no_exact_bookmaker_market"]},
        "calibration": {"status": "calibrated", "n": 40},
        "prediction": {"provider": "history"},
        "venue_rest_evidence": {
            "neutral_site": None,
            "neutral_site_source": None,
            "home_rest_hours": None,
            "away_rest_hours": None,
            "rest_diff_hours": None,
            "home_back_to_back": None,
            "away_back_to_back": None,
            "any_back_to_back": None,
        },
    }
    line = _venue_rest_evidence_line(row)
    assert line is not None
    assert "neutral_site=unknown" in line
    assert "home_rest=unknown" in line
    assert "away_rest=unknown" in line
    assert "diff=unknown" in line
    assert "any_b2b=unknown" in line
    assert "0.0h" not in line
    assert "0h" not in line
    assert PENDING_GRADE_RE.search(line) is None

    # Card without venue_rest_evidence vs card with venue_rest_evidence must not
    # trigger settled-card anti-regression checks in publish_localdata.
    row_before = deepcopy(row)
    row_before.pop("venue_rest_evidence")
    card_before = _render_picks(
        [row_before],
        league="NCAA_M",
        target_date="2026-11-25",
        generated_utc="2026-11-26T06:00:00Z",
        timezone_name="Africa/Johannesburg",
    )
    card_after = _render_picks(
        [row],
        league="NCAA_M",
        target_date="2026-11-25",
        generated_utc="2026-11-26T06:00:00Z",
        timezone_name="Africa/Johannesburg",
    )
    assert line in card_after
    dest = tmp_path / "picks_2026-11-25.txt"
    dest.write_text(card_before, encoding="utf-8")
    assert _would_regress_settled_card(dest, card_after.encode("utf-8")) is False
    dest.write_text(card_after, encoding="utf-8")
    assert _would_regress_settled_card(dest, card_before.encode("utf-8")) is False

    # Compaction retains structured venue_rest_evidence and changes detailed hash deterministically.
    compact_before = _compact_frozen_forecast(row_before)
    compact_after = _compact_frozen_forecast(row)
    assert isinstance(compact_before, dict) and isinstance(compact_after, dict)
    assert "venue_rest_evidence" not in compact_before
    assert compact_after["venue_rest_evidence"] == row["venue_rest_evidence"]
    assert compact_before["detailed_evidence_sha256"] != compact_after["detailed_evidence_sha256"]
