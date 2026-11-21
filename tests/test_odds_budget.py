"""Unit tests for the serialized The Odds API free-tier reservation guard."""

from __future__ import annotations

from datetime import datetime, timezone

import pytest

from ingestion.odds_budget import (
    DAILY_RESERVATION_CAP,
    HARD_RESERVE,
    MONTHLY_RESERVATION_CAP,
    BudgetStateError,
    empty_state,
    reconcile_provider_quota,
    reserve,
    validate_permit,
    validate_state,
)

NOW = datetime(2026, 11, 18, 9, 0, tzinfo=timezone.utc)  # 11:00 SAST


def _matrix(*leagues: str):
    return {
        "include": [
            {
                "league": league,
                "target_date": "2026-11-18",
                "gate_reason": "ticket_creation_or_acquisition_due",
            }
            for league in leagues
        ]
    }


def test_reservation_is_global_prioritised_and_bounded_to_four_full_game_pulls():
    state, permits = reserve(
        empty_state(NOW),
        _matrix("WNBA", "NBA", "NCAA_M", "NCAA_W", "AUSTRALIA_NBL"),
        run_id="run-1",
        now=NOW,
        automated_tick=True,
    )

    # The unverified/probationary NBL lane is never promoted by budget plumbing.
    assert list(permits) == [
        "NCAA_W|2026-11-18",
        "NBA|2026-11-18",
        "NCAA_M|2026-11-18",
        "WNBA|2026-11-18",
    ]
    assert sum(item["maximum_credits"] for item in state["reservations"]) == DAILY_RESERVATION_CAP
    assert all(permit["maximum_credits"] == 3 for permit in permits.values())
    assert all(permit["markets"] == ["h2h", "totals", "spreads"] for permit in permits.values())
    assert all(validate_permit(permit, league=permit["league"], target_date="2026-11-18") for permit in permits.values())


def test_non_ticket_reasons_manual_rehearsal_and_duplicates_never_spend_a_slot():
    draft_only = {
        "include": [
            {
                "league": "NCAA_W",
                "target_date": "2026-11-18",
                "gate_reason": "pick_draft_publication_due; frozen_pick_settlement_due",
            }
        ]
    }
    state = empty_state(NOW)
    same_state, permits = reserve(
        state, draft_only, run_id="manual", now=NOW, automated_tick=True
    )
    assert same_state == state
    assert permits == {}

    state, permits = reserve(
        state, _matrix("NCAA_W"), run_id="run-1", now=NOW, automated_tick=True
    )
    assert permits
    unchanged, duplicate = reserve(
        state, _matrix("NCAA_W"), run_id="run-2", now=NOW, automated_tick=True
    )
    assert duplicate == {}
    assert unchanged == state

    unchanged, manual = reserve(
        state, _matrix("NBA"), run_id="manual", now=NOW, automated_tick=False
    )
    assert manual == {}
    assert unchanged == state


def test_monthly_cap_and_provider_hard_reserve_fail_closed():
    cap_now = datetime(2026, 12, 31, 9, 0, tzinfo=timezone.utc)
    state = empty_state(cap_now)
    # Thirty earlier SAST days times four allowed lanes consume all 360 planned
    # credits while preserving unique league/date identities.
    for day_number in range(1, 31):
        day = f"2026-12-{day_number:02d}"
        for league in ("NCAA_W", "NBA", "NCAA_M", "WNBA"):
            state["reservations"].append(
                {
                    "permit_id": f"{league}-{day}",
                    "run_id": "old",
                    "league": league,
                    "target_date": day,
                    "budget_date": day,
                    "issued_utc": "2026-12-01T00:00:00Z",
                    "maximum_credits": 3,
                    "markets": ["h2h", "totals", "spreads"],
                }
            )
    state = validate_state(state, cap_now)
    assert sum(item["maximum_credits"] for item in state["reservations"]) == MONTHLY_RESERVATION_CAP
    unchanged, permits = reserve(
        state,
        {"include": [{"league": "NCAA_W", "target_date": "2026-12-31", "gate_reason": "ticket_creation_or_acquisition_due"}]},
        run_id="cap",
        now=cap_now,
        automated_tick=True,
    )
    assert permits == {}
    assert unchanged == state

    # A provider-reported reserve threshold blocks a new month even where the
    # local reservation ledger has plenty of headroom.
    state = empty_state(NOW)
    state["last_provider_quota"] = {
        "x-requests-remaining": HARD_RESERVE,
        "x-requests-used": 360,
        "x-requests-last": 3,
        "observed_utc": "2026-11-18T08:00:00Z",
    }
    unchanged, permits = reserve(
        state, _matrix("NCAA_W"), run_id="run-reserve", now=NOW, automated_tick=True
    )
    assert permits == {}
    assert unchanged == state

    # The final three-credit request must not consume the 140-credit reserve.
    state["last_provider_quota"]["x-requests-remaining"] = HARD_RESERVE + 2
    unchanged, permits = reserve(
        state, _matrix("NCAA_W"), run_id="run-near-reserve", now=NOW, automated_tick=True
    )
    assert permits == {}
    assert unchanged == state


def test_reconcile_uses_the_most_conservative_parallel_header_values():
    state = reconcile_provider_quota(
        empty_state(NOW),
        [
            {"x-requests-remaining": 491, "x-requests-used": 9, "x-requests-last": 3},
            {"x-requests-remaining": 488, "x-requests-used": 12, "x-requests-last": 3},
            {"x-requests-remaining": "bad", "x-requests-used": 14, "x-requests-last": 3},
        ],
        now=NOW,
    )
    assert state["last_provider_quota"] == {
        "x-requests-remaining": 488,
        "x-requests-used": 12,
        "x-requests-last": 3,
        "observed_utc": "2026-11-18T09:00:00Z",
    }


def test_invalid_calendar_dates_are_rejected_not_relaxed():
    invalid_target = {
        "include": [
            {
                "league": "NBA",
                "target_date": "2026-02-30",
                "gate_reason": "ticket_creation_or_acquisition_due",
            }
        ]
    }
    with pytest.raises(BudgetStateError, match="target date"):
        reserve(empty_state(NOW), invalid_target, run_id="invalid-date", now=NOW, automated_tick=True)

    state = empty_state(NOW)
    state["reservations"] = [
        {
            "permit_id": "bad-date",
            "run_id": "r",
            "league": "NBA",
            "target_date": "2026-02-30",
            "budget_date": "2026-11-18",
            "issued_utc": "2026-11-18T09:00:00Z",
            "maximum_credits": 3,
            "markets": ["h2h", "totals", "spreads"],
        }
    ]
    with pytest.raises(BudgetStateError, match="target/budget date"):
        validate_state(state, NOW)


def test_invalid_state_and_permit_are_rejected_not_relaxed():
    with pytest.raises(BudgetStateError, match="unknown schema"):
        validate_state({"schema_version": "wrong"}, NOW)
    with pytest.raises(BudgetStateError, match="provider permit target date"):
        validate_permit(
            {
                "schema_version": "magolide.theoddsapi_permit",
                "permit_id": "p",
                "run_id": "r",
                "league": "NBA",
                "target_date": "2026-02-30",
                "issued_utc": "2026-11-18T09:00:00Z",
                "maximum_credits": 3,
                "markets": ["h2h", "totals", "spreads"],
            },
            league="NBA",
            target_date="2026-02-30",
        )
    with pytest.raises(BudgetStateError, match="unsupported markets"):
        validate_permit(
            {
                "schema_version": "magolide.theoddsapi_permit",
                "permit_id": "p",
                "run_id": "r",
                "league": "NBA",
                "target_date": "2026-11-18",
                "issued_utc": "2026-11-18T09:00:00Z",
                "maximum_credits": 3,
                "markets": ["h2h"],
            },
            league="NBA",
            target_date="2026-11-18",
        )
