from __future__ import annotations

import json
from pathlib import Path

import pytest

from ingestion.league_registry import SPECS
from scripts.publish_localdata import (
    STATE_KEYS,
    allowed_lifecycle_recovery_file,
    allowed_operator_file,
    merge_downloaded_localdata,
)


def artifact_file(tmp_path: Path, artifact: str, state: str, name: str, content: bytes):
    path = tmp_path / "artifacts" / artifact / "localdata" / state / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def test_merge_publishes_only_operator_files_and_retains_history(tmp_path):
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_today.json", b"today\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-12.txt", b"picks\n")
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"receipt\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12.txt",
        b"frozen-ticket-card\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-13.txt",
        b"hourly-candidate-card\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-14.txt",
        b"hourly-no-bet-card\n",
    )
    artifact_file(tmp_path, "nba-artifact", "nba", "warehouse.duckdb", b"canonical")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_unapproved.json", b"unexpected")
    artifact_file(tmp_path, "cba-artifact", "cba", "picks_today.json", b"cba\n")
    artifact_file(tmp_path, "wnba-artifact", "wnba", "picks_today.json", b"wnba\n")
    skipped = tmp_path / "artifacts" / "skipped-cba"
    skipped.mkdir(parents=True)
    (skipped / "daily-output.json").write_text("{}")
    repo = tmp_path / "repo"
    historical = repo / "localdata" / "nba" / "picks_2026-11-11.txt"
    historical.parent.mkdir(parents=True)
    historical.write_text("retain me")
    stale_no_bet_card = repo / "localdata" / "nba" / "auto_tickets_2026-11-14.txt"
    stale_no_bet_card.write_text("stale pending candidate\n")

    report = merge_downloaded_localdata(tmp_path / "artifacts", repo)

    assert report["copied_count"] == 5
    compact_pick_card = (repo / "localdata" / "nba" / "picks_2026-11-12.txt").read_text()
    assert "Ma Golide compact pick lifecycle marker" in compact_pick_card
    assert "target_date=2026-11-12" in compact_pick_card
    assert "detailed_card_sha256=" in compact_pick_card
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12_decision.json").read_bytes() == b"receipt\n"
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12.txt").read_bytes() == b"frozen-ticket-card\n"
    # A pre-cutoff card is a mutable published preview, not a receipt or date lock.
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-13.txt").read_bytes() == b"hourly-candidate-card\n"
    assert not (repo / "localdata" / "nba" / "auto_tickets_2026-11-13_decision.json").exists()
    # The terminal state is inline in auto_tickets_state.json; the dated card
    # remains the only per-date no-bet publication.
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-14.txt").read_bytes() == b"hourly-no-bet-card\n"
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12.txt",
        b"later-settlement-rewrite\n",
    )
    settled = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert settled["copied"] == ["localdata/nba/auto_tickets_2026-11-12.txt"]
    assert (repo / "localdata" / "nba" / "auto_tickets_2026-11-12.txt").read_bytes() == b"later-settlement-rewrite\n"
    for state in ("nba", "cba", "wnba"):
        assert not (repo / "localdata" / state / "picks_today.json").exists()
    assert historical.read_text() == "retain me"
    assert not (repo / "localdata" / "nba" / "warehouse.duckdb").exists()
    assert not (repo / "localdata" / "nba" / "picks_unapproved.json").exists()


def test_publisher_unions_aba_history_receipts_and_rejects_count_conflicts(tmp_path):
    receipt = {
        "season_end_year": 2026,
        "calendar_final_count": 185,
        "reconciled_final_count": 183,
        "administrative_disposition_count": 2,
        "not_applicable_rows": 0,
        "blocking_gap_count": 0,
        "status": "calendar_terminal_count_reconciled",
        "source": "aba_liga",
        "evidence": "ROUTE_CONTRACT_REFERENCED",
        "completed_at_utc": "2026-11-20T13:40:00Z",
    }

    def payload(entry):
        return json.dumps(
            {
                "schema_version": "magolide.aba_history_receipts",
                "league": "ABA_LEAGUE",
                "provider": "aba_liga",
                "receipts": {str(entry["season_end_year"]): entry},
            }
        ).encode()

    artifact_file(
        tmp_path, "aba-artifact", "aba_league", "history_receipts.json", payload(receipt)
    )
    # This exact file name is ABA-only, even though the publisher scans all
    # registered state directories.
    artifact_file(tmp_path, "cba-artifact", "cba", "history_receipts.json", payload(receipt))

    repo = tmp_path / "repo"
    first = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    destination = repo / "localdata" / "aba_league" / "history_receipts.json"
    assert first["copied"] == ["localdata/aba_league/history_receipts.json"]
    assert not (repo / "localdata" / "cba" / "history_receipts.json").exists()
    assert json.loads(destination.read_text())["receipts"] == {"2026": receipt}

    artifact_file(
        tmp_path,
        "aba-artifact",
        "aba_league",
        "history_receipts.json",
        payload({**receipt, "completed_at_utc": "2026-11-20T14:00:00Z"}),
    )
    repeated = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert repeated["unchanged"] == ["localdata/aba_league/history_receipts.json"]
    assert json.loads(destination.read_text())["receipts"] == {"2026": receipt}

    artifact_file(
        tmp_path,
        "aba-artifact",
        "aba_league",
        "history_receipts.json",
        payload({**receipt, "calendar_final_count": 186, "reconciled_final_count": 184}),
    )
    with pytest.raises(ValueError, match="ABA history receipt conflicts"):
        merge_downloaded_localdata(tmp_path / "artifacts", repo)


def test_publisher_does_not_regress_settled_pick_or_ticket_card_to_pending(tmp_path):
    artifact_file(
        tmp_path,
        "stale-nba-artifact",
        "nba",
        "picks_2026-11-18.txt",
        "bucket=CAUTION · grade=⏳ PENDING\n".encode(),
    )
    artifact_file(
        tmp_path,
        "stale-nba-artifact",
        "nba",
        "auto_tickets_2026-11-18.txt",
        "[⏳ PENDING] shadow\n".encode(),
    )
    repo = tmp_path / "repo"
    state = repo / "localdata" / "nba"
    state.mkdir(parents=True)
    settled_pick = state / "picks_2026-11-18.txt"
    settled_pick.write_text("grade=❌ LOSS · lifecycle=SETTLED_UNFROZEN\n")
    settled_ticket = state / "auto_tickets_2026-11-18.txt"
    settled_ticket.write_text("[✅ WIN] settled shadow\n")

    report = merge_downloaded_localdata(tmp_path / "artifacts", repo)

    assert report["status"] == "no_changes"
    assert report["settled_card_regressions_skipped"] == [
        "localdata/nba/auto_tickets_2026-11-18.txt",
        "localdata/nba/picks_2026-11-18.txt",
    ]
    assert "LOSS" in settled_pick.read_text()
    assert "WIN" in settled_ticket.read_text()


def test_publisher_accepts_progressive_settlement_with_a_late_pending_card_line(tmp_path):
    artifact_file(
        tmp_path,
        "newer-nba-artifact",
        "nba",
        "auto_tickets_2026-11-18.txt",
        (
            "[PAPER 3-LEG ACCA #1] · ❌ LOSS\n"
            "Early game · ✅ WIN\n"
            "Middle game · ❌ LOSS\n"
            "Late game · ⏳ PENDING\n"
        ).encode(),
    )
    repo = tmp_path / "repo"
    card = repo / "localdata" / "nba" / "auto_tickets_2026-11-18.txt"
    card.parent.mkdir(parents=True)
    card.write_text(
        "[PAPER 3-LEG ACCA #1] · ⏳ PENDING\n"
        "Early game · ✅ WIN\n"
        "Middle game · ⏳ PENDING\n"
        "Late game · ⏳ PENDING\n"
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", repo)

    assert report["settled_card_regressions_skipped"] == []
    assert report["copied"] == ["localdata/nba/auto_tickets_2026-11-18.txt"]
    assert card.read_text() == (
        "[PAPER 3-LEG ACCA #1] · ❌ LOSS\n"
        "Early game · ✅ WIN\n"
        "Middle game · ❌ LOSS\n"
        "Late game · ⏳ PENDING\n"
    )


def test_publisher_rejects_a_same_line_settlement_downgrade_despite_other_progress(tmp_path):
    artifact_file(
        tmp_path,
        "stale-nba-artifact",
        "nba",
        "auto_tickets_2026-11-18.txt",
        (
            "First game · ⏳ PENDING\n"
            "Second game · ❌ LOSS\n"
            "Late game · ✅ WIN\n"
        ).encode(),
    )
    repo = tmp_path / "repo"
    card = repo / "localdata" / "nba" / "auto_tickets_2026-11-18.txt"
    card.parent.mkdir(parents=True)
    original = (
        "First game · ✅ WIN\n"
        "Second game · ❌ LOSS\n"
        "Late game · ⏳ PENDING\n"
    )
    card.write_text(original)

    report = merge_downloaded_localdata(tmp_path / "artifacts", repo)

    assert report["settled_card_regressions_skipped"] == [
        "localdata/nba/auto_tickets_2026-11-18.txt"
    ]
    assert card.read_text() == original


def test_publisher_state_keys_follow_the_central_registry_and_publish_expansion_lanes(tmp_path):
    assert STATE_KEYS == {spec.state_key for spec in SPECS}
    artifact_file(
        tmp_path,
        "acb-artifact",
        "spain_acb",
        "picks_2026-11-18.txt",
        b"expansion draft\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["copied"] == ["localdata/spain_acb/picks_2026-11-18.txt"]
    assert report["skipped_states"] == []


def test_publisher_reports_unknown_state_directories_without_publishing_them(tmp_path):
    artifact_file(tmp_path, "unknown-artifact", "unregistered_league", "picks_2026-11-18.txt", b"x\n")

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["status"] == "no_changes"
    assert report["skipped_states"] == [
        (tmp_path / "artifacts" / "unknown-artifact" / "localdata" / "unregistered_league").resolve().as_posix()
    ]
    assert not (tmp_path / "repo" / "localdata" / "unregistered_league").exists()


def test_publisher_keeps_structured_pick_json_artifact_only_globally(tmp_path):
    dated_json_name = "picks_2026-11-18.json"
    lifecycle_name = "picks_2026-11-18_lifecycle.json"
    retired_freeze_name = "pick_freeze_2026-11-18_0123456789abcdef.json"

    def lifecycle_state(rows):
        return json.dumps(
            {
                "schema_version": "magolide.pick_lifecycle_state",
                "league": "NBA",
                "pending_frozen_games": [["game-1", "2026-11-18T20:00:00Z"]],
                "dated_reports": {"2026-11-18": {"frozen_fixtures": 1}},
                "frozen_by_date": {
                    "2026-11-18": {
                        "game-1": {
                            "schema_version": "magolide.pick_freeze",
                            "target_date": "2026-11-18",
                            "league": "NBA",
                            "game_id": "game-1",
                            "rows": rows,
                        }
                    }
                },
            }
        ).encode()

    for state in ("nba", "ncaa_w"):
        artifact_file(tmp_path, f"{state}-artifact", state, lifecycle_name, b"detailed lifecycle\n")
    artifact_file(tmp_path, "nba-artifact", "nba", dated_json_name, b"draft one\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-18.txt", b"draft one\n")
    artifact_file(
        tmp_path, "nba-artifact", "nba", "pick_lifecycle_state.json", lifecycle_state([{"id": 1}])
    )
    artifact_file(tmp_path, "nba-artifact", "nba", retired_freeze_name, b"retired sidecar\n")

    repo = tmp_path / "repo"
    first = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert first["copied_count"] == 2
    assert not allowed_operator_file(dated_json_name)
    assert not allowed_operator_file(lifecycle_name)
    assert not allowed_operator_file(retired_freeze_name)
    assert not allowed_lifecycle_recovery_file(dated_json_name)
    assert not allowed_lifecycle_recovery_file(lifecycle_name)
    assert not allowed_lifecycle_recovery_file(retired_freeze_name)
    for state in ("nba", "ncaa_w"):
        assert not (repo / "localdata" / state / lifecycle_name).exists()
    destination = repo / "localdata" / "nba"
    assert not (destination / dated_json_name).exists()
    assert not (destination / retired_freeze_name).exists()
    compact_pick_card = (destination / "picks_2026-11-18.txt").read_text()
    assert "Ma Golide compact pick lifecycle marker" in compact_pick_card
    assert "detailed_card_sha256=" in compact_pick_card
    persisted = json.loads((destination / "pick_lifecycle_state.json").read_text())
    assert persisted["frozen_by_date"]["2026-11-18"]["game-1"]["rows"] == [{"id": 1}]

    artifact_file(tmp_path, "nba-artifact", "nba", dated_json_name, b"draft two\n")
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_2026-11-18.txt", b"draft two\n")
    second = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert second["copied"] == ["localdata/nba/picks_2026-11-18.txt"]
    assert not (destination / dated_json_name).exists()
    second_marker = (destination / "picks_2026-11-18.txt").read_text()
    assert "Ma Golide compact pick lifecycle marker" in second_marker
    assert "detailed_card_sha256=" in second_marker

    artifact_file(
        tmp_path, "nba-artifact", "nba", "pick_lifecycle_state.json", lifecycle_state([{"id": 2}])
    )
    with pytest.raises(ValueError, match="immutable inline pick freeze differs"):
        merge_downloaded_localdata(tmp_path / "artifacts", repo)


def test_publisher_compacts_frozen_fixture_and_ticket_evidence_without_changing_identity(tmp_path):
    full_row = {
        "forecast_id": "forecast-1",
        "pick_id": "pick-1",
        "game_id": "game-1",
        "league": "NBA",
        "date": "2026-11-18",
        "scheduled_utc": "2026-11-18T20:00:00Z",
        "home": "Home",
        "away": "Away",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": None,
        "raw_probability": 0.64,
        "calibrated_probability": 0.62,
        "push_probability": 0.0,
        "model": "team_history_empirical",
        "model_version": "model-v1",
        "role": "champion",
        "bucket": "ELIGIBLE_CLEAN",
        "price_evidence": "NAMED_BOOKMAKER_PRICE",
        "paper_only": True,
        "alternative_probabilities": {"away": 0.36},
        "explanation": {"very_large": "x" * 10_000},
        "team_history": {"rows": list(range(100))},
        "context": {"verbose": True},
        "calibration": {
            "status": "calibrated",
            "n": 40,
            "minimum_samples": 30,
            "sufficient_for_accas": True,
            "verbose": "x" * 1_000,
        },
        "eligibility": {"acca_eligible": True, "reasons": [], "verbose": "x" * 1_000},
        "prediction": {"provider": "espn", "observed_utc": "2026-11-18T18:00:00Z"},
        "book": {
            "provider": "theoddsapi",
            "book": "DraftKings",
            "observed_utc": "2026-11-18T18:00:00Z",
            "selected_decimal_price": 1.8,
            "selected_opposite_decimal_price": 2.1,
            "price_provenance": "named_book",
            "source_permission": "unresolved",
            "operational_validation": "unverified",
            "verbose": "x" * 1_000,
        },
        "paper_roi_band": {
            "market": "moneyline",
            "period": "full_game",
            "historical_pick_count": 18,
            "observed_roi": 0.12,
            "adjusted_roi": 0.08,
            "qualifies": True,
            "verbose": "x" * 1_000,
        },
    }
    lifecycle = {
        "schema_version": "magolide.pick_lifecycle_state",
        "league": "NBA",
        "pending_frozen_games": [["game-1", "2026-11-18T20:00:00Z"]],
        "dated_reports": {"2026-11-18": {"frozen_fixtures": 1}},
        "frozen_by_date": {
            "2026-11-18": {
                "game-1": {
                    "schema_version": "magolide.pick_freeze",
                    "target_date": "2026-11-18",
                    "league": "NBA",
                    "game_id": "game-1",
                    "scheduled_utc": "2026-11-18T20:00:00Z",
                    "decision_cutoff_utc": "2026-11-18T19:00:00Z",
                    "frozen_at": "2026-11-18T19:00:00Z",
                    "rows": [full_row],
                }
            }
        },
    }
    receipt = {
        "schema_version": "magolide.paper_ticket_receipt",
        "date": "2026-11-18",
        "league": "NBA",
        "decision_id": "decision-1",
        "status": "frozen",
        "paper_only": True,
        "captured_at": "2026-11-18T19:00:00Z",
        "frozen_at": "2026-11-18T19:00:00Z",
        "stake_units_per_acca": 1.0,
        "stake_units_per_position": 1.0,
        "accas": [
            {
                "ticket_id": "ticket-1",
                "ticket_type": "acca",
                "publication_index": 1,
                "leg_count": 1,
                "combined_decimal_odds": 1.8,
                "stake_units": 1.0,
                "estimated_joint_win_probability": 0.62,
                "paper_only": True,
                "legs": [full_row],
                "individual_pick_filter": {
                    "status": "qualified",
                    "candidate_pick_count": 20,
                    "qualifying_pick_count": 1,
                    "qualifying_band_count": 1,
                    "historical_unique_pick_count": 40,
                    "candidate_band_details": {
                        "forecast-1": full_row["paper_roi_band"],
                        "other": {"verbose": "x" * 1_000},
                    },
                    "bands": [{"verbose": "x" * 1_000}],
                },
            }
        ],
        "singles": [],
    }
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "pick_lifecycle_state.json",
        json.dumps(lifecycle).encode(),
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-18_decision.json",
        json.dumps(receipt).encode(),
    )

    repo = tmp_path / "repo"
    first = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    state = json.loads((repo / "localdata/nba/pick_lifecycle_state.json").read_text())
    compact_row = state["frozen_by_date"]["2026-11-18"]["game-1"]["rows"][0]
    assert compact_row["evidence_schema"] == "magolide.frozen_forecast_evidence.v1"
    assert compact_row["book"]["selected_decimal_price"] == 1.8
    assert compact_row["book"]["source_permission"] == "unresolved"
    assert compact_row["eligibility"] == {"acca_eligible": True, "reasons": []}
    assert "explanation" not in compact_row
    assert "team_history" not in compact_row
    assert len(json.dumps(compact_row)) < len(json.dumps(full_row)) / 3
    # The normal operator can still rebuild a gradeable forecast from the
    # compact inline evidence restored from Git on a later run.
    from forecasting.operator import _forecast_from_row

    restored_forecast = _forecast_from_row(compact_row)
    assert restored_forecast.selection == "home"
    assert restored_forecast.book is not None
    assert restored_forecast.book.price_for("home") == 1.8

    compact_receipt = json.loads(
        (repo / "localdata/nba/auto_tickets_2026-11-18_decision.json").read_text()
    )
    compact_position = compact_receipt["accas"][0]
    assert "bands" not in compact_position["individual_pick_filter"]
    assert set(compact_position["individual_pick_filter"]["candidate_band_details"]) == {"forecast-1"}
    assert "explanation" not in compact_position["legs"][0]

    # A later full artifact is semantically the same immutable receipt/freeze,
    # so the compact Git form is accepted without a conflict.
    repeated = merge_downloaded_localdata(tmp_path / "artifacts", repo)
    assert repeated["status"] == "no_changes"
    assert first["copied_count"] == 2


def test_interrupted_lane_artifact_can_publish_only_durable_pick_lifecycle_evidence(tmp_path):
    artifact = "MaGolide-lifecycle-nba-run"
    dated_json_name = "picks_2026-11-18.json"
    retired_freeze_name = "pick_freeze_2026-11-18_0123456789abcdef.json"
    artifact_file(tmp_path, artifact, "nba", dated_json_name, b"raw draft\n")
    artifact_file(tmp_path, artifact, "nba", "picks_2026-11-18.txt", b"draft card\n")
    artifact_file(tmp_path, artifact, "nba", retired_freeze_name, b"retired sidecar\n")
    lifecycle_state = json.dumps(
        {
            "schema_version": "magolide.pick_lifecycle_state",
            "league": "NBA",
            "pending_frozen_games": [],
            "dated_reports": {},
            "frozen_by_date": {},
        }
    ).encode()
    artifact_file(tmp_path, artifact, "nba", "pick_lifecycle_state.json", lifecycle_state)
    artifact_file(
        tmp_path,
        artifact,
        "nba",
        "auto_tickets_2026-11-18_decision.json",
        b"must-not-publish\n",
    )
    artifact_file(
        tmp_path,
        artifact,
        "nba",
        "auto_tickets_2026-11-18.txt",
        b"must-not-publish\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    destination = tmp_path / "repo" / "localdata" / "nba"
    assert not (destination / dated_json_name).exists()
    assert "Ma Golide compact pick lifecycle marker" in (
        destination / "picks_2026-11-18.txt"
    ).read_text()
    assert not (destination / retired_freeze_name).exists()
    assert (destination / "pick_lifecycle_state.json").read_bytes() == (
        json.dumps(json.loads(lifecycle_state), indent=2, sort_keys=True) + "\n"
    ).encode()
    assert not (destination / "auto_tickets_2026-11-18_decision.json").exists()
    assert not (destination / "auto_tickets_2026-11-18.txt").exists()
    assert {Path(path).name for path in report["recovery_files_skipped"]} == {
        dated_json_name,
        retired_freeze_name,
        "auto_tickets_2026-11-18_decision.json",
        "auto_tickets_2026-11-18.txt",
    }
    assert allowed_lifecycle_recovery_file("pick_lifecycle_state.json")
    assert not allowed_lifecycle_recovery_file(retired_freeze_name)
    assert not allowed_lifecycle_recovery_file("auto_tickets_2026-11-18_decision.json")


def test_publisher_rejects_retired_isolated_no_bet_sidecars(tmp_path):
    artifact_file(
        tmp_path,
        "elw-artifact",
        "euroleague_w",
        "auto_tickets_2026-11-18_no_bet.json",
        b"retired-sidecar\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["status"] == "no_changes"
    assert report["copied"] == []


def test_publisher_strips_hourly_telemetry_before_git_publication(tmp_path):
    state = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_state.json",
        json.dumps(
            {
                "schema_version": "magolide.paper_ticket_state",
                "frozen_by_date": {"2026-11-12": {"final": True, "decision_id": "receipt"}},
                "scans": [{"run_id": "hourly-1", "large": "x" * 100_000}],
                "last_generated_utc": "2026-11-12T18:00:00Z",
            }
        ).encode(),
    )
    performance = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_performance.json",
        json.dumps(
            {
                "schema_version": "magolide.operator",
                "generated_utc": "2026-11-12T18:00:00Z",
                "history": [{"date": "2026-11-12", "positions": []}],
            }
        ).encode(),
    )

    first = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    destination = tmp_path / "repo" / "localdata" / "nba"
    saved_state = json.loads((destination / state.name).read_text())
    saved_performance = json.loads((destination / performance.name).read_text())
    assert first["copied_count"] == 2
    assert saved_state["frozen_by_date"]["2026-11-12"]["decision_id"] == "receipt"
    assert "scans" not in saved_state and "last_generated_utc" not in saved_state
    assert "generated_utc" not in saved_performance

    # A new hourly scan/clock value alone cannot create another Git update.
    state.write_text(
        json.dumps(
            {
                **json.loads(state.read_text()),
                "scans": [{"run_id": "hourly-2", "large": "y" * 100_000}],
                "last_generated_utc": "2026-11-12T19:00:00Z",
            }
        )
    )
    performance.write_text(
        json.dumps(
            {
                **json.loads(performance.read_text()),
                "generated_utc": "2026-11-12T19:00:00Z",
            }
        )
    )
    second = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert second["status"] == "no_changes"


def test_git_no_bet_decision_blocks_a_stale_cached_ticket_receipt(tmp_path):
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"stale-ticket\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba"
    destination.mkdir(parents=True)
    (destination / "auto_tickets_state.json").write_text(
        json.dumps(
            {
                "schema_version": "magolide.paper_ticket_state",
                "frozen_by_date": {"2026-11-12": {"no_bet": True}},
            }
        )
        + "\n"
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")

    assert report["status"] == "no_changes"
    assert report["no_bet_receipts_skipped"] == [
        "localdata/nba/auto_tickets_2026-11-12_decision.json"
    ]
    assert not (destination / "auto_tickets_2026-11-12_decision.json").exists()


def test_immutable_receipt_conflict_aborts_without_overwriting(tmp_path):
    source = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"new-ticket\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba" / source.name
    destination.parent.mkdir(parents=True)
    destination.write_bytes(b"original-frozen-ticket\n")

    with pytest.raises(ValueError, match="immutable ticket receipt differs"):
        merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert destination.read_bytes() == b"original-frozen-ticket\n"


def test_identical_receipt_is_verified_and_not_rewritten(tmp_path):
    source = artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_decision.json",
        b"same-receipt\n",
    )
    destination = tmp_path / "repo" / "localdata" / "nba" / source.name
    destination.parent.mkdir(parents=True)
    destination.write_bytes(source.read_bytes())

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert report["unchanged_count"] == 1
    assert report["immutable_receipts_verified"] == [
        "localdata/nba/auto_tickets_2026-11-12_decision.json"
    ]


def test_withdrawn_cached_receipt_is_not_republished(tmp_path):
    artifact_file(
        tmp_path,
        "ncaa-m-artifact",
        "ncaa_m",
        "auto_tickets_2026-11-13_decision.json",
        json.dumps({"decision_id": "defective-run"}).encode(),
    )
    localdata = tmp_path / "repo" / "localdata" / "ncaa_m"
    localdata.mkdir(parents=True)
    (localdata / "auto_tickets_state.json").write_text(
        json.dumps(
            {
                "frozen_by_date": {
                    "2026-11-13": {
                        "withdrawn": True,
                        "decision_id": "defective-run",
                    }
                }
            }
        )
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert report["withdrawn_receipts_skipped"] == [
        "localdata/ncaa_m/auto_tickets_2026-11-13_decision.json"
    ]
    assert not (localdata / "auto_tickets_2026-11-13_decision.json").exists()


def test_allowlist_is_narrow_and_missing_artifact_directory_is_a_noop(tmp_path):
    expected = {
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
        "auto_tickets_performance.txt",
        "auto_tickets_2026-11-12.txt",
        "auto_tickets_2026-11-12_decision.json",
        "coverage_gap_2026-11-12.json",
    }
    assert all(allowed_operator_file(name) for name in expected)
    assert all(
        allowed_operator_file(name)
        for name in {
            "picks_2026-11-12.txt",
            "pick_lifecycle_state.json",
        }
    )
    assert not allowed_operator_file("picks_2026-11-12.json")
    assert not allowed_operator_file("pick_freeze_2026-11-12_0123456789abcdef.json")
    assert not allowed_operator_file("picks_2026-11-12_lifecycle.json")
    assert not allowed_operator_file("picks_today.json")
    assert not allowed_operator_file("picks_audit_rolling.json")
    assert not allowed_operator_file("picks_audit_2026-11-12.md")
    assert not allowed_operator_file("picks_unapproved.json")
    assert not allowed_operator_file("auto_tickets_2026-11-12.json")
    assert not allowed_operator_file("auto_tickets_2026-11-12_no_bet.json")
    assert not allowed_operator_file("auto_tickets_2026-11-12_preview.txt")
    assert not allowed_operator_file("warehouse.duckdb")
    report = merge_downloaded_localdata(tmp_path / "missing", tmp_path / "repo")
    assert report["status"] == "no_changes"
    assert not (tmp_path / "repo" / "localdata").exists()


def test_preview_sidecar_artifacts_are_ignored_by_the_publisher(tmp_path):
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12_preview.txt",
        b"stale preview sidecar\n",
    )
    artifact_file(
        tmp_path,
        "nba-artifact",
        "nba",
        "auto_tickets_2026-11-12.txt",
        b"current dated card\n",
    )

    report = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    localdata = tmp_path / "repo" / "localdata" / "nba"

    assert report["copied"] == ["localdata/nba/auto_tickets_2026-11-12.txt"]
    assert report["copied_count"] == 1
    assert (localdata / "auto_tickets_2026-11-12.txt").exists()
    assert not (localdata / "auto_tickets_2026-11-12_preview.txt").exists()


def test_cli_reports_json(tmp_path, capsys, monkeypatch):
    artifact_file(tmp_path, "nba-artifact", "nba", "picks_today.json", b"[]\n")
    from scripts.publish_localdata import main

    import sys

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "publish_localdata.py",
            "--artifacts",
            str(tmp_path / "artifacts"),
            "--repo",
            str(tmp_path / "repo"),
        ],
    )
    assert main() == 0
    output = json.loads(capsys.readouterr().out)
    assert output["copied_count"] == 0
    assert not (tmp_path / "repo" / "localdata" / "nba" / "picks_today.json").exists()


def test_publisher_preserves_time_locked_euroleague_consensus_roster(tmp_path):
    roster_name = "euroleague_consensus_paper_2026-11-18_roster.json"
    card_name = "euroleague_consensus_paper_2026-11-18.txt"
    roster = artifact_file(
        tmp_path,
        "euroleague-artifact",
        "euroleague",
        roster_name,
        b'{"frozen": true, "selections": []}\n',
    )
    artifact_file(
        tmp_path,
        "euroleague-artifact",
        "euroleague",
        card_name,
        b"initial consensus grades\n",
    )

    first = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert set(first["copied"]) == {
        f"localdata/euroleague/{roster_name}",
        f"localdata/euroleague/{card_name}",
    }
    assert allowed_operator_file(roster_name)
    assert allowed_operator_file(card_name)
    assert allowed_lifecycle_recovery_file(roster_name)
    assert not allowed_lifecycle_recovery_file(card_name)

    # The text card is a mutable official-result view, while the no-stake
    # roster is the immutable selection/evidence record.
    artifact_file(
        tmp_path,
        "euroleague-artifact",
        "euroleague",
        card_name,
        b"refreshed consensus grades\n",
    )
    second = merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
    assert second["copied"] == [f"localdata/euroleague/{card_name}"]

    roster.write_bytes(b'{"frozen": true, "selections": ["replacement"]}\n')
    with pytest.raises(ValueError, match="immutable EuroLeague consensus-paper roster differs"):
        merge_downloaded_localdata(tmp_path / "artifacts", tmp_path / "repo")
