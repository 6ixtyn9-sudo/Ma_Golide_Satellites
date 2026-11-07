from __future__ import annotations

import json

import pytest

from scripts.notify import notify_summary, notification_text


def _summary():
    return {
        "league": "NBA",
        "target_date": "2026-11-07",
        "mode": "write",
        "verification": "LIVE_ATTEMPT",
        "counts": {"games": 1, "predictions": 1, "book_snapshots": 0},
        "commands": {
            "pregame": {"fixtures_found": 1, "gaps": [{"required_field": "book_Q1"}]},
            "results": {"finals_written": 0, "gaps": []},
        },
        "warehouse_read_available": True,
    }


def test_notification_is_operational_and_deduplicated(tmp_path):
    summary = _summary()
    message = notification_text(summary)
    assert "1 games" in message
    assert "1 reported gaps" in message
    assert "no betting recommendation" in message
    assert "WHATSAPP_TOKEN" not in message

    ledger = tmp_path / "notification_ledger.json"
    dry = notify_summary(summary, ledger_path=str(ledger))
    assert dry["status"] == "dry_run" and not dry["sent"]
    assert not ledger.exists()

    sent_messages = []
    first = notify_summary(
        summary, send=True, ledger_path=str(ledger), dispatcher=sent_messages.append
    )
    second = notify_summary(
        summary, send=True, ledger_path=str(ledger), dispatcher=sent_messages.append
    )
    assert first["status"] == "sent" and first["sent"]
    assert second["status"] == "already_sent" and not second["sent"]
    assert sent_messages == [message]
    assert len(json.loads(ledger.read_text())["sent_sha256"]) == 1


def test_fixture_report_can_preview_but_never_send(tmp_path):
    summary = _summary() | {"verification": "FIXTURE_VERIFIED"}
    preview = notify_summary(summary, ledger_path=str(tmp_path / "ledger.json"))
    assert preview["status"] == "dry_run"
    with pytest.raises(RuntimeError, match="cannot trigger external"):
        notify_summary(
            summary,
            send=True,
            ledger_path=str(tmp_path / "ledger.json"),
            dispatcher=lambda message: None,
        )
