#!/usr/bin/env python3
"""Optional WhatsApp operations heartbeat for a completed daily report.

This sends acquisition/coverage facts, never invents a betting pick. Dry run is
the default. A content-hash delivery ledger prevents duplicate notifications
when the three-hour orchestrator repeats an unchanged report summary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def notification_text(summary: Dict[str, Any]) -> str:
    counts = summary.get("counts", {})
    commands = summary.get("commands", {})
    pregame = commands.get("pregame", {})
    results = commands.get("results", {})
    gaps = sum(len(report.get("gaps", [])) for report in commands.values())
    return "\n".join(
        [
            f"Ma Golide {summary.get('league', '')} · {summary.get('target_date', '')}",
            (
                f"Mode: {summary.get('mode', 'unknown')} · "
                f"Verification: {summary.get('verification', 'unknown')}"
            ),
            (
                f"Warehouse: {counts.get('games', 0)} games, "
                f"{counts.get('predictions', 0)} prediction slots, "
                f"{counts.get('book_snapshots', 0)} book snapshots."
            ),
            (
                f"This run: {pregame.get('fixtures_found', 0)} fixtures, "
                f"{results.get('finals_written', 0)} reconciled finals, "
                f"{gaps} reported gaps."
            ),
            "Direct reads: local DuckDB warehouse.",
            "Operational heartbeat only; no betting recommendation is implied.",
        ]
    )


def message_hash(message: str) -> str:
    return hashlib.sha256(message.encode()).hexdigest()


def _load_sent(path: Path) -> Dict[str, str]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text())
    deliveries = data.get("deliveries")
    if isinstance(deliveries, list):
        return {
            str(item["sha256"]): str(item["sent_utc"])
            for item in deliveries
            if isinstance(item, dict) and item.get("sha256") and item.get("sent_utc")
        }
    # Read the original unbounded hash-list format once; entries without a
    # timestamp are eligible for immediate retention cleanup on the next save.
    return {str(digest): "" for digest in data.get("sent_sha256", [])}


def _save_sent(
    path: Path,
    sent: Dict[str, str],
    *,
    now: datetime,
    retention_days: int,
    max_entries: int,
) -> int:
    if retention_days < 1 or max_entries < 1:
        raise ValueError("notification retention and max entries must be positive")
    cutoff = now - timedelta(days=retention_days)
    deliveries = []
    for digest, timestamp in sent.items():
        try:
            sent_at = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
            if sent_at.tzinfo is None:
                continue
            sent_at = sent_at.astimezone(timezone.utc)
        except (AttributeError, ValueError):
            continue
        if sent_at >= cutoff:
            deliveries.append((digest, sent_at))
    deliveries.sort(key=lambda item: (item[1], item[0]), reverse=True)
    deliveries = deliveries[:max_entries]
    deliveries.sort(key=lambda item: (item[1], item[0]))
    payload = {
        "retention_days": retention_days,
        "max_entries": max_entries,
        "deliveries": [
            {
                "sha256": digest,
                "sent_utc": sent_at.isoformat().replace("+00:00", "Z"),
            }
            for digest, sent_at in deliveries
        ],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(payload, indent=2) + "\n")
    os.replace(temporary, path)
    return len(deliveries)


def dispatch_whatsapp(message: str) -> None:
    """Send one Meta WhatsApp text message; credentials stay in the environment."""
    import requests

    token = os.environ.get("WHATSAPP_TOKEN", "").strip()
    phone_id = os.environ.get("WHATSAPP_PHONE_NUMBER_ID", "").strip()
    recipient = os.environ.get("WHATSAPP_RECIPIENT", "").strip()
    if not all((token, phone_id, recipient)):
        raise RuntimeError(
            "WHATSAPP_TOKEN, WHATSAPP_PHONE_NUMBER_ID and WHATSAPP_RECIPIENT are required"
        )
    response = requests.post(
        f"https://graph.facebook.com/v22.0/{phone_id}/messages",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "messaging_product": "whatsapp",
            "to": recipient,
            "type": "text",
            "text": {"body": message},
        },
        timeout=30,
    )
    response.raise_for_status()


def notify_summary(
    summary: Dict[str, Any],
    *,
    send: bool = False,
    ledger_path: str = "localdata/notification_ledger.json",
    dispatcher: Callable[[str], None] | None = None,
    retention_days: int = 30,
    max_entries: int = 1000,
    now: datetime | None = None,
) -> Dict[str, Any]:
    message = notification_text(summary)
    digest = message_hash(message)
    result: Dict[str, Any] = {
        "message_sha256": digest,
        "message": message,
        "sent": False,
        "status": "dry_run",
    }
    if not send:
        return result
    if summary.get("verification") == "FIXTURE_VERIFIED":
        raise RuntimeError("fixture reports cannot trigger external notifications")

    path = Path(ledger_path)
    sent = _load_sent(path)
    if digest in sent:
        result["status"] = "already_sent"
        return result
    if retention_days < 1 or max_entries < 1:
        raise ValueError("notification retention and max entries must be positive")
    sent_at = now or datetime.now(timezone.utc)
    if sent_at.tzinfo is None:
        raise ValueError("notification timestamp must be timezone-aware")
    sent_at = sent_at.astimezone(timezone.utc)
    (dispatcher or dispatch_whatsapp)(message)
    sent[digest] = sent_at.isoformat().replace("+00:00", "Z")
    retained = _save_sent(
        path,
        sent,
        now=sent_at,
        retention_days=retention_days,
        max_entries=max_entries,
    )
    result.update({"sent": True, "status": "sent", "ledger_entries": retained})
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, help="daily report JSON")
    parser.add_argument("--send", action="store_true")
    parser.add_argument("--ledger", default="localdata/notification_ledger.json")
    parser.add_argument("--retention-days", type=int, default=30)
    parser.add_argument("--max-entries", type=int, default=1000)
    args = parser.parse_args()
    if args.send and os.environ.get("ALLOW_NOTIFICATION_WRITES", "").lower() != "true":
        parser.error("--send requires ALLOW_NOTIFICATION_WRITES=true")
    summary = json.loads(Path(args.report).read_text())
    result = notify_summary(
        summary,
        send=args.send,
        ledger_path=args.ledger,
        retention_days=args.retention_days,
        max_entries=args.max_entries,
    )
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
