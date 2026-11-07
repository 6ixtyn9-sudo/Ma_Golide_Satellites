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


def _load_sent(path: Path) -> set[str]:
    if not path.exists():
        return set()
    data = json.loads(path.read_text())
    return set(data.get("sent_sha256", []))


def _save_sent(path: Path, sent: set[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(json.dumps({"sent_sha256": sorted(sent)}, indent=2) + "\n")
    os.replace(temporary, path)


def dispatch_whatsapp(message: str) -> None:
    """Send one Meta WhatsApp text message; credentials stay in the environment."""
    import requests

    token = os.environ.get("WHATSAPP_TOKEN", "").strip()
    phone_id = os.environ.get("WHATSAPP_PHONE_NUMBER_ID", "").strip()
    recipient = os.environ.get("WHATSAPP_RECIPIENT", "").strip()
    if not all((token, phone_id, recipient)):
        raise RuntimeError(
            "WHATSAPP_TOKEN, WHATSAPP_PHONE_NUMBER_ID and "
            "WHATSAPP_RECIPIENT are required"
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
    (dispatcher or dispatch_whatsapp)(message)
    sent.add(digest)
    _save_sent(path, sent)
    result.update({"sent": True, "status": "sent"})
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, help="daily report JSON")
    parser.add_argument("--send", action="store_true")
    parser.add_argument("--ledger", default="localdata/notification_ledger.json")
    args = parser.parse_args()
    if args.send and os.environ.get("ALLOW_NOTIFICATION_WRITES", "").lower() != "true":
        parser.error("--send requires ALLOW_NOTIFICATION_WRITES=true")
    summary = json.loads(Path(args.report).read_text())
    result = notify_summary(summary, send=args.send, ledger_path=args.ledger)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
