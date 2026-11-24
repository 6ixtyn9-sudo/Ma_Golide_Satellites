#!/usr/bin/env python3
"""Render the redacted operational summary for $GITHUB_STEP_SUMMARY.

GitHub aborts the step-summary upload above 1024k of content, and the naive
heredoc this script replaces embedded unbounded blobs (``operator`` alone was
519,765 bytes in run 45385819819; NCAA_W's total hit 1,207k and the summary was
lost).  This renderer emits the same summary but, when the text exceeds
``MAX_CHARS``, replaces the largest embedded sections with explicit omission
markers so the summary is always delivered.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

#: Stay under GitHub's 1024k upload cap with room for the fenced wrapper.
MAX_CHARS = 900_000


def _len(value) -> int:  # type: ignore[no-untyped-def]
    try:
        return len(json.dumps(value, default=str))
    except (TypeError, ValueError):
        return len(str(value))


def _trim(summary: dict) -> dict:
    """Replace the largest top-level sections until the summary fits."""
    summary = dict(summary)
    while True:
        text = json.dumps(summary, indent=2, default=str)
        if len(text) <= MAX_CHARS:
            return summary
        biggest = max(summary, key=lambda key: _len(summary[key]))
        omitted = (
            f"<omitted: {_len(summary[biggest])} bytes; "
            "summary exceeded GitHub's 1024k cap>"
        )
        if _len(omitted) >= _len(summary[biggest]):
            # Nothing left to trim that would shrink the output.
            return summary
        summary[biggest] = omitted


def render(data: dict) -> str:
    """Build the markdown summary body for one lane's daily report."""
    markets = data.get("markets") or {}
    walk_forward = markets.get("walk_forward") or {}
    accas = markets.get("accas") or {}
    price_coverage: dict = {}
    for forecast in markets.get("forecasts") or []:
        book = forecast.get("book") or {}
        price = book.get("selected_decimal_price")
        exact_price = (
            isinstance(price, (int, float)) and not isinstance(price, bool) and price > 1.0
        )
        state = (
            "exact_selected_price"
            if exact_price
            else "matched_selected_price_missing"
            if book
            else "unmatched_market"
        )
        market_key = f"{forecast.get('market')}/{forecast.get('period')}"
        price_coverage.setdefault(market_key, {}).setdefault(state, 0)
        price_coverage[market_key][state] += 1
    pregame = (data.get("commands") or {}).get("pregame") or {}
    summary = {
        key: data.get(key)
        for key in (
            "status",
            "reason",
            "capabilities",
            "run_id",
            "target_date",
            "league",
            "mode",
            "verification",
            "system_of_record",
            "counts",
            "ledgers",
            "notification",
            "report_path",
            "market_report_path",
            "operator",
        )
        if key in data
    }
    summary["markets"] = {
        "forecast_count": markets.get("forecast_count"),
        "graded_walk_forward_forecasts": walk_forward.get("graded_forecasts"),
        "acca_count": len(accas.get("results") or []),
        "price_coverage_by_market_period": price_coverage,
        "book_provider_status": pregame.get("book_provider_status") or {},
    }
    heading = f"## {summary['league']}\n" if summary.get("league") else ""
    return heading + "```json\n" + json.dumps(_trim(summary), indent=2, default=str) + "\n```"


def main() -> int:
    try:
        data = json.loads(Path("daily-output.json").read_text())
    except (OSError, ValueError):
        print("No valid daily report was produced; inspect the job log.")
        return 0
    print(render(data))
    return 0


if __name__ == "__main__":
    sys.exit(main())
