"""Edge-style operator pick cards, paper tickets, settlement, and audits.

This module is a presentation/settlement layer only. It does not select a
winner, alter a probability, or place a bet. Current forecasts come from
``MarketForecastService``; outcomes come only from reconciled official finals.

Durable contracts (all beneath one league's localdata directory):
- ``picks_YYYY-MM-DD.json``: cumulative pick evidence keyed by forecast id.
- ``picks_YYYY-MM-DD.txt``: latest logical picks, with started/settled archive sections.
- ``picks_today.json``: the current scan, not a replacement for dated history.
- ``picks_audit_rolling.json`` / ``picks_audit_YYYY-MM-DD.md``: rolling grades.
- ``auto_tickets_YYYY-MM-DD_decision.json``: immutable paper ticket receipt.
- ``auto_tickets_YYYY-MM-DD.txt``: receipt rendered with current leg grades.
- ``auto_tickets_state.json``: write-once date locks and bounded scan telemetry.
- ``auto_tickets_performance.json`` / ``.txt``: reconstructed paper performance.

A known losing leg settles its acca as a loss even when other legs are still
pending. Push legs are removed from the effective price once every other leg
has resolved. Pending tickets never count as wins or losses. Empty scans never
cancel or rotate a frozen receipt. All stakes are flat one paper unit per acca.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime, timedelta, timezone
import json
import math
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from zoneinfo import ZoneInfo

from ingestion.identity import parse_utc
from ingestion.models import Game, GameStatus, League, Period

from .types import BookEvidence, Forecast, _period_scores, grade_forecast, opposite_selection

OPERATOR_SCHEMA = "magolide.operator.v1"
FADE_AUDIT_SCHEMA = "magolide.paper_fade_audit.v1"
TICKET_SCHEMA = "magolide.paper_ticket_receipt.v1"
STATE_SCHEMA = "magolide.paper_ticket_state.v1"
BASE_PAPER_UNITS = 100.0
STAKE_UNITS_PER_ACCA = 1.0
BUCKET_ORDER = [
    "ELIGIBLE_CLEAN",
    "CAUTION",
    "WATCHLIST_NO_ODDS",
    "WATCHLIST_INSUFFICIENT_CONTEXT",
    "SKIPPED_INELIGIBLE",
]
BUCKET_LABELS = {
    "ELIGIBLE_CLEAN": "ELIGIBLE CLEAN PICKS",
    "CAUTION": "CAUTION PICKS",
    "WATCHLIST_NO_ODDS": "WATCHLIST — NO MATCHED ODDS",
    "WATCHLIST_INSUFFICIENT_CONTEXT": "WATCHLIST — INSUFFICIENT CONTEXT",
    "SKIPPED_INELIGIBLE": "SKIPPED — INELIGIBLE",
}
UTC = timezone.utc


def _atomic_bytes(path: Path, raw: bytes) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.read_bytes() == raw:
        return str(path)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, path)
    return str(path)


def _write_json(path: Path, value: Any) -> str:
    raw = (json.dumps(value, indent=2, sort_keys=True, default=str) + "\n").encode()
    return _atomic_bytes(path, raw)


def _write_text(path: Path, text: str) -> str:
    return _atomic_bytes(path, text.encode())


def _withdrawn_sidecar_path(localdata: Path, target_date: str) -> Path:
    return localdata / f"auto_tickets_{target_date}_withdrawn.json"


def _withdrawn_ids_from_sidecars(localdata: Path) -> set[str]:
    ids: set[str] = set()
    for path in localdata.glob("auto_tickets_*_withdrawn.json"):
        try:
            data = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        did = data.get("decision_id")
        if did:
            ids.add(str(did))
    return ids


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return deepcopy(default)
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as exc:
        # Financial/evidence state must never silently reset to a fresh bank or
        # an empty card. Corruption is an operator-visible failure.
        raise ValueError(f"cannot parse operator state {path}: {exc}") from exc


def _local_time(value: str, timezone_name: str) -> str:
    return parse_utc(value).astimezone(ZoneInfo(timezone_name)).strftime("%Y-%m-%d %H:%M %Z")


def _bucket(forecast: Dict[str, Any]) -> str:
    book = forecast.get("book") or {}
    price = book.get("selected_decimal_price")
    eligibility = forecast.get("eligibility") or {}
    reasons = set(eligibility.get("reasons") or [])
    if eligibility.get("acca_eligible"):
        return "ELIGIBLE_CLEAN"
    if price is None:
        return "WATCHLIST_NO_ODDS"
    if reasons == {"insufficient_walk_forward_samples"}:
        return "CAUTION"
    if "insufficient_walk_forward_samples" in reasons and not (
        reasons - {"insufficient_walk_forward_samples", "below_calibrated_probability_floor"}
    ):
        return "WATCHLIST_INSUFFICIENT_CONTEXT"
    return "SKIPPED_INELIGIBLE"


def _price_evidence(forecast: Dict[str, Any]) -> str:
    book = forecast.get("book") or {}
    if book.get("selected_decimal_price") is not None:
        return "NAMED_BOOKMAKER_PRICE"
    if book:
        return "MATCHED_BUT_SELECTED_PRICE_MISSING"
    return "UNMATCHED"


def _opposite_selection_for_row(row: Dict[str, Any]) -> Optional[str]:
    book = row.get("book")
    side_prices = book.get("side_prices") if isinstance(book, dict) else None
    has_draw = isinstance(side_prices, dict) and side_prices.get("draw") is not None
    return opposite_selection(
        str(row.get("market") or ""),
        str(row.get("period") or ""),
        str(row.get("selection") or ""),
        has_draw=has_draw,
    )


def _same_line(left: Any, right: Any) -> bool:
    if left is None or right is None:
        return left is right
    if isinstance(left, bool) or isinstance(right, bool):
        return False
    try:
        left_value, right_value = float(left), float(right)
    except (OverflowError, TypeError, ValueError):
        return False
    return (
        math.isfinite(left_value)
        and math.isfinite(right_value)
        and math.isclose(left_value, right_value, rel_tol=0.0, abs_tol=1e-9)
    )


def opposite_quote(row: Dict[str, Any]) -> tuple[Optional[float], Optional[str]]:
    """Return the exact opposite-side quote or a reason it is unavailable.

    The quote must be present in this row's preserved BookSnapshot evidence,
    whose game, market, period, line, bookmaker, and observation are checked.
    No odds or missing market sides are inferred.
    """
    selection = _opposite_selection_for_row(row)
    if selection is None:
        return None, "market_has_no_single_opposite_side"

    book = row.get("book")
    if not isinstance(book, dict):
        return None, "no_book_evidence"
    side_prices = book.get("side_prices")
    if not isinstance(side_prices, dict):
        return None, "opposite_quote_not_preserved"
    if selection not in side_prices or side_prices[selection] is None:
        return None, "opposite_quote_missing_in_payload"

    selected_price = book.get("selected_decimal_price")
    selected_side = str(row.get("selection") or "")
    if selected_price is not None:
        preserved_selected = side_prices.get(selected_side)
        if (
            isinstance(selected_price, bool)
            or not isinstance(selected_price, (int, float))
            or isinstance(preserved_selected, bool)
            or not isinstance(preserved_selected, (int, float))
        ):
            return None, "selected_quote_inconsistent"
        try:
            selected_value = float(selected_price)
            preserved_value = float(preserved_selected)
        except OverflowError:
            return None, "selected_quote_inconsistent"
        if (
            not math.isfinite(selected_value)
            or not math.isfinite(preserved_value)
            or not math.isclose(selected_value, preserved_value, rel_tol=0.0, abs_tol=1e-9)
        ):
            return None, "selected_quote_inconsistent"

    if str(row.get("market") or "") in {"spread", "total"}:
        if row.get("line") is None:
            return None, "market_line_missing"
        if not _same_line(row.get("line"), row.get("line")):
            return None, "invalid_market_line"

    if (
        not row.get("game_id")
        or book.get("game_id") != row.get("game_id")
        or book.get("market") != row.get("market")
        or book.get("period") != row.get("period")
        or not _same_line(book.get("line"), row.get("line"))
    ):
        return None, "opposite_quote_scope_mismatch"
    if book.get("pre_match") is not True:
        return None, "not_pre_match"
    if not str(book.get("provider") or "").strip() or not str(book.get("book") or "").strip():
        return None, "unnamed_book_quote"
    try:
        observed = parse_utc(str(book["observed_utc"]))
        scheduled = parse_utc(str(row["scheduled_utc"]))
    except (KeyError, TypeError, ValueError):
        return None, "missing_quote_timestamp"
    if observed >= scheduled:
        return None, "not_pre_tip"

    raw_price = side_prices[selection]
    if isinstance(raw_price, bool) or not isinstance(raw_price, (int, float)):
        return None, "invalid_opposite_quote"
    try:
        price = float(raw_price)
    except OverflowError:
        return None, "invalid_opposite_quote"
    if not math.isfinite(price) or price <= 1.0:
        return None, "invalid_opposite_quote"
    return price, None


def _confidence_probability(row: Dict[str, Any]) -> tuple[Optional[float], str]:
    value = row.get("calibrated_probability")
    basis = "calibrated_probability"
    if value is None:
        value = row.get("raw_probability")
        basis = "raw_probability"
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None, basis
    try:
        probability = float(value)
    except OverflowError:
        return None, basis
    if not math.isfinite(probability) or not 0.0 <= probability <= 1.0:
        return None, basis
    return probability, basis


def _confidence_band(row: Dict[str, Any]) -> str:
    probability, _ = _confidence_probability(row)
    if probability is None:
        return "unknown"
    if probability < 0.5:
        return "<0.50"
    if probability < 0.6:
        return "0.50-<0.60"
    if probability < 0.7:
        return "0.60-<0.70"
    if probability < 0.8:
        return "0.70-<0.80"
    if probability < 0.9:
        return "0.80-<0.90"
    return "0.90-1.00"


def _fade_metrics(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Describe counterfactual fade results using exact opposite-side quotes."""
    missing_reasons: Counter[str] = Counter()
    wins = losses = pushes = priced_sample = 0
    profit = 0.0
    implied_break_even_sum = 0.0
    decision_quotes = 0
    quoted_prices: List[float] = []

    for row in rows:
        price, reason = opposite_quote(row)
        if price is None:
            missing_reasons[reason or "opposite_quote_unavailable"] += 1
            continue

        fade_grade = {"win": "loss", "loss": "win", "push": "push"}.get(
            str(row.get("grade_status"))
        )
        if fade_grade is None:
            missing_reasons["pick_not_settled"] += 1
            continue

        priced_sample += 1
        quoted_prices.append(price)
        if fade_grade == "win":
            wins += 1
            profit += price - 1.0
        elif fade_grade == "loss":
            losses += 1
            profit -= 1.0
        else:
            pushes += 1
        if fade_grade != "push":
            decision_quotes += 1
            implied_break_even_sum += 1.0 / price

    decisions = wins + losses
    return {
        "candidate_sample": len(rows),
        "priced_sample": priced_sample,
        "unpriced_sample": len(rows) - priced_sample,
        "unpriced_reasons": dict(sorted(missing_reasons.items())),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "decisions": decisions,
        "hit_rate": wins / decisions if decisions else None,
        "profit_units": profit,
        "roi": profit / priced_sample if priced_sample else None,
        "mean_opposite_decimal_price": (
            sum(quoted_prices) / len(quoted_prices) if quoted_prices else None
        ),
        "break_even_hit_rate": (
            implied_break_even_sum / decision_quotes if decision_quotes else None
        ),
        "break_even_sample": decision_quotes,
    }


def _fade_breakdown(rows: List[Dict[str, Any]], key_for_row) -> Dict[str, Any]:
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[str(key_for_row(row))].append(row)
    return {key: _fade_metrics(grouped[key]) for key in sorted(grouped)}


def _pick_row(forecast: Dict[str, Any], target_date: str) -> Dict[str, Any]:
    row = deepcopy(forecast)
    row["pick_id"] = row["forecast_id"]
    row["date"] = target_date
    row["bucket"] = _bucket(row)
    row["price_evidence"] = _price_evidence(row)
    row["grade_status"] = "pending"
    row["outcome"] = None
    row["paper_only"] = True
    return row


def _forecast_from_row(row: Dict[str, Any]) -> Forecast:
    book_row = row.get("book")
    book = None
    if book_row:
        stored_prices = book_row.get("side_prices")
        if isinstance(stored_prices, dict):
            prices = {str(key): value for key, value in stored_prices.items()}
            prices.setdefault(
                str(row["selection"]), book_row.get("selected_decimal_price")
            )
        else:
            # Legacy rows retain only the selected quote; never backfill sides.
            prices = {str(row["selection"]): book_row.get("selected_decimal_price")}
        book = BookEvidence(
            provider=str(book_row.get("provider") or "unknown"),
            book=str(book_row.get("book") or "unknown"),
            observed_utc=str(book_row["observed_utc"]),
            source_modified_utc=book_row.get("source_modified_utc"),
            pre_match=bool(book_row.get("pre_match", True)),
            main_line=bool(book_row.get("main_line", True)),
            frozen=bool(book_row.get("frozen", False)),
            prices=prices,
            game_id=book_row.get("game_id"),
            market=book_row.get("market"),
            period=book_row.get("period"),
            line=book_row.get("line"),
        )
    prediction = row.get("prediction") or {}
    explanation = row.get("explanation") or {}
    shadow_explanation = explanation.get("shadow_calibration") or {}
    shadow_calibration = dict(shadow_explanation) if isinstance(shadow_explanation, dict) else {}
    if shadow_calibration:
        candidate_probability = shadow_calibration.pop("candidate_probability", None)
        shadow_calibration.setdefault("probability", candidate_probability)
        if "ess" not in shadow_calibration and "effective_sample_size" in explanation:
            shadow_calibration["ess"] = explanation["effective_sample_size"]
    return Forecast(
        game_id=str(row["game_id"]),
        league=str(row["league"]),
        scheduled_utc=str(row["scheduled_utc"]),
        home_name=str(row["home"]),
        away_name=str(row["away"]),
        market=str(row["market"]),
        period=str(row["period"]),
        selection=str(row["selection"]),
        raw_probability=float(row["raw_probability"]),
        alternative_probabilities=dict(row.get("alternative_probabilities") or {}),
        source=str(row.get("model") or "unknown"),
        line=row.get("line"),
        push_probability=float(row.get("push_probability") or 0.0),
        expected_margin=row.get("expected_margin"),
        expected_total=row.get("expected_total"),
        training_games=int(row.get("training_games") or 0),
        residual_games=int(row.get("residual_games") or 0),
        team_history=(dict(row["team_history"]) if row.get("team_history") is not None else None),
        book=book,
        prediction_provider=prediction.get("provider"),
        prediction_observed_utc=prediction.get("observed_utc"),
        prediction_frozen=bool(prediction.get("frozen", False)),
        prediction_model=prediction.get("model"),
        prediction_model_version=prediction.get("model_version"),
        prediction_produced_by=prediction.get("produced_by"),
        role=row.get("role"),
        outcome=row.get("outcome"),
        calibrated_probability=row.get("calibrated_probability"),
        calibration_version=row.get("calibration_version"),
        shadow_calibration=shadow_calibration,
        calibration=dict(row.get("calibration") or {}),
        eligibility=dict(row.get("eligibility") or {}),
    )


def _grade_row(row: Dict[str, Any], games: Dict[str, Game]) -> Optional[str]:
    game = games.get(str(row["game_id"]))
    if game is None or game.status is not GameStatus.FINAL:
        return None
    return grade_forecast(_forecast_from_row(row), game)


def _set_grade(row: Dict[str, Any], games: Dict[str, Game]) -> None:
    outcome = _grade_row(row, games)
    row["outcome"] = outcome
    row["grade_status"] = outcome or "pending"


def _spread_line_for_selection(
    row: Dict[str, Any], selection: Optional[str] = None
) -> Optional[float]:
    """Convert a canonical home-oriented spread line to the selected side."""
    line = row.get("line")
    side = str(
        selection if selection is not None else row.get("selection") or ""
    ).lower()
    if (
        side not in {"home", "away"}
        or isinstance(line, bool)
        or not isinstance(line, (int, float))
    ):
        return None
    value = float(line)
    if not math.isfinite(value):
        return None
    return value if side == "home" else -value


def _selection_label(row: Dict[str, Any]) -> str:
    market = str(row.get("market") or "?").replace("_", " ").upper()
    period = str(row.get("period") or "?")
    selection = str(row.get("selection") or "?").upper()
    line = row.get("line")
    if market == "SPREAD":
        selected_line = _spread_line_for_selection(row)
        line_text = "" if selected_line is None else f" {selected_line:+g}"
    else:
        line_text = "" if line is None else f" {float(line):g}"
    return f"{market} {period}{line_text} -> {selection}"


def _canonical_book_for_key(book_name: str | None) -> str | None:
    """Canonicalize book names stored in historical pick rows.

    Older rows used ``Draft Kings`` (space) while newer rows use
    ``DraftKings``. Without normalization the same market would be counted
    as two logical picks, inflating archived/settled totals and ROI.
    This is a presentation-layer fix; canonical snapshots are already
    normalized at ingestion time.
    """
    if not book_name:
        return book_name
    # Local import to avoid circular dependency at module import time.
    try:
        from ingestion.identity import canonical_book_name

        canonical = canonical_book_name(book_name)
        return canonical or book_name
    except Exception:
        return book_name


def _logical_pick_key(row: Dict[str, Any]) -> tuple:
    """Identify repeated observations of one displayed pick, not its evidence ID."""

    book = row.get("book") or {}
    line = row.get("line")
    return (
        row.get("game_id"),
        row.get("market"),
        row.get("period"),
        row.get("selection"),
        float(line) if line is not None else None,
        row.get("model"),
        book.get("provider"),
        _canonical_book_for_key(book.get("book")),
    )


def _evidence_time(row: Dict[str, Any]) -> datetime:
    timestamps = [
        (row.get("book") or {}).get("observed_utc"),
        (row.get("prediction") or {}).get("observed_utc"),
    ]
    observed = [parse_utc(value) for value in timestamps if value]
    return max(observed, default=parse_utc(row["scheduled_utc"]))


def _latest_evidence_utc(row: Dict[str, Any]) -> Optional[str]:
    timestamps = [
        value
        for value in (
            (row.get("book") or {}).get("observed_utc"),
            (row.get("prediction") or {}).get("observed_utc"),
        )
        if value
    ]
    return max(timestamps, key=parse_utc) if timestamps else None


def _latest_logical_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Select the latest observation per logical pick without deleting evidence."""

    groups: Dict[tuple, List[Dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[_logical_pick_key(row)].append(row)
    selected = [
        max(group, key=lambda row: (_evidence_time(row), str(row["pick_id"])))
        for group in groups.values()
    ]
    return sorted(
        selected,
        key=lambda row: (
            row["scheduled_utc"],
            row["game_id"],
            row["market"],
            row["period"],
            row["selection"],
            str(row["pick_id"]),
        ),
    )


def _render_picks(
    rows: List[Dict[str, Any]],
    *,
    league: str,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
) -> str:
    logical_rows = _latest_logical_rows(rows)
    revision_counts = Counter(_logical_pick_key(row) for row in rows)
    generated = parse_utc(generated_utc)
    settled = [row for row in logical_rows if row["grade_status"] in {"win", "loss", "push"}]
    pending = [row for row in logical_rows if row["grade_status"] == "pending"]
    upcoming = [row for row in pending if parse_utc(row["scheduled_utc"]) > generated]
    started = [row for row in pending if parse_utc(row["scheduled_utc"]) <= generated]

    lines = [
        f"Ma Golide Picks — {league} {target_date}",
        "=" * 72,
        f"Generated at: {_local_time(generated_utc, timezone_name)}",
        "PAPER ONLY · eligibility labels are evidence states, not profit guarantees.",
        "Active sections show the latest pending pre-tip observation per logical pick.",
        "Started and settled games are retained separately as archive status.",
        "",
    ]
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for row in upcoming:
        grouped[str(row.get("bucket") or "UNKNOWN")].append(row)
    order = [*BUCKET_ORDER, *sorted(set(grouped) - set(BUCKET_ORDER))]

    def append_pick(row: Dict[str, Any]) -> None:
        bucket = str(row.get("bucket") or "UNKNOWN")
        probability = row.get("calibrated_probability")
        if probability is None:
            probability = row["raw_probability"]
        price = (row.get("book") or {}).get("selected_decimal_price")
        price_text = f"@{float(price):.3f}" if price is not None else "@n/a"
        book_name = (
            _canonical_book_for_key((row.get("book") or {}).get("book")) or "no matched book"
        )
        lines.append(
            f"  [{row.get('model')}/{row.get('model_version')}] "
            f"{row['home']} vs {row['away']} · "
            f"KO {_local_time(row['scheduled_utc'], timezone_name)}"
        )
        lines.append(
            f"      {_selection_label(row)} · p={probability:.1%} "
            f"(raw {row['raw_probability']:.1%}) {price_text} {book_name}"
        )
        eligibility = row.get("eligibility") or {}
        reasons = ",".join(eligibility.get("reasons") or []) or "none"
        calibration = row.get("calibration") or {}
        lines.append(
            f"      bucket={bucket} · price_evidence={row['price_evidence']} "
            f"· acca_eligible={bool(eligibility.get('acca_eligible'))} "
            f"· grade={row['grade_status'].upper()}"
        )
        lines.append(
            f"      calibration={calibration.get('status', 'unknown')} "
            f"n={calibration.get('n', 0)} · reasons={reasons} "
            f"· prediction={(row.get('prediction') or {}).get('provider', 'history')}"
        )
        revisions = revision_counts[_logical_pick_key(row)]
        if revisions > 1:
            latest = _latest_evidence_utc(row) or "unknown"
            lines.append(f"      archived_revisions={revisions} · latest_evidence_utc={latest}")

    def sort_picks(picks: List[Dict[str, Any]]) -> None:
        picks.sort(
            key=lambda row: (
                -(row.get("calibrated_probability") or row["raw_probability"]),
                row["home"],
                row["away"],
                row["pick_id"],
            )
        )

    for bucket in order:
        picks = grouped.get(bucket, [])
        lines.extend(["", BUCKET_LABELS.get(bucket, bucket), "=" * 72])
        if not picks:
            lines.append("  (none)")
            continue
        sort_picks(picks)
        for row in picks:
            append_pick(row)

    for heading, picks in (
        ("IN PROGRESS / AWAITING OFFICIAL RESULT — ARCHIVE ONLY", started),
        ("SETTLED PICKS — ARCHIVE ONLY", settled),
    ):
        lines.extend(["", heading, "=" * 72])
        if not picks:
            lines.append("  (none)")
            continue
        sort_picks(picks)
        for row in picks:
            append_pick(row)

    rendered = len(upcoming) + len(started) + len(settled)
    if rendered != len(logical_rows):
        raise RuntimeError(
            f"daily report dropped logical picks: ledger={len(logical_rows)} rendered={rendered}"
        )
    counts = Counter(row["bucket"] for row in upcoming)
    lines.extend(
        [
            "",
            "Active summary: " + " ".join(f"{bucket}={counts.get(bucket, 0)}" for bucket in order),
            f"Archive status: IN_PROGRESS_OR_AWAITING_RESULT={len(started)} SETTLED={len(settled)}",
            f"Logical picks shown: {len(logical_rows)} · "
            f"superseded observations: {len(rows) - len(logical_rows)}",
            f"Total archived pick evidence rows: {len(rows)}",
            "Flat one-paper-unit tickets only. No real-money stakes or automatic promotion.",
            "Bet only what you can afford to lose.",
            "",
        ]
    )
    return "\n".join(lines)


def _dated_pick_files(localdata: Path) -> List[Path]:
    return sorted(localdata.glob("picks_????-??-??.json"))


def _refresh_dated_pick_reports(
    localdata: Path,
    *,
    games: Dict[str, Game],
    league: str,
    generated_utc: str,
    timezone_name: str,
) -> None:
    """Regrade every retained date before rendering or auditing its picks."""

    for path in _dated_pick_files(localdata):
        rows = _read_json(path, [])
        if not isinstance(rows, list):
            raise ValueError(f"dated pick ledger must be a list: {path}")
        for row in rows:
            _set_grade(row, games)
        _write_json(path, rows)
        target_date = path.stem.removeprefix("picks_")
        _write_text(
            path.with_suffix(".txt"),
            _render_picks(
                rows,
                league=league,
                target_date=target_date,
                generated_utc=generated_utc,
                timezone_name=timezone_name,
            ),
        )


def _audit_period_label(period: str) -> str:
    return {
        "full_game": "Full-game",
        "H1": "First-half",
        "H2": "Second-half",
        "Q1": "First-quarter",
        "Q2": "Second-quarter",
        "Q3": "Third-quarter",
        "Q4": "Fourth-quarter",
    }.get(period, period.replace("_", " "))


def _spread_side_label(row: Dict[str, Any], selection: str) -> str:
    side = str(selection or "").lower()
    if side == "home":
        team = row.get("home")
    elif side == "away":
        team = row.get("away")
    else:
        team = None
    line = _spread_line_for_selection(row, side)
    parts = [str(team)] if team else []
    if line is not None:
        parts.append(f"{line:+g}")
    if not parts:
        return side or "?"
    label = " ".join(parts)
    return f"{label} ({side})" if side else label


def _audit_market_selection(row: Dict[str, Any]) -> str:
    period = _audit_period_label(str(row["period"]))
    market = str(row["market"])
    selection = str(row["selection"])
    selected_team = (
        row["home"] if selection == "home" else row["away"] if selection == "away" else None
    )
    if market == "moneyline":
        side = f"{selected_team} ({selection})" if selected_team else selection.title()
        return f"{period} moneyline: {side}"
    if market == "total":
        line = row.get("line")
        line_text = f" {float(line):g}" if line is not None else ""
        return f"{period} total{line_text}: {selection.title()}"
    if market == "spread":
        line = _spread_line_for_selection(row)
        line_text = f" {line:+g}" if line is not None else ""
        side = f"{selected_team} ({selection})" if selected_team else selection.title()
        return f"{period} spread{line_text}: {side}"
    if market == "highest_quarter":
        return f"Highest quarter: {selection.upper()}"
    return f"{period} {market.replace('_', ' ')}: {selection.title()}"


def _audit_actual_result(row: Dict[str, Any], games: Dict[str, Game] | None) -> str | None:
    game = (games or {}).get(str(row["game_id"]))
    if game is None or game.actual is None:
        return None
    try:
        scores = _period_scores(game, Period(str(row["period"])))
    except ValueError:
        return None
    if scores is None:
        return None
    home, away = scores
    period = str(row["period"])
    market = str(row["market"])
    if market == "total":
        return f"{_audit_period_label(period)} total: **{home + away:g}**"
    if market == "spread":
        return f"{_audit_period_label(period)} margin: **{home - away:+g}**"
    if period == Period.FULL_GAME.value:
        # The fixture heading already carries the full-game final score.
        return None
    return f"{_audit_period_label(period)} score: **{home:g}–{away:g}**"


def _render_pick_audit_markdown(
    report: Dict[str, Any],
    *,
    games: Dict[str, Game] | None,
    timezone_name: str,
) -> str:
    """Render the rolling machine audit as a readable report."""
    settled = list(report.get("settled_ledger") or [])
    wins = int(report["wins"])
    losses = int(report["losses"])
    decisions = wins + losses
    priced_sample = int(report["priced_sample"])
    lines = [
        f"# Ma Golide — Recent picks audit ({report['window']['start']} to {report['window']['end']})",
        "",
        f"Generated: {_local_time(report['generated_utc'], timezone_name)} · "
        f"{report['window']['days']}-day rolling window · Paper only",
        "",
        "## Overall",
        "",
        f"- archived logical picks: {report['archived_pick_rows']}",
        f"- dated evidence rows: {report['archived_evidence_rows']} "
        f"({report['superseded_observations']} superseded observations)",
        f"- settled picks: {report['settled_picks']} "
        f"({wins} wins / {losses} losses / {report['pushes']} pushes)",
        (
            f"- hit rate: {report['hit_rate']:.1%} ({wins}/{decisions} win/loss decisions)"
            if decisions
            else "- hit rate: n/a (no win/loss decisions)"
        ),
        f"- pending/unmatched picks: {report['pending_picks']} prior-day · "
        f"{report['pending_picks_total']} total",
        f"- priced picks: {priced_sample}",
    ]
    if priced_sample:
        lines.extend(
            [
                f"- paper P&L: {report['profit_units']:+.3f} units",
                f"- ROI: {report['roi']:+.1%} (P&L / {priced_sample} priced "
                f"{'pick' if priced_sample == 1 else 'picks'})"
                + (" · **provisional — fewer than 30 priced picks**" if priced_sample < 30 else ""),
            ]
        )
    else:
        lines.append("- ROI: n/a (no settled priced picks)")

    lines.extend(
        [
            "",
            "## Settlement policy",
            "",
            "- reconciled official finals settle immediately; pending rows are never losses",
            f"- local-day pending cutoff: {report['same_day_cutoff']}",
            f"- same-day pending rows outside prior-day counters: {report['same_day_pending_picks']}",
            f"- same-day official-final picks included: {report['same_day_settled_picks']}",
            f"- eligible prior-day picks: {report['eligible_prior_picks']}",
            "- latest observation per logical pick; superseded dated evidence remains in JSON",
            "",
            "## Bucket performance",
            "",
        ]
    )
    by_bucket = report.get("by_bucket") or {}
    if by_bucket:
        lines.extend(
            [
                "| Bucket | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for bucket, group in sorted(by_bucket.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if row["bucket"] == bucket
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{bucket}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    # Market breakdown - all leagues, requested for NCAA M/W watchlist analysis
    by_market = report.get("by_market") or {}
    lines.extend(["", "## By market", ""])
    if by_market:
        lines.extend(
            [
                "| Market | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for market, group in sorted(by_market.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if str(row.get("market")) == market
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{market}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    # Over/under + side bias — selection within market, every league
    by_selection = report.get("by_selection") or {}
    lines.extend(["", "## By selection — over/under & side bias", ""])
    if by_selection:
        lines.extend(
            [
                "| Market:Selection | Settled | W–L–P | Hit rate | Priced | ROI |",
                "| --- | ---: | ---: | ---: | ---: | ---: |",
            ]
        )
        for key, group in sorted(by_selection.items()):
            group_settled = int(group["settled"])
            group_wins = int(group["wins"])
            group_pushes = int(group["pushes"])
            group_decisions = group_settled - group_pushes
            group_losses = group_decisions - group_wins
            group_priced = sum(
                (row.get("book") or {}).get("selected_decimal_price") is not None
                for row in settled
                if f"{str(row.get('market'))}:{str(row.get('selection'))}" == key
            )
            hit_rate = f"{group_wins / group_decisions:.1%}" if group_decisions else "n/a"
            roi = f"{group['priced_profit_units'] / group_priced:+.1%}" if group_priced else "n/a"
            lines.append(
                f"| `{key}` | {group_settled} | {group_wins}–{group_losses}–{group_pushes} | "
                f"{hit_rate} | {group_priced} | {roi} |"
            )
    else:
        lines.append("No settled picks in this window.")

    fade = report.get("fade_audit") or {}
    fade_priced = int(fade.get("priced_sample") or 0)
    fade_candidates = int(fade.get("candidate_sample") or 0)
    fade_decisions = int(fade.get("decisions") or 0)
    fade_hit_rate = fade.get("hit_rate")
    fade_roi = fade.get("roi")
    break_even_hit_rate = fade.get("break_even_hit_rate")
    lines.extend(
        [
            "",
            "## Paper-only opposite-side fade audit",
            "",
            "Counterfactual reporting only — this is not a strategy, pick, or ticket. "
            "A settled result is reversed (win↔loss; push stays push) and priced only "
            "when the exact opposite quote is preserved in the same pre-tip book "
            "snapshot, market, period, line, and observation. No quote is inferred "
            "or historically backfilled; selected-side price is not required.",
            "",
            f"- settled rows with exact opposite quotes: {fade_priced}/{fade_candidates}",
            (
                f"- fade record: {fade.get('wins', 0)}–{fade.get('losses', 0)}–"
                f"{fade.get('pushes', 0)}"
            ),
            (
                f"- fade hit rate: {fade_hit_rate:.1%} "
                f"({fade.get('wins', 0)}/{fade_decisions} decisions)"
                if fade_hit_rate is not None
                else "- fade hit rate: n/a (no priced win/loss decisions)"
            ),
            (
                f"- paper P&L: {fade.get('profit_units', 0.0):+.3f} units · "
                f"ROI: {fade_roi:+.1%} (one unit per priced fade, pushes included)"
                if fade_roi is not None
                else "- fade ROI: n/a (no settled exact opposite-side quotes)"
            ),
            (
                f"- mean market-implied break-even hit rate: {break_even_hit_rate:.1%} "
                f"(n={fade.get('break_even_sample', 0)} non-push quotes; includes vig)"
                if break_even_hit_rate is not None
                else "- mean market-implied break-even hit rate: n/a"
            ),
        ]
    )
    fade_reasons = fade.get("unpriced_reasons") or {}
    if fade_reasons:
        lines.append(
            "- unpriced settled rows: "
            + ", ".join(f"`{reason}`={count}" for reason, count in sorted(fade_reasons.items()))
        )
    fade_by_market = fade.get("by_market") or {}
    if fade_by_market:
        lines.extend(
            [
                "",
                "| Market | Settled candidates | Exact quotes | Fade W–L–P | ROI |",
                "| --- | ---: | ---: | ---: | ---: |",
            ]
        )
        for market, group in sorted(fade_by_market.items()):
            roi = f"{group['roi']:+.1%}" if group.get("roi") is not None else "n/a"
            lines.append(
                f"| `{market}` | {group['candidate_sample']} | {group['priced_sample']} | "
                f"{group['wins']}–{group['losses']}–{group['pushes']} | {roi} |"
            )
    fade_by_band = fade.get("by_confidence_band") or {}
    if fade_by_band:
        lines.extend(
            [
                "",
                "Numeric selected-probability bands (calibrated probability, otherwise raw):",
                "",
                "| Probability band | Settled candidates | Exact quotes | Fade W–L–P | ROI |",
                "| --- | ---: | ---: | ---: | ---: |",
            ]
        )
        for band, group in sorted(fade_by_band.items()):
            roi = f"{group['roi']:+.1%}" if group.get("roi") is not None else "n/a"
            lines.append(
                f"| `{band}` | {group['candidate_sample']} | {group['priced_sample']} | "
                f"{group['wins']}–{group['losses']}–{group['pushes']} | {roi} |"
            )

    lines.extend(
        [
            "",
            "## Settled picks — official-final audit",
            "",
            "Grouped by fixture. Every settled logical pick in the window is included; "
            "unpriced grades do not enter ROI.",
            "",
        ]
    )
    if not settled:
        lines.extend(["No settled picks in this window.", ""])

    grouped: Dict[tuple, List[Dict[str, Any]]] = {}
    for row in settled:
        key = (row["date"], row["game_id"], row["home"], row["away"])
        grouped.setdefault(key, []).append(row)
    for (pick_date, game_id, home, away), rows in grouped.items():
        game = (games or {}).get(str(game_id))
        if (
            game is not None
            and game.actual is not None
            and game.actual.home_final is not None
            and game.actual.away_final is not None
        ):
            final = f"Final: **{game.actual.home_final:g}–{game.actual.away_final:g}**"
        else:
            final = "Official final reconciled"
        lines.extend([f"### {pick_date}: {home} vs {away} ({final})", ""])
        for row in rows:
            grade = str(row["grade_status"])
            icon, result = {
                "win": ("🟢", "WON"),
                "loss": ("🔴", "LOST"),
                "push": ("⚪", "PUSH"),
            }[grade]
            lines.append(f"- {icon} **{result} — {_audit_market_selection(row)}**")
            probability = row.get("calibrated_probability")
            if probability is None:
                probability = row["raw_probability"]
            probability_text = f"**{probability:.1%}**"
            if abs(probability - row["raw_probability"]) >= 0.0005:
                probability_text += f" calibrated (raw {row['raw_probability']:.1%})"
            facts = []
            actual = _audit_actual_result(row, games)
            if actual:
                facts.append(f"Actual {actual}")
            facts.extend(
                [
                    f"Expected: {probability_text}",
                ]
            )
            price = (row.get("book") or {}).get("selected_decimal_price")
            book_display = _canonical_book_for_key((row.get("book") or {}).get("book"))
            price_text = (
                f"{float(price):.3f} · {book_display or 'book unknown'}"
                if price is not None
                else "No exact bookmaker price"
            )
            facts.append(f"Price: **{price_text}**")
            opposite_price, opposite_reason = opposite_quote(row)
            opposite_side = _opposite_selection_for_row(row)
            if opposite_price is None:
                fade_text = f"not priced ({opposite_reason})"
            else:
                fade_result = {
                    "win": "LOST",
                    "loss": "WON",
                    "push": "PUSH",
                }[grade]
                fade_side = opposite_side
                if row.get("market") == "spread" and opposite_side in {"home", "away"}:
                    fade_side = _spread_side_label(row, opposite_side)
                fade_text = (
                    f"{fade_result} @ {opposite_price:.3f} {fade_side} · "
                    f"{book_display or 'book unknown'}"
                )
            facts.append(f"Counterfactual fade: **{fade_text}**")
            lines.append("  - " + " · ".join(facts))
            calibration = row.get("calibration") or {}
            lines.append(
                f"  - Calibration: **n={calibration.get('n', 0)}** · "
                f"{str(calibration.get('status', 'unknown')).replace('_', ' ')} · "
                f"Bucket: `{row['bucket']}` · "
                f"Model: `{row.get('model') or 'unknown'}`"
            )
        lines.append("")

    lines.extend(
        [
            "## Evidence notes",
            "",
            "- Grades use reconciled official finals only; missing or conflicting results stay pending.",
            "- Dated JSON preserves every forecast, price, model, calibration, and eligibility field.",
            "- Flat one-paper-unit ROI is descriptive history, not proof of future performance.",
            "",
        ]
    )
    return "\n".join(lines)


def _write_pick_audit(
    localdata: Path,
    *,
    target_date: str,
    generated_utc: str,
    timezone_name: str,
    games: Dict[str, Game] | None = None,
) -> Dict[str, Any]:
    end = datetime.strptime(target_date, "%Y-%m-%d").date()
    start = end - timedelta(days=29)
    ledger: List[Dict[str, Any]] = []
    for path in _dated_pick_files(localdata):
        payload = _read_json(path, [])
        if not isinstance(payload, list):
            raise ValueError(f"dated pick ledger must be a list: {path}")
        ledger.extend(payload)
    window = [row for row in ledger if start.isoformat() <= row["date"] <= target_date]
    logical_window = _latest_logical_rows(window)
    local_cutoff = parse_utc(generated_utc).astimezone(ZoneInfo(timezone_name)).date().isoformat()
    # Match Edge's local-day pending cutoff, not the acquisition/slate date.
    # Unlike Edge's source-score default, our grades require reconciled official
    # finals, so a same-day final is already safe to include in settled metrics.
    prior = [row for row in logical_window if row["date"] < local_cutoff]
    same_day_pending = [
        row
        for row in logical_window
        if row["date"] >= local_cutoff and row["grade_status"] == "pending"
    ]
    settled = sorted(
        (row for row in logical_window if row["grade_status"] in {"win", "loss", "push"}),
        key=lambda row: (
            row["date"],
            row["scheduled_utc"],
            row["game_id"],
            str(row["pick_id"]),
        ),
        reverse=True,
    )
    same_day_settled = sum(row["date"] >= local_cutoff for row in settled)
    decisions = [row for row in settled if row["grade_status"] != "push"]
    wins = sum(row["grade_status"] == "win" for row in decisions)
    pushes = sum(row["grade_status"] == "push" for row in settled)
    pending = sum(row["grade_status"] == "pending" for row in prior)
    pending_total = sum(row["grade_status"] == "pending" for row in logical_window)
    priced = [
        row for row in settled if (row.get("book") or {}).get("selected_decimal_price") is not None
    ]
    profit = 0.0
    for row in priced:
        price = float(row["book"]["selected_decimal_price"])
        if row["grade_status"] == "win":
            profit += price - 1.0
        elif row["grade_status"] == "loss":
            profit -= 1.0
    by_bucket: Dict[str, Dict[str, Any]] = {}
    by_market: Dict[str, Dict[str, Any]] = {}
    by_selection: Dict[str, Dict[str, Any]] = {}
    for row in settled:
        group = by_bucket.setdefault(
            row["bucket"],
            {"settled": 0, "wins": 0, "pushes": 0, "priced_profit_units": 0.0},
        )
        group["settled"] += 1
        group["wins"] += int(row["grade_status"] == "win")
        group["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                group["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                group["priced_profit_units"] -= 1.0
        # Market breakdown (all leagues) - requested for NCAA M/W watchlist analysis
        m = str(row.get("market") or "unknown")
        mgroup = by_market.setdefault(
            m, {"settled": 0, "wins": 0, "pushes": 0, "priced_profit_units": 0.0}
        )
        mgroup["settled"] += 1
        mgroup["wins"] += int(row["grade_status"] == "win")
        mgroup["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                mgroup["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                mgroup["priced_profit_units"] -= 1.0
        # Over/under + side bias audit — selection within market (over/under, home/away) — same regime, no guards lowered
        sel = str(row.get("selection") or "unknown")
        key = f"{m}:{sel}"
        sgroup = by_selection.setdefault(
            key,
            {
                "settled": 0,
                "wins": 0,
                "pushes": 0,
                "priced_profit_units": 0.0,
                "market": m,
                "selection": sel,
            },
        )
        sgroup["settled"] += 1
        sgroup["wins"] += int(row["grade_status"] == "win")
        sgroup["pushes"] += int(row["grade_status"] == "push")
        if (row.get("book") or {}).get("selected_decimal_price") is not None:
            price = float(row["book"]["selected_decimal_price"])
            if row["grade_status"] == "win":
                sgroup["priced_profit_units"] += price - 1.0
            elif row["grade_status"] == "loss":
                sgroup["priced_profit_units"] -= 1.0

    fade_audit = {
        "schema_version": FADE_AUDIT_SCHEMA,
        "paper_only": True,
        "evidence_policy": (
            "use only explicit opposite-side prices preserved in the pick's "
            "same BookSnapshot; never infer or historically backfill quotes"
        ),
        "selected_price_required": False,
        "unit_stake": 1.0,
        "profit_basis": (
            "fade win earns opposite decimal price minus one; fade loss loses one; "
            "push returns stake; ROI divides by one unit per priced fade, including pushes"
        ),
        "break_even_basis": (
            "mean inverse opposite decimal price among non-push fade quotes; "
            "book-implied and includes market margin"
        ),
        "confidence_probability_basis": (
            "numeric calibrated_probability, falling back to raw_probability; "
            "calibration status is not used"
        ),
        **_fade_metrics(settled),
        "by_market": _fade_breakdown(
            settled, lambda row: row.get("market") or "unknown"
        ),
        "by_selection": _fade_breakdown(
            settled,
            lambda row: f"{row.get('market') or 'unknown'}:{row.get('selection') or 'unknown'}",
        ),
        "by_confidence_band": _fade_breakdown(settled, _confidence_band),
        "by_market_period_model_confidence_band": _fade_breakdown(
            settled,
            lambda row: "|".join(
                (
                    str(row.get("market") or "unknown"),
                    str(row.get("period") or "unknown"),
                    str(row.get("model") or "unknown"),
                    _confidence_band(row),
                )
            ),
        ),
    }
    report = {
        "schema_version": OPERATOR_SCHEMA,
        "generated_utc": generated_utc,
        "window": {"start": start.isoformat(), "end": target_date, "days": 30},
        "archived_pick_rows": len(logical_window),
        "archived_evidence_rows": len(window),
        "superseded_observations": len(window) - len(logical_window),
        "same_day_cutoff": local_cutoff,
        "include_same_day_settled": True,
        "same_day_rows_excluded": len(same_day_pending),
        "same_day_pending_picks": len(same_day_pending),
        "same_day_settled_picks": same_day_settled,
        "eligible_prior_picks": len(prior),
        "settled_picks": len(settled),
        "pending_picks": pending,
        "pending_picks_total": pending_total,
        "wins": wins,
        "losses": len(decisions) - wins,
        "pushes": pushes,
        "hit_rate": wins / len(decisions) if decisions else None,
        "priced_sample": len(priced),
        "roi": profit / len(priced) if priced else None,
        "profit_units": profit,
        "stake_policy": "flat_one_paper_unit_per_priced_pick",
        "fade_audit": fade_audit,
        "by_bucket": by_bucket,
        "by_market": by_market,
        "by_selection": by_selection,
        "grading_basis": (
            "reconciled_official_finals_only; official_finals_settle_immediately; "
            "pending_is_not_a_loss; local_day_pending_cutoff; "
            "latest_observation_per_logical_pick"
        ),
        "ledger": logical_window,
        "settled_ledger": settled,
    }
    _write_json(localdata / "picks_audit_rolling.json", report)
    _write_text(
        localdata / f"picks_audit_{target_date}.md",
        _render_pick_audit_markdown(report, games=games, timezone_name=timezone_name),
    )
    return report


def _load_state(localdata: Path) -> Dict[str, Any]:
    state = _read_json(
        localdata / "auto_tickets_state.json",
        {
            "schema_version": STATE_SCHEMA,
            "base_paper_units": BASE_PAPER_UNITS,
            "stake_units_per_acca": STAKE_UNITS_PER_ACCA,
            "frozen_by_date": {},
            "scans": [],
        },
    )
    if not isinstance(state, dict) or state.get("schema_version") != STATE_SCHEMA:
        raise ValueError("unsupported or corrupt auto ticket state; refusing to reset it")
    state.setdefault("frozen_by_date", {})
    state.setdefault("scans", [])
    return state


def _new_receipt(
    *,
    league: League,
    target_date: str,
    captured_utc: str,
    run_id: str,
    accas: List[Dict[str, Any]],
) -> Dict[str, Any]:
    ticket_rows = []
    for index, acca in enumerate(accas, 1):
        legs = [_pick_row(leg, target_date) for leg in acca["legs"]]
        books = sorted({(leg.get("book") or {}).get("book") for leg in legs} - {None})
        ticket_rows.append(
            {
                "ticket_id": acca["acca_id"],
                "publication_index": index,
                "combined_decimal_odds": acca["combined_decimal_odds"],
                "stake_units": STAKE_UNITS_PER_ACCA,
                "estimated_joint_win_probability": acca["estimated_joint_win_probability"],
                "paper_only": True,
                "provisional_sample_gate_bypassed": bool(
                    acca.get("provisional_sample_gate_bypassed", False)
                ),
                "provisional_leg_count": int(acca.get("provisional_leg_count") or 0),
                "ranking_status": acca.get("ranking_status"),
                "bookmakers": books,
                "legs": legs,
            }
        )
    return {
        "schema_version": TICKET_SCHEMA,
        "date": target_date,
        "league": league.value,
        "decision_id": run_id,
        "status": "frozen",
        "paper_only": True,
        "captured_at": captured_utc,
        "frozen_at": captured_utc,
        "stake_units_per_acca": STAKE_UNITS_PER_ACCA,
        "accas": ticket_rows,
    }


def _ticket_receipt(
    localdata: Path,
    *,
    state: Dict[str, Any],
    league: League,
    target_date: str,
    captured_utc: str,
    run_id: str,
    accas: List[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    path = localdata / f"auto_tickets_{target_date}_decision.json"
    # Sidecar is authoritative even if state cache rolls back.
    if _withdrawn_sidecar_path(localdata, target_date).exists():
        return None
    lock = state["frozen_by_date"].get(target_date)
    if isinstance(lock, dict) and lock.get("withdrawn") is True:
        # Keep immutable receipt bytes for evidence, but never adopt, rotate,
        # settle, or replace a withdrawn publication. This also makes the
        # withdrawal survive restoration of an older Actions cache.
        if not lock.get("decision_id"):
            raise ValueError(f"withdrawn ticket lock lacks decision_id: {target_date}")
        # Persist sidecar so future state rollbacks still see the withdrawal.
        sidecar = _withdrawn_sidecar_path(localdata, target_date)
        if not sidecar.exists():
            _write_json(
                sidecar,
                {
                    "decision_id": lock.get("decision_id"),
                    "withdrawn": True,
                    "target_date": target_date,
                },
            )
        return None
    if lock is not None:
        locked_path = localdata / str(lock.get("receipt_path") or path.name)
        if not locked_path.exists():
            raise FileNotFoundError(
                f"frozen ticket receipt is missing: {locked_path}; refusing to rotate the card"
            )
        receipt = _read_json(locked_path, {})
        if receipt.get("schema_version") != TICKET_SCHEMA:
            raise ValueError(f"invalid frozen ticket receipt: {locked_path}")
        return receipt
    if path.exists():
        # Do not adopt a withdrawn receipt even via immutable-receipt path.
        sidecar_ids = _withdrawn_ids_from_sidecars(localdata)
        if sidecar_ids:
            try:
                cand = json.loads(path.read_text())
                if str(cand.get("decision_id")) in sidecar_ids:
                    return None
            except (OSError, json.JSONDecodeError):
                pass
        # A receipt is itself immutable evidence. Adopt its lock rather than
        # rebuilding a potentially different card after a state-cache loss.
        receipt = _read_json(path, {})
        if receipt.get("schema_version") != TICKET_SCHEMA:
            raise ValueError(f"invalid frozen ticket receipt: {path}")
        state["frozen_by_date"][target_date] = {
            "frozen_at": receipt["frozen_at"],
            "final": True,
            "lock_source": "immutable_receipt",
            "receipt_path": path.name,
        }
        return receipt
    if not accas:
        return None
    receipt = _new_receipt(
        league=league,
        target_date=target_date,
        captured_utc=captured_utc,
        run_id=run_id,
        accas=accas,
    )
    _write_json(path, receipt)
    state["frozen_by_date"][target_date] = {
        "frozen_at": captured_utc,
        "final": True,
        "lock_source": "state",
        "receipt_path": path.name,
    }
    return receipt


def _grade_receipt(receipt: Dict[str, Any], games: Dict[str, Game]) -> Dict[str, Any]:
    graded = deepcopy(receipt)
    for acca in graded["accas"]:
        outcomes: List[Optional[str]] = []
        for leg in acca["legs"]:
            _set_grade(leg, games)
            outcomes.append(leg["outcome"])
        acca["leg_outcomes"] = outcomes
        acca["push_legs"] = sum(outcome == "push" for outcome in outcomes)
        acca["unresolved_legs"] = sum(outcome is None for outcome in outcomes)
        stake = float(acca["stake_units"])
        effective_odds = math.prod(
            float((leg.get("book") or {})["selected_decimal_price"])
            for leg in acca["legs"]
            if leg["outcome"] != "push"
        )
        acca["effective_decimal_odds"] = effective_odds
        if any(outcome == "loss" for outcome in outcomes):
            status = "loss"
            profit = -stake
        elif any(outcome is None for outcome in outcomes):
            status = "pending"
            profit = None
        elif outcomes and all(outcome == "push" for outcome in outcomes):
            status = "push"
            profit = 0.0
        else:
            status = "win"
            profit = stake * (effective_odds - 1.0)
        acca["grade_status"] = status
        acca["profit_units"] = profit
    return graded


def _performance(receipts: List[Dict[str, Any]], games: Dict[str, Game]) -> Dict[str, Any]:
    history = [_grade_receipt(receipt, games) for receipt in receipts]
    accas = [acca for receipt in history for acca in receipt["accas"]]
    settled = [acca for acca in accas if acca["profit_units"] is not None]
    pending = [acca for acca in accas if acca["profit_units"] is None]
    profit = sum(float(acca["profit_units"]) for acca in settled)
    staked = sum(float(acca["stake_units"]) for acca in settled)
    committed = sum(float(acca["stake_units"]) for acca in pending)
    wins = sum(acca["grade_status"] == "win" for acca in settled)
    losses = sum(acca["grade_status"] == "loss" for acca in settled)
    pushes = sum(acca["grade_status"] == "push" for acca in settled)
    unresolved_legs = sum(acca["unresolved_legs"] for acca in accas)
    brier = []
    seen_legs = set()
    for receipt in history:
        for acca in receipt["accas"]:
            for leg in acca["legs"]:
                identity = (receipt["date"], leg["pick_id"])
                if identity in seen_legs or leg["outcome"] not in {"win", "loss"}:
                    continue
                seen_legs.add(identity)
                probability = leg.get("calibrated_probability")
                if probability is None:
                    probability = leg["raw_probability"]
                brier.append((probability - float(leg["outcome"] == "win")) ** 2)
    days = []
    equity = BASE_PAPER_UNITS
    for receipt in sorted(history, key=lambda row: row["date"]):
        daily = receipt["accas"]
        daily_settled = [acca for acca in daily if acca["profit_units"] is not None]
        daily_pending = [acca for acca in daily if acca["profit_units"] is None]
        daily_stake = sum(float(acca["stake_units"]) for acca in daily_settled)
        daily_profit = sum(float(acca["profit_units"]) for acca in daily_settled)
        daily_pending_stake = sum(float(acca["stake_units"]) for acca in daily_pending)
        equity += daily_profit
        days.append(
            {
                "date": receipt["date"],
                "league": receipt["league"],
                "decision_id": receipt["decision_id"],
                "frozen_at": receipt["frozen_at"],
                "wins": sum(acca["grade_status"] == "win" for acca in daily_settled),
                "losses": sum(acca["grade_status"] == "loss" for acca in daily_settled),
                "pushes": sum(acca["grade_status"] == "push" for acca in daily_settled),
                "staked_units": daily_stake,
                "returned_units": daily_stake + daily_profit if daily_settled else None,
                "profit_units": daily_profit if daily_settled else None,
                "pending_units": daily_pending_stake,
                "bank_units": equity,
                "accas": daily,
            }
        )
    return {
        "schema_version": OPERATOR_SCHEMA,
        "paper_only": True,
        "base_paper_units": BASE_PAPER_UNITS,
        "paper_bank_units": BASE_PAPER_UNITS + profit,
        "paper_free_units": BASE_PAPER_UNITS + profit - committed,
        "paper_committed_units": committed,
        "settled_stake_units": staked,
        "settled_accas": len(settled),
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "hit_rate": wins / (wins + losses) if wins + losses else None,
        "pending_accas": len(pending),
        "unresolved_result_legs": unresolved_legs,
        "profit_units": profit,
        "roi": profit / staked if staked else None,
        "selected_leg_brier_score": sum(brier) / len(brier) if brier else None,
        "stake_policy": "flat_one_paper_unit_per_acca_no_refill_no_real_money",
        "drawdown_basis": "not_a_settlement_time_equity_curve",
        "history": days,
    }


def _render_ticket_card(
    receipt: Dict[str, Any], performance: Dict[str, Any], timezone_name: str
) -> str:
    lines = [
        f"MA GOLIDE PAPER TICKETS — {receipt['league']} {receipt['date']}",
        "=" * 72,
        f"CURRENT PAPER BANK (all dates): {performance['paper_bank_units']:.3f} units "
        f"· free {performance['paper_free_units']:.3f} / reserved {performance['paper_committed_units']:.3f}",
        f"{len(receipt['accas'])} paper accas · 1 ticket publication · times in {timezone_name}",
        "PAPER ONLY — flat one unit per acca; no automatic betting writes.",
        "Frozen tickets retain their original decision basis; empty scans do not cancel them.",
        "",
    ]
    if not receipt["accas"]:
        lines.extend(["NO PAPER BET — no strict or provisional paper acca at publication.", ""])
    for index, acca in enumerate(receipt["accas"], 1):
        books = ", ".join(acca.get("bookmakers") or []) or "multiple/no book"
        provisional = bool(acca.get("provisional_sample_gate_bypassed"))
        provisional_label = " · PROVISIONAL SAMPLE GATE" if provisional else ""
        lines.append(
            f"[PAPER ACCA #{index}]{provisional_label} @{acca['combined_decimal_odds']:.3f} "
            f"— {acca['stake_units']:.6f} paper units — {books} "
            f"· {acca['grade_status'].upper()}"
        )
        lines.append(
            f"   Published {_local_time(receipt['frozen_at'], timezone_name)} "
            f"· reference P1/{index} · joint p={acca['estimated_joint_win_probability']:.1%}"
        )
        if provisional:
            provisional_legs = [
                leg for leg in acca["legs"] if _is_provisional_sample_only_row(leg)
            ]
            samples = ", ".join(
                f"{leg.get('game_id', '?')} n={(leg.get('calibration') or {}).get('n', '?')}"
                f"/{(leg.get('calibration') or {}).get('minimum_samples', '?')}"
                for leg in provisional_legs
            )
            detail = f" · {samples}" if samples else ""
            lines.append(
                "   PROVISIONAL PAPER ONLY — sample-size gate bypassed; all other "
                f"eligibility gates remain enforced{detail}"
            )
        for leg in acca["legs"]:
            price = (leg.get("book") or {}).get("selected_decimal_price")
            probability = leg.get("calibrated_probability")
            if probability is None:
                probability = leg["raw_probability"]
            lines.extend(
                [
                    f"   {leg['home']} vs {leg['away']} · "
                    f"TIP-OFF {_local_time(leg['scheduled_utc'], timezone_name)}",
                    f"      {_selection_label(leg)} @{float(price):.3f} "
                    f"· ML {probability:.1%} · {leg['bucket']} · {leg['grade_status'].upper()}",
                ]
            )
        lines.append("")
    lines.extend(
        [
            f"FROZEN AT {_local_time(receipt['frozen_at'], timezone_name)[11:16]} - FINAL",
            "Existing tickets are unchanged by later scans or corrected forecasts.",
            "Paper only. No real-money stakes or automatic promotion.",
            "",
        ]
    )
    return "\n".join(lines)


def _is_provisional_sample_only_row(row: Dict[str, Any]) -> bool:
    eligibility = row.get("eligibility") or {}
    if eligibility.get("acca_eligible") is not False:
        return False
    reasons = eligibility.get("reasons")
    if not isinstance(reasons, (list, tuple, set)):
        return False
    if set(reasons) != {"insufficient_walk_forward_samples"}:
        return False
    calibration = row.get("calibration") or {}
    if (
        calibration.get("status") != "provisional"
        or calibration.get("sufficient_for_accas") is not False
    ):
        return False
    n = calibration.get("n")
    if isinstance(n, bool) or not isinstance(n, int) or n < 1:
        return False
    minimum = calibration.get("minimum_samples")
    if minimum is not None and (
        isinstance(minimum, bool)
        or not isinstance(minimum, int)
        or minimum < 1
        or n >= minimum
    ):
        return False
    return True


def _ticket_candidate_funnel(market_report: Dict[str, Any] | None) -> Dict[str, Any]:
    """Explain strict and paper-exploratory supply without changing the report."""
    report = market_report or {}
    rows = list(report.get("forecasts") or [])
    acca_report = report.get("accas") or {}

    def exact_price(row: Dict[str, Any]) -> bool:
        price = (row.get("book") or {}).get("selected_decimal_price")
        return (
            isinstance(price, (int, float))
            and not isinstance(price, bool)
            and math.isfinite(price)
            and price > 1.0
        )

    priced = [row for row in rows if exact_price(row)]
    calibrated = [
        row for row in rows if (row.get("calibration") or {}).get("sufficient_for_accas") is True
    ]
    eligible = [row for row in rows if (row.get("eligibility") or {}).get("acca_eligible")]
    provisional = [
        row for row in rows if exact_price(row) and _is_provisional_sample_only_row(row)
    ]
    eligible_games = {row.get("game_id") for row in eligible}
    provisional_games = {row.get("game_id") for row in provisional}
    paper_results = acca_report.get("paper_results")
    if paper_results is None:  # legacy reports expose only the strict result surface
        paper_results = acca_report.get("results")
    paper_accas = list(paper_results or [])
    strict_accas = acca_report.get("strict_results")
    if strict_accas is None:  # pre-exploration report compatibility
        strict_accas = list(acca_report.get("results") or [])
    else:
        strict_accas = list(strict_accas or [])
    blockers = Counter(
        reason for row in rows for reason in ((row.get("eligibility") or {}).get("reasons") or [])
    )
    try:
        required_games = int(acca_report.get("requested_legs") or 3)
    except (TypeError, ValueError):
        required_games = 3
    return {
        "scan_rows": len(rows),
        "exact_priced_rows": len(priced),
        "calibration_sufficient_rows": len(calibrated),
        "priced_and_calibrated_rows": sum(
            exact_price(row) and (row.get("calibration") or {}).get("sufficient_for_accas") is True
            for row in rows
        ),
        "eligible_rows": len(eligible),
        "eligible_games": len(eligible_games),
        "provisional_sample_only_rows": len(provisional),
        "provisional_sample_only_games": len(provisional_games),
        "paper_candidate_rows": len(eligible) + len(provisional),
        "paper_candidate_games": len(eligible_games | provisional_games),
        "paper_accas": len(paper_accas),
        "strict_accas": len(strict_accas),
        "required_games": required_games,
        "eligibility_blockers": dict(sorted(blockers.items())),
    }


def _render_empty_ticket_scan(
    *,
    league: League,
    target_date: str,
    generated_utc: str,
    performance: Dict[str, Any],
    timezone_name: str,
    candidate_funnel: Dict[str, Any],
    withdrawal: Optional[Dict[str, Any]] = None,
) -> str:
    """A mutable latest-scan view, never a frozen no-bet receipt."""
    lines = [
        f"MA GOLIDE PAPER TICKETS — {league.value} {target_date}",
        "=" * 72,
        f"CURRENT PAPER BANK (all dates): {performance['paper_bank_units']:.3f} units "
        f"· free {performance['paper_free_units']:.3f} / reserved {performance['paper_committed_units']:.3f}",
        f"Latest scan: {_local_time(generated_utc, timezone_name)}",
        "",
        (
            "NO PAPER BET — the earlier ticket publication was withdrawn for model-integrity failure."
            if withdrawal
            else "NO PAPER BET — no strict or provisional paper acca in this scan."
        ),
        "",
        "CANDIDATE FUNNEL",
        f"  {candidate_funnel['scan_rows']} scan rows · "
        f"{candidate_funnel['exact_priced_rows']} exact priced · "
        f"{candidate_funnel['calibration_sufficient_rows']} calibration-sufficient",
        f"  {candidate_funnel['priced_and_calibrated_rows']} priced + calibrated · "
        f"{candidate_funnel['eligible_rows']} eligible "
        f"{'row' if candidate_funnel['eligible_rows'] == 1 else 'rows'} · "
        f"{candidate_funnel['eligible_games']} eligible "
        f"{'game' if candidate_funnel['eligible_games'] == 1 else 'games'}",
        f"  {candidate_funnel['provisional_sample_only_rows']} provisional sample-only "
        f"{'row' if candidate_funnel['provisional_sample_only_rows'] == 1 else 'rows'} · "
        f"{candidate_funnel['provisional_sample_only_games']} "
        f"{'game' if candidate_funnel['provisional_sample_only_games'] == 1 else 'games'} · "
        f"{candidate_funnel['paper_accas']} paper accas "
        f"({candidate_funnel['strict_accas']} strict)",
        f"  Need {candidate_funnel['required_games']} independent games for one acca; "
        "same-game legs are prohibited.",
    ]
    blockers = candidate_funnel.get("eligibility_blockers") or {}
    if blockers:
        lines.append("  Blockers (a row can have more than one):")
        for reason, count in sorted(blockers.items(), key=lambda item: (-item[1], item[0])):
            lines.append(f"    {count:>3} {reason}")
    elif not candidate_funnel["scan_rows"]:
        lines.append("  No current forecasts were generated.")
    lines.extend(
        [
            "",
            (
                "This date is closed to replacement ticket publication after the withdrawal."
                if withdrawal
                else "This scan card is not a frozen receipt; a later qualifying scan may publish tickets."
            ),
            (
                "The withdrawn receipt is excluded from performance and retained unchanged as evidence."
                if withdrawal
                else "Empty scans never cancel or resize earlier frozen positions."
            ),
            "Paper only. No real-money stakes or automatic promotion.",
            "",
        ]
    )
    return "\n".join(lines)


def _render_performance(report: Dict[str, Any], timezone_name: str) -> str:
    lines = [
        "MA GOLIDE AUTO-TICKETS (ROLLING) — paper performance",
        "=" * 72,
        f"BANK     {report['paper_bank_units']:.3f} units   "
        f"x{report['paper_bank_units'] / BASE_PAPER_UNITS:.3f} of the {BASE_PAPER_UNITS:.0f}-unit base",
        f"CAPITAL  {report['paper_free_units']:.3f} free / "
        f"{report['paper_committed_units']:.3f} reserved paper units",
    ]
    if report["wins"] + report["losses"]:
        lines.append(
            f"RECORD   {report['wins']}W / {report['losses']}L / {report['pushes']}P "
            f"({report['hit_rate']:.1%} hit) · ROI {report['roi']:+.1%}"
        )
    elif report["settled_accas"]:
        lines.append(
            f"RECORD   no win/loss decisions yet · {report['pushes']} pushes · "
            f"ROI {report['roi']:+.1%}"
        )
    else:
        lines.append("RECORD   no settled accas yet")
    lines.extend(
        [
            f"OPEN     {report['pending_accas']} held accas · "
            f"{report['unresolved_result_legs']} unresolved leg results",
            f"LEGS     Brier {report['selected_leg_brier_score']:.4f}"
            if report["selected_leg_brier_score"] is not None
            else "LEGS     Brier n/a",
            "",
            f"BET-DAYS, oldest first ({len(report['history'])})",
        ]
    )
    for day in report["history"]:
        returned = f"{day['returned_units']:.3f}" if day["returned_units"] is not None else "—"
        profit = f"{day['profit_units']:+.3f}" if day["profit_units"] is not None else "—"
        lines.append(
            f"  {day['date']}  {day['wins']}W/{day['losses']}L/{day['pushes']}P "
            f"staked {day['staked_units']:.3f} returned {returned} net {profit} "
            f"pending {day['pending_units']:.3f} bank* {day['bank_units']:.3f}"
        )
    lines.append("")
    lines.append("HELD (not yet settled; stake counted as committed)")
    held = False
    for day in report["history"]:
        for index, acca in enumerate(day["accas"], 1):
            if acca["grade_status"] != "pending":
                continue
            held = True
            lines.append(
                f"  {day['date']} ACCA #{index} @{acca['combined_decimal_odds']:.3f} "
                f"stake {acca['stake_units']:.3f} · receipt {day['decision_id']}"
            )
            for leg in acca["legs"]:
                lines.append(
                    f"      {leg['home']} vs {leg['away']} · {_selection_label(leg)} "
                    f"· {leg['grade_status'].upper()}"
                )
    if not held:
        lines.append("  None.")
    lines.extend(
        [
            "",
            "NOTES",
            "  bank* is reconstructed from latest known results by ticket date; pending is held at cost.",
            "  This is not a settlement-time equity curve or certified betting ROI.",
            "  A known losing leg settles its acca as a loss even if other legs remain unresolved.",
            "  Push legs are removed from effective price after every other leg resolves.",
            "  Empty scans never cancel tickets. Frozen receipts are never resized or rotated.",
            "  Paper only. No real-money stakes or automatic promotion.",
            "",
        ]
    )
    return "\n".join(lines)


def publish_operator_reports(
    store,
    league: League,
    market_report: Optional[Dict[str, Any]],
    localdata: Path,
    *,
    target_date: str,
    generated_utc: str,
    run_id: str,
    timezone_name: str = "Africa/Johannesburg",
) -> Dict[str, Any]:
    """Publish current picks and grade every retained pick/ticket in localdata."""
    localdata.mkdir(parents=True, exist_ok=True)
    games = {game.game_id: game for game in store.list_games(league=league)}
    current = [
        _pick_row(row, target_date)
        for row in deepcopy((market_report or {}).get("forecasts") or [])
    ]
    candidate_funnel = _ticket_candidate_funnel(market_report)
    dated_path = localdata / f"picks_{target_date}.json"
    prior = _read_json(dated_path, [])
    if not isinstance(prior, list):
        raise ValueError(f"dated pick ledger must be a list: {dated_path}")
    merged = {str(row["pick_id"]): row for row in prior}
    merged.update({row["pick_id"]: row for row in current})
    archived = sorted(
        merged.values(),
        key=lambda row: (row["scheduled_utc"], row["game_id"], row["pick_id"]),
    )
    for row in archived:
        _set_grade(row, games)
    # picks_today is the current scan only; the dated file retains earlier
    # frozen rows even when a later scan produces no fresh forecast for them.
    for row in current:
        if row["pick_id"] in merged:
            row.update(merged[row["pick_id"]])
    _write_json(localdata / "picks_today.json", current)
    _write_json(dated_path, archived)
    _refresh_dated_pick_reports(
        localdata,
        games=games,
        league=league.value,
        generated_utc=generated_utc,
        timezone_name=timezone_name,
    )
    audit = _write_pick_audit(
        localdata,
        target_date=target_date,
        generated_utc=generated_utc,
        timezone_name=timezone_name,
        games=games,
    )

    state = _load_state(localdata)
    acca_report = (market_report or {}).get("accas") or {}
    paper_results = acca_report.get("paper_results")
    if paper_results is None:  # legacy reports expose only the strict result surface
        paper_results = acca_report.get("results") or []
    ticket_accas = deepcopy(paper_results or [])
    receipt = _ticket_receipt(
        localdata,
        state=state,
        league=league,
        target_date=target_date,
        captured_utc=generated_utc,
        run_id=run_id,
        accas=ticket_accas,
    )
    state["scans"].append(
        {
            "run_id": run_id,
            "captured_at": generated_utc,
            "fresh_accas": len(ticket_accas),
            "receipt_present": receipt is not None,
            "candidate_funnel": deepcopy(candidate_funnel),
        }
    )
    state["scans"] = state["scans"][-100:]
    withdrawal = state["frozen_by_date"].get(target_date)
    if not (isinstance(withdrawal, dict) and withdrawal.get("withdrawn") is True):
        # Sidecar withdrawal also counts for display (survives state rollback)
        if _withdrawn_sidecar_path(localdata, target_date).exists():
            try:
                sc = json.loads(_withdrawn_sidecar_path(localdata, target_date).read_text())
                withdrawal = {
                    "withdrawn": True,
                    "decision_id": sc.get("decision_id"),
                    "sidecar": True,
                }
            except (OSError, json.JSONDecodeError):
                withdrawal = {"withdrawn": True, "sidecar": True}
        else:
            withdrawal = None
    withdrawn_decision_ids = {
        str(lock.get("decision_id"))
        for lock in state["frozen_by_date"].values()
        if isinstance(lock, dict) and lock.get("withdrawn") is True and lock.get("decision_id")
    } | _withdrawn_ids_from_sidecars(localdata)
    receipts = []
    for path in sorted(localdata.glob("auto_tickets_????-??-??_decision.json")):
        candidate = _read_json(path, {})
        if candidate.get("schema_version") != TICKET_SCHEMA:
            raise ValueError(f"invalid paper ticket receipt: {path}")
        if str(candidate.get("decision_id")) in withdrawn_decision_ids:
            continue
        receipts.append(candidate)
    performance = _performance(receipts, games)
    performance["generated_utc"] = generated_utc
    performance["league"] = league.value
    _write_json(localdata / "auto_tickets_performance.json", performance)
    _write_text(
        localdata / "auto_tickets_performance.txt",
        _render_performance(performance, timezone_name),
    )
    ticket_path = None
    if receipt is not None:
        graded_receipt = next(
            row
            for row in performance["history"]
            if (row["date"], row["decision_id"]) == (receipt["date"], receipt["decision_id"])
        )
        ticket_path = _write_text(
            localdata / f"auto_tickets_{target_date}.txt",
            _render_ticket_card(graded_receipt, performance, timezone_name),
        )
    else:
        ticket_path = _write_text(
            localdata / f"auto_tickets_{target_date}.txt",
            _render_empty_ticket_scan(
                league=league,
                target_date=target_date,
                generated_utc=generated_utc,
                performance=performance,
                timezone_name=timezone_name,
                candidate_funnel=candidate_funnel,
                withdrawal=withdrawal,
            ),
        )
    state["last_generated_utc"] = generated_utc
    state["paper_bank_units"] = performance["paper_bank_units"]
    _write_json(localdata / "auto_tickets_state.json", state)
    return {
        "paper_only": True,
        "picks_today": str(localdata / "picks_today.json"),
        "dated_picks": str(dated_path),
        "picks_text": str(localdata / f"picks_{target_date}.txt"),
        "rolling_audit": str(localdata / "picks_audit_rolling.json"),
        "dated_audit": str(localdata / f"picks_audit_{target_date}.md"),
        "ticket_receipt": (
            str(localdata / f"auto_tickets_{target_date}_decision.json")
            if receipt is not None
            else None
        ),
        "ticket_text": ticket_path,
        "ticket_state": str(localdata / "auto_tickets_state.json"),
        "performance": str(localdata / "auto_tickets_performance.json"),
        "performance_text": str(localdata / "auto_tickets_performance.txt"),
        "candidate_funnel": candidate_funnel,
        "counts": {
            "current_picks": len(current),
            "archived_picks": len(archived),
            "graded_picks": audit["settled_picks"],
            # Backward-compatible alias; this now includes same-day finals.
            "graded_prior_picks": audit["settled_picks"],
            "published_accas": len(receipt["accas"]) if receipt else 0,
            "settled_accas": performance["settled_accas"],
            "pending_accas": performance["pending_accas"],
        },
    }


__all__ = ["publish_operator_reports"]
