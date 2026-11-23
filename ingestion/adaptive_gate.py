"""Fail-open, lightweight scheduler gate for unattended league runs.

The full daily pipeline deliberately does expensive acquisition, history
backfill, market construction, and settlement work.  External scheduler ticks
are more frequent than all of that work is needed.  This module is the small,
read-only decision layer placed in front of the workflow matrix.

A league is allowed to skip only when all of the following are true:

* its compact Git-authoritative operator state and performance summary parse;
* a bounded, current-day schedule source was read successfully;
* no unsettled/awaiting operator position requires reconciliation; and
* there is no unfrozen slate at or inside the documented 60-minute ticket
  decision window.

Any ambiguity is deliberately a reason to run the full pipeline.  In
particular, EuroLeague currently exposes only the adapter's whole-season
schedule endpoint, not a verified date-bounded endpoint.  Fetching that on
every half-hour tick would only disguise an expensive acquisition as a gate,
so EuroLeague is conservatively fail-open until such a source is verified.

This module has no third-party dependencies.  Keeping the gate dependency-free
lets the Actions gate job make its decision before installing DuckDB or the
full project package.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time, timedelta, timezone
import json
import os
import re
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from ingestion.league_registry import SPECS
from ingestion.providers._cba_aes import Undecryptable, decrypt_body
from ingestion.providers.fiba_euroleague_w_contract import (
    COMPETITION_CODE,
    COMPETITION_ID,
    GDAP_BASE_URL,
    KEY_HEADER,
    KEY_PAGE_URL,
    PUBLIC_FALLBACK_SUBSCRIPTION_KEY,
    SCHEDULE_ENDPOINT,
    game_datetime_utc,
    has_complete_team_identity,
    has_unassigned_team_slot,
    page_subscription_key,
    schedule_status,
)

SCHEMA_VERSION = "magolide.adaptive-gate"
DECISION_LEAD_MINUTES = 120
# Freezes written under the previous one-hour lead remain readable evidence;
# see the matching legacy tuple in forecasting.operator.
LEGACY_DECISION_LEAD_MINUTES = (60,)
# Keep this universal policy aligned with forecasting.operator. Every dated
# ticket decision is at midday Johannesburg time, unless an earlier fixture
# needs the ordinary 60-minute pre-tip safety cutoff.
TICKET_PUBLICATION_TIMEZONE = "Africa/Johannesburg"
TICKET_PUBLICATION_LOCAL_TIME = time(hour=12)
HTTP_TIMEOUT_SECONDS = 15
MAX_RESPONSE_BYTES = 5 * 1024 * 1024
COVERAGE_GAP_MARKER_SCHEMA = "magolide.coverage_gap_marker"
HISTORICAL_CARD_RE = re.compile(r"^(picks|auto_tickets)_(\d{4}-\d{2}-\d{2})\.txt$")
# Keep this aligned with the recovery utility. It deliberately detects rendered
# grades only, never explanatory prose mentioning a pending observation.
PENDING_GRADE_RE = re.compile(r"(?:grade=|\[\s*)PENDING\b|⏳\s+PENDING\b")
# The temporary 2026-11-20/21 marker format is opaque to Git readers. It is
# not a ticket or forecast signal, but it must wake the serialized publisher so
# the bounded report-only recovery can replace it with retained readable text.
COMPACT_PICK_CARD_MARKER = b"Ma Golide compact pick lifecycle marker\n"
# Public browser-bundle fallback for the documented CBA data response format;
# it is not a repository credential. An operator may override it for a rotated
# public format through CBA_DECRYPT_KEY without exposing the value in reports.
PUBLIC_CBA_AES_KEY = b"uVayqL4ONKjFbVzQ"

# The FIBA wire contract (including public-key fallback and status mapping) is
# shared with the full provider above.  This gate imports only its stdlib-only
# contract module, preserving its dependency-free early-workflow role.

# This is kept in the same order as the workflow matrix so its JSON ``include``
# rows can be consumed directly by GitHub Actions.
@dataclass(frozen=True)
class LeagueGateConfig:
    label: str
    # CLI/workflow spelling retained at the outer boundary. ``league_key`` is
    # the canonical enum value used for every persisted-state comparison.
    league: str
    league_key: str
    state_key: str
    acquisition_timezone: str
    schedule_source: str


LEAGUES: tuple[LeagueGateConfig, ...] = tuple(
    LeagueGateConfig(
        spec.label,
        # Keep the established workflow spelling as a boundary alias; the
        # daily CLI immediately normalises it to League.EUROLEAGUE.
        "EuroLeague" if spec.league.value == "EUROLEAGUE" else spec.league.value,
        spec.league.value,
        spec.state_key,
        spec.acquisition_timezone,
        spec.gate_schedule_source,
    )
    for spec in SPECS
)


class GateUncertain(RuntimeError):
    """A condition which must make the gate run the full pipeline."""


@dataclass(frozen=True)
class ScheduleGame:
    """The minimal schedule fact the gate needs; never canonical evidence."""

    scheduled_utc: datetime
    status: str
    # Source event identity is used only to wake a per-fixture draft/freeze
    # lifecycle. Empty remains supported for hermetic legacy schedule tests.
    game_id: str = ""


@dataclass(frozen=True)
class GateDecision:
    """One league's read-only gate decision."""

    config: LeagueGateConfig
    target_date: str
    should_run: bool
    reasons: tuple[str, ...]
    scheduled_game_count: int | None

    def as_dict(self) -> dict[str, Any]:
        return {
            "label": self.config.label,
            "league": self.config.league,
            "league_key": self.config.league_key,
            "state_key": self.config.state_key,
            "target_date": self.target_date,
            "should_run": self.should_run,
            "reasons": list(self.reasons),
            "scheduled_game_count": self.scheduled_game_count,
        }

    def matrix_row(self) -> dict[str, str]:
        return {
            "label": self.config.label,
            "league": self.config.league,
            "league_key": self.config.league_key,
            "state_key": self.config.state_key,
            "target_date": self.target_date,
            "gate_reason": "; ".join(self.reasons),
        }


ScheduleFetcher = Callable[[LeagueGateConfig, str], Iterable[ScheduleGame]]


def utc_now() -> datetime:
    """Return a second-precision aware UTC instant for stable reports."""

    return datetime.now(timezone.utc).replace(microsecond=0)


def _parse_utc(value: Any, *, context: str) -> datetime:
    if not isinstance(value, str) or not value.strip():
        raise GateUncertain(f"{context} has no timestamp")
    try:
        parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError as exc:
        raise GateUncertain(f"{context} timestamp is invalid") from exc
    if parsed.tzinfo is None:
        raise GateUncertain(f"{context} timestamp has no timezone")
    return parsed.astimezone(timezone.utc)


def target_date_for(config: LeagueGateConfig, now: datetime) -> str:
    """Return today's date in the same acquisition timezone as ``daily.py``."""

    if now.tzinfo is None:
        raise ValueError("gate clock must be timezone-aware")
    return now.astimezone(ZoneInfo(config.acquisition_timezone)).strftime("%Y-%m-%d")


def _ticket_decision_cutoff(
    config: LeagueGateConfig, target_date: str, games: Iterable[ScheduleGame]
) -> datetime | None:
    """Return the league card cutoff, bounded by the first game's safety window."""

    try:
        target_day = datetime.strptime(target_date, "%Y-%m-%d").date()
    except ValueError as exc:  # pragma: no cover - target dates are internal ISO facts
        raise GateUncertain("target date is invalid") from exc
    tips = [
        game.scheduled_utc.astimezone(timezone.utc)
        for game in games
        if game.status in {"scheduled", "live"}
    ]
    if not tips:
        return None
    publication = datetime.combine(
        target_day,
        TICKET_PUBLICATION_LOCAL_TIME,
        tzinfo=ZoneInfo(TICKET_PUBLICATION_TIMEZONE),
    ).astimezone(timezone.utc)
    return min(publication, min(tips) - timedelta(minutes=DECISION_LEAD_MINUTES))


def _coverage_gap_marker_path(state_dir: Path, target_date: str) -> Path:
    return state_dir / f"coverage_gap_{target_date}.json"


def _source_unavailable_reasons(
    state_dir: Path, *, config: LeagueGateConfig, target_date: str
) -> list[str]:
    """Throttle known-unavailable lanes without inventing an empty schedule.

    A Git-durable marker records that today's explicit CoverageGap was already
    published. Existing ticket or pick-lifecycle work still wins over the
    marker; unreadable/missing evidence remains a run decision.
    """

    marker_path = _coverage_gap_marker_path(state_dir, target_date)
    if not marker_path.exists():
        return ["coverage_gap_publication_due"]
    marker = _read_json(marker_path, description="coverage-gap marker")
    if (
        marker.get("schema_version") != COVERAGE_GAP_MARKER_SCHEMA
        or marker.get("league") != config.league_key
        or marker.get("target_date") != target_date
        or marker.get("schedule_source") != "source_unavailable"
    ):
        raise GateUncertain("coverage-gap marker has an unknown identity or schema")
    # The gate cannot discover a source-unavailable fixture itself. A dated
    # durable text card is therefore proof of a manually supplied lifecycle
    # that needs the full pipeline rather than permission to skip it as an
    # empty slate. The corresponding structured JSON stays cache/artifact-only.
    if (state_dir / f"picks_{target_date}.txt").is_file():
        return ["source_unavailable_pick_lifecycle_due"]
    return ()


def _inline_frozen_game_ids(
    lifecycle: dict[str, Any], *, config: LeagueGateConfig, target_date: str
) -> set[str]:
    """Validate and return one date's immutable inline fixture snapshots."""

    frozen_by_date = lifecycle.get("frozen_by_date", {})
    if not isinstance(frozen_by_date, dict):
        raise GateUncertain("pick lifecycle state has invalid frozen fixture locks")
    frozen = frozen_by_date.get(target_date, {})
    if not isinstance(frozen, dict):
        raise GateUncertain("pick lifecycle state has invalid frozen fixture date locks")
    frozen_ids: set[str] = set()
    for game_id, payload in frozen.items():
        if not isinstance(payload, dict) or str(game_id) != str(payload.get("game_id") or ""):
            raise GateUncertain("pick lifecycle state has an invalid frozen fixture lock")
        if (
            payload.get("schema_version") != "magolide.pick_freeze"
            or payload.get("target_date") != target_date
            or payload.get("league") != config.league_key
            or not isinstance(payload.get("rows"), list)
        ):
            raise GateUncertain("pick lifecycle state has an invalid frozen fixture payload")
        tip = _parse_utc(payload.get("scheduled_utc"), context="frozen pick fixture")
        cutoff = _parse_utc(payload.get("decision_cutoff_utc"), context="frozen pick cutoff")
        frozen_at = _parse_utc(payload.get("frozen_at"), context="frozen pick timestamp")
        accepted_cutoffs = {
            tip - timedelta(minutes=lead)
            for lead in (DECISION_LEAD_MINUTES, *LEGACY_DECISION_LEAD_MINUTES)
        }
        if cutoff not in accepted_cutoffs or not cutoff <= frozen_at < tip:
            raise GateUncertain("pick lifecycle state has an invalid frozen fixture cutoff")
        frozen_ids.add(str(game_id))
    return frozen_ids


def _pick_lifecycle_reasons(
    state_dir: Path,
    *,
    config: LeagueGateConfig,
    target_date: str,
    games: Iterable[ScheduleGame],
    now: datetime,
) -> list[str]:
    """Decide whether draft, fixture-freeze, or frozen-result work is due."""

    reasons: list[str] = []
    games = list(games)
    lifecycle_path = state_dir / "pick_lifecycle_state.json"
    lifecycle = None
    frozen_game_ids: set[str] = set()
    if lifecycle_path.exists():
        lifecycle = _read_json(lifecycle_path, description="pick lifecycle state")
        if lifecycle.get("schema_version") != "magolide.pick_lifecycle_state":
            raise GateUncertain("pick lifecycle state has an unknown schema")
        if lifecycle.get("league") != config.league_key:
            raise GateUncertain("pick lifecycle state belongs to a different or unknown league")
        frozen_game_ids = _inline_frozen_game_ids(
            lifecycle, config=config, target_date=target_date
        )
    # One early run publishes the current-day draft; after that the gate wakes
    # only at each fixture's own pre-tip window or for result reconciliation.
    if (
        any(game.game_id for game in games)
        and not (state_dir / f"picks_{target_date}.txt").is_file()
    ):
        reasons.append("pick_draft_publication_due")
    for game in games:
        if game.status not in {"scheduled", "live"}:
            continue
        tip = game.scheduled_utc.astimezone(timezone.utc)
        if (
            game.game_id
            and tip - timedelta(minutes=DECISION_LEAD_MINUTES) <= now < tip
            and game.game_id not in frozen_game_ids
        ):
            reasons.append("fixture_pick_freeze_due")
            break
    if lifecycle is None:
        return reasons
    pending = lifecycle.get("pending_frozen_games")
    if not isinstance(pending, list):
        raise GateUncertain("pick lifecycle state has invalid pending_frozen_games")
    for item in pending:
        if not isinstance(item, (list, tuple)) or len(item) != 2:
            raise GateUncertain("pick lifecycle state has invalid pending fixture")
        tip = _parse_utc(item[1], context="frozen pick fixture")
        if tip <= now:
            reasons.append("frozen_pick_settlement_due")
            break
    return reasons


def _read_json(path: Path, *, description: str) -> dict[str, Any]:
    try:
        raw = json.loads(path.read_text())
    except FileNotFoundError as exc:
        raise GateUncertain(f"missing {description}") from exc
    except (OSError, json.JSONDecodeError) as exc:
        raise GateUncertain(f"unreadable {description}") from exc
    if not isinstance(raw, dict):
        raise GateUncertain(f"invalid {description}")
    return raw


def _nonnegative_int(value: Any, *, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise GateUncertain(f"invalid performance field {field}")
    return value


def _load_operator_summary(
    state_dir: Path, *, league: str
) -> tuple[dict[str, Any], dict[str, Any]]:
    state = _read_json(state_dir / "auto_tickets_state.json", description="operator state")
    performance = _read_json(
        state_dir / "auto_tickets_performance.json", description="operator performance"
    )
    if not isinstance(state.get("frozen_by_date"), dict):
        raise GateUncertain("operator state has invalid frozen_by_date")
    if performance.get("league") != league:
        raise GateUncertain("operator performance belongs to a different or unknown league")
    # The current compact schema always writes all of these.  Treat an older
    # or partial summary as unknown rather than accidentally assuming no work.
    for field in (
        "pending_positions",
        "pending_accas",
        "pending_singles",
        "awaiting_game_legs",
        "unresolved_result_legs",
        "result_unresolvable_legs",
    ):
        _nonnegative_int(performance.get(field), field=field)
    return state, performance


def _validate_sidecar(
    path: Path,
    *,
    target_date: str,
    league: str,
    no_bet: bool,
) -> None:
    decision = _read_json(path, description=f"frozen decision {path.name}")
    date = decision.get("target_date") if no_bet else decision.get("date")
    if date != target_date or decision.get("league") != league:
        raise GateUncertain(f"frozen decision identity mismatch for {target_date}")
    if no_bet:
        if decision.get("status") != "frozen_no_bet":
            raise GateUncertain(f"invalid no-bet decision for {target_date}")
    elif decision.get("status") != "frozen":
        raise GateUncertain(f"invalid ticket decision for {target_date}")


def date_is_frozen(
    state_dir: Path,
    state: dict[str, Any],
    *,
    target_date: str,
    league: str,
) -> bool:
    """Return whether today's decision is durably frozen and internally sound.

    A ticket lock without its matching immutable receipt is not enough to skip
    a run. No-bet evidence is instead embedded in its durable state lock, so a
    missing per-date sidecar cannot create artifact bloat or a false no-work
    decision.
    """

    locks = state["frozen_by_date"]
    lock = locks.get(target_date)
    decision_path = state_dir / f"auto_tickets_{target_date}_decision.json"
    withdrawn_path = state_dir / f"auto_tickets_{target_date}_withdrawn.json"

    if lock is None:
        if decision_path.exists() or withdrawn_path.exists():
            raise GateUncertain(f"decision sidecar lacks a state lock for {target_date}")
        return False
    if not isinstance(lock, dict) or lock.get("final") is not True:
        raise GateUncertain(f"invalid state lock for {target_date}")

    if lock.get("withdrawn") is True:
        withdrawal = _read_json(withdrawn_path, description=f"withdrawn decision {target_date}")
        if (
            withdrawal.get("withdrawn") is not True
            or withdrawal.get("target_date") != target_date
            or withdrawal.get("decision_id") != lock.get("decision_id")
        ):
            raise GateUncertain(f"invalid withdrawn sidecar for {target_date}")
        return True
    if lock.get("no_bet") is True:
        if decision_path.exists():
            raise GateUncertain(f"no-bet lock conflicts with ticket receipt for {target_date}")
        decision = lock.get("no_bet_decision")
        if (
            not isinstance(decision, dict)
            or decision.get("target_date") != target_date
            or decision.get("league") != league
            or decision.get("status") != "frozen_no_bet"
            or decision.get("decision_id") != lock.get("decision_id")
        ):
            raise GateUncertain(f"invalid inline no-bet decision for {target_date}")
        return True
    _validate_sidecar(decision_path, target_date=target_date, league=league, no_bet=False)
    return True


def _pending_reasons(performance: dict[str, Any]) -> list[str]:
    pending_positions = sum(
        _nonnegative_int(performance[field], field=field)
        for field in ("pending_positions", "pending_accas", "pending_singles")
    )
    waiting_for_results = sum(
        _nonnegative_int(performance[field], field=field)
        for field in (
            "awaiting_game_legs",
            "unresolved_result_legs",
            "result_unresolvable_legs",
        )
    )
    reasons = []
    if pending_positions:
        reasons.append("pending_settlement")
    if waiting_for_results:
        reasons.append("result_reconciliation")
    return reasons


def _historical_pending_card_reasons(state_dir: Path, *, target_date: str) -> list[str]:
    """Wake recovery when an earlier public card still displays a real grade.

    Structured dated pick ledgers are intentionally cache/artifact-only, so a
    cold runner cannot inspect their rows before the full pipeline starts. The
    Git-durable cards are sufficient to request a *bounded recovery attempt*.
    Only a recovered original ledger permits normal lifecycle rendering to
    settle a row; if retained evidence is absent, the operator's missing-ledger
    guards leave the historical card untouched. This catches unfrozen pick
    cards and audit-only auto-ticket shadow cards that have no receipt/no-bet
    sidecar, categories not represented in performance.
    """

    reasons: list[str] = []
    try:
        cards = sorted(path for path in state_dir.iterdir() if path.is_file())
    except FileNotFoundError:
        return reasons
    except OSError as exc:
        raise GateUncertain("cannot inspect historical operator cards") from exc
    for path in cards:
        match = HISTORICAL_CARD_RE.fullmatch(path.name)
        if match is None:
            continue
        kind, card_date = match.groups()
        # ISO dates have chronological lexical order. Future/current cards
        # must remain draft-capable and are intentionally outside recovery.
        if card_date >= target_date:
            continue
        try:
            has_pending_grade = PENDING_GRADE_RE.search(path.read_text(encoding="utf-8")) is not None
        except OSError as exc:
            raise GateUncertain(f"cannot read historical {kind} card {card_date}") from exc
        if has_pending_grade:
            reasons.append(f"historical_{kind}_reconciliation_due:{card_date}")
    return reasons


def _legacy_compact_pick_card_reasons(state_dir: Path) -> list[str]:
    """Wake the publisher to replace an opaque legacy dated-card marker.

    A compact marker is presentation debt rather than a forecast/ticket input.
    Its card content is recovered later by the publisher from a selected Git
    snapshot or a verified release asset; the gate only ensures an otherwise
    idle external tick reaches that serialized migration.
    """

    try:
        cards = sorted(state_dir.glob("picks_*.txt")) if state_dir.is_dir() else []
        for path in cards:
            if path.is_symlink() or not path.is_file() or not HISTORICAL_CARD_RE.fullmatch(path.name):
                continue
            if path.read_bytes().startswith(COMPACT_PICK_CARD_MARKER):
                return ["legacy_compact_pick_card_restore_due"]
    except OSError as exc:
        raise GateUncertain("cannot read legacy compact pick card") from exc
    return []


def _settlement_view_reasons(state_dir: Path, performance: dict[str, Any]) -> list[str]:
    """Wake a lane when its public receipt card lags settled compact state.

    Decision JSON stays immutable. The dated ``.txt`` card is intentionally a
    derived settlement view, so a completed receipt must be re-rendered if it
    is absent or still advertises a pending position after all grades settle.
    Malformed history is state uncertainty rather than a false no-work result.
    """

    history = performance.get("history") or []
    if not isinstance(history, list):
        raise GateUncertain("operator performance has invalid history")
    state = _read_json(state_dir / "auto_tickets_state.json", description="operator state")
    locks = state.get("frozen_by_date")
    if not isinstance(locks, dict):
        raise GateUncertain("operator state has invalid frozen date locks")
    reasons: list[str] = []
    settled_fields = (
        "pending_positions",
        "pending_accas",
        "pending_singles",
        "awaiting_game_legs",
        "unresolved_result_legs",
        "result_unresolvable_legs",
    )
    for row in history:
        if not isinstance(row, dict):
            raise GateUncertain("operator performance history has an invalid row")
        date = row.get("date")
        if not isinstance(date, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", date):
            raise GateUncertain("operator performance history has an invalid date")
        if any(_nonnegative_int(row.get(field), field=f"history.{field}") for field in settled_fields):
            continue
        receipt = state_dir / f"auto_tickets_{date}_decision.json"
        lock = locks.get(date)
        if (isinstance(lock, dict) and lock.get("no_bet") is True) or not receipt.is_file():
            continue
        card = state_dir / f"auto_tickets_{date}.txt"
        try:
            text = card.read_text(encoding="utf-8")
        except FileNotFoundError:
            reasons.append(f"settlement_view_refresh_due:{date}")
            continue
        except OSError as exc:
            raise GateUncertain(f"cannot read settlement view {date}") from exc
        # The renderer uses this stable grade label in both ticket headers and
        # legs. A zero-pending compact receipt with the label still present is
        # provably a stale human view, not a new ticket decision.
        if "PENDING" in text:
            reasons.append(f"settlement_view_refresh_due:{date}")
    return reasons


def evaluate_league(
    config: LeagueGateConfig,
    *,
    state_root: Path,
    now: datetime,
    schedule_fetcher: ScheduleFetcher,
) -> GateDecision:
    """Evaluate one league, never returning a skip after uncertainty.

    ``schedule_fetcher`` is injectable so ordinary tests are entirely
    hermetic.  Exceptions from it are intentionally folded into a run decision
    rather than propagated to the workflow as an unsafe skipped matrix.
    """

    target_date = target_date_for(config, now)
    state_dir = state_root / config.state_key
    reasons: list[str] = []
    schedule_count: int | None = None
    try:
        state, performance = _load_operator_summary(state_dir, league=config.league_key)
        frozen = date_is_frozen(
            state_dir, state, target_date=target_date, league=config.league_key
        )
        reasons.extend(_pending_reasons(performance))
        reasons.extend(_settlement_view_reasons(state_dir, performance))
        reasons.extend(_historical_pending_card_reasons(state_dir, target_date=target_date))
        reasons.extend(_legacy_compact_pick_card_reasons(state_dir))
    except GateUncertain as exc:
        return GateDecision(
            config, target_date, True, (f"state_uncertain: {exc}",), schedule_count
        )

    if config.schedule_source == "source_unavailable":
        try:
            # There is no compliant schedule probe to call. Do not represent
            # that as an authoritative empty slate; only skip once today's
            # explicit CoverageGap is Git durable and no lifecycle work exists.
            reasons.extend(
                _pick_lifecycle_reasons(
                    state_dir,
                    config=config,
                    target_date=target_date,
                    games=(),
                    now=now,
                )
            )
            reasons.extend(
                _source_unavailable_reasons(
                    state_dir, config=config, target_date=target_date
                )
            )
        except GateUncertain as exc:
            return GateDecision(
                config,
                target_date,
                True,
                tuple(dict.fromkeys([*reasons, f"state_uncertain: {exc}"])),
                schedule_count,
            )
        unique_reasons = tuple(dict.fromkeys(reasons))
        return GateDecision(
            config,
            target_date,
            bool(unique_reasons),
            unique_reasons or ("no_work",),
            schedule_count,
        )

    try:
        games = list(schedule_fetcher(config, target_date))
        schedule_count = len(games)
        for game in games:
            if not isinstance(game, ScheduleGame):
                raise GateUncertain("schedule source returned an invalid game")
            if game.scheduled_utc.tzinfo is None:
                raise GateUncertain("schedule source returned a timezone-naive tip")
            if game.status not in {"scheduled", "live", "final"}:
                raise GateUncertain("schedule source returned an unknown game status")
    except Exception as exc:  # fail open for provider/network/parser uncertainty
        detail = str(exc).strip() or type(exc).__name__
        return GateDecision(
            config,
            target_date,
            True,
            tuple(dict.fromkeys([*reasons, f"schedule_uncertain: {detail}"])),
            schedule_count,
        )

    try:
        reasons.extend(
            _pick_lifecycle_reasons(
                state_dir,
                config=config,
                target_date=target_date,
                games=games,
                now=now,
            )
        )
    except GateUncertain as exc:
        return GateDecision(
            config,
            target_date,
            True,
            tuple(dict.fromkeys([*reasons, f"state_uncertain: {exc}"])),
            schedule_count,
        )
    if not frozen:
        decision_cutoff = _ticket_decision_cutoff(config, target_date, games)
        # Once the publication cutoff arrives, wake the lane until it records an
        # immutable ticket or no-bet.  If the run arrives late it will record
        # the explicit missed-window no-bet rather than creating a post-tip
        # position.  A no-game slate stays idle after its normal draft work.
        if decision_cutoff is not None and now >= decision_cutoff:
            reasons.append("ticket_creation_or_acquisition_due")

    unique_reasons = tuple(dict.fromkeys(reasons))
    return GateDecision(config, target_date, bool(unique_reasons), unique_reasons or ("no_work",), schedule_count)


def evaluate_gate(
    *,
    state_root: Path,
    now: datetime | None = None,
    schedule_fetcher: ScheduleFetcher | None = None,
    force_all_reason: str | None = None,
) -> list[GateDecision]:
    """Evaluate the registry-driven league matrix, with explicit force-open mode."""

    current = now or utc_now()
    if current.tzinfo is None:
        raise ValueError("gate clock must be timezone-aware")
    if force_all_reason:
        return [
            GateDecision(
                config,
                target_date_for(config, current),
                True,
                (force_all_reason,),
                None,
            )
            for config in LEAGUES
        ]
    fetch = schedule_fetcher or fetch_schedule
    return [
        evaluate_league(config, state_root=state_root, now=current, schedule_fetcher=fetch)
        for config in LEAGUES
    ]


def report_for(decisions: Iterable[GateDecision], *, generated_utc: datetime) -> dict[str, Any]:
    """Serialise decisions for Actions outputs and the human gate summary."""

    rows = list(decisions)
    include = [row.matrix_row() for row in rows if row.should_run]
    return {
        "schema_version": SCHEMA_VERSION,
        "generated_utc": generated_utc.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "decision_lead_minutes": DECISION_LEAD_MINUTES,
        "run_any": bool(include),
        "matrix": {"include": include},
        "decisions": [row.as_dict() for row in rows],
    }


def _read_limited_response(url: str, *, headers: dict[str, str] | None = None) -> str:
    request = Request(
        url,
        headers={"User-Agent": "MaGolideAdaptiveGate/1.0", **(headers or {})},
    )
    try:
        with urlopen(request, timeout=HTTP_TIMEOUT_SECONDS) as response:  # noqa: S310 public fixed hosts
            body = response.read(MAX_RESPONSE_BYTES + 1)
    except Exception as exc:
        raise GateUncertain(f"schedule request failed ({type(exc).__name__})") from exc
    if len(body) > MAX_RESPONSE_BYTES:
        raise GateUncertain("schedule response exceeded gate size limit")
    try:
        return body.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise GateUncertain("schedule response was not UTF-8") from exc


def _read_json_response(url: str, *, headers: dict[str, str] | None = None) -> Any:
    try:
        return json.loads(_read_limited_response(url, headers=headers))
    except json.JSONDecodeError as exc:
        raise GateUncertain("schedule response was not JSON") from exc


def _bal_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """One BAL day from the season blob embedded in bal.nba.com/schedule.

    The blob holds the whole season, so the gate reads one page and keeps the
    rows whose UTC tip falls on the target date. State is derived from scores
    (a settled basketball game cannot be 0-0) rather than from the unverified
    ``gameStatus`` code.
    """

    from ingestion.providers import bal

    html = _read_limited_response(bal.SCHEDULE_URL)
    payload = bal.extract_schedule(html)
    games: list[ScheduleGame] = []
    for group in bal._game_dates(payload):
        for row in group.get("games") or []:
            if not isinstance(row, dict):
                continue
            stamp = row.get("gameDateTimeUTC")
            if not isinstance(stamp, str) or not stamp.startswith(target_date):
                continue
            game_id = str(row.get("gameId") or "").strip()
            if not game_id:
                continue
            home = (row.get("homeTeam") or {}).get("score")
            away = (row.get("awayTeam") or {}).get("score")
            postponed = str(row.get("postponedStatus") or "N").strip().upper() == "P"
            if postponed:
                continue
            played = (
                isinstance(home, (int, float))
                and isinstance(away, (int, float))
                and (home > 0 or away > 0)
            )
            status = "final" if played else "scheduled"
            tip = datetime.fromisoformat(stamp.replace("Z", "+00:00"))
            games.append(
                ScheduleGame(
                    scheduled_utc=tip.astimezone(timezone.utc),
                    status=status,
                    game_id=f"BAL-{game_id}",
                )
            )
    return games


def _bcl_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """One BCL day from the FIBA GDAP season index (one call, filtered here).

    The index is not date-parameterised, so the gate reads it once and keeps the
    rows whose UTC tip falls on the target date. An unclassifiable status skips
    the row rather than asserting a state the gate cannot act on.
    """

    from ingestion.providers import bcl
    from ingestion.providers.fiba_euroleague_w_contract import schedule_status

    url = bcl.schedule_url()
    body = _read_json_response(url, headers=bcl.request_headers())
    rows = body if isinstance(body, list) else None
    if rows is None and isinstance(body, dict):
        rows = body.get("value") if isinstance(body.get("value"), list) else None
    if rows is None:
        raise GateUncertain("BCL GDAP index has no row list")
    games: list[ScheduleGame] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        stamp = row.get("gameDateTimeUTC")
        if not isinstance(stamp, str) or not stamp.startswith(target_date):
            continue
        if row.get("teamA") is None or row.get("teamB") is None:
            continue
        status = schedule_status(row)
        if status not in ("scheduled", "live", "final"):
            continue
        game_id = row.get("gameId")
        if game_id is None:
            continue
        tip = datetime.fromisoformat(
            stamp[:-1] + "+00:00" if stamp.endswith(("Z", "z")) else stamp
        )
        if tip.tzinfo is None:
            tip = tip.replace(tzinfo=timezone.utc)
        games.append(
            ScheduleGame(
                scheduled_utc=tip.astimezone(timezone.utc),
                status=status,
                game_id=f"BCL-{int(game_id)}",
            )
        )
    return games


def _bleague_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """One B.League B1 day from the site's server-rendered slate.

    The slate is asked for in JST terms and already carries the tip time and the
    status text, so the gate needs two page reads and no per-game snapshot. A
    page with no recognisable game block fails open rather than asserting an
    empty slate.
    """

    from ingestion.providers import bleague

    day = datetime.strptime(target_date.replace("-", ""), "%Y%m%d")
    local_day = (day.replace(tzinfo=timezone.utc) + timedelta(hours=9)).date()
    games: list[ScheduleGame] = []
    seen: set[str] = set()
    for offset in (0, 1):
        url = bleague.schedule_url(local_day + timedelta(days=offset), bleague.TAB_B1)
        blocks = bleague.parse_slate(_read_limited_response(url))
        if not blocks:
            raise GateUncertain("B.LEAGUE schedule page has no recognisable game block")
        for block in blocks:
            key = block["schedule_key"]
            if key in seen or not block["tip_local"]:
                continue
            try:
                hour, minute = (int(part) for part in block["tip_local"].split(":"))
                tip = datetime(
                    local_day.year, local_day.month, local_day.day, hour, minute
                ) + timedelta(days=offset)
            except (TypeError, ValueError):
                continue
            tip = tip.replace(tzinfo=timezone(timedelta(hours=9)))
            if tip.astimezone(timezone.utc).strftime("%Y-%m-%d") != target_date:
                continue
            seen.add(key)
            games.append(
                ScheduleGame(
                    scheduled_utc=tip.astimezone(timezone.utc),
                    status=bleague.status_from_text(block["status_text"]),
                    game_id=f"BLEAGUE-{key}",
                )
            )
    return games


def _tbf_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """One TBF day from the federation's daily-matches route (no key).

    The route is keyed by Turkey local date and answers every competition the
    federation runs in one body, so the gate reads the two local dates a UTC
    target day can span and keeps only BSL rows whose UTC tip lands inside it.
    An unreadable body fails open instead of asserting an empty slate.
    """

    day = datetime.strptime(target_date.replace("-", ""), "%Y%m%d")
    local_day = (day.replace(tzinfo=timezone.utc) + timedelta(hours=3)).date()
    games: list[ScheduleGame] = []
    seen: set[str] = set()
    for offset in (0, 1):
        url = (
            "https://miniappapi.tbf.org.tr/webapi-service/api/Match/get-daily-matches"
            f"?match_date={(local_day + timedelta(days=offset)).isoformat()}"
        )
        body = _read_json_response(url, headers={"Accept": "application/json"})
        if not isinstance(body, dict) or not isinstance(body.get("data"), list):
            raise GateUncertain("TBF get-daily-matches has no data list")
        for group in body["data"]:
            matches = group.get("matches") if isinstance(group, dict) else None
            for row in matches or []:
                if not isinstance(row, dict):
                    continue
                if str(row.get("activityCode") or "").strip().lower() != "bsl":
                    continue
                match_id = row.get("matchId")
                if match_id is None:
                    continue
                key = str(int(float(match_id)))
                if key in seen:
                    continue
                try:
                    tip = datetime.fromisoformat(str(row.get("matchDate") or "").strip())
                except ValueError:
                    continue
                if tip.tzinfo is None:
                    tip = tip.replace(tzinfo=timezone(timedelta(hours=3)))
                if tip.astimezone(timezone.utc).strftime("%Y-%m-%d") != target_date:
                    continue
                seen.add(key)
                played = str(row.get("matchStatusId") or "").strip()
                games.append(
                    ScheduleGame(
                        scheduled_utc=tip.astimezone(timezone.utc),
                        status="final" if played in {"1", "1.0"} else "scheduled",
                        game_id=f"TBF-{key}",
                    )
                )
    return games


def _kbl_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """One KBL day from the documented date-window route (no key, two headers).

    The host answers a window inside one calendar year in a single request, so
    the half-hour gate reads exactly the KST target day plus the UTC day either
    side of it - never a season. Rows whose tip falls outside the target UTC
    date are dropped, and an unreadable body fails open instead of asserting an
    empty slate.
    """

    day = datetime.strptime(target_date.replace("-", ""), "%Y%m%d")
    local = day.replace(tzinfo=timezone.utc) + timedelta(hours=9)
    start = (local - timedelta(days=1)).strftime("%Y%m%d")
    end = (local + timedelta(days=1)).strftime("%Y%m%d")
    season = local.year + 1 if local.month >= 8 else local.year
    url = (
        f"https://api.kbl.or.kr/match/list?season={season}"
        f"&fromDate={start}&toDate={end}"
    )
    body = _read_json_response(url, headers={"Channel": "WEB", "TeamCode": "XX"})
    rows = body if isinstance(body, list) else None
    if rows is None and isinstance(body, dict):
        for key in ("data", "list", "games", "items"):
            if isinstance(body.get(key), list):
                rows = body[key]
                break
    if rows is None:
        raise GateUncertain("KBL match/list has no recognisable row list")
    games: list[ScheduleGame] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        raw_date = str(row.get("game_date") or "").replace("-", "")
        raw_start = str(row.get("game_start") or "").strip().zfill(4)
        if len(raw_date) != 8 or not raw_date.isdigit() or not raw_start.isdigit():
            continue
        tip = datetime.strptime(raw_date + raw_start, "%Y%m%d%H%M").replace(
            tzinfo=timezone(timedelta(hours=9))
        )
        status = (
            "final"
            if str(row.get("is_ended") or "0").strip() == "1"
            else "live"
            if str(row.get("is_started") or "0").strip() == "1"
            else "scheduled"
        )
        games.append(
            ScheduleGame(
                scheduled_utc=tip,
                status=status,
                game_id=f"KBL-{row.get('gmkey')}",
            )
        )
    return [
        game
        for game in games
        if game.scheduled_utc.astimezone(timezone.utc).strftime("%Y-%m-%d")
        == target_date
    ]


def _espn_scoreboard_url(config: LeagueGateConfig, target_date: str | None) -> str:
    suffix: str
    if config.league == "NBA":
        slug, suffix = "nba", "limit=100"
    elif config.league == "WNBA":
        slug, suffix = "wnba", "limit=100"
    elif config.league == "NCAA_M":
        slug, suffix = "mens-college-basketball", "groups=50&limit=500"
    elif config.league == "NCAA_W":
        slug, suffix = "womens-college-basketball", "groups=50&limit=500"
    elif config.league == "AUSTRALIA_NBL":
        # ESPN's own NBL competition route (league UID 55); date-bounded, so
        # the half-hour gate may read it directly.
        slug, suffix = "nbl", "limit=100"
    elif config.league == "NBA_G_LEAGUE":
        # Validated separately from NBA: this is ESPN's NBA G League route,
        # not the NBA scoreboard with a group filter.
        slug, suffix = "nba-development", "limit=100"
    else:  # pragma: no cover - guarded by LEAGUES
        raise GateUncertain("unsupported ESPN league")
    date_query = f"dates={target_date.replace('-', '')}&" if target_date else ""
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
        f"scoreboard?{date_query}{suffix}"
    )


def _espn_games(payload: Any) -> list[ScheduleGame]:
    if not isinstance(payload, dict) or not isinstance(payload.get("events"), list):
        raise GateUncertain("ESPN scoreboard lacks an events list")
    games: list[ScheduleGame] = []
    for event in payload["events"]:
        if not isinstance(event, dict):
            raise GateUncertain("ESPN scoreboard has an invalid event")
        competitions = event.get("competitions")
        if not isinstance(competitions, list) or len(competitions) != 1:
            raise GateUncertain("ESPN event lacks exactly one competition")
        competition = competitions[0]
        if not isinstance(competition, dict):
            raise GateUncertain("ESPN competition is invalid")
        status = ((competition.get("status") or {}).get("type") or {}).get("state")
        mapped_status = {"pre": "scheduled", "in": "live", "post": "final"}.get(status)
        if mapped_status is None:
            raise GateUncertain("ESPN event has an unknown state")
        event_id = str(event.get("id") or "").strip()
        if not event_id:
            raise GateUncertain("ESPN event lacks an identity")
        games.append(
            ScheduleGame(
                _parse_utc(competition.get("date") or event.get("date"), context="ESPN event"),
                mapped_status,
                event_id,
            )
        )
    return games


def _espn_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    games = _espn_games(_read_json_response(_espn_scoreboard_url(config, target_date)))
    if games:
        return games
    # A valid empty date-filtered payload is not sufficient proof that no work
    # exists. Corroborate it once with ESPN's unfiltered current scoreboard;
    # any same-local-date event means the date endpoint is uncertain and the
    # gate must run rather than skip a fixture's freeze window.
    unfiltered = _espn_games(_read_json_response(_espn_scoreboard_url(config, None)))
    wanted = datetime.strptime(target_date, "%Y-%m-%d").date()
    zone = ZoneInfo(config.acquisition_timezone)
    if any(game.scheduled_utc.astimezone(zone).date() == wanted for game in unfiltered):
        raise GateUncertain("empty ESPN date-filtered scoreboard conflicts with current scoreboard")
    return []


def _cba_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    del config
    url = f"https://data-server.cbaleague.com/api/schedules/day?auto=0&day={target_date}"
    body = _read_limited_response(url, headers={"Referer": "https://www.cbaleague.com/data/"})
    key = os.environ.get("CBA_DECRYPT_KEY", "").encode() or PUBLIC_CBA_AES_KEY
    try:
        rows = decrypt_body(body, key)
    except Undecryptable as exc:
        raise GateUncertain("CBA schedule body did not decrypt") from exc
    if not isinstance(rows, list):
        raise GateUncertain("CBA schedule did not return a list")
    games: list[ScheduleGame] = []
    for row in rows:
        if not isinstance(row, dict):
            raise GateUncertain("CBA schedule has an invalid row")
        status_value = row.get("status")
        mapped_status = {1: "scheduled", "1": "scheduled", 4: "final", "4": "final"}.get(
            status_value
        )
        if mapped_status is None:
            raise GateUncertain("CBA schedule has an unknown status")
        raw_time = row.get("scheduleTime")
        if not isinstance(raw_time, str):
            raise GateUncertain("CBA schedule row has no scheduleTime")
        try:
            tip = datetime.strptime(raw_time.strip(), "%Y-%m-%d %H:%M:%S").replace(
                tzinfo=ZoneInfo("Asia/Shanghai")
            )
        except ValueError as exc:
            raise GateUncertain("CBA scheduleTime is invalid") from exc
        schedule_id = str(row.get("scheduleId") or "").strip()
        if not schedule_id:
            raise GateUncertain("CBA schedule row lacks scheduleId")
        games.append(ScheduleGame(tip.astimezone(timezone.utc), mapped_status, schedule_id))
    return games


def _fiba_schedule_payload(url: str) -> Any:
    """Read GDAP via header auth, refreshing the public page key only on need."""

    headers = {KEY_HEADER: PUBLIC_FALLBACK_SUBSCRIPTION_KEY, "Accept": "application/json"}
    try:
        return _read_json_response(url, headers=headers)
    except GateUncertain as first_error:
        # A rotated public browser key is operational uncertainty, not a reason
        # to silently skip a league. Refresh once without ever placing it in a
        # URL, then let any failure become the gate's normal fail-open path.
        page = _read_limited_response(KEY_PAGE_URL)
        refreshed = page_subscription_key(page)
        if not refreshed:
            raise GateUncertain("FIBA public key refresh did not expose a valid key") from first_error
        return _read_json_response(url, headers={KEY_HEADER: refreshed, "Accept": "application/json"})


def _fiba_elw_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    """Read assigned FIBA ELW fixtures and fail open on status/schema drift."""

    url = f"{GDAP_BASE_URL}/{SCHEDULE_ENDPOINT}?gdapCompetitionId={COMPETITION_ID}"
    payload = _fiba_schedule_payload(url)
    if not isinstance(payload, list) or not payload:
        raise GateUncertain("FIBA EuroLeague Women schedule lacks a non-empty game list")
    wanted = datetime.strptime(target_date, "%Y-%m-%d").date()
    games: list[ScheduleGame] = []
    for row in payload:
        if not isinstance(row, dict):
            raise GateUncertain("FIBA EuroLeague Women schedule has an invalid row")
        competition = row.get("competition")
        if (
            not isinstance(competition, dict)
            or competition.get("competitionId") != COMPETITION_ID
            or competition.get("competitionCode") not in {None, COMPETITION_CODE}
        ):
            raise GateUncertain("FIBA EuroLeague Women schedule row is outside competition 209127")
        if has_unassigned_team_slot(row):
            # Future bracket placeholders cannot create a ticket or canonical
            # fixture. Do not require their still-unassigned timestamp before
            # skipping them; the full provider emits the coverage telemetry.
            continue
        try:
            tip = game_datetime_utc(row.get("gameDateTimeUTC"))
        except ValueError as exc:
            raise GateUncertain("FIBA EuroLeague Women game timestamp is invalid") from exc
        if tip.astimezone(ZoneInfo(config.acquisition_timezone)).date() != wanted:
            continue
        if not has_complete_team_identity(row):
            raise GateUncertain("FIBA EuroLeague Women row has incomplete assigned team identity")
        status = schedule_status(row)
        if status not in {"scheduled", "live", "final"}:
            raise GateUncertain(f"FIBA EuroLeague Women row has {status} status")
        game_id = str(row.get("gameId") or "").strip()
        if not game_id:
            raise GateUncertain("FIBA EuroLeague Women row lacks gameId")
        games.append(ScheduleGame(tip, status, f"fiba-elw-{COMPETITION_ID}-{game_id}"))
    return games


def fetch_schedule(config: LeagueGateConfig, target_date: str) -> Iterable[ScheduleGame]:
    """Read one bounded league-day schedule, or fail open when none exists."""

    if config.schedule_source == "espn":
        return _espn_schedule(config, target_date)
    if config.schedule_source == "kbl":
        return _kbl_schedule(config, target_date)
    if config.schedule_source == "tbf":
        return _tbf_schedule(config, target_date)
    if config.schedule_source == "bleague":
        return _bleague_schedule(config, target_date)
    if config.schedule_source == "bcl":
        return _bcl_schedule(config, target_date)
    if config.schedule_source == "bal":
        return _bal_schedule(config, target_date)
    if config.schedule_source == "cba":
        return _cba_schedule(config, target_date)
    if config.schedule_source == "fiba_elw":
        return _fiba_elw_schedule(config, target_date)
    if config.schedule_source == "euroleague":
        # The adapter's only verified endpoint is a whole-season feed.  Do not
        # call it from a half-hour gate and claim this job is lightweight.
        raise GateUncertain("EuroLeague has no verified date-bounded schedule source")
    if config.schedule_source == "lba":
        # The calendar route always serves the league's current focus day; a
        # target date has no verified date-bounded source, so the gate fails
        # open rather than calling a focus-day fetch a probe.
        raise GateUncertain("Italy LBA has no verified date-bounded gate source")
    if config.schedule_source == "israel_ibl":
        # The board route is season-wide (cYear + board). A half-hour gate must
        # fail open rather than pull a whole season and call it a probe.
        raise GateUncertain("Israel Premier League has no verified date-bounded gate source")
    if config.schedule_source == "acb":
        # ACB's calendar route is season-scoped and identity is resolved from a
        # per-game header; a half-hour gate must fail open rather than fetch a
        # whole season and call it a probe.
        raise GateUncertain("ACB has no verified date-bounded gate source")
    if config.schedule_source == "aba_liga":
        # ABA's externally referenced calendar route is season-wide. The full
        # pipeline owns its identity/detail checks; a half-hour gate must fail
        # open rather than duplicate an HTML parser and call a whole season a
        # probe.
        raise GateUncertain("ABA League has no independent date-bounded gate source")
    raise GateUncertain("no gate schedule source is configured")
