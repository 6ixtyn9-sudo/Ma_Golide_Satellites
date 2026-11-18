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
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from ingestion.providers._cba_aes import Undecryptable, decrypt_body

SCHEMA_VERSION = "magolide.adaptive-gate"
DECISION_LEAD_MINUTES = 60
HTTP_TIMEOUT_SECONDS = 15
MAX_RESPONSE_BYTES = 5 * 1024 * 1024

# This is kept in the same order as the workflow matrix so its JSON ``include``
# rows can be consumed directly by GitHub Actions.
@dataclass(frozen=True)
class LeagueGateConfig:
    label: str
    league: str
    state_key: str
    acquisition_timezone: str
    schedule_source: str


LEAGUES: tuple[LeagueGateConfig, ...] = (
    LeagueGateConfig("NBA", "NBA", "nba", "America/New_York", "espn"),
    LeagueGateConfig("WNBA", "WNBA", "wnba", "America/New_York", "espn"),
    LeagueGateConfig("EuroLeague", "EuroLeague", "euroleague", "Europe/Paris", "euroleague"),
    LeagueGateConfig("Chinese CBA", "CBA", "cba", "Asia/Shanghai", "cba"),
    LeagueGateConfig("NCAA men", "NCAA_M", "ncaa_m", "America/New_York", "espn"),
    LeagueGateConfig("NCAA women", "NCAA_W", "ncaa_w", "America/New_York", "espn"),
)


class GateUncertain(RuntimeError):
    """A condition which must make the gate run the full pipeline."""


@dataclass(frozen=True)
class ScheduleGame:
    """The minimal schedule fact the gate needs; never canonical evidence."""

    scheduled_utc: datetime
    status: str


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

    A lock without the matching immutable receipt/no-bet sidecar is not enough
    to skip a run: cache recovery could otherwise turn missing evidence into a
    false no-work decision.
    """

    locks = state["frozen_by_date"]
    lock = locks.get(target_date)
    decision_path = state_dir / f"auto_tickets_{target_date}_decision.json"
    no_bet_path = state_dir / f"auto_tickets_{target_date}_no_bet.json"
    withdrawn_path = state_dir / f"auto_tickets_{target_date}_withdrawn.json"

    if decision_path.exists() and no_bet_path.exists():
        raise GateUncertain(f"conflicting ticket and no-bet decisions for {target_date}")
    if lock is None:
        if decision_path.exists() or no_bet_path.exists() or withdrawn_path.exists():
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
        _validate_sidecar(no_bet_path, target_date=target_date, league=league, no_bet=True)
        return True
    if no_bet_path.exists():
        raise GateUncertain(f"ticket lock conflicts with no-bet receipt for {target_date}")
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
        state, performance = _load_operator_summary(state_dir, league=config.league)
        frozen = date_is_frozen(
            state_dir, state, target_date=target_date, league=config.league
        )
        reasons.extend(_pending_reasons(performance))
    except GateUncertain as exc:
        return GateDecision(
            config, target_date, True, (f"state_uncertain: {exc}",), schedule_count
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

    if not frozen:
        deadline = now + timedelta(minutes=DECISION_LEAD_MINUTES)
        if any(
            game.status in {"scheduled", "live"}
            and game.scheduled_utc.astimezone(timezone.utc) <= deadline
            for game in games
        ):
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
    """Evaluate the fixed six-league matrix, with an explicit force-open mode."""

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


def _espn_scoreboard_url(config: LeagueGateConfig, target_date: str) -> str:
    yyyymmdd = target_date.replace("-", "")
    if config.league == "NBA":
        slug, suffix = "nba", "limit=100"
    elif config.league == "WNBA":
        slug, suffix = "wnba", "limit=100"
    elif config.league == "NCAA_M":
        slug, suffix = "mens-college-basketball", "groups=50&limit=500"
    elif config.league == "NCAA_W":
        slug, suffix = "womens-college-basketball", "groups=50&limit=500"
    else:  # pragma: no cover - guarded by LEAGUES
        raise GateUncertain("unsupported ESPN league")
    return (
        f"https://site.api.espn.com/apis/site/v2/sports/basketball/{slug}/"
        f"scoreboard?dates={yyyymmdd}&{suffix}"
    )


def _espn_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    payload = _read_json_response(_espn_scoreboard_url(config, target_date))
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
        games.append(
            ScheduleGame(
                _parse_utc(competition.get("date") or event.get("date"), context="ESPN event"),
                mapped_status,
            )
        )
    return games


def _cba_schedule(config: LeagueGateConfig, target_date: str) -> list[ScheduleGame]:
    del config
    url = f"https://data-server.cbaleague.com/api/schedules/day?auto=0&day={target_date}"
    body = _read_limited_response(url, headers={"Referer": "https://www.cbaleague.com/data/"})
    key = os.environ.get("CBA_DECRYPT_KEY", "").encode() or b"uVayqL4ONKjFbVzQ"
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
        games.append(ScheduleGame(tip.astimezone(timezone.utc), mapped_status))
    return games


def fetch_schedule(config: LeagueGateConfig, target_date: str) -> Iterable[ScheduleGame]:
    """Read one bounded league-day schedule, or fail open when none exists."""

    if config.schedule_source == "espn":
        return _espn_schedule(config, target_date)
    if config.schedule_source == "cba":
        return _cba_schedule(config, target_date)
    if config.schedule_source == "euroleague":
        # The adapter's only verified endpoint is a whole-season feed.  Do not
        # call it from a half-hour gate and claim this job is lightweight.
        raise GateUncertain("EuroLeague has no verified date-bounded schedule source")
    raise GateUncertain("no gate schedule source is configured")
