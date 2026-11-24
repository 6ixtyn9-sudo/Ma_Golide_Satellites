"""The config tuner: read graded bets, move the gates, stamp the result.

Ported from ``Config_Tuner.gs``
------------------------------
The retired tuner computed three things and turned them into advice on a
``Config_Tuner_Suggestions`` sheet:

* **Calibration** -- bucket graded picks by stated confidence, compare each bucket's
  hit rate to its target (``TARGET_ELITE_ACC`` 0.70, ``TARGET_STRONG_ACC`` 0.62,
  ``TARGET_MEDIUM_ACC`` 0.55), and compute a Brier score per bucket. Anything above
  ``BRIER_THRESHOLD_BAD`` 0.22 was called badly calibrated.
* **Directional bias** -- grade OVER and UNDER separately; a gap wider than 0.10
  over more than 50 settled picks meant the model was systematically one-sided.
* **League performance** -- per-league hit rate, with ``MIN_SAMPLES_FOR_LEAGUE_TUNING``
  20 before a league was worth talking about, suggesting
  ``league_weight_<LEAGUE> = 0.7`` for a league under 50%.

What changed
------------
The analysis is the same. The *output* is not advice -- you asked for the tuner to
apply its own findings -- so each rule now resolves to one bounded numeric move:

===============  ==========================================  ==========================
finding          move                                        bound
===============  ==========================================  ==========================
Brier > 0.22     ``confidence_scale`` + 10 (shrink harder,   1 .. 200, one step
                 the sheet's own "40+" advice from 30)
OVER/UNDER gap   ``ftouEdgeThreshold`` + 0.01 (demand more   0 .. 1, one step
> 0.10, n > 50   edge on totals where the model is skewed)
league < 50%     ``league_weights[LEAGUE] = 0.7``             0.1 .. 1.0
(n >= 20)
===============  ==========================================  ==========================

Every move goes through :meth:`forecasting.bet_config.BetConfig.with_overrides`, so
bounds and the tier invariants are enforced *after* the tuner speaks, and every
applied change is written to the ledger with before/after values. The tuner can
therefore never produce a configuration a human could not have typed, and the ledger
still answers "why is the system betting differently this week?".
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from forecasting.bet_config import DEFAULTS, KEYS, SHIPPED_PATH, BetConfig, load_config
from forecasting.config_ledger import DEFAULT_LEDGER_PATH, ConfigStamp, record, stamp_for

#: Carried verbatim from ``TUNER_CONFIG``.
CONF_BUCKETS = (45.0, 50.0, 55.0, 60.0, 65.0, 70.0, 75.0, 80.0, 85.0, 90.0, 95.0)
TARGET_ELITE_ACC = 0.70
TARGET_STRONG_ACC = 0.62
TARGET_MEDIUM_ACC = 0.55
MIN_SAMPLES_FOR_LEAGUE_TUNING = 20
BRIER_THRESHOLD_BAD = 0.22
DIRECTIONAL_MIN_SAMPLES = 50
DIRECTIONAL_GAP = 0.10

#: Where auto-applied overrides live. The shipped baseline stays untouched, so a
#: tuned run is always diffable against the defaults it started from.
TUNED_PATH = Path("localdata/magolide_bets.tuned.json")


@dataclass(frozen=True)
class GradedBet:
    """One settled bet, as the tuner needs it."""

    bet_type: str
    league: str
    stated_confidence: float
    won: Optional[bool]  # None when the bet pushed
    direction: Optional[str] = None  # "OVER" / "UNDER" for totals

    @property
    def push(self) -> bool:
        return self.won is None


@dataclass(frozen=True)
class LeagueReport:
    league: str
    count: int
    wins: int
    pushes: int
    win_rate: float
    brier: float


@dataclass(frozen=True)
class TuningReport:
    """What the graded sample says, before any change is made."""

    total: int
    overall_brier: Optional[float]
    buckets: Mapping[float, Mapping[str, float]] = field(default_factory=dict)
    directional: Mapping[str, Mapping[str, float]] = field(default_factory=dict)
    leagues: Mapping[str, LeagueReport] = field(default_factory=dict)
    changes: Mapping[str, Any] = field(default_factory=dict)
    reasons: Tuple[str, ...] = ()

    @property
    def anything_to_do(self) -> bool:
        return bool(self.changes)


def _bucket_for(confidence: float) -> float:
    return min(CONF_BUCKETS, key=lambda bucket: abs(bucket - confidence))


def analyse(bets: Sequence[GradedBet]) -> TuningReport:
    """The three analyses, in one pass, with no changes applied yet."""

    buckets: Dict[float, Dict[str, float]] = {
        bucket: {"count": 0.0, "wins": 0.0, "pushes": 0.0, "brier_sum": 0.0}
        for bucket in CONF_BUCKETS
    }
    directional: Dict[str, Dict[str, float]] = {
        "OVER": {"count": 0.0, "wins": 0.0, "pushes": 0.0},
        "UNDER": {"count": 0.0, "wins": 0.0, "pushes": 0.0},
    }
    league_acc: Dict[str, Dict[str, float]] = {}

    for bet in bets:
        outcome = 0.5 if bet.push else (1.0 if bet.won else 0.0)
        probability = bet.stated_confidence / 100.0

        bucket = buckets[_bucket_for(bet.stated_confidence)]
        bucket["count"] += 1
        bucket["wins"] += 0.0 if bet.push else float(bool(bet.won))
        bucket["pushes"] += 1.0 if bet.push else 0.0
        bucket["brier_sum"] += (probability - outcome) ** 2

        if bet.direction in directional:
            entry = directional[bet.direction]
            entry["count"] += 1
            entry["wins"] += 0.0 if bet.push else float(bool(bet.won))
            entry["pushes"] += 1.0 if bet.push else 0.0

        if bet.league:
            entry = league_acc.setdefault(
                bet.league, {"count": 0.0, "wins": 0.0, "pushes": 0.0, "brier_sum": 0.0}
            )
            entry["count"] += 1
            entry["wins"] += 0.0 if bet.push else float(bool(bet.won))
            entry["pushes"] += 1.0 if bet.push else 0.0
            entry["brier_sum"] += (probability - outcome) ** 2

    total_brier_sum = sum(bucket["brier_sum"] for bucket in buckets.values())
    total_count = sum(bucket["count"] for bucket in buckets.values())
    overall_brier = (total_brier_sum / total_count) if total_count else None

    for bucket in buckets.values():
        settled = bucket["count"] - bucket["pushes"]
        bucket["win_rate"] = (bucket["wins"] / settled) if settled else 0.0
        bucket["brier"] = (bucket["brier_sum"] / bucket["count"]) if bucket["count"] else 0.0

    for entry in directional.values():
        settled = entry["count"] - entry["pushes"]
        entry["win_rate"] = (entry["wins"] / settled) if settled else 0.0

    leagues: Dict[str, LeagueReport] = {}
    for name, entry in league_acc.items():
        settled = entry["count"] - entry["pushes"]
        leagues[name] = LeagueReport(
            league=name,
            count=int(entry["count"]),
            wins=int(entry["wins"]),
            pushes=int(entry["pushes"]),
            win_rate=(entry["wins"] / settled) if settled else 0.0,
            brier=(entry["brier_sum"] / entry["count"]) if entry["count"] else 0.0,
        )

    changes, reasons = _decide(overall_brier, directional, leagues, int(total_count))
    return TuningReport(
        total=int(total_count),
        overall_brier=overall_brier,
        buckets={bucket: dict(entry) for bucket, entry in buckets.items()},
        directional={key: dict(entry) for key, entry in directional.items()},
        leagues=leagues,
        changes=changes,
        reasons=tuple(reasons),
    )


def _decide(
    overall_brier: Optional[float],
    directional: Mapping[str, Mapping[str, float]],
    leagues: Mapping[str, LeagueReport],
    total: int,
) -> Tuple[Dict[str, Any], List[str]]:
    changes: Dict[str, Any] = {}
    reasons: List[str] = []

    if total == 0:
        return changes, ["no settled bets: nothing to tune"]

    if overall_brier is not None and overall_brier > BRIER_THRESHOLD_BAD:
        changes["confidence_scale"] = ("confidence_scale", 10.0)
        reasons.append(
            f"Brier {overall_brier:.3f} > {BRIER_THRESHOLD_BAD}: shrink probabilities harder"
        )

    over = directional["OVER"]
    under = directional["UNDER"]
    over_settled = over["count"] - over["pushes"]
    under_settled = under["count"] - under["pushes"]
    if (
        over_settled + under_settled > DIRECTIONAL_MIN_SAMPLES
        and abs(over["win_rate"] - under["win_rate"]) > DIRECTIONAL_GAP
    ):
        better = "OVER" if over["win_rate"] > under["win_rate"] else "UNDER"
        worse = "UNDER" if better == "OVER" else "OVER"
        changes["ftouEdgeThreshold"] = ("ftouEdgeThreshold", 0.01)
        reasons.append(
            f"{better} hits {max(over['win_rate'], under['win_rate']):.1%} vs "
            f"{min(over['win_rate'], under['win_rate']):.1%} for {worse}: demand more edge on totals"
        )

    for name, report in sorted(leagues.items()):
        if report.count >= MIN_SAMPLES_FOR_LEAGUE_TUNING and report.win_rate < 0.50:
            changes[f"league_weights/{name}"] = (name, 0.7)
            reasons.append(
                f"{name}: {report.win_rate:.1%} over {report.count} settled picks: discount to 0.7"
            )

    return changes, reasons


def apply(
    report: TuningReport,
    config: BetConfig,
    *,
    tuned_path: Path | str = TUNED_PATH,
    ledger_path: Path | str = DEFAULT_LEDGER_PATH,
    write: bool = True,
) -> Tuple[BetConfig, ConfigStamp, Dict[str, Tuple[Any, Any]]]:
    """Turn findings into a new configuration, clamped, stamped and ledgered.

    Returns ``(config, stamp, applied)`` where ``applied`` maps each changed key to
    ``(before, after)``. With nothing to do the same config comes back untouched --
    a tuner that fires on every quiet run is just noise in the ledger.
    """

    if not report.changes:
        note = "; ".join(report.reasons) or "no change"
        stamp = stamp_for(config, note=note) if not write else record(config, ledger_path, note=note)
        return config, stamp, {}

    changes: Dict[str, Any] = {}
    league_weights: Dict[str, float] = dict(config.league_weights)
    for key, (name, delta) in report.changes.items():
        if key.startswith("league_weights/"):
            league_weights[name] = float(delta)
        else:
            # Clamp to the key's bounds *before* handing the value to the config:
            # the sheet loader treats an out-of-range number as corrupt input and
            # resets it to the default, which for a tuner step would be a wild
            # swing. A step that reaches a bound simply stops there.
            spec = KEYS[key]
            target = float(config.get(key)) + float(delta)
            changes[key] = min(max(target, spec.minimum), spec.maximum)

    tuned = config.with_overrides(changes, source="tuner")
    if league_weights:
        tuned = tuned.with_leagues(league_weights, source="tuner")

    applied = {
        key: (config.get(key), tuned.get(key))
        for key in changes
        if tuned.get(key) != config.get(key)
    }
    for name, weight in league_weights.items():
        if config.league_weight(name) != weight:
            applied[f"league_weights/{name}"] = (config.league_weight(name), weight)

    note = "; ".join(
        f"{key} {before}->{after}" for key, (before, after) in sorted(applied.items())
    ) or "no change"
    if report.reasons:
        note = f"{note} | {'; '.join(report.reasons)}"

    if not write:
        # A dry run reports exactly what it would do and touches nothing.
        return tuned, stamp_for(tuned, note=note), applied

    if applied:
        _write_tuned(tuned, Path(tuned_path))

    stamp = record(tuned, ledger_path, note=note)
    return tuned, stamp, applied


def _write_tuned(config: BetConfig, path: Path) -> None:
    """Persist the tuned overrides, key-sorted so the file is diff-stable."""

    payload: Dict[str, Any] = {
        "_comment": (
            "Written by forecasting.config_tuner.apply() from settled results. "
            "Load after config/magolide_bets.json: load_config(SHIPPED_PATH, TUNED_PATH)."
        ),
        "version": config.version,
        "generated_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        # Anything that differs from the shipped defaults, not just what this run
        # touched: a second tuning run must not drop the first run's overrides.
        "keys": {
            name: value
            for name, value in config.as_dict().items()
            if value != DEFAULTS[name]
        },
        "league_weights": dict(sorted(config.league_weights.items())),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=False) + "\n")


def load_tuned(
    *, tuned_path: Path | str = TUNED_PATH, baseline: Optional[Path | str] = None
) -> BetConfig:
    """The baseline plus whatever the tuner last applied."""

    return load_config(baseline or SHIPPED_PATH, tuned_path)


def graded_from_rows(rows: Iterable[Mapping[str, Any]]) -> List[GradedBet]:
    """Build tuner input from ``{"bet_type", "league", "confidence", "won", ...}``.

    A convenience for callers holding a results table: it converts row dictionaries
    into :class:`GradedBet` and skips rows with no confidence, rather than guessing.
    """

    out: List[GradedBet] = []
    for row in rows:
        try:
            confidence = float(row.get("confidence"))
        except (TypeError, ValueError):
            continue
        won = row.get("won")
        out.append(
            GradedBet(
                bet_type=str(row.get("bet_type", "")),
                league=str(row.get("league", "")),
                stated_confidence=confidence,
                won=None if won is None else bool(won),
                direction=row.get("direction"),
            )
        )
    return out


__all__ = [
    "BRIER_THRESHOLD_BAD",
    "CONF_BUCKETS",
    "DIRECTIONAL_GAP",
    "DIRECTIONAL_MIN_SAMPLES",
    "MIN_SAMPLES_FOR_LEAGUE_TUNING",
    "TARGET_ELITE_ACC",
    "TARGET_MEDIUM_ACC",
    "TARGET_STRONG_ACC",
    "TUNED_PATH",
    "GradedBet",
    "LeagueReport",
    "TuningReport",
    "analyse",
    "apply",
    "graded_from_rows",
    "load_tuned",
]
