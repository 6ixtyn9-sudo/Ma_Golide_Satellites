"""Prequential market calibration and acca-eligibility policy."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import math
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from ingestion.identity import parse_utc

from assayer.assayer_engine import wilson_lower_bound

from .types import LEGACY_VERSION_TAG, MODEL_VERSION, Forecast, current_spelling

# Hard-reset: calibration must not borrow from pre-feature-freeze history.
# Any downward threshold change would be a CRITICAL regime violation; this
# epoch is the honest alternative — count n only from the freeze date.
CALIBRATION_EPOCH_UTC = "2026-11-10T00:00:00Z"

SegmentKey = Tuple[str, ...]

# Reserved 5th/6th key dimension for diagnostics-only rows that did not clear
# team-history sufficiency; these profiles are never applied.
DIAG_POOL_TAG = "team_history_insufficient"


def _requires_team_history(forecast: Forecast) -> bool:
    """Only the history model (and history-based hybrids) need team evidence."""
    source = current_spelling(forecast.source)  # history stored before the rename
    return source == MODEL_VERSION or source.endswith(f"+{MODEL_VERSION}")


def config_fingerprint(components: Mapping[str, Any]) -> str:
    """Deterministic 12-hex fingerprint of prediction-relevant configuration.

    The fingerprint covers exactly the components passed in: callers must pass
    only the parameters that change prediction behaviour, so irrelevant or
    volatile metadata (paths, timestamps, seeds, host data) cannot enter.
    Mapping insertion order is irrelevant (canonical JSON with sorted keys),
    and top-level float-integers normalize to integers, so ``30`` and ``30.0``
    fingerprint identically. Any non-JSON-serializable component raises
    ``ValueError`` instead of silently degrading the version identity.
    """
    normalized = {
        str(key): int(value) if isinstance(value, float) and value.is_integer() else value
        for key, value in components.items()
    }
    try:
        payload = json.dumps(normalized, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"config components must be JSON-serializable: {exc}") from exc
    return hashlib.sha256(payload.encode()).hexdigest()[:12]


@dataclass(frozen=True)
class CalibrationConfig:
    prior_strength: float = 30.0
    min_samples: int = 30
    min_probability: float = 0.50
    max_push_probability: float = 0.005
    probability_floor: float = 0.01
    probability_ceiling: float = 0.99
    calibration_epoch_utc: Optional[str] = CALIBRATION_EPOCH_UTC

    def __post_init__(self) -> None:
        if self.prior_strength <= 0 or self.min_samples < 1:
            raise ValueError("calibration prior and sample minimum must be positive")
        if not 0 <= self.min_probability <= 1:
            raise ValueError("minimum probability must be between zero and one")
        if not 0 <= self.max_push_probability < 1:
            raise ValueError("push tolerance must be in [0, 1)")
        if not 0 < self.probability_floor < self.probability_ceiling < 1:
            raise ValueError("invalid calibrated probability bounds")


@dataclass(frozen=True)
class CalibrationProfile:
    key: SegmentKey
    wins: int
    losses: int
    pushes: int
    mean_raw_probability: float
    observed_win_rate: float
    adjustment: float
    wilson_lower_95: float
    brier_score: float
    log_loss: float
    min_samples: int

    @property
    def n(self) -> int:
        return self.wins + self.losses

    @property
    def sufficient(self) -> bool:
        return self.n >= self.min_samples

    def calibrate(self, probability: float, config: CalibrationConfig) -> float:
        adjusted = probability + self.adjustment
        return min(config.probability_ceiling, max(config.probability_floor, adjusted))

    def to_dict(self) -> Dict[str, Any]:
        # Diagnostic rows carry the reserved DIAG_POOL_TAG as the final key
        # dimension (6-tuples since version-aware segmentation; the version tag
        # precedes it). Diagnostic profiles are visible in walk_forward.segments
        # but are structurally never applied: apply_profiles matches the
        # forecast's versioned 5-tuple key only.
        is_diag = bool(self.key) and self.key[-1] == DIAG_POOL_TAG
        result = {
            "league": self.key[0],
            "market": self.key[1],
            "period": self.key[2],
            "source": self.key[3],
            "calibration_version": self.key[4] if len(self.key) > 4 else LEGACY_VERSION_TAG,
            "n": self.n,
            "wins": self.wins,
            "losses": self.losses,
            "pushes": self.pushes,
            "mean_raw_probability": round(self.mean_raw_probability, 6),
            "observed_win_rate": round(self.observed_win_rate, 6),
            "probability_adjustment": round(self.adjustment, 6),
            "wilson_lower_95": round(self.wilson_lower_95, 6),
            "brier_score": round(self.brier_score, 6),
            "log_loss": round(self.log_loss, 6),
            "minimum_samples": self.min_samples,
            "sufficient_for_accas": False if is_diag else self.sufficient,
            "pool": "league_prior_in_diagnostics_only" if is_diag else "eligible",
        }
        if is_diag:
            result["team_history_sufficient"] = False
            result["segment_key_5d"] = list(self.key)
        return result


def _log_loss(probability: float, won: bool) -> float:
    bounded = min(1 - 1e-12, max(1e-12, probability))
    return -math.log(bounded if won else 1 - bounded)


_PROFILE_STAGE_FIELDS = (
    "input_records",
    "ungraded_records",
    "graded_records",
    "invalid_scheduled_times",
    "excluded_before_epoch",
    "epoch_eligible_records",
    "team_history_insufficient_records",
    "team_history_insufficient_samples",
    "team_history_insufficient_pushes",
    "profile_pool_records",
    "profile_pool_samples",
    "profile_pool_pushes",
)


def _empty_profile_stage_counts() -> Dict[str, int]:
    return {field: 0 for field in _PROFILE_STAGE_FIELDS}


def _segment_dimensions(key: SegmentKey) -> Dict[str, str]:
    return {
        "league": key[0],
        "market": key[1],
        "period": key[2],
        "source": key[3],
        "calibration_version": key[4] if len(key) > 4 else LEGACY_VERSION_TAG,
    }


def _group_profile_records(
    records: Iterable[Forecast], cfg: CalibrationConfig
) -> tuple[
    Dict[SegmentKey, List[Forecast]],
    Dict[SegmentKey, List[Forecast]],
    Dict[str, Any],
]:
    """Apply the exact profile admission rules and retain stage counts."""
    epoch_raw = getattr(cfg, "calibration_epoch_utc", CALIBRATION_EPOCH_UTC)
    try:
        epoch = parse_utc(epoch_raw) if epoch_raw else None
    except Exception:
        epoch = None

    grouped: Dict[SegmentKey, List[Forecast]] = {}
    grouped_insufficient: Dict[SegmentKey, List[Forecast]] = {}
    totals = _empty_profile_stage_counts()
    by_segment: Dict[SegmentKey, Dict[str, int]] = {}

    def increment(key: SegmentKey, field: str) -> None:
        totals[field] += 1
        counts = by_segment.setdefault(key, _empty_profile_stage_counts())
        counts[field] += 1

    for record in records:
        key = record.versioned_segment_key
        increment(key, "input_records")
        if record.outcome not in {"win", "loss", "push"}:
            increment(key, "ungraded_records")
            continue
        increment(key, "graded_records")

        # Hard-reset: ignore any graded record before the feature-freeze epoch.
        if epoch is not None:
            try:
                scheduled = parse_utc(record.scheduled_utc)
            except Exception:
                increment(key, "invalid_scheduled_times")
                continue
            if scheduled < epoch:
                increment(key, "excluded_before_epoch")
                continue
        increment(key, "epoch_eligible_records")

        # Version-aware segmentation: evidence from materially different model
        # or config versions carries a distinct version tag and never merges.
        # Records without version information carry the explicit ``legacy`` tag,
        # which reproduces the historical grouping exactly.
        # Calibrate the population that may actually supply accas. League-prior
        # fallback rows remain visible in diagnostics, but cannot lend their
        # sample count or reliability adjustment to team-specific selections.
        if _requires_team_history(record) and (record.team_history or {}).get(
            "sufficient_for_accas"
        ) is False:
            grouped_insufficient.setdefault(key, []).append(record)
            increment(key, "team_history_insufficient_records")
            increment(
                key,
                "team_history_insufficient_pushes"
                if record.outcome == "push"
                else "team_history_insufficient_samples",
            )
            continue

        grouped.setdefault(key, []).append(record)
        increment(key, "profile_pool_records")
        increment(
            key,
            "profile_pool_pushes" if record.outcome == "push" else "profile_pool_samples",
        )

    diagnostics = {
        "calibration_epoch_utc": epoch_raw,
        "calibration_epoch_enabled": epoch is not None,
        **totals,
        "segments": [
            {**_segment_dimensions(key), **by_segment[key]}
            for key in sorted(by_segment)
        ],
    }
    return grouped, grouped_insufficient, diagnostics


def calibration_population_diagnostics(
    records: Iterable[Forecast], config: CalibrationConfig | None = None
) -> Dict[str, Any]:
    """Count record inclusion and exclusion at each calibration-profile stage.

    The same grouping routine powers :func:`build_profiles`, so these counts
    cannot drift from its graded-outcome, epoch, and team-history filters.
    ``profile_pool_samples`` counts wins and losses; pushes are reported
    separately and never contribute to profile ``n``.
    """
    cfg = config or CalibrationConfig()
    _, _, diagnostics = _group_profile_records(records, cfg)
    return diagnostics


def build_profiles(
    records: Iterable[Forecast], config: CalibrationConfig | None = None
) -> Dict[SegmentKey, CalibrationProfile]:
    """Fit additive reliability adjustments from graded past forecasts only."""
    cfg = config or CalibrationConfig()
    grouped, grouped_insufficient, _ = _group_profile_records(records, cfg)
    profiles: Dict[SegmentKey, CalibrationProfile] = {}
    for key, rows in grouped.items():
        graded = [row for row in rows if row.outcome in {"win", "loss"}]
        pushes = sum(row.outcome == "push" for row in rows)
        n = len(graded)
        if not n:
            continue
        wins = sum(row.outcome == "win" for row in graded)
        raw_mean = sum(row.raw_probability for row in graded) / n
        observed = wins / n
        # Small samples stay near the model's own mean confidence. The
        # correction approaches observed-minus-claimed only as evidence grows.
        weight = n / (n + cfg.prior_strength)
        adjustment = weight * (observed - raw_mean)
        profiles[key] = CalibrationProfile(
            key=key,
            wins=wins,
            losses=n - wins,
            pushes=pushes,
            mean_raw_probability=raw_mean,
            observed_win_rate=observed,
            adjustment=adjustment,
            wilson_lower_95=wilson_lower_bound(wins, n),
            brier_score=sum(
                (row.raw_probability - float(row.outcome == "win")) ** 2 for row in graded
            )
            / n,
            log_loss=sum(_log_loss(row.raw_probability, row.outcome == "win") for row in graded)
            / n,
            min_samples=cfg.min_samples,
        )
    # Diagnostics-only parallel pool for rows that did not clear team-history.
    # These profiles are visible in walk_forward.segments with the reserved
    # DIAG_POOL_TAG as the final key dimension and are structurally never
    # applied (apply_profiles matches the forecast's versioned 5-tuple only).
    for key, rows in grouped_insufficient.items():
        graded = [row for row in rows if row.outcome in {"win", "loss"}]
        pushes = sum(row.outcome == "push" for row in rows)
        n = len(graded)
        if not n:
            continue
        wins = sum(row.outcome == "win" for row in graded)
        raw_mean = sum(row.raw_probability for row in graded) / n
        observed = wins / n
        weight = n / (n + cfg.prior_strength)
        adjustment = weight * (observed - raw_mean)
        diag_key = (*key, DIAG_POOL_TAG)
        profiles[diag_key] = CalibrationProfile(
            key=diag_key,
            wins=wins,
            losses=n - wins,
            pushes=pushes,
            mean_raw_probability=raw_mean,
            observed_win_rate=observed,
            adjustment=adjustment,
            wilson_lower_95=wilson_lower_bound(wins, n),
            brier_score=sum(
                (row.raw_probability - float(row.outcome == "win")) ** 2 for row in graded
            )
            / n,
            log_loss=sum(_log_loss(row.raw_probability, row.outcome == "win") for row in graded)
            / n,
            min_samples=cfg.min_samples,
        )
    return profiles


class IncrementalProfileBuilder:
    """O(1)-per-record sufficient-statistics tracker identical to ``build_profiles``.

    Used by prequential walk-forward replay so advancing the availability
    cutoff over ``N`` historical finals updates only newly settled forecasts
    instead of rescanning and refitting every bucket from scratch on each game.
    """

    def __init__(self, config: CalibrationConfig | None = None) -> None:
        self.cfg = config or CalibrationConfig()
        epoch_raw = getattr(self.cfg, "calibration_epoch_utc", CALIBRATION_EPOCH_UTC)
        try:
            self._epoch = parse_utc(epoch_raw) if epoch_raw else None
        except Exception:
            self._epoch = None
        # bucket_key -> [wins, losses, pushes, sum_raw, sum_brier, sum_log_loss]
        self._stats: Dict[SegmentKey, List[Any]] = {}
        self._cache: Dict[SegmentKey, CalibrationProfile] = {}

    def add(self, record: Forecast) -> None:
        if record.outcome not in {"win", "loss", "push"}:
            return
        if self._epoch is not None:
            try:
                scheduled = parse_utc(record.scheduled_utc)
            except Exception:
                return
            if scheduled < self._epoch:
                return
        key: SegmentKey = record.versioned_segment_key
        if (
            _requires_team_history(record)
            and (record.team_history or {}).get("sufficient_for_accas") is False
        ):
            bucket: SegmentKey = (*key, DIAG_POOL_TAG)
        else:
            bucket = key
        stats = self._stats.get(bucket)
        if stats is None:
            stats = [0, 0, 0, 0.0, 0.0, 0.0]
            self._stats[bucket] = stats
        if record.outcome == "push":
            stats[2] += 1
        else:
            won = record.outcome == "win"
            if won:
                stats[0] += 1
            else:
                stats[1] += 1
            prob = record.raw_probability
            stats[3] += prob
            stats[4] += (prob - float(won)) ** 2
            stats[5] += _log_loss(prob, won)
        self._cache.pop(bucket, None)

    def get(
        self, key: SegmentKey, default: Optional[CalibrationProfile] = None
    ) -> Optional[CalibrationProfile]:
        cached = self._cache.get(key)
        if cached is not None:
            return cached
        stats = self._stats.get(key)
        if stats is None:
            return default
        wins, losses, pushes, sum_raw, sum_brier, sum_log_loss = stats
        n = wins + losses
        if not n:
            return default
        raw_mean = sum_raw / n
        observed = wins / n
        weight = n / (n + self.cfg.prior_strength)
        adjustment = weight * (observed - raw_mean)
        profile = CalibrationProfile(
            key=key,
            wins=wins,
            losses=losses,
            pushes=pushes,
            mean_raw_probability=raw_mean,
            observed_win_rate=observed,
            adjustment=adjustment,
            wilson_lower_95=wilson_lower_bound(wins, n),
            brier_score=sum_brier / n,
            log_loss=sum_log_loss / n,
            min_samples=self.cfg.min_samples,
        )
        self._cache[key] = profile
        return profile

    def profiles(self) -> Dict[SegmentKey, CalibrationProfile]:
        out: Dict[SegmentKey, CalibrationProfile] = {}
        for key in self._stats:
            profile = self.get(key)
            if profile is not None:
                out[key] = profile
        return out


#: Reason stamped on a row that is a calibration twin of a better-evidenced row.
TWIN_SUPERSEDED_REASON = "superseded_by_calibration_twin"
#: Advisory stamped when the twins would have disagreed about eligibility.
TWIN_DISAGREEMENT_ADVISORY = "calibration_twin_disagreement"


def _wager_identity(forecast: Forecast) -> tuple:
    """The identity of a *bet*, independent of which model produced it.

    Two rows that agree on all six fields price the same wager. When they come
    from different calibration segments they are twins, not two bets.
    """

    return (
        forecast.league,
        forecast.game_id,
        forecast.market,
        forecast.period,
        forecast.selection,
        forecast.line,
    )


def _calibration_evidence_rank(forecast: Forecast) -> tuple:
    """Order twins by how much verified calibration skill stands behind them.

    Sufficiency first (``n >= minimum_samples``), then lower walk-forward
    ``brier_score`` among sufficient twins so a sharper calibrated model (e.g.
    ``provider:espn``) is never suppressed merely because a weaker baseline
    replayed more early-season games, then sample size, probability, and row
    id so the winner never depends on emission order. Deliberately *not* "the
    stricter verdict wins": that reads as cautious, but it hands a veto to
    whichever model is worse, and a lane could then block every bet by also
    emitting a badly-calibrated pessimistic twin.
    """

    calibration = forecast.calibration or {}
    sample = calibration.get("n")
    sufficient = 1 if calibration.get("sufficient_for_accas") is True else 0
    brier = calibration.get("brier_score")
    has_brier = (
        sufficient == 1
        and isinstance(brier, (int, float))
        and not isinstance(brier, bool)
        and math.isfinite(float(brier))
    )
    return (
        sufficient,
        1 if has_brier else 0,
        -float(brier) if has_brier else float("-inf"),
        int(sample) if isinstance(sample, int) and not isinstance(sample, bool) else 0,
        float(forecast.calibrated_probability or 0.0),
        str(forecast.forecast_id),
    )


def resolve_calibration_twins(forecasts: Iterable[Forecast]) -> List[Forecast]:
    """Make one betting key yield exactly one wager-eligible row.

    A lane can emit one priced row per calibration segment -- on 2026-11-25
    EuroLeague did it for every priced key: a ``provider_prediction`` twin at
    p≈0.48 (blocked by the probability floor) and a ``team_history_empirical``
    twin at p≈0.55 (eligible), same fixture, market, period, side and price.
    Evaluating them independently made the floor selectable by choosing a
    calibration: the same bet was eligible or not depending on which twin the
    gate happened to read, and the pick card showed both verdicts at once.

    The better-evidenced row wins (:func:`_calibration_evidence_rank`); every
    other twin stays on the card as audit-only, carrying
    ``superseded_by_calibration_twin``, the winner's id and version, and -- when
    the verdicts would have differed -- a ``calibration_twin_disagreement``
    advisory so the disagreement is stated rather than resolved silently.
    """

    rows = list(forecasts)
    groups: Dict[tuple, List[Forecast]] = {}
    for forecast in rows:
        groups.setdefault(_wager_identity(forecast), []).append(forecast)
    for group in groups.values():
        if len(group) < 2:
            continue
        canonical = max(group, key=_calibration_evidence_rank)
        canonical_eligible = bool((canonical.eligibility or {}).get("acca_eligible"))
        for row in group:
            if row is canonical:
                continue
            eligibility = dict(row.eligibility or {})
            reasons = list(eligibility.get("reasons") or [])
            if TWIN_SUPERSEDED_REASON not in reasons:
                reasons.append(TWIN_SUPERSEDED_REASON)
            advisories = list(eligibility.get("advisories") or [])
            if bool(eligibility.get("acca_eligible")) != canonical_eligible:
                if TWIN_DISAGREEMENT_ADVISORY not in advisories:
                    advisories.append(TWIN_DISAGREEMENT_ADVISORY)
            eligibility["reasons"] = reasons
            eligibility["advisories"] = advisories
            eligibility["acca_eligible"] = False
            eligibility["superseded_by_forecast_id"] = canonical.forecast_id
            eligibility["superseded_by_calibration_version"] = canonical.calibration_version
            row.eligibility = eligibility
    return rows


def apply_profiles(
    forecasts: Iterable[Forecast],
    profiles: Mapping[SegmentKey, CalibrationProfile],
    config: CalibrationConfig | None = None,
) -> List[Forecast]:
    cfg = config or CalibrationConfig()
    calibrated: List[Forecast] = []
    for forecast in forecasts:
        profile = profiles.get(forecast.versioned_segment_key)
        if profile is None:
            probability = forecast.raw_probability
            forecast.calibration = {
                "status": "no_history",
                "n": 0,
                "adjustment": 0.0,
                "sufficient_for_accas": False,
                "calibration_version": forecast.version_tag,
            }
        else:
            probability = profile.calibrate(forecast.raw_probability, cfg)
            forecast.calibration = profile.to_dict()
            forecast.calibration["status"] = "calibrated" if profile.sufficient else "provisional"
        forecast.calibrated_probability = probability
        reasons: List[str] = []
        if forecast.book is None:
            reasons.append("no_exact_bookmaker_market")
        elif forecast.decimal_price is None:
            reasons.append("missing_selected_decimal_price")
        elif not forecast.book.ticket_price_eligible:
            # Still emitted for the provenances the paper lane does not accept
            # (a price-less record, an exchange match, a prohibited source).
            # Named-book and aggregate-consensus quotes both clear the gate;
            # which of the two it was stays on ``price_provenance_label``.
            reasons.append("price_not_named_bookmaker_quote")
        # Evidence-thinness is an advisory, not a refusal. Operator decision
        # 2026-11-22: a lane must be able to publish a pick before it has
        # settled anything, otherwise walk-forward can never accumulate the
        # very sample it is waiting for. The shortfall is still named on every
        # row (``advisories``), still drives the PRIOR_HEAVY/CAUTION buckets,
        # and still keeps ``calibration.sufficient_for_accas`` false so the
        # recency-weighted ROI band filter sees the thin sample.
        advisories: List[str] = []
        if profile is None or not profile.sufficient:
            advisories.append("insufficient_walk_forward_samples")
        if _requires_team_history(forecast) and (forecast.team_history or {}).get(
            "sufficient_for_accas"
        ) is not True:
            advisories.append("insufficient_team_specific_history")
        # These two remain hard refusals: they are facts about *this* bet (a
        # push rule the receipt cannot express, a probability the configured
        # floor rejects), not about how much history the lane has collected.
        if forecast.push_probability > cfg.max_push_probability:
            reasons.append("push_risk_requires_leg_void_rules")
        if probability < cfg.min_probability:
            reasons.append("below_calibrated_probability_floor")
        forecast.eligibility = {
            "acca_eligible": not reasons,
            "reasons": reasons,
            "advisories": advisories,
            "evidence_state": (
                "frozen_pre_tip"
                if forecast.book is not None and forecast.book.frozen
                else "current_pre_tip_unfrozen"
            ),
        }
        calibrated.append(forecast)
    return resolve_calibration_twins(calibrated)


def probability_metrics(
    records: Sequence[Forecast], *, calibrated: bool = False
) -> Dict[str, Optional[float] | int]:
    graded = [row for row in records if row.outcome in {"win", "loss"}]
    if not graded:
        return {"n": 0, "brier_score": None, "log_loss": None, "accuracy": None}
    probabilities = []
    for row in graded:
        value = row.calibrated_probability if calibrated else row.raw_probability
        if value is None:
            value = row.raw_probability
        probabilities.append(value)
    return {
        "n": len(graded),
        "brier_score": sum(
            (probability - float(row.outcome == "win")) ** 2
            for probability, row in zip(probabilities, graded)
        )
        / len(graded),
        "log_loss": sum(
            _log_loss(probability, row.outcome == "win")
            for probability, row in zip(probabilities, graded)
        )
        / len(graded),
        "accuracy": sum(row.outcome == "win" for row in graded) / len(graded),
    }


__all__ = [
    "CalibrationConfig",
    "CalibrationProfile",
    "SegmentKey",
    "apply_profiles",
    "build_profiles",
    "calibration_population_diagnostics",
    "config_fingerprint",
    "probability_metrics",
]
