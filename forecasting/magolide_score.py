"""Ma Golide score — cold-start standings scorer, shadow lane only.

This is a faithful Python port of the retired ``calculateMaGolideScore`` core
from ``docs/Forecaster.gs`` (deleted 2026-11-07 in ``de966196``), reduced to the
one capability the current calibrated machinery does not have: producing a
pre-tip read for a fixture **with zero reconciled results history**, using
standings-derived feature differences only.

Everything here is descriptive shadow evidence.  A score from this module can
never be a ticket, a price-backed pick, or a settlement grade:

* the feature weights are the retired tuner's *baseline*, not a measured fit
  for any lane this module may run on;
* the lanes it targets have no calibrated model of their own. Until 2026-11-25
  they also had no reconciled final to calibrate against; three of the four now
  settle, so what this module still lacks is a measurement, not a route;
* the repo's walk-forward and sample gates (30-row minimum, Wilson/shrinkage)
  apply to any future promotion and none of them can pass on a lane without
  results.

Weighted hand-picked form features are deliberately ``shadow_candidate_only``
in ``forecasting/feature_manifest.py``; this module keeps that posture.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence

from ingestion.league_registry import SPECS, LeagueSpec
from ingestion.models import Game, League

SCHEMA_VERSION = "magolide.magolide_score_shadow"
SOURCE_NAME = "magolide_score/shadow"

#: Defaults, NOT a fitted configuration.
#:
#: These numbers are the *validator's* defaults -- ``validateTier1Config`` in the
#: deleted ``docs/Forecaster.gs`` (line 5396: threshold default 25, confMin 50,
#: confMax 95, homeAdv 5) plus the empty-stub values handed to the Elite proposal
#: writer (line 3956). They are the values a config falls back to when a sheet
#: has no stored entry; they were never fitted to anything.
#:
#: The retired system did not run on a static config at all. Every analysis run
#: began with ``loadTier1Config(ss)`` -- a versioned config read from the
#: spreadsheet (line 2213, and again in the tuner at 3797, which logs
#: ``Current config version: <v>``). The Elite tuner then re-tuned that config
#: and ``writeEliteProposalSheet_`` promoted a new ``config_version`` with the
#: current config and ranked #2/#3 alternatives side by side (line 4731). So the
#: live behaviour came from a stored, versioned artifact that changed over time,
#: per sheet, and the constants below are only its cold-start fallback.
VALIDATOR_DEFAULTS: Dict[str, float] = {
    "rank": 0.0,
    "form": 2.5,
    "h2h": 1.5,
    "forebet": 3.0,
    "variance": 1.0,
    "pct": 3.0,
    "net_rtg": 2.0,
    "home_court": 1.0,
    "momentum": 1.0,
    "streak": 1.0,
    "home_adv": 5.0,
    "threshold": 25.0,
    "conf_min": 50.0,
    "conf_max": 95.0,
    "confidence_scale": 30.0,
}


def _safe_num(value: Any, fallback: float) -> float:
    """Port of ``safeNum_``: coerce to float, fall back on garbage."""
    if value is None:
        return fallback
    try:
        number = float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return fallback
    return number if math.isfinite(number) else fallback


def _clamp(number: float, low: float, high: float) -> float:
    return max(low, min(high, number))


@dataclass(frozen=True)
class StandingsRow:
    """The standings-derived inputs the cold-start scorer can use.

    Every field is optional: a missing standings row falls back to the same
    neutral defaults the retired script used (rank 15, pct 0.5, streak 0), so a
    brand-new lane scores honestly ``RISKY`` instead of inventing a read.
    """

    team: str
    rank: Optional[float] = None
    pct: Optional[float] = None
    net_rtg: Optional[float] = None
    home_pct: Optional[float] = None
    away_pct: Optional[float] = None
    streak: Optional[int] = None
    l10_wins: Optional[float] = None
    l10_losses: Optional[float] = None


@dataclass(frozen=True)
class MaGolideFeatures:
    """Signed, home-positive feature differences (port of Section A)."""

    rank_diff: float = 0.0
    pct_diff: float = 0.0
    net_rtg_diff: float = 0.0
    home_court_effect: float = 0.0
    momentum_diff: float = 0.0
    streak_diff: float = 0.0
    form_diff: float = 0.0
    h2h_diff: float = 0.0
    forebet_diff: float = 0.0
    variance_penalty: float = 0.0


@dataclass(frozen=True)
class MaGolideConfig:
    """Weights and gates (port of Section B/E/F)."""

    rank: float = VALIDATOR_DEFAULTS["rank"]
    form: float = VALIDATOR_DEFAULTS["form"]
    h2h: float = VALIDATOR_DEFAULTS["h2h"]
    forebet: float = VALIDATOR_DEFAULTS["forebet"]
    variance: float = VALIDATOR_DEFAULTS["variance"]
    pct: float = VALIDATOR_DEFAULTS["pct"]
    net_rtg: float = VALIDATOR_DEFAULTS["net_rtg"]
    home_court: float = VALIDATOR_DEFAULTS["home_court"]
    momentum: float = VALIDATOR_DEFAULTS["momentum"]
    streak: float = VALIDATOR_DEFAULTS["streak"]
    home_adv: float = VALIDATOR_DEFAULTS["home_adv"]
    threshold: float = VALIDATOR_DEFAULTS["threshold"]
    conf_min: float = VALIDATOR_DEFAULTS["conf_min"]
    conf_max: float = VALIDATOR_DEFAULTS["conf_max"]
    confidence_scale: float = VALIDATOR_DEFAULTS["confidence_scale"]
    use_legacy_sigmoid: bool = False
    cap_risky_confidence: bool = False
    risky_confidence_cap: Optional[float] = None
    #: Which stored configuration these weights came from. The retired system
    #: read a versioned config per sheet (``loadTier1Config``); a caller that
    #: holds such a config must say so here, because an unversioned default must
    #: never be mistaken for a measured fit.
    config_version: str = "unversioned-validator-defaults"


@dataclass(frozen=True)
class MaGolideVerdict:
    """Port of the scorer's return contract (Section G)."""

    score: float
    prediction: str  # HOME / AWAY / RISKY
    confidence: float
    #: The logistic quantity ``sigmoid(score / confidence_scale)`` for the
    #: predicted side. Named for what it is: a transform of an unmeasured
    #: hand-weighted score, not a calibrated probability. Anything that wants
    #: a probability must calibrate first, which a lane without results
    #: cannot do.
    uncalibrated_logistic_pct: float
    factor_breakdown: Dict[str, float]
    meta: Dict[str, Any]


def build_standings_features(
    home: Optional[StandingsRow],
    away: Optional[StandingsRow],
    *,
    h2h_diff: float = 0.0,
    forebet_diff: float = 0.0,
    variance_penalty: float = 0.0,
    momentum_diff: float = 0.0,
) -> MaGolideFeatures:
    """Standings-derived cold-start features, ported from the ``…Difference_`` helpers.

    Neutral fallbacks match the retired script: missing rank 15, missing pct
    0.5, missing streak 0.  A standings row absent entirely scores as that
    neutral row, so an unpopulated lane reads ``RISKY`` rather than guessing.
    """

    def _standings(row: Optional[StandingsRow]) -> StandingsRow:
        return row if row is not None else StandingsRow(team="")

    home_s = _standings(home)
    away_s = _standings(away)

    rank_diff = _safe_num(away_s.rank, 15.0) - _safe_num(home_s.rank, 15.0)
    pct_diff = (_safe_num(home_s.pct, 0.5) - _safe_num(away_s.pct, 0.5)) * 100.0
    net_rtg_diff = _safe_num(home_s.net_rtg, 0.0) - _safe_num(away_s.net_rtg, 0.0)
    home_court = (
        _safe_num(home_s.home_pct, 0.5) - _safe_num(away_s.away_pct, 0.5)
    ) * 100.0
    streak_diff = float(_safe_num(home_s.streak, 0)) - float(_safe_num(away_s.streak, 0))

    home_l10 = (
        _safe_num(home_s.l10_wins, 0.0) - _safe_num(home_s.l10_losses, 0.0)
        if home_s.l10_wins is not None or home_s.l10_losses is not None
        else 0.0
    )
    away_l10 = (
        _safe_num(away_s.l10_wins, 0.0) - _safe_num(away_s.l10_losses, 0.0)
        if away_s.l10_wins is not None or away_s.l10_losses is not None
        else 0.0
    )
    form_diff = home_l10 - away_l10 if home_l10 or away_l10 else streak_diff

    return MaGolideFeatures(
        rank_diff=rank_diff,
        pct_diff=pct_diff,
        net_rtg_diff=net_rtg_diff,
        home_court_effect=home_court,
        momentum_diff=momentum_diff,
        streak_diff=streak_diff,
        form_diff=form_diff,
        h2h_diff=h2h_diff,
        forebet_diff=forebet_diff,
        variance_penalty=variance_penalty,
    )


def score_features(features: MaGolideFeatures, config: MaGolideConfig) -> MaGolideVerdict:
    """Port of ``calculateMaGolideScore`` Sections C–G."""

    weighted = {
        "rank": config.rank * features.rank_diff,
        "pct": config.pct * features.pct_diff,
        "netRtg": config.net_rtg * features.net_rtg_diff,
        "homeCourt": config.home_court * features.home_court_effect,
        "momentum": config.momentum * features.momentum_diff,
        "streak": config.streak * features.streak_diff,
        "form": config.form * features.form_diff,
        "h2h": config.h2h * features.h2h_diff,
        "forebet": config.forebet * features.forebet_diff,
        # Variance is a PENALTY (always subtracted).
        "variance": config.variance * (-features.variance_penalty),
    }

    use_new_features = any(
        _nonzero(getattr(config, key))
        for key in ("pct", "net_rtg", "home_court", "momentum", "streak")
    )
    if use_new_features:
        score = sum(weighted.values())
    else:
        score = weighted["rank"] + weighted["form"] + weighted["h2h"] + weighted["forebet"] + weighted["variance"]
    score += config.home_adv
    weighted["homeAdv"] = config.home_adv

    if not math.isfinite(score):
        score = 0.0

    threshold = config.threshold if math.isfinite(config.threshold) and config.threshold >= 0 else 5.0
    abs_score = abs(score)
    if abs_score < threshold:
        prediction = "RISKY"
    elif score > 0:
        prediction = "HOME"
    else:
        prediction = "AWAY"

    # ``is not None`` rather than truthiness: safeNum_ keeps an explicit 0.
    conf_min = _clamp(config.conf_min if config.conf_min is not None else 50.0, 0.0, 100.0)
    conf_max = _clamp(config.conf_max if config.conf_max is not None else 95.0, 0.0, 100.0)
    if conf_min >= conf_max:
        conf_min, conf_max = 50.0, 95.0

    scale = config.confidence_scale if math.isfinite(config.confidence_scale) and config.confidence_scale > 0 else 30.0

    home_win_prob = 1.0 / (1.0 + math.exp(-score / scale))
    if prediction == "HOME":
        predicted_win_prob = home_win_prob
    elif prediction == "AWAY":
        predicted_win_prob = 1.0 - home_win_prob
    else:
        predicted_win_prob = 0.5

    if config.use_legacy_sigmoid:
        sigmoid_abs = 1.0 / (1.0 + math.exp(-abs_score / scale))
        confidence = conf_min + (conf_max - conf_min) * sigmoid_abs
    else:
        confidence = _clamp(predicted_win_prob * 100.0, conf_min, conf_max)

    if config.cap_risky_confidence and prediction == "RISKY":
        confidence = min(confidence, config.risky_confidence_cap if config.risky_confidence_cap is not None else conf_min + 5.0)

    confidence = _clamp(confidence, conf_min, conf_max)
    if not math.isfinite(confidence):
        confidence = conf_min

    meta = {
        "mode": "NEW" if use_new_features else "LEGACY",
        "confidence_mode": "LEGACY_SIGMOID" if config.use_legacy_sigmoid else "PROBABILITY",
        "scale_factor": scale,
        "threshold": threshold,
        "conf_bounds": {"min": conf_min, "max": conf_max},
        "risky_cap_applied": bool(config.cap_risky_confidence and prediction == "RISKY"),
        "home_win_prob_pct": round(home_win_prob * 1000, 1) / 10,
        "predicted_win_prob_pct": round(predicted_win_prob * 1000, 1) / 10,
        "weights_source": "validator_defaults_not_a_fitted_config",
        "config_version": config.config_version,
    }

    return MaGolideVerdict(
        score=round(score, 2),
        prediction=prediction,
        confidence=round(confidence, 1),
        uncalibrated_logistic_pct=round(predicted_win_prob * 1000, 1) / 10,
        factor_breakdown={key: round(value, 2) for key, value in weighted.items()},
        meta=meta,
    )


def _nonzero(value: float) -> bool:
    return abs(value) > 1e-9


# ---------------------------------------------------------------------------
# Shadow-lane adapter: fixtures-only leagues
# ---------------------------------------------------------------------------


def shadow_lane_leagues() -> List[LeagueSpec]:
    """The lanes this module may run on, straight from the registry contract.

    A lane qualifies when its schedule rides the aggregate-price fixtures table
    and it is active — the "fixtures exist, no model of their own" class.
    Retired and off-season lanes never qualify, whatever their sources say.

    The results source is deliberately *not* part of this rule any more. Three of
    these lanes gained a settlement route on 2026-11-25 (the aggregator's match
    page states the split beside the total), and a rule keyed on "no results
    route" would have stopped scoring exactly the lanes that can now be graded.
    What makes this module their scorer is that no calibrated model runs on them,
    and that has not changed.
    """

    return [
        spec
        for spec in SPECS
        if spec.schedule_source == "betexplorer_fixtures" and spec.status.value == "active"
    ]


@dataclass(frozen=True)
class ShadowPickRow:
    """One shadow observation.  Display/audit only — never a ticket row."""

    league: str
    game_id: Optional[str]
    home_team: str
    away_team: str
    scheduled_utc: Optional[str]
    verdict: MaGolideVerdict
    evidence: Dict[str, Any] = field(default_factory=dict)

    def as_card_row(self) -> str:
        v = self.verdict
        return (
            f"  [{SOURCE_NAME}] {self.home_team} vs {self.away_team}"
            f"{f' · KO {self.scheduled_utc}' if self.scheduled_utc else ''}\n"
            f"      MAGOLIDE_SCORE full_game -> {v.prediction} · score={v.score:+.2f}"
            f" · uncalibrated_logistic={v.uncalibrated_logistic_pct}%"
            f" · confidence={v.confidence}%\n"
            f"      bucket=WATCHLIST_SHADOW · price_evidence=NONE · acca_eligible=False"
            f" · grade=NOT_SETTLEABLE · lifecycle=SHADOW\n"
            f"      reasons=magolide_score_shadow,no_settlement_route,"
            f"cold_start_no_walk_forward · weights={v.meta['weights_source']}"
            f" · config_version={v.meta['config_version']}"
        )


def shadow_rows(
    league: League,
    games: Sequence[Game],
    standings: Mapping[str, StandingsRow],
    *,
    config: Optional[MaGolideConfig] = None,
) -> List[ShadowPickRow]:
    """Score a fixtures-only league's slate as shadow observations.

    Refuses any league that is not in :func:`shadow_lane_leagues` — the scorer
    must not silently run where results exist (the calibrated machinery owns
    that) or where acquisition is retired.
    """

    allowed = {spec.league for spec in shadow_lane_leagues()}
    if league not in allowed:
        raise ValueError(
            f"{league.value} is not a fixtures-only shadow lane "
            f"(eligible: {sorted(s.value for s in allowed)})"
        )

    resolved_config = config or MaGolideConfig()
    rows: List[ShadowPickRow] = []
    for game in games:
        home = standings.get(game.home_name)
        away = standings.get(game.away_name)
        features = build_standings_features(home, away)
        verdict = score_features(features, resolved_config)
        evidence = {
            "home_standings_present": home is not None,
            "away_standings_present": away is not None,
            "neutral_defaults_applied": home is None or away is None,
            "schema_version": SCHEMA_VERSION,
        }
        rows.append(
            ShadowPickRow(
                league=league.value,
                game_id=game.game_id,
                home_team=game.home_name,
                away_team=game.away_name,
                scheduled_utc=game.scheduled_utc,
                verdict=verdict,
                evidence=evidence,
            )
        )
    return rows


def render_shadow_card(league_value: str, rows: Iterable[ShadowPickRow], *, date_label: str) -> str:
    """A clearly-labelled shadow card, in the house card style but inert."""

    lines = [
        f"Ma Golide SHADOW SCORES — {league_value} {date_label}",
        "=" * 72,
        "SHADOW ONLY · cold-start standings scorer ported from the retired",
        "Forecaster.gs calculateMaGolideScore (weights = tuner baseline,",
        "unmeasured for this lane). Not picks, not tickets, not settlement",
        "evidence. The lane has no results route: nothing here can ever be",
        "graded, banked, or ROI-qualified.",
        "=" * 72,
        "",
        "WATCHLIST_SHADOW — COLD-START STANDINGS READ",
        "=" * 72,
    ]
    row_list = list(rows)
    if not row_list:
        lines.append("  (no fixtures supplied)")
    else:
        lines.extend(row.as_card_row() for row in row_list)
    lines += [
        "",
        f"Logical shadow rows: {len(row_list)} · tickets: 0 · settleable: 0",
        "Flat one-paper-unit ticket rules do not apply to shadow rows.",
    ]
    return "\n".join(lines) + "\n"


__all__ = [
    "SCHEMA_VERSION",
    "SOURCE_NAME",
    "VALIDATOR_DEFAULTS",
    "MaGolideConfig",
    "MaGolideFeatures",
    "MaGolideVerdict",
    "ShadowPickRow",
    "StandingsRow",
    "build_standings_features",
    "render_shadow_card",
    "score_features",
    "shadow_lane_leagues",
    "shadow_rows",
]
