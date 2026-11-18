"""Availability-adjusted empirical forecast surface from point-in-time ESPN context.

This module promotes the previously candidate-only ESPN injury/minutes evidence
into a bounded, deterministic adjustment applied on a *separate* empirical
forecast surface (``context_availability``). The base empirical surface and all
provider surfaces remain bit-for-bit unchanged.

Method (no invented benchmark coefficient): for each team, the window-average
scoring contribution of every player is ``appearance_rate x per-appearance
points`` from ``build_espn_player_context``. Players ESPN reports as Out /
Out For Season / Inactive are confirmed out for the target game, so their
window-average contribution is the estimated scoring loss. The team's expected
points are scaled by ``1 - loss_share`` (bounded by ``max_scale``). Replacement
production is not modeled, which overstates the loss; the surface carries its
own calibration segment so walk-forward results can validate (or retire) it.

Guards inherited from the evidence layer: the injury snapshot must be visible
by the forecast cutoff and before tip; a window of prior boxscores must exist
for BOTH teams (asymmetric evidence is rejected rather than half-applied);
missing injury entries are never read as "healthy"; a zero adjustment emits no
forecast at all.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
import math
from typing import Any, Dict, Optional

from .types import MODEL_VERSION, id_spelling

CONTEXT_SURFACE = "context_availability"
CONTEXT_SOURCE = f"{CONTEXT_SURFACE}+{MODEL_VERSION}"


@dataclass(frozen=True)
class ContextAdjustmentConfig:
    """Knobs for the availability surface. Kept out of ForecastConfig on
    purpose: ForecastConfig fields fingerprint the *base* calibration
    version, and these knobs must never silently re-version base or provider
    segments."""

    points_coefficient: float = 1.0
    max_scale: float = 0.05
    min_window_games: int = 3
    min_rotation_players: int = 6

    def __post_init__(self) -> None:
        for name, value in (("points_coefficient", self.points_coefficient),):
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ValueError(f"{name} must be a finite number")
            if not 0.0 < float(value) <= 2.0:
                raise ValueError(f"{name} must be in (0, 2]")
        if (
            isinstance(self.max_scale, bool)
            or not isinstance(self.max_scale, (int, float))
            or not 0.0 < float(self.max_scale) <= 0.25
        ):
            raise ValueError("max_scale must be in (0, 0.25]")
        for name in ("min_window_games", "min_rotation_players"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ValueError(f"{name} must be a positive integer")


@dataclass(frozen=True)
class ContextAdjustment:
    home_scale: float
    away_scale: float
    home_evidence: Dict[str, Any]
    away_evidence: Dict[str, Any]

    def __post_init__(self) -> None:
        for name in ("home_scale", "away_scale"):
            value = getattr(self, name)
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise ValueError(f"{name} must be a finite number")
            if not 0.0 < float(value) <= 1.0:
                raise ValueError(f"{name} must be in (0, 1]")

    def to_metadata(self) -> Dict[str, Any]:
        """Deterministic, JSON-safe description attached to each forecast."""
        return {
            "surface": CONTEXT_SURFACE,
            "home_scale": round(float(self.home_scale), 6),
            "away_scale": round(float(self.away_scale), 6),
            "home": self.home_evidence,
            "away": self.away_evidence,
        }


def _team_lost_points_share(
    team: Any, config: ContextAdjustmentConfig
) -> Optional[Dict[str, Any]]:
    """Return the bounded lost-scoring share for one team, or None.

    None means "this team's evidence is insufficient" and rejects the whole
    adjustment: applying the scale to only one side would fabricate a margin
    signal out of missingness.
    """
    if not isinstance(team, dict):
        return None
    window = team.get("historical_boxscore_games")
    players = team.get("recent_players")
    if not isinstance(window, int) or isinstance(window, bool) or window < config.min_window_games:
        return None
    if not isinstance(players, list) or len(players) < config.min_rotation_players:
        return None
    out_ids = set(team.get("confirmed_out_player_ids") or [])
    total_contribution = 0.0
    lost_contribution = 0.0
    usable_players = 0
    for row in players:
        if not isinstance(row, dict):
            continue
        appearances = row.get("appearances")
        mean_points = row.get("mean_points")
        if (
            not isinstance(appearances, int)
            or isinstance(appearances, bool)
            or appearances <= 0
            or appearances > window
            or not isinstance(mean_points, (int, float))
            or isinstance(mean_points, bool)
            or not math.isfinite(float(mean_points))
            or mean_points < 0
        ):
            continue
        usable_players += 1
        contribution = (appearances / window) * float(mean_points)
        total_contribution += contribution
        if row.get("athlete_id") in out_ids:
            lost_contribution += contribution
    if usable_players < config.min_rotation_players or total_contribution <= 0.0:
        return None
    raw_share = lost_contribution / total_contribution
    if not math.isfinite(raw_share) or raw_share < 0.0:
        return None
    bounded = min(max(raw_share * config.points_coefficient, 0.0), config.max_scale)
    if not math.isfinite(bounded):
        return None
    return {
        "window_games": window,
        "players_with_minutes": len(players),
        "usable_players": usable_players,
        "confirmed_out_players": len(out_ids),
        "raw_lost_points_share": round(raw_share, 6),
        "applied_lost_points_share": round(bounded, 6),
        "scale": round(1.0 - bounded, 6),
    }


def build_context_adjustment(
    record: Any,
    *,
    config: ContextAdjustmentConfig | None = None,
) -> Optional[ContextAdjustment]:
    """Convert a candidate context record into a bounded availability scale.

    Returns None unless a point-in-time injury snapshot exists AND both teams
    carry sufficient window evidence AND at least one confirmed-out player
    actually reduces expected scoring. Zero adjustments never emit forecasts:
    a scale of exactly 1.0 on both sides would only duplicate the base
    empirical surface under a different source label.
    """
    cfg = config or ContextAdjustmentConfig()
    if not isinstance(record, dict):
        return None
    if record.get("injury_snapshot_status") != "snapshot_present":
        return None
    teams = record.get("teams")
    if not isinstance(teams, dict):
        return None
    home = _team_lost_points_share(teams.get("home"), cfg)
    away = _team_lost_points_share(teams.get("away"), cfg)
    if home is None or away is None:
        return None
    home_scale = home.pop("scale")
    away_scale = away.pop("scale")
    if home_scale >= 1.0 and away_scale >= 1.0:
        return None
    return ContextAdjustment(
        home_scale=home_scale,
        away_scale=away_scale,
        home_evidence=home,
        away_evidence=away,
    )


def context_calibration_components(
    config: "ContextAdjustmentConfig", *, base_config_asdict: Dict[str, Any]
) -> Dict[str, Any]:
    """Fingerprint components for the context surface's calibration version.

    Includes the full base model config (the surface is an empirical-history
    surface) plus the context knobs, but nothing else: paths, timestamps and
    host data must never enter version identity.
    """
    return {
        "model_version": id_spelling(MODEL_VERSION),
        "surface": CONTEXT_SURFACE,
        **base_config_asdict,
        **asdict(config),
    }
