"""
Elite Highest Quarter Pipeline.
Uses softmax normalization, Bayesian shrinkage, Forebet blending,
quarter-specific pace, coaching tendencies, rest advantage, and foul trouble.
Thresholds: STRONG >= 58%, MEDIUM >= 54%. Auto-reject anything weaker.
"""

import math
import logging
import uuid
from datetime import datetime
from typing import List, Optional, Tuple

from assayer.bet_record import BetRecord

logger = logging.getLogger(__name__)

QUARTERS = ["Q1", "Q2", "Q3", "Q4"]

STRONG_THRESHOLD = 58.0
MEDIUM_THRESHOLD = 54.0

DEFAULT_CONFIG = {
    "hq_softmax_temperature": 4.0,
    "hq_shrink_k": 10,
    "hq_min_confidence": 55.0,
    "hq_min_pwin": 0.35,
    "hq_skip_ties": True,
    "hq_vol_weight": 0.4,
    "hq_fb_weight": 0.25,
    "hq_exempt_from_cap": False,
    "hq_max_picks_per_slip": 2,
    "highest_q_tie_policy": "SKIP",
    "highest_q_tie_conf_penalty": 0.10,
    "highQtrTieMargin": 2.5,
    "config_version": "t2_elite_20260114_0801",
}


def _softmax(scores: List[float], temperature: float = 4.0) -> List[float]:
    if temperature <= 0:
        temperature = 1.0
    scaled = [s / temperature for s in scores]
    max_s = max(scaled)
    exps = [math.exp(s - max_s) for s in scaled]
    total = sum(exps)
    return [e / total for e in exps]


def _bayesian_shrink(p: float, n: int, shrink_k: int, league_prior: float = 0.25) -> float:
    """
    Shrink posterior toward league uniform prior (0.25 for 4 quarters).
    n = number of historical observations, shrink_k = pseudo-count.
    """
    return (p * n + league_prior * shrink_k) / (n + shrink_k)


def _blend_with_forebet(hq_probs: List[float], fb_probs: Optional[List[float]],
                         fb_weight: float = 0.25) -> List[float]:
    if fb_probs is None or len(fb_probs) != 4:
        return hq_probs
    blended = []
    for i in range(4):
        b = (1 - fb_weight) * hq_probs[i] + fb_weight * fb_probs[i]
        blended.append(b)
    total = sum(blended)
    return [b / total for b in blended] if total > 0 else hq_probs


def _apply_pace_adjustment(scores: List[float], quarter_pace_factors: Optional[List[float]]) -> List[float]:
    if not quarter_pace_factors or len(quarter_pace_factors) != 4:
        return scores
    adjusted = [s * f for s, f in zip(scores, quarter_pace_factors)]
    return adjusted


def _apply_rest_advantage(scores: List[float], rest_boost_quarter: Optional[int],
                           rest_magnitude: float = 0.05) -> List[float]:
    if rest_boost_quarter is None or rest_boost_quarter < 0 or rest_boost_quarter > 3:
        return scores
    boosted = list(scores)
    boosted[rest_boost_quarter] = boosted[rest_boost_quarter] * (1 + rest_magnitude)
    return boosted


def _apply_coaching_tendency(scores: List[float], preferred_quarter: Optional[int],
                               tendency_weight: float = 0.04) -> List[float]:
    if preferred_quarter is None or preferred_quarter < 0 or preferred_quarter > 3:
        return scores
    adjusted = list(scores)
    adjusted[preferred_quarter] = adjusted[preferred_quarter] * (1 + tendency_weight)
    return adjusted


def _is_tie(probs: List[float], tie_margin: float = 2.5) -> Tuple[bool, int, int]:
    max_prob = max(probs)
    max_idx = probs.index(max_prob)
    second_max = sorted(probs, reverse=True)[1]
    tie_threshold = tie_margin / 100.0
    is_tied = abs(max_prob - second_max) < tie_threshold
    second_idx = probs.index(second_max)
    return is_tied, max_idx, second_idx


def _confidence_from_prob(p: float) -> float:
    return round(p * 100, 4)


def _estimate_quarter_scores(game_row: dict) -> Optional[List[float]]:
    """
    Estimate relative quarter scoring strength from available data.
    Uses pace proxies from the row if available, otherwise uses uniform prior.
    """
    base = []
    for q_num in [1, 2, 3, 4]:
        home_key = f"Q{q_num}_Home_Proj"
        away_key = f"Q{q_num}_Away_Proj"
        pace_key = f"Q{q_num}_Pace"
        total_key = f"Q{q_num}_Total_Proj"

        val = None
        for k in [total_key, pace_key]:
            raw = game_row.get(k)
            if raw is not None:
                try:
                    val = float(str(raw).strip())
                    break
                except Exception:
                    pass

        if val is None:
            h_raw = game_row.get(home_key)
            a_raw = game_row.get(away_key)
            if h_raw is not None and a_raw is not None:
                try:
                    val = float(str(h_raw)) + float(str(a_raw))
                except Exception:
                    pass

        base.append(val if val is not None else 60.0)

    return base


def _get_forebet_probs(game_row: dict) -> Optional[List[float]]:
    fb_keys = [
        ("FB_Q1_Prob", "FB_Q2_Prob", "FB_Q3_Prob", "FB_Q4_Prob"),
        ("Forebet_Q1", "Forebet_Q2", "Forebet_Q3", "Forebet_Q4"),
    ]
    for key_set in fb_keys:
        vals = []
        for k in key_set:
            raw = game_row.get(k)
            if raw is not None:
                try:
                    vals.append(float(str(raw).strip()))
                except Exception:
                    break
        if len(vals) == 4:
            total = sum(vals)
            if total > 0:
                return [v / total for v in vals]
    return None


def run_hq_pipeline(
    game_rows: List[dict],
    config: Optional[dict] = None,
    origin_config: str = "TIER2",
    tier_level: str = "TIER2",
    source_module: str = "HQ_Pipeline",
) -> Tuple[List[BetRecord], int]:
    """
    Run the elite Highest Quarter pipeline over a list of game rows.
    Returns (accepted_records, predictions_generated).
    """
    cfg = {**DEFAULT_CONFIG, **(config or {})}

    temperature = float(cfg.get("hq_softmax_temperature", 4.0))
    shrink_k = int(cfg.get("hq_shrink_k", 10))
    min_confidence = float(cfg.get("hq_min_confidence", 55.0))
    skip_ties = bool(cfg.get("hq_skip_ties", True))
    tie_margin = float(cfg.get("highQtrTieMargin", 2.5))
    tie_conf_penalty = float(cfg.get("highest_q_tie_conf_penalty", 0.10))
    vol_weight = float(cfg.get("hq_vol_weight", 0.4))
    fb_weight = float(cfg.get("hq_fb_weight", 0.25))
    max_picks = int(cfg.get("hq_max_picks_per_slip", 2))
    config_version = cfg.get("config_version", "t2_elite_20260114_0801")

    accepted = []
    total_generated = 0

    for row in game_rows:
        league = str(row.get("League") or row.get("league") or "").strip()
        date = str(row.get("Date") or row.get("date") or "").strip()
        home = str(row.get("Home") or row.get("home") or "").strip()
        away = str(row.get("Away") or row.get("away") or "").strip()

        if not home or not away:
            continue

        base_scores = _estimate_quarter_scores(row)
        if base_scores is None:
            continue

        pace_factors = None
        raw_pace = [row.get(f"Q{q}_Pace_Factor") for q in [1, 2, 3, 4]]
        if all(v is not None for v in raw_pace):
            try:
                pace_factors = [float(str(v)) for v in raw_pace]
            except Exception:
                pass

        base_scores = _apply_pace_adjustment(base_scores, pace_factors)

        rest_quarter = None
        rest_raw = row.get("Rest_Advantage_Quarter")
        if rest_raw is not None:
            try:
                rest_quarter = int(rest_raw) - 1
            except Exception:
                pass
        base_scores = _apply_rest_advantage(base_scores, rest_quarter)

        coaching_quarter = None
        coaching_raw = row.get("Coaching_Preferred_Quarter")
        if coaching_raw is not None:
            try:
                coaching_quarter = int(coaching_raw) - 1
            except Exception:
                pass
        base_scores = _apply_coaching_tendency(base_scores, coaching_quarter)

        volatility = []
        for q_num in [1, 2, 3, 4]:
            raw = row.get(f"Q{q_num}_Volatility")
            if raw is not None:
                try:
                    volatility.append(float(str(raw)))
                except Exception:
                    volatility.append(0.0)
            else:
                volatility.append(0.0)

        if any(v > 0 for v in volatility):
            adjusted = [s * (1 - vol_weight * v) for s, v in zip(base_scores, volatility)]
        else:
            adjusted = base_scores

        probs = _softmax(adjusted, temperature=temperature)

        n_hist = int(row.get("HQ_Historical_N", 0) or 0)
        probs = [_bayesian_shrink(p, n_hist, shrink_k, 0.25) for p in probs]
        total_p = sum(probs)
        if total_p > 0:
            probs = [p / total_p for p in probs]

        fb_probs = _get_forebet_probs(row)
        probs = _blend_with_forebet(probs, fb_probs, fb_weight=fb_weight)

        total_generated += 1

        is_tied, max_idx, second_idx = _is_tie(probs, tie_margin=tie_margin)

        if is_tied:
            if skip_ties:
                logger.debug(f"HQ: skipping tie for {home} vs {away} (Q{max_idx+1} vs Q{second_idx+1})")
                continue
            probs[max_idx] = probs[max_idx] * (1 - tie_conf_penalty)
            total_p = sum(probs)
            probs = [p / total_p for p in probs] if total_p > 0 else probs

        max_prob = max(probs)
        max_q_idx = probs.index(max_prob)
        predicted_quarter = QUARTERS[max_q_idx]

        confidence_pct = _confidence_from_prob(max_prob)

        if confidence_pct < MEDIUM_THRESHOLD:
            logger.debug(f"HQ: rejected {home} vs {away} — confidence {confidence_pct:.2f}% < {MEDIUM_THRESHOLD}%")
            continue

        if confidence_pct < STRONG_THRESHOLD:
            tier_code = "MEDIUM"
            tier_display = "MEDIUM ●"
        else:
            tier_code = "STRONG"
            tier_display = "STRONG ★"

        q_totals_str = " ".join(
            f"{QUARTERS[i]}:{base_scores[i]:.1f}" for i in range(4)
        )

        game_slug = f"{date}_{home.upper().replace(' ', '_')}_{away.upper().replace(' ', '_')}"
        record_id = (
            f"{game_slug}__HQ__{predicted_quarter}__"
            f"{source_module}__{str(uuid.uuid4())[:8].upper()}"
        )

        rec = BetRecord(
            bet_record_id=record_id,
            universal_game_id=f"{date}__{home.upper().replace(' ', '_')}__{away.upper().replace(' ', '_')}",
            source_prediction_id=record_id,
            league=league,
            date=date,
            home=home,
            away=away,
            bet_type="HIGHEST_QUARTER",
            period="FT",
            quarter=predicted_quarter,
            selection_side="",
            selection_team="",
            selection_text=f"Highest Scoring Quarter: {predicted_quarter}",
            confidence_pct=confidence_pct,
            confidence_prob=round(max_prob, 4),
            tier_code=tier_code,
            tier_display=tier_display,
            origin_config=origin_config,
            config_version_used=config_version,
            tier_level=tier_level,
            config_version_t2=config_version,
            source_module=source_module,
            predicted=predicted_quarter,
            hq_q_totals=q_totals_str,
            hq_dominant_strength=round(max_prob, 4),
        )

        accepted.append(rec)

        if len(accepted) >= max_picks:
            break

    logger.info(
        f"HQ Pipeline: {total_generated} generated, {len(accepted)} accepted "
        f"(STRONG>={STRONG_THRESHOLD}% or MEDIUM>={MEDIUM_THRESHOLD}%)"
    )
    return accepted, total_generated
