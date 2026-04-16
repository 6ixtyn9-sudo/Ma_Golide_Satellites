"""
MA GOLIDE COMPLETE ACCURACY REPORT — unified accuracy engine.
Single source of truth: reads exclusively from Bet_Slips + ResultsClean.
Grades all 7 bet types. Supports origin_config filter.
"""

import logging
import re
from collections import defaultdict
from datetime import datetime
from typing import List, Optional

from assayer.bet_record import BetRecord

logger = logging.getLogger(__name__)

ALL_BET_TYPES = [
    "BANKER", "ROBBER", "SNIPER_OU", "SNIPER_MARGIN",
    "FIRST_HALF_1X2", "FT_OU", "HIGHEST_QUARTER",
]

VALID_ORIGIN_CONFIGS = ("TIER1", "TIER2", "BLENDED", "LEGACY")


def _norm(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip().lower())


def _norm_team(v) -> str:
    v = _norm(v)
    v = re.sub(r"[^a-z0-9 ]", "", v)
    return v.strip()


def _norm_date(v) -> str:
    v = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return v


def _parse_float(v) -> Optional[float]:
    try:
        return float(str(v).replace("%", "").strip())
    except Exception:
        return None


def _match_key(date: str, home: str, away: str) -> str:
    return f"{_norm_date(date)}|{_norm_team(home)}|{_norm_team(away)}"


def _build_results_map(results_rows: List[dict]) -> dict:
    result_map = {}
    for row in results_rows:
        date = row.get("Date") or row.get("date") or ""
        home = row.get("Home") or row.get("home") or row.get("Home_Team") or ""
        away = row.get("Away") or row.get("away") or row.get("Away_Team") or ""
        if not date or not home or not away:
            continue
        key = _match_key(date, home, away)
        result_map[key] = row
    return result_map


def _get_result_value(row: dict, *keys):
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return None


def _parse_q_score(row: dict, quarter: str, side: str) -> Optional[int]:
    q = quarter.upper().replace("Q", "")
    prefix_map = {"1": "Q1", "2": "Q2", "3": "Q3", "4": "Q4"}
    q_label = prefix_map.get(q, f"Q{q}")
    s = "H" if _norm(side) in ("home", "h", "1") else "A"
    candidates = [
        f"{q_label}_{s}", f"{q_label}{s}",
        f"Q{q}_Home" if s == "H" else f"Q{q}_Away",
        f"Q{q}Home" if s == "H" else f"Q{q}Away",
    ]
    for c in candidates:
        v = row.get(c)
        if v is not None:
            try:
                return int(float(str(v).strip()))
            except Exception:
                pass
    return None


def _parse_q_total(row: dict, quarter: str) -> Optional[float]:
    q = quarter.upper().replace("Q", "")
    h = _parse_q_score(row, quarter, "H")
    a = _parse_q_score(row, quarter, "A")
    if h is not None and a is not None:
        return float(h + a)
    total_candidates = [f"Q{q}_Total", f"Q{q}Total", f"Q{q}Pts"]
    for c in total_candidates:
        v = row.get(c)
        if v is not None:
            try:
                return float(str(v).strip())
            except Exception:
                pass
    return None


def _parse_ft_score(row: dict) -> tuple:
    home_keys = ["FT_Home", "FT_H", "Home_Score", "Score_Home", "Home_FT"]
    away_keys = ["FT_Away", "FT_A", "Away_Score", "Score_Away", "Away_FT"]
    raw_score = _get_result_value(row, "Score", "Final_Score", "FT_Score", "Result_Score")
    if raw_score:
        parts = re.split(r"[-–—:]", raw_score)
        if len(parts) == 2:
            try:
                return int(float(parts[0].strip())), int(float(parts[1].strip()))
            except Exception:
                pass
    h = None
    a = None
    for k in home_keys:
        v = row.get(k)
        if v is not None:
            try:
                h = int(float(str(v)))
                break
            except Exception:
                pass
    for k in away_keys:
        v = row.get(k)
        if v is not None:
            try:
                a = int(float(str(v)))
                break
            except Exception:
                pass
    return h, a


def _parse_1h_score(row: dict) -> tuple:
    raw = _get_result_value(row, "1H_Score", "Half_Score", "HT_Score", "1H")
    if raw:
        parts = re.split(r"[-–—:]", raw)
        if len(parts) == 2:
            try:
                return int(float(parts[0].strip())), int(float(parts[1].strip()))
            except Exception:
                pass
    h = None
    a = None
    for k in ["1H_Home", "HT_Home", "Half_Home"]:
        v = row.get(k)
        if v is not None:
            try:
                h = int(float(str(v)))
                break
            except Exception:
                pass
    for k in ["1H_Away", "HT_Away", "Half_Away"]:
        v = row.get(k)
        if v is not None:
            try:
                a = int(float(str(v)))
                break
            except Exception:
                pass
    return h, a


def _grade_banker_robber(rec: BetRecord, result_row: dict) -> dict:
    """Grade BANKER / ROBBER bets — predicted winner vs actual winner."""
    h_score, a_score = _parse_ft_score(result_row)
    actual_score = f"{h_score} - {a_score}" if h_score is not None and a_score is not None else ""

    if h_score is None or a_score is None:
        return {"outcome": "NO_RESULT", "actual_score": actual_score, "actual_winner": ""}

    if h_score > a_score:
        actual_winner = rec.home
    elif a_score > h_score:
        actual_winner = rec.away
    else:
        actual_winner = "DRAW"

    predicted_winner = rec.selection_team or rec.predicted
    pred_norm = _norm_team(predicted_winner)
    home_norm = _norm_team(rec.home)
    away_norm = _norm_team(rec.away)

    if _norm_team(actual_winner) in (home_norm, away_norm):
        hit = pred_norm == _norm_team(actual_winner)
    else:
        hit = False

    return {
        "outcome": "HIT" if hit else "MISS",
        "actual_score": actual_score,
        "actual_winner": actual_winner,
    }


def _grade_sniper_ou(rec: BetRecord, result_row: dict) -> dict:
    """Grade SNIPER_OU — quarter over/under vs actual quarter total."""
    quarter = rec.quarter or rec.period or "Q1"
    line = rec.selection_line
    direction = _norm(rec.selection_side)

    if line is None:
        direction_from_text = _norm(rec.selection_text)
        if "over" in direction_from_text:
            direction = "over"
        elif "under" in direction_from_text:
            direction = "under"

    actual_total = _parse_q_total(result_row, quarter)
    if actual_total is None:
        return {"outcome": "NO_RESULT", "actual_total": None}

    if line is None:
        return {"outcome": "NO_RESULT", "actual_total": actual_total}

    if abs(actual_total - line) < 0.01:
        return {"outcome": "PUSH", "actual_total": actual_total}

    if direction in ("over", "o"):
        hit = actual_total > line
    elif direction in ("under", "u"):
        hit = actual_total < line
    else:
        return {"outcome": "NO_RESULT", "actual_total": actual_total}

    return {"outcome": "HIT" if hit else "MISS", "actual_total": actual_total}


def _grade_sniper_margin(rec: BetRecord, result_row: dict) -> dict:
    """Grade SNIPER_MARGIN — quarter spread bet."""
    quarter = rec.quarter or rec.period or "Q1"
    line = rec.selection_line or 0.0
    side = _norm(rec.selection_side)

    h_score = _parse_q_score(result_row, quarter, "H")
    a_score = _parse_q_score(result_row, quarter, "A")

    if h_score is None or a_score is None:
        return {"outcome": "NO_RESULT", "actual_q_home": None, "actual_q_away": None}

    if side in ("home", "h", "1"):
        covered = (h_score + line) > a_score
    elif side in ("away", "a", "2"):
        covered = (a_score + line) > h_score
    else:
        return {"outcome": "NO_RESULT", "actual_q_home": h_score, "actual_q_away": a_score}

    return {
        "outcome": "HIT" if covered else "MISS",
        "actual_q_home": h_score,
        "actual_q_away": a_score,
    }


def _grade_first_half_1x2(rec: BetRecord, result_row: dict) -> dict:
    """Grade FIRST_HALF_1X2 — 1H winner."""
    h_score, a_score = _parse_1h_score(result_row)
    if h_score is None or a_score is None:
        return {"outcome": "NO_RESULT", "actual_1h_score": ""}

    actual_1h_score = f"{h_score} - {a_score}"
    if h_score > a_score:
        actual_result = "1"
    elif a_score > h_score:
        actual_result = "2"
    else:
        actual_result = "X"

    predicted = _norm(rec.selection_side or rec.predicted or rec.selection_text)
    if "home" in predicted or predicted == "1":
        predicted_code = "1"
    elif "away" in predicted or predicted == "2":
        predicted_code = "2"
    elif predicted in ("x", "draw"):
        predicted_code = "X"
    else:
        return {"outcome": "NO_RESULT", "actual_1h_score": actual_1h_score, "actual_result": actual_result}

    return {
        "outcome": "HIT" if predicted_code == actual_result else "MISS",
        "actual_1h_score": actual_1h_score,
        "actual_result": actual_result,
    }


def _grade_ft_ou(rec: BetRecord, result_row: dict) -> dict:
    """Grade FT_OU — full-time over/under."""
    h_score, a_score = _parse_ft_score(result_row)
    if h_score is None or a_score is None:
        return {"outcome": "NO_RESULT", "actual_ft_total": None}

    actual_total = float(h_score + a_score)
    line = rec.selection_line
    direction = _norm(rec.selection_side or rec.selection_text)

    if line is None:
        return {"outcome": "NO_RESULT", "actual_ft_total": actual_total}

    if "over" in direction or direction == "o":
        hit = actual_total > line
    elif "under" in direction or direction == "u":
        hit = actual_total < line
    else:
        return {"outcome": "NO_RESULT", "actual_ft_total": actual_total}

    return {
        "outcome": "HIT" if hit else "MISS",
        "actual_ft_total": actual_total,
    }


def _grade_highest_quarter(rec: BetRecord, result_row: dict) -> dict:
    """Grade HIGHEST_QUARTER — predicted highest scoring quarter vs actual."""
    q_totals = {}
    for q_num in [1, 2, 3, 4]:
        total = _parse_q_total(result_row, f"Q{q_num}")
        if total is not None:
            q_totals[f"Q{q_num}"] = total

    if not q_totals:
        return {"outcome": "NO_RESULT", "q_totals_str": "", "actual_highest": ""}

    q_totals_str = " ".join(f"{q}:{int(v)}" for q, v in sorted(q_totals.items()))
    max_total = max(q_totals.values())
    actual_highest_quarters = [q for q, v in q_totals.items() if v == max_total]
    actual_highest = actual_highest_quarters[0] if actual_highest_quarters else ""

    predicted_q = _norm(rec.predicted or rec.selection_text or rec.quarter)
    for q in ["q4", "q3", "q2", "q1"]:
        if q in predicted_q:
            predicted_q = q.upper()
            break
    else:
        predicted_q = predicted_q.upper() if predicted_q else ""

    hit = predicted_q in actual_highest_quarters

    return {
        "outcome": "HIT" if hit else "MISS",
        "q_totals_str": q_totals_str,
        "actual_highest": f"{actual_highest} ({int(max_total)})",
    }


GRADER_MAP = {
    "BANKER":          _grade_banker_robber,
    "ROBBER":          _grade_banker_robber,
    "SNIPER_OU":       _grade_sniper_ou,
    "SNIPER_MARGIN":   _grade_sniper_margin,
    "FIRST_HALF_1X2": _grade_first_half_1x2,
    "FT_OU":           _grade_ft_ou,
    "HIGHEST_QUARTER": _grade_highest_quarter,
}


def run_accuracy_report(
    bet_slip_rows: List[dict],
    results_rows: List[dict],
    origin_config_filter: Optional[str] = None,
) -> dict:
    """
    Unified accuracy report — MA GOLIDE COMPLETE ACCURACY REPORT.
    Reads exclusively from Bet_Slips and ResultsClean.
    """
    result_map = _build_results_map(results_rows)

    records = [BetRecord.from_bet_slip_row(r) for r in bet_slip_rows]

    if origin_config_filter and origin_config_filter.upper() in VALID_ORIGIN_CONFIGS:
        records = [r for r in records if r.origin_config == origin_config_filter.upper()]

    by_type = defaultdict(lambda: {
        "found": 0, "matched": 0, "hits": 0, "misses": 0, "pushes": 0, "details": []
    })

    total_graded = 0
    total_hits = 0
    total_misses = 0

    for rec in records:
        bet_type_key = rec.bet_type.upper().replace(" ", "_").replace("-", "_")
        if bet_type_key not in GRADER_MAP:
            bet_type_key = "BANKER"

        by_type[bet_type_key]["found"] += 1

        match_key = _match_key(rec.date, rec.home, rec.away)
        result_row = result_map.get(match_key)

        if result_row is None:
            by_type[bet_type_key]["details"].append({
                "rec": rec.to_dict(),
                "outcome": "NO_MATCH",
                "detail": {},
            })
            continue

        by_type[bet_type_key]["matched"] += 1

        grader = GRADER_MAP[bet_type_key]
        grade_result = grader(rec, result_row)
        outcome = grade_result.get("outcome", "NO_RESULT")

        if outcome == "HIT":
            by_type[bet_type_key]["hits"] += 1
            total_hits += 1
        elif outcome == "MISS":
            by_type[bet_type_key]["misses"] += 1
            total_misses += 1
        elif outcome == "PUSH":
            by_type[bet_type_key]["pushes"] += 1

        if outcome in ("HIT", "MISS", "PUSH"):
            total_graded += 1

        rec.outcome = outcome
        rec.actual_result = grade_result.get("actual_result", "")
        rec.actual_score = grade_result.get("actual_score", "")
        rec.actual_winner = grade_result.get("actual_winner", "")
        if bet_type_key == "HIGHEST_QUARTER":
            rec.hq_q_totals = grade_result.get("q_totals_str", "")

        by_type[bet_type_key]["details"].append({
            "rec": rec.to_dict(),
            "outcome": outcome,
            "detail": grade_result,
        })

    total_scored = total_hits + total_misses
    overall_hit_rate = round((total_hits / total_scored * 100), 2) if total_scored > 0 else 0.0

    by_type_summary = {}
    for bt, data in sorted(by_type.items()):
        scored = data["hits"] + data["misses"]
        hit_rate = round((data["hits"] / scored * 100), 2) if scored > 0 else 0.0
        by_type_summary[bt] = {
            "bet_type": bt,
            "found": data["found"],
            "matched": data["matched"],
            "hits": data["hits"],
            "misses": data["misses"],
            "pushes": data["pushes"],
            "hit_rate_pct": hit_rate,
            "details": data["details"],
        }

    return {
        "report_name": "MA GOLIDE COMPLETE ACCURACY REPORT",
        "generated_at": datetime.utcnow().isoformat(),
        "origin_config_filter": origin_config_filter or "ALL",
        "total_bets_found": len(records),
        "total_bets_graded": total_graded,
        "total_hits": total_hits,
        "total_misses": total_misses,
        "overall_hit_rate_pct": overall_hit_rate,
        "by_bet_type": by_type_summary,
        "result_map_size": len(result_map),
    }
