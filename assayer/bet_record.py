"""
Unified BetRecord — Single source of truth for all pipelines.
Strict traceability: origin_config is always TIER1, TIER2, BLENDED, or LEGACY.
"""

import datetime
from dataclasses import dataclass, field, asdict
from typing import Optional

BET_TYPES = {
    "BANKER": "Moneyline — high-confidence winner",
    "ROBBER": "Moneyline — underdog pick",
    "SNIPER_OU": "Quarter Over/Under",
    "SNIPER_MARGIN": "Quarter Spread",
    "FIRST_HALF_1X2": "First Half Winner",
    "FT_OU": "Full Time Over/Under",
    "HIGHEST_QUARTER": "Highest Scoring Quarter",
}

VALID_ORIGINS = {"TIER1", "TIER2", "BLENDED", "LEGACY"}


def _tier_from_confidence(conf_pct: float, bet_type: str = "") -> tuple:
    if bet_type == "HIGHEST_QUARTER":
        if conf_pct >= 58.0:
            return "STRONG", "STRONG ★"
        if conf_pct >= 54.0:
            return "MEDIUM", "MEDIUM ●"
        return "WEAK", "WEAK ○"
    
    if conf_pct >= 80:   return "ELITE",  f"★ ({int(conf_pct)}%) ★"
    if conf_pct >= 65:   return "STRONG", "STRONG ★"
    if conf_pct >= 54:   return "MEDIUM", "MEDIUM ●"
    return "WEAK", "WEAK ○"


@dataclass
class BetRecord:
    bet_record_id: str = ""
    universal_game_id: str = ""
    source_prediction_id: str = ""

    league: str = ""
    date: str = ""
    time: str = ""
    home: str = ""
    away: str = ""

    bet_type: str = ""
    period: str = ""
    quarter: str = ""

    selection_side: str = ""
    selection_line: Optional[float] = None
    selection_team: str = ""
    selection_text: str = ""

    odds: Optional[float] = None
    confidence_pct: float = 0.0
    confidence_prob: Optional[float] = None
    ev: Optional[float] = None

    tier_code: str = ""
    tier_display: str = ""

    origin_config: str = "LEGACY"      # TIER1, TIER2, BLENDED, LEGACY
    config_version_used: str = ""
    tier_level: str = ""               # TIER1 or TIER2

    config_version_t1: str = ""
    config_version_t2: str = ""
    source_module: str = ""

    predicted: str = ""
    actual_result: str = ""
    actual_score: str = ""
    actual_winner: str = ""
    outcome: str = ""

    hq_q_totals: str = ""
    hq_dominant_strength: Optional[float] = None

    created_at: str = field(default_factory=lambda: datetime.datetime.utcnow().isoformat())

    def finalize(self):
        """Call this before writing to Bet_Slips."""
        self.tier_code, self.tier_display = _tier_from_confidence(
            self.confidence_pct, self.bet_type
        )
        if not self.origin_config or self.origin_config not in VALID_ORIGINS:
            self.origin_config = "TIER2" if self.tier_level == "TIER2" else "TIER1"

    def to_dict(self) -> dict:
        return asdict(self)

    def to_bet_slip_row(self) -> list:
        return [
            self.bet_record_id, self.universal_game_id, self.source_prediction_id,
            self.league, self.date, self.time, self.home, self.away,
            self.bet_type, self.period, self.quarter,
            self.selection_side, self.selection_line or "", self.selection_team, self.selection_text,
            self.odds or "", self.confidence_pct, self.confidence_prob or "", self.ev or "",
            self.tier_code, self.tier_display,
            self.origin_config, self.config_version_used, self.tier_level,
            self.config_version_t1, self.config_version_t2, self.source_module,
            self.predicted, self.hq_q_totals, self.created_at,
        ]

    @classmethod
    def bet_slip_headers(cls) -> list:
        return [
            "Bet_Record_ID", "Universal_Game_ID", "Source_Prediction_Record_ID",
            "League", "Date", "Time", "Home", "Away", "Market", "Period", "Quarter",
            "Selection_Side", "Selection_Line", "Selection_Team", "Selection_Text",
            "Odds", "Confidence_Pct", "Confidence_Prob", "EV",
            "Tier_Code", "Tier_Display", "Origin_Config", "Config_Version_Used",
            "Tier_Level", "Config_Version_T1", "Config_Version_T2", "Source_Module",
            "Predicted", "HQ_Q_Totals", "Created_At"
        ]
