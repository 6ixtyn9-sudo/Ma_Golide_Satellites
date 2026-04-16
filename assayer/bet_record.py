"""
Unified BetRecord — The Single Contract for All Pipelines
============================================================

This is the central data structure used by every betting pipeline 
(BANKER, ROBBER, SNIPER_OU, SNIPER_MARGIN, FIRST_HALF_1X2, FT_OU, HIGHEST_QUARTER).

Every prediction must be converted into a BetRecord before being written to 
the Bet_Slips sheet or graded by the accuracy engine.

Key Traceability Features:
- origin_config:    "TIER1", "TIER2", "BLENDED", or "LEGACY"
- tier_level:       "TIER1" or "TIER2"
- config_version_used: Exact version string of the config that produced it

Config 1 is strictly used for Tier 1 bets. Config 2 is strictly used for Tier 2 bets.
Fallbacks have been minimized. Legacy code paths are clearly marked.
"""

import datetime
import logging
from dataclasses import dataclass, field, asdict
from typing import Optional, Any

logger = logging.getLogger(__name__)

BET_TYPES = {
    "BANKER":           "Moneyline — high-confidence home/away winner",
    "ROBBER":           "Moneyline — upset/underdog pick",
    "SNIPER_OU":        "Quarter Over/Under totals",
    "SNIPER_MARGIN":    "Quarter Spread / Side margin",
    "FIRST_HALF_1X2":   "First Half winner (1=home, 2=away)",
    "FT_OU":            "Full-Time Over/Under totals",
    "HIGHEST_QUARTER":  "Highest-scoring quarter prediction",
}

VALID_ORIGIN_CONFIGS = {"TIER1", "TIER2", "BLENDED", "LEGACY"}
VALID_TIER_LEVELS = {"TIER1", "TIER2"}


def _tier_from_confidence(conf_pct: float, bet_type: str = "") -> tuple[str, str]:
    """Return tier code and display string based on confidence and bet type."""
    if bet_type == "HIGHEST_QUARTER":
        if conf_pct >= 58.0:
            return "STRONG", "STRONG ★"
        if conf_pct >= 54.0:
            return "MEDIUM", "MEDIUM ●"
        return "WEAK", "WEAK ○"

    # Standard tier logic for all other bet types
    if conf_pct >= 80.0:
        return "ELITE", f"★ ({int(conf_pct)}%) ★"
    if conf_pct >= 65.0:
        return "STRONG", "STRONG ★"
    if conf_pct >= 54.0:
        return "MEDIUM", "MEDIUM ●"
    return "WEAK", "WEAK ○"


def _resolve_origin_config(
    origin_config: str,
    tier_level: str,
    config_version_used: str,
    config_version_t1: str,
    config_version_t2: str
) -> str:
    """Strict logic to determine origin_config with minimal fallbacks."""
    if origin_config in VALID_ORIGIN_CONFIGS:
        return origin_config

    if config_version_t1 and not config_version_t2:
        return "TIER1"
    if config_version_t2 and not config_version_t1:
        return "TIER2"
    if config_version_t1 and config_version_t2:
        return "BLENDED"

    if tier_level in VALID_TIER_LEVELS:
        return tier_level

    if config_version_used:
        return "LEGACY"

    logger.warning("Could not determine origin_config. Defaulting to LEGACY.")
    return "LEGACY"


@dataclass
class BetRecord:
    """Core data structure for all betting predictions."""
    
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

    # === Traceability Fields ===
    origin_config: str = "LEGACY"          # TIER1, TIER2, BLENDED, LEGACY
    config_version_used: str = ""
    tier_level: str = ""                   # TIER1 or TIER2

    config_version_t1: str = ""
    config_version_t2: str = ""
    config_version_acc: str = ""
    source_module: str = ""

    predicted: str = ""
    actual_result: str = ""
    actual_score: str = ""
    actual_winner: str = ""
    outcome: str = ""

    reject_reason: str = ""

    hq_q_totals: str = ""
    hq_dominant_strength: Optional[float] = None

    created_at: str = field(default_factory=lambda: datetime.datetime.utcnow().isoformat())

    def finalize(self):
        """Must be called before writing to Bet_Slips. Derives tier and origin."""
        self.tier_code, self.tier_display = _tier_from_confidence(
            self.confidence_pct, self.bet_type
        )
        
        self.origin_config = _resolve_origin_config(
            self.origin_config,
            self.tier_level,
            self.config_version_used,
            self.config_version_t1,
            self.config_version_t2
        )

        # Final validation
        if self.origin_config not in VALID_ORIGIN_CONFIGS:
            logger.error(f"Invalid origin_config '{self.origin_config}' for bet {self.bet_record_id}. Forcing TIER2.")
            self.origin_config = "TIER2"

    def to_dict(self) -> dict:
        return asdict(self)

    def to_bet_slip_row(self) -> list:
        """Returns a list formatted for direct insertion into the Bet_Slips sheet."""
        return [
            self.bet_record_id,
            self.universal_game_id,
            self.source_prediction_id,
            self.league,
            self.date,
            self.time,
            self.home,
            self.away,
            self.bet_type,
            self.period,
            self.quarter,
            self.selection_side,
            self.selection_line if self.selection_line is not None else "",
            self.selection_team,
            self.selection_text,
            self.odds if self.odds is not None else "",
            self.confidence_pct,
            self.confidence_prob if self.confidence_prob is not None else "",
            self.ev if self.ev is not None else "",
            self.tier_code,
            self.tier_display,
            self.origin_config,
            self.config_version_used,
            self.tier_level,
            self.config_version_t1,
            self.config_version_t2,
            self.config_version_acc,
            self.source_module,
            self.predicted,
            self.hq_q_totals,
            self.created_at,
        ]

    @classmethod
    def bet_slip_headers(cls) -> list:
        """Returns the exact column headers used in the Bet_Slips sheet."""
        return [
            "Bet_Record_ID", "Universal_Game_ID", "Source_Prediction_Record_ID",
            "League", "Date", "Time", "Home", "Away",
            "Market", "Period", "Quarter",
            "Selection_Side", "Selection_Line", "Selection_Team", "Selection_Text",
            "Odds", "Confidence_Pct", "Confidence_Prob", "EV",
            "Tier_Code", "Tier_Display",
            "Origin_Config", "Config_Version_Used", "Tier_Level",
            "Config_Version_T1", "Config_Version_T2", "Config_Version_Acc",
            "Source_Module", "Predicted", "HQ_Q_Totals", "Created_At",
        ]

    @classmethod
    def from_bet_slip_row(cls, row: dict) -> "BetRecord":
        """Parse a row from the Bet_Slips sheet into a BetRecord object."""
        def _s(*keys: str) -> str:
            for key in keys:
                val = row.get(key, "")
                if val not in (None, "", " "):
                    return str(val).strip()
            return ""

        def _f(*keys: str) -> Optional[float]:
            for key in keys:
                val = row.get(key)
                if val not in (None, "", " "):
                    try:
                        return float(str(val).replace("%", "").strip())
                    except Exception:
                        continue
            return None

        conf_raw = _f("Confidence_Pct", "Confidence", "confidence_pct")
        if conf_raw is not None and conf_raw <= 1.0:
            conf_raw *= 100.0

        bet_type = _s("Market", "bet_type", "Type", "Bet_Type").upper()
        t1 = _s("Config_Version_T1", "config_version_t1")
        t2 = _s("Config_Version_T2", "config_version_t2")

        rec = cls(
            bet_record_id=_s("Bet_Record_ID", "bet_record_id"),
            universal_game_id=_s("Universal_Game_ID", "universal_game_id"),
            source_prediction_id=_s("Source_Prediction_Record_ID", "source_prediction_id"),
            league=_s("League", "league"),
            date=_s("Date", "date"),
            time=_s("Time", "time"),
            home=_s("Home", "home"),
            away=_s("Away", "away"),
            bet_type=bet_type,
            period=_s("Period", "period"),
            quarter=_s("Quarter", "quarter"),
            selection_side=_s("Selection_Side", "selection_side"),
            selection_line=_f("Selection_Line", "selection_line", "Line"),
            selection_team=_s("Selection_Team", "selection_team"),
            selection_text=_s("Selection_Text", "selection_text", "Pick"),
            odds=_f("Odds", "odds"),
            confidence_pct=conf_raw or 0.0,
            confidence_prob=_f("Confidence_Prob", "confidence_prob"),
            ev=_f("EV", "ev"),
            tier_code=_s("Tier_Code", "tier_code"),
            tier_display=_s("Tier_Display", "tier_display"),
            origin_config=_s("Origin_Config", "origin_config"),
            config_version_used=_s("Config_Version_Used", "config_version_used"),
            tier_level=_s("Tier_Level", "tier_level"),
            config_version_t1=t1,
            config_version_t2=t2,
            config_version_acc=_s("Config_Version_Acc", "config_version_acc"),
            source_module=_s("Source_Module", "source_module"),
            predicted=_s("Predicted", "predicted"),
            hq_q_totals=_s("HQ_Q_Totals", "hq_q_totals"),
        )

        rec.finalize()
        return rec
