"""
Unified BetRecord — the single contract all pipelines feed into.
Every bet type must produce a BetRecord before being written to
Bet_Slips or graded by the accuracy engine.
"""

import datetime
from dataclasses import dataclass, field, asdict
from typing import Optional

BET_TYPES = {
    "BANKER":           "Moneyline — high-confidence home/away winner",
    "ROBBER":           "Moneyline — upset/underdog pick",
    "SNIPER_OU":        "Quarter Over/Under totals",
    "SNIPER_MARGIN":    "Quarter Spread / Side margin",
    "FIRST_HALF_1X2":  "First Half winner (1=home, 2=away)",
    "FT_OU":            "Full-Time Over/Under totals",
    "HIGHEST_QUARTER":  "Highest-scoring quarter prediction",
}

ORIGIN_CONFIGS = ("TIER1", "TIER2", "BLENDED", "LEGACY")
TIER_LEVELS    = ("TIER1", "TIER2")


def _tier_from_confidence(conf_pct: float, bet_type: str = "") -> tuple:
    if bet_type == "HIGHEST_QUARTER":
        if conf_pct >= 58:
            return "STRONG", "STRONG ★"
        if conf_pct >= 54:
            return "MEDIUM", "MEDIUM ●"
        return "WEAK", "WEAK ○"
    if conf_pct >= 80:
        return "ELITE", "★ ({}%) ★".format(int(conf_pct))
    if conf_pct >= 65:
        return "STRONG", "STRONG ★"
    if conf_pct >= 54:
        return "MEDIUM", "MEDIUM ●"
    return "WEAK", "WEAK ○"


def _resolve_origin_config(origin_config: str, tier_level: str, config_version_used: str) -> str:
    if origin_config in ORIGIN_CONFIGS:
        return origin_config
    if tier_level == "TIER1":
        return "TIER1"
    if tier_level == "TIER2":
        return "TIER2"
    if config_version_used:
        return "LEGACY"
    return "LEGACY"


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

    origin_config: str = "LEGACY"
    config_version_used: str = ""
    tier_level: str = ""

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

    def derive_tier(self):
        self.tier_code, self.tier_display = _tier_from_confidence(
            self.confidence_pct, self.bet_type
        )

    def derive_origin(self):
        self.origin_config = _resolve_origin_config(
            self.origin_config, self.tier_level, self.config_version_used
        )

    def finalize(self):
        self.derive_tier()
        self.derive_origin()

    def to_dict(self) -> dict:
        return asdict(self)

    def to_bet_slip_row(self) -> list:
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
        def _s(*keys):
            for k in keys:
                v = row.get(k, "")
                if v != "":
                    return str(v).strip()
            return ""

        def _f(*keys):
            for k in keys:
                v = row.get(k, "")
                if v != "":
                    try:
                        return float(str(v).replace("%", "").strip())
                    except Exception:
                        pass
            return None

        conf_raw = _f("Confidence_Pct", "confidence_pct", "Confidence")
        if conf_raw is not None and conf_raw <= 1.0:
            conf_raw *= 100

        bet_type = _s("Market", "bet_type", "Type").upper()
        t1 = _s("Config_Version_T1", "config_version_t1")
        t2 = _s("Config_Version_T2", "config_version_t2")
        origin = _s("Origin_Config", "origin_config")
        ver_used = _s("Config_Version_Used", "config_version_used")
        tier_lvl = _s("Tier_Level", "tier_level")

        tier_code, tier_display = _tier_from_confidence(conf_raw or 0, bet_type)

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
            selection_line=_f("Selection_Line", "selection_line"),
            selection_team=_s("Selection_Team", "selection_team"),
            selection_text=_s("Selection_Text", "selection_text"),
            odds=_f("Odds", "odds"),
            confidence_pct=conf_raw or 0.0,
            confidence_prob=_f("Confidence_Prob", "confidence_prob"),
            ev=_f("EV", "ev"),
            tier_code=tier_code,
            tier_display=tier_display,
            origin_config=origin if origin in ORIGIN_CONFIGS else "LEGACY",
            config_version_used=ver_used,
            tier_level=tier_lvl if tier_lvl in TIER_LEVELS else "",
            config_version_t1=t1,
            config_version_t2=t2,
            config_version_acc=_s("Config_Version_Acc", "config_version_acc"),
            source_module=_s("Source_Module", "source_module"),
            predicted=_s("Predicted", "predicted"),
            hq_q_totals=_s("HQ_Q_Totals", "hq_q_totals"),
        )
        return rec
