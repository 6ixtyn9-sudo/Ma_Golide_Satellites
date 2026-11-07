"""Canonical, provider-independent data model for the acquisition layer.

These types are the contract between provider adapters, the prediction
selector, the derivation layer and the Sheets writer.  Adapters translate
provider payloads into these records; nothing downstream ever sees a raw
provider payload.

Design rules (see master prompt section 7):

- Predictions and bookmaker lines are NEVER conflated: ``Prediction`` and
  ``BookSnapshot`` are separate types and stay separate.
- All timestamps are timezone-aware UTC ISO 8601 strings internally.
- Money keeps raw precision in ``expected_*`` fields; rounding happens only
  at sheet rendering.
- A ``CoverageGap`` is emitted instead of guessing when a provider lacks a
  field.  No silent imputation.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple


# ---------------------------------------------------------------------------
# Enums / controlled vocabularies
# ---------------------------------------------------------------------------


class League(str, Enum):
    NBA = "NBA"
    EUROLEAGUE = "EUROLEAGUE"
    WNBA = "WNBA"
    NCAA_M = "NCAA_M"
    NCAA_W = "NCAA_W"

    @classmethod
    def display_value(cls, league: "League") -> str:
        """Canonical display value used in sheet ``League`` column."""
        return {
            League.NBA: "NBA",
            League.EUROLEAGUE: "EuroLeague",
            League.WNBA: "WNBA",
            League.NCAA_M: "NCAA M",
            League.NCAA_W: "NCAA W",
        }[league]


class GameStatus(str, Enum):
    SCHEDULED = "SCHEDULED"
    LIVE = "LIVE"
    FINAL = "FINAL"
    POSTPONED = "POSTPONED"
    CANCELLED = "CANCELLED"
    ABANDONED = "ABANDONED"
    UNKNOWN = "UNKNOWN"


class FinalStatus(str, Enum):
    """Status column values allowed on a result row (existing convention)."""

    FT = "FT"  # full time / regulation
    AOT = "AOT"  # after overtime
    ABD = "ABD"  # abandoned (kept from legacy parser; quarantined upstream)


class Market(str, Enum):
    MONEYLINE = "moneyline"
    SPREAD = "spread"
    TOTAL = "total"


class Period(str, Enum):
    FULL_GAME = "full_game"
    Q1 = "Q1"
    Q2 = "Q2"
    Q3 = "Q3"
    Q4 = "Q4"
    H1 = "H1"
    H2 = "H2"
    OT = "OT"


class CoverageState(str, Enum):
    """Coverage-report states (Milestone 0 requirement).

    Distinguishes *why* a field is empty so reports never hide the cause.
    """

    OK = "ok"
    MISSING = "missing"  # provider had the game but not the field
    STALE = "stale"  # value exists but failed the freshness policy
    UNSUPPORTED = "unsupported"  # provider/plan does not offer it
    UNAUTHORIZED = "unauthorized"  # credentials missing/invalid
    PROVIDER_ERROR = "provider_error"  # fetch/parse failure
    NOT_APPLICABLE = "not_applicable"  # e.g. OT cell on upcoming games


class PeriodKind(str, Enum):
    """What a provider's period data actually represents.

    Guards the "never relabel" rules: NCAA halves are HALVES, EuroLeague
    cumulative EndOfQuarter rows are CUMULATIVE.
    """

    QUARTERS = "quarters"
    HALVES = "halves"
    CUMULATIVE = "cumulative"
    UNKNOWN = "unknown"


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


def payload_hash(obj: Any) -> str:
    """Stable short hash of a JSON-serialisable payload (for provenance)."""
    canonical = json.dumps(obj, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:16]


def _require_utc(value: str, field_name: str) -> str:
    if not value.endswith("Z") and not re.search(r"[+-]\d{2}:\d{2}$", value):
        raise ValueError(
            f"{field_name} must be timezone-aware UTC ISO 8601 (got {value!r})"
        )
    return value


def round_half_up(value: float, ndigits: int = 0) -> float:
    """Decimal-compatible half-up rounding (matches sheet display rounding)."""
    from decimal import Decimal, ROUND_HALF_UP

    q = Decimal(1).scaleb(-ndigits)
    return float(Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP))


def fmt_prob(prob_home: Optional[float], prob_away: Optional[float]) -> Optional[str]:
    """Display ``H - A`` percent string, one decimal (e.g. ``23.1 - 76.9``).

    Inputs are 0..1 fractions (canonical storage) and are converted to
    percentages only here, at render time.
    """
    if prob_home is None or prob_away is None:
        return None
    return (
        f"{round_half_up(prob_home * 100.0, 1):.1f} - "
        f"{round_half_up(prob_away * 100.0, 1):.1f}"
    )


def fmt_score_pair(home: Optional[float], away: Optional[float]) -> Optional[str]:
    """Display ``home-away`` integer score pair (e.g. ``111-120``)."""
    if home is None or away is None:
        return None
    return f"{round_half_up(home, 0):.0f}-{round_half_up(away, 0):.0f}"


def fmt_total(total: Optional[float]) -> Optional[str]:
    """Display a model expected total at one decimal (e.g. ``230.6``)."""
    if total is None:
        return None
    return f"{round_half_up(total, 1):.1f}"


def pred_from_probabilities(
    prob_home: Optional[float], prob_away: Optional[float]
) -> Optional[int]:
    """``1`` when the home probability is higher, else ``2``.

    Exact ties resolve to ``2`` (the away side), matching the documented
    tie behaviour; ``None`` when probabilities are absent/invalid.
    """
    if prob_home is None or prob_away is None:
        return None
    if prob_away > prob_home:
        return 2
    if prob_home > prob_away:
        return 1
    return 2  # documented exact-tie behaviour


# ---------------------------------------------------------------------------
# Canonical records
# ---------------------------------------------------------------------------


@dataclass
class Team:
    """Canonical team identity plus provider ids/aliases."""

    canonical_name: str
    league: League
    provider_ids: Dict[str, str] = field(default_factory=dict)  # provider -> id
    aliases: List[str] = field(default_factory=list)

    def normalized(self) -> "Team":
        return self


@dataclass
class ActualResult:
    """Official actual periods and final, per side.

    ``overtime`` is the combined OT team-pair score; multiple OT periods are
    summed by the adapter and the aggregation is documented on the record.
    """

    period_kind: PeriodKind
    home_quarters: List[Optional[float]]  # [Q1..Q4] for quarters kind
    away_quarters: List[Optional[float]]
    home_overtime: Optional[float] = None
    away_overtime: Optional[float] = None
    home_final: Optional[float] = None
    away_final: Optional[float] = None
    overtime_periods: int = 0
    source: str = ""

    def reconciles(self, tolerance: float = 0.0) -> Tuple[bool, str]:
        """Check ``Q1+Q2+Q3+Q4+OT == FT`` independently for both teams."""
        if self.period_kind != PeriodKind.QUARTERS:
            return False, f"period kind {self.period_kind.value} is not quarters"
        if not self.home_quarters or not self.away_quarters:
            return False, "missing period values"
        if any(v is None for v in self.home_quarters + self.away_quarters):
            return False, "incomplete periods"
        if self.home_final is None or self.away_final is None:
            return False, "missing final"
        home_sum = sum(v for v in self.home_quarters if v is not None)
        away_sum = sum(v for v in self.away_quarters if v is not None)
        if self.home_overtime is not None:
            home_sum += self.home_overtime
        if self.away_overtime is not None:
            away_sum += self.away_overtime
        dh = abs(home_sum - self.home_final)
        da = abs(away_sum - self.away_final)
        if dh > tolerance or da > tolerance:
            return False, f"sum {home_sum:g}-{away_sum:g} != final {self.home_final:g}-{self.away_final:g}"
        return True, ""


@dataclass
class Game:
    """Canonical game identity (provider event ids live in ``provider_ids``)."""

    game_id: str
    league: League
    season: str
    home_team_id: str
    away_team_id: str
    home_name: str
    away_name: str
    scheduled_utc: str  # ISO 8601 Z
    status: GameStatus = GameStatus.SCHEDULED
    phase: str = "League"  # Game Type display value
    provider_ids: Dict[str, str] = field(default_factory=dict)  # provider -> event id
    actual: Optional[ActualResult] = None
    rescheduled_from: Optional[str] = None

    def __post_init__(self) -> None:
        _require_utc(self.scheduled_utc, "scheduled_utc")

    @property
    def utc_datetime(self):
        from datetime import datetime, timezone

        return datetime.fromisoformat(self.scheduled_utc.replace("Z", "+00:00")).astimezone(timezone.utc)

    def final_status(self) -> Optional[FinalStatus]:
        """Status cell value for a finished game (``FT``/``AOT`` convention)."""
        if self.status is not GameStatus.FINAL or self.actual is None:
            return None
        if self.actual.overtime_periods > 0:
            return FinalStatus.AOT
        if self.actual.home_final is None or self.actual.away_final is None:
            return None
        return FinalStatus.FT


@dataclass
class Prediction:
    """One provider's pre-tip prediction with full provenance.

    ``expected_total`` is the model expected total (the new ``Avg``
    semantics); it is NEVER a bookmaker line.
    """

    game_id: str
    provider: str
    observed_utc: str
    source_modified_utc: Optional[str] = None
    model: Optional[str] = None
    model_version: Optional[str] = None
    produced_by: Optional[str] = None
    prob_home: Optional[float] = None  # 0..1
    prob_away: Optional[float] = None  # 0..1
    expected_home_score: Optional[float] = None
    expected_away_score: Optional[float] = None
    expected_total: Optional[float] = None
    expected_margin: Optional[float] = None  # home minus away
    source_revision: Optional[str] = None  # artifact hash / commit / lastModified
    pre_tip: bool = True

    def __post_init__(self) -> None:
        _require_utc(self.observed_utc, "observed_utc")
        for name in ("prob_home", "prob_away"):
            v = getattr(self, name)
            if v is not None and not (0.0 <= v <= 1.0):
                raise ValueError(f"{name} must be a 0..1 probability (got {v})")

    def validate_orientation(self, tolerance: float = 0.25) -> None:
        """Validate probability sum/range and score/margin orientation.

        Raises ``ValueError`` on contradiction so callers quarantine instead
        of writing a wrong-orientation forecast.
        """
        if self.prob_home is not None and self.prob_away is not None:
            if abs((self.prob_home + self.prob_away) - 1.0) > 0.01:
                raise ValueError(
                    f"probabilities sum to {self.prob_home + self.prob_away}, not 1"
                )
        if (
            self.expected_home_score is not None
            and self.expected_away_score is not None
        ):
            margin = self.expected_home_score - self.expected_away_score
            if self.expected_margin is not None and abs(margin - self.expected_margin) > tolerance:
                raise ValueError(
                    f"expected score margin {margin:.3f} contradicts expected_margin "
                    f"{self.expected_margin:.3f} beyond tolerance"
                )
            if self.expected_total is not None:
                total = self.expected_home_score + self.expected_away_score
                if abs(total - self.expected_total) > tolerance:
                    raise ValueError(
                        f"expected scores sum to {total:.3f}, contradicts expected_total "
                        f"{self.expected_total:.3f} beyond tolerance"
                    )

    def display(self) -> Dict[str, Optional[str]]:
        """Render the ``G:J`` bundle exactly as the base sheet expects it."""
        return {
            "Prob %": fmt_prob(self.prob_home, self.prob_away),
            "Pred": pred_from_probabilities(self.prob_home, self.prob_away),
            "Pred Score": fmt_score_pair(self.expected_home_score, self.expected_away_score),
            "Avg": fmt_total(self.expected_total),
        }


@dataclass
class BookSnapshot:
    """One book's price/line for one market/period, fully normalised.

    Invariants enforced by construction and by ``books.py``:
    - ``period`` is exact (FULL_GAME / Q1..Q4); no derived periods exist.
    - ``line`` for totals markets is the LINE VALUE (e.g. 232.0), never a price.
    - prices are decimal; American odds are converted at ingestion.
    """

    game_id: str
    provider: str
    book: str
    market: Market
    period: Period
    observed_utc: str
    source_modified_utc: Optional[str] = None
    line: Optional[float] = None  # total line / spread value
    over_price_decimal: Optional[float] = None
    under_price_decimal: Optional[float] = None
    home_price_decimal: Optional[float] = None
    away_price_decimal: Optional[float] = None
    pre_match: bool = True
    main_line: bool = True
    payload_hash: Optional[str] = None
    # Raw provider identity is retained until the pipeline resolves the
    # snapshot to a canonical Game.  Names/start are optional for manual
    # canonical-ID imports, but provider adapters populate them whenever
    # available so cross-provider joins never rely on ID coincidence.
    provider_event_id: Optional[str] = None
    home_name: Optional[str] = None
    away_name: Optional[str] = None
    scheduled_utc: Optional[str] = None

    def __post_init__(self) -> None:
        _require_utc(self.observed_utc, "observed_utc")
        if self.scheduled_utc:
            _require_utc(self.scheduled_utc, "scheduled_utc")

    def moneyline_for_side(self, side: int) -> Optional[float]:
        """Decimal price for side 1 (home) or 2 (away)."""
        if side == 1:
            return self.home_price_decimal
        if side == 2:
            return self.away_price_decimal
        return None


@dataclass
class StandingsRow:
    """One standings row as the legacy ``loadStandings`` expects it.

    ``points_for``/``points_against`` are **cumulative** season totals (the
    legacy loader computes ``netRtg = (PF-PA)/GP``); per-game averages must
    never be written into these fields.
    """

    team_name: str
    position: Optional[int] = None
    games_played: Optional[int] = None
    wins: Optional[int] = None
    losses: Optional[int] = None
    points_for: Optional[float] = None
    points_against: Optional[float] = None
    win_pct: Optional[float] = None  # 0..1
    streak: Optional[int] = None  # signed: +W3 / -L2
    last10_pct: Optional[float] = None
    home_record: Optional[str] = None
    away_record: Optional[str] = None
    section: Optional[str] = None  # conference/division label, e.g. "East"
    source: str = ""


@dataclass
class CoverageGap:
    game_id: Optional[str]
    required_field: str
    reason: CoverageState
    provider: Optional[str] = None
    detail: str = ""
    observed_utc: str = ""

    def __post_init__(self) -> None:
        if not self.observed_utc:
            from datetime import datetime, timezone

            self.observed_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class ProviderHealth:
    provider: str
    league: League
    last_attempt_utc: Optional[str] = None
    last_success_utc: Optional[str] = None
    coverage: Dict[str, Any] = field(default_factory=dict)  # complete/total counters
    quota: Dict[str, Any] = field(default_factory=dict)  # remaining/used headers
    state: CoverageState = CoverageState.OK
    detail: str = ""


@dataclass
class CoverageReport:
    """Machine-readable acquisition summary (workflow Step Summary payload)."""

    league: str
    generated_utc: str
    mode: str  # live | dry_run | fixture
    fixtures_found: int = 0
    predictions_selected: Dict[str, int] = field(default_factory=dict)  # provider -> n
    full_game_line_coverage: Dict[str, int] = field(default_factory=dict)
    quarter_line_coverage: Dict[str, Dict[str, int]] = field(default_factory=dict)
    finals_written: int = 0
    frozen_predictions_used: int = 0
    gaps: List[Dict[str, str]] = field(default_factory=list)
    provider_health: List[Dict[str, Any]] = field(default_factory=list)
    quota: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "league": self.league,
            "generated_utc": self.generated_utc,
            "mode": self.mode,
            "fixtures_found": self.fixtures_found,
            "predictions_selected": self.predictions_selected,
            "full_game_line_coverage": self.full_game_line_coverage,
            "quarter_line_coverage": self.quarter_line_coverage,
            "finals_written": self.finals_written,
            "frozen_predictions_used": self.frozen_predictions_used,
            "gaps": self.gaps,
            "provider_health": self.provider_health,
            "quota": self.quota,
        }
