"""Research pick rows for lanes that have a table and a price but no ticket route.

The problem this solves
-----------------------
Four aggregate-only lanes -- Greece, WCBA, WNBL, NKL -- have never shown a single
pick row. Two different reasons, both structural:

* Greece, WCBA and WNBL produce no forecast at all. Nothing scores their fixtures,
  so the card has nothing to label.
* NKL does have a results source, so a forecast exists, but its price is the site
  average (``aggregate_consensus``), which fails
  ``price_not_named_bookmaker_quote`` and is bucketed ``SKIPPED_INELIGIBLE``.

Both end in the same place: a lane that quotes fixtures and prices but never says
anything about who wins. This module says something -- and says exactly what it is.

Banker or robber
----------------
The operator's rule (2026-11-26): Ma Golide always made a bet, and it was either a
**banker** or a **robber**. Which one is not a matter of taste -- the retired
system encoded it, and the encoding was re-read from source (``de966196^``,
``docs/Forecaster.gs`` and ``docs/Contract_Enforcement.gs``) rather than inferred:

* Tiers, from ``Forecaster.gs`` ``getConfidenceTier`` / ``evaluateConfig``: STRONG
  at ``abs(score) >= 75`` (``tier_strong_min_score``), MEDIUM at ``>= 60``
  (``tier_medium_min_score``), WEAK at ``>= 50`` (``tier_weak_min_score``), and
  RISKY below ``threshold`` (25 in the validator defaults). ``confidence_tier``
  here is that ladder, applied to ``abs(score)``.
* ``Forecaster.gs``'s own tally: ``Bankers = STRONG-tier picks`` on either side;
  ``Robbers = AWAY (upset) picks that cleared MEDIUM threshold``.
* ``Contract_Enforcement.gs``'s sheet builder makes every named pick a banker once
  it clears ``bankerThreshold`` (default 60 -- the medium floor), and its separate
  Robbers engine bets the *underdog*, taken from the longer price when odds exist.

So a fixture is a **banker** when the model's side clears the banker gate -- STRONG
on either side, or MEDIUM on the home/favourite side -- and a **robber** otherwise:
an AWAY pick that only reached MEDIUM is the script's upset robber, and a fixture
that never reached MEDIUM (WEAK, LOW, or RISKY) has no model edge to bank, so the
robber takes the dog -- the longer observed price, else the weaker published row,
labelled for which of those it was rather than dressed up as a model claim.

``NO_EVIDENCE`` is not a third kind of bet: it is the statement that nothing was
read about either club *and* no price exists, so there is no favourite and no dog
to take. It names why instead of pretending a draw of two unknowns is a selection.

The rest of the inventory
-------------------------
The retired system wrote seven kinds of bet (``BET_TYPES`` in the old
``assayer/bet_record.py``): moneyline bankers and robbers, first-half 1X2, full-time
over/under, quarter margin and quarter totals ("snipers"), and highest quarter. A
standings table plus a moneyline price supports three of those honestly and one more
when a total line was observed, so each row carries a proposal per type -- including
the ones that did not fire, with a stable reason token such as
``ft_ou=UNAVAILABLE(no_totals_line)``. The quarter-level types are not offered at
all: inventing a quarter split from a season table would be a fabrication dressed as
a bet.

Config as data
--------------
Every gate above is a number in ``config/magolide_bets.json``, read through
``forecasting.bet_config`` (per-key bounds, monotonic tier invariants) and stamped
``CFG_<djb2>``. ``forecasting.config_tuner`` may amend those numbers from settled
results and records each change in ``localdata/config_ledger.jsonl``, so a card's
stamp always names the era of the system that produced it.

What a row is
-------------
A ``ResearchPickRow`` pairs one fixture with a ``MaGolideVerdict`` produced by
``forecasting/magolide_score.py`` from a *published* standings table
(``forecasting/standings.from_published``), plus the price observed at read time
when one exists. Every row carries its own provenance: weights source, standings
source, price source, and the eligibility the row does **not** have.

What a row is not
-----------------
It is not a probability. ``uncalibrated_logistic_pct`` is
``sigmoid(score / 30)`` over validator-default weights; it is printed as a research
number and never as a price, and the card says so on every line.

It is not a ticket leg. ``acca_eligible`` is ``False`` on every row, always -- not
because a gate happened to fail today, but because a site average is not a
named-book quote and no amount of research labelling changes that. Nothing here
touches a stake, a receipt, a bank balance or ROI.

What a research ticket is
-------------------------
A paper slip whose legs are research rows, priced at the observed average so the
group is auditable, and stamped ``RESEARCH PAPER — NO STAKE``. A single leg is a
single; two legs become a slip with a combined decimal price and **no** combined
probability (``Group success is not estimated`` -- the project's existing
shadow-group language). The combined price is arithmetic on two published numbers,
which is all a paper slip needs and all this module claims.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field, replace
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from ingestion.identity import canonical_team_name
from ingestion.models import BookSnapshot, Game, Market, Period
from ingestion.models import StandingsRow as PublishedStandingsRow

from .bet_config import BetConfig, default_config
from .magolide_score import (
    MaGolideConfig,
    MaGolideVerdict,
    StandingsRow,
    build_standings_features,
    score_features,
)
from .standings import PUBLISHED_SOURCE, from_published

#: The heading a card section built from these rows must carry.
RESEARCH_LABEL = "MAGOLIDE_SCORE RESEARCH-ONLY"

#: The bucket every row reports. Deliberately not one of the ticket buckets in
#: ``forecasting/operator.py``: a research row must never be mistaken for a
#: candidate that failed a gate and might still clear one.
RESEARCH_BUCKET = "RESEARCH_ONLY_NO_TICKET_ROUTE"

#: Printed on every line that carries a price.
PRICE_LABEL = "AGGREGATE AVERAGE — NOT A NAMED-BOOK QUOTE"

#: Printed on every research slip.
TICKET_LABEL = "RESEARCH PAPER — NO STAKE / NO RECEIPT / NO BANK IMPACT"

_SELECTIONS = {"HOME": "home", "AWAY": "away"}

#: The two kinds of bet Ma Golide makes: a **banker** backs the side the model
#: will bank, a **robber** takes the dog when the model has no edge to bank. There
#: is no third kind of bet -- a fixture with something readable gets one or the
#: other, which is what "Ma Golide always made a bet" means in practice.
KIND_BANKER = "BANKER"
KIND_ROBBER = "ROBBER"

#: The one case that is not a bet at all: neither club has a table row and no
#: market was observed, so there is nothing to be the favourite and nothing to pay
#: for the dog. Named rather than left blank so a reader sees a fact, not an
#: omission.
KIND_NO_EVIDENCE = "NO_EVIDENCE"

#: Tier ladder, exactly as the retired scorer thresholds it (``Forecaster.gs``
#: ``getConfidenceTier`` and the ``tier_*_min_score`` keys in its config master
#: list, both recovered from ``de966196^``). The RISKY floor is not a constant: it
#: is ``config.threshold``, carried on each verdict's ``meta``.
TIER_STRONG = "STRONG"
TIER_MEDIUM = "MEDIUM"
TIER_WEAK = "WEAK"
TIER_LOW = "LOW"
TIER_RISKY = "RISKY"

TIER_STRONG_MIN_SCORE = 75.0
TIER_MEDIUM_MIN_SCORE = 60.0
TIER_WEAK_MIN_SCORE = 50.0
TIER_LOW_MIN_SCORE = 25.0


def confidence_tier(score: float, *, risky_below: float = TIER_LOW_MIN_SCORE) -> str:
    """The ``Forecaster.gs`` tier for a signed score, by its magnitude.

    ``risky_below`` is the config's ``threshold``: below it the script's evaluator
    does ``continue`` -- no prediction, no tier. Passing the threshold that scored
    the game (``verdict.meta["threshold"]``) keeps this honest when a lane overrides
    the default.
    """

    magnitude = abs(float(score))
    if magnitude < risky_below:
        return TIER_RISKY
    if magnitude >= TIER_STRONG_MIN_SCORE:
        return TIER_STRONG
    if magnitude >= TIER_MEDIUM_MIN_SCORE:
        return TIER_MEDIUM
    if magnitude >= TIER_WEAK_MIN_SCORE:
        return TIER_WEAK
    return TIER_LOW


def _risky_below(verdict: MaGolideVerdict) -> float:
    """The ``threshold`` that actually scored this verdict, or the default."""

    raw = verdict.meta.get("threshold", TIER_LOW_MIN_SCORE)
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return TIER_LOW_MIN_SCORE
    return value if value >= 0 else TIER_LOW_MIN_SCORE



@dataclass
class ResearchPickRow:
    """One scored fixture, with everything needed to audit it and nothing more."""

    league: str
    game_id: str
    home: str
    away: str
    scheduled_utc: str
    verdict: MaGolideVerdict
    selection: Optional[str]  # "HOME" / "AWAY" / None only when nothing was read
    kind: Optional[str] = None  # BANKER / ROBBER / NO_EVIDENCE
    basis: str = ""  # why that kind, in words a reader can check
    decimal_price: Optional[float] = None
    price_book: Optional[str] = None
    price_provenance: Optional[str] = None
    observed_utc: Optional[str] = None
    #: True when a full-game moneyline for this fixture was observed at all, even
    #: if the verdict named no side and so no side's price could be taken.
    market_observed: bool = False
    standings_evidence: Dict[str, object] = field(default_factory=dict)
    #: Every bet type this module evaluated for the fixture, headline first. The
    #: ones that could not be priced are present too, carrying the reason, so the
    #: card shows the retired system's inventory rather than only its hits.
    bets: Tuple[BetProposal, ...] = ()
    #: ``CFG_<djb2>`` of the bet configuration that produced these proposals.
    config_stamp: str = ""

    @property
    def tier(self) -> str:
        """Which rung of the scorer's ladder this fixture landed on."""

        return confidence_tier(self.verdict.score, risky_below=_risky_below(self.verdict))

    @property
    def rank_suppressed(self) -> bool:
        """True when the two clubs sit in different stages/groups of one table."""

        return bool(self.standings_evidence.get("rank_comparison_suppressed"))

    @property
    def observed(self) -> bool:
        return self.market_observed

    @property
    def priced(self) -> bool:
        """A price *for the selected side*, which is what a slip needs."""

        return self.decimal_price is not None

    @property
    def has_selection(self) -> bool:
        return self.selection is not None

    def as_dict(self) -> Dict[str, object]:
        return {
            "league": self.league,
            "game_id": self.game_id,
            "home": self.home,
            "away": self.away,
            "scheduled_utc": self.scheduled_utc,
            "prediction": self.verdict.prediction,
            "selection": self.selection,
            "kind": self.kind,
            "kind_basis": self.basis,
            "bets": [proposal.as_dict() for proposal in self.bets],
            "config_stamp": self.config_stamp,
            "score": self.verdict.score,
            "confidence": self.verdict.confidence,
            "uncalibrated_logistic_pct": self.verdict.uncalibrated_logistic_pct,
            "decimal_price": self.decimal_price,
            "price_book": self.price_book,
            "price_provenance": self.price_provenance,
            "market_observed": self.market_observed,
            "observed_utc": self.observed_utc,
            "bucket": RESEARCH_BUCKET,
            "rank_comparison_suppressed": self.rank_suppressed,
            "acca_eligible": False,
            "bank_impact": "none",
            "weights_source": self.verdict.meta.get("weights_source"),
            "standings_source": PUBLISHED_SOURCE,
            "standings_evidence": dict(self.standings_evidence),
        }

    def as_line(self) -> str:
        """One audit line, in the vocabulary the existing cards already use."""

        v = self.verdict
        price = (
            f"@{self.decimal_price:.3f} {self.price_book or 'average'} "
            f"({PRICE_LABEL})"
            if self.priced
            else "no observed price"
        )
        selection = self.selection or "NONE"
        kind = self.kind or KIND_NO_EVIDENCE
        return (
            f"    [{RESEARCH_BUCKET}] {kind} · {self.home} vs {self.away} · "
            f"{self.scheduled_utc} · "
            f"MAGOLIDE_SCORE full_game -> {v.prediction} · tier={self.tier} · "
            f"selection={selection} · "
            f"basis={self.basis} · "
            f"bets=[{', '.join(proposal.as_token() for proposal in self.bets[1:])}] · "
            f"config_stamp={self.config_stamp} · "
            f"score={v.score:+.2f} · uncalibrated_logistic={v.uncalibrated_logistic_pct}% "
            f"· confidence={v.confidence}% · {price} · acca_eligible=False · "
            f"grade=NOT_SETTLEABLE_HERE · reasons=no_named_book_price,"
            f"weights={v.meta.get('weights_source')},standings={PUBLISHED_SOURCE}"
        )


@dataclass
class ResearchTicket:
    """A paper slip over research legs. Never a position in the bank."""

    legs: Tuple[ResearchPickRow, ...]
    combined_decimal: float

    @property
    def label(self) -> str:
        kind = "SINGLE" if len(self.legs) == 1 else f"{len(self.legs)}-LEG SLIP"
        return f"[{TICKET_LABEL}] {kind} @{self.combined_decimal:.3f}"

    def as_lines(self) -> List[str]:
        """The slip, one leg per fixture, each naming the bet type it carries.

        One leg per fixture is the rule the other lanes' auto-tickets follow, and it
        is why the headline proposal is the one that travels: a fixture that was also
        evaluated for a first half or a total contributes its moneyline bet, once.
        """

        lines = [f"  {self.label}"]
        lines.extend(
            f"    leg: {leg.kind or '-':<9} {leg.selection} {leg.home} vs {leg.away} "
            f"@{float(leg.decimal_price):.3f} {leg.price_book or 'average'}"
            for leg in self.legs
        )
        lines.append(
            "    combined probability: NOT ESTIMATED — multiplying uncalibrated "
            "scores would invent a number this module cannot support"
        )
        return lines


def _price_index(
    snapshots: Iterable[BookSnapshot],
) -> Tuple[Dict[str, BookSnapshot], Dict[Tuple[str, str], BookSnapshot]]:
    """Index full-game moneylines by game id and by canonical team pair.

    The pair index exists because a fixture's id in our store and the
    aggregator's ``betexplorer:<match_id>`` id are not always the same string,
    while the two club names are what a reader would check anyway.
    """

    by_id: Dict[str, BookSnapshot] = {}
    by_pair: Dict[Tuple[str, str], BookSnapshot] = {}
    for snapshot in snapshots:
        if snapshot.market is not Market.MONEYLINE or snapshot.period is not Period.FULL_GAME:
            continue
        if snapshot.home_price_decimal is None and snapshot.away_price_decimal is None:
            continue
        by_id.setdefault(snapshot.game_id, snapshot)
        pair = (
            canonical_team_name(snapshot.home_name or ""),
            canonical_team_name(snapshot.away_name or ""),
        )
        by_pair.setdefault(pair, snapshot)
    return by_id, by_pair


def _market_index(
    snapshots: Sequence[BookSnapshot], market: Market, period: Period
) -> Tuple[Dict[str, BookSnapshot], Dict[Tuple[str, str], BookSnapshot]]:
    """Index one market/period by game id and by club pair.

    Separate from the moneyline index so a totals line or a first-half price can be
    found without loosening the moneyline lookup -- a split-market row must never
    be mistaken for a full-game price.
    """

    by_id: Dict[str, BookSnapshot] = {}
    by_pair: Dict[Tuple[str, str], BookSnapshot] = {}
    for snapshot in snapshots:
        if snapshot.market is not market or snapshot.period is not period:
            continue
        if snapshot.line is None and snapshot.market is Market.TOTAL:
            continue
        by_id.setdefault(snapshot.game_id, snapshot)
        by_pair.setdefault(
            (
                canonical_team_name(snapshot.home_name or ""),
                canonical_team_name(snapshot.away_name or ""),
            ),
            snapshot,
        )
    return by_id, by_pair


def _longer_priced_side(snapshot: Optional[BookSnapshot]) -> Optional[str]:
    """The dog: the side with the longer price, or ``None`` when unreadable.

    Only unambiguous when *both* prices are present. A lone price says nothing
    about which side is the longer one, so it answers ``None`` and the caller falls
    back to the table rather than calling a favourite a dog.
    """

    if snapshot is None:
        return None
    home = snapshot.home_price_decimal
    away = snapshot.away_price_decimal
    if home is None or away is None or home == away:
        return None
    return "HOME" if home > away else "AWAY"


def _table_stronger_side(
    home: Optional[StandingsRow], away: Optional[StandingsRow]
) -> Optional[str]:
    """The stronger side by the published table alone, or ``None``.

    Win percentage first, then points difference. A club with a row beats a club
    without one -- that is a statement about evidence, not strength, and the basis
    string says so.
    """

    if home is None and away is None:
        return None
    if home is None or away is None:
        return "AWAY" if home is None else "HOME"
    for field_name in ("pct", "net_rtg"):
        left = getattr(home, field_name)
        right = getattr(away, field_name)
        if left is None or right is None:
            continue
        if left > right:
            return "HOME"
        if right > left:
            return "AWAY"
    return None


def _table_weaker_side(
    home: Optional[StandingsRow], away: Optional[StandingsRow]
) -> Optional[str]:
    """The dog by the published table alone, or ``None`` when the table is level.

    The inverse of :func:`_table_stronger_side`, which already answers ``None`` for
    a level table and for two absent rows -- so this cannot invent a dog from a
    tie. A club with a row is stronger than one without, which makes the club
    *without* a row the dog; that is a statement about evidence, and the basis
    string says so.
    """

    stronger = _table_stronger_side(home, away)
    if stronger is None:
        return None
    return "AWAY" if stronger == "HOME" else "HOME"


def _dog_side(
    verdict: MaGolideVerdict,
    home: Optional[StandingsRow],
    away: Optional[StandingsRow],
    snapshot: Optional[BookSnapshot],
) -> Tuple[Optional[str], str]:
    """``(side, reason)`` for the dog, in the order the evidence deserves.

    Price first: with both sides quoted, the longer price *is* the dog, whatever the
    table says. Then the table: the weaker published row. Only then the model, by
    its inverse -- a named side means the other side is the dog, which is the
    weakest of the three claims and is labelled as such.
    """

    priced = _longer_priced_side(snapshot)
    if priced is not None:
        return priced, "market_dog"
    weaker = _table_weaker_side(home, away)
    if weaker is not None:
        return weaker, "table_weak_side"
    if verdict.prediction in _SELECTIONS:
        return ("AWAY" if verdict.prediction == "HOME" else "HOME"), "opposite_the_model_pick"
    if home is None and away is None:
        return None, "no_readable_table_and_no_price"
    # Both clubs have rows and they are level on every field this reads, with no
    # market pair to break the tie. That is a measured draw, not a missing read.
    return None, "table_level_and_no_price"


#: The banker gate. ``Contract_Enforcement.gs`` builds its BANKERS section from
#: every named pick clearing ``config.bankerThreshold`` (default 60), which is the
#: medium tier floor; this module applies the same number to the score, because the
#: score is the quantity the research card publishes.
BANKER_MIN_SCORE = TIER_MEDIUM_MIN_SCORE


def decide_bet(
    verdict: MaGolideVerdict,
    home: Optional[StandingsRow],
    away: Optional[StandingsRow],
    snapshot: Optional[BookSnapshot],
) -> Tuple[Optional[str], Optional[str], str]:
    """``(selection, kind, basis)`` -- the bet Ma Golide makes on a fixture.

    The retired contract, in three lines:

    * **banker** -- the model's own side, when it clears the banker gate
      (``BANKER_MIN_SCORE``, the same 60 as ``bankerThreshold``): STRONG on either
      side, or MEDIUM on the home side. Backing a side you scored that highly is
      the banker bet.
    * **robber** -- everything else. An AWAY pick that only reached MEDIUM is
      ``Forecaster.gs``'s upset robber by name. WEAK, LOW and RISKY have no model
      edge to bank, so the robber takes the dog: longer price, else weaker row,
      else the inverse of whatever the model said.
    * **NO_EVIDENCE** -- no row for either club and no price. Not a bet, and it
      says so.

    The basis string always starts with the tier, so a reader can see which rung of
    the ladder put the fixture where it is and check it against the score on the
    same line.
    """

    tier = confidence_tier(verdict.score, risky_below=_risky_below(verdict))

    clears_gate = abs(float(verdict.score)) >= BANKER_MIN_SCORE
    if verdict.prediction in _SELECTIONS and clears_gate:
        if tier == TIER_STRONG or (tier == TIER_MEDIUM and verdict.prediction == "HOME"):
            return verdict.prediction, KIND_BANKER, f"{tier.lower()}_tier_model_pick"

    if tier == TIER_MEDIUM and verdict.prediction == "AWAY":
        # ''Robbers = AWAY (upset) picks that cleared MEDIUM threshold'' -- the
        # script's own words, and the one case it names a robber outright.
        return "AWAY", KIND_ROBBER, "medium_tier_away_upset"

    side, reason = _dog_side(verdict, home, away, snapshot)
    if side is not None:
        return side, KIND_ROBBER, f"{tier.lower()}_tier_{reason}"
    return None, KIND_NO_EVIDENCE, f"{tier.lower()}_tier_{reason}"



#: The bet types this module can honestly price from a standings table plus the
#: prices a lane publishes. The retired system had seven (``BANKER``, ``ROBBER``,
#: ``FIRST_HALF_1X2``, ``FT_OU``, ``SNIPER_MARGIN``, ``SNIPER_OU``,
#: ``HIGHEST_QUARTER``); the three quarter-level ones need per-quarter splits and
#: quarter lines that a standings table does not carry, so they are not offered
#: here at all rather than offered and left empty.
BET_FIRST_HALF = "FIRST_HALF_1X2"
BET_FT_OU = "FT_OU"


@dataclass(frozen=True)
class BetProposal:
    """One bet type's answer for one fixture: a side, or the reason there is none.

    ``reason`` is a stable token, not prose: a card reader can count
    ``no_totals_line`` across a slate and learn something, where "no bet" teaches
    nothing.
    """

    bet_type: str
    available: bool
    selection: Optional[str] = None
    confidence: float = 0.0
    price: Optional[float] = None
    line: Optional[float] = None
    basis: str = ""
    reason: str = ""
    #: The league discount the tuner applied, 1.0 when it has not. Shown in the
    #: token only when it bit, so a reader can see *why* a confidence fell short.
    league_weight: float = 1.0

    def as_token(self) -> str:
        """``first_half_1x2=HOME@72.3/1.950`` or ``ft_ou=UNAVAILABLE(no_totals_line)``."""

        name = self.bet_type.lower()
        if not self.available:
            token = f"{name}=UNAVAILABLE({self.reason})"
            if self.league_weight != 1.0:
                token += f" lw={self.league_weight:.2f}"
            return token
        token = f"{name}={self.selection}"
        if self.line is not None:
            token += f" line={self.line:g}"
        token += f" conf={self.confidence:.1f}"
        if self.league_weight != 1.0:
            token += f" lw={self.league_weight:.2f}"
        if self.price is not None:
            token += f" @{self.price:.3f}"
        return token

    def as_dict(self) -> Dict[str, object]:
        return {
            "bet_type": self.bet_type,
            "available": self.available,
            "selection": self.selection,
            "confidence": self.confidence,
            "price": self.price,
            "line": self.line,
            "basis": self.basis,
            "reason": self.reason,
            "league_weight": self.league_weight,
        }


def _h1_norm_cdf(z: float) -> float:
    """``Accumulator_Builder._h1_normCdf``: the logistic stand-in for a normal CDF."""

    return 1.0 / (1.0 + math.exp(-1.7 * z))


def _norm_cdf(z: float) -> float:
    """The real thing, for totals, where the script used an erf-based CDF."""

    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def _clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


#: The margin signal's share of the first-half combiner's weights
#: (``H1_CONFIG.weights``: margin .50, h2h .25, form .15, odds .10). A standings
#: table supports exactly one of the four, so the combiner's own data-quality
#: divisor (``totalWeight / 0.8``, clamped to [0.5, 1.0]) applies as 0.50/0.8.
H1_MARGIN_WEIGHT = 0.50
H1_QUALITY_REFERENCE = 0.80


def first_half_bet(
    verdict: MaGolideVerdict,
    snapshot: Optional[BookSnapshot],
    config: BetConfig,
    league: str = "",
) -> BetProposal:
    """``FIRST_HALF_1X2`` from the score, on the combiner's own arithmetic.

    The retired ``predictFirstHalf1x2`` combines four signals (margin, head to head,
    form, odds) and falls back to ``margin = combinedSignal * 8`` when no margin
    stats exist. A standings table gives us the score and nothing else, so:

    * the score is divided by ``firstHalfSignalDivisor`` to land on the combiner's
      ``[-1, 1]`` signal scale;
    * ``firstHalfHomeAdvantage`` is added, as the script adds it;
    * the winner is two-way (``firstHalfTwoWay``), the script's own default;
    * confidence is the script's formula -- ``50 + (pFavor*100 - 50) * dataQuality``
      with ``pFavor = _h1_norm_cdf(absSignal * 2)`` -- and ``dataQuality`` is the
      margin weight over the combiner's 0.8 reference, because that is the only
      signal we have.

    ``basis`` says ``h1_signal_from_magolide_score`` on every available proposal, so
    a reader is never told this is the full four-signal combiner.
    """

    if not config.get("enableFirstHalf"):
        return BetProposal(BET_FIRST_HALF, False, reason="enableFirstHalf_off")

    signal = _clamp(float(verdict.score) / config.get("firstHalfSignalDivisor"), -1.0, 1.0)
    margin = signal * config.get("firstHalfMarginPerSignal") + config.get(
        "firstHalfHomeAdvantage"
    )
    if not config.get("firstHalfTwoWay") and abs(margin) < config.get("firstHalfDrawThreshold"):
        return BetProposal(BET_FIRST_HALF, False, reason="draw_within_threshold")

    selection = "HOME" if margin >= 0 else "AWAY"
    p_favour = _h1_norm_cdf(config.get("firstHalfSignalScale") * abs(signal))
    quality = _clamp(H1_MARGIN_WEIGHT / H1_QUALITY_REFERENCE, 0.5, 1.0)
    confidence = 50.0 + (p_favour * 100.0 - 50.0) * quality
    confidence = _clamp(
        confidence, config.get("firstHalfMinConfidence"), config.get("firstHalfMaxConfidence")
    )
    # The tuner's per-league discount applies to the confidence a gate compares,
    # which is exactly what the sheet's ``league_weight_<LEAGUE> = 0.7`` advice was
    # for. The undiscounted signal is still what chose the side.
    weight = config.league_weight(league)
    confidence *= weight

    if confidence < config.get("firstHalfMinConf"):
        return BetProposal(
            BET_FIRST_HALF,
            False,
            confidence=round(confidence, 1),
            league_weight=weight,
            reason="confidence_below_first_half_gate",
        )

    price = _side_price(snapshot, selection)
    if price is not None:
        edge = confidence / 100.0 - (1.0 / price)
        if edge < config.get("firstHalfMinEdge"):
            return BetProposal(
                BET_FIRST_HALF,
                False,
                confidence=round(confidence, 1),
                price=price,
                league_weight=weight,
                reason="edge_below_first_half_gate",
            )

    return BetProposal(
        BET_FIRST_HALF,
        True,
        selection=selection,
        confidence=round(confidence, 1),
        price=price,
        basis="h1_signal_from_magolide_score",
        league_weight=weight,
    )


def ft_ou_bet(
    home: Optional[PublishedStandingsRow],
    away: Optional[PublishedStandingsRow],
    snapshot: Optional[BookSnapshot],
    config: BetConfig,
    league: str = "",
) -> BetProposal:
    """``FT_OU`` against an observed total line, or the reason there is none.

    ``predictFTOverUnder`` skips with ``No valid FT line found`` when the sheet has
    no total, and this does the same rather than inventing a line: a totals bet is a
    comparison to a number somebody posted.

    With a line, the model total is the sum of each side's expected points
    (``(own PF + opponent PA) / 2``), shrunk towards the posted line by
    ``ftouShrinkK`` exactly as the script shrinks, and turned into a probability with
    a normal CDF whose sigma is the configured floor -- a standings table carries no
    per-game dispersion, so the floor is the only sigma it can defend.
    """

    if not config.get("enableFTOU"):
        return BetProposal(BET_FT_OU, False, reason="enableFTOU_off")
    if snapshot is None or snapshot.line is None:
        return BetProposal(BET_FT_OU, False, reason="no_totals_line")
    if snapshot.over_price_decimal is None or snapshot.under_price_decimal is None:
        return BetProposal(BET_FT_OU, False, reason="no_over_under_prices")
    if home is None or away is None:
        return BetProposal(BET_FT_OU, False, reason="no_standings_row")

    played = min(home.games_played or 0, away.games_played or 0)
    if played < config.get("ftouMinSamples"):
        return BetProposal(BET_FT_OU, False, reason="sample_below_ftou_minimum")

    home_pf = (home.points_for or 0.0) / played
    home_pa = (home.points_against or 0.0) / played
    away_pf = (away.points_for or 0.0) / played
    away_pa = (away.points_against or 0.0) / played

    line = float(snapshot.line)
    raw_mu = (home_pf + away_pa) / 2.0 + (away_pf + home_pa) / 2.0
    shrink = config.get("ftouShrinkK")
    mu = (played * raw_mu + shrink * line) / (played + shrink) if (played + shrink) else raw_mu
    sigma = max(config.get("ftouSigmaFloor"), config.get("ftouSigmaFloor") * config.get("ftouSigmaScale"))
    if sigma <= 0:
        return BetProposal(BET_FT_OU, False, reason="no_usable_sigma")

    p_over = 1.0 - _norm_cdf((line - mu) / sigma)
    selection = "OVER" if p_over >= 0.5 else "UNDER"
    probability = p_over if selection == "OVER" else 1.0 - p_over
    price = (
        snapshot.over_price_decimal if selection == "OVER" else snapshot.under_price_decimal
    )
    implied = 1.0 / float(price)
    edge = probability - implied
    ev = probability * (float(price) - 1.0) - (1.0 - probability)
    weight = config.league_weight(league)
    confidence = probability * 100.0 * weight

    if confidence < config.get("ftOuMinConf"):
        return BetProposal(
            BET_FT_OU, False, confidence=round(confidence, 1), line=line,
            league_weight=weight, reason="confidence_below_ftou_gate",
        )
    if edge < config.get("ftouEdgeThreshold"):
        return BetProposal(
            BET_FT_OU, False, confidence=round(confidence, 1), price=float(price), line=line,
            league_weight=weight, reason="edge_below_ftou_gate",
        )
    if ev < config.get("ftOuMinEv"):
        return BetProposal(
            BET_FT_OU, False, confidence=round(confidence, 1), price=float(price), line=line,
            league_weight=weight, reason="ev_below_ftou_gate",
        )

    return BetProposal(
        BET_FT_OU,
        True,
        selection=selection,
        confidence=round(confidence, 1),
        price=float(price),
        line=line,
        basis="model_total_shrunk_to_observed_line",
        league_weight=weight,
    )

def _side_price(snapshot: Optional[BookSnapshot], selection: Optional[str]) -> Optional[float]:
    if snapshot is None or selection is None:
        return None
    side = _SELECTIONS.get(selection)
    if side == "home":
        return snapshot.home_price_decimal
    if side == "away":
        return snapshot.away_price_decimal
    return None


def research_rows(
    league: str,
    games: Sequence[Game],
    published: Sequence[PublishedStandingsRow],
    *,
    snapshots: Sequence[BookSnapshot] = (),
    derived: Optional[Mapping[str, StandingsRow]] = None,
    config: Optional[MaGolideConfig] = None,
    bet_config: Optional[BetConfig] = None,
) -> List[ResearchPickRow]:
    """Score a slate from a published table, with the price that was observed.

    ``games`` is read for four attributes only -- ``game_id``, ``home_name``,
    ``away_name``, ``scheduled_utc`` -- so a caller may pass canonical ``Game``
    objects or a fixture taken straight off a published page. That matters for the
    lanes this exists for: three of them have no warehouse at all, and requiring a
    ``Game`` would mean requiring a store they do not have.

    A fixture whose clubs are absent from the table is still scored -- the scorer's
    neutral defaults make it ``RISKY`` on purpose -- and the row records
    ``neutral_defaults_applied`` so a reader can tell a thin table from a confident
    read. That is the difference between this module and a filter.
    """

    bets_config = bet_config or default_config()
    # The scorer runs under the same tunable numbers the bet gates do, unless a
    # caller pins a scorer config explicitly. Without this the tuner's
    # ``confidence_scale`` (its main calibration lever) would move a stamp and
    # nothing else: the card would keep scoring at the shipped default.
    resolved_config = config or MaGolideConfig(
        confidence_scale=bets_config.get("confidence_scale"),
        threshold=bets_config.get("threshold"),
        conf_min=bets_config.get("confMin"),
        conf_max=bets_config.get("confMax"),
    )
    scorer_rows = from_published(published, derived=derived)
    index = {row.team: row for row in scorer_rows}
    published_by_team = {
        canonical_team_name(str(row.team_name)): row for row in published
    }
    sections = {
        canonical_team_name(str(row.team_name)): row.section
        for row in published
        if row.section
    }
    by_id, by_pair = _price_index(snapshots)
    h1_by_id, h1_by_pair = _market_index(snapshots, Market.MONEYLINE, Period.H1)
    totals_by_id, totals_by_pair = _market_index(snapshots, Market.TOTAL, Period.FULL_GAME)

    out: List[ResearchPickRow] = []
    for game in games:
        home_name = canonical_team_name(game.home_name)
        away_name = canonical_team_name(game.away_name)
        home = index.get(home_name)
        away = index.get(away_name)
        # A published rank is a rank *within its stage or group*. NKL's table lists
        # the winners and losers stages as one 18-row list, so rank 2 in one group
        # and rank 16 in the other are not two positions on one ladder. When the
        # clubs sit in different sections the rank difference is dropped for both
        # (the scorer's neutral 15 applies to neither side's advantage) and the
        # suppression is recorded, rather than quoting a number that compares
        # nothing.
        home_section = sections.get(home_name)
        away_section = sections.get(away_name)
        cross_group = bool(home_section and away_section and home_section != away_section)
        if cross_group:
            if home is not None:
                home = replace(home, rank=None)
            if away is not None:
                away = replace(away, rank=None)
        features = build_standings_features(home, away)
        verdict = score_features(features, resolved_config)

        snapshot = by_id.get(game.game_id) or by_pair.get((home_name, away_name))
        pair = (home_name, away_name)
        selection, kind, basis = decide_bet(verdict, home, away, snapshot)

        # The headline proposal is the moneyline bet; the other types follow it so a
        # reader sees the whole inventory, including the ones that could not be
        # priced and why.
        headline = BetProposal(
            bet_type=kind or KIND_NO_EVIDENCE,
            available=selection is not None,
            selection=selection,
            confidence=verdict.confidence,
            price=_side_price(snapshot, selection),
            basis=basis,
            reason="" if selection is not None else basis,
        )
        proposals = (
            headline,
            first_half_bet(
                verdict,
                h1_by_id.get(game.game_id) or h1_by_pair.get(pair),
                bets_config,
                league,
            ),
            ft_ou_bet(
                published_by_team.get(home_name),
                published_by_team.get(away_name),
                totals_by_id.get(game.game_id) or totals_by_pair.get(pair),
                bets_config,
                league,
            ),
        )
        out.append(
            ResearchPickRow(
                league=league,
                game_id=game.game_id,
                home=game.home_name,
                away=game.away_name,
                scheduled_utc=game.scheduled_utc,
                verdict=verdict,
                selection=selection,
                kind=kind,
                basis=basis,
                decimal_price=_side_price(snapshot, selection),
                price_book=(snapshot.book if snapshot is not None else None),
                price_provenance=(
                    snapshot.price_provenance.value
                    if snapshot is not None and snapshot.price_provenance is not None
                    else None
                ),
                observed_utc=(snapshot.observed_utc if snapshot is not None else None),
                market_observed=snapshot is not None,
                bets=proposals,
                config_stamp=bets_config.stamp_id,
                standings_evidence={
                    "home_standings_present": home is not None,
                    "away_standings_present": away is not None,
                    "neutral_defaults_applied": home is None or away is None,
                    "home_rank": None if home is None else home.rank,
                    "away_rank": None if away is None else away.rank,
                    "home_section": home_section,
                    "away_section": away_section,
                    "rank_comparison_suppressed": cross_group,
                    "rows_in_table": len(scorer_rows),
                },
            )
        )
    return out


def research_tickets(rows: Sequence[ResearchPickRow]) -> List[ResearchTicket]:
    """Singles for every priced selection, plus two-leg slips over priced pairs.

    A slip needs a side and a number. The side is now always there when the table
    or the market can name one (see :func:`decide_bet`), but the number is not: a
    banker whose fixture carries no observed price is a bet with no paper to write,
    so it stays a labelled row rather than a slip. Slips are formed adjacent-first
    and never repeat a fixture, matching the grouping the existing shadow sections
    use.
    """

    priced = [row for row in rows if row.has_selection and row.priced]
    tickets = [
        ResearchTicket(legs=(row,), combined_decimal=float(row.decimal_price))
        for row in priced
    ]
    pending: List[ResearchPickRow] = []
    used: set[str] = set()
    for row in priced:
        if len(pending) == 1 and row.game_id not in used and row.game_id != pending[0].game_id:
            legs = (pending.pop(), row)
            tickets.append(
                ResearchTicket(
                    legs=legs,
                    # Rounded to the six decimals two three-decimal prices can
                    # produce: 1.15 * 1.40 is 1.6099999999999999 in binary
                    # floating point, and a paper slip shows a price, not an
                    # artefact of IEEE-754.
                    combined_decimal=round(
                        float(legs[0].decimal_price) * float(legs[1].decimal_price), 6
                    ),
                )
            )
            used.update(leg.game_id for leg in legs)
        else:
            pending = [row]
    return tickets


def render_research_section(
    rows: Sequence[ResearchPickRow],
    *,
    date_label: str,
    league: str,
) -> List[str]:
    """A fenced block for a pick or ticket card.

    No line carries a ``PENDING`` marker: the cards' readers use that word to
    decide whether an older card needs rebuilding, and a research row is not a
    grade awaiting settlement -- it is a number with no settlement route at all.
    """

    if not rows:
        return []
    tickets = research_tickets(rows)
    priced = sum(1 for row in rows if row.priced)
    bankers = sum(1 for row in rows if row.kind == KIND_BANKER)
    robbers = sum(1 for row in rows if row.kind == KIND_ROBBER)
    neutral = sum(
        1 for row in rows if row.standings_evidence.get("neutral_defaults_applied")
    )
    first_half_bets = sum(
        1
        for row in rows
        for proposal in row.bets
        if proposal.bet_type == BET_FIRST_HALF and proposal.available
    )
    totals_bets = sum(
        1
        for row in rows
        for proposal in row.bets
        if proposal.bet_type == BET_FT_OU and proposal.available
    )
    stamps = sorted({row.config_stamp for row in rows if row.config_stamp})
    lines = [
        "",
        f"{RESEARCH_LABEL} — {league} {date_label} — {len(rows)} fixture(s), "
        f"{bankers} banker(s), {robbers} robber(s), {priced} priced, "
        f"{len(tickets)} paper slip(s), {first_half_bets} first-half, "
        f"{totals_bets} totals",
        "- Research only: a published standings table and an aggregate average price. "
        "No named-book quote, no stake, no receipt, no bank or ROI effect.",
        "- MAGOLIDE_SCORE is not a probability: uncalibrated_logistic = "
        "sigmoid(score/30) over validator defaults, shown as a research number.",
        f"- Standings provenance: {PUBLISHED_SOURCE}; price provenance: {PRICE_LABEL}.",
        "- Every fixture gets a bet, as Ma Golide always did. Tiers are the retired "
        "scorer's (STRONG>=75, MEDIUM>=60, WEAK>=50, RISKY below threshold). BANKER = "
        "STRONG on either side, or MEDIUM on the home side. ROBBER = an AWAY pick that "
        "only reached MEDIUM, or any WEAK/LOW/RISKY fixture, where the dog is the "
        "longer price, else the weaker table row. The basis on each line names the "
        "tier it came from.",
        "- Bet inventory per fixture: the moneyline bet is the headline; "
        "FIRST_HALF_1X2 is derived from the same score with the retired combiner's "
        "arithmetic (one of its four signals, labelled as such); FT_OU is offered only "
        "against an observed total line. Quarter SNIPER and HIGHEST_QUARTER need "
        "per-quarter splits a standings table cannot give, so they are not offered.",
        f"- Bet config: {'/'.join(stamps) or 'none'} — every gate, bound and tier "
        "invariant lives in config/magolide_bets.json, and the tuner may amend those "
        "numbers from settled results; the stamp on every line names the era it ran in.",
        "- A league the tuner has discounted carries its weight on the confidence-gated "
        "bet types (first half, totals), where `lw=` shows it; the moneyline tier is a "
        "score, not a confidence, so it is not discounted.",
    ]
    if neutral:
        lines.append(
            f"- {neutral} fixture(s) had at least one club absent from the table and "
            "scored on the scorer's neutral defaults; those rows read RISKY by design."
        )
    suppressed = sum(1 for row in rows if row.rank_suppressed)
    if suppressed:
        lines.append(
            f"- {suppressed} fixture(s) paired clubs from different stages/groups of "
            "the same table, so rank was not used as a difference: a rank inside one "
            "group is not a rank inside another."
        )
    lines.append("- acca_eligible=False on every row; this section never becomes a ticket leg.")
    lines.append("")
    lines.extend(row.as_line() for row in rows)
    for ticket in tickets:
        lines.extend(ticket.as_lines())
    return lines


__all__ = [
    "PRICE_LABEL",
    "RESEARCH_BUCKET",
    "RESEARCH_LABEL",
    "TICKET_LABEL",
    "ResearchPickRow",
    "ResearchTicket",
    "render_research_section",
    "research_rows",
    "research_tickets",
]
