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

from dataclasses import dataclass, field, replace
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from ingestion.identity import canonical_team_name
from ingestion.models import BookSnapshot, Game, Market, Period
from ingestion.models import StandingsRow as PublishedStandingsRow

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


@dataclass
class ResearchPickRow:
    """One scored fixture, with everything needed to audit it and nothing more."""

    league: str
    game_id: str
    home: str
    away: str
    scheduled_utc: str
    verdict: MaGolideVerdict
    selection: Optional[str]  # "HOME" / "AWAY" / None when the verdict is RISKY
    decimal_price: Optional[float] = None
    price_book: Optional[str] = None
    price_provenance: Optional[str] = None
    observed_utc: Optional[str] = None
    #: True when a full-game moneyline for this fixture was observed at all, even
    #: if the verdict named no side and so no side's price could be taken.
    market_observed: bool = False
    standings_evidence: Dict[str, object] = field(default_factory=dict)

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
        return (
            f"    [{RESEARCH_BUCKET}] {self.home} vs {self.away} · {self.scheduled_utc} · "
            f"MAGOLIDE_SCORE full_game -> {v.prediction} · selection={selection} · "
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
        lines = [f"  {self.label}"]
        lines.extend(
            f"    leg: {leg.selection} {leg.home} vs {leg.away} "
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

    resolved_config = config or MaGolideConfig()
    scorer_rows = from_published(published, derived=derived)
    index = {row.team: row for row in scorer_rows}
    sections = {
        canonical_team_name(str(row.team_name)): row.section
        for row in published
        if row.section
    }
    by_id, by_pair = _price_index(snapshots)

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

        selection = verdict.prediction if verdict.prediction in _SELECTIONS else None
        snapshot = by_id.get(game.game_id) or by_pair.get((home_name, away_name))
        out.append(
            ResearchPickRow(
                league=league,
                game_id=game.game_id,
                home=game.home_name,
                away=game.away_name,
                scheduled_utc=game.scheduled_utc,
                verdict=verdict,
                selection=selection,
                decimal_price=_side_price(snapshot, selection),
                price_book=(snapshot.book if snapshot is not None else None),
                price_provenance=(
                    snapshot.price_provenance.value
                    if snapshot is not None and snapshot.price_provenance is not None
                    else None
                ),
                observed_utc=(snapshot.observed_utc if snapshot is not None else None),
                market_observed=snapshot is not None,
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

    A row without a selection (``RISKY``) or without a price proposes nothing: a
    slip needs a side and a number, and inventing either is how a research card
    starts looking like advice. Slips are formed adjacent-first and never repeat a
    fixture, matching the grouping the existing shadow sections use.
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
    selected = sum(1 for row in rows if row.has_selection)
    neutral = sum(
        1 for row in rows if row.standings_evidence.get("neutral_defaults_applied")
    )
    lines = [
        "",
        f"{RESEARCH_LABEL} — {league} {date_label} — {len(rows)} fixture(s), "
        f"{selected} with a side, {priced} priced, {len(tickets)} paper slip(s)",
        "- Research only: a published standings table and an aggregate average price. "
        "No named-book quote, no stake, no receipt, no bank or ROI effect.",
        "- MAGOLIDE_SCORE is not a probability: uncalibrated_logistic = "
        "sigmoid(score/30) over validator defaults, shown as a research number.",
        f"- Standings provenance: {PUBLISHED_SOURCE}; price provenance: {PRICE_LABEL}.",
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
