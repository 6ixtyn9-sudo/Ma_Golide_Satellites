"""An exchange price is not a bookmaker's price, and must not be stamped as one.

Nothing in the repository could express an exchange quote: ``PriceProvenance``
had ``NAMED_BOOK``, ``AGGREGATE_CONSENSUS`` and ``NO_PRICE``, so a Betfair
API-NG offer -- a real matched counterparty price, but nobody's bookmaker quote
-- had to be recorded as one of those three, all of them wrong in a way that
matters: ``NAMED_BOOK`` is what ``ticket_price_eligible`` trusts.

These tests pin the vocabulary and the gate that keeps an exchange observation
out of the paper-ticket lane until a capture and its terms authorise otherwise.
"""

from __future__ import annotations

from forecasting.types import BookEvidence
from ingestion.league_registry import price_evidence_policy
from ingestion.models import (
    OperationalValidation,
    PriceProvenance,
    SourcePermission,
    League,
)


def _evidence(provenance: PriceProvenance) -> BookEvidence:
    return BookEvidence(
        provider="betfair",
        book="betfair_exchange",
        observed_utc="2026-11-23T18:40:00Z",
        source_modified_utc=None,
        pre_match=True,
        main_line=True,
        frozen=False,
        prices={"home": 1.98, "away": 1.99},
        price_provenance=provenance,
    )


def test_exchange_is_its_own_provenance():
    assert PriceProvenance.EXCHANGE.value == "exchange"
    assert PriceProvenance.EXCHANGE is not PriceProvenance.NAMED_BOOK
    assert PriceProvenance.EXCHANGE is not PriceProvenance.AGGREGATE_CONSENSUS
    # The wire values of the three pre-existing members are untouched.
    assert {
        PriceProvenance.NAMED_BOOK.value,
        PriceProvenance.AGGREGATE_CONSENSUS.value,
        PriceProvenance.NO_PRICE.value,
    } == {"named_book", "aggregate_consensus", "no_price"}


def test_an_unregistered_exchange_is_never_stamped_as_a_named_book():
    """The unknown-provider fallback must not launder an exchange identity."""

    for spelling in ("betfair", "Betfair", "BETFAIR_EXCHANGE", " betfair "):
        policy = price_evidence_policy(League.SPAIN_ACB, spelling)
        assert policy.price_provenance is PriceProvenance.EXCHANGE, spelling
        assert policy.acquire is False, spelling
        assert policy.source_permission is SourcePermission.UNRESOLVED, spelling
        assert policy.operational_validation is OperationalValidation.UNVERIFIED, spelling


def test_a_genuinely_unknown_provider_keeps_the_documented_fallback():
    # Pinned so the exchange rule above is not implemented by loosening this.
    policy = price_evidence_policy(League.SPAIN_ACB, "some_new_book")
    assert policy.price_provenance is PriceProvenance.NAMED_BOOK
    assert policy.acquire is False


def test_an_exchange_quote_can_be_graded_but_never_paper_ticketed():
    exchange = _evidence(PriceProvenance.EXCHANGE)
    named = _evidence(PriceProvenance.NAMED_BOOK)
    aggregate = _evidence(PriceProvenance.AGGREGATE_CONSENSUS)

    assert exchange.ticket_price_eligible is False
    assert named.ticket_price_eligible is True  # the lane itself still works
    assert aggregate.ticket_price_eligible is False
    # The price is still there to grade against -- refusing the ticket is not
    # the same as discarding the observation.
    assert exchange.prices["home"] == 1.98
