"""When a host publishes no licence, the only thing that can open a lane is a person.

China (``cba.net.cn``) and Lithuania (``nkl.lt``) publish no terms document of any
kind — not a permissive one, not a prohibitive one. Greece publishes terms whose
licence is scoped to personal non-commercial use. In all three the read is not
authorised by the host; it is authorised, if at all, by a named operator decision.

This records those decisions as data: who, when, what scope, on what basis. It
deliberately does **not** switch a lane on. An acceptance opens the licence gate
and nothing else — the adapter still has to exist and be replayed against a served
page, which is the same rule the other lanes were held to. Keeping that separation
matters: an acceptance that silently activated a lane would be a permission nobody
could audit, granted by editing a registry string.
"""

from __future__ import annotations

import pytest

from ingestion.league_registry import OperatorAcceptance, league_spec, parse_league
from ingestion.models import League

#: Lanes whose licence gate was opened by a person rather than by a host.
ACCEPTED = (
    "GREECE_BASKET_LEAGUE",
    "CHINA_WCBA_W",
    "LITHUANIA_NKL",
    "AUSTRALIA_WNBL_W",
)


@pytest.mark.parametrize("league", ACCEPTED)
def test_every_recorded_acceptance_names_who_decided_and_what_they_accepted(league):
    record = league_spec(parse_league(league)).operator_acceptance
    assert record is not None, f"{league} has no recorded operator decision"
    assert record.decided_by.strip()
    assert record.decided_on == "2026-11-24"
    assert record.scope.strip()
    assert record.basis.strip()
    # The scope says what may be read and bounds what may be done with it.
    assert "no redistribution" in record.scope.lower()


#: Lanes whose acceptance opened the licence gate but whose adapter has not yet
#: been replayed against served pages. NKL is deliberately absent: its adapter was
#: replayed live on 2026-11-24 and settled three games, which is the only thing
#: that turns an acceptance into a lit lane.
ACCEPTED_BUT_DARK = tuple(league for league in ACCEPTED if league != "LITHUANIA_NKL")


@pytest.mark.parametrize("league", ACCEPTED_BUT_DARK)
def test_an_acceptance_does_not_by_itself_switch_a_results_lane_on(league):
    """The rule the whole project runs on: a decision is not a served page."""

    spec = league_spec(parse_league(league))
    assert spec.operator_acceptance is not None
    assert spec.results_source == "source_unavailable", (
        f"{league} must stay dark until an adapter is replayed against a captured page"
    )


def test_nkl_is_lit_by_a_served_page_rather_than_by_its_acceptance():
    """The other half of the same rule, in the one lane where it has been met."""

    from ingestion.providers import nkl

    spec = league_spec(League.LITHUANIA_NKL)

    assert spec.operator_acceptance is not None
    assert spec.results_source == nkl.PROVIDER
    # The lane names the reader that settled a game, not a promise of one.
    assert spec.results_source == "nkl"
    assert "Crawl-delay" in spec.operator_acceptance.scope


def test_the_two_hosts_without_any_terms_document_say_so_in_their_basis():
    for league in ("CHINA_WCBA_W", "LITHUANIA_NKL"):
        basis = league_spec(parse_league(league)).operator_acceptance.basis
        assert "no terms document" in basis


def test_greeces_record_carries_the_clause_that_was_actually_read():
    basis = league_spec(League.GREECE_BASKET_LEAGUE).operator_acceptance.basis
    assert "personal" in basis and "non-commercial" in basis


def test_a_record_without_a_decider_is_refused():
    with pytest.raises(ValueError, match="non-empty decided_by"):
        OperatorAcceptance(decided_by="  ", decided_on="2026-11-24", scope="s", basis="b")
    with pytest.raises(ValueError, match="non-empty basis"):
        OperatorAcceptance(decided_by="operator", decided_on="2026-11-24", scope="s", basis="")


def test_a_lane_with_no_decision_records_none_rather_than_a_default():
    """Absence is not refusal, and it must not look like consent."""

    assert league_spec(League.NBA).operator_acceptance is None
    assert league_spec(League.ISRAEL_PREMIER_LEAGUE).operator_acceptance is None


def test_the_record_is_reported_as_data_not_as_prose():
    record = league_spec(League.CHINA_WCBA_W).operator_acceptance
    assert record.as_dict() == {
        "decided_by": "operator",
        "decided_on": "2026-11-24",
        "scope": record.scope,
        "basis": record.basis,
    }
