"""Dependency-free dated paper-ticket publication timing policy.

The adaptive scheduler imports this module before project dependencies are
installed, and the full operator imports the same values when it constructs an
immutable ticket/no-bet decision.  Keeping the small league-specific exception
here prevents their cutoffs from drifting apart.
"""

from __future__ import annotations

from datetime import time


TICKET_PUBLICATION_TIMEZONE = "Africa/Johannesburg"
DEFAULT_TICKET_PUBLICATION_LOCAL_TIME = time(hour=9)
NBA_TICKET_PUBLICATION_LOCAL_TIME = time(hour=18)
DEFAULT_TICKET_DECISION_POLICY = "0900_sast_or_first_slate_tip_minus_60_minutes"
NBA_TICKET_DECISION_POLICY = "1800_sast_or_first_slate_tip_minus_60_minutes"


def ticket_publication_local_time(league_key: str) -> time:
    """Return the local dated-card publication time for one canonical league."""

    if league_key == "NBA":
        return NBA_TICKET_PUBLICATION_LOCAL_TIME
    return DEFAULT_TICKET_PUBLICATION_LOCAL_TIME


def ticket_decision_policy(league_key: str) -> str:
    """Return the audit label paired with the league's publication time."""

    if league_key == "NBA":
        return NBA_TICKET_DECISION_POLICY
    return DEFAULT_TICKET_DECISION_POLICY
