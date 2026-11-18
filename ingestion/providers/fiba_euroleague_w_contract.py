"""Pure, shared wire-contract helpers for FIBA EuroLeague Women.

The provider and the dependency-free adaptive gate both consume the same GDAP
schedule.  Keep the small facts they must agree on here, without importing the
HTTP client, models, or any third-party dependency.  This prevents status
semantics or the public APIM fallback value from drifting between those two
paths.
"""

from __future__ import annotations

import re
from typing import Any, Mapping

COMPETITION_ID = 209127
COMPETITION_CODE = "ELW"
GDAP_BASE_URL = "https://digital-api.fiba.basketball/hapi"
SCHEDULE_ENDPOINT = "getgdapgamesbycompetitionid"
KEY_PAGE_URL = "https://www.fiba.basketball/en"
KEY_HEADER = "Ocp-Apim-Subscription-Key"
# Public value distributed to browsers by FIBA's web application.  It is a
# fallback, never an application secret.  Runtime readers refresh from the
# public page when required and must never log the value.
PUBLIC_FALLBACK_SUBSCRIPTION_KEY = "898cd5e7389140028ecb42943c47eb74"

# GDAP statusCode values live-observed for competition 209127 (INIT, VALID),
# plus explicit terminal values the FIBA adapter supports.  Anything else is
# unknown rather than silently treated as an ordinary scheduled fixture.
_RECOGNISED_OPEN_STATUS_CODES = frozenset({"", "INIT", "VALID"})
_CANCELLED_STATUS_CODES = frozenset({"CANCELLED", "CANCELED"})
_ABANDONED_STATUS_CODES = frozenset({"ABANDONED", "VOID"})


def scores_present(row: Mapping[str, Any]) -> bool:
    """Whether GDAP has two numeric final-score values (excluding booleans)."""

    return all(
        isinstance(row.get(field), (int, float)) and not isinstance(row.get(field), bool)
        for field in ("teamAScore", "teamBScore")
    )


def schedule_status(row: Mapping[str, Any]) -> str:
    """Return the shared conservative GDAP status vocabulary.

    Values are lower-case strings so this module remains model-free.  The
    provider maps them to ``GameStatus``; the gate accepts only scheduled,
    live, and final rows and fail-opens on terminal or unknown rows.
    """

    if row.get("isPostponed") is True:
        return "postponed"
    status_code = str(row.get("statusCode") or "").upper()
    if status_code in _CANCELLED_STATUS_CODES:
        return "cancelled"
    if status_code in _ABANDONED_STATUS_CODES:
        return "abandoned"
    if row.get("isLive") is True:
        return "live"
    if (
        str(row.get("gameResultStatusCode") or "").upper() == "VALID"
        and row.get("liveGameStatus") == 999
        and scores_present(row)
    ):
        return "final"
    if status_code in _RECOGNISED_OPEN_STATUS_CODES:
        return "scheduled"
    return "unknown"


def has_complete_team_identity(row: Mapping[str, Any]) -> bool:
    """Whether the row can become a canonical fixture with two FIBA teams."""

    for side in ("teamA", "teamB"):
        team = row.get(side)
        if not isinstance(team, Mapping):
            return False
        if (
            team.get("teamId") is None
            or not str(team.get("officialName") or team.get("shortName") or "").strip()
            or not str(team.get("code") or "").strip()
        ):
            return False
    return True


def has_unassigned_team_slot(row: Mapping[str, Any]) -> bool:
    """Whether GDAP has not assigned one/both bracket teams yet.

    A null/absent team object is a normal future-bracket placeholder.  A team
    object with malformed identity is *not* this benign condition and must be
    treated as provider uncertainty by callers.
    """

    return row.get("teamA") is None or row.get("teamB") is None


def page_subscription_key(page: str) -> str | None:
    """Extract FIBA's public browser key from normal or Flight-escaped text."""

    if not isinstance(page, str):
        return None
    normalized = page
    # Flight strings can escape the quote once or twice.  Repeated replacement
    # makes the public assignment readable without trying to decode arbitrary
    # JavaScript.
    for _ in range(2):
        decoded = normalized.replace(r'\"', '"')
        if decoded == normalized:
            break
        normalized = decoded
    match = re.search(
        r"NEXT_CLIENT_APIM_SUBSCRIPTION_KEY[\"']?\s*[:=]\s*[\"']([0-9a-f]{32})[\"']",
        normalized,
        flags=re.IGNORECASE,
    )
    return match.group(1) if match else None


__all__ = [
    "COMPETITION_CODE",
    "COMPETITION_ID",
    "GDAP_BASE_URL",
    "KEY_HEADER",
    "KEY_PAGE_URL",
    "PUBLIC_FALLBACK_SUBSCRIPTION_KEY",
    "SCHEDULE_ENDPOINT",
    "has_complete_team_identity",
    "has_unassigned_team_slot",
    "page_subscription_key",
    "schedule_status",
    "scores_present",
]
