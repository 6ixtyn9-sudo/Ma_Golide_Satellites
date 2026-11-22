"""Predixsport prediction provider — authenticated, disabled by default.

Terms & budget (DOCUMENTATION_ONLY)
-----------------------------------
Predixsport's free tier is 200 SHARED calls per calendar month for personal,
non-commercial use (see https://www.predixsport.com/api-access and
https://www.predixsport.com/pricing).  Budget discipline: ONE upcoming-list
call per refresh plus per-match detail calls only where a total is needed,
served from the caller's cache — this module performs a call only when asked
and never caches internally, and the 200-call allowance is shared across
keys/connectors, so callers must not poll each match hourly.

Evidence status: DOCUMENTATION_ONLY
-----------------------------------
The endpoints below are documented publicly (OpenAPI / docs). The
unauthenticated live sweep on 2026-11-06 returned ``Missing Authorization
header`` — recorded here as evidence that ``Authorization: Bearer <key>`` is
required.  No AUTHENTICATED live response has ever been observed, so the
real response schema is unconfirmed and everything in this module stays
labelled ``DOCUMENTATION_ONLY`` until a real-key smoke test is run and
recorded.  The committed fixtures under ``tests/fixtures/predixsport/`` are
SYNTHETIC payloads authored from the published documentation — including the
2026-11-22 quickstart and ``/v1/predictions/nba/{match_id}`` example, which
return ``home_team`` / ``away_team`` objects with ``win_probability``; they are
never ``LIVE_VERIFIED`` evidence.

Endpoints consumed (documented schema)
--------------------------------------
- ``GET {BASE_URL}/v1/predictions/{league}/upcoming?days={days}``
  (league ``nba`` | ``euroleague``) — list rows with ``match_id``,
  ``probabilities`` (e.g. ``{"home": 0.42, "away": 0.58}``), and, per the
  observed docs, ``predicted_point_diff`` (HOME minus AWAY) and
  ``produced_by``.
- ``GET {BASE_URL}/v1/predictions/{league}/{match_id}`` — adds
  ``predicted_total_points`` (plus spread/total distributions, which this
  adapter does not consume).

Normalisation rules
-------------------
- Derived scores when detail has BOTH ``predicted_total_points`` AND
  ``predicted_point_diff``: ``home = (total + diff) / 2``,
  ``away = (total - diff) / 2`` — ``diff`` is HOME MINUS AWAY, so a negative
  diff flips the score edge.  Probabilities-only records keep scores
  ``None`` (no fabrication from probabilities).
- The downstream moneyline selection comes from win probability only. The
  winner, spread and total models are independent: a probability favourite
  that disagrees with the margin sign is PRESERVED, never "repaired" into
  artificial consistency.
- ``produced_by`` is preserved from the payload (early-season fallback
  predictions must stay distinguishable from the mature model).

Auth & failure behaviour
------------------------
- ``Authorization: Bearer <PREDIXSPORT_API_KEY>`` is read from the
  environment inside this module; the key value is never logged
  (``http.redact`` scrubs it from any surfaced error text).
- ``is_enabled()`` is False without a non-empty key; the fetch functions
  then return ``[]``/``None`` WITHOUT calling or raising, so the keyless
  pipeline never breaks (the caller records ``UNAUTHORIZED`` health).
  Calling ``_headers()`` (or the prod fetch factory) directly without a key
  raises :class:`UnauthorizedError`.
- ``fetch_factory(provider)`` returns the ``FetchFn`` used for the actual
  calls (prod: :func:`default_fetch_factory` builds an ``HttpClient`` with
  the Bearer header baked in; tests: a fixture loader).
"""

from __future__ import annotations

import os
from typing import Any, Callable, Dict, List, Optional

from ..http import HttpClient, redact
from ..identity import now_utc_iso
from ..models import CoverageGap, CoverageState, Prediction
from . import FetchFn

BASE_URL = "https://api.predixsport.com"
# Horizon for the upcoming list. The provider documents one day as the correct
# window — "NBA predictions land on game day (days=1 is the right horizon;
# days=14 returns 0 matches in steady state)" — so the adapter's former fixed
# days=7 asked for a window it documents as empty.
UPCOMING_DAYS = 1
# Minimum spacing between upstream calls, per league. The free tier is 200
# SHARED requests per calendar month, so an every-tick call (~32/day observed)
# exhausts it in about a week. 18h gives ~40 calls/month per league — 80 for
# NBA + EuroLeague — while still refreshing well inside the 1-day horizon the
# provider publishes on.
UPCOMING_WINDOW_HOURS = 18
ENV_KEY = "PREDIXSPORT_API_KEY"
PROVIDER = "predixsport"
PREDICTION_MODEL_VERSION = "adapter:predixsport"
VALID_LEAGUES = ("nba", "euroleague")

EVIDENCE = "DOCUMENTATION_ONLY"

# fetch_factory contract: called with the provider name ("predixsport") and
# returns the FetchFn that performs GETs (prod: an authenticated HttpClient
# wrapper; tests: a fixture loader).
FetchFactory = Callable[[str], FetchFn]


class UnauthorizedError(RuntimeError):
    """Predixsport credentials are missing — the API requires a Bearer key.

    Recorded evidence: unauthenticated calls on 2026-11-06 were rejected with
    ``Missing Authorization header``.
    """


def is_enabled() -> bool:
    """True when ``PREDIXSPORT_API_KEY`` is present and non-empty."""
    return bool(os.environ.get(ENV_KEY, "").strip())


def _headers() -> Dict[str, str]:
    """Bearer auth header built from the environment.

    Raises :class:`UnauthorizedError` when the key is missing/blank — direct
    callers get a precise signal instead of a header that is guaranteed to
    be rejected.  The key value is never logged.
    """
    key = os.environ.get(ENV_KEY, "").strip()
    if not key:
        raise UnauthorizedError(
            f"{ENV_KEY} is not set; Predixsport requires "
            "'Authorization: Bearer <key>' (unauthenticated calls on "
            "2026-11-06 returned 'Missing Authorization header')"
        )
    # HTTP header values are latin-1: a pasted placeholder ('…'), a smart
    # quote, or any other non-ASCII character otherwise surfaces as an
    # opaque UnicodeEncodeError from deep inside requests/urllib3. Report
    # the offending position and code point, never the value itself.
    for index, char in enumerate(key):
        if ord(char) > 0x7E or ord(char) < 0x21:
            raise UnauthorizedError(
                f"{ENV_KEY} contains a non-printable-ASCII character at index "
                f"{index} (U+{ord(char):04X}); HTTP header values are latin-1, "
                "so this cannot be sent. Re-paste the key — documentation "
                "placeholders such as '\u2026' are not credentials."
            )
    return {"Authorization": f"Bearer {key}"}


def default_fetch_factory(provider: str) -> FetchFn:
    """Production ``fetch_factory``: an HttpClient with the Bearer header.

    ``_headers()`` is evaluated once at factory time, so a missing key fails
    fast with :class:`UnauthorizedError` instead of mid-sweep.
    """
    if provider != PROVIDER:  # defensive: factories are provider-keyed
        raise ValueError(f"predixsport factory asked for provider {provider!r}")
    headers = _headers()
    client = HttpClient()
    return lambda url: client.get_json(url, headers=headers)


def _league(league: str) -> str:
    league_l = str(league or "").strip().lower()
    if league_l not in VALID_LEAGUES:
        raise ValueError(f"predixsport league must be one of {VALID_LEAGUES}, got {league!r}")
    return league_l


def _num(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _team_block(row: Dict[str, Any], side: str) -> Dict[str, Any]:
    """Return the nested ``home_team`` / ``away_team`` object, if present."""

    nested = row.get(f"{side}_team")
    return nested if isinstance(nested, dict) else {}


def _team_name(row: Dict[str, Any], side: str) -> Optional[str]:
    """Team name from the documented shape, with older flat keys as fallback.

    The published quickstart returns ``home_team`` / ``away_team`` objects
    carrying ``name`` and ``win_probability``; earlier synthetic fixtures used
    flat ``home`` / ``away`` strings. Both are read; nothing is invented.
    """

    direct = row.get(side)
    if isinstance(direct, str) and direct.strip():
        return direct
    if isinstance(direct, dict) and isinstance(direct.get("name"), str):
        return direct["name"]
    nested = _team_block(row, side).get("name")
    return nested if isinstance(nested, str) and nested.strip() else None


def _tipoff_utc(row: Dict[str, Any]) -> Optional[str]:
    """Tip-off timestamp from the documented field names, if any is present.

    The published list/detail examples carry no tip-off, only
    ``prediction_date``; a date is not a tip-off, so it is never promoted into
    one. A row without a timestamp keeps ``source_tipoff_utc`` empty and the
    pipeline's identity join reports a coverage gap instead of guessing.
    """

    for field in ("tipoff_utc", "commence_time", "start_time", "start_utc"):
        value = row.get(field)
        if isinstance(value, str) and value.strip():
            return value
    return None


def _probabilities(row: Dict[str, Any]) -> Any:
    """Extract (p_home, p_away) in 0..1 from the documented locations.

    Read order: ``probabilities.home`` / ``probabilities.away`` (earlier
    documented normalisation), then the live-documented
    ``home_team.win_probability`` / ``away_team.win_probability``, then flat
    ``p_home`` / ``p_away``. Anything else stays ``None`` — no invention.
    """
    probs = row.get("probabilities")
    p_home = p_away = None
    if isinstance(probs, dict):
        p_home = _num(probs.get("home"))
        p_away = _num(probs.get("away"))
    if p_home is None:
        p_home = _num(_team_block(row, "home").get("win_probability"))
    if p_away is None:
        p_away = _num(_team_block(row, "away").get("win_probability"))
    if p_home is None:
        p_home = _num(row.get("p_home"))
    if p_away is None:
        p_away = _num(row.get("p_away"))
    return p_home, p_away


def upcoming_url(league: str, days: int) -> str:
    return f"{BASE_URL}/v1/predictions/{_league(league)}/upcoming?days={int(days)}"


def detail_url(league: str, match_id: str) -> str:
    return f"{BASE_URL}/v1/predictions/{_league(league)}/{match_id}"


def fetch_upcoming(
    league: str,
    days: int,
    fetch_factory: FetchFactory,
    gaps: Optional[List[CoverageGap]] = None,
    diagnostics: Optional[Dict[str, Any]] = None,
) -> List[Prediction]:
    """Fetch the upcoming predictions list for ``league``.

    Returns ``[]`` when the provider is disabled (no key; the caller records
    ``UNAUTHORIZED`` health) — it never raises for a missing key.  Makes
    exactly ONE list call through ``fetch_factory(PROVIDER)``.  Malformed
    entries are skipped defensively and, when the caller passes a ``gaps``
    list, a ``CoverageGap`` is appended for each skip so nothing disappears
    silently.

    ``diagnostics`` (optional dict, filled in place) records the *shape* of
    the response — top-level key names, the reader used, and the row count —
    never values.  This exists because the schema has never been observed
    live: a populated payload delivered under any key other than ``matches``
    would otherwise be indistinguishable from a genuinely empty feed, and
    both would report ``records_returned: 0``.
    """
    if not is_enabled():
        return []
    league = _league(league)
    fetch = fetch_factory(PROVIDER)
    payload = fetch(upcoming_url(league, days))
    if isinstance(payload, dict):
        rows = payload.get("matches")
        row_source = "matches" if "matches" in payload else "absent"
        payload_keys: Optional[List[str]] = sorted(str(key) for key in payload.keys())
    else:
        rows = payload
        row_source = "bare_list" if isinstance(payload, list) else "unexpected_type"
        payload_keys = None
    if diagnostics is not None:
        diagnostics.update(
            {
                "response_type": type(payload).__name__,
                "row_source": row_source,
                "payload_keys": payload_keys,
                "rows_seen": len(rows) if isinstance(rows, list) else 0,
            }
        )
    out: List[Prediction] = []
    for row in rows or []:
        try:
            if not isinstance(row, dict):
                raise ValueError("entry is not a JSON object")
            match_id = row.get("match_id") or row.get("id")
            if match_id is None or str(match_id).strip() == "":
                raise ValueError("missing match_id")
            p_home, p_away = _probabilities(row)
            if p_home is None or p_away is None:
                raise ValueError("missing probabilities")
            # Upcoming rows carry no predicted_total_points, so scores stay
            # None (no fabrication).  Prediction.__post_init__ rejects
            # probabilities outside 0..1 (e.g. percentage-scale values) and
            # validate_orientation rejects sums that are not ~1 — both
            # quarantine the entry via the gap below.
            prediction = Prediction(
                # ``game_id`` is the provider's event identity at this adapter
                # boundary.  The pipeline joins it to one canonical game only
                # by an existing provider ID or by strict home/away/UTC-tip
                # identity, then preserves this ID in ``auxiliary_metrics``.
                game_id=str(match_id),
                provider=PROVIDER,
                observed_utc=now_utc_iso(),
                model=f"predixsport-{_league(league)}",
                model_version=PREDICTION_MODEL_VERSION,
                produced_by=row.get("produced_by") or PROVIDER,
                prob_home=p_home,
                prob_away=p_away,
                expected_margin=_num(row.get("predicted_point_diff")),
                auxiliary_metrics={
                    "source_match_id": str(match_id),
                    "source_home_name": _team_name(row, "home"),
                    "source_away_name": _team_name(row, "away"),
                    "source_tipoff_utc": _tipoff_utc(row),
                },
            )
            prediction.validate_orientation()
            out.append(prediction)
        except (TypeError, ValueError, AttributeError, KeyError) as exc:
            if gaps is not None:
                gaps.append(
                    CoverageGap(
                        game_id=str(row.get("match_id")) if isinstance(row, dict) else None,
                        required_field="prediction",
                        reason=CoverageState.MISSING,
                        provider=PROVIDER,
                        detail=f"malformed upcoming entry skipped: {type(exc).__name__}: {exc}",
                    )
                )
    return out


def fetch_match_detail(
    match_id: str, league: str, fetch_factory: FetchFactory
) -> Optional[Prediction]:
    """Fetch one match's detail and derive expected scores.

    ``home = (total + diff) / 2`` and ``away = (total - diff) / 2`` with
    ``diff`` = predicted_point_diff = HOME minus AWAY (sign-correct: a
    negative diff flips the score edge).  When only probabilities exist the
    scores stay ``None``. The downstream moneyline winner is computed from
    probabilities and is deliberately NOT forced to agree with the margin
    sign — independent-model disagreement is preserved. Returns ``None``
    when disabled; a contradictory payload
    raises ``ValueError`` (caller quarantines).
    """
    if not is_enabled():
        return None
    league = _league(league)
    fetch = fetch_factory(PROVIDER)
    payload = fetch(detail_url(league, match_id))
    row = payload if isinstance(payload, dict) else {}
    p_home, p_away = _probabilities(row)
    total = _num(row.get("predicted_total_points"))
    diff = _num(row.get("predicted_point_diff"))
    home_score = away_score = None
    if total is not None and diff is not None:
        home_score = (total + diff) / 2.0
        away_score = (total - diff) / 2.0
    prediction = Prediction(
        game_id=str(row.get("match_id") or match_id),
        provider=PROVIDER,
        observed_utc=now_utc_iso(),
        model=f"predixsport-{_league(league)}",
        model_version=PREDICTION_MODEL_VERSION,
        produced_by=row.get("produced_by") or PROVIDER,
        prob_home=p_home,
        prob_away=p_away,
        expected_home_score=home_score,
        expected_away_score=away_score,
        expected_total=total,
        expected_margin=diff,  # home minus away, preserved as-is
    )
    prediction.validate_orientation()
    return prediction


def smoke(league: str, fetch_factory: FetchFactory) -> Dict[str, Any]:
    """Provider smoke report; evidence is always labelled, never claims LIVE.

    Disabled: ``{'enabled': False, 'evidence': 'DOCUMENTATION_ONLY', ...}``
    with no call made.  Enabled: performs the one upcoming-list call through
    ``fetch_factory`` and reports counts; the evidence label remains
    ``DOCUMENTATION_ONLY`` because no authenticated response has been
    observed — operators must relabel to LIVE_VERIFIED only after a real-key
    live smoke run is recorded.
    """
    if not is_enabled():
        return {
            "provider": PROVIDER,
            "league": league,
            "enabled": False,
            "evidence": EVIDENCE,
            "note": f"{ENV_KEY} not set; provider disabled, no call made.",
        }
    gaps: List[CoverageGap] = []
    try:
        predictions = fetch_upcoming(league, days=7, fetch_factory=fetch_factory, gaps=gaps)
    except Exception as exc:
        return {
            "provider": PROVIDER,
            "league": league,
            "enabled": True,
            "evidence": EVIDENCE,
            "upcoming_count": 0,
            "gaps": 0,
            "error": redact(f"{type(exc).__name__}: {exc}"),
            "note": "Call failed; key material is redacted from this text.",
        }
    return {
        "provider": PROVIDER,
        "league": league,
        "enabled": True,
        "evidence": EVIDENCE,
        "evidence_note": (
            "Schema is documentation/fixture-verified only; no authenticated "
            "live response has been observed (unauthenticated sweep 2026-11-06 "
            "got 'Missing Authorization header'). Relabel to LIVE_VERIFIED "
            "only after a real-key live smoke test is run and recorded."
        ),
        "upcoming_count": len(predictions),
        "gaps": len(gaps),
    }
