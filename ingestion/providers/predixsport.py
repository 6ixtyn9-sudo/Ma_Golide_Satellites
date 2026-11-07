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
The endpoints below are documented publicly (OpenAPI / docs), but the
unauthenticated live sweep on 2026-11-06 returned ``Missing Authorization
header`` — recorded here as evidence that ``Authorization: Bearer <key>`` is
required.  No AUTHENTICATED live response has ever been observed, so the
real response schema is unconfirmed and everything in this module stays
labelled ``DOCUMENTATION_ONLY`` until a real-key smoke test is run and
recorded.  The committed fixtures under ``tests/fixtures/predixsport/`` are
SYNTHETIC payloads authored from the documented schema (see that directory's
README); they are never ``LIVE_VERIFIED`` evidence.

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
ENV_KEY = "PREDIXSPORT_API_KEY"
PROVIDER = "predixsport"
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
        raise ValueError(
            f"predixsport league must be one of {VALID_LEAGUES}, got {league!r}"
        )
    return league_l


def _num(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _probabilities(row: Dict[str, Any]) -> Any:
    """Extract (p_home, p_away) in 0..1 from the documented locations.

    Normalised field names (from the synthetic fixtures built on the
    documented schema): ``probabilities.home`` / ``probabilities.away``, with
    flat ``p_home`` / ``p_away`` accepted as a fallback.  Anything else stays
    ``None`` — no invention.
    """
    probs = row.get("probabilities")
    p_home = p_away = None
    if isinstance(probs, dict):
        p_home = _num(probs.get("home"))
        p_away = _num(probs.get("away"))
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
) -> List[Prediction]:
    """Fetch the upcoming predictions list for ``league``.

    Returns ``[]`` when the provider is disabled (no key; the caller records
    ``UNAUTHORIZED`` health) — it never raises for a missing key.  Makes
    exactly ONE list call through ``fetch_factory(PROVIDER)``.  Malformed
    entries are skipped defensively and, when the caller passes a ``gaps``
    list, a ``CoverageGap`` is appended for each skip so nothing disappears
    silently.
    """
    if not is_enabled():
        return []
    league = _league(league)
    fetch = fetch_factory(PROVIDER)
    payload = fetch(upcoming_url(league, days))
    rows = payload.get("matches") if isinstance(payload, dict) else payload
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
                game_id=str(match_id),
                provider=PROVIDER,
                observed_utc=now_utc_iso(),
                model=f"predixsport-{_league(league)}",
                produced_by=row.get("produced_by") or PROVIDER,
                prob_home=p_home,
                prob_away=p_away,
                expected_margin=_num(row.get("predicted_point_diff")),
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
                        detail=f"malformed upcoming entry skipped: "
                        f"{type(exc).__name__}: {exc}",
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
