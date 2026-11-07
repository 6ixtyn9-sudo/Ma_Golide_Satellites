"""Hardwood (roni-altshuler/nba_predictor) forecast artifact adapter.

Evidence status
---------------
The artifact at ``FORECASTS_URL`` is a public JSON file from an MIT-licensed
repository, observed live on 2026-11-06; the committed fixture
``tests/fixtures/hardwood/game_forecasts.json`` is a slice captured from that
live artifact, so this provider is ``LIVE_VERIFIED``.

Sign convention (verified against the live artifact, 2026-11-06)
----------------------------------------------------------------
``exp_margin`` is HOME minus AWAY.  For game 401909287 (Washington Wizards @
Atlanta Hawks) the home side Atlanta is favoured: p_home 0.713845,
exp_home_score 117.1 vs exp_away_score 109.47, exp_margin 7.636 = 117.1 -
109.47 (within rounding; all five fixture games agree with the
home-minus-away reading).  Values are stored AS-IS — the adapter never
rewrites, re-derives or "repairs" provider numbers (master prompt section 8
rules 1 and 9).  ``Prediction.validate_orientation()`` is run so a
self-contradictory entry (margin flipped versus its own expected scores)
quarantines as a gap instead of being silently corrected or written.

Join policy
-----------
Join by ``games[].game_id`` == ``game.provider_ids['espn']`` (the ESPN event
id; the canonical ``game.game_id`` is the fallback), then VALIDATE identity:
home/away names must match case-insensitively via
``identity.normalize_team_name`` AND the entry's ``date_utc`` must sit within
``identity.TIME_TOLERANCE`` (30 min) of the game's scheduled time.  ANY
mismatch returns ``None`` plus a recorded ``CoverageGap`` — never a
best-guess join, and never a team-only join across different dates.

Provenance recorded on every prediction: ``model`` = ``"hardwood"``,
``model_version`` = artifact ``model_version``, ``produced_by`` =
``"roni-altshuler/nba_predictor"``, ``source_revision`` = ``artifact_sha256``,
``observed_utc`` = ``source_modified_utc`` = artifact ``generated_at`` (the
artifact generation time IS the observation).  Selection/averaging with other
providers happens in the prediction selector, never here — this module
returns one coherent single-provider bundle or nothing.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from ..identity import (
    TIME_TOLERANCE,
    normalize_team_name,
    parse_utc,
    times_within,
    to_utc_iso,
)
from ..models import CoverageGap, CoverageState, Game, Prediction
from . import FetchFn

FORECASTS_URL = (
    "https://raw.githubusercontent.com/roni-altshuler/nba_predictor/main/"
    "backend/data/predictions/game_forecasts.json"
)

PROVIDER = "hardwood"
PRODUCED_BY = "roni-altshuler/nba_predictor"

# Same default tolerance as Prediction.validate_orientation: expected-score
# differences below this are treated as rounding, not contradiction.
ORIENTATION_TOLERANCE = 0.25

# Tolerated clock skew when generated_at lies slightly in the future.
_FUTURE_SKEW = timedelta(minutes=10)


# ---------------------------------------------------------------------------
# Artifact container and loading
# ---------------------------------------------------------------------------


@dataclass
class ForecastsArtifact:
    """Validated Hardwood artifact metadata plus its games keyed by id.

    ``games_by_id`` maps ``str(games[].game_id)`` (the ESPN event id) to the
    raw game entry dict; duplicate ids keep the FIRST occurrence.
    """

    generated_at: str
    ingestion_status: str
    model_version: str
    trained_through: str
    artifact_sha256: str
    games_by_id: Dict[str, dict] = field(default_factory=dict)


def load_forecasts(fetch: FetchFn) -> ForecastsArtifact:
    """Fetch and structurally validate the raw forecast artifact.

    Required metadata (master prompt section 5 / PROVIDER_GUIDE): a parseable
    ``generated_at``, ``ingestion.status`` present, ``trained_through``
    present, plus ``model_version``/``artifact_sha256`` for provenance and a
    list of ``games``.  Raises ``ValueError`` naming every missing/invalid
    field so callers treat the provider as unhealthy instead of parsing a
    half-broken payload.  ``ingestion.status != 'ok'`` is NOT structural: the
    artifact is returned and :func:`freshness_state` maps it to
    ``CoverageState.PROVIDER_ERROR``.
    """
    artifact = fetch(FORECASTS_URL)
    if not isinstance(artifact, dict):
        raise ValueError("hardwood artifact is not a JSON object")

    problems: List[str] = []
    generated = artifact.get("generated_at")
    if not generated:
        problems.append("generated_at (missing)")
    else:
        try:
            parse_utc(generated)
        except (TypeError, ValueError) as exc:
            problems.append(f"generated_at (unparseable: {exc})")

    ingestion = artifact.get("ingestion")
    if not isinstance(ingestion, dict) or not ingestion.get("status"):
        problems.append("ingestion.status (missing)")
    for name in ("model_version", "trained_through", "artifact_sha256"):
        if not artifact.get(name):
            problems.append(f"{name} (missing)")
    if not isinstance(artifact.get("games"), list):
        problems.append("games (not a list)")
    if problems:
        raise ValueError("hardwood artifact invalid metadata: " + "; ".join(problems))

    games_by_id: Dict[str, dict] = {}
    for row in artifact["games"]:
        if isinstance(row, dict) and row.get("game_id") is not None:
            games_by_id.setdefault(str(row["game_id"]), row)

    return ForecastsArtifact(
        generated_at=str(generated),
        ingestion_status=str(ingestion.get("status", "")),
        model_version=str(artifact["model_version"]),
        trained_through=str(artifact["trained_through"]),
        artifact_sha256=str(artifact["artifact_sha256"]),
        games_by_id=games_by_id,
    )


def freshness_state(
    artifact: ForecastsArtifact,
    now_utc: Optional[str] = None,
    freshness_hours: float = 12.0,
) -> CoverageState:
    """Coverage state of an artifact against the freshness policy.

    ``STALE`` when ``generated_at`` is older than ``freshness_hours`` relative
    to ``now_utc`` (default: current time).  ``PROVIDER_ERROR`` for an
    ingestion that did not end ``ok``, unparseable timestamps, a
    ``trained_through`` after ``generated_at``, or a ``generated_at`` further
    than 10 minutes in the future.  ``OK`` otherwise.  Never raises.
    """
    if artifact.ingestion_status != "ok":
        return CoverageState.PROVIDER_ERROR
    try:
        generated_dt = parse_utc(artifact.generated_at)
        trained_dt = parse_utc(artifact.trained_through)
    except (TypeError, ValueError):
        return CoverageState.PROVIDER_ERROR
    if not artifact.trained_through or trained_dt > generated_dt:
        return CoverageState.PROVIDER_ERROR

    now_dt = parse_utc(now_utc) if now_utc else datetime.now(timezone.utc)
    age = now_dt - generated_dt
    if age < -_FUTURE_SKEW:
        return CoverageState.PROVIDER_ERROR
    if age > timedelta(hours=freshness_hours):
        return CoverageState.STALE
    return CoverageState.OK


# ---------------------------------------------------------------------------
# Prediction join
# ---------------------------------------------------------------------------


def _num(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def fetch_prediction(
    game: Game,
    fetch: FetchFn,
    freshness_hours: float = 12.0,
    now_utc: Optional[str] = None,
    gaps: Optional[List[CoverageGap]] = None,
) -> Optional[Prediction]:
    """Join one canonical game to its Hardwood forecast, or return ``None``.

    Returns the single-provider :class:`Prediction` only when the artifact is
    healthy, the ESPN game id matches, and the teams AND start time validate
    (30 min tolerance).  Every rejection records a ``CoverageGap`` (appended
    to ``gaps`` when the caller supplies a list) with the precise reason:
    ``PROVIDER_ERROR`` for an invalid/unhealthy artifact or a
    self-contradictory entry, ``STALE`` for a freshness failure, ``MISSING``
    for an absent game or an identity mismatch.  Never raises for expected
    data problems, never guesses, never blends providers.
    """

    def _gap(reason: CoverageState, detail: str) -> None:
        if gaps is not None:
            gaps.append(
                CoverageGap(
                    game_id=game.game_id,
                    required_field="prediction",
                    reason=reason,
                    provider=PROVIDER,
                    detail=detail,
                )
            )

    try:
        artifact = load_forecasts(fetch)
    except ValueError as exc:
        _gap(CoverageState.PROVIDER_ERROR, f"hardwood artifact invalid: {exc}")
        return None

    state = freshness_state(artifact, now_utc=now_utc, freshness_hours=freshness_hours)
    if state is not CoverageState.OK:
        _gap(
            state,
            f"hardwood artifact not usable: freshness_state={state.value} "
            f"(ingestion_status={artifact.ingestion_status!r}, "
            f"generated_at={artifact.generated_at!r}, "
            f"freshness_hours={freshness_hours:g})",
        )
        return None

    # game_id first: the artifact's game_id is the ESPN event id.
    espn_id = str(game.provider_ids.get("espn") or game.game_id)
    entry = artifact.games_by_id.get(espn_id)
    if entry is None:
        _gap(
            CoverageState.MISSING,
            f"no hardwood forecast for game_id {espn_id!r} in artifact "
            f"({len(artifact.games_by_id)} games indexed)",
        )
        return None

    entry_home = str((entry.get("home") or {}).get("name") or "")
    entry_away = str((entry.get("away") or {}).get("name") or "")
    if (
        not entry_home
        or not entry_away
        or normalize_team_name(entry_home) != normalize_team_name(game.home_name)
        or normalize_team_name(entry_away) != normalize_team_name(game.away_name)
    ):
        _gap(
            CoverageState.MISSING,
            "team identity mismatch: artifact "
            f"home={entry_home!r} away={entry_away!r} vs game "
            f"home={game.home_name!r} away={game.away_name!r}; never a "
            "best-guess or swapped-side join",
        )
        return None

    entry_date = entry.get("date_utc")
    if not entry_date or not times_within(str(entry_date), game.scheduled_utc):
        _gap(
            CoverageState.MISSING,
            "date mismatch beyond tolerance: artifact "
            f"date_utc={entry_date!r} vs game scheduled_utc="
            f"{game.scheduled_utc!r} (tolerance {TIME_TOLERANCE}); "
            "never join across dates",
        )
        return None

    try:
        generated_iso = to_utc_iso(parse_utc(artifact.generated_at))
        prediction = Prediction(
            game_id=game.game_id,
            provider=PROVIDER,
            observed_utc=generated_iso,  # artifact generation IS the observation
            source_modified_utc=generated_iso,
            model=PROVIDER,
            model_version=artifact.model_version,
            produced_by=PRODUCED_BY,
            # Stored as-is: probabilities are 0..1 in the artifact, exp_margin
            # is home-minus-away.  A contradiction with the expected scores is
            # caught by validate_orientation below (quarantine, not repair).
            prob_home=_num(entry.get("p_home")),
            prob_away=_num(entry.get("p_away")),
            expected_home_score=_num(entry.get("exp_home_score")),
            expected_away_score=_num(entry.get("exp_away_score")),
            expected_total=_num(entry.get("exp_total")),
            expected_margin=_num(entry.get("exp_margin")),
            source_revision=artifact.artifact_sha256,
            pre_tip=True,
        )
        prediction.validate_orientation()
    except (TypeError, ValueError, AttributeError, KeyError) as exc:
        _gap(
            CoverageState.PROVIDER_ERROR,
            f"hardwood forecast for {espn_id!r} failed validation: "
            f"{type(exc).__name__}: {exc}",
        )
        return None

    return prediction
