"""Prediction selection, fallback coherence, and pre-tip freezing.

Rules (master prompt section 8, binding):

- G:J (``Prob %``, ``Pred``, ``Pred Score``, ``Avg``) is ONE coherent
  provider bundle: a complete provider wins whole; fields are never mixed
  across providers to synthesise a forecast.
- NBA default order: fresh complete ESPN BPI, then fresh matching Hardwood.
  Predixsport joins the order only after authenticated verification.
- A partial provider may fill only what it genuinely supports (e.g.
  probabilities), with explicit per-field provenance and no invented values.
- Hardwood is used only when its ingestion is healthy, fresh, identity-
  verified, and the game is still upcoming (enforced in its adapter).
- Before tip, the latest observation under the refresh policy wins.  At or
  after tip, the latest valid PRE-TIP observation is frozen and never
  replaced by a later value.
- Shadow predictions (non-selected providers) are preserved for independent
  evaluation; providers are never averaged or blended.
- Rounded display scores may tie or disagree with the probability favorite;
  that is preserved, never "repaired".
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .identity import parse_utc
from .models import BookSnapshot, CoverageGap, CoverageState, Game, Prediction, pred_from_probabilities

# Provider priority for complete-bundle selection (NBA). "espn" is the
# provider label used by providers/espn.py for BPI predictions.
DEFAULT_PROVIDER_ORDER = ("espn", "hardwood", "predixsport")

# A "complete" bundle needs the whole G:J span.
COMPLETE_FIELDS = ("prob_home", "prob_away", "expected_home_score",
                   "expected_away_score", "expected_total")


def is_complete(pred: Prediction) -> bool:
    return all(getattr(pred, f) is not None for f in COMPLETE_FIELDS)


# Tolerated provider/caller clock skew. `now` snapshots and provider
# timestamps (e.g. ESPN lastModified, which has been observed running tens
# of minutes ahead) must not disqualify a just-fetched observation, but an
# observation meaningfully in the future is a bug and stays rejected.
CLOCK_SKEW_TOLERANCE_SECONDS = 300.0


def is_fresh(pred: Prediction, now_utc: str, freshness_hours: float) -> bool:
    """Observed before ``now`` (within skew tolerance) and inside the window."""
    from datetime import timedelta

    now = parse_utc(now_utc)
    obs = parse_utc(pred.observed_utc)
    if obs > now + timedelta(seconds=CLOCK_SKEW_TOLERANCE_SECONDS):
        return False  # future observation is a bug, treat as not fresh
    return (now - obs) <= timedelta(hours=freshness_hours)


@dataclass
class SelectionResult:
    selected: Optional[Prediction]
    provider: Optional[str]
    reason: str
    gaps: List[CoverageGap] = field(default_factory=list)
    shadows: List[Prediction] = field(default_factory=list)
    partial_fields: Tuple[str, ...] = ()

    @property
    def is_partial(self) -> bool:
        return self.selected is not None and bool(self.partial_fields)


def _valid_candidate(
    pred: Prediction, game: Game, now_utc: str, freshness_hours: float
) -> bool:
    """Identity + freshness + pre-tip guards for one candidate."""
    if pred.game_id != game.game_id:
        return False
    if not is_fresh(pred, now_utc, freshness_hours):
        return False
    return True


def select_prediction(
    candidates: List[Prediction],
    game: Game,
    now_utc: str,
    freshness_hours: float = 12.0,
    provider_order: Tuple[str, ...] = DEFAULT_PROVIDER_ORDER,
) -> SelectionResult:
    """Pick ONE coherent bundle; never mix providers.

    Selection order:
    1. Complete bundles, in ``provider_order`` (freshest per provider wins
       inside a provider).
    2. Otherwise the highest-priority partial bundle (probabilities only,
       e.g. WNBA/NCAA BPI) — its supported fields carry explicit partial
       provenance.
    """
    valid = [c for c in candidates if _valid_candidate(c, game, now_utc, freshness_hours)]
    grouped: Dict[str, List[Prediction]] = {}
    for candidate in valid:
        grouped.setdefault(candidate.provider, []).append(candidate)

    # Prefer the freshest COMPLETE observation within each provider. A newer
    # partial payload must not displace a still-fresh complete bundle merely
    # because it arrived later; if no complete observation exists, use the
    # freshest partial one with explicit missing-field provenance.
    by_provider: Dict[str, Prediction] = {}
    for provider, candidates in grouped.items():
        complete = [candidate for candidate in candidates if is_complete(candidate)]
        pool = complete or candidates
        by_provider[provider] = max(pool, key=lambda candidate: candidate.observed_utc)

    ordered = [by_provider[p] for p in provider_order if p in by_provider]
    ordered += [c for p, c in by_provider.items() if p not in provider_order]

    shadows: List[Prediction] = []
    for pred in ordered:
        if is_complete(pred):
            shadows.extend(c for c in ordered if c is not pred)
            return SelectionResult(
                selected=pred, provider=pred.provider,
                reason="complete_bundle", shadows=shadows,
            )

    gaps: List[CoverageGap] = []
    if not ordered:
        gaps.append(CoverageGap(
            game_id=game.game_id, required_field="pregame_prediction",
            reason=CoverageState.MISSING, detail="no fresh provider prediction",
        ))
        return SelectionResult(None, None, "no_fresh_candidate", gaps=gaps)

    partial = ordered[0]
    missing = tuple(f for f in COMPLETE_FIELDS if getattr(partial, f) is None)
    gaps.append(CoverageGap(
        game_id=game.game_id, required_field="prediction_scores",
        reason=CoverageState.MISSING, provider=partial.provider,
        detail=f"partial bundle missing: {', '.join(missing) or 'none'}",
    ))
    shadows.extend(c for c in ordered if c is not partial)
    return SelectionResult(
        selected=partial, provider=partial.provider,
        reason="partial_bundle_probabilities_only",
        gaps=gaps, shadows=shadows, partial_fields=missing,
    )


# ---------------------------------------------------------------------------
# Freezing
# ---------------------------------------------------------------------------


def latest_pre_tip(
    predictions: List[Prediction], tip_utc: str
) -> Optional[Prediction]:
    """Latest observation strictly usable as a pre-tip value.

    Accepts observations up to tip time (observed_utc <= tip); a value first
    seen after tip is NOT a pre-tip observation and never qualifies.
    """
    tip = parse_utc(tip_utc)
    usable = [
        p for p in predictions
        if p.pre_tip and parse_utc(p.observed_utc) <= tip
    ]
    if not usable:
        return None
    usable.sort(key=lambda p: p.observed_utc)
    return usable[-1]


def freeze_game_predictions(store, game: Game, tip_utc: str) -> Dict[str, str]:
    """Freeze the latest pre-tip prediction per provider; never overwrite.

    Post-tip observations (``pre_tip=False``, stamped by the pipeline at
    storage time) are ignored: a value first seen after tip is never a
    pre-tip prediction. Returns ``{provider: status}`` with status in
    ``frozen`` | ``already_frozen`` | ``no_pregame_prediction``.
    """
    from .store import FrozenWriteError

    statuses: Dict[str, str] = {}
    by_provider: Dict[str, Prediction] = {}
    for pred in store.get_predictions(game.game_id):
        if not pred.pre_tip:
            continue
        prev = by_provider.get(pred.provider)
        if prev is None or pred.observed_utc > prev.observed_utc:
            by_provider[pred.provider] = pred
    for provider, pred in sorted(by_provider.items()):
        chosen = latest_pre_tip([pred], tip_utc)
        if chosen is None:
            statuses[provider] = "no_pregame_prediction"
            continue
        if store.is_prediction_frozen(game.game_id, provider):
            statuses[provider] = "already_frozen"
            continue
        try:
            store.upsert_prediction(chosen, frozen=True)
            statuses[provider] = "frozen"
        except FrozenWriteError:
            statuses[provider] = "already_frozen"
    return statuses


def resolve_frozen_bundle(
    store, game_id: str, provider_order: Tuple[str, ...] = DEFAULT_PROVIDER_ORDER
) -> Tuple[Optional[Prediction], bool]:
    """The frozen prediction to render on a result row.

    Returns ``(prediction, was_frozen)``. Only PRE-TIP observations qualify
    (frozen or not): a value first seen after tip never renders. Prefers
    frozen providers in the selection order; falls back to the latest
    pre-tip value of any provider when nothing is frozen (unattended runs
    without a freeze step must still render the correct pre-tip snapshot).
    """
    predictions = [p for p in store.get_predictions(game_id) if p.pre_tip]
    frozen = [p for p in predictions if store.is_prediction_frozen(game_id, p.provider)]
    pool = frozen or predictions
    if not pool:
        return None, False
    by_provider = {p.provider: p for p in pool}
    for provider in provider_order:
        pred = by_provider.get(provider)
        if pred is not None and is_complete(pred):
            return pred, bool(frozen)
    for provider in provider_order:
        if provider in by_provider:
            return by_provider[provider], bool(frozen)
    pool.sort(key=lambda p: p.observed_utc)
    return pool[-1], bool(frozen)


# ---------------------------------------------------------------------------
# Odds for the predicted side
# ---------------------------------------------------------------------------


def odds_for_prediction(
    snapshot: BookSnapshot, prediction: Optional[Prediction]
) -> Optional[float]:
    """Decimal moneyline of the team selected by ``Pred``, same snapshot.

    If the selected prediction changes, the caller re-runs this with the new
    selection: the price always follows the predicted side of the SAME book
    snapshot (section 6 ``Odds`` semantics).  Never a live/post-game price.
    """
    if prediction is None or snapshot.market.value != "moneyline":
        return None
    side = pred_from_probabilities(prediction.prob_home, prediction.prob_away)
    if side is None:
        return None
    return snapshot.moneyline_for_side(side)
