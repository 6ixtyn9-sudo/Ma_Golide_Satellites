"""Deterministic stress: 10,000 synthetic chronological observations through
the prequential shadow evaluation. Fixed seed; structural invariants only (no
statistical thresholds); determinism re-checked on a rebuilt dataset; runtime
and peak memory are reported, not asserted."""

import json
import random
import time
import tracemalloc

from forecasting.shadow import RecencyConfig, prequential_shadow_evaluation
from forecasting.types import Forecast

SEED = 20261113
ROWS = 10_000

# The calibration epoch is 2026-11-10; the synthetic clock starts there and
# advances hourly, so every observation sits inside the calibratable window.
START_TIP = "2026-11-10T00:00:00Z"


def _build(count: int) -> list:
    rng = random.Random(SEED)
    from datetime import datetime, timedelta, timezone

    start = datetime(2026, 11, 10, tzinfo=timezone.utc)
    rows = []
    for index in range(count):
        tip = start + timedelta(hours=index)
        league = "NBA" if index % 2 else "NCAA"
        source = "provider:espn" if index % 3 else "empirical:synthetic"
        probability = round(rng.uniform(0.55, 0.92), 3)
        # Distribution drift in the final 40%: outcomes become more favorable
        # than the raw probabilities claim, so both calibrators must react.
        if index >= 0.6 * count:
            win_probability = min(0.99, probability + 0.06)
        else:
            win_probability = max(0.01, probability - 0.04)
        outcome = "win" if rng.random() < win_probability else "loss"
        rows.append(
            Forecast(
                game_id=f"s{index}",
                league=league,
                scheduled_utc=tip.strftime("%Y-%m-%dT%H:%M:%SZ"),
                home_name="Home",
                away_name="Away",
                market="moneyline",
                period="full_game",
                selection="home",
                raw_probability=probability,
                alternative_probabilities={"home": probability, "away": 1 - probability},
                source=source,
                outcome=outcome,
            )
        )
    return rows


def test_ten_thousand_row_prequential_stress():
    rows = _build(ROWS)
    tracemalloc.start()
    started = time.perf_counter()
    summary = prequential_shadow_evaluation(rows, RecencyConfig())
    duration = time.perf_counter() - started
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    assert summary["rows_evaluated"] == ROWS
    assert summary["rows_ungraded_or_ineligible_skipped"] == 0
    # Structural invariants only — no statistical or timing thresholds.
    for row in rows:
        shadow = row.shadow_calibration
        assert shadow["status"] in {"no_history", "provisional", "sufficient"}
        assert RecencyConfig().probability_floor <= shadow["probability"] <= RecencyConfig().probability_ceiling
        assert shadow["ess"] <= shadow["n"] + 1e-9
        assert shadow["champion_status"] in {"no_history", "provisional", "sufficient"}
    assert sum(summary["challenger"]["coverage"].values()) == ROWS
    assert sum(summary["champion"]["coverage"].values()) == ROWS
    assert summary["challenger"]["mean_effective_sample_size"] is not None

    # Determinism: a fresh rebuild of the same seeded dataset yields an
    # identical aggregate (and the same per-row shadows).
    rebuilt = _build(ROWS)
    summary_again = prequential_shadow_evaluation(rebuilt, RecencyConfig())
    assert json.dumps(summary, sort_keys=True) == json.dumps(summary_again, sort_keys=True)
    for left, right in zip(rows, rebuilt):
        assert left.shadow_calibration == right.shadow_calibration

    print(
        f"\nstress: rows={ROWS} duration={duration:.3f}s peak_mem={peak / 1e6:.1f}MB "
        f"brier raw={summary['raw_metrics']['brier_score']} "
        f"champion={summary['champion']['brier_score']} "
        f"challenger={summary['challenger']['brier_score']} "
        f"mean_ess={summary['challenger']['mean_effective_sample_size']} "
        f"mean_n={summary['challenger']['mean_raw_n']}"
    )
