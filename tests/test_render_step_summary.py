"""The lane job summary must survive GitHub's 1024k step-summary cap.

Run 45385819819: NCAA_W's naive summary hit 1,207k (``operator`` alone was
519,765 bytes) and GitHub aborted the upload with ``$GITHUB_STEP_SUMMARY upload
aborted, supports content up to a size of 1024k, got 1207k`` -- the lane's
summary was silently lost from the run's UI record.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

_SPEC = importlib.util.spec_from_file_location(
    "render_step_summary",
    Path(__file__).resolve().parents[1] / "scripts" / "render_step_summary.py",
)
render_step_summary = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(render_step_summary)


def test_a_small_report_is_rendered_in_full():
    data = {
        "league": "NBA",
        "mode": "write",
        "counts": {"predictions": 75, "book_snapshots": 271},
        "operator": {"paper_only": True, "picks_text": "localdata/nba/picks_2026-11-24.txt"},
        "markets": {"forecast_count": 60, "walk_forward": {"graded_forecasts": 7230}, "accas": {}},
        "commands": {
            "pregame": {
                "book_provider_status": {"espn": {"status": "observed_pre_match_markets"}}
            }
        },
    }

    text = render_step_summary.render(data)

    # The heredoc this replaced titled the block with the lane name; keep it.
    assert text.startswith("## NBA")
    assert '"predictions": 75' in text
    assert '"forecast_count": 60' in text
    assert "omitted" not in text


def test_an_oversized_report_loses_sections_not_the_whole_summary():
    data = {
        "league": "NCAA_W",
        "mode": "write",
        "counts": {"predictions": 300},
        "status": "ok",
        # 519,765 bytes of operator audit in the real run; bigger here to be sure.
        "operator": {"rolling_audit": "x" * (render_step_summary.MAX_CHARS + 40_000)},
        "markets": {"forecast_count": 184, "walk_forward": {}, "accas": {}},
        "commands": {"pregame": {"book_provider_status": {}}},
    }

    text = render_step_summary.render(data)

    assert len(text) <= render_step_summary.MAX_CHARS
    assert "<omitted:" in text
    # The small, decision-bearing keys survive the trim.
    assert '"league": "NCAA_W"' in text
    assert '"status": "ok"' in text


def test_a_missing_report_is_said_not_rendered(capsys, monkeypatch, tmp_path):
    # Parity with the heredoc it replaced: no file -> plain sentence, exit 0.
    monkeypatch.chdir(tmp_path)
    assert render_step_summary.main() == 0
    assert "No valid daily report" in capsys.readouterr().out
