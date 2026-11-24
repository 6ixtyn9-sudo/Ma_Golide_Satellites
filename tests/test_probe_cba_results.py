"""The CBA round probe: bounded, quiet about failure, and honest about what it saw.

Run 2026-11-24T14:02Z: the documented WCBA results URL answered HTTP 500 in 179 B
with ``param name: roundId value not exists``. The error names the missing
parameter and says nothing about its values, so the probe enumerates a bounded
range and records every answer. These tests pin the parts that decide what the
evidence means, without touching the network.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_SPEC = importlib.util.spec_from_file_location(
    "probe_cba_results",
    Path(__file__).resolve().parents[1] / "scripts" / "probe_cba_results.py",
)
probe = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(probe)


def test_the_url_carries_the_parameter_the_host_asked_for():
    url = probe.rounds_url(7)
    assert url.startswith("https://www.cba.net.cn/datahub/cbamatch/games/matchschedules?")
    assert "competitionId=56" in url
    assert "seasonId=2286" in url
    assert url.endswith("&roundId=7")


def test_rounds_are_parsed_bounded_and_ordered():
    assert probe.parse_rounds("1-5") == [1, 2, 3, 4, 5]
    assert probe.parse_rounds("3,1,3") == [1, 3]
    assert probe.parse_rounds(" 2 - 4 ") == [2, 3, 4]
    with pytest.raises(ValueError, match="no rounds"):
        probe.parse_rounds(" , ")


def test_a_range_wider_than_the_budget_is_refused_not_quietly_trimmed():
    with pytest.raises(ValueError, match="bounded budget"):
        probe.parse_rounds(f"1-{probe.REQUEST_BUDGET + 1}")


def test_the_observed_500_is_read_as_an_error_not_as_an_empty_slate():
    summary = probe.summarise(
        '{"timestamp":"2026-11-24T14:02:50.016+00:00","path":"/cbamatch/games/matchschedules",'
        '"status":500,"error":"Internal Server Error",'
        '"message":"param name: roundId value not exists"}'
    )

    assert summary["looks_like_json"] is True
    assert summary["game_count"] == 0
    assert summary["error"] == "param name: roundId value not exists"
    # The 179 bytes the host actually served, not a rounded guess.
    assert summary["bytes"] == 179


def test_a_body_with_games_is_counted_without_assuming_one_key():
    body = '{"data":{"list":[{"gameId":1,"homeScore":88},{"gameId":2,"homeScore":91}]}}'
    assert probe.summarise(body)["game_count"] == 2


def test_a_non_json_body_is_recorded_without_inventing_a_shape():
    summary = probe.summarise("<html>页面找不到</html>")

    assert summary["looks_like_json"] is False
    assert summary["game_count"] == 0
    assert summary["error"] == ""
    assert summary["sha256"]
