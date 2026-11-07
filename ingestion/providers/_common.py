"""Shared provider helpers: fetch function type, fixture loading, coverage."""

from __future__ import annotations

import json
import os
from typing import Any, Callable, Dict, List, Optional

FetchFn = Callable[[str], Any]

FIXTURES_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    "tests",
    "fixtures",
)


class IdentityError(ValueError):
    """Provider data could not be reconciled to canonical identity.

    Callers quarantine the record; they must NOT guess or swap home/away.
    """


def fixture_fetch(*paths: str) -> FetchFn:
    """Build a test FetchFn serving JSON files from ``tests/fixtures``.

    ``fixture_fetch(("espn", "scoreboard_20261106.json"))`` — a single tuple
    serves one URL independent of the requested URL. Alternatively pass
    ``{url_substring: relative_path}`` mappings via :func:`fixture_fetch_map`.
    """

    if len(paths) == 1 and isinstance(paths[0], tuple):
        rel = os.path.join(*paths[0])
    elif len(paths) == 1 and isinstance(paths[0], str):
        rel = paths[0]
    else:  # pragma: no cover - misuse guard
        raise TypeError("fixture_fetch expects one tuple path")

    def _fetch(url: str) -> Any:
        full = os.path.join(FIXTURES_DIR, rel)
        with open(full) as f:
            return json.load(f)

    return _fetch


def fixture_fetch_map(mapping: Dict[str, str]) -> FetchFn:
    """Build a test FetchFn that picks the fixture by URL substring."""

    def _fetch(url: str) -> Any:
        for needle, rel in mapping.items():
            if needle in url:
                full = os.path.join(FIXTURES_DIR, rel)
                with open(full) as f:
                    return json.load(f)
        raise AssertionError(f"no fixture mapped for URL: {url}")

    return _fetch


def bound_client_quota(fetch: FetchFn) -> Dict[str, str]:
    """Read quota headers captured by a bound ``HttpClient.get_json``.

    Fixture callables return no headers; production bound methods expose the
    client as ``__self__``. Explicit response envelopes still take precedence
    in each provider's unwrap helper.
    """
    client = getattr(fetch, "__self__", None)
    return dict(getattr(client, "last_quota", {}) or {})


class CoverageAccumulator:
    """Collects complete/total coverage counts and gaps per field."""

    def __init__(self) -> None:
        self.counts: Dict[str, Dict[str, int]] = {}
        self.gaps: List[Any] = []

    def observe(
        self,
        field: str,
        complete: bool,
        detail: str = "",
        game_id: Optional[str] = None,
        provider: Optional[str] = None,
        state: Optional[Any] = None,
    ) -> bool:
        slot = self.counts.setdefault(field, {"complete": 0, "total": 0})
        slot["total"] += 1
        if complete:
            slot["complete"] += 1
        else:
            from ..models import CoverageGap, CoverageState

            self.gaps.append(
                CoverageGap(
                    game_id=game_id,
                    required_field=field,
                    reason=state or CoverageState.MISSING,
                    provider=provider,
                    detail=detail,
                )
            )
        return complete

    def summary(self) -> Dict[str, Dict[str, int]]:
        return {k: dict(v) for k, v in self.counts.items()}
