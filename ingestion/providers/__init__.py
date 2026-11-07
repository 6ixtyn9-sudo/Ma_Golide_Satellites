"""Provider adapters. See docs/PROVIDERS.md for coverage and constraints."""

from __future__ import annotations

from ._common import CoverageAccumulator, FetchFn, IdentityError, fixture_fetch, fixture_fetch_map

__all__ = [
    "CoverageAccumulator",
    "FetchFn",
    "IdentityError",
    "fixture_fetch",
    "fixture_fetch_map",
]
