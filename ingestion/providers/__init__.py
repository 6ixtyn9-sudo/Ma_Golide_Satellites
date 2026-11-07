"""Provider adapters. See ingestion/PROVIDER_GUIDE.md for the contract."""

from __future__ import annotations

from ._common import CoverageAccumulator, FetchFn, IdentityError, fixture_fetch, fixture_fetch_map

__all__ = [
    "CoverageAccumulator",
    "FetchFn",
    "IdentityError",
    "fixture_fetch",
    "fixture_fetch_map",
]
