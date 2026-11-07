"""Bounded public-payload acquisition with durable cache and safe telemetry.

This is the basketball-neutral part of the acquisition discipline used by the
rolling website captures in Edge Factory: request budgets, cooldowns, durable
cache entries, and attempt receipts. Provider adapters still own payload
semantics and entity reconciliation.

Only callers that explicitly pass ``use_cache=True`` read/write the durable
cache. Current predictions and book observations therefore always receive a
fresh network observation timestamp instead of accidentally re-dating a cached
payload. Historical scoreboards are the intended cache users.

Cache filenames contain a SHA-256 URL digest. URLs, query strings, response
headers, and exception bodies are never written to telemetry or cache metadata;
this prevents provider credentials in query parameters from leaking to disk.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import tempfile
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, Optional
from urllib.parse import urlsplit

from .http import redact
from .identity import now_utc_iso, parse_utc, to_utc_iso


class AcquisitionBudgetExceeded(RuntimeError):
    """The per-run network-request budget was exhausted."""


class CachedProviderFailure(RuntimeError):
    """A recent failure is inside its explicit negative-cache cooldown."""


def _utc(value: datetime) -> str:
    return to_utc_iso(value.astimezone(timezone.utc))


def _endpoint(url: str) -> str:
    """Return a query-free endpoint label safe for receipts."""

    parsed = urlsplit(url)
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}"


class AcquisitionClient:
    """Wrap a JSON fetcher with a bounded network budget and optional cache."""

    def __init__(
        self,
        fetch: Callable[[str], Any],
        *,
        cache_dir: Optional[str] = None,
        request_budget: int = 1200,
        failure_cooldown_seconds: int = 900,
        now_fn: Callable[[], str] = now_utc_iso,
    ) -> None:
        if isinstance(request_budget, bool) or int(request_budget) < 1:
            raise ValueError("request_budget must be a positive integer")
        self.fetch = fetch
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self.request_budget = int(request_budget)
        self.failure_cooldown_seconds = max(0, int(failure_cooldown_seconds))
        self.now_fn = now_fn
        self.network_requests = 0
        self.fixture_requests = 0
        self.cache_hits = 0
        self.cache_writes = 0
        self.negative_cache_hits = 0
        self.last_observed_utc: Optional[str] = None
        self.attempts = []

    def __call__(self, url: str) -> Any:
        """Fetch without a durable read, preserving live observation semantics."""

        return self.get_json(url, use_cache=False)

    def get_json(
        self,
        url: str,
        *,
        use_cache: bool = True,
        ttl_seconds: int = 86400,
    ) -> Any:
        """Return JSON, optionally using a bounded durable positive/negative cache."""

        now = parse_utc(self.now_fn())
        key = hashlib.sha256(url.encode("utf-8")).hexdigest()
        if use_cache:
            cached = self._read(key)
            if cached:
                expires = parse_utc(cached["expires_utc"])
                if expires > now:
                    if cached["kind"] == "payload":
                        self.cache_hits += 1
                        self.last_observed_utc = cached["observed_utc"]
                        self._attempt(url, "cache_hit", cached["observed_utc"])
                        return cached["payload"]
                    self.negative_cache_hits += 1
                    self.last_observed_utc = cached["observed_utc"]
                    self._attempt(url, "negative_cache_hit", cached["observed_utc"])
                    raise CachedProviderFailure(
                        f"temporary acquisition cooldown for {_endpoint(url)}"
                    )
        if getattr(self.fetch, "magolide_acquisition_source", None) == "fixture":
            self.fixture_requests += 1
            observed = _utc(now)
            self.last_observed_utc = observed
            try:
                payload = self.fetch(url)
            except Exception as exc:
                self._attempt(
                    url, "fixture_error", observed, type(exc).__name__
                )
                raise
            self._attempt(url, "fixture_success", observed)
            return payload
        if self.network_requests >= self.request_budget:
            self._attempt(url, "budget_exceeded", _utc(now))
            raise AcquisitionBudgetExceeded(
                f"network request budget {self.request_budget} exhausted"
            )

        self.network_requests += 1
        observed = _utc(now)
        self.last_observed_utc = observed
        try:
            payload = self.fetch(url)
        except Exception as exc:
            # Only the redacted exception class is durable. Message text may
            # contain a provider URL/query and is deliberately not persisted.
            detail = type(exc).__name__
            self._attempt(url, "provider_error", observed, detail)
            if use_cache and self.failure_cooldown_seconds:
                self._write(
                    key,
                    {
                        "kind": "error",
                        "observed_utc": observed,
                        "expires_utc": _utc(
                            now + timedelta(seconds=self.failure_cooldown_seconds)
                        ),
                        "error_type": detail,
                    },
                )
            raise

        self._attempt(url, "network_success", observed)
        if use_cache:
            self._write(
                key,
                {
                    "kind": "payload",
                    "observed_utc": observed,
                    "expires_utc": _utc(now + timedelta(seconds=max(1, ttl_seconds))),
                    "payload": payload,
                },
            )
        return payload

    def telemetry(self) -> Dict[str, Any]:
        """Return bounded, secret-free counters and at most 100 recent attempts."""

        return {
            "schema_version": "magolide.acquisition.v1",
            "network_requests": self.network_requests,
            "fixture_requests": self.fixture_requests,
            "request_budget": self.request_budget,
            "cache_hits": self.cache_hits,
            "cache_writes": self.cache_writes,
            "negative_cache_hits": self.negative_cache_hits,
            "cache_configured": self.cache_dir is not None,
            "attempts_retained": min(len(self.attempts), 100),
            "attempts_total": len(self.attempts),
            "recent_attempts": self.attempts[-100:],
        }

    def _attempt(
        self, url: str, outcome: str, observed_utc: str, detail: str = ""
    ) -> None:
        self.attempts.append(
            {
                "endpoint": _endpoint(redact(url)),
                "outcome": outcome,
                "observed_utc": observed_utc,
                "detail": detail,
            }
        )

    def _path(self, key: str) -> Optional[Path]:
        return self.cache_dir / f"{key}.json" if self.cache_dir else None

    def _read(self, key: str) -> Optional[Dict[str, Any]]:
        path = self._path(key)
        if path is None or not path.exists():
            return None
        try:
            value = json.loads(path.read_text())
            if (
                value.get("schema_version") == "magolide.http-cache.v1"
                and value.get("url_sha256") == key
                and value.get("kind") in {"payload", "error"}
            ):
                parse_utc(value["observed_utc"])
                parse_utc(value["expires_utc"])
                return value
        except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
            return None
        return None

    def _write(self, key: str, value: Dict[str, Any]) -> None:
        path = self._path(key)
        if path is None:
            return
        document = {
            "schema_version": "magolide.http-cache.v1",
            "url_sha256": key,
            **value,
        }
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(
                "w", dir=path.parent, prefix=f".{key}.", delete=False
            ) as handle:
                temporary = Path(handle.name)
                json.dump(document, handle, sort_keys=True, separators=(",", ":"))
                handle.write("\n")
            os.replace(temporary, path)
            self.cache_writes += 1
        finally:
            if temporary is not None and temporary.exists():
                temporary.unlink()
