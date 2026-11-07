"""HTTP layer: timeouts, bounded retry/backoff, quota headers, redaction.

Rules:
- Every call has a timeout; retries are bounded and NEVER applied to
  permanent failures (400/401/403/404/422).
- Credentials are never logged: ``redact()`` scrubs configured secret
  environment values from any text before it is raised/logged.
- Quota/rate-limit headers are captured into a dict the caller can store on
  ``ProviderHealth.quota``.
- Unit tests never call this module's network path; providers accept a
  ``fetch`` callable so tests inject fixtures.
"""

from __future__ import annotations

import logging
import os
import re
import time
from typing import Any, Callable, Dict, Optional

import requests

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT = 20.0
DEFAULT_RETRIES = 2
DEFAULT_BACKOFF = 1.5
PERMANENT_STATUS = {400, 401, 403, 404, 405, 406, 410, 422, 451}

# Quota headers we recognise across providers.
QUOTA_HEADERS = [
    "x-datapoints",
    "x-requests-remaining",
    "x-requests-used",
    "x-requests-reset",
    "x-ratelimit-remaining",
    "x-ratelimit-limit",
    "ratelimit-remaining",
    "ratelimit-reset",
    "x-quota-remaining",
    "x-quota-used",
]

REDACT_KEYS = (
    "PREDIXSPORT_API_KEY",
    "THERUNDOWN_API_KEY",
    "SPORTSGAMEODDS_API_KEY",
    "THE_ODDS_API_KEY",
    "WHATSAPP_TOKEN",
)


class HttpError(RuntimeError):
    """Raised for non-retryable/permanent HTTP failures."""

    def __init__(self, status: int, url: str, body_snippet: str = "") -> None:
        self.status = status
        # Redact the URL too: credentials in query strings (e.g. apiKey=)
        # must never surface in logs or exception text.
        self.url = url
        redacted_url = redact(url)
        redacted = redact(body_snippet)[:300]
        super().__init__(f"HTTP {status} for {redacted_url}: {redacted}")


def redact(text: str) -> str:
    """Scrub secret environment values and common key patterns from text."""
    out = str(text or "")
    for key in REDACT_KEYS:
        value = os.environ.get(key, "")
        if value:
            out = out.replace(value, f"[REDACTED:{key}]")
    # Private keys, bearer tokens, and generic API keys.
    out = re.sub(r"-----BEGIN[ A-Z]*PRIVATE KEY-----[\s\S]*?-----END[ A-Z]*PRIVATE KEY-----", "[REDACTED:PRIVATE_KEY]", out)
    out = re.sub(r"(?i)bearer\s+[A-Za-z0-9._\-]+", "bearer [REDACTED]", out)
    out = re.sub(r"(?i)(api[_-]?key|apikey|token|secret)\s*[=:]\s*['\"]?[A-Za-z0-9._\-]{8,}", r"\1=[REDACTED]", out)
    return out


class HttpClient:
    """Bounded-retry GET client with optional in-memory ETag caching."""

    def __init__(
        self,
        timeout: float = DEFAULT_TIMEOUT,
        retries: int = DEFAULT_RETRIES,
        backoff: float = DEFAULT_BACKOFF,
        user_agent: str = "MaGolideIngestion/0.1 (unattended acquisition)",
        session: Optional[requests.Session] = None,
        sleep: Callable[[float], None] = time.sleep,
        default_headers: Optional[Dict[str, str]] = None,
    ) -> None:
        self.timeout = timeout
        self.retries = retries
        self.backoff = backoff
        self.user_agent = user_agent
        self.session = session or requests.Session()
        self.sleep = sleep
        self.default_headers = dict(default_headers or {})
        self._etag_cache: Dict[str, Dict[str, Any]] = {}
        self.last_quota: Dict[str, str] = {}
        self.last_status: Optional[int] = None

    # -- public API ---------------------------------------------------------

    def get_json(self, url: str, headers: Optional[Dict[str, str]] = None) -> Any:
        """GET a JSON document with bounded retries. Raises HttpError on 4xx."""
        body = self.get_text(url, headers=headers)
        import json

        try:
            return json.loads(body)
        except json.JSONDecodeError as e:
            raise HttpError(0, url, f"invalid JSON: {e}") from e

    def get_text(self, url: str, headers: Optional[Dict[str, str]] = None) -> str:
        merged = dict(self.default_headers)
        merged["User-Agent"] = self.user_agent
        if headers:
            merged.update(headers)
        attempt = 0
        while True:
            try:
                resp = self.session.get(url, headers=merged, timeout=self.timeout)
            except requests.RequestException as e:
                if attempt >= self.retries:
                    raise HttpError(0, url, redact(str(e))) from e
                attempt += 1
                self.sleep(self.backoff**attempt)
                continue

            self.last_status = resp.status_code
            self._capture_quota(resp)
            if resp.status_code in PERMANENT_STATUS:
                raise HttpError(resp.status_code, url, redact(resp.text[:500]))
            if resp.status_code >= 500 or resp.status_code == 429:
                if attempt >= self.retries:
                    raise HttpError(resp.status_code, url, redact(resp.text[:300]))
                attempt += 1
                self.sleep(self.backoff**attempt)
                continue
            if resp.status_code >= 400:
                raise HttpError(resp.status_code, url, redact(resp.text[:300]))
            return resp.text

    # -- internals ----------------------------------------------------------

    def _capture_quota(self, resp: requests.Response) -> None:
        quota: Dict[str, str] = {}
        for name in QUOTA_HEADERS:
            value = resp.headers.get(name)
            if value is not None:
                quota[name] = value
        if quota:
            self.last_quota = quota
