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

import gzip
import zlib

import logging
import os
import re
import time
from typing import Any, Callable, Dict, Optional

import requests

logger = logging.getLogger(__name__)

# Process-wide transport meter. Providers that receive no ``fetch`` build their
# own ``HttpClient`` via ``default_fetch()``, which the pipeline's injected
# meter never sees; a lane could therefore read thirty pages and still report
# ``network_requests: 0``. Count every attempt here, at the one place a request
# is actually issued, so "we did not ask" can never be reported for a lane that
# did ask.
_TRANSPORT_METER: Dict[str, int] = {"attempts": 0, "responses": 0, "errors": 0}


def transport_meter() -> Dict[str, int]:
    """HTTP attempts seen by every HttpClient in this process."""

    return dict(_TRANSPORT_METER)


def reset_transport_meter() -> None:
    for key in _TRANSPORT_METER:
        _TRANSPORT_METER[key] = 0

DEFAULT_TIMEOUT = 20.0
DEFAULT_RETRIES = 2
DEFAULT_BACKOFF = 1.5
PERMANENT_STATUS = {400, 401, 403, 404, 405, 406, 410, 422, 451}

#: A host that answers ``Retry-After`` is telling us when to come back. Honouring
#: it is the polite reading of a rate limit; ignoring it and retrying on our own
#: schedule is the impolite one. The wait stays bounded so a lane cannot sleep
#: through its own tick -- a header longer than this is a "come back later", and
#: the caller degrades to the stored slate instead of camping on the run.
MAX_RETRY_AFTER_SECONDS = 30.0


def retry_after_seconds(value: Any) -> Optional[float]:
    """Parse a ``Retry-After`` header (delta-seconds or HTTP-date) to seconds.

    Returns ``None`` for anything unparseable so the caller falls back to its own
    bounded backoff rather than trusting a malformed header.
    """

    text = str(value or "").strip()
    if not text:
        return None
    try:
        return max(0.0, float(text))
    except ValueError:
        pass
    try:
        from datetime import datetime, timezone
        from email.utils import parsedate_to_datetime

        when = parsedate_to_datetime(text)
    except (TypeError, ValueError):
        return None
    if when is None:
        return None
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    return max(0.0, (when - datetime.now(timezone.utc)).total_seconds())

# Quota headers we recognise across providers.
QUOTA_HEADERS = [
    "x-datapoints",
    "x-requests-remaining",
    "x-requests-used",
    "x-requests-last",
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
    out = re.sub(
        r"-----BEGIN[ A-Z]*PRIVATE KEY-----[\s\S]*?-----END[ A-Z]*PRIVATE KEY-----",
        "[REDACTED:PRIVATE_KEY]",
        out,
    )
    out = re.sub(r"(?i)bearer\s+[A-Za-z0-9._\-]+", "bearer [REDACTED]", out)
    out = re.sub(
        r"(?i)(api[_-]?key|apikey|token|secret)\s*[=:]\s*['\"]?[A-Za-z0-9._\-]{8,}",
        r"\1=[REDACTED]",
        out,
    )
    return out


GZIP_MAGIC = b"\x1f\x8b"


def _decoded_body(resp: Any) -> str:
    """Return the response body as text, decompressing an undeclared gzip body.

    B.League's S3 context snapshots are served gzip-compressed while declaring
    ``content-type: application/json`` and no ``Content-Encoding``, so
    ``requests`` leaves the bytes alone and ``resp.text`` is replacement
    characters. Parsing that fails as if the JSON were corrupt. Verified from an
    external probe on 2026-11-23: body magic ``1f 8b``, 543,231 bytes once
    decompressed, full ``PlayByPlays``/``Game`` payload.

    A body whose magic bytes say gzip is decompressed regardless of what the
    headers claim; anything else keeps the normal ``resp.text`` path, so a
    correctly declared response is untouched.
    """

    raw = getattr(resp, "content", None)
    if isinstance(raw, (bytes, bytearray)) and bytes(raw[:2]) == GZIP_MAGIC:
        try:
            return gzip.decompress(bytes(raw)).decode("utf-8", "replace")
        except (OSError, EOFError, zlib.error):
            # Not a usable gzip stream despite the magic -- a truncated body, a
            # bad CRC, or (the case an external audit found) a valid gzip header
            # wrapped around a corrupt deflate stream, where gzip.decompress
            # raises zlib.error. zlib.error is a zlib-module exception and is
            # neither OSError nor EOFError, so catching only those let a raw
            # zlib.error escape get_text() instead of falling back to resp.text
            # -- an exception type this module's contract says it never raises.
            pass
    return resp.text


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
        # Header values must be latin-1, and ``requests``/``urllib3`` raise a
        # bare UnicodeEncodeError (not RequestException) for anything else —
        # a traceback instead of a fail-closed provider error. Name the header
        # and the code point; never echo the value, which may be a credential.
        for name, value in merged.items():
            try:
                str(value).encode("latin-1")
            except UnicodeEncodeError as exc:
                raise HttpError(
                    0,
                    url,
                    f"header {name!r} is not latin-1 encodable at index "
                    f"{exc.start} (U+{ord(str(value)[exc.start]):04X}); "
                    "check the credential value for pasted placeholders",
                ) from exc
        attempt = 0
        while True:
            _TRANSPORT_METER["attempts"] += 1
            try:
                resp = self.session.get(url, headers=merged, timeout=self.timeout)
            except requests.RequestException as e:
                _TRANSPORT_METER["errors"] += 1
                if attempt >= self.retries:
                    raise HttpError(0, url, redact(str(e))) from e
                attempt += 1
                self.sleep(self.backoff**attempt)
                continue

            _TRANSPORT_METER["responses"] += 1
            self.last_status = resp.status_code
            self._capture_quota(resp)
            if resp.status_code in PERMANENT_STATUS:
                raise HttpError(resp.status_code, url, redact(resp.text[:500]))
            if resp.status_code >= 500 or resp.status_code == 429:
                if attempt >= self.retries:
                    raise HttpError(resp.status_code, url, redact(resp.text[:300]))
                attempt += 1
                self.sleep(self._retry_delay(resp, attempt))
                continue
            if resp.status_code >= 400:
                raise HttpError(resp.status_code, url, redact(resp.text[:300]))
            return _decoded_body(resp)

    # -- internals ----------------------------------------------------------

    def _retry_delay(self, resp: requests.Response, attempt: int) -> float:
        """Bounded wait before retry: the host's ``Retry-After``, else backoff.

        ``Retry-After`` is evidence about the host, not a preference of ours, so
        it wins when present and parseable -- clamped to
        :data:`MAX_RETRY_AFTER_SECONDS` because a longer one means "not now",
        which the caller handles by degrading rather than sleeping.
        """

        promised = retry_after_seconds(resp.headers.get("Retry-After"))
        if promised is not None:
            return min(promised, MAX_RETRY_AFTER_SECONDS)
        return self.backoff**attempt

    def _capture_quota(self, resp: requests.Response) -> None:
        quota: Dict[str, str] = {}
        for name in QUOTA_HEADERS:
            value = resp.headers.get(name)
            if value is not None:
                quota[name] = value
        if quota:
            self.last_quota = quota
