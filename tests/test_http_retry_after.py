"""A host that names its wait gets it honoured; a vague one gets our backoff.

BetExplorer's fixtures page answered ``HTTP 429`` three times in ~5s on run
45376480832 and the tick died. The retry schedule was ours alone (1.5s, 2.25s):
the transport never read ``Retry-After``, so when the host did say how long to
wait, the answer was ignored and the lane failed anyway.

The contract asserted here:

* ``Retry-After`` in delta-seconds is the wait;
* ``Retry-After`` as an HTTP-date is converted to the remaining wait;
* it is clamped -- a header longer than :data:`MAX_RETRY_AFTER_SECONDS` means
  "not now", and the caller degrades rather than camping on the run;
* no header (or a malformed one) keeps the existing bounded exponential wait;
* the wait never extends the retry budget: 429 still raises once retries run out.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from email.utils import format_datetime

import pytest

from ingestion import http
from ingestion.http import HttpClient, HttpError


class _Response:
    def __init__(self, status_code: int, headers=None, text: str = "") -> None:
        self.status_code = status_code
        self.headers = dict(headers or {})
        self.text = text
        self.content = text.encode()

    @property
    def encoding(self):
        return "utf-8"


class _Session:
    def __init__(self, responses) -> None:
        self._responses = list(responses)
        self.calls = 0

    def get(self, url, headers=None, timeout=None):
        self.calls += 1
        return self._responses.pop(0)


def _client(responses, waits):
    return HttpClient(session=_Session(responses), sleep=waits.append, retries=2)


def test_delta_seconds_is_honoured():
    waits: list[float] = []
    client = _client([_Response(429, {"Retry-After": "7"}), _Response(200, text="ok")], waits)

    assert client.get_text("https://example.test/page") == "ok"
    assert waits == [7.0]


def test_http_date_is_converted_to_the_remaining_wait():
    waits: list[float] = []
    when = datetime.now(timezone.utc) + timedelta(seconds=20)
    client = _client(
        [_Response(429, {"Retry-After": format_datetime(when, usegmt=True)}), _Response(200, text="ok")],
        waits,
    )

    client.get_text("https://example.test/page")
    # Clock skew makes this a range, not an equality.
    assert len(waits) == 1 and 15.0 <= waits[0] <= 21.0


def test_a_long_retry_after_is_clamped_rather_than_obeyed():
    waits: list[float] = []
    client = _client([_Response(429, {"Retry-After": "3600"}), _Response(200, text="ok")], waits)

    client.get_text("https://example.test/page")
    assert waits == [http.MAX_RETRY_AFTER_SECONDS]


def test_no_header_keeps_the_bounded_backoff():
    waits: list[float] = []
    client = _client([_Response(503), _Response(200, text="ok")], waits)

    client.get_text("https://example.test/page")
    assert waits == [1.5]


def test_a_malformed_header_falls_back_to_backoff():
    waits: list[float] = []
    client = _client([_Response(429, {"Retry-After": "soon-ish"}), _Response(200, text="ok")], waits)

    client.get_text("https://example.test/page")
    assert waits == [1.5]


def test_retry_after_does_not_extend_the_retry_budget():
    waits: list[float] = []
    client = _client([_Response(429, {"Retry-After": "1"})] * 3, waits)

    with pytest.raises(HttpError) as raised:
        client.get_text("https://example.test/page")
    assert raised.value.status == 429
    assert len(waits) == 2, "two bounded retries, then the refusal stands"


@pytest.mark.parametrize(
    "value, expected",
    [
        (None, None),
        ("", None),
        ("12", 12.0),
        (" 0 ", 0.0),
        ("-4", 0.0),
        ("nonsense", None),
    ],
)
def test_retry_after_seconds_parsing(value, expected):
    assert http.retry_after_seconds(value) == expected
