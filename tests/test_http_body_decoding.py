"""Undecodable bodies: an undeclared gzip stream must still become text.

Evidence: external probe 2026-11-23 (Z.ai), recorded in
``lane-probe-2026-11-23.md``. B.League's S3 context snapshot for ScheduleKey
506602 came back with body magic ``1f 8b``, ``content-type: application/json``
and no usable text form, while ``latestid`` beside it returns a bare integer.
Decompressed it is a complete 543,231-byte payload with 725 ``PlayByPlays``
rows. A client that returns ``resp.text`` for that body hands the JSON parser
replacement characters, so the lane reads nothing and blames the payload.
"""

from __future__ import annotations

import gzip
import json
from dataclasses import dataclass

from ingestion.http import GZIP_MAGIC, _decoded_body


@dataclass
class _Response:
    content: bytes
    text: str
    status_code: int = 200


def test_undeclared_gzip_body_is_decompressed():
    payload = json.dumps({"PlayByPlays": [{"Period": 1}], "Game": {"ScheduleKey": 506602}})
    compressed = gzip.compress(payload.encode("utf-8"))
    assert compressed[:2] == GZIP_MAGIC

    response = _Response(content=compressed, text=compressed.decode("latin-1"))

    decoded = _decoded_body(response)

    assert json.loads(decoded)["Game"]["ScheduleKey"] == 506602


def test_a_plain_body_is_returned_untouched():
    response = _Response(content=b'{"ok": true}', text='{"ok": true}')
    assert _decoded_body(response) == '{"ok": true}'


def test_broken_gzip_falls_back_to_the_text_attribute():
    # Magic bytes but not a real stream: prefer the declared text over raising
    # on a provider read that might still be usable.
    response = _Response(content=GZIP_MAGIC + b"not really gzip", text="fallback text")
    assert _decoded_body(response) == "fallback text"


def test_response_without_a_content_attribute_still_works():
    class _TextOnly:
        text = '{"ok": true}'

    assert _decoded_body(_TextOnly()) == '{"ok": true}'


def test_corrupt_deflate_behind_a_valid_header_falls_back():
    """The case an external audit found in the first version of this helper.

    Catching only ``OSError``/``EOFError`` was not enough: ``gzip.decompress``
    raises ``zlib.error`` -- a zlib-module exception that is neither -- when the
    gzip header parses but the deflate stream behind it is corrupt. A raw
    ``zlib.error`` then escaped ``get_text``, violating this module's contract of
    raising only ``HttpError``. The test above cannot catch that: its fixture has
    a broken *header* (BadGzipFile, an OSError).
    """

    raw = bytearray(gzip.compress(b'{"a": 1}'))
    raw[12] ^= 0xFF
    raw[13] ^= 0xFF  # header intact, deflate stream corrupt

    response = _Response(content=bytes(raw), text="fallback text")

    assert _decoded_body(response) == "fallback text"


@dataclass
class _Encoded:
    content: bytes
    charset: str | None = None

    @property
    def headers(self) -> dict:
        if self.charset is None:
            return {"content-type": "text/html"}
        return {"content-type": f"text/html; charset={self.charset}"}


def test_utf8_with_no_declared_charset_is_read_as_utf8():
    """The bug that hid a Greek date behind a control character.

    ``Σαβ 21 Νοε`` served as UTF-8 with no charset header became
    ``Î£Î±Î² 21 Î\\x9dÎ¿Îµ`` under requests' latin-1 default: two characters for
    every Greek letter, one of them a control code. No table of Greek month names
    can match that, and no terminal can show why.
    """

    from ingestion.http import decode_body_bytes

    body = "Σαβ 21 Νοε - 15:00".encode("utf-8")

    assert _decoded_body(_Encoded(content=body)) == "Σαβ 21 Νοε - 15:00"
    assert decode_body_bytes(body) == "Σαβ 21 Νοε - 15:00"


def test_a_declared_charset_is_still_honoured():
    from ingestion.http import decode_body_bytes

    body = "Σαβ 21 Νοε".encode("iso-8859-7")

    assert decode_body_bytes(body, "iso-8859-7") == "Σαβ 21 Νοε"
    assert _decoded_body(_Encoded(content=body, charset="iso-8859-7")) == "Σαβ 21 Νοε"


def test_bytes_that_are_not_utf8_fall_back_to_latin1_rather_than_raising():
    from ingestion.http import decode_body_bytes

    assert decode_body_bytes(b"caf\xe9 21") == "café 21"
