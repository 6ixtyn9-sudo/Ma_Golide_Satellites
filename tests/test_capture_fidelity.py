"""Byte fidelity of the capture-backed fixtures.

The 2026-11-22 probe package reported a sha256 and a byte count for every
artifact it transmitted. Where a whole JSON body was transmissible, this
repository commits it as received and re-hashes it here, so a later edit to a
"captured" fixture fails loudly instead of quietly becoming shape-only again.
"""

from __future__ import annotations

import hashlib

# path -> (bytes, sha256) exactly as the package reported them.
CAPTURES = {
    "tests/fixtures/lba/championships_2026.json": (
        867,
        "e6d648fb4f57e5a44c6fc3053c974402d103866a95291d758aba6f3a266e70cc",
    ),
    "tests/fixtures/lba/calendar_602.json": (
        11636,
        "9a0e5d28c18eac20d06aeef94b9ca354e0597319301b2ba48628c5853c155101",
    ),
}


def test_committed_captures_are_byte_exact():
    for path, (size, digest) in CAPTURES.items():
        with open(path, "rb") as handle:
            data = handle.read()
        assert len(data) == size, f"{path} is {len(data)} bytes, capture was {size}"
        assert hashlib.sha256(data).hexdigest() == digest, f"{path} no longer matches its capture"
