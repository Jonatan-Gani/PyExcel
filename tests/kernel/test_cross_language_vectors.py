"""Golden cross-language frame vectors.

Each vector is ``(frame_type, meta, payloads, expected_hex)``. The C#
counterpart at ``tests/PyExcel.Bridge.Tests/CrossLanguageVectorsTests.cs``
asserts encoding the same inputs produces the same bytes. If both pass,
the two implementations produce byte-identical output for these inputs
and their decoders accept the same input.

These vectors deliberately use only types whose JSON encoding is fully
specified — null, bool, integer, ASCII string. Float repr depends on the
language runtime; floats are roundtrip-tested but not byte-pinned.

If a vector below changes, update both sides in the same commit.
"""

from __future__ import annotations

from typing import Tuple

import pytest

from pyexcel.kernel.framing import (
    Frame,
    FrameType,
    buffer_reader,
    encode_frame,
    read_frame,
)


VECTORS: list[Tuple[FrameType, dict, Tuple[bytes, ...], str]] = [
    # PING with empty meta, no payloads. Empty dict still serialises to b"{}".
    # body = 1 + 4 + 2 + 4 = 11 = 0x0b
    # 0b000000 | 02 | 02000000 | 7b7d | 00000000
    (FrameType.PING, {}, (), "0b00000002020000007b7d00000000"),
    # PONG with {"i": 42}, no payloads.
    # meta_bytes = b'{"i":42}'  (8 bytes)
    # body = 1 + 4 + 8 + 4 = 17 = 0x11
    # 11000000 | 03 | 08000000 | 7b226922 3a3432 7d | 00000000
    (
        FrameType.PONG,
        {"i": 42},
        (),
        "1100000003080000007b2269223a34327d00000000",
    ),
    # ERROR with {"code":"BAD"}.
    # meta_bytes = b'{"code":"BAD"}'  (14 bytes)
    # body = 1 + 4 + 14 + 4 = 23 = 0x17
    (
        FrameType.ERROR,
        {"code": "BAD"},
        (),
        "17000000050e0000007b22636f6465223a22424144227d00000000",
    ),
    # RUN_RESULT with {} meta, single payload 0xdeadbeef.
    # meta_bytes = b'{}'  (2 bytes)
    # body = 1 + 4 + 2 + 4 + 4 + 4 = 19 = 0x13
    (
        FrameType.RUN_RESULT,
        {},
        (b"\xde\xad\xbe\xef",),
        "130000000b020000007b7d0100000004000000deadbeef",
    ),
    # HELLO with {"kernel":"py","protocol":2} (alphabetised by sort_keys).
    # meta_bytes = b'{"kernel":"py","protocol":2}'  (28 bytes)
    # body = 1 + 4 + 28 + 4 = 37 = 0x25
    (
        FrameType.HELLO,
        {"protocol": 2, "kernel": "py"},
        (),
        "25000000011c0000007b226b65726e656c223a227079222c2270726f"
        "746f636f6c223a327d00000000",
    ),
    # LOG with {"empty":null,"flag":true,"n":-7}.
    # meta_bytes = b'{"empty":null,"flag":true,"n":-7}'  (33 bytes)
    # body = 1 + 4 + 33 + 4 = 42 = 0x2a
    (
        FrameType.LOG,
        {"flag": True, "n": -7, "empty": None},
        (),
        "2a0000000d210000007b22656d707479223a6e756c6c2c22666c6167"
        "223a747275652c226e223a2d377d00000000",
    ),
]


@pytest.mark.parametrize("ftype,meta,payloads,expected_hex", VECTORS)
def test_encode_matches_golden(
    ftype: FrameType, meta: dict, payloads: Tuple[bytes, ...], expected_hex: str
) -> None:
    wire = encode_frame(ftype, meta, payloads)
    assert wire.hex() == expected_hex


@pytest.mark.parametrize("ftype,meta,payloads,expected_hex", VECTORS)
def test_decode_matches_golden(
    ftype: FrameType, meta: dict, payloads: Tuple[bytes, ...], expected_hex: str
) -> None:
    wire = bytes.fromhex(expected_hex)
    decoded = read_frame(buffer_reader(wire))
    assert decoded == Frame(type=ftype, meta=meta, payloads=list(payloads))
