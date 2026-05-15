"""
Tests for ``pyexcel.kernel.framing``.

Covers:
- Roundtrip identity for a representative set of frame shapes.
- Bounded behavior on oversized frames (encode and decode side).
- Truncation handling (peer disconnect mid-frame).
- Malformed-frame detection (inconsistent internal lengths, bad JSON,
  unknown frame types).
- Deterministic encoding (sorted JSON keys, no extra whitespace).
- Property fuzz: random meta + random payload sets roundtrip byte-identically.
"""

from __future__ import annotations

import json
import os
import random
from typing import List, Tuple

import pytest

from pyexcel.kernel.framing import (
    DEFAULT_MAX_FRAME_BYTES,
    Frame,
    FrameTooLargeError,
    FrameType,
    MalformedFrameError,
    PROTOCOL_VERSION,
    TruncatedFrameError,
    buffer_reader,
    encode_frame,
    read_frame,
)


# -----------------------------------------------------------------------------
# Roundtrip
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "ftype,meta,payloads",
    [
        (FrameType.HELLO, {"protocol": PROTOCOL_VERSION, "kernel": "abc"}, ()),
        (FrameType.PING, {}, ()),
        (FrameType.PONG, {}, ()),
        (FrameType.RUN_REQUEST, {"script": "foo.py", "run_id": "r1"}, (b"\x00\x01\x02",)),
        (
            FrameType.RUN_RESULT,
            {"status": "done", "duration": 0.123},
            (b"table-bytes", b"another-payload-blob"),
        ),
        (FrameType.PROGRESS, {"pct": 42, "message": "halfway"}, ()),
        (FrameType.LOG, {"level": "info", "text": "hello"}, ()),
        (FrameType.CANCEL, {"run_id": "r1"}, ()),
        (FrameType.SHUTDOWN, {"drain": True, "timeout_ms": 3000}, ()),
        (FrameType.ERROR, {"code": "PIPE_ACL_FAILURE", "message": "no SID match"}, ()),
    ],
)
def test_roundtrip(ftype: FrameType, meta: dict, payloads: Tuple[bytes, ...]) -> None:
    wire = encode_frame(ftype, meta, payloads)
    decoded = read_frame(buffer_reader(wire))

    assert decoded.type == ftype
    assert decoded.meta == meta
    assert decoded.payloads == list(payloads)


def test_unicode_meta_roundtrip() -> None:
    meta = {"name": "Säles", "emoji": "✓", "ja": "日本語"}
    wire = encode_frame(FrameType.LOG, meta)
    decoded = read_frame(buffer_reader(wire))
    assert decoded.meta == meta


def test_empty_meta_roundtrip() -> None:
    wire = encode_frame(FrameType.PING, {})
    decoded = read_frame(buffer_reader(wire))
    assert decoded.type == FrameType.PING
    assert decoded.meta == {}
    assert decoded.payloads == []


def test_zero_length_payload_roundtrip() -> None:
    wire = encode_frame(FrameType.RUN_RESULT, {"k": "v"}, (b"", b"after-empty"))
    decoded = read_frame(buffer_reader(wire))
    assert decoded.payloads == [b"", b"after-empty"]


def test_large_payload_roundtrip() -> None:
    # 8 MiB payload, well under the default 256 MiB cap.
    payload = os.urandom(8 * 1024 * 1024)
    wire = encode_frame(FrameType.RUN_RESULT, {"size": len(payload)}, (payload,))
    decoded = read_frame(buffer_reader(wire))
    assert decoded.payloads == [payload]


# -----------------------------------------------------------------------------
# Determinism
# -----------------------------------------------------------------------------


def test_deterministic_encoding() -> None:
    # Same logical frame must produce byte-identical output regardless of dict
    # insertion order -- meta is encoded with sort_keys=True.
    a = encode_frame(FrameType.LOG, {"b": 2, "a": 1})
    b = encode_frame(FrameType.LOG, {"a": 1, "b": 2})
    assert a == b


def test_encoding_has_no_whitespace_padding() -> None:
    # Confirm the JSON separator choice (no spaces). Catches accidental
    # `json.dumps(..., indent=2)` regressions.
    wire = encode_frame(FrameType.LOG, {"a": 1, "b": 2})
    # The meta JSON appears somewhere inside the body.
    assert b'{"a":1,"b":2}' in wire
    assert b'{"a": 1' not in wire  # no space-after-colon


# -----------------------------------------------------------------------------
# Size cap (encode side)
# -----------------------------------------------------------------------------


def test_encode_rejects_oversize_frame() -> None:
    cap = 1024
    too_big_payload = b"x" * (cap + 1)
    with pytest.raises(FrameTooLargeError):
        encode_frame(FrameType.RUN_RESULT, {}, (too_big_payload,), max_frame_bytes=cap)


def test_encode_accepts_at_cap() -> None:
    # The cap applies to the body length; pick a payload size that fits with
    # the smallest possible meta and headers.
    cap = 256
    # meta is {} -> 2 bytes; body_len = 1 + 4 + 2 + 4 + 4 + N = 15 + N
    payload = b"y" * (cap - 15)
    wire = encode_frame(FrameType.RUN_RESULT, {}, (payload,), max_frame_bytes=cap)
    decoded = read_frame(buffer_reader(wire), max_frame_bytes=cap)
    assert decoded.payloads == [payload]


# -----------------------------------------------------------------------------
# Size cap (decode side)
# -----------------------------------------------------------------------------


def test_decode_rejects_oversize_announcement() -> None:
    # Hand-craft a frame whose length prefix exceeds the cap. Decoder must
    # reject before allocating the body.
    import struct

    announced_body = 10 * 1024 * 1024  # 10 MiB
    cap = 1024
    bogus = struct.pack("<I", announced_body)
    with pytest.raises(FrameTooLargeError):
        read_frame(buffer_reader(bogus), max_frame_bytes=cap)


# -----------------------------------------------------------------------------
# Truncation handling
# -----------------------------------------------------------------------------


def test_truncated_before_length_prefix() -> None:
    with pytest.raises(TruncatedFrameError):
        read_frame(buffer_reader(b""))


def test_truncated_length_prefix() -> None:
    with pytest.raises(TruncatedFrameError):
        read_frame(buffer_reader(b"\x01\x02"))  # only 2 of 4 bytes


def test_truncated_body() -> None:
    wire = encode_frame(FrameType.RUN_RESULT, {"k": "v"}, (b"abc",))
    # Lop off the last byte.
    with pytest.raises(TruncatedFrameError):
        read_frame(buffer_reader(wire[:-1]))


# -----------------------------------------------------------------------------
# Malformed-frame detection
# -----------------------------------------------------------------------------


def test_malformed_unknown_frame_type() -> None:
    import struct

    # Minimum body: type(1) + meta_len(4) + 0-byte meta + n_payload(4) = 9 bytes.
    body = struct.pack("<B", 99) + struct.pack("<I", 0) + struct.pack("<I", 0)
    wire = struct.pack("<I", len(body)) + body
    with pytest.raises(MalformedFrameError, match="unknown frame type"):
        read_frame(buffer_reader(wire))


def test_malformed_meta_len_overruns_body() -> None:
    import struct

    # Claim meta_len = 100 but only ship 9-byte body.
    body = struct.pack("<B", FrameType.PING.value) + struct.pack("<I", 100) + struct.pack("<I", 0)
    wire = struct.pack("<I", len(body)) + body
    with pytest.raises(MalformedFrameError):
        read_frame(buffer_reader(wire))


def test_malformed_meta_not_json() -> None:
    import struct

    bad_meta = b"\xff\xfe not json"
    body = (
        struct.pack("<B", FrameType.LOG.value)
        + struct.pack("<I", len(bad_meta))
        + bad_meta
        + struct.pack("<I", 0)
    )
    wire = struct.pack("<I", len(body)) + body
    with pytest.raises(MalformedFrameError, match="meta is not valid UTF-8 JSON"):
        read_frame(buffer_reader(wire))


def test_malformed_meta_not_object() -> None:
    # Top-level JSON must be an object (the contract). A JSON array is invalid.
    import struct

    arr_bytes = b"[1,2,3]"
    body = (
        struct.pack("<B", FrameType.LOG.value)
        + struct.pack("<I", len(arr_bytes))
        + arr_bytes
        + struct.pack("<I", 0)
    )
    wire = struct.pack("<I", len(body)) + body
    with pytest.raises(MalformedFrameError, match="meta must be a JSON object"):
        read_frame(buffer_reader(wire))


def test_malformed_payload_size_overruns() -> None:
    import struct

    # meta = {}, payload count = 1, payload size = 1000, but only 1 byte present.
    meta_bytes = b"{}"
    body = (
        struct.pack("<B", FrameType.RUN_RESULT.value)
        + struct.pack("<I", len(meta_bytes))
        + meta_bytes
        + struct.pack("<I", 1)
        + struct.pack("<I", 1000)
        + b"\x00"
    )
    wire = struct.pack("<I", len(body)) + body
    with pytest.raises(MalformedFrameError):
        read_frame(buffer_reader(wire))


def test_malformed_body_minimum() -> None:
    import struct

    # Announce a body smaller than the absolute minimum.
    wire = struct.pack("<I", 3) + b"\x01\x02\x03"
    with pytest.raises(MalformedFrameError, match="below minimum"):
        read_frame(buffer_reader(wire))


# -----------------------------------------------------------------------------
# Encoding-side type checks
# -----------------------------------------------------------------------------


def test_encode_rejects_non_bytes_payload() -> None:
    with pytest.raises(TypeError):
        encode_frame(FrameType.RUN_RESULT, {}, ("not bytes",))  # type: ignore[arg-type]


def test_encode_rejects_non_json_meta() -> None:
    class NotSerializable:
        pass

    with pytest.raises(TypeError):
        encode_frame(FrameType.LOG, {"obj": NotSerializable()})


def test_encode_rejects_unknown_frame_type_int() -> None:
    with pytest.raises(ValueError, match="unknown frame type"):
        encode_frame(999, {})  # type: ignore[arg-type]


def test_encode_accepts_bytearray_and_memoryview() -> None:
    ba = bytearray(b"hello")
    mv = memoryview(b"world")
    wire = encode_frame(FrameType.RUN_RESULT, {}, (ba, mv))
    decoded = read_frame(buffer_reader(wire))
    assert decoded.payloads == [b"hello", b"world"]


# -----------------------------------------------------------------------------
# Streaming behavior: read_frame is correctly composable
# -----------------------------------------------------------------------------


def test_back_to_back_frames_stream() -> None:
    # Multiple frames in one buffer must decode sequentially via the same
    # read_exact callable.
    a = encode_frame(FrameType.PING, {"i": 1})
    b = encode_frame(FrameType.LOG, {"text": "two"}, (b"payload",))
    c = encode_frame(FrameType.PONG, {"i": 1})

    reader = buffer_reader(a + b + c)
    f1 = read_frame(reader)
    f2 = read_frame(reader)
    f3 = read_frame(reader)

    assert f1.type == FrameType.PING and f1.meta == {"i": 1}
    assert f2.type == FrameType.LOG and f2.payloads == [b"payload"]
    assert f3.type == FrameType.PONG


# -----------------------------------------------------------------------------
# Property-based fuzz
# -----------------------------------------------------------------------------


def _random_meta(rng: random.Random) -> dict:
    keys = ["a", "b", "c", "name", "id", "value", "kv", "nested"]
    out: dict = {}
    for _ in range(rng.randint(0, 6)):
        k = rng.choice(keys)
        kind = rng.choice(["int", "float", "str", "bool", "none", "list"])
        if kind == "int":
            out[k] = rng.randint(-(2**31), 2**31 - 1)
        elif kind == "float":
            out[k] = rng.random() * 1e6
        elif kind == "str":
            out[k] = "".join(rng.choices("abcdef日本✓", k=rng.randint(0, 16)))
        elif kind == "bool":
            out[k] = rng.choice([True, False])
        elif kind == "none":
            out[k] = None
        elif kind == "list":
            out[k] = [rng.randint(0, 99) for _ in range(rng.randint(0, 4))]
    return out


def _random_payloads(rng: random.Random) -> List[bytes]:
    n = rng.randint(0, 4)
    return [os.urandom(rng.randint(0, 256)) for _ in range(n)]


@pytest.mark.parametrize("seed", range(20))
def test_fuzz_roundtrip(seed: int) -> None:
    rng = random.Random(seed)
    ftype = rng.choice(list(FrameType))
    meta = _random_meta(rng)
    payloads = _random_payloads(rng)

    wire = encode_frame(ftype, meta, tuple(payloads))
    decoded = read_frame(buffer_reader(wire))

    assert decoded.type == ftype
    # Compare via JSON canonical form to be robust to nested-list/tuple parity.
    assert json.dumps(decoded.meta, sort_keys=True) == json.dumps(meta, sort_keys=True)
    assert decoded.payloads == payloads


# -----------------------------------------------------------------------------
# Protocol-version constant is exposed
# -----------------------------------------------------------------------------


def test_protocol_version_is_int() -> None:
    assert isinstance(PROTOCOL_VERSION, int)
    assert PROTOCOL_VERSION >= 2  # the v2 contract requires PROTOCOL_VERSION == 2 currently
