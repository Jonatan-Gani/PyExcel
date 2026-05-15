"""
Wire framing for the PyExcel v2 kernel protocol.

Frame layout (all integers little-endian, unsigned):

    +---------+----+----------+-----------+-----------+--------------+
    | total   | t  | meta_len | meta_json | n_payload | payloads...  |
    | u32     | u8 | u32      | bytes     | u32       |              |
    +---------+----+----------+-----------+-----------+--------------+

    total      = byte length of everything after this field (i.e. the frame body)
    t          = frame type, see FrameType
    meta_len   = byte length of meta_json
    meta_json  = UTF-8 encoded JSON object, application-defined per frame type
    n_payload  = number of binary payloads (often 0)
    payload_i  = u32 size + raw bytes, repeated n_payload times

Design rules (see docs/v2-safety-contract.md):

* Pure stdlib. No pyarrow/numpy/lxml imports here -- those live in artifact
  serialization layers above. Framing must stay loadable from the supervisor
  before any user dependency has been validated.
* Bounded. A maximum frame size is enforced both on encode and decode so a
  malformed peer cannot exhaust memory. The cap is large enough for realistic
  Variant blobs (default 256 MiB).
* Stream-friendly. ``read_frame`` reads exactly the bytes it needs from a
  ``read_exact`` callable and returns; partial reads are the caller's problem
  to compose with select/overlapped I/O.
* Deterministic. Same inputs always produce byte-identical frames; the
  roundtrip property is fuzz-tested.

The protocol version is the v2 safety contract version. Bumping the wire
format bumps ``PROTOCOL_VERSION`` and the matching ``KC_PROTOCOL_VERSION``
constant in ``src/module/kernelClient.bas`` in the same change.
"""

from __future__ import annotations

import enum
import json
import struct
from dataclasses import dataclass, field
from typing import Callable, List, Tuple

PROTOCOL_VERSION = 2

DEFAULT_MAX_FRAME_BYTES = 256 * 1024 * 1024  # 256 MiB

# Header field widths, in bytes.
_U8 = 1
_U32 = 4

# Pre-built struct objects -- avoids re-parsing the format string per call.
_U32_LE = struct.Struct("<I")
_U8_LE = struct.Struct("<B")


class FrameType(enum.IntEnum):
    """
    All wire frame types. Values are stable across the v2 protocol; new types
    append, existing types never re-number.

    The split between control frames (Hello/Ping/Pong/Shutdown/Error) and run
    frames (RunRequest/RunResult/Progress/Log/Cancel) is documented in
    docs/wire-protocol.md.
    """

    HELLO = 1
    PING = 2
    PONG = 3
    SHUTDOWN = 4
    ERROR = 5

    RUN_REQUEST = 10
    RUN_RESULT = 11
    PROGRESS = 12
    LOG = 13
    CANCEL = 14


class FramingError(Exception):
    """Base for all framing-layer errors."""


class FrameTooLargeError(FramingError):
    """Frame size exceeds the configured cap. Decoded *before* the body is read."""


class TruncatedFrameError(FramingError):
    """``read_exact`` returned fewer bytes than requested. Connection is likely dead."""


class MalformedFrameError(FramingError):
    """Frame header parsed but contents are inconsistent (e.g. meta_len > body)."""


@dataclass(frozen=True)
class Frame:
    """
    One decoded wire frame.

    ``meta`` is a parsed JSON object (always a ``dict`` at the top level by
    convention; callers may assert this if they want).

    ``payloads`` is a list of raw byte buffers in the order they appeared on the
    wire. Lengths are implicit in ``len(bytes)``.
    """

    type: FrameType
    meta: dict
    payloads: List[bytes] = field(default_factory=list)


# -----------------------------------------------------------------------------
# Encode
# -----------------------------------------------------------------------------


def encode_frame(
    frame_type: FrameType,
    meta: dict,
    payloads: Tuple[bytes, ...] = (),
    *,
    max_frame_bytes: int = DEFAULT_MAX_FRAME_BYTES,
) -> bytes:
    """
    Serialize a frame to a single ``bytes`` object ready to be written to the
    transport.

    Raises:
        FrameTooLargeError: if the encoded frame would exceed ``max_frame_bytes``.
        TypeError: if ``meta`` is not JSON-serializable.
    """
    if not isinstance(frame_type, FrameType):
        # Accept int values too, but reject anything outside the enum.
        try:
            frame_type = FrameType(frame_type)
        except ValueError as exc:
            raise ValueError(f"unknown frame type: {frame_type!r}") from exc

    meta_bytes = json.dumps(
        meta,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")

    # Validate payloads up-front.
    for i, p in enumerate(payloads):
        if not isinstance(p, (bytes, bytearray, memoryview)):
            raise TypeError(
                f"payload {i} must be bytes-like, got {type(p).__name__}"
            )

    payload_count = len(payloads)
    payload_bytes_total = sum(len(p) for p in payloads)

    # body = u8(type) + u32(meta_len) + meta + u32(n_payload) + (u32 + bytes) * n
    body_len = (
        _U8
        + _U32
        + len(meta_bytes)
        + _U32
        + payload_count * _U32
        + payload_bytes_total
    )

    # Frame total on the wire = u32(total) + body. ``total`` itself counts the body
    # only, so the on-wire size is body_len + _U32. Enforce the cap on the body
    # size; that's what a downstream reader has to allocate.
    if body_len > max_frame_bytes:
        raise FrameTooLargeError(
            f"encoded frame body of {body_len} bytes exceeds cap of "
            f"{max_frame_bytes} bytes"
        )

    out = bytearray()
    out += _U32_LE.pack(body_len)
    out += _U8_LE.pack(int(frame_type))
    out += _U32_LE.pack(len(meta_bytes))
    out += meta_bytes
    out += _U32_LE.pack(payload_count)
    for p in payloads:
        out += _U32_LE.pack(len(p))
        # bytes-like is fine for extend; bytearray() coerces memoryview correctly.
        out += bytes(p)

    return bytes(out)


# -----------------------------------------------------------------------------
# Decode
# -----------------------------------------------------------------------------


ReadExact = Callable[[int], bytes]
"""
Callable contract: ``read_exact(n)`` returns exactly ``n`` bytes, blocking as
needed. Returns ``b""`` only when the peer has cleanly closed -- shorter reads
on an open connection are errors. The framing layer treats short returns as
end-of-stream and raises ``TruncatedFrameError``.
"""


def read_frame(
    read_exact: ReadExact,
    *,
    max_frame_bytes: int = DEFAULT_MAX_FRAME_BYTES,
) -> Frame:
    """
    Read one frame from a stream.

    ``read_exact(n)`` must return exactly ``n`` bytes or fewer (only on clean
    EOF). This function never partially advances state on its caller -- if a
    read fails, no half-frame is buffered; the caller may safely close the
    transport.

    Raises:
        TruncatedFrameError: connection closed mid-frame.
        FrameTooLargeError: header announced a body larger than ``max_frame_bytes``.
        MalformedFrameError: header parsed but internal lengths are inconsistent.
    """
    # Length prefix.
    total_buf = read_exact(_U32)
    if len(total_buf) < _U32:
        raise TruncatedFrameError("EOF reading frame length prefix")
    (body_len,) = _U32_LE.unpack(total_buf)

    if body_len > max_frame_bytes:
        raise FrameTooLargeError(
            f"peer announced frame of {body_len} bytes; cap is {max_frame_bytes}"
        )
    if body_len < _U8 + _U32 + _U32:
        # Minimum body: type(1) + meta_len(4) + 0-byte meta + n_payload(4)
        raise MalformedFrameError(f"frame body length {body_len} is below minimum")

    body = read_exact(body_len)
    if len(body) < body_len:
        raise TruncatedFrameError(
            f"EOF in frame body: expected {body_len} bytes, got {len(body)}"
        )

    offset = 0

    # Frame type.
    (type_byte,) = _U8_LE.unpack_from(body, offset)
    offset += _U8
    try:
        ftype = FrameType(type_byte)
    except ValueError as exc:
        raise MalformedFrameError(f"unknown frame type byte {type_byte}") from exc

    # Meta.
    (meta_len,) = _U32_LE.unpack_from(body, offset)
    offset += _U32
    if offset + meta_len > body_len:
        raise MalformedFrameError(
            f"meta_len {meta_len} would read past end of frame body"
        )
    meta_bytes = bytes(body[offset : offset + meta_len])
    offset += meta_len

    try:
        meta = json.loads(meta_bytes.decode("utf-8")) if meta_len else {}
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise MalformedFrameError(f"meta is not valid UTF-8 JSON: {exc}") from exc

    if not isinstance(meta, dict):
        raise MalformedFrameError(
            f"meta must be a JSON object, got {type(meta).__name__}"
        )

    # Payload count.
    if offset + _U32 > body_len:
        raise MalformedFrameError("frame truncated before payload count")
    (n_payload,) = _U32_LE.unpack_from(body, offset)
    offset += _U32

    payloads: List[bytes] = []
    for i in range(n_payload):
        if offset + _U32 > body_len:
            raise MalformedFrameError(
                f"frame truncated before size of payload {i}"
            )
        (psize,) = _U32_LE.unpack_from(body, offset)
        offset += _U32
        if offset + psize > body_len:
            raise MalformedFrameError(
                f"payload {i} size {psize} would read past end of frame body"
            )
        payloads.append(bytes(body[offset : offset + psize]))
        offset += psize

    if offset != body_len:
        raise MalformedFrameError(
            f"frame has {body_len - offset} trailing bytes after parsing"
        )

    return Frame(type=ftype, meta=meta, payloads=payloads)


# -----------------------------------------------------------------------------
# Convenience: build a ReadExact from a byte buffer (used by tests and by the
# stdin-driven launcher mode).
# -----------------------------------------------------------------------------


def buffer_reader(buf: bytes) -> ReadExact:
    """
    Return a ``read_exact`` callable backed by an in-memory buffer. Yields
    ``b""`` once the buffer is drained, which the framing layer treats as
    EOF.
    """
    pos = [0]
    blen = len(buf)

    def _read(n: int) -> bytes:
        start = pos[0]
        end = min(start + n, blen)
        chunk = buf[start:end]
        pos[0] = end
        return chunk

    return _read
