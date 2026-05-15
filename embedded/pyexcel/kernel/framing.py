"""
Wire framing for the PyExcel v2 kernel protocol.

Frame layout (all integers little-endian, unsigned):

    +---------+----+----------+-----------+-----------+--------------+
    | body    | t  | meta_len | meta_json | n_payload | payloads...  |
    | u32     | u8 | u32      | bytes     | u32       |              |
    +---------+----+----------+-----------+-----------+--------------+

    body       = byte length of everything after this field (the frame body)
    t          = frame type, see FrameType
    meta_len   = byte length of meta_json
    meta_json  = UTF-8 encoded JSON object, application-defined per frame type
    n_payload  = number of binary payloads (typically 1 Arrow IPC stream per run)
    payload_i  = u32 size + raw bytes, repeated n_payload times

Design rules:

* **Pure stdlib.** No pyarrow/lxml imports here — those live in higher
  layers. Framing must stay loadable from the supervisor before any user
  dependency has been validated.
* **Bounded.** A maximum frame size is enforced on both encode and decode
  (default 256 MiB). A malformed peer cannot exhaust memory.
* **Stream-friendly.** ``read_frame`` reads exactly the bytes it needs from
  a ``read_exact`` callable. Partial reads are the caller's problem to
  compose with select/overlapped I/O.
* **Deterministic.** Same logical inputs always produce byte-identical
  frames; sorted JSON keys, no whitespace.

The C# counterpart in ``src/PyExcel.Bridge/Framing.cs`` (added in Phase 2)
must match this layout byte-for-byte. ``PROTOCOL_VERSION`` is the agreed
on-wire schema version both sides advertise in the ``Hello`` frame.
"""

from __future__ import annotations

import enum
import json
import struct
from dataclasses import dataclass, field
from typing import Callable, List, Tuple

PROTOCOL_VERSION = 2

DEFAULT_MAX_FRAME_BYTES = 256 * 1024 * 1024  # 256 MiB

_U8 = 1
_U32 = 4

_U32_LE = struct.Struct("<I")
_U8_LE = struct.Struct("<B")


class FrameType(enum.IntEnum):
    """Stable across the v2 protocol; new types append, existing never re-number."""

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
    LIST_JOBS = 15


class FramingError(Exception):
    """Base for all framing-layer errors."""


class FrameTooLargeError(FramingError):
    """Frame size exceeds the configured cap (caught *before* allocating)."""


class TruncatedFrameError(FramingError):
    """``read_exact`` returned fewer bytes than requested. Peer likely closed."""


class MalformedFrameError(FramingError):
    """Frame parsed but internal lengths are inconsistent."""


@dataclass(frozen=True)
class Frame:
    """One decoded wire frame.

    ``meta`` is always a parsed JSON object (dict).
    ``payloads`` is raw byte buffers in wire order.
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
    """Serialize a frame to a single ``bytes`` ready to write to the transport.

    Raises:
        FrameTooLargeError: encoded frame body would exceed ``max_frame_bytes``.
        TypeError: ``meta`` not JSON-serializable, or a payload isn't bytes-like.
        ValueError: ``frame_type`` not a member of :class:`FrameType`.
    """
    if not isinstance(frame_type, FrameType):
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

    for i, p in enumerate(payloads):
        if not isinstance(p, (bytes, bytearray, memoryview)):
            raise TypeError(f"payload {i} must be bytes-like, got {type(p).__name__}")

    payload_count = len(payloads)
    payload_bytes_total = sum(len(p) for p in payloads)

    body_len = (
        _U8
        + _U32
        + len(meta_bytes)
        + _U32
        + payload_count * _U32
        + payload_bytes_total
    )

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
        out += bytes(p)

    return bytes(out)


# -----------------------------------------------------------------------------
# Decode
# -----------------------------------------------------------------------------


ReadExact = Callable[[int], bytes]
"""``read_exact(n)`` returns exactly ``n`` bytes, blocking as needed.

Returns shorter buffer only on clean EOF; framing treats short returns as
end-of-stream and raises :class:`TruncatedFrameError`.
"""


def read_frame(
    read_exact: ReadExact,
    *,
    max_frame_bytes: int = DEFAULT_MAX_FRAME_BYTES,
) -> Frame:
    """Read one frame from a stream.

    Raises:
        TruncatedFrameError: connection closed mid-frame.
        FrameTooLargeError: header announced a body larger than ``max_frame_bytes``.
        MalformedFrameError: header parsed but internal lengths inconsistent.
    """
    total_buf = read_exact(_U32)
    if len(total_buf) < _U32:
        raise TruncatedFrameError("EOF reading frame length prefix")
    (body_len,) = _U32_LE.unpack(total_buf)

    if body_len > max_frame_bytes:
        raise FrameTooLargeError(
            f"peer announced frame of {body_len} bytes; cap is {max_frame_bytes}"
        )
    if body_len < _U8 + _U32 + _U32:
        raise MalformedFrameError(f"frame body length {body_len} is below minimum")

    body = read_exact(body_len)
    if len(body) < body_len:
        raise TruncatedFrameError(
            f"EOF in frame body: expected {body_len} bytes, got {len(body)}"
        )

    offset = 0

    (type_byte,) = _U8_LE.unpack_from(body, offset)
    offset += _U8
    try:
        ftype = FrameType(type_byte)
    except ValueError as exc:
        raise MalformedFrameError(f"unknown frame type byte {type_byte}") from exc

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

    if offset + _U32 > body_len:
        raise MalformedFrameError("frame truncated before payload count")
    (n_payload,) = _U32_LE.unpack_from(body, offset)
    offset += _U32

    payloads: List[bytes] = []
    for i in range(n_payload):
        if offset + _U32 > body_len:
            raise MalformedFrameError(f"frame truncated before size of payload {i}")
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


def buffer_reader(buf: bytes) -> ReadExact:
    """Return a ``read_exact`` callable backed by an in-memory buffer.

    Used by tests and by the stdin-driven launcher mode. Yields ``b""``
    once the buffer is drained, which framing treats as EOF.
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
