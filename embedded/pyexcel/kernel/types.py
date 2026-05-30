"""User-facing types that flow across the kernel boundary.

These are the typed values a user :func:`transform` may return (or accept
as input) where a plain Python primitive isn't enough to express the
intent. Each one has a wire representation in :mod:`pyexcel.kernel.arrow_io`
and a matching .NET type on the host side.

Currently:

* :class:`Formula` — an Excel A1-mode formula string. The kernel encodes
  a ``Formula`` as a string Arrow column with field-level metadata
  ``pyexcel-cell-type = formula``; the host writes it via
  ``Range.Formula`` (not ``Range.Value2``) so Excel recomputes it on
  every recalc.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Formula:
    """An Excel A1-mode formula (e.g. ``=SUM(A1:B2)``).

    The text must start with ``=`` — the host rejects everything else as
    "not a formula" rather than guessing.
    """

    text: str

    def __post_init__(self) -> None:
        if not isinstance(self.text, str):
            raise TypeError(
                f"Formula.text must be a string, got {type(self.text).__name__}"
            )
        if not self.text.startswith("="):
            raise ValueError(f"formula must start with '=': {self.text!r}")
