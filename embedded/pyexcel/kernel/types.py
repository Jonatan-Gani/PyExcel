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
* :class:`ChartSpec` — a JSON chart specification (schema documented in
  :mod:`pyexcel.kernel.chart`). Produced automatically when a user
  ``transform()`` returns a Plotly figure; the host renders it as a
  native Excel chart. Encoded as a string Arrow scalar with schema
  metadata ``pyexcel-shape = chart``.
* :class:`ChartImage` — a rendered figure image (SVG, PNG fallback).
  Produced automatically when a user ``transform()`` returns a
  Matplotlib figure; the host embeds it as a picture. Encoded as a
  binary Arrow scalar with schema metadata ``pyexcel-shape = image``
  and field-level metadata ``pyexcel-image-format = svg|png``.
"""

from __future__ import annotations

from dataclasses import dataclass

# Image formats a ChartImage may carry. SVG is preferred (vector, crisp at
# any zoom); PNG is the fallback when the figure can't render to SVG.
CHART_IMAGE_FORMATS = ("svg", "png")


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


@dataclass(frozen=True)
class ChartSpec:
    """A JSON chart specification the host renders as a native Excel chart.

    ``json`` holds the serialised spec document — see
    :func:`pyexcel.kernel.chart.plotly_figure_to_spec` for the schema.
    Users normally never construct this directly: returning a Plotly
    figure from ``transform()`` converts automatically. Constructing one
    by hand is supported for advanced cases (a spec assembled without
    Plotly).
    """

    json: str

    def __post_init__(self) -> None:
        if not isinstance(self.json, str):
            raise TypeError(
                f"ChartSpec.json must be a string, got {type(self.json).__name__}"
            )
        if not self.json.strip():
            raise ValueError("ChartSpec.json must be non-empty")


@dataclass(frozen=True)
class ChartImage:
    """A rendered figure image the host embeds as a worksheet picture.

    ``data`` is the raw image bytes; ``format`` is one of
    :data:`CHART_IMAGE_FORMATS`. Users normally never construct this
    directly: returning a Matplotlib figure from ``transform()`` converts
    automatically.
    """

    data: bytes
    format: str

    def __post_init__(self) -> None:
        if not isinstance(self.data, bytes):
            raise TypeError(
                f"ChartImage.data must be bytes, got {type(self.data).__name__}"
            )
        if len(self.data) == 0:
            raise ValueError("ChartImage.data must be non-empty")
        if self.format not in CHART_IMAGE_FORMATS:
            raise ValueError(
                f"ChartImage.format must be one of {CHART_IMAGE_FORMATS}, "
                f"got {self.format!r}"
            )
