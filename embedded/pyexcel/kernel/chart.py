"""
Figure → wire-value conversion for the v2 kernel.

A user ``transform()`` may return a Plotly figure or a Matplotlib figure;
this module turns each into a typed wire value :mod:`pyexcel.kernel.arrow_io`
knows how to encode:

* Plotly figure → :class:`pyexcel.kernel.types.ChartSpec` carrying a JSON
  **chart spec** document. The host parses the spec and builds a native
  Excel chart (``PyExcel.Excel.ChartBuilder``). This is the v2 port of
  v1's ``PlotlyToExcelXMLConverter`` traversal — same figure walk, but the
  output is JSON carried inside the ``RUN_RESULT`` frame instead of an XML
  file on disk, and data lands as native JSON arrays instead of
  comma-joined strings (so values containing commas can't corrupt a row).
* Matplotlib figure → :class:`pyexcel.kernel.types.ChartImage` carrying
  rendered image bytes — SVG preferred, PNG fallback. The host embeds the
  image as a worksheet picture.

Neither Plotly nor Matplotlib is imported by this module — detection is
duck-typed on the value's class, so the kernel boots and runs fine in an
environment without either package. If the user returned a figure, the
package that produced it is necessarily already imported.

Chart spec JSON schema (version 1)
----------------------------------

::

    {
      "version": 1,
      "chart_type": "xy" | "line" | "bar" | "area" | "pie",
      "title": str,                       # "" when the figure has none
      "x_axis": {"title": str, "min": num|null, "max": num|null,
                 "log_scale": bool},
      "y_axis": {  same shape  },
      "legend": {"visible": bool, "position": "right" | "bottom"},
      "barmode": str,                     # "group" | "stack" | "overlay" | …
      "traces": [
        {
          "id": int,                      # 1-based, unique
          "x": [num | str | null, ...],
          "y": [num | null, ...],
          "text": [str, ...] | null,      # per-point labels
          "size": [num, ...] | null,      # bubble sizes
          "style": {
            "series_type": str,           # see SUPPORTED_SERIES_TYPES
            "name": str,
            "axis_group": "primary" | "secondary",
            "visible": bool,
            "line": {"color": str, "dash": str, "width": num|null},
            "marker": {"size": num, "color": str, "shape": str},
            "fill_color": str | null,
            "fill_opacity": num | null
          }
        }
      ],
      "annotations": [
        {
          "id": int,
          "type": "event_line" | "threshold",
          "axis": "x" | "y",
          "value": num | str,
          "style": {"color": str|null, "width": num|null,
                    "dash": str, "opacity": num|null},
          "label": str | null
        }
      ]
    }

Unsupported Plotly trace types (box, violin, heatmap, …) raise
:class:`UnsupportedChartTypeError` with the offending type named, so the
failure surfaces to the user as an actionable error instead of a silently
wrong chart.
"""

from __future__ import annotations

import datetime as _dt
import json
import math
from typing import Any, List, Optional

from .types import ChartImage, ChartSpec

CHART_SPEC_VERSION = 1

# The series types the host's ChartBuilder renders. The spec contract:
# every trace's style.series_type is one of these.
SUPPORTED_SERIES_TYPES = (
    "scatter",
    "scatter_lines",
    "scatter_lines_markers",
    "line",
    "column",
    "bar",
    "area",
    "pie",
    "histogram",
    "bubble",
)

# Plotly trace types we deliberately do not render — there is no faithful
# native Excel counterpart for them in the supported set. Kept explicit so
# the error message can say "unsupported" rather than "unknown".
_UNSUPPORTED_TRACE_TYPES = (
    "box",
    "violin",
    "heatmap",
    "waterfall",
    "candlestick",
    "ohlc",
    "scatter3d",
    "surface",
    "sunburst",
    "treemap",
    "funnel",
)

# Plotly's default qualitative palette (D3 category10) — used when a
# bar/column/histogram trace carries no explicit colour, matching what the
# user saw in the Plotly preview.
_DEFAULT_COLORS = (
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
)

_MARKER_SHAPES = {
    "circle": "circle",
    "square": "square",
    "diamond": "diamond",
    "cross": "x",
    "x": "plus",
    "triangle-up": "triangle",
    "triangle-down": "triangle",
}


class UnsupportedChartTypeError(TypeError):
    """A figure uses a chart/trace type the v2 host cannot render natively.

    Subclasses :class:`TypeError` so the worker folds it into the existing
    ``BadReturnType`` error code — the message names the offending type and
    the supported set, which is what reaches the user.
    """


# -----------------------------------------------------------------------------
# Figure detection (duck-typed — no plotly/matplotlib import)
# -----------------------------------------------------------------------------


def is_plotly_figure(value: Any) -> bool:
    """Whether ``value`` is a ``plotly.graph_objs.Figure`` (or subclass),
    detected without importing Plotly."""
    for cls in type(value).__mro__:
        module = cls.__module__ or ""
        if cls.__name__ == "Figure" and module.startswith("plotly."):
            return True
    return False


def is_matplotlib_figure(value: Any) -> bool:
    """Whether ``value`` is a ``matplotlib.figure.Figure`` (or subclass),
    detected without importing Matplotlib."""
    for cls in type(value).__mro__:
        module = cls.__module__ or ""
        if cls.__name__ == "Figure" and module.startswith("matplotlib"):
            return True
    return False


def convert_figure(value: Any) -> Any:
    """Convert a figure return value to its typed wire value.

    Plotly figures become :class:`ChartSpec`; Matplotlib figures become
    :class:`ChartImage`; anything else passes through unchanged. This is
    the single hook :func:`pyexcel.kernel.arrow_io.encode` calls on every
    outbound value.
    """
    if is_plotly_figure(value):
        return ChartSpec(json.dumps(plotly_figure_to_spec(value)))
    if is_matplotlib_figure(value):
        data, fmt = matplotlib_figure_to_image(value)
        return ChartImage(data, fmt)
    return value


# -----------------------------------------------------------------------------
# Plotly figure → chart spec
# -----------------------------------------------------------------------------


def plotly_figure_to_spec(figure: Any) -> dict:
    """Walk a Plotly figure and emit the chart-spec document (a plain dict
    ready for ``json.dumps``). Schema documented in the module docstring.

    Raises:
        UnsupportedChartTypeError: a trace's type has no native Excel
            counterpart in the supported set.
    """
    layout = figure.layout

    spec: dict = {
        "version": CHART_SPEC_VERSION,
        "chart_type": _chart_type_for(figure),
        "title": _layout_title(layout),
        "x_axis": _axis_spec(getattr(layout, "xaxis", None)),
        "y_axis": _axis_spec(getattr(layout, "yaxis", None)),
        "legend": _legend_spec(layout),
        "barmode": getattr(layout, "barmode", None) or "group",
        "traces": [
            _trace_spec(idx, trace)
            for idx, trace in enumerate(figure.data, start=1)
        ],
        "annotations": _annotations_spec(layout),
    }
    return spec


def _chart_type_for(figure: Any) -> str:
    """The chart-level base type, from the first trace (v1 behaviour)."""
    if len(figure.data) == 0:
        return "xy"
    first_type = (getattr(figure.data[0], "type", None) or "").lower()
    mapping = {
        "scatter": "xy",
        "scattergl": "xy",
        "bar": "bar",
        "histogram": "bar",
        "pie": "pie",
    }
    return mapping.get(first_type, "xy")


def _layout_title(layout: Any) -> str:
    title = getattr(layout, "title", None)
    text = getattr(title, "text", None) if title is not None else None
    return "" if text is None else str(text)


def _axis_spec(axis: Any) -> dict:
    if axis is None:
        return {"title": "", "min": None, "max": None, "log_scale": False}
    title = getattr(axis, "title", None)
    text = getattr(title, "text", None) if title is not None else None
    rng = getattr(axis, "range", None)
    return {
        "title": "" if text is None else str(text),
        "min": _json_value(rng[0]) if rng else None,
        "max": _json_value(rng[1]) if rng else None,
        "log_scale": getattr(axis, "type", None) == "log",
    }


def _legend_spec(layout: Any) -> dict:
    # Plotly's visibility switch is layout.showlegend (None = auto = on);
    # legend.visible also exists on newer Plotly versions — either turns
    # the legend off.
    show = getattr(layout, "showlegend", None)
    legend = getattr(layout, "legend", None)
    legend_visible = getattr(legend, "visible", None) if legend is not None else None
    visible = (show is not False) and (legend_visible is not False)

    orientation = getattr(legend, "orientation", None) if legend is not None else None
    position = "bottom" if orientation == "h" else "right"
    return {"visible": bool(visible), "position": position}


def _trace_spec(idx: int, trace: Any) -> dict:
    series_type = _series_type_for(trace)

    if series_type == "histogram":
        x, y, text = _histogram_data(trace)
        size = None
    elif series_type == "pie":
        # Plotly pie traces carry labels/values, not x/y.
        x = _json_list(getattr(trace, "labels", None)) or []
        y = _json_list(getattr(trace, "values", None)) or []
        text = _str_list(getattr(trace, "text", None))
        size = None
    elif series_type == "bar" :
        # Horizontal bars: Plotly stores x=values, y=categories; the spec
        # (like the v1 XML) always carries x=categories, y=values.
        x = _json_list(getattr(trace, "y", None)) or []
        y = _json_list(getattr(trace, "x", None)) or []
        text = _str_list(getattr(trace, "text", None))
        size = None
    else:
        x = _json_list(getattr(trace, "x", None)) or []
        y = _json_list(getattr(trace, "y", None)) or []
        text = _str_list(getattr(trace, "text", None))
        size = _bubble_sizes(trace) if series_type == "bubble" else None

    return {
        "id": idx,
        "x": x,
        "y": y,
        "text": text,
        "size": size,
        "style": _style_spec(idx, trace, series_type),
    }


def _series_type_for(trace: Any) -> str:
    trace_type = (getattr(trace, "type", None) or "").lower()

    if trace_type in ("scatter", "scattergl"):
        # Area: any Plotly fill mode turns the trace into an area series.
        if getattr(trace, "fill", None) not in (None, "none"):
            return "area"
        # Bubble: a per-point marker size array.
        marker = getattr(trace, "marker", None)
        if _is_sized_array(getattr(marker, "size", None) if marker is not None else None):
            return "bubble"
        mode = getattr(trace, "mode", None) or ""
        tokens = {m.strip().lower() for m in mode.split("+") if m.strip()}
        if "lines" in tokens and "markers" not in tokens:
            return "scatter_lines"
        if "lines" in tokens and "markers" in tokens:
            return "scatter_lines_markers"
        return "scatter"

    if trace_type == "bar":
        orientation = (getattr(trace, "orientation", None) or "v").lower()
        return "bar" if orientation == "h" else "column"

    if trace_type == "histogram":
        return "histogram"

    if trace_type == "pie":
        return "pie"

    if trace_type in _UNSUPPORTED_TRACE_TYPES:
        raise UnsupportedChartTypeError(
            f"Plotly trace type {trace_type!r} has no native Excel chart "
            f"counterpart; supported series types: "
            f"{', '.join(SUPPORTED_SERIES_TYPES)}"
        )

    raise UnsupportedChartTypeError(
        f"unknown Plotly trace type {trace_type!r}; supported series types: "
        f"{', '.join(SUPPORTED_SERIES_TYPES)}"
    )


def _is_sized_array(value: Any) -> bool:
    """True for numpy-array-like marker.size values (anything indexable
    with a length that isn't a string/number)."""
    if value is None or isinstance(value, (str, bytes, int, float)):
        return False
    try:
        len(value)
        return True
    except TypeError:
        return False


def _histogram_data(trace: Any) -> tuple:
    """Pre-bin a histogram trace: Excel gets bin midpoints as x, counts as
    y, and human-readable ``lo:hi`` range labels as text (used as category
    labels by the host). Mirrors the v1 converter, including honouring an
    explicit ``xbins`` start/end/size when fully specified."""
    import numpy as np  # hard kernel dep; local import keeps module load light

    raw = getattr(trace, "x", None)
    if raw is None or len(raw) == 0:
        return [], [], None

    values = np.asarray(raw, dtype=float)
    values = values[~np.isnan(values)]
    if values.size == 0:
        return [], [], None

    xbins = getattr(trace, "xbins", None)
    if (
        xbins is not None
        and getattr(xbins, "start", None) is not None
        and getattr(xbins, "end", None) is not None
        and getattr(xbins, "size", None) is not None
    ):
        bins = np.arange(xbins.start, xbins.end + xbins.size, xbins.size)
        counts, edges = np.histogram(values, bins=bins)
    else:
        counts, edges = np.histogram(values, bins="auto")

    midpoints = ((edges[:-1] + edges[1:]) / 2.0).tolist()
    labels = [
        f"{round(float(edges[i]), 4)}:{round(float(edges[i + 1]), 4)}"
        for i in range(len(edges) - 1)
    ]
    return midpoints, [int(c) for c in counts], labels


def _bubble_sizes(trace: Any) -> Optional[List[Any]]:
    marker = getattr(trace, "marker", None)
    if marker is None:
        return None
    return _json_list(getattr(marker, "size", None))


def _style_spec(idx: int, trace: Any, series_type: str) -> dict:
    name = getattr(trace, "name", None)

    axis_group = "secondary" if (getattr(trace, "yaxis", None) or "y") != "y" else "primary"

    visible = getattr(trace, "visible", True)
    # Plotly visible may be True / False / "legendonly"; only True renders.
    is_visible = visible is True or visible is None

    line = getattr(trace, "line", None)
    line_color = "#000000"
    line_dash = "solid"
    line_width: Optional[float] = None
    if line is not None:
        if getattr(line, "color", None):
            line_color = str(line.color)
        if getattr(line, "dash", None):
            line_dash = str(line.dash)
        if getattr(line, "width", None) is not None:
            line_width = float(line.width)

    marker = getattr(trace, "marker", None)
    marker_size: float = 6.0
    marker_color = "#000000"
    marker_shape = "circle"
    if series_type == "scatter_lines":
        marker_size = 0.0
    if marker is not None:
        size = getattr(marker, "size", None)
        if size is not None:
            if _is_sized_array(size):
                # Bubble sizes ride in data.size; the style slot gets a
                # representative scalar for non-bubble fallbacks.
                marker_size = float(size[0]) if len(size) > 0 else marker_size
            else:
                marker_size = float(size)
        color = getattr(marker, "color", None)
        if color is not None and not _is_sized_array(color):
            marker_color = str(color)
        symbol = getattr(marker, "symbol", None)
        if symbol is not None:
            marker_shape = _MARKER_SHAPES.get(str(symbol).lower(), "circle")

    fill_color = getattr(trace, "fillcolor", None) or None
    if fill_color is None and marker is not None:
        color = getattr(marker, "color", None)
        if color is not None and not _is_sized_array(color):
            fill_color = str(color)
    if fill_color is None and series_type in ("bar", "column", "histogram"):
        fill_color = _DEFAULT_COLORS[(idx - 1) % len(_DEFAULT_COLORS)]

    fill_opacity = getattr(trace, "opacity", None)

    return {
        "series_type": series_type,
        "name": str(name) if name is not None else f"Series {idx}",
        "axis_group": axis_group,
        "visible": bool(is_visible),
        "line": {"color": line_color, "dash": line_dash, "width": line_width},
        "marker": {"size": marker_size, "color": marker_color, "shape": marker_shape},
        "fill_color": fill_color,
        "fill_opacity": _json_value(fill_opacity),
    }


def _annotations_spec(layout: Any) -> List[dict]:
    """Vertical/horizontal line shapes become event_line / threshold
    annotations; every other shape kind is skipped (v1 behaviour)."""
    shapes = getattr(layout, "shapes", None) or ()
    annotations: List[dict] = []
    for idx, shape in enumerate(shapes, start=1):
        if getattr(shape, "type", None) != "line":
            continue
        x0, x1 = getattr(shape, "x0", None), getattr(shape, "x1", None)
        y0, y1 = getattr(shape, "y0", None), getattr(shape, "y1", None)
        if x0 is not None and x0 == x1:
            ann_type, axis, value = "event_line", "x", _json_value(x0)
        elif y0 is not None and y0 == y1:
            ann_type, axis, value = "threshold", "y", _json_value(y0)
        else:
            continue

        line = getattr(shape, "line", None)
        style = {
            "color": str(line.color) if line is not None and getattr(line, "color", None) else None,
            "width": _json_value(getattr(line, "width", None)) if line is not None else None,
            "dash": str(getattr(line, "dash", None) or "solid") if line is not None else "solid",
            "opacity": _json_value(getattr(shape, "opacity", None)),
        }
        label = getattr(shape, "name", None)
        annotations.append(
            {
                "id": idx,
                "type": ann_type,
                "axis": axis,
                "value": value,
                "style": style,
                "label": str(label) if label else None,
            }
        )
    return annotations


# -----------------------------------------------------------------------------
# Matplotlib figure → image
# -----------------------------------------------------------------------------


def matplotlib_figure_to_image(figure: Any) -> tuple:
    """Render a Matplotlib figure to image bytes: ``(data, format)``.

    SVG first (vector — crisp at any worksheet zoom); PNG fallback if the
    SVG backend fails for any reason. If both backends fail the PNG error
    propagates as :class:`TypeError` so the worker reports BadReturnType
    with the underlying render error in the message.
    """
    import io

    svg_buf = io.BytesIO()
    try:
        figure.savefig(svg_buf, format="svg", bbox_inches="tight")
        return svg_buf.getvalue(), "svg"
    except Exception:  # noqa: BLE001 — any SVG failure falls back to PNG
        pass

    png_buf = io.BytesIO()
    try:
        figure.savefig(png_buf, format="png", dpi=150, bbox_inches="tight")
        return png_buf.getvalue(), "png"
    except Exception as exc:  # noqa: BLE001
        raise TypeError(
            f"could not render the Matplotlib figure to SVG or PNG: {exc}"
        ) from exc


# -----------------------------------------------------------------------------
# JSON value normalisation
# -----------------------------------------------------------------------------


def _json_value(value: Any) -> Any:
    """Coerce one figure value to a JSON-representable primitive.

    Numbers stay numbers (NaN → null — JSON has no NaN), datetimes become
    ISO-8601 strings, numpy scalars unwrap, everything else stringifies.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, str)):
        return value
    if isinstance(value, float):
        return None if math.isnan(value) else value
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()
    item = getattr(value, "item", None)
    if callable(item):
        # numpy scalar (np.float64, np.int64, np.datetime64 via .item(), …)
        try:
            return _json_value(item())
        except (TypeError, ValueError):
            pass
    return str(value)


def _json_list(values: Any) -> Optional[List[Any]]:
    if values is None:
        return None
    return [_json_value(v) for v in values]


def _str_list(values: Any) -> Optional[List[str]]:
    if values is None:
        return None
    out = ["" if v is None else str(v) for v in values]
    return out if out else None
