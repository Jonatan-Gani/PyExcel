"""Tests for ``pyexcel.kernel.chart`` — the chart-spec contract.

Covers the Plotly figure → JSON chart-spec traversal (the v2 port of
v1's ``PlotlyToExcelXMLConverter``), the Matplotlib figure → image
rendering, duck-typed figure detection, the typed wire values
(:class:`ChartSpec` / :class:`ChartImage`), and the Arrow encode/decode
round-trip for both new shapes.

Plotly and Matplotlib are declared kernel dependencies (requirements.txt)
but the kernel itself never imports them — tests that need a real figure
``importorskip`` so an Arrow-only environment still runs the rest.
"""

from __future__ import annotations

import json

import pytest

from pyexcel.kernel.arrow_io import Shape, decode, encode
from pyexcel.kernel.chart import (
    SUPPORTED_SERIES_TYPES,
    UnsupportedChartTypeError,
    convert_figure,
    is_matplotlib_figure,
    is_plotly_figure,
    matplotlib_figure_to_image,
    plotly_figure_to_spec,
)
from pyexcel.kernel.types import ChartImage, ChartSpec

go = pytest.importorskip("plotly.graph_objects")


def spec_of(fig) -> dict:
    """Build + JSON-round-trip the spec so tests assert on exactly what
    the host will parse (catches non-JSON-serialisable leaks)."""
    return json.loads(json.dumps(plotly_figure_to_spec(fig)))


# -----------------------------------------------------------------------------
# Wire types
# -----------------------------------------------------------------------------


class TestWireTypes:
    def test_chart_spec_requires_string(self):
        with pytest.raises(TypeError):
            ChartSpec(b"{}")  # type: ignore[arg-type]

    def test_chart_spec_rejects_empty(self):
        with pytest.raises(ValueError):
            ChartSpec("   ")

    def test_chart_image_requires_bytes(self):
        with pytest.raises(TypeError):
            ChartImage("not bytes", "png")  # type: ignore[arg-type]

    def test_chart_image_rejects_empty_data(self):
        with pytest.raises(ValueError):
            ChartImage(b"", "png")

    def test_chart_image_rejects_unknown_format(self):
        with pytest.raises(ValueError):
            ChartImage(b"\x89PNG", "jpeg")

    def test_chart_image_accepts_svg_and_png(self):
        assert ChartImage(b"<svg/>", "svg").format == "svg"
        assert ChartImage(b"\x89PNG", "png").format == "png"


# -----------------------------------------------------------------------------
# Figure detection
# -----------------------------------------------------------------------------


class TestDetection:
    def test_plotly_figure_detected(self):
        assert is_plotly_figure(go.Figure())

    def test_plotly_figure_is_not_matplotlib(self):
        assert not is_matplotlib_figure(go.Figure())

    def test_plain_values_are_not_figures(self):
        for value in (None, 42, "Figure", {"a": 1}, [1, 2]):
            assert not is_plotly_figure(value)
            assert not is_matplotlib_figure(value)

    def test_convert_figure_passes_non_figures_through(self):
        sentinel = object()
        assert convert_figure(sentinel) is sentinel

    def test_convert_figure_wraps_plotly_as_chart_spec(self):
        converted = convert_figure(go.Figure(data=[go.Scatter(x=[1], y=[2])]))
        assert isinstance(converted, ChartSpec)
        assert json.loads(converted.json)["version"] == 1


# -----------------------------------------------------------------------------
# Spec document: meta
# -----------------------------------------------------------------------------


class TestChartMeta:
    def test_minimal_scatter_spec(self):
        spec = spec_of(go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4])]))
        assert spec["version"] == 1
        assert spec["chart_type"] == "xy"
        assert spec["title"] == ""
        assert len(spec["traces"]) == 1

    def test_title_carried(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.update_layout(title="Quarterly Revenue")
        assert spec_of(fig)["title"] == "Quarterly Revenue"

    def test_empty_figure_defaults_to_xy(self):
        spec = spec_of(go.Figure())
        assert spec["chart_type"] == "xy"
        assert spec["traces"] == []

    def test_bar_figure_chart_type(self):
        spec = spec_of(go.Figure(data=[go.Bar(x=["a", "b"], y=[1, 2])]))
        assert spec["chart_type"] == "bar"

    def test_pie_figure_chart_type(self):
        spec = spec_of(go.Figure(data=[go.Pie(labels=["a", "b"], values=[1, 2])]))
        assert spec["chart_type"] == "pie"

    def test_axis_titles_and_range(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.update_layout(
            xaxis=dict(title="Time", range=[0, 10]),
            yaxis=dict(title="Price", range=[-5, 5]),
        )
        spec = spec_of(fig)
        assert spec["x_axis"] == {
            "title": "Time", "min": 0, "max": 10, "log_scale": False,
        }
        assert spec["y_axis"] == {
            "title": "Price", "min": -5, "max": 5, "log_scale": False,
        }

    def test_axis_defaults_when_unset(self):
        spec = spec_of(go.Figure(data=[go.Scatter(x=[1], y=[2])]))
        assert spec["x_axis"] == {
            "title": "", "min": None, "max": None, "log_scale": False,
        }

    def test_log_axis(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.update_layout(yaxis_type="log")
        assert spec_of(fig)["y_axis"]["log_scale"] is True

    def test_legend_default_visible_right(self):
        spec = spec_of(go.Figure(data=[go.Scatter(x=[1], y=[2])]))
        assert spec["legend"] == {"visible": True, "position": "right"}

    def test_legend_hidden_via_showlegend(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.update_layout(showlegend=False)
        assert spec_of(fig)["legend"]["visible"] is False

    def test_legend_horizontal_maps_to_bottom(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.update_layout(legend_orientation="h")
        assert spec_of(fig)["legend"]["position"] == "bottom"

    def test_barmode_default_group(self):
        spec = spec_of(go.Figure(data=[go.Bar(x=["a"], y=[1])]))
        assert spec["barmode"] == "group"

    def test_barmode_stack_carried(self):
        fig = go.Figure(data=[go.Bar(x=["a"], y=[1])])
        fig.update_layout(barmode="stack")
        assert spec_of(fig)["barmode"] == "stack"


# -----------------------------------------------------------------------------
# Spec document: traces
# -----------------------------------------------------------------------------


class TestTraces:
    def test_trace_ids_are_one_based_and_unique(self):
        fig = go.Figure(
            data=[
                go.Scatter(x=[1], y=[2]),
                go.Scatter(x=[3], y=[4]),
                go.Scatter(x=[5], y=[6]),
            ]
        )
        ids = [t["id"] for t in spec_of(fig)["traces"]]
        assert ids == [1, 2, 3]

    def test_scatter_markers_only(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4], mode="markers")])
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "scatter"
        assert trace["x"] == [1, 2]
        assert trace["y"] == [3, 4]

    def test_scatter_lines_only(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], mode="lines")])
        assert spec_of(fig)["traces"][0]["style"]["series_type"] == "scatter_lines"

    def test_scatter_lines_and_markers(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], mode="lines+markers")])
        assert (
            spec_of(fig)["traces"][0]["style"]["series_type"]
            == "scatter_lines_markers"
        )

    def test_scatter_with_fill_becomes_area(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4], fill="tozeroy")])
        assert spec_of(fig)["traces"][0]["style"]["series_type"] == "area"

    def test_vertical_bar_is_column(self):
        fig = go.Figure(data=[go.Bar(x=["a", "b"], y=[1, 2])])
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "column"
        assert trace["x"] == ["a", "b"]
        assert trace["y"] == [1, 2]

    def test_horizontal_bar_swaps_axes(self):
        # Plotly horizontal bars store x=values, y=categories; the spec
        # always carries x=categories, y=values.
        fig = go.Figure(data=[go.Bar(x=[1, 2], y=["a", "b"], orientation="h")])
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "bar"
        assert trace["x"] == ["a", "b"]
        assert trace["y"] == [1, 2]

    def test_pie_labels_and_values(self):
        fig = go.Figure(data=[go.Pie(labels=["a", "b", "c"], values=[1, 2, 3])])
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "pie"
        assert trace["x"] == ["a", "b", "c"]
        assert trace["y"] == [1, 2, 3]

    def test_bubble_when_marker_size_is_array(self):
        fig = go.Figure(
            data=[go.Scatter(x=[1, 2], y=[3, 4], marker=dict(size=[10, 20]))]
        )
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "bubble"
        assert trace["size"] == [10, 20]

    def test_histogram_is_prebinned(self):
        values = [1.0, 1.1, 1.2, 5.0, 5.1, 5.2, 9.0, 9.1, 9.2]
        fig = go.Figure(data=[go.Histogram(x=values)])
        trace = spec_of(fig)["traces"][0]
        assert trace["style"]["series_type"] == "histogram"
        # Pre-binned: counts sum to the sample count, labels are lo:hi pairs.
        assert sum(trace["y"]) == len(values)
        assert len(trace["x"]) == len(trace["y"]) == len(trace["text"])
        assert all(":" in label for label in trace["text"])

    def test_histogram_explicit_bins(self):
        fig = go.Figure(
            data=[go.Histogram(x=[1, 2, 3, 4], xbins=dict(start=0, end=4, size=2))]
        )
        trace = spec_of(fig)["traces"][0]
        assert trace["y"] == [1, 3]  # [0,2) → {1}; [2,4] → {2,3,4}

    def test_empty_histogram_yields_empty_data(self):
        fig = go.Figure(data=[go.Histogram(x=[])])
        trace = spec_of(fig)["traces"][0]
        assert trace["x"] == []
        assert trace["y"] == []

    def test_text_labels_carried(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4], text=["p", "q"])])
        assert spec_of(fig)["traces"][0]["text"] == ["p", "q"]

    def test_none_values_become_null(self):
        fig = go.Figure(data=[go.Scatter(x=[1, None, 3], y=[4, 5, None])])
        trace = spec_of(fig)["traces"][0]
        assert trace["x"] == [1, None, 3]
        assert trace["y"] == [4, 5, None]

    def test_numpy_data_serialises(self):
        np = pytest.importorskip("numpy")
        fig = go.Figure(
            data=[go.Scatter(x=np.array([1.5, 2.5]), y=np.array([3, 4]))]
        )
        trace = spec_of(fig)["traces"][0]
        assert trace["x"] == [1.5, 2.5]
        assert trace["y"] == [3, 4]

    def test_datetime_x_serialises_as_iso(self):
        import datetime as dt

        fig = go.Figure(
            data=[go.Scatter(x=[dt.date(2026, 1, 2), dt.date(2026, 1, 3)], y=[1, 2])]
        )
        assert spec_of(fig)["traces"][0]["x"] == ["2026-01-02", "2026-01-03"]


# -----------------------------------------------------------------------------
# Spec document: per-trace style
# -----------------------------------------------------------------------------


class TestTraceStyle:
    def test_default_name_when_unnamed(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        assert spec_of(fig)["traces"][0]["style"]["name"] == "Series 1"

    def test_explicit_name_carried(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], name="prices")])
        assert spec_of(fig)["traces"][0]["style"]["name"] == "prices"

    def test_line_styling(self):
        fig = go.Figure(
            data=[
                go.Scatter(
                    x=[1], y=[2], mode="lines",
                    line=dict(color="#ff0000", dash="dash", width=2.5),
                )
            ]
        )
        line = spec_of(fig)["traces"][0]["style"]["line"]
        assert line == {"color": "#ff0000", "dash": "dash", "width": 2.5}

    def test_line_defaults(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        line = spec_of(fig)["traces"][0]["style"]["line"]
        assert line == {"color": "#000000", "dash": "solid", "width": None}

    def test_marker_styling(self):
        fig = go.Figure(
            data=[
                go.Scatter(
                    x=[1], y=[2], mode="markers",
                    marker=dict(size=12, color="#00ff00", symbol="square"),
                )
            ]
        )
        marker = spec_of(fig)["traces"][0]["style"]["marker"]
        assert marker == {"size": 12.0, "color": "#00ff00", "shape": "square"}

    def test_marker_shape_mapping(self):
        for plotly_symbol, expected in [
            ("circle", "circle"),
            ("square", "square"),
            ("diamond", "diamond"),
            ("cross", "x"),
            ("x", "plus"),
            ("triangle-up", "triangle"),
            ("triangle-down", "triangle"),
            ("hexagon", "circle"),  # unmapped symbol falls back to circle
        ]:
            fig = go.Figure(
                data=[go.Scatter(x=[1], y=[2], marker=dict(symbol=plotly_symbol))]
            )
            shape = spec_of(fig)["traces"][0]["style"]["marker"]["shape"]
            assert shape == expected, plotly_symbol

    def test_scatter_lines_suppresses_markers(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], mode="lines")])
        assert spec_of(fig)["traces"][0]["style"]["marker"]["size"] == 0.0

    def test_secondary_axis_group(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], yaxis="y2")])
        assert spec_of(fig)["traces"][0]["style"]["axis_group"] == "secondary"

    def test_primary_axis_group_default(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        assert spec_of(fig)["traces"][0]["style"]["axis_group"] == "primary"

    def test_column_without_color_gets_palette_default(self):
        fig = go.Figure(
            data=[go.Bar(x=["a"], y=[1]), go.Bar(x=["a"], y=[2])]
        )
        traces = spec_of(fig)["traces"]
        assert traces[0]["style"]["fill_color"] == "#1f77b4"
        assert traces[1]["style"]["fill_color"] == "#ff7f0e"

    def test_marker_color_used_as_fill(self):
        fig = go.Figure(data=[go.Bar(x=["a"], y=[1], marker=dict(color="#123456"))])
        assert spec_of(fig)["traces"][0]["style"]["fill_color"] == "#123456"

    def test_scatter_without_fill_color_stays_null(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        assert spec_of(fig)["traces"][0]["style"]["fill_color"] is None

    def test_opacity_carried(self):
        fig = go.Figure(data=[go.Bar(x=["a"], y=[1], opacity=0.4)])
        assert spec_of(fig)["traces"][0]["style"]["fill_opacity"] == 0.4

    def test_legendonly_trace_marked_invisible(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2], visible="legendonly")])
        assert spec_of(fig)["traces"][0]["style"]["visible"] is False

    def test_series_types_are_all_in_supported_set(self):
        figs = [
            go.Figure(data=[go.Scatter(x=[1], y=[2], mode="markers")]),
            go.Figure(data=[go.Scatter(x=[1], y=[2], mode="lines")]),
            go.Figure(data=[go.Scatter(x=[1], y=[2], mode="lines+markers")]),
            go.Figure(data=[go.Scatter(x=[1], y=[2], fill="tozeroy")]),
            go.Figure(data=[go.Scatter(x=[1], y=[2], marker=dict(size=[5]))]),
            go.Figure(data=[go.Bar(x=["a"], y=[1])]),
            go.Figure(data=[go.Bar(x=[1], y=["a"], orientation="h")]),
            go.Figure(data=[go.Pie(labels=["a"], values=[1])]),
            go.Figure(data=[go.Histogram(x=[1, 2, 3])]),
        ]
        for fig in figs:
            for trace in spec_of(fig)["traces"]:
                assert trace["style"]["series_type"] in SUPPORTED_SERIES_TYPES


# -----------------------------------------------------------------------------
# Unsupported chart types — explicit, surfaced
# -----------------------------------------------------------------------------


class TestUnsupportedTypes:
    def test_box_trace_raises_with_message(self):
        fig = go.Figure(data=[go.Box(y=[1, 2, 3])])
        with pytest.raises(UnsupportedChartTypeError, match="box"):
            plotly_figure_to_spec(fig)

    def test_heatmap_trace_raises(self):
        fig = go.Figure(data=[go.Heatmap(z=[[1, 2], [3, 4]])])
        with pytest.raises(UnsupportedChartTypeError, match="heatmap"):
            plotly_figure_to_spec(fig)

    def test_error_message_names_supported_set(self):
        fig = go.Figure(data=[go.Box(y=[1])])
        with pytest.raises(UnsupportedChartTypeError, match="scatter_lines"):
            plotly_figure_to_spec(fig)

    def test_unsupported_error_is_a_type_error(self):
        # The worker maps TypeError from encode to the BadReturnType error
        # code; UnsupportedChartTypeError must ride that path.
        assert issubclass(UnsupportedChartTypeError, TypeError)


# -----------------------------------------------------------------------------
# Annotations (layout.shapes → event_line / threshold)
# -----------------------------------------------------------------------------


class TestAnnotations:
    def test_vertical_line_becomes_event_line(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4])])
        fig.add_shape(type="line", x0=1.5, x1=1.5, y0=0, y1=1)
        anns = spec_of(fig)["annotations"]
        assert len(anns) == 1
        assert anns[0]["type"] == "event_line"
        assert anns[0]["axis"] == "x"
        assert anns[0]["value"] == 1.5

    def test_horizontal_line_becomes_threshold(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4])])
        fig.add_shape(
            type="line", x0=0, x1=2, y0=3.5, y1=3.5,
            line=dict(color="#ff0000", width=2, dash="dot"),
        )
        anns = spec_of(fig)["annotations"]
        assert anns[0]["type"] == "threshold"
        assert anns[0]["axis"] == "y"
        assert anns[0]["value"] == 3.5
        assert anns[0]["style"]["color"] == "#ff0000"
        assert anns[0]["style"]["width"] == 2
        assert anns[0]["style"]["dash"] == "dot"

    def test_rect_shapes_skipped(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.add_shape(type="rect", x0=0, x1=1, y0=0, y1=1)
        assert spec_of(fig)["annotations"] == []

    def test_diagonal_line_skipped(self):
        fig = go.Figure(data=[go.Scatter(x=[1], y=[2])])
        fig.add_shape(type="line", x0=0, x1=1, y0=0, y1=1)
        assert spec_of(fig)["annotations"] == []


# -----------------------------------------------------------------------------
# Matplotlib → image
# -----------------------------------------------------------------------------


class TestMatplotlibImage:
    @pytest.fixture()
    def mpl_figure(self):
        matplotlib = pytest.importorskip("matplotlib")
        matplotlib.use("Agg")  # headless backend for CI
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots()
        ax.plot([1, 2, 3], [4, 5, 6])
        yield fig
        plt.close(fig)

    def test_detected_as_matplotlib(self, mpl_figure):
        assert is_matplotlib_figure(mpl_figure)
        assert not is_plotly_figure(mpl_figure)

    def test_renders_svg(self, mpl_figure):
        data, fmt = matplotlib_figure_to_image(mpl_figure)
        assert fmt == "svg"
        assert b"<svg" in data[:1024]

    def test_png_fallback_when_svg_fails(self, mpl_figure, monkeypatch):
        real_savefig = mpl_figure.savefig

        def failing_svg(buf, *args, **kwargs):
            if kwargs.get("format") == "svg":
                raise RuntimeError("svg backend unavailable")
            return real_savefig(buf, *args, **kwargs)

        monkeypatch.setattr(mpl_figure, "savefig", failing_svg)
        data, fmt = matplotlib_figure_to_image(mpl_figure)
        assert fmt == "png"
        assert data[:8] == b"\x89PNG\r\n\x1a\n"

    def test_raises_type_error_when_both_backends_fail(self, mpl_figure, monkeypatch):
        def always_failing(buf, *args, **kwargs):
            raise RuntimeError("no render backends")

        monkeypatch.setattr(mpl_figure, "savefig", always_failing)
        with pytest.raises(TypeError, match="SVG or PNG"):
            matplotlib_figure_to_image(mpl_figure)

    def test_convert_figure_wraps_as_chart_image(self, mpl_figure):
        converted = convert_figure(mpl_figure)
        assert isinstance(converted, ChartImage)
        assert converted.format == "svg"


# -----------------------------------------------------------------------------
# Arrow wire round-trip for the new shapes
# -----------------------------------------------------------------------------


class TestArrowRoundTrip:
    def test_plotly_figure_encodes_as_chart_shape(self):
        import pyarrow as pa
        import pyarrow.ipc as ipc

        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4])])
        buf = encode(fig)

        reader = ipc.open_stream(pa.BufferReader(buf))
        metadata = reader.schema.metadata or {}
        assert metadata.get(b"pyexcel-shape") == Shape.CHART.value

    def test_plotly_figure_round_trips_to_chart_spec(self):
        fig = go.Figure(data=[go.Scatter(x=[1, 2], y=[3, 4], name="prices")])
        decoded = decode(encode(fig))
        assert isinstance(decoded, ChartSpec)
        doc = json.loads(decoded.json)
        assert doc["version"] == 1
        assert doc["traces"][0]["style"]["name"] == "prices"

    def test_chart_spec_value_round_trips_directly(self):
        spec = ChartSpec('{"version": 1, "traces": []}')
        decoded = decode(encode(spec))
        assert decoded == spec

    def test_chart_image_round_trips_with_format(self):
        image = ChartImage(b"<svg>chart</svg>", "svg")
        decoded = decode(encode(image))
        assert isinstance(decoded, ChartImage)
        assert decoded.data == image.data
        assert decoded.format == "svg"

    def test_png_image_round_trips(self):
        image = ChartImage(b"\x89PNG\r\n\x1a\nrest", "png")
        decoded = decode(encode(image))
        assert decoded.format == "png"
        assert decoded.data == image.data

    def test_image_shape_metadata(self):
        import pyarrow as pa
        import pyarrow.ipc as ipc

        buf = encode(ChartImage(b"\x89PNG", "png"))
        reader = ipc.open_stream(pa.BufferReader(buf))
        metadata = reader.schema.metadata or {}
        assert metadata.get(b"pyexcel-shape") == Shape.IMAGE.value
        field_md = reader.schema.field(0).metadata or {}
        assert field_md.get(b"pyexcel-image-format") == b"png"

    def test_unsupported_figure_type_raises_at_encode(self):
        fig = go.Figure(data=[go.Box(y=[1, 2, 3])])
        with pytest.raises(TypeError, match="box"):
            encode(fig)
