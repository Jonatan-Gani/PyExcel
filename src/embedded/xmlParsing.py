from __future__ import annotations

import os
import re

import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import datetime as dt

from datetime import datetime
from plotly.graph_objs import Figure
from lxml import etree as LET

from typing import Union, Sequence, Mapping, Optional, Dict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field


_ANCHOR_RE = re.compile(r'^\$?[A-Z]+\$?\d+$')

@dataclass(frozen=True)
class ExcelFormula:
    mode: str                   # "a1" or "r1c1"
    a1: Optional[str] = None
    anchor: Optional[str] = None
    r1c1: Optional[str] = None
    # arbitrary optional attributes for future use
    attrs: Dict[str, str] = field(default_factory=dict)


def excel_formula(a1: str, anchor: str = None, **options) -> ExcelFormula:
    """
    Create an A1-mode ExcelFormula without requiring an anchor.
    If an anchor is provided, it is validated then ignored (deprecated).
    """
    # Minimal validation
    if not isinstance(a1, str) or not a1.startswith("="):
        raise ValueError("a1 must start with '='.")

    # Back-compat: accept but ignore anchor
    if anchor is not None:
        if not isinstance(anchor, str) or not _ANCHOR_RE.match(anchor):
            raise ValueError("anchor must be a valid A1 like 'E7' or '$E$7'.")
        # intentionally ignored

    # Normalize common options to strings
    def b(x): return "1" if bool(x) else "0"
    normalized: Dict[str, str] = {}
    for k, v in options.items():
        if v is None:
            continue
        if k in {"spill", "volatile", "protect"}:
            normalized[k] = b(v)
        else:
            normalized[k] = str(v)

    # NOTE: no 'anchor' in the payload anymore
    return ExcelFormula(mode="a1", a1=a1, attrs=normalized)

# ----------------------------
# XML to and from DataFrames
# ----------------------------

def read_xml(path: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    tables: dict[str, pd.DataFrame] = {}

    # State for the current <table>
    table_name = None
    col_names: list[str] | None = None
    col_types: list[str] | None = None
    col_buffers: list[list[str]] | None = None

    # Stream parse: handle 'end' events
    for event, elem in LET.iterparse(path, events=("end",)):
        tag = elem.tag

        # -----------------------------
        # TABLE: <columns>
        # -----------------------------
        if tag == "columns" and elem.getparent() is not None and elem.getparent().tag == "table":
            cols = elem.findall("col")
            col_names = [c.get("name", "") for c in cols]
            col_types = [c.get("type", "string") for c in cols]
            col_buffers = [[] for _ in col_names]

            elem.clear()
            parent = elem.getparent()
            if parent is not None:
                while elem.getprevious() is not None:
                    del parent[0]

        # -----------------------------
        # TABLE: <row>
        # -----------------------------
        elif tag == "row":
            if col_buffers is None or col_names is None:
                raise ValueError("Encountered <row> before <columns>.")

            cells = elem.findall("col")
            n = len(col_names)

            row_vals = [(c.text or "") for c in cells]
            if len(row_vals) < n:
                row_vals.extend([""] * (n - len(row_vals)))
            elif len(row_vals) > n:
                row_vals = row_vals[:n]

            for j, v in enumerate(row_vals):
                col_buffers[j].append(v)

            elem.clear()
            parent = elem.getparent()
            if parent is not None:
                while elem.getprevious() is not None:
                    del parent[0]

        # -----------------------------
        # TABLE: <table>
        # -----------------------------
        elif tag == "table":
            table_name = elem.get("name", "")
            if not table_name:
                raise ValueError("Table missing 'name' attribute")
            if col_names is None or col_types is None or col_buffers is None:
                raise ValueError(f"Malformed table '{table_name}': missing columns/rows")

            df_dict: dict[str, pd.Series] = {}

            for cname, ctype, raw in zip(col_names, col_types, col_buffers):
                s = pd.Series(raw, copy=False)

                if ctype == "int":
                    empty_mask = s.eq("")
                    num = pd.to_numeric(s.mask(empty_mask), errors="coerce")
                    num[empty_mask] = 0
                    try:
                        df_dict[cname] = num.astype("Int64")
                    except TypeError:
                        print(f"[read_xml] Column '{cname}' contains non-integers, keeping as float")
                        print(num.head(20).to_list())
                        df_dict[cname] = num

                elif ctype == "float":
                    empty_mask = s.eq("")
                    num = pd.to_numeric(s.mask(empty_mask), errors="coerce")
                    num[empty_mask] = 0.0
                    df_dict[cname] = num

                elif ctype == "bool":
                    df_dict[cname] = s.str.lower().eq("true")

                elif ctype == "timestamp":
                    df_dict[cname] = pd.to_datetime(s, errors="coerce", utc=True)

                elif ctype == "blank":
                    df_dict[cname] = pd.Series(pd.NA, index=s.index)

                else:
                    df_dict[cname] = s.astype(str)

            tables[table_name] = pd.DataFrame(df_dict, columns=col_names)

            # Reset state
            elem.clear()
            parent = elem.getparent()
            if parent is not None:
                while elem.getprevious() is not None:
                    del parent[0]
            table_name = None
            col_names = None
            col_types = None
            col_buffers = None

        # -----------------------------
        # NEW: <list>
        # -----------------------------
        elif tag == "list":
            name = elem.get("name")
            if not name:
                raise ValueError("<list> missing required 'name' attribute")

            items = [(child.text or "") for child in elem.findall("item")]
            result[name] = items

            elem.clear()
            parent = elem.getparent()
            if parent is not None:
                while elem.getprevious() is not None:
                    del parent[0]

        # -----------------------------
        # NEW: <value>
        # -----------------------------
        elif tag == "value":
            name = elem.get("name")
            dtype = elem.get("datatype", "string")
            if not name:
                raise ValueError("<value> missing required 'name' attribute")

            raw = elem.text or ""

            if dtype == "int":
                out = int(raw)
            elif dtype == "decimal":
                out = float(raw)
            elif dtype == "bool":
                out = raw.lower() == "true"
            elif dtype == "timestamp":
                out = pd.to_datetime(raw, errors="coerce", utc=True)
            else:
                out = raw

            result[name] = out

            elem.clear()
            parent = elem.getparent()
            if parent is not None:
                while elem.getprevious() is not None:
                    del parent[0]

    # Merge tables and simple values/lists
    result.update(tables)
    return result



def write_xml(path: str,
              tables: Union[pd.DataFrame,
                            Sequence[pd.DataFrame],
                            Mapping[str, pd.DataFrame]]):
    try:
        ExcelFormulaBase = ExcelFormula  # type: ignore[name-defined]
    except NameError:
        ExcelFormulaBase = ()  # type: ignore[assignment]

    def infer_column_type(s: pd.Series) -> str:
        ss = s.dropna()
        if ss.empty:
            return "blank"
        ts = ss.map(type).unique()
        if all(t is bool for t in ts):
            return "bool"
        if all(issubclass(t, (int, np.integer)) for t in ts):
            return "int"
        if all(issubclass(t, (float, int, np.floating, np.integer)) for t in ts):
            return "float"
        if all(isinstance(v, (datetime, pd.Timestamp)) for v in ss):
            return "date"
        return "string"

    def is_formula_obj(v) -> bool:
        if isinstance(v, ExcelFormulaBase):
            return True
        if isinstance(v, dict) and ("a1" in v or "r1c1" in v):
            return True
        return False

    def extract_meta(s: pd.Series) -> Optional[Dict[str, str]]:
        for v in s:
            if isinstance(v, ExcelFormulaBase):
                meta: Dict[str, str] = {"mode": v.mode}
                if v.mode == "a1":
                    meta["a1"] = v.a1  # type: ignore[arg-type]
                elif v.mode == "r1c1":
                    meta["r1c1"] = v.r1c1  # type: ignore[arg-type]
                for k, vv in (getattr(v, "attrs", None) or {}).items():
                    meta[k] = vv if isinstance(vv, str) else str(vv)
                return meta
        for v in s:
            if isinstance(v, dict) and ("a1" in v or "r1c1" in v):
                meta = {k: (vv if isinstance(vv, str) else str(vv)) for k, vv in v.items()}
                if "mode" not in meta:
                    meta["mode"] = "a1" if "a1" in meta else "r1c1"
                if meta.get("mode") == "a1":
                    meta.pop("anchor", None)
                return meta
        for v in s:
            if isinstance(v, str) and v.startswith("="):
                return {"mode": "r1c1", "r1c1": v}
        return None

    def serialize_value(v):
        if v is None or pd.isna(v):
            return None
        if is_formula_obj(v):
            return None
        if isinstance(v, str):
            if v.startswith("=") or v == "":
                return None
            return v
        if isinstance(v, bool):
            return str(v).lower()
        if isinstance(v, (int, float, np.integer, np.floating)):
            return str(v)
        if isinstance(v, (datetime, pd.Timestamp)):
            return v.strftime("%Y-%m-%d")
        return str(v)

    if isinstance(tables, pd.DataFrame):
        name_df_pairs = [("df1", tables)]
    elif isinstance(tables, Mapping):
        name_df_pairs = [(str(name), df) for name, df in tables.items()]
    else:
        name_df_pairs = [(f"df{i+1}", df) for i, df in enumerate(tables)]

    root = LET.Element("data")

    for name, df in name_df_pairs:
        table_el = LET.SubElement(root, "table", attrib={"name": name})

        cols = LET.SubElement(table_el, "columns")
        for col in df.columns:
            s = df[col]
            attrib: Dict[str, str] = {"name": str(col), "type": infer_column_type(s)}
            meta = extract_meta(s)
            if meta:
                meta.pop("anchor", None)
                attrib.update({k: v if isinstance(v, str) else str(v) for k, v in meta.items()})
            LET.SubElement(cols, "col", attrib=attrib)

        rows = LET.SubElement(table_el, "rows")
        for _, r in df.iterrows():
            row = LET.SubElement(rows, "row")
            for v in r:
                cell = LET.SubElement(row, "col")
                sv = serialize_value(v)
                if sv is not None:
                    cell.text = sv

    tree = LET.ElementTree(root)
    tree.write(path, encoding="utf-8", xml_declaration=True, pretty_print=False)


class PlotlyToExcelXMLConverter:
    SCHEMA_VERSION = "2.0"

    def __init__(self, figure: Figure):
        if not isinstance(figure, Figure):
            raise TypeError("Input must be a plotly.graph_objs.Figure")
        self.figure = figure

    def _safe_text(self, value) -> str:
        return "" if value is None else str(value)

    def _comma_join(self, values) -> str:
        if values is None:
            return ""
        return ",".join(str(v) for v in values)

    def _map_plotly_type_to_chart_type(self, plotly_type: str) -> str:
        mapping = {
            "scatter": "xy",
            "scattergl": "xy",
            "bar": "bar",
            "line": "line",
            "area": "area",
            "pie": "pie",
        }
        return mapping.get(plotly_type.lower(), "xy")

    def _map_plotly_type_to_series_type(self, trace) -> str:
        trace_type = getattr(trace, "type", "").lower()

        if trace_type in ("scatter", "scattergl"):
            mode = getattr(trace, "mode", "") or ""
            mode_tokens = {m.strip().lower() for m in mode.split("+")}
            if "lines" in mode_tokens and "markers" not in mode_tokens:
                return "scatter_lines"
            if "lines" in mode_tokens and "markers" in mode_tokens:
                return "scatter_lines_markers"
            return "scatter"

        mapping = {
            "bar": "bar",
            "histogram": "histogram",
            "area": "area",
            "pie": "pie",
            "box": "box",
            "violin": "box",
            "heatmap": "heatmap",
            "waterfall": "waterfall",
            "bubble": "bubble",
        }
        return mapping.get(trace_type, "scatter")

    def _map_marker_shape(self, plotly_shape: str) -> str:
        mapping = {
            "circle": "xlMarkerStyleCircle",
            "square": "xlMarkerStyleSquare",
            "diamond": "xlMarkerStyleDiamond",
            "cross": "xlMarkerStyleX",
            "x": "xlMarkerStylePlus",
            "triangle-up": "xlMarkerStyleTriangle",
            "triangle-down": "xlMarkerStyleTriangle",
        }
        return mapping.get(plotly_shape.lower(), "xlMarkerStyleCircle")

    def build_xml_tree(self) -> LET.etree.ElementTree:
        root = LET.etree.Element("plotly_excel_chart", version=self.SCHEMA_VERSION)
        self._build_chart_meta(root)
        self._build_traces(root)
        self._build_extras(root)
        return LET.etree.ElementTree(root)

    def _build_chart_meta(self, root):
        layout = self.figure.layout
        chart_meta = LET.etree.SubElement(root, "chart_meta")

        if len(self.figure.data) == 0:
            chart_type_value = "xy"
        else:
            first_type = self.figure.data[0].type
            chart_type_value = self._map_plotly_type_to_chart_type(first_type)
        LET.etree.SubElement(chart_meta, "chart_type").text = chart_type_value

        LET.etree.SubElement(chart_meta, "title").text = self._safe_text(
            layout.title.text if layout.title else ""
        )

        x_axis = LET.etree.SubElement(chart_meta, "x_axis")
        LET.etree.SubElement(x_axis, "title").text = self._safe_text(
            layout.xaxis.title.text if layout.xaxis.title else ""
        )
        LET.etree.SubElement(x_axis, "min").text = self._safe_text(
            layout.xaxis.range[0] if layout.xaxis.range else ""
        )
        LET.etree.SubElement(x_axis, "max").text = self._safe_text(
            layout.xaxis.range[1] if layout.xaxis.range else ""
        )
        LET.etree.SubElement(x_axis, "log_scale").text = (
            "true" if layout.xaxis.type == "log" else "false"
        )

        y_axis = LET.etree.SubElement(chart_meta, "y_axis")
        LET.etree.SubElement(y_axis, "title").text = self._safe_text(
            layout.yaxis.title.text if layout.yaxis.title else ""
        )
        LET.etree.SubElement(y_axis, "min").text = self._safe_text(
            layout.yaxis.range[0] if layout.yaxis.range else ""
        )
        LET.etree.SubElement(y_axis, "max").text = self._safe_text(
            layout.yaxis.range[1] if layout.yaxis.range else ""
        )
        LET.etree.SubElement(y_axis, "log_scale").text = (
            "true" if layout.yaxis.type == "log" else "false"
        )

        legend = LET.etree.SubElement(chart_meta, "legend")
        legend_visible = getattr(layout.legend, "visible", True)
        legend_orientation = getattr(layout.legend, "orientation", "v")
        LET.etree.SubElement(legend, "visible").text = (
            "true" if bool(legend_visible) else "false"
        )
        LET.etree.SubElement(legend, "position").text = (
            "right" if legend_orientation == "v" else "bottom"
        )

        # NEW: barmode for stacked/grouped
        LET.etree.SubElement(chart_meta, "barmode").text = getattr(layout, "barmode", "group")

        timestamp = dt.datetime.utcnow().isoformat(timespec="microseconds") + "+00:00"
        LET.etree.SubElement(chart_meta, "export_timestamp").text = timestamp

    def _build_traces(self, root):
        traces_el = LET.etree.SubElement(root, "traces")

        for idx, trace in enumerate(self.figure.data, start=1):
            trace_el = LET.etree.SubElement(traces_el, "trace", id=str(idx))
            data_el = LET.etree.SubElement(trace_el, "data")

            if trace.type == "histogram":
                if hasattr(trace, "xbins") and trace.xbins.start is not None and trace.xbins.end is not None and trace.xbins.size is not None:
                    bins = np.arange(trace.xbins.start, trace.xbins.end + trace.xbins.size, trace.xbins.size)
                    counts, bin_edges = np.histogram(trace.x, bins=bins)
                else:
                    counts, bin_edges = np.histogram(trace.x, bins="auto")

                bin_midpoints = (bin_edges[:-1] + bin_edges[1:]) / 2.0

                if hasattr(trace, "text") and trace.text is not None and len(trace.text) > 0:
                    bin_labels = [
                        str(trace.text[i]) if i < len(trace.text) and trace.text[i] is not None else ""
                        for i in range(len(bin_midpoints))
                    ]
                else:
                    bin_labels = [
                        f"{round(bin_edges[i], 4)}:{round(bin_edges[i + 1], 4)}"
                        for i in range(len(bin_edges) - 1)
                    ]

                LET.etree.SubElement(data_el, "x").text = self._comma_join(bin_midpoints)
                LET.etree.SubElement(data_el, "y").text = self._comma_join(counts)
                LET.etree.SubElement(data_el, "text").text = self._comma_join(bin_labels)

            elif trace.type == "bar" and (not hasattr(trace, "y") or trace.y is None or len(trace.y) == 0):
                counts, bin_edges = np.histogram(trace.x, bins="auto")
                x_labels = [f"{bin_edges[i]}-{bin_edges[i+1]}" for i in range(len(bin_edges)-1)]
                LET.etree.SubElement(data_el, "x").text = self._comma_join(x_labels)
                LET.etree.SubElement(data_el, "y").text = self._comma_join(counts)

            else:
                LET.etree.SubElement(data_el, "x").text = self._comma_join(getattr(trace, "x", []))
                LET.etree.SubElement(data_el, "y").text = self._comma_join(getattr(trace, "y", []))
                LET.etree.SubElement(data_el, "z").text = self._comma_join(getattr(trace, "z", []))

                marker_size_values = getattr(trace.marker, "size", None) if hasattr(trace, "marker") else None
                if isinstance(marker_size_values, (list, tuple)):
                    marker_size_str = self._comma_join(marker_size_values)
                elif marker_size_values is not None:
                    marker_size_str = str(marker_size_values)
                else:
                    marker_size_str = ""
                LET.etree.SubElement(data_el, "size").text = marker_size_str

                LET.etree.SubElement(data_el, "text").text = self._comma_join(getattr(trace, "text", []))

                # NEW: marker.opacity array support
                marker_opacity = getattr(trace.marker, "opacity", None) if hasattr(trace, "marker") else None
                if isinstance(marker_opacity, (list, tuple)):
                    LET.etree.SubElement(data_el, "opacity").text = self._comma_join(marker_opacity)
                elif marker_opacity is not None:
                    LET.etree.SubElement(data_el, "opacity").text = str(marker_opacity)

            style_el = LET.etree.SubElement(trace_el, "style")
            LET.etree.SubElement(style_el, "series_type").text = self._map_plotly_type_to_series_type(trace)
            LET.etree.SubElement(style_el, "name").text = self._safe_text(getattr(trace, "name", f"Series {idx}"))

            # NEW: secondary axis support
            axis_group = "secondary" if getattr(trace, "yaxis", "y") != "y" else "primary"
            LET.etree.SubElement(style_el, "axis_group").text = axis_group

            LET.etree.SubElement(style_el, "visibility").text = "true"

            line_color = "#000000"
            line_style = "solid"
            line_width = None
            if hasattr(trace, "line") and trace.line:
                if trace.line.color:
                    line_color = str(trace.line.color)
                if trace.line.dash:
                    line_style = str(trace.line.dash)
                if trace.line.width is not None:
                    line_width = str(trace.line.width)
            LET.etree.SubElement(style_el, "line_color").text = line_color
            LET.etree.SubElement(style_el, "line_style").text = line_style
            if line_width is not None:
                LET.etree.SubElement(style_el, "line_width").text = line_width

            marker_el = LET.etree.SubElement(style_el, "marker")
            marker_size = "6"
            marker_color = "#000000"
            marker_shape = "xlMarkerStyleCircle"

            series_type = self._map_plotly_type_to_series_type(trace)
            if series_type == "scatter_lines":
                marker_size = "0"

            if hasattr(trace, "marker") and trace.marker:
                if getattr(trace.marker, "size", None) is not None:
                    if isinstance(trace.marker.size, (list, tuple)):
                        marker_size = ",".join(str(v) for v in trace.marker.size)
                    else:
                        marker_size = str(trace.marker.size)
                if getattr(trace.marker, "color", None) is not None:
                    marker_color = str(trace.marker.color)
                if getattr(trace.marker, "symbol", None) is not None:
                    marker_shape = self._map_marker_shape(str(trace.marker.symbol))

                # NEW: marker.line handling
                if hasattr(trace.marker, "line") and trace.marker.line:
                    LET.etree.SubElement(marker_el, "line_color").text = str(trace.marker.line.color or "")
                    LET.etree.SubElement(marker_el, "line_width").text = str(trace.marker.line.width or "")

            LET.etree.SubElement(marker_el, "size").text = marker_size
            LET.etree.SubElement(marker_el, "color").text = marker_color
            LET.etree.SubElement(marker_el, "shape").text = marker_shape

            LET.etree.SubElement(style_el, "fill_color").text = self._safe_text(getattr(trace, "fillcolor", ""))
            LET.etree.SubElement(style_el, "fill_opacity").text = self._safe_text(getattr(trace, "opacity", ""))

            # NEW: error bars
            for axis in ["x", "y"]:
                err_attr = f"error_{axis}"
                if hasattr(trace, err_attr) and getattr(trace, err_attr).visible:
                    err = getattr(trace, err_attr)
                    err_el = LET.etree.SubElement(trace_el, f"error_{axis}")
                    LET.etree.SubElement(err_el, "type").text = getattr(err, "type", "data")
                    LET.etree.SubElement(err_el, "symmetric").text = str(getattr(err, "symmetric", True)).lower()
                    LET.etree.SubElement(err_el, "array").text = self._comma_join(getattr(err, "array", []))
                    if hasattr(err, "arrayminus"):
                        LET.etree.SubElement(err_el, "arrayminus").text = self._comma_join(getattr(err, "arrayminus", []))

    def _build_extras(self, root):
        extras_el = LET.etree.SubElement(root, "extras")

        if hasattr(self.figure.layout, "shapes") and self.figure.layout.shapes:
            for idx, shape in enumerate(self.figure.layout.shapes, start=1):
                ann_type = None
                axis = None
                value = None
                span_axis = None
                span_mode = None

                if shape.type == "line":
                    if shape.x0 == shape.x1:
                        ann_type = "event_line"
                        axis = "x"
                        value = self._safe_text(shape.x0)
                        span_axis = "y"
                        span_mode = "full"
                    elif shape.y0 == shape.y1:
                        ann_type = "threshold"
                        axis = "y"
                        value = self._safe_text(shape.y0)
                        span_axis = "x"
                        span_mode = "full"

                # Skip "rect" or "region" shapes completely
                if not ann_type:
                    continue

                ann_el = LET.etree.SubElement(extras_el, "annotation", id=str(idx), type=ann_type)
                LET.etree.SubElement(ann_el, "axis").text = axis
                LET.etree.SubElement(ann_el, "value").text = value
                span_el = LET.etree.SubElement(ann_el, "span", axis=span_axis)
                span_el.set("mode", span_mode)

                style_el = LET.etree.SubElement(ann_el, "style")
                if hasattr(shape, "line") and shape.line:
                    LET.etree.SubElement(style_el, "color").text = self._safe_text(shape.line.color)
                    LET.etree.SubElement(style_el, "width").text = self._safe_text(shape.line.width)
                    LET.etree.SubElement(style_el, "dash").text = self._safe_text(getattr(shape.line, "dash", "solid"))
                if hasattr(shape, "opacity"):
                    LET.etree.SubElement(style_el, "opacity").text = self._safe_text(shape.opacity)

                if hasattr(shape, "name"):
                    LET.etree.SubElement(ann_el, "label").text = self._safe_text(shape.name)

    def save_to_file(self, output_path: str):
        tree = self.build_xml_tree()
        tree.write(output_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")

def figure_to_excel_xml(fig: Figure, output_path: str):
    converter = PlotlyToExcelXMLConverter(fig)
    converter.save_to_file(output_path)



# =======================================================
# ----------------------- Testing -----------------------
# =======================================================

if __name__ == "__main__":

    # ---------------- Example ----------------
    def _mk_plotly_charts() -> dict[str, object]:
        """
        Build four Plotly figures based on randomly generated data:
        1) Grouped columns with error bars + target line on secondary axis
        2) Overlaid histograms + smoothed density + CDF on secondary axis
        3) Two-cloud scatter with OLS trendlines and 95% CI bands (CI constrained to observed support)
        4) Two random walks, rolling means, volatility (secondary axis), range slider, and event markers
        """
        import numpy as np
        import pandas as pd
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots

        rng = np.random.default_rng()  # unseeded → different each run

        # ---------- 1) Grouped column chart + target line + errors ----------
        cats = [f"C{i}" for i in range(1, 8)]
        vals_a = rng.normal(loc=10, scale=3, size=len(cats))
        vals_b = vals_a * rng.normal(loc=1.05, scale=0.15, size=len(cats))
        target = np.full(len(cats), 11.0)
        err_a = np.clip(rng.normal(1.0, 0.4, len(cats)), 0.2, None)
        err_b = np.clip(rng.normal(0.8, 0.4, len(cats)), 0.2, None)

        fig_col = make_subplots(specs=[[{"secondary_y": True}]])
        fig_col.add_trace(
            go.Bar(name="Series A", x=cats, y=vals_a,
                error_y=dict(type="data", array=err_a, visible=True),
                marker_color="#3b82f6"),
            secondary_y=False
        )
        fig_col.add_trace(
            go.Bar(name="Series B", x=cats, y=vals_b,
                error_y=dict(type="data", array=err_b, visible=True),
                marker_color="#22c55e"),
            secondary_y=False
        )
        fig_col.add_trace(
            go.Scatter(name="Target", x=cats, y=target, mode="lines",
                    line=dict(dash="dash", width=2, color="#ef4444")),
            secondary_y=True
        )
        last_idx = -1
        delta = vals_a[last_idx] - target[last_idx]
        fig_col.add_annotation(
            x=cats[last_idx], y=max(vals_a[last_idx], vals_b[last_idx]),
            text=f"Δ to target (A): {delta:+.2f}", showarrow=True, arrowhead=2
        )
        fig_col.update_layout(
            title="Random Column Chart (Grouped) + Target + Errors",
            barmode="group", legend_orientation="h", legend_y=1.15,
            xaxis_title="Category", yaxis_title="Value",
        )
        fig_col.update_yaxes(title_text="Value (bars)", secondary_y=False)
        fig_col.update_yaxes(title_text="Target", secondary_y=True, showgrid=False)

        # ---------- 2) Overlaid histograms + KDE-ish line + cumulative on 2nd axis ----------
        samples1 = rng.normal(loc=0.0, scale=1.0, size=2000)
        samples2 = rng.normal(loc=0.5, scale=1.2, size=2000)
        bins = 40

        hist_y, hist_x = np.histogram(np.concatenate([samples1, samples2]), bins=bins, density=True)
        centers = 0.5 * (hist_x[1:] + hist_x[:-1])
        grid = centers[:, None] - centers[None, :]
        kernel = np.exp(-0.5 * (grid / 0.35) ** 2)
        kernel /= kernel.sum(axis=1, keepdims=True)
        kde = kernel @ hist_y

        xs = np.sort(samples1)
        cdf = np.linspace(0, 1, xs.size)

        # Create bin labels
        bin_labels = [f"{round(hist_x[i], 2)} - {round(hist_x[i+1], 2)}" for i in range(len(hist_x)-1)]

        fig_hist = make_subplots(specs=[[{"secondary_y": True}]])
        fig_hist.add_trace(
            go.Histogram(
                name="Dist A",
                x=samples1,
                nbinsx=bins,
                opacity=0.55,
                marker_color="#2563eb",
                histnorm="probability density",
                text=bin_labels  # added labels
            ),
            secondary_y=False
        )
        fig_hist.add_trace(
            go.Histogram(
                name="Dist B",
                x=samples2,
                nbinsx=bins,
                opacity=0.55,
                marker_color="#16a34a",
                histnorm="probability density",
                text=bin_labels  # added labels
            ),
            secondary_y=False
        )
        fig_hist.add_trace(
            go.Scatter(
                name="Smoothed Density",
                x=centers,
                y=kde,
                mode="lines",
                line=dict(width=3, color="#f59e0b")
            ),
            secondary_y=False
        )
        fig_hist.add_trace(
            go.Scatter(
                name="CDF (A)",
                x=xs,
                y=cdf,
                mode="lines",
                line=dict(dash="dot", width=2, color="#dc2626")
            ),
            secondary_y=True
        )
        fig_hist.update_layout(
            title="Random Histograms (Overlaid) + Smoothed Density + CDF",
            barmode="overlay",
            xaxis_title="Value",
            yaxis_title="Density",
            legend_orientation="h",
            legend_y=1.15
        )
        fig_hist.update_yaxes(title_text="Density", secondary_y=False)
        fig_hist.update_yaxes(title_text="Cumulative (A)", secondary_y=True, range=[0, 1])


        # ---------- 3) Scatter with two clouds + OLS trendlines + 95% CI band (constrained) ----------
        n = 450
        x = rng.normal(0, 1, n)
        y1 = 0.6 * x + rng.normal(0, 0.8, n)
        y2 = -0.4 * x + rng.normal(0, 0.7, n) + 0.5

        def ols_with_ci(xv, yv, qlo: float = 0.02, qhi: float = 0.98):
            import numpy as np
            X = np.c_[np.ones_like(xv), xv]
            beta = np.linalg.lstsq(X, yv, rcond=None)[0]
            yhat = X @ beta
            resid = yv - yhat
            s2 = np.sum(resid ** 2) / (len(xv) - 2)

            # limit evaluation grid to inner support to avoid edge inflation
            lo, hi = np.quantile(xv, [qlo, qhi])
            x0 = np.linspace(lo, hi, 200)
            X0 = np.c_[np.ones_like(x0), x0]

            XtX_inv = np.linalg.inv(X.T @ X)
            se = np.sqrt(np.sum((X0 @ XtX_inv) * X0, axis=1) * s2)
            y0 = X0 @ beta
            z = 1.96
            return x0, y0, y0 - z * se, y0 + z * se, beta

        x0, y0_1, lo1, hi1, beta1 = ols_with_ci(x, y1)
        x0b, y0_2, lo2, hi2, beta2 = ols_with_ci(x, y2)

        fig_scatter = go.Figure()
        fig_scatter.add_trace(go.Scatter(name="Cloud A", x=x, y=y1, mode="markers",
                                        marker=dict(size=6, color="#2563eb"), opacity=0.7))
        fig_scatter.add_trace(go.Scatter(name="Cloud B", x=x, y=y2, mode="markers",
                                        marker=dict(size=6, color="#16a34a"), opacity=0.7))
        # CI/Trend A
        fig_scatter.add_trace(go.Scatter(name="95% CI A", x=np.r_[x0, x0[::-1]],
                                        y=np.r_[hi1, lo1[::-1]], fill="toself",
                                        line=dict(width=0), fillcolor="rgba(37, 99, 235, 0.15)", showlegend=False))
        fig_scatter.add_trace(go.Scatter(name="Trend A", x=x0, y=y0_1, mode="lines",
                                        line=dict(width=3, color="#1d4ed8")))
        # CI/Trend B
        fig_scatter.add_trace(go.Scatter(name="95% CI B", x=np.r_[x0b, x0b[::-1]],
                                        y=np.r_[hi2, lo2[::-1]], fill="toself",
                                        line=dict(width=0), fillcolor="rgba(22, 163, 74, 0.15)", showlegend=False))
        fig_scatter.add_trace(go.Scatter(name="Trend B", x=x0b, y=y0_2, mode="lines",
                                        line=dict(width=3, color="#15803d")))
        fig_scatter.update_layout(
            title=f"Random Scatter with Two Trends (β₁={beta1[1]:.2f}, β₂={beta2[1]:.2f})",
            xaxis_title="X", yaxis_title="Y",
            legend_orientation="h", legend_y=1.14
        )
        fig_scatter.update_xaxes(zeroline=True)
        fig_scatter.update_yaxes(zeroline=True)

        # clamp axes to central quantiles
        xlo, xhi = np.quantile(x, [0.01, 0.99])
        ydata = np.r_[y1, y2]
        ylo, yhi = np.quantile(ydata, [0.01, 0.99])
        fig_scatter.update_xaxes(range=[float(xlo), float(xhi)])
        fig_scatter.update_yaxes(range=[float(ylo), float(yhi)])

        # ---------- 4) Random walks ----------
        n = 400
        steps1 = rng.normal(loc=0, scale=1, size=n)
        steps2 = rng.normal(loc=0.05, scale=1.1, size=n)
        walk1 = np.cumsum(steps1)
        walk2 = np.cumsum(steps2)
        t = np.arange(n)
        win = 20
        import pandas as pd
        roll1 = pd.Series(walk1).rolling(win, min_periods=1).mean().to_numpy()
        roll2 = pd.Series(walk2).rolling(win, min_periods=1).mean().to_numpy()
        vol = pd.Series(steps1).rolling(win, min_periods=1).std().to_numpy()

        event_xs = [int(0.6 * n)]

        from plotly.subplots import make_subplots
        fig_walk = make_subplots(specs=[[{"secondary_y": True}]])
        fig_walk.add_trace(go.Scatter(name="Walk A", x=t, y=walk1, mode="lines",
                                    line=dict(width=2, color="#0ea5e9")),
                        secondary_y=False)
        fig_walk.add_trace(go.Scatter(name="Walk B (drift)", x=t, y=walk2, mode="lines",
                                    line=dict(width=2, color="#22c55e")),
                        secondary_y=False)
        fig_walk.add_trace(go.Scatter(name=f"Rolling mean A ({win})", x=t, y=roll1, mode="lines",
                                    line=dict(width=3, dash="dash", color="#0369a1")),
                        secondary_y=False)
        fig_walk.add_trace(go.Scatter(name=f"Rolling mean B ({win})", x=t, y=roll2, mode="lines",
                                    line=dict(width=3, dash="dash", color="#15803d")),
                        secondary_y=False)
        fig_walk.add_trace(go.Scatter(name=f"Volatility A ({win})", x=t, y=vol, mode="lines",
                                    line=dict(width=2, color="#f59e0b")),
                        secondary_y=True)

        ev_shapes = []
        ev_ann = []
        for i, ex in enumerate(event_xs, 1):
            ev_shapes.append(dict(
                type="line", xref="x", yref="paper",
                x0=ex, x1=ex, y0=0, y1=1,
                line=dict(color="#ef4444", width=2, dash="dash")
            ))
            ev_ann.append(dict(
                x=ex, y=1.02, xref="x", yref="paper",
                text=f"Event {i}", showarrow=False,
                font=dict(size=12, color="#ef4444")
            ))

        fig_walk.update_layout(
            title="Random Walks + Rolling Means + Volatility",
            xaxis_title="t", yaxis_title="Level",
            legend_orientation="h", legend_y=1.14,
            xaxis_rangeslider=dict(visible=True),
            shapes=[dict(type="line", xref="x", yref="y",
                        x0=t.min(), x1=t.max(), y0=0, y1=0,
                        line=dict(dash="dot", width=1))] + ev_shapes,
            annotations=ev_ann
        )
        fig_walk.update_yaxes(title_text="Level", secondary_y=False)
        fig_walk.update_yaxes(title_text="Volatility", secondary_y=True, showgrid=False)

        return {
            "plotly_col_chart": fig_col,
            "plotly_histogram": fig_hist,
            "plotly_scatter": fig_scatter,
            "plotly_random_walk": fig_walk,
        }

    ###
    # ----------------------------------------------------------
    # Directories
    # ----------------------------------------------------------
    main_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    print(main_dir)

    data_dir = fr"C:\Users\Jonatan\Documents\Education\Msc\Master Banking and Finance\Semester C\Market structure\Assignment 2\data"
    os.makedirs(data_dir, exist_ok=True)

    # ----------------------------------------------------------
    # Utility: Validate required columns
    # ----------------------------------------------------------
    def validate_columns(df, required_cols, name):
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in {name}: {missing}")

    # ----------------------------------------------------------
    # Load and clean Quotes
    # ----------------------------------------------------------
    def load_quotes():
        quotes_path = os.path.join(data_dir, 'AVANZ_2weeks_quotes_ST.csv')
        quotes_df = pd.read_csv(quotes_path)

        print("\nQuotes DataFrame Columns:")
        print(quotes_df.columns.tolist())
        print("\nQuotes DataFrame Preview:")
        print(quotes_df.head())

        # Validate required columns
        required_cols = ['Type', 'Bid Price', 'Ask Price', 'Bid Size', 'Ask Size', 'Date-Time']
        validate_columns(quotes_df, required_cols, "Quotes")

        # ---- Cleaning and Feature Engineering ----
        quotes_df = quotes_df[quotes_df['Type'] != 'Auction']  # Drop auction quotes

        # Parse Date-Time and GMT Offset
        quotes_df['Date-Time'] = pd.to_datetime(quotes_df['Date-Time'], errors='coerce')
        quotes_df = quotes_df[quotes_df['Date-Time'].notna()]  # remove invalid datetimes

        quotes_df['GMT Offset'] = pd.to_numeric(quotes_df['GMT Offset'], errors='coerce')
        quotes_df = quotes_df[quotes_df['GMT Offset'].notna()]  # remove rows with NaN offset

        # Adjust Date-Time to LOCAL time
        offset_td = pd.to_timedelta(quotes_df['GMT Offset'], unit='h')
        quotes_df['Date-Time'] = quotes_df['Date-Time'] + offset_td  # overwrite with local time

        # Filter by local trading hours
        start_time = pd.to_datetime('09:00', format='%H:%M').time()
        end_time = pd.to_datetime('17:25', format='%H:%M').time()
        quotes_df = quotes_df[
            (quotes_df['Date-Time'].dt.time >= start_time) &
            (quotes_df['Date-Time'].dt.time <= end_time)
        ]

        print(f"Filtered Trades: {quotes_df.shape}")

        # Convert numeric fields
        quotes_df['Bid Price'] = pd.to_numeric(quotes_df['Bid Price'], errors='coerce')
        quotes_df['Ask Price'] = pd.to_numeric(quotes_df['Ask Price'], errors='coerce')

        # Filter out invalid bid/ask
        quotes_df = quotes_df[(quotes_df['Bid Price'] > 0) & (quotes_df['Ask Price'] > 0)]

        # Spread and mid-price
        quotes_df['spread'] = quotes_df['Ask Price'] - quotes_df['Bid Price']
        quotes_df['mid'] = (quotes_df['Ask Price'] + quotes_df['Bid Price']) / 2
        quotes_df['midspread'] = (quotes_df['spread'] / quotes_df['mid']) * 10**4

        # Bid/Ask price changes
        quotes_df['bid_change'] = quotes_df['Bid Price'].diff().abs()
        quotes_df['ask_change'] = quotes_df['Ask Price'].diff().abs()

        # Tick size
        changes = quotes_df[['bid_change', 'ask_change']].copy()
        changes = changes[(changes['bid_change'] != 0) | (changes['ask_change'] != 0)].dropna()
        tick_size = round(changes[['bid_change', 'ask_change']].min().min(), 1)
        print(f"Tick Size: {tick_size}")

        # Tick spread
        quotes_df['tick_spread'] = quotes_df['spread'] / tick_size

        # Depth (value-weighted depth in thousands)
        quotes_df['depth'] = (
            ((quotes_df['Bid Size'] * quotes_df['Bid Price']) +
            (quotes_df['Ask Size'] * quotes_df['Ask Price'])) / 2
        ) / 1000

        # Ensure standardized datetime
        quotes_df['Date-Time'] = pd.to_datetime(quotes_df['Date-Time'], errors='coerce')

        print("\nQuotes DataFrame with new columns:")
        print(quotes_df.head())

        return quotes_df

    # ----------------------------------------------------------
    # Load, clean, and standardize Trades
    # ----------------------------------------------------------
    def load_trades():
        trades_path = os.path.join(data_dir, 'AVANZ_2weeks_trades_ST.csv')
        trades_df = pd.read_csv(trades_path)

        print("\nInitial Trades DataFrame Columns:")
        print(trades_df.columns.tolist())
        print("\nInitial Trades DataFrame:")
        print(trades_df.head())

        # Validate required columns
        required_cols = ['Type', 'Date-Time', 'GMT Offset', 'Price']
        validate_columns(trades_df, required_cols, "Trades")

        # Drop auction trades
        trades_df = trades_df[trades_df['Type'] != 'Auction']
        print(f"\nTrades after dropping 'Auction': {trades_df.shape}")

        # Parse Date-Time and GMT Offset
        trades_df['Date-Time'] = pd.to_datetime(trades_df['Date-Time'], errors='coerce')
        trades_df = trades_df[trades_df['Date-Time'].notna()]  # remove invalid datetimes

        trades_df['GMT Offset'] = pd.to_numeric(trades_df['GMT Offset'], errors='coerce')
        trades_df = trades_df[trades_df['GMT Offset'].notna()]  # remove rows with NaN offset

        # Adjust Date-Time to LOCAL time
        offset_td = pd.to_timedelta(trades_df['GMT Offset'], unit='h')
        trades_df['Date-Time'] = trades_df['Date-Time'] + offset_td  # overwrite with local time

        # Filter by local trading hours
        start_time = pd.to_datetime('09:00', format='%H:%M').time()
        end_time = pd.to_datetime('17:25', format='%H:%M').time()
        trades_df = trades_df[
            (trades_df['Date-Time'].dt.time >= start_time) &
            (trades_df['Date-Time'].dt.time <= end_time)
        ]

        print(f"Filtered Trades: {trades_df.shape}")

        # Sort for plotting
        trades_df = trades_df.sort_values('Date-Time')

        # ---- Plot ----
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(trades_df['Date-Time'], trades_df['Price'], label='Trade Price', color='blue')
        ax.set_xlabel('Local Date-Time')
        ax.set_ylabel('Trade Price', color='blue')
        ax.tick_params(axis='y', labelcolor='blue')

        # ADD vertical red bold line
        april_2 = pd.to_datetime('2025-04-02')
        ax.axvline(april_2, color='red', linewidth=2, label='April 2, 2025')

        plt.xticks(rotation=45)
        ax.legend()
        plt.tight_layout()
        plt.savefig(os.path.join(data_dir, 'trades_plot.png'))

        return trades_df

    # ----------------------------------------------------------
    # Combine Quotes and Trades
    # ----------------------------------------------------------
    def combine_tables(quotes_df, trades_df):
        # Combine and sort
        combined_df = pd.concat([quotes_df, trades_df], ignore_index=True, sort=False)
        combined_df = combined_df.sort_values('Date-Time').reset_index(drop=True)

        # Forward-fill Bid/Ask for trades
        combined_df['Bid Price'] = combined_df['Bid Price'].ffill()
        combined_df['Ask Price'] = combined_df['Ask Price'].ffill()

        # Midpoint
        combined_df['mid'] = (combined_df['Bid Price'] + combined_df['Ask Price']) / 2

        # ---- Trade direction ----
        mask_trade = combined_df['Type'] == 'Trade'
        trades_only = combined_df.loc[mask_trade]

        directions = np.select(
            [
                np.isclose(trades_only['Price'], trades_only['Ask Price']),
                np.isclose(trades_only['Price'], trades_only['Bid Price'])
            ],
            [1, -1],
            default=np.nan
        )
        combined_df.loc[mask_trade, 'direction'] = directions

        # ---- Effective spread (bps) ----
        combined_df['effective_spread'] = (
            2 * combined_df['direction'] * (combined_df['Price'] - combined_df['mid']) / combined_df['mid']
        ) * 10**4

        # ---- Daily averages ----
        trades_df_only = combined_df.loc[mask_trade].copy()
        trades_df_only['date'] = trades_df_only['Date-Time'].dt.date
        daily_avg_effective_spread = trades_df_only.groupby('date')['effective_spread'].mean()

        mask_quote = combined_df['Type'] == 'Quote'
        quotes_df_only = combined_df.loc[mask_quote].copy()
        quotes_df_only['date'] = quotes_df_only['Date-Time'].dt.date
        daily_avg_quoted_spread = quotes_df_only.groupby('date')['midspread'].mean()

        # Create figure
        fig = go.Figure()

        # Effective Spread
        fig.add_trace(go.Scatter(
            x=daily_avg_effective_spread.index,
            y=daily_avg_effective_spread.values,
            mode='lines+markers',
            marker=dict(symbol='circle'),
            line=dict(color='green'),
            name='Average Effective Spread (Trades)'
        ))

        # Quoted Spread
        fig.add_trace(go.Scatter(
            x=daily_avg_quoted_spread.index,
            y=daily_avg_quoted_spread.values,
            mode='lines+markers',
            marker=dict(symbol='square'),
            line=dict(color='blue'),
            name='Average Quoted Spread (Quotes)'
        ))

        # ---- Extras for testing ----
        april_2 = pd.to_datetime('2025-04-02')
        fig.add_vline(
            x=april_2,
            line=dict(color='red', width=2, dash='dash'),
            name='April 2, 2025'
        )

        # Horizontal threshold line at spread = 5 bps
        fig.add_hline(
            y=5,
            line=dict(color='orange', width=2, dash='dot'),
            name='Threshold = 5 bps'
        )

        # Vertical region (band) from April 5–April 7
        start_band = pd.to_datetime('2025-04-05')
        end_band = pd.to_datetime('2025-04-07')
        fig.add_vrect(
            x0=start_band, x1=end_band,
            fillcolor='lightblue', opacity=0.3,
            layer='below', line_width=0,
        )

        # Text annotation
        fig.add_annotation(
            x=april_2, y=5,
            text="Event: April 2",
            showarrow=True,
            arrowhead=2
        )

        # Layout settings
        fig.update_layout(
            title='Daily Average Effective Spread vs Quoted Spread',
            xaxis_title='Date',
            yaxis_title='Spread (bps)',
            xaxis=dict(tickangle=45),
            legend=dict(title='Legend'),
            template='simple_white'
        )

        fig.show()

        # ---- Preview combined dataframe ----
        print("\nCombined DataFrame with effective spread:")
        print(
            combined_df.loc[combined_df['Type'] == 'Trade',
                            ['Date-Time', 'Price', 'Bid Price', 'Ask Price', 'mid', 'direction', 'effective_spread']]
            .head()
        )

        return fig



    # ----------------------------------------------------------
    # Main
    # ----------------------------------------------------------
    def main():
        quotes_plt = load_quotes()
        trades_plt = load_trades()
        result = combine_tables(quotes_plt, trades_plt)
        return {'result': result}

    ###

    # Generate test Plotly figures
    charts = _mk_plotly_charts()
    # charts = main()

    # Base folder where XML files will be stored
    output_folder = "plotly_xml_outputs"
    os.makedirs(output_folder, exist_ok=True)  # Create base folder

    for name, fig in charts.items():
        # Option 1: Save directly as a single file (recommended)
        output_path = os.path.join(output_folder, f"{name}_output.xml")

        # Option 2: If you want a subfolder per chart, uncomment this block:
        # chart_folder = os.path.join(output_folder, name)
        # os.makedirs(chart_folder, exist_ok=True)
        # output_path = os.path.join(chart_folder, "_output.xml")

        # Generate the XML
        figure_to_excel_xml(fig, output_path)

        print(f"XML written to {output_path}")





# ----------------------------
# NOT USED CONSIDER DELETING!
# # ----------------------------
# def _ensure_dir(p: str) -> None:
#     os.makedirs(p, exist_ok=True)

# def _attrs(d: Dict[str, Any]) -> Dict[str, str]:
#     return {k: "" if v is None else str(v) for k, v in d.items()}

# _XMLNS = "urn:example:chartspec:1.0"

# def write_chart_xml(path: str, chart_spec: Dict[str, Any]) -> None:
#     NSMAP = {None: _XMLNS}
#     root = LET.Element("ChartSpec", nsmap=NSMAP, version="1.0")

#     chart = LET.SubElement(root, "chart", type=chart_spec.get("type", "line"))

#     if cat := chart_spec.get("categoryAxis"):
#         LET.SubElement(chart, "categoryAxis", **_attrs(cat))
#     if val := chart_spec.get("valueAxis"):
#         LET.SubElement(chart, "valueAxis", **_attrs(val))
#     if val2 := chart_spec.get("valueAxis2"):
#         LET.SubElement(chart, "valueAxis2", **_attrs(val2))

#     series_root = LET.SubElement(root, "series")
#     for s in chart_spec.get("series", []):
#         s_attrs = {}
#         name = s.get("name")
#         if name is not None:
#             s_attrs["name"] = str(name)
#         axis = s.get("axis")
#         if axis:
#             s_attrs["yAxis"] = str(axis)
#         s_elem = LET.SubElement(series_root, "s", **s_attrs)

#         x_elem = LET.SubElement(s_elem, "x")
#         # Pre-stringify once to reduce Python-level overhead
#         for xv in s.get("x", []):
#             n = LET.SubElement(x_elem, "n")
#             n.text = "" if xv is None else str(xv)

#         y_elem = LET.SubElement(s_elem, "y")
#         for yv in s.get("y", []):
#             n = LET.SubElement(y_elem, "n")
#             n.text = "" if yv is None else str(yv)

#         if st := s.get("style"):
#             LET.SubElement(s_elem, "style", **_attrs(st))

#     if layout := chart_spec.get("layout"):
#         layout_elem = LET.SubElement(root, "layout")
#         if "title" in layout:
#             LET.SubElement(layout_elem, "title", text=str(layout["title"]))
#         if "legend" in layout:
#             LET.SubElement(layout_elem, "legend", show=("true" if layout["legend"] else "false"))
#         for k, v in layout.items():
#             if k in ("title", "legend"):
#                 continue
#             LET.SubElement(layout_elem, "opt", key=str(k), value=("" if v is None else str(v)))

#     _ensure_dir(os.path.dirname(path) or ".")
#     # No pretty print; use binary write; default UTF-8
#     data = LET.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=False)
#     with open(path, "wb") as f:
#         f.write(data)
# ----------------------------
# ----------------------------






# def dict_to_xml(path: str, chart_spec: Dict[str, Any]) -> None:
#     """
#     Extremely fast conversion of any nested dict/list/scalar (e.g., Plotly fig.to_dict()) to XML.
#     Uses lxml's streaming writer (C backend) to minimize Python overhead and memory usage.
    
#     Args:
#         path: Output XML file path.
#         data: Dictionary to serialize.
#         root_name: Top-level XML tag name.
#     """

#     def _write_element(xml_writer, key, value):
#         """Fast type-specific dispatch with minimal recursion."""
#         if isinstance(value, dict):
#             with xml_writer.element(str(key)):
#                 for k, v in value.items():
#                     _write_element(xml_writer, k, v)
#         elif isinstance(value, list):
#             with xml_writer.element(str(key)):
#                 for item in value:
#                     _write_element(xml_writer, "item", item)
#         else:
#             # Scalar: write as text
#             text_value = "" if value is None else str(value)
#             xml_writer.write(LET.Element(str(key)))
#             # Manually set text directly in the last node for speed
#             xml_writer._file.write(f"<{key}>{text_value}</{key}>".encode("utf-8"))

#     os.makedirs(os.path.dirname(path) or ".", exist_ok=True)

#     with LET.xmlfile(path, encoding="UTF-8") as xf:
#         # Root element is streamed, not built
#         with xf.element(root_name):
#             for k, v in data.items():
#                 _write_element(xf, k, v)



# --------- Utility: Sanitize text for XML ---------
# XML 1.0 invalid characters:
# [#x00-#x08] | #x0B | #x0C | [#x0E-#x1F] | #x7F
# INVALID_XML_RE = re.compile(
#     r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]"
# )

# def sanitize_xml_text(text: str) -> str:
#     """Remove invalid XML 1.0 characters including NULL bytes."""
#     return INVALID_XML_RE.sub("", text)


# # --------- Plotly Figure to XML Converter ---------
# class FigureToXML:
#     ROOT_TAG = "plotly_figure"

#     def __init__(self, ensure_ascii: bool = False) -> None:
#         self.ensure_ascii = ensure_ascii

#     def to_file(self, fig: go.Figure, out_path: Union[str, os.PathLike]) -> None:
#         with open(out_path, "wb") as out_fp:
#             self._write_xml_stream(fig, out_fp)

#     def to_bytes(self, fig: go.Figure) -> bytes:
#         buf = io.BytesIO()
#         self._write_xml_stream(fig, buf)
#         return buf.getvalue()

#     def to_string(self, fig: go.Figure) -> str:
#         return self.to_bytes(fig).decode("utf-8")

#     def _write_xml_stream(self, fig: go.Figure, out_fp: IO[bytes]) -> None:
#         """
#         Convert a Plotly figure to JSON and stream it to XML,
#         removing any illegal XML characters.
#         """
#         # Plotly older versions don't have include_defaults
#         full_json_str = pio.to_json(fig)
#         full_dict = json.loads(full_json_str)

#         # Remove default template noise manually
#         if "layout" in full_dict and "template" in full_dict["layout"]:
#             full_dict["layout"].pop("template", None)

#         # Write cleaned JSON to temporary file for streaming
#         jtmp = tempfile.NamedTemporaryFile(prefix="plotly_json_", suffix=".json", delete=False)
#         jtmp_name = jtmp.name
#         try:
#             rj.dump(
#                 full_dict,
#                 jtmp,
#                 ensure_ascii=self.ensure_ascii,
#                 number_mode=rj.NM_NATIVE,
#                 datetime_mode=rj.DM_NONE,
#                 uuid_mode=rj.UM_NONE,
#                 write_mode=rj.WM_COMPACT,
#             )
#             jtmp.flush()
#             jtmp.close()

#             with open(jtmp_name, "rb") as jf, xmlfile(out_fp, encoding="utf-8") as xf:
#                 xf.write_declaration()
#                 try:
#                     with xf.element(self.ROOT_TAG):
#                         self._json_events_to_xml(jf, xf)
#                 except Exception as e:
#                     print(f"Streaming error: {e}")
#                     raise
#         finally:
#             try:
#                 os.unlink(jtmp_name)
#             except OSError:
#                 pass

#     def _json_events_to_xml(self, jf: IO[bytes], xf: xmlfile) -> None:
#         """Convert JSON tokens to hierarchical XML."""
#         stack = []

#         for prefix, event, value in ijson.parse(jf):
#             if event == "start_map":
#                 if stack and stack[-1]["type"] == "array":
#                     ctx = xf.element("item")
#                 elif stack and stack[-1]["type"] == "object":
#                     ctx = xf.element(stack[-1]["current_key"])
#                 else:
#                     ctx = None
#                 stack.append({"type": "object", "current_key": None, "ctx": ctx})
#                 if ctx:
#                     ctx.__enter__()

#             elif event == "map_key":
#                 stack[-1]["current_key"] = value

#             elif event == "end_map":
#                 current = stack.pop()
#                 if current["ctx"]:
#                     current["ctx"].__exit__(None, None, None)
#                 if stack and stack[-1]["type"] == "object":
#                     stack[-1]["current_key"] = None

#             elif event == "start_array":
#                 if stack[-1]["type"] == "object":
#                     ctx = xf.element(stack[-1]["current_key"])
#                 else:
#                     ctx = xf.element("item")
#                 stack.append({"type": "array", "ctx": ctx})
#                 ctx.__enter__()

#             elif event == "end_array":
#                 current = stack.pop()
#                 current["ctx"].__exit__(None, None, None)
#                 if stack and stack[-1]["type"] == "object":
#                     stack[-1]["current_key"] = None

#             elif event in ("string", "number", "boolean", "null"):
#                 raw_text = (
#                     "true" if value is True
#                     else "false" if value is False
#                     else "" if value is None
#                     else str(value)
#                 )

#                 # Sanitize to ensure valid XML
#                 safe_text = sanitize_xml_text(raw_text)

#                 # Debug logging to identify problematic values
#                 if raw_text != safe_text:
#                     print(f"Sanitized value: {repr(raw_text)} → {repr(safe_text)}")

#                 if stack[-1]["type"] == "object":
#                     with xf.element(stack[-1]["current_key"]):
#                         xf.write(safe_text.encode("utf-8"))
#                     stack[-1]["current_key"] = None
#                 elif stack[-1]["type"] == "array":
#                     with xf.element("item"):
#                         xf.write(safe_text.encode("utf-8"))




# def dict_to_xml_file(data: dict, dst: str, root_tag: str = "root") -> None:
#     """
#     Convert a Python dictionary to an XML file.

#     Args:
#         data (dict): Input dictionary to convert.
#         dst (str): Destination file path to save the XML.
#         root_tag (str): Root XML tag name. Default is "root".
#     """
#     if not isinstance(data, Mapping):
#         raise TypeError("Input data must be a dictionary")

#     def build_xml(parent, value):
#         if isinstance(value, Mapping):
#             for key, val in value.items():
#                 if isinstance(val, Sequence) and not isinstance(val, (str, bytes, bytearray)):
#                     # Handle list: wrap in parent, repeat child element
#                     list_parent = etree.SubElement(parent, key)
#                     singular = key[:-1] if key.endswith('s') else "item"
#                     for item in val:
#                         item_tag = etree.SubElement(list_parent, singular)
#                         build_xml(item_tag, item)
#                 else:
#                     child = etree.SubElement(parent, key)
#                     build_xml(child, val)
#         else:
#             parent.text = str(value)

#     root = etree.Element(root_tag)
#     build_xml(root, data)

#     tree = etree.ElementTree(root)
#     with open(dst, "wb") as f:
#         tree.write(f, xml_declaration=True, encoding="UTF-8", pretty_print=True)




# from lxml import etree
# from plotly.graph_objs import Figure
# import datetime
# import numpy as np


# class PlotlyToExcelXMLConverter:
#     """
#     Converts a Plotly Figure into an XML format strictly conforming
#     to the XML → Excel Chart Specification (version 2.0).
#     """

#     SCHEMA_VERSION = "2.0"

#     def __init__(self, figure: Figure):
#         if not isinstance(figure, Figure):
#             raise TypeError("Input must be a plotly.graph_objs.Figure")
#         self.figure = figure

#     # ===============================
#     # Helpers
#     # ===============================
#     def _safe_text(self, value) -> str:
#         return "" if value is None else str(value)

#     def _comma_join(self, values) -> str:
#         if values is None:
#             return ""
#         return ",".join(str(v) for v in values)

#     def _map_plotly_type_to_chart_type(self, plotly_type: str) -> str:
#         mapping = {
#             "scatter": "xy",
#             "scattergl": "xy",
#             "bar": "bar",
#             "line": "line",
#             "area": "area",
#             "pie": "pie",
#         }
#         return mapping.get(plotly_type.lower(), "xy")

#     def _map_plotly_type_to_series_type(self, trace) -> str:
#         trace_type = getattr(trace, "type", "").lower()

#         if trace_type in ("scatter", "scattergl"):
#             mode = getattr(trace, "mode", "") or ""
#             mode_tokens = {m.strip().lower() for m in mode.split("+")}
#             if "lines" in mode_tokens and "markers" not in mode_tokens:
#                 return "scatter_lines"
#             if "lines" in mode_tokens and "markers" in mode_tokens:
#                 return "scatter_lines_markers"
#             return "scatter"

#         mapping = {
#             "bar": "bar",
#             "histogram": "histogram",
#             "area": "area",
#             "pie": "pie",
#             "box": "box",
#             "violin": "box",
#             "heatmap": "heatmap",
#             "waterfall": "waterfall",
#             "bubble": "bubble",
#         }
#         return mapping.get(trace_type, "scatter")

#     def _map_marker_shape(self, plotly_shape: str) -> str:
#         mapping = {
#             "circle": "xlMarkerStyleCircle",
#             "square": "xlMarkerStyleSquare",
#             "diamond": "xlMarkerStyleDiamond",
#             "cross": "xlMarkerStyleX",
#             "x": "xlMarkerStylePlus",
#             "triangle-up": "xlMarkerStyleTriangle",
#             "triangle-down": "xlMarkerStyleTriangle",  # rotation handled in VBA
#         }
#         return mapping.get(plotly_shape.lower(), "xlMarkerStyleCircle")

#     # ===============================
#     # Build Full XML
#     # ===============================
#     def build_xml_tree(self) -> etree.ElementTree:
#         root = etree.Element("plotly_excel_chart", version=self.SCHEMA_VERSION)
#         self._build_chart_meta(root)
#         self._build_traces(root)
#         self._build_extras(root)
#         return etree.ElementTree(root)

#     # ===============================
#     # Build <chart_meta>
#     # ===============================
#     def _build_chart_meta(self, root):
#         layout = self.figure.layout
#         chart_meta = etree.SubElement(root, "chart_meta")

#         # <chart_type>
#         if len(self.figure.data) == 0:
#             chart_type_value = "xy"
#         else:
#             first_type = self.figure.data[0].type
#             chart_type_value = self._map_plotly_type_to_chart_type(first_type)
#         etree.SubElement(chart_meta, "chart_type").text = chart_type_value

#         # <title>
#         etree.SubElement(chart_meta, "title").text = self._safe_text(
#             layout.title.text if layout.title else ""
#         )

#         # <x_axis>
#         x_axis = etree.SubElement(chart_meta, "x_axis")
#         etree.SubElement(x_axis, "title").text = self._safe_text(
#             layout.xaxis.title.text if layout.xaxis.title else ""
#         )
#         etree.SubElement(x_axis, "min").text = self._safe_text(
#             layout.xaxis.range[0] if layout.xaxis.range else ""
#         )
#         etree.SubElement(x_axis, "max").text = self._safe_text(
#             layout.xaxis.range[1] if layout.xaxis.range else ""
#         )
#         etree.SubElement(x_axis, "log_scale").text = (
#             "true" if layout.xaxis.type == "log" else "false"
#         )

#         # <y_axis>
#         y_axis = etree.SubElement(chart_meta, "y_axis")
#         etree.SubElement(y_axis, "title").text = self._safe_text(
#             layout.yaxis.title.text if layout.yaxis.title else ""
#         )
#         etree.SubElement(y_axis, "min").text = self._safe_text(
#             layout.yaxis.range[0] if layout.yaxis.range else ""
#         )
#         etree.SubElement(y_axis, "max").text = self._safe_text(
#             layout.yaxis.range[1] if layout.yaxis.range else ""
#         )
#         etree.SubElement(y_axis, "log_scale").text = (
#             "true" if layout.yaxis.type == "log" else "false"
#         )

#         # <legend>
#         legend = etree.SubElement(chart_meta, "legend")
#         legend_visible = getattr(layout.legend, "visible", True)
#         legend_orientation = getattr(layout.legend, "orientation", "v")
#         etree.SubElement(legend, "visible").text = (
#             "true" if bool(legend_visible) else "false"
#         )
#         etree.SubElement(legend, "position").text = (
#             "right" if legend_orientation == "v" else "bottom"
#         )

#         # <export_timestamp>
#         timestamp = dt.datetime.utcnow().isoformat(timespec="microseconds") + "+00:00"
#         etree.SubElement(chart_meta, "export_timestamp").text = timestamp

#     # ===============================
#     # Build <traces>
#     # ===============================
#     def _build_traces(self, root):
#         traces_el = etree.SubElement(root, "traces")

#         for idx, trace in enumerate(self.figure.data, start=1):
#             trace_el = etree.SubElement(traces_el, "trace", id=str(idx))
#             data_el = etree.SubElement(trace_el, "data")



#             # ---------- HISTOGRAM HANDLING ----------
#             if trace.type == "histogram":
#                 # Use Plotly bin info if available
#                 if hasattr(trace, "xbins") and trace.xbins.start is not None and trace.xbins.end is not None and trace.xbins.size is not None:
#                     # Construct bin edges using Plotly's own settings
#                     bins = np.arange(trace.xbins.start, trace.xbins.end + trace.xbins.size, trace.xbins.size)
#                     counts, bin_edges = np.histogram(trace.x, bins=bins)
#                 else:
#                     # If no explicit bin info exists, let numpy determine bins automatically
#                     counts, bin_edges = np.histogram(trace.x, bins="auto")

#                 # Numeric midpoints for Excel X axis
#                 bin_midpoints = (bin_edges[:-1] + bin_edges[1:]) / 2.0

#                 # ---------- Handle Labels ----------
#                 # If Plotly provided explicit text labels, align them to bins
#                 if hasattr(trace, "text") and trace.text is not None and len(trace.text) > 0:
#                     bin_labels = [
#                         str(trace.text[i]) if i < len(trace.text) and trace.text[i] is not None else ""
#                         for i in range(len(bin_midpoints))
#                     ]
#                 else:
#                     # Fallback: generate default range labels
#                     bin_labels = [
#                         f"{round(bin_edges[i], 4)}:{round(bin_edges[i + 1], 4)}"
#                         for i in range(len(bin_edges) - 1)
#                     ]

#                 # ---------- Write to XML ----------
#                 # <x> must remain numeric for Excel scaling
#                 etree.SubElement(data_el, "x").text = self._comma_join(bin_midpoints)
#                 etree.SubElement(data_el, "y").text = self._comma_join(counts)

#                 # <text> is always same length as bin_midpoints
#                 etree.SubElement(data_el, "text").text = self._comma_join(bin_labels)




#             elif trace.type == "bar" and (not hasattr(trace, "y") or trace.y is None or len(trace.y) == 0):
#                 counts, bin_edges = np.histogram(trace.x, bins="auto")
#                 x_labels = [f"{bin_edges[i]}-{bin_edges[i+1]}" for i in range(len(bin_edges)-1)]
#                 etree.SubElement(data_el, "x").text = self._comma_join(x_labels)
#                 etree.SubElement(data_el, "y").text = self._comma_join(counts)

#             else:
#                 etree.SubElement(data_el, "x").text = self._comma_join(getattr(trace, "x", []))
#                 etree.SubElement(data_el, "y").text = self._comma_join(getattr(trace, "y", []))

#                 # <z>
#                 etree.SubElement(data_el, "z").text = self._comma_join(getattr(trace, "z", []))

#                 # <size> FIXED IMPLEMENTATION
#                 marker_size_values = getattr(trace.marker, "size", None) if hasattr(trace, "marker") else None
#                 if isinstance(marker_size_values, (list, tuple)):
#                     marker_size_str = self._comma_join(marker_size_values)
#                 elif marker_size_values is not None:
#                     marker_size_str = str(marker_size_values)
#                 else:
#                     marker_size_str = ""
#                 etree.SubElement(data_el, "size").text = marker_size_str

#                 # <text>
#                 etree.SubElement(data_el, "text").text = self._comma_join(getattr(trace, "text", []))

#             # ==============================
#             # <style>
#             # ==============================
#             style_el = etree.SubElement(trace_el, "style")
#             etree.SubElement(style_el, "series_type").text = self._map_plotly_type_to_series_type(trace)
#             etree.SubElement(style_el, "name").text = self._safe_text(getattr(trace, "name", f"Series {idx}"))
#             etree.SubElement(style_el, "axis_group").text = "primary"
#             etree.SubElement(style_el, "visibility").text = "true"

#             # Line
#             line_color = "#000000"
#             line_style = "solid"
#             line_width = None
#             if hasattr(trace, "line") and trace.line:
#                 if trace.line.color:
#                     line_color = str(trace.line.color)
#                 if trace.line.dash:
#                     line_style = str(trace.line.dash)
#                 if trace.line.width is not None:
#                     line_width = str(trace.line.width)
#             etree.SubElement(style_el, "line_color").text = line_color
#             etree.SubElement(style_el, "line_style").text = line_style
#             if line_width is not None:
#                 etree.SubElement(style_el, "line_width").text = line_width

#             # Marker
#             marker_el = etree.SubElement(style_el, "marker")
#             marker_size = "6"
#             marker_color = "#000000"
#             marker_shape = "xlMarkerStyleCircle"

#             series_type = self._map_plotly_type_to_series_type(trace)
#             if series_type == "scatter_lines":
#                 marker_size = "0"

#             if hasattr(trace, "marker") and trace.marker:
#                 if getattr(trace.marker, "size", None) is not None:
#                     if isinstance(trace.marker.size, (list, tuple)):
#                         marker_size = ",".join(str(v) for v in trace.marker.size)
#                     else:
#                         marker_size = str(trace.marker.size)
#                 if getattr(trace.marker, "color", None) is not None:
#                     marker_color = str(trace.marker.color)
#                 if getattr(trace.marker, "symbol", None) is not None:
#                     marker_shape = self._map_marker_shape(str(trace.marker.symbol))

#             etree.SubElement(marker_el, "size").text = marker_size
#             etree.SubElement(marker_el, "color").text = marker_color
#             etree.SubElement(marker_el, "shape").text = marker_shape

#             # Fill
#             etree.SubElement(style_el, "fill_color").text = self._safe_text(getattr(trace, "fillcolor", ""))
#             etree.SubElement(style_el, "fill_opacity").text = self._safe_text(getattr(trace, "opacity", ""))

#     # ===============================
#     # Build <extras>
#     # ===============================
#     def _build_extras(self, root):
#         etree.SubElement(root, "extras")

#     # ===============================
#     # Save to File
#     # ===============================
#     def save_to_file(self, output_path: str):
#         tree = self.build_xml_tree()
#         tree.write(output_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")


# # ================================
# # Convenience function
# # ================================
# def figure_to_excel_xml(fig: Figure, output_path: str):
#     converter = PlotlyToExcelXMLConverter(fig)
#     converter.save_to_file(output_path)
