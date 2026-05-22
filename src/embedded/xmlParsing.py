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
                    # Try to convert to numeric first (Excel format)
                    # Excel stores dates as days since 1899-12-30
                    # This handles: dates (45741), times (0.5), datetimes (45741.65625)
                    numeric_vals = pd.to_numeric(s, errors="coerce")

                    if numeric_vals.notna().any():
                        # Excel epoch: 1899-12-30 (Excel incorrectly treats 1900 as leap year)
                        df_dict[cname] = pd.to_datetime(
                            numeric_vals,
                            unit='D',
                            origin='1899-12-30',
                            utc=True,
                            errors='coerce'
                        )
                    else:
                        # Fallback for string dates (backwards compatibility)
                        df_dict[cname] = pd.to_datetime(s, dayfirst=True, errors="coerce", utc=True)

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
        # NEW: <list> with optional datatype
        # -----------------------------
        elif tag == "list":
            name = elem.get("name")
            if not name:
                raise ValueError("<list> missing required 'name' attribute")

            datatype = elem.get("datatype", "string")
            raw_items = [(child.text or "") for child in elem.findall("item")]

            # Convert items based on datatype
            if datatype == "int":
                items = []
                for raw in raw_items:
                    if raw == "":
                        items.append(0)
                    else:
                        try:
                            items.append(int(float(raw)))
                        except (ValueError, TypeError):
                            items.append(0)

            elif datatype == "float":
                items = []
                for raw in raw_items:
                    if raw == "":
                        items.append(0.0)
                    else:
                        try:
                            items.append(float(raw))
                        except (ValueError, TypeError):
                            items.append(0.0)

            elif datatype == "bool":
                items = [raw.lower() == "true" for raw in raw_items]

            elif datatype == "timestamp":
                items = []
                for raw in raw_items:
                    if raw == "":
                        items.append(pd.NaT)
                    else:
                        try:
                            # Try numeric first (Excel serial date format)
                            numeric_val = float(raw)
                            items.append(pd.to_datetime(numeric_val, unit='D', origin='1899-12-30', utc=True))
                        except (ValueError, TypeError):
                            # Fallback to string parsing
                            items.append(pd.to_datetime(raw, dayfirst=True, errors="coerce", utc=True))

            else:
                # Default: keep as strings
                items = raw_items

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
                # Try numeric first (Excel serial date format)
                try:
                    numeric_val = float(raw)
                    # Excel epoch: 1899-12-30
                    out = pd.to_datetime(numeric_val, unit='D', origin='1899-12-30', utc=True)
                except (ValueError, TypeError):
                    # Fallback to string parsing
                    out = pd.to_datetime(raw, dayfirst=True, errors="coerce", utc=True)
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
            # Convert to Excel serial date (days since 1899-12-30)
            # This preserves time components and is unambiguous
            excel_epoch = pd.Timestamp('1899-12-30', tz='UTC')
            ts = pd.Timestamp(v)
            if ts.tzinfo is None:
                ts = ts.tz_localize('UTC')
            delta = ts - excel_epoch
            return str(delta.total_seconds() / 86400)  # Convert seconds to days
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

        # Handle bar orientation: Plotly "bar" with orientation="v" (default) = vertical columns
        # Excel: xlColumnClustered = vertical, xlBarClustered = horizontal
        if trace_type == "bar":
            orientation = getattr(trace, "orientation", "v") or "v"
            if orientation.lower() == "h":
                return "bar"  # horizontal bars -> Excel xlBarClustered
            else:
                return "column"  # vertical bars -> Excel xlColumnClustered

        mapping = {
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

    def build_xml_tree(self) -> LET.ElementTree:
        root = LET.Element("plotly_excel_chart", version=self.SCHEMA_VERSION)
        self._build_chart_meta(root)
        self._build_traces(root)
        self._build_extras(root)
        return LET.ElementTree(root)

    def _build_chart_meta(self, root):
        layout = self.figure.layout
        chart_meta = LET.SubElement(root, "chart_meta")

        if len(self.figure.data) == 0:
            chart_type_value = "xy"
        else:
            first_type = self.figure.data[0].type
            chart_type_value = self._map_plotly_type_to_chart_type(first_type)
        LET.SubElement(chart_meta, "chart_type").text = chart_type_value

        LET.SubElement(chart_meta, "title").text = self._safe_text(
            layout.title.text if layout.title else ""
        )

        x_axis = LET.SubElement(chart_meta, "x_axis")
        LET.SubElement(x_axis, "title").text = self._safe_text(
            layout.xaxis.title.text if layout.xaxis.title else ""
        )
        LET.SubElement(x_axis, "min").text = self._safe_text(
            layout.xaxis.range[0] if layout.xaxis.range else ""
        )
        LET.SubElement(x_axis, "max").text = self._safe_text(
            layout.xaxis.range[1] if layout.xaxis.range else ""
        )
        LET.SubElement(x_axis, "log_scale").text = (
            "true" if layout.xaxis.type == "log" else "false"
        )

        y_axis = LET.SubElement(chart_meta, "y_axis")
        LET.SubElement(y_axis, "title").text = self._safe_text(
            layout.yaxis.title.text if layout.yaxis.title else ""
        )
        LET.SubElement(y_axis, "min").text = self._safe_text(
            layout.yaxis.range[0] if layout.yaxis.range else ""
        )
        LET.SubElement(y_axis, "max").text = self._safe_text(
            layout.yaxis.range[1] if layout.yaxis.range else ""
        )
        LET.SubElement(y_axis, "log_scale").text = (
            "true" if layout.yaxis.type == "log" else "false"
        )

        legend = LET.SubElement(chart_meta, "legend")
        legend_visible = getattr(layout.legend, "visible", True)
        legend_orientation = getattr(layout.legend, "orientation", "v")
        LET.SubElement(legend, "visible").text = (
            "true" if bool(legend_visible) else "false"
        )
        LET.SubElement(legend, "position").text = (
            "right" if legend_orientation == "v" else "bottom"
        )

        # NEW: barmode for stacked/grouped
        LET.SubElement(chart_meta, "barmode").text = getattr(layout, "barmode", "group")

        timestamp = dt.datetime.utcnow().isoformat(timespec="microseconds") + "+00:00"
        LET.SubElement(chart_meta, "export_timestamp").text = timestamp

    def _build_traces(self, root):
        traces_el = LET.SubElement(root, "traces")

        for idx, trace in enumerate(self.figure.data, start=1):
            trace_el = LET.SubElement(traces_el, "trace", id=str(idx))
            data_el = LET.SubElement(trace_el, "data")

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

                LET.SubElement(data_el, "x").text = self._comma_join(bin_midpoints)
                LET.SubElement(data_el, "y").text = self._comma_join(counts)
                LET.SubElement(data_el, "text").text = self._comma_join(bin_labels)

            elif trace.type == "bar" and (not hasattr(trace, "y") or trace.y is None or len(trace.y) == 0):
                counts, bin_edges = np.histogram(trace.x, bins="auto")
                x_labels = [f"{bin_edges[i]}-{bin_edges[i+1]}" for i in range(len(bin_edges)-1)]
                LET.SubElement(data_el, "x").text = self._comma_join(x_labels)
                LET.SubElement(data_el, "y").text = self._comma_join(counts)

            # Handle bar orientation: Plotly horizontal bars have x=values, y=categories
            # But VBA expects x=categories, y=values for categorical chart types
            elif trace.type == "bar":
                orientation = getattr(trace, "orientation", "v") or "v"
                if orientation.lower() == "h":
                    # Horizontal bars: swap x and y so VBA gets x=categories, y=values
                    LET.SubElement(data_el, "x").text = self._comma_join(getattr(trace, "y", []))
                    LET.SubElement(data_el, "y").text = self._comma_join(getattr(trace, "x", []))
                else:
                    LET.SubElement(data_el, "x").text = self._comma_join(getattr(trace, "x", []))
                    LET.SubElement(data_el, "y").text = self._comma_join(getattr(trace, "y", []))

            else:
                LET.SubElement(data_el, "x").text = self._comma_join(getattr(trace, "x", []))
                LET.SubElement(data_el, "y").text = self._comma_join(getattr(trace, "y", []))
                LET.SubElement(data_el, "z").text = self._comma_join(getattr(trace, "z", []))

                marker_size_values = getattr(trace.marker, "size", None) if hasattr(trace, "marker") else None
                if isinstance(marker_size_values, (list, tuple)):
                    marker_size_str = self._comma_join(marker_size_values)
                elif marker_size_values is not None:
                    marker_size_str = str(marker_size_values)
                else:
                    marker_size_str = ""
                LET.SubElement(data_el, "size").text = marker_size_str

                LET.SubElement(data_el, "text").text = self._comma_join(getattr(trace, "text", []))

                # NEW: marker.opacity array support
                marker_opacity = getattr(trace.marker, "opacity", None) if hasattr(trace, "marker") else None
                if isinstance(marker_opacity, (list, tuple)):
                    LET.SubElement(data_el, "opacity").text = self._comma_join(marker_opacity)
                elif marker_opacity is not None:
                    LET.SubElement(data_el, "opacity").text = str(marker_opacity)

            style_el = LET.SubElement(trace_el, "style")
            LET.SubElement(style_el, "series_type").text = self._map_plotly_type_to_series_type(trace)
            LET.SubElement(style_el, "name").text = self._safe_text(getattr(trace, "name", f"Series {idx}"))

            # NEW: secondary axis support
            axis_group = "secondary" if getattr(trace, "yaxis", "y") != "y" else "primary"
            LET.SubElement(style_el, "axis_group").text = axis_group

            LET.SubElement(style_el, "visibility").text = "true"

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
            LET.SubElement(style_el, "line_color").text = line_color
            LET.SubElement(style_el, "line_style").text = line_style
            if line_width is not None:
                LET.SubElement(style_el, "line_width").text = line_width

            marker_el = LET.SubElement(style_el, "marker")
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
                    LET.SubElement(marker_el, "line_color").text = str(trace.marker.line.color or "")
                    LET.SubElement(marker_el, "line_width").text = str(trace.marker.line.width or "")

            LET.SubElement(marker_el, "size").text = marker_size
            LET.SubElement(marker_el, "color").text = marker_color
            LET.SubElement(marker_el, "shape").text = marker_shape

            # For bar/column/histogram traces, use marker.color as fill color if fillcolor is not set
            fill_color_value = getattr(trace, "fillcolor", "") or ""
            if not fill_color_value and hasattr(trace, "marker") and trace.marker:
                marker_color_val = getattr(trace.marker, "color", None)
                # Only use marker.color if it's a single color (not an array)
                if marker_color_val is not None and not isinstance(marker_color_val, (list, tuple)):
                    fill_color_value = str(marker_color_val)

            # If still no fill color for bar/column/histogram, use a default color from Plotly's palette
            series_type = self._map_plotly_type_to_series_type(trace)
            if not fill_color_value and series_type in ("bar", "column", "histogram"):
                # Plotly's default color sequence (D3 category10)
                default_colors = [
                    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
                    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"
                ]
                fill_color_value = default_colors[(idx - 1) % len(default_colors)]

            LET.SubElement(style_el, "fill_color").text = self._safe_text(fill_color_value)
            LET.SubElement(style_el, "fill_opacity").text = self._safe_text(getattr(trace, "opacity", ""))

            # NEW: error bars
            for axis in ["x", "y"]:
                err_attr = f"error_{axis}"
                if hasattr(trace, err_attr) and getattr(trace, err_attr).visible:
                    err = getattr(trace, err_attr)
                    err_el = LET.SubElement(trace_el, f"error_{axis}")
                    LET.SubElement(err_el, "type").text = getattr(err, "type", "data")
                    LET.SubElement(err_el, "symmetric").text = str(getattr(err, "symmetric", True)).lower()
                    LET.SubElement(err_el, "array").text = self._comma_join(getattr(err, "array", []))
                    if hasattr(err, "arrayminus"):
                        LET.SubElement(err_el, "arrayminus").text = self._comma_join(getattr(err, "arrayminus", []))

    def _build_extras(self, root):
        extras_el = LET.SubElement(root, "extras")

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

                ann_el = LET.SubElement(extras_el, "annotation", id=str(idx), type=ann_type)
                LET.SubElement(ann_el, "axis").text = axis
                LET.SubElement(ann_el, "value").text = value
                span_el = LET.SubElement(ann_el, "span", axis=span_axis)
                span_el.set("mode", span_mode)

                style_el = LET.SubElement(ann_el, "style")
                if hasattr(shape, "line") and shape.line:
                    LET.SubElement(style_el, "color").text = self._safe_text(shape.line.color)
                    LET.SubElement(style_el, "width").text = self._safe_text(shape.line.width)
                    LET.SubElement(style_el, "dash").text = self._safe_text(getattr(shape.line, "dash", "solid"))
                if hasattr(shape, "opacity"):
                    LET.SubElement(style_el, "opacity").text = self._safe_text(shape.opacity)

                if hasattr(shape, "name"):
                    LET.SubElement(ann_el, "label").text = self._safe_text(shape.name)

    def save_to_file(self, output_path: str):
        tree = self.build_xml_tree()
        tree.write(output_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")

def figure_to_excel_xml(fig: Figure, output_path: str):
    converter = PlotlyToExcelXMLConverter(fig)
    converter.save_to_file(output_path)
