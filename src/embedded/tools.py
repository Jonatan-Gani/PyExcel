import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
import xmlParsing as xmlp # Internal XML parsing utilities

import os
# import re
import time
import sys
import traceback
import argparse
import threading
import shutil
import hashlib
# import base64
# from dataclasses import dataclass, field

from typing import Optional, Dict, Callable, Mapping, Any, Tuple, List
from matplotlib.figure import Figure
from matplotlib.axes import Axes
from datetime import datetime

# Imports not used directly in this file but re-exported
from xmlParsing import ExcelFormula, excel_formula


def _mpl_as_figure(obj: Any):
    """Return (fig, axes_or_none) if obj is matplotlib-like, else (None, None)."""
    try:
        # Figure
        if isinstance(obj, Figure):
            return obj, None
        # Axes -> Figure
        if isinstance(obj, Axes):
            return obj.get_figure(), obj
        # Common cases: pyplot, artists exposing 'figure' attr, etc.
        if hasattr(obj, "figure") and isinstance(obj.figure, Figure):
            return obj.figure, None
        # Objects with gcf/gca semantics (rare in modern code)
        if hasattr(obj, "gcf"):
            fig = obj.gcf()
            if isinstance(fig, Figure):
                return fig, None
    except Exception:
        pass
    return None, None

def _ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)

def _hash_file(path: str, algo: str = "sha256") -> str:
    h = hashlib.new(algo)
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def _file_info(path: str) -> Tuple[int, str]:
    return os.path.getsize(path), _hash_file(path, "sha256")

def _rel_href(absolute_path: str, meta_dir: str) -> str:
    return os.path.normpath(os.path.relpath(absolute_path, start=meta_dir)).replace("\\", "/")

def _mime_for(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    return {
        ".xml": "application/xml",
        ".txt": "text/plain",
        ".emf": "image/x-emf",
        ".svg": "image/svg+xml",
        ".png": "image/png",
        ".csv": "text/csv",
    }.get(ext, "application/octet-stream")

def _write_xml_atomic(path: str, root_elem: ET.Element) -> None:
    data = ET.tostring(root_elem, encoding="utf-8", xml_declaration=True)
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    os.replace(tmp, path)

# ----------------------------
# Plot adapters -> ChartSpec
# ----------------------------

def _is_matplotlib_fig(obj: Any) -> bool:
    try:
        from matplotlib.figure import Figure
        from matplotlib.axes import Axes
        return isinstance(obj, (Figure, Axes)) or (obj.__class__.__module__.startswith("matplotlib") and hasattr(obj, "savefig"))
    except Exception:
        return False

def _is_plotly_fig(obj: Any) -> bool:
    return obj.__class__.__module__.startswith("plotly") and hasattr(obj, "to_dict")


# ----------------------------
# Writers for list/value artifacts (independent XML files)
# ----------------------------

def write_list_xml(path: str, items: List[Any], item_tag: str = "item") -> None:
    root = ET.Element("list")
    for x in items:
        e = ET.SubElement(root, item_tag)
        e.text = "" if x is None else str(x)
    _ensure_dir(os.path.dirname(path) or ".")
    _write_xml_atomic(path, root)

def write_value_xml(path: str, value: Any, datatype: str = "") -> None:
    root = ET.Element("value")
    if datatype:
        root.set("datatype", datatype)
    root.text = "" if value is None else str(value)
    _ensure_dir(os.path.dirname(path) or ".")
    _write_xml_atomic(path, root)

# ----------------------------
# Meta/status XML
# ----------------------------

def _build_meta_xml(payload: Dict[str, Any], artifacts: List[Dict[str, Any]] = None) -> ET.Element:
    root = ET.Element("meta", {"version": "1.1"})
    def add(tag, text):
        e = ET.SubElement(root, tag)
        e.text = "" if text is None else str(text)

    add("run_id", payload.get("run_id", ""))
    add("status", payload.get("status", ""))
    add("timestamp", payload.get("timestamp", ""))
    if "duration" in payload: add("duration", payload.get("duration"))
    if "message" in payload: add("message", payload.get("message"))
    if "stderr"  in payload: add("stderr",  payload.get("stderr"))

    if artifacts:
        arts = ET.SubElement(root, "artifacts")
        for a in artifacts:
            # required
            attrs = {
                "type": a["type"],
                "id": a["id"],
                "href": a["href"],
            }
            # optional
            if a.get("mime"):   attrs["mime"] = a["mime"]
            if a.get("bytes") is not None: attrs["bytes"] = str(a["bytes"])
            if a.get("sha256"): attrs["sha256"] = a["sha256"]
            ET.SubElement(arts, "artifact", attrs)

    return root

def write_answer(path: str, status: str, message: str = "", duration: float = 0.0,
                 stderr: str = "", data_file: str = "", run_id: str = "",
                 artifacts: List[Dict[str, Any]] = None) -> None:
    print(path)
    payload = {
        "run_id": run_id,
        "status": status,
        "timestamp": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "message": message,
        "duration": round(duration, 3),
        "stderr": stderr,
        # retained for compatibility with your old signature; not written to XML
        "data_file": data_file,
    }
    root = _build_meta_xml(payload, artifacts=artifacts)
    _write_xml_atomic(path, root)

def write_in_progress_status(path: str, message: str = "in_progress", run_id: str = "") -> None:
    print(path)
    payload = {
        "run_id": run_id,
        "status": "in_progress",
        "timestamp": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "message": message,
    }
    root = _build_meta_xml(payload, artifacts=None)
    _write_xml_atomic(path, root)

# ----------------------------
# Heartbeat
# ----------------------------

def _start_heartbeat(meta_file: str, run_id: str, interval: float = 10.0):
    stop_event = threading.Event()
    def heartbeat_loop():
        while not stop_event.wait(interval):
            write_in_progress_status(meta_file, run_id=run_id)
    thread = threading.Thread(target=heartbeat_loop, daemon=True)
    thread.start()
    return stop_event

# ----------------------------
# Artifact materialization
# ----------------------------

def _artifact_paths(out_file: str, meta_dir: str) -> Dict[str, str]:
    # All artifacts live in the same temp dir where meta.xml is
    temp_dir = meta_dir

    stem = os.path.splitext(os.path.basename(out_file))[0] or "out"
    return {
        "tables_dir": os.path.join(temp_dir, "tables"),
        "lists_dir":  os.path.join(temp_dir, "lists"),
        "values_dir": os.path.join(temp_dir, "values"),
        "assets_dir": os.path.join(temp_dir, "assets"),
        "out_file":   os.path.join(temp_dir, f"{stem}.xml"),
    }
    
    
def _materialize_outputs(
    out_file: str,
    meta_dir: str,
    result: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """
    Enforces the dict-based return protocol.
    Returns a list of artifact dicts: {type, id, path, href, mime, bytes, sha256}.
    """
    if not isinstance(result, dict):
        raise TypeError("Transform must return a dict[str, Any].")

    # paths = _artifact_paths(out_file)
    paths = _artifact_paths(out_file, meta_dir)
    artifacts: List[Dict[str, Any]] = []

    def reg(path: str, typ: str, id_: str) -> None:
        size, sha = _file_info(path)
        artifacts.append({
            "type": typ,
            "id": id_,
            "path": path,
            "href": _rel_href(path, meta_dir),
            "mime": _mime_for(path),
            "bytes": size,
            "sha256": sha,
        })

    # ensure dirs
    _ensure_dir(paths["tables_dir"])
    _ensure_dir(paths["lists_dir"])
    _ensure_dir(paths["values_dir"])
    _ensure_dir(paths["assets_dir"])

    for key, val in result.items():
        # DataFrame or Series
        if isinstance(val, (pd.DataFrame, pd.Series)):
            if isinstance(val, pd.Series):
                val = val.to_frame(name=val.name or "value")  # ensure one-column DataFrame with a column name

            table_path = os.path.join(paths["tables_dir"], f"{key}.xml")
            xmlp.write_xml(table_path, val)
            reg(table_path, "table", key)

        # Dict[str, DataFrame]
        elif isinstance(val, Mapping) and all(isinstance(x, pd.DataFrame) for x in val.values()):
            table_path = os.path.join(paths["tables_dir"], f"{key}.xml")
            xmlp.write_xml(table_path, val)  # keys become table names
            reg(table_path, "table", key)

        # List/Tuple
        elif isinstance(val, (list, tuple)):
            if all(isinstance(x, pd.DataFrame) for x in val):
                # list of DataFrames → single XML with df1, df2, ...
                table_path = os.path.join(paths["tables_dir"], f"{key}.xml")
                xmlp.write_xml(table_path, list(val))
                reg(table_path, "table", key)
            elif all(isinstance(x, (int, float, str)) or x is None for x in val):
                # list of scalars
                list_path = os.path.join(paths["lists_dir"], f"{key}.xml")
                write_list_xml(list_path, list(val))
                reg(list_path, "list", key)
            else:
                raise TypeError(
                    f"List for key '{key}' contains mixed or unsupported types: {[type(x) for x in val]}"
                )

        # Scalar
        elif isinstance(val, (int, float, str)):
            val_path = os.path.join(paths["values_dir"], f"{key}.xml")
            dt = "integer" if isinstance(val, int) else "decimal" if isinstance(val, float) else "string"
            write_value_xml(val_path, val, datatype=dt)
            reg(val_path, "value", key)

        # Chart objects
        # elif hasattr(val, "savefig"):  # matplotlib
        #     dst = os.path.join(paths["assets_dir"], f"{key}.svg")
        #     val.savefig(dst, format="svg")
        #     reg(dst, "chart", key)
         # Chart-like objects (matplotlib / plotly)
        # elif hasattr(val, "savefig") or _is_matplotlib_fig(val) or _is_plotly_fig(val):
        #     # 1) Prefer ChartSpec (plot2.0)
        #     spec = None
        #     if _is_plotly_fig(val):
        #         spec = _chartspec_from_plotly(val)
        #     if spec is None and _is_matplotlib_fig(val):
        #         spec = _chartspec_from_matplotlib(val)
        #     # Try to normalize matplotlib objects (Axes -> Figure) if needed
        #     if spec is None:
        #         fig = None
        #         try:
        #             # normalize to a Figure for matplotlib fallbacks
        #             fig, _axes = _mpl_as_figure(val)  # helper: returns (Figure or None, Axes or None)
        #         except Exception:
        #             fig = None
        #         if fig is not None:
        #             spec = _chartspec_from_matplotlib(fig)

        #     if spec:
        #         chart_path = os.path.join(paths["assets_dir"], f"{key}.xml")
        #         xmlp.write_chart_xml(chart_path, spec)
        #         size, sha = _file_info(chart_path)
        #         artifacts.append({
        #             "type": "plot2.0",
        #             "id": key,
        #             "path": chart_path,
        #             "href": _rel_href(chart_path, meta_dir),
        #             "mime": "application/xml",
        #             "bytes": size,
        #             "sha256": sha,
        #         })
        #     else:
        #         # 2) Fallback: persist as image correctly per backend
        #         dst_svg = os.path.join(paths["assets_dir"], f"{key}.svg")
        #         written = None
        #         if _is_plotly_fig(val) or hasattr(val, "write_image"):
        #             try:
        #                 val.write_image(dst_svg, format="svg")
        #                 written = dst_svg
        #             except Exception:
        #                 dst_png = os.path.splitext(dst_svg)[0] + ".png"
        #                 val.write_image(dst_png, format="png")
        #                 written = dst_png
        #         else:
        #             # matplotlib (normalize to Figure)
        #             fig, _axes = _mpl_as_figure(val)
        #             if fig is None and hasattr(val, "savefig"):
        #                 # last resort: object itself exposes savefig
        #                 fig = val
        #             if fig is None:
        #                 raise TypeError(f"Object of type {type(val)} is not a supported plot.")
        #             try:
        #                 fig.savefig(dst_svg, format="svg")
        #                 written = dst_svg
        #             except Exception:
        #                 dst_png = os.path.splitext(dst_svg)[0] + ".png"
        #                 fig.savefig(dst_png, format="png", dpi=144)
        #                 written = dst_png

        #         reg(written, "chart", key)

        elif hasattr(val, "savefig") or _is_matplotlib_fig(val) or _is_plotly_fig(val):
            # New: Plotly → Excel XML (treated as plot2.0)
            if _is_plotly_fig(val):
                chart_path = os.path.join(paths["assets_dir"], f"{key}.xml")

                conv = xmlp.PlotlyToExcelXMLConverter(val)
                tree = conv.build_xml_tree()
                tree.write(chart_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")

                size, sha = _file_info(chart_path)
                artifacts.append({
                    "type": "plot2.0",
                    "id": key,
                    "path": chart_path,
                    "href": _rel_href(chart_path, meta_dir),
                    "mime": "application/xml",
                    "bytes": size,
                    "sha256": sha,
                })
                continue

            # Matplotlib fallback → export image
            dst_svg = os.path.join(paths["assets_dir"], f"{key}.svg")
            written = None

            fig, _axes = _mpl_as_figure(val)
            if fig is None and hasattr(val, "savefig"):
                fig = val
            if fig is None:
                raise TypeError(f"Object of type {type(val)} is not a supported plot.")

            try:
                fig.savefig(dst_svg, format="svg")
                written = dst_svg
            except Exception:
                dst_png = os.path.splitext(dst_svg)[0] + ".png"
                fig.savefig(dst_png, format="png", dpi=144)
                written = dst_png

            reg(written, "chart", key)




        # Explicit asset path (now also supports ChartSpec .xml)
        elif isinstance(val, str) and os.path.exists(val):
            ext = os.path.splitext(val)[1].lower()
            if ext not in (".emf", ".svg", ".png", ".xml"):
                raise ValueError(f"Unsupported chart format for '{key}': {ext}. Expected .emf/.svg/.png/.xml")
            dst = os.path.join(paths["assets_dir"], f"{key}{ext}")
            if os.path.abspath(val) != os.path.abspath(dst):
                shutil.copy2(val, dst)
            if ext == ".xml":
                size, sha = _file_info(dst)
                artifacts.append({
                    "type": "plot2.0",
                    "id": key,
                    "path": dst,
                    "href": _rel_href(dst, meta_dir),
                    "mime": "application/xml",
                    "bytes": size,
                    "sha256": sha,
                })
            else:
                reg(dst, "chart", key)

        elif hasattr(val, "write_image"):  # plotly
            dst = os.path.join(paths["assets_dir"], f"{key}.svg")
            val.write_image(dst, format="svg")
            reg(dst, "chart", key)

        elif isinstance(val, str) and os.path.exists(val):
            ext = os.path.splitext(val)[1].lower()
            if ext not in (".emf", ".svg", ".png"):
                raise ValueError(f"Unsupported chart format for '{key}': {ext}. Expected .emf/.svg/.png")
            dst = os.path.join(paths["assets_dir"], f"{key}{ext}")
            if os.path.abspath(val) != os.path.abspath(dst):
                shutil.copy2(val, dst)
            reg(dst, "chart", key)

        else:
            raise TypeError(f"Unsupported return type for key '{key}': {type(val)}")

    return artifacts

# ----------------------------
# Runner
# ----------------------------

def run_script(
    transform: Callable[[List[pd.DataFrame]], Any],
    *,
    in_file: str,
    out_file: str,
    meta_file: str,
    run_id: str,
    heartbeat_interval: float = 10.0
):
    """
    Run a transform function that consumes a list of input DataFrames
    and produces either:
      - a dict[str, Any] of artifacts
      - a single value
      - a tuple/list of values
    Non-dict outputs are wrapped in a dict with generic keys like "result_0", "result_1", etc.
    """
    start_time = time.time()
    write_in_progress_status(meta_file, message="Script started", run_id=run_id)
    heartbeat_stop = _start_heartbeat(meta_file, run_id=run_id, interval=heartbeat_interval)

    try:
        df_in = xmlp.read_xml(in_file)  # type: ignore[name-defined]
        raw_result = transform(df_in)

        # Normalize result to a dict
        if isinstance(raw_result, dict):
            result = raw_result
        elif isinstance(raw_result, (list, tuple, set)):
            # preserve order for list/tuple, arbitrary for set
            result = {f"result_{i}": v for i, v in enumerate(raw_result)}
        else:
            # single value
            result = {"result_0": raw_result}

        meta_dir = os.path.dirname(os.path.abspath(meta_file)) or "."
        artifacts = _materialize_outputs(out_file, meta_dir, result)

        duration = time.time() - start_time
        heartbeat_stop.set()
        write_answer(
            path=meta_file,
            status="done",
            message="Success",
            duration=duration,
            run_id=run_id,
            artifacts=artifacts
        )
    except Exception as e:
        duration = time.time() - start_time
        heartbeat_stop.set()
        
        # Format exception to include type, e.g. "KeyError: 'col'" instead of just "'col'"
        msg_str = f"{type(e).__name__}: {str(e)}"
        
        write_answer(
            path=meta_file,
            status="error",
            message=msg_str,
            duration=duration,
            stderr=traceback.format_exc(),
            run_id=run_id,
            artifacts=None
        )
        sys.exit(1)
        

def run_script_cli(
    transform: Callable[[List[pd.DataFrame]], Dict[str, Any]],
    heartbeat_interval: float = 10.0
):
    """
    Command-line entrypoint.
    `transform` must return a dict[str, Any] according to the standard protocol:
      - DataFrame
      - list/tuple of DataFrames
      - list/tuple of scalars
      - scalar
      - figure-like object (matplotlib/plotly) or path to .emf/.svg/.png
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="in_file", required=True)
    parser.add_argument("--out", dest="out_file", required=True)
    parser.add_argument("--meta", dest="meta_file", required=True)
    parser.add_argument("--run-id", dest="run_id", required=True)
    args = parser.parse_args()

    run_script(
        transform,
        in_file=args.in_file,
        out_file=args.out_file,
        meta_file=args.meta_file,
        run_id=args.run_id,
        heartbeat_interval=heartbeat_interval
    )



# legacy code


# _ANCHOR_RE = re.compile(r'^\$?[A-Z]+\$?\d+$')

# @dataclass(frozen=True)
# class ExcelFormula:
#     mode: str                   # "a1" or "r1c1"
#     a1: Optional[str] = None
#     anchor: Optional[str] = None
#     r1c1: Optional[str] = None
#     # arbitrary optional attributes for future use
#     attrs: Dict[str, str] = field(default_factory=dict)


# def excel_formula(a1: str, anchor: str = None, **options) -> ExcelFormula:
#     """
#     Create an A1-mode ExcelFormula without requiring an anchor.
#     If an anchor is provided, it is validated then ignored (deprecated).
#     """
#     # Minimal validation
#     if not isinstance(a1, str) or not a1.startswith("="):
#         raise ValueError("a1 must start with '='.")

#     # Back-compat: accept but ignore anchor
#     if anchor is not None:
#         if not isinstance(anchor, str) or not _ANCHOR_RE.match(anchor):
#             raise ValueError("anchor must be a valid A1 like 'E7' or '$E$7'.")
#         # intentionally ignored

#     # Normalize common options to strings
#     def b(x): return "1" if bool(x) else "0"
#     normalized: Dict[str, str] = {}
#     for k, v in options.items():
#         if v is None:
#             continue
#         if k in {"spill", "volatile", "protect"}:
#             normalized[k] = b(v)
#         else:
#             normalized[k] = str(v)

#     # NOTE: no 'anchor' in the payload anymore
#     return ExcelFormula(mode="a1", a1=a1, attrs=normalized)


# ----------------------------
# XML to and from DataFrames
# ----------------------------

# def read_xml(path: str) -> dict[str, pd.DataFrame]:
#     tables: dict[str, pd.DataFrame] = {}

#     # State for the current <table>
#     table_name = None
#     col_names: list[str] | None = None
#     col_types: list[str] | None = None
#     col_buffers: list[list[str]] | None = None

#     # Stream parse: handle 'end' events so children text is available
#     for event, elem in LET.iterparse(path, events=("end",)):
#         tag = elem.tag

#         if tag == "columns" and elem.getparent() is not None and elem.getparent().tag == "table":
#             # Collect column metadata
#             cols = elem.findall("col")
#             col_names = [c.get("name", "") for c in cols]
#             col_types = [c.get("type", "string") for c in cols]
#             col_buffers = [[] for _ in col_names]  # one list per column

#             # Free memory: drop the <columns> subtree
#             elem.clear()
#             parent = elem.getparent()
#             if parent is not None:
#                 while elem.getprevious() is not None:
#                     del parent[0]

#         elif tag == "row":
#             # Append cell text into column-wise buffers (pad/truncate to metadata length)
#             if col_buffers is None or col_names is None:
#                 raise ValueError("Encountered <row> before <columns>.")

#             cells = elem.findall("col")
#             n = len(col_names)
#             # Extract text (empty string if missing)
#             row_vals = [(c.text or "") for c in cells]
#             if len(row_vals) < n:
#                 row_vals.extend([""] * (n - len(row_vals)))
#             elif len(row_vals) > n:
#                 row_vals = row_vals[:n]

#             for j, v in enumerate(row_vals):
#                 col_buffers[j].append(v)

#             # Clear the processed row to keep memory low
#             elem.clear()
#             parent = elem.getparent()
#             if parent is not None:
#                 while elem.getprevious() is not None:
#                     del parent[0]

#         elif tag == "table":
#             # Closing a table: build DataFrame from buffers and cast types
#             table_name = elem.get("name", "")
#             if not table_name:
#                 raise ValueError("Table missing 'name' attribute")
#             if col_names is None or col_types is None or col_buffers is None:
#                 raise ValueError(f"Malformed table '{table_name}': missing columns/rows")

#             # Construct DataFrame column-wise (avoids list-of-lists materialization)
#             df_dict: dict[str, pd.Series] = {}
#             for cname, ctype, raw in zip(col_names, col_types, col_buffers):
#                 s = pd.Series(raw, copy=False)

#                 if ctype == "int":
#                     empty_mask = s.eq("")
#                     num = pd.to_numeric(s.mask(empty_mask), errors="coerce")
#                     num[empty_mask] = 0

#                     # Try safe cast: if any non-integer floats are present, leave as float
#                     try:
#                         df_dict[cname] = num.astype("Int64")
#                     except TypeError:
#                         # Debug print to show which column broke and sample of values
#                         print(f"[read_xml] Column '{cname}' contains non-integers, keeping as float")
#                         print(num.head(20).to_list())
#                         df_dict[cname] = num

#                 elif ctype == "float":
#                     # Empty -> 0.0; invalid -> NaN
#                     empty_mask = s.eq("")
#                     num = pd.to_numeric(s.mask(empty_mask), errors="coerce")
#                     num[empty_mask] = 0.0
#                     df_dict[cname] = num

#                 elif ctype == "bool":
#                     df_dict[cname] = s.str.lower().eq("true")

#                 elif ctype == "timestamp":
#                     # ISO-8601 text with Z suffix -> datetime64[ns, UTC]
#                     df_dict[cname] = pd.to_datetime(s, errors="coerce", utc=True)

#                 elif ctype == "blank":
#                     df_dict[cname] = pd.Series(pd.NA, index=s.index)

#                 else:
#                     df_dict[cname] = s.astype(str)

#             tables[table_name] = pd.DataFrame(df_dict, columns=col_names)

#             # Clear the processed table subtree and reset state
#             elem.clear()
#             parent = elem.getparent()
#             if parent is not None:
#                 while elem.getprevious() is not None:
#                     del parent[0]
#             table_name = None
#             col_names = None
#             col_types = None
#             col_buffers = None

#     return tables

# def write_xml(path: str,
#               tables: Union[pd.DataFrame,
#                             Sequence[pd.DataFrame],
#                             Mapping[str, pd.DataFrame]]):
#     try:
#         ExcelFormulaBase = ExcelFormula  # type: ignore[name-defined]
#     except NameError:
#         ExcelFormulaBase = ()  # type: ignore[assignment]

#     def infer_column_type(s: pd.Series) -> str:
#         ss = s.dropna()
#         if ss.empty:
#             return "blank"
#         ts = ss.map(type).unique()
#         if all(t is bool for t in ts):
#             return "bool"
#         if all(issubclass(t, (int, np.integer)) for t in ts):
#             return "int"
#         if all(issubclass(t, (float, int, np.floating, np.integer)) for t in ts):
#             return "float"
#         if all(isinstance(v, (datetime, pd.Timestamp)) for v in ss):
#             return "date"
#         return "string"

#     def is_formula_obj(v) -> bool:
#         if isinstance(v, ExcelFormulaBase):
#             return True
#         if isinstance(v, dict) and ("a1" in v or "r1c1" in v):
#             return True
#         return False

#     def extract_meta(s: pd.Series) -> Optional[Dict[str, str]]:
#         for v in s:
#             if isinstance(v, ExcelFormulaBase):
#                 meta: Dict[str, str] = {"mode": v.mode}
#                 if v.mode == "a1":
#                     meta["a1"] = v.a1  # type: ignore[arg-type]
#                 elif v.mode == "r1c1":
#                     meta["r1c1"] = v.r1c1  # type: ignore[arg-type]
#                 for k, vv in (getattr(v, "attrs", None) or {}).items():
#                     meta[k] = vv if isinstance(vv, str) else str(vv)
#                 return meta
#         for v in s:
#             if isinstance(v, dict) and ("a1" in v or "r1c1" in v):
#                 meta = {k: (vv if isinstance(vv, str) else str(vv)) for k, vv in v.items()}
#                 if "mode" not in meta:
#                     meta["mode"] = "a1" if "a1" in meta else "r1c1"
#                 if meta.get("mode") == "a1":
#                     meta.pop("anchor", None)
#                 return meta
#         for v in s:
#             if isinstance(v, str) and v.startswith("="):
#                 return {"mode": "r1c1", "r1c1": v}
#         return None

#     def serialize_value(v):
#         if v is None or pd.isna(v):
#             return None
#         if is_formula_obj(v):
#             return None
#         if isinstance(v, str):
#             if v.startswith("=") or v == "":
#                 return None
#             return v
#         if isinstance(v, bool):
#             return str(v).lower()
#         if isinstance(v, (int, float, np.integer, np.floating)):
#             return str(v)
#         if isinstance(v, (datetime, pd.Timestamp)):
#             return v.strftime("%Y-%m-%d")
#         return str(v)

#     if isinstance(tables, pd.DataFrame):
#         name_df_pairs = [("df1", tables)]
#     elif isinstance(tables, Mapping):
#         name_df_pairs = [(str(name), df) for name, df in tables.items()]
#     else:
#         name_df_pairs = [(f"df{i+1}", df) for i, df in enumerate(tables)]

#     root = LET.Element("data")

#     for name, df in name_df_pairs:
#         table_el = LET.SubElement(root, "table", attrib={"name": name})

#         cols = LET.SubElement(table_el, "columns")
#         for col in df.columns:
#             s = df[col]
#             attrib: Dict[str, str] = {"name": str(col), "type": infer_column_type(s)}
#             meta = extract_meta(s)
#             if meta:
#                 meta.pop("anchor", None)
#                 attrib.update({k: v if isinstance(v, str) else str(v) for k, v in meta.items()})
#             LET.SubElement(cols, "col", attrib=attrib)

#         rows = LET.SubElement(table_el, "rows")
#         for _, r in df.iterrows():
#             row = LET.SubElement(rows, "row")
#             for v in r:
#                 cell = LET.SubElement(row, "col")
#                 sv = serialize_value(v)
#                 if sv is not None:
#                     cell.text = sv

#     tree = LET.ElementTree(root)
#     tree.write(path, encoding="utf-8", xml_declaration=True, pretty_print=False)


# def write_chart_xml(path: str, chart_spec: Dict[str, Any]) -> None:
    # root = ET.Element("ChartSpec", {"version": "1.0", "xmlns": "urn:example:chartspec:1.0"})

    # chart = ET.SubElement(root, "chart", {"type": chart_spec.get("type", "line")})
    # if cat := chart_spec.get("categoryAxis"):
        # ET.SubElement(chart, "categoryAxis", {k: str(v) for k, v in cat.items()})
    # if val := chart_spec.get("valueAxis"):
        # ET.SubElement(chart, "valueAxis", {k: str(v) for k, v in val.items()})

    # series_root = ET.SubElement(root, "series")
    # for s in chart_spec.get("series", []):
        # s_attrs = {}
        # if "name" in s: s_attrs["name"] = str(s["name"])
        # s_elem = ET.SubElement(series_root, "s", s_attrs)

        # x_elem = ET.SubElement(s_elem, "x")
        # for xv in s.get("x", []):
            # ET.SubElement(x_elem, "n").text = "" if xv is None else str(xv)

        # y_elem = ET.SubElement(s_elem, "y")
        # for yv in s.get("y", []):
            # ET.SubElement(y_elem, "n").text = "" if yv is None else str(yv)

        # if st := s.get("style"):
            # ET.SubElement(s_elem, "style", {k: str(v) for k, v in st.items()})

    # if layout := chart_spec.get("layout"):
        # layout_elem = ET.SubElement(root, "layout")
        # if "title" in layout:
            # ET.SubElement(layout_elem, "title", {"text": str(layout["title"])})

        # if legend := layout.get("legend"):
            # ET.SubElement(layout_elem, "legend", {"show": "true" if legend else "false"})

    # _ensure_dir(os.path.dirname(path) or ".")
    # _write_xml_atomic(path, root)

# _XMLNS = "urn:example:chartspec:1.0"

# def _attrs(d: Dict[str, Any]) -> Dict[str, str]:
#     return {k: "" if v is None else str(v) for k, v in d.items()}

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





# def _chartspec_from_plotly(fig: Any) -> Optional[Dict[str, Any]]:
#     """
#     Robust Plotly → ChartSpec adapter.
#     Supports: bar, histogram, scatter, scattergl (markers|lines|lines+markers).
#     Emits per-series style hints, axis assignment, axis titles (x, y, y2),
#     and sparse per-index shape overrides as multiple <shapes> groups.
#     """
#     try:
#         d = fig.to_dict()
#         traces = (d.get("data") or [])
#         layout = (d.get("layout") or {})
#         if not traces:
#             return None

#         series: List[Dict[str, Any]] = []
#         root_kinds: set[str] = set()

#         def _axis_title_from_layout(ax: Dict[str, Any]) -> str:
#             t = ax.get("title")
#             if isinstance(t, dict):
#                 return t.get("text") or ""
#             if isinstance(t, str):
#                 return t
#             return ""

#         def _is_seq(v) -> bool:
#             return isinstance(v, (list, tuple, np.ndarray, pd.Series))

#         def _to_list(v) -> Optional[List[Any]]:
#             if v is None:
#                 return None
#             if _is_seq(v):
#                 return list(v)
#             return [v]

#         def _normalize_xy(vals):
#             return [_to_serializable(vv) for vv in (vals if _is_seq(vals) else [vals])]

#         def _build_shape_groups(marker: Dict[str, Any], n: int,
#                                 def_shape: Optional[str],
#                                 def_size: Optional[Any],
#                                 def_color: Optional[str]) -> Optional[List[Dict[str, Any]]]:
#             """
#             From Plotly marker arrays, produce override groups:
#             [{"indices":[...], "shape":..., "shape_size":..., "shape_color":...}, ...]
#             Grouped by identical override tuples; missing attrs fall back to series defaults.
#             """
#             if not marker or n <= 0:
#                 return None

#             sym = _to_list(marker.get("symbol"))
#             sz  = _to_list(marker.get("size"))
#             col = _to_list(marker.get("color"))

#             # If all are scalar (len==1) or None → no groups needed
#             if (sym is None or len(sym) <= 1) and (sz is None or len(sz) <= 1) and (col is None or len(col) <= 1):
#                 return None

#             # Safe element access with bounds/None
#             def at(arr, i):
#                 if arr is None:
#                     return None
#                 return arr[i] if i < len(arr) else None

#             groups: Dict[Tuple[Tuple[str, str], ...], List[int]] = {}

#             for i in range(n):
#                 s_i = at(sym, i)
#                 z_i = at(sz, i)
#                 c_i = at(col, i)

#                 # Determine if this index wants any override. Rules:
#                 # - If any of symbol/size/color is explicitly provided at i:
#                 #     include this index; build override attrs only for provided values
#                 # - 'none' symbol is a valid override (explicit suppression)
#                 include = (s_i is not None) or (z_i is not None) or (c_i is not None)
#                 if not include:
#                     continue

#                 # Build sparse override dict; only include attributes that are set and differ from defaults.
#                 o: Dict[str, Any] = {}

#                 if s_i is not None:
#                     s_val = str(s_i)
#                     # Include even if equal to default when s_val == "none" (explicit suppression)
#                     if (def_shape is None) or (s_val != def_shape) or (s_val == "none"):
#                         o["shape"] = s_val

#                 if z_i is not None:
#                     # Size may be numeric or string; stringify later in XML
#                     if (def_size is None) or (z_i != def_size):
#                         o["shape_size"] = z_i

#                 if c_i is not None:
#                     c_val = str(c_i)
#                     if (def_color is None) or (c_val != def_color):
#                         o["shape_color"] = c_val

#                 if not o:
#                     # All provided values equal to defaults → no override needed
#                     continue

#                 # Hashable grouping key from override attrs (order-independent)
#                 key = tuple(sorted((k, str(v)) for k, v in o.items()))
#                 groups.setdefault(key, []).append(i)

#             if not groups:
#                 return None

#             out: List[Dict[str, Any]] = []
#             for key, idxs in groups.items():
#                 g: Dict[str, Any] = {"indices": idxs}
#                 for k, v in key:
#                     # v already stringified above for key; retain original semantics
#                     if k == "shape_size":
#                         # best-effort numeric; keep as-is if not numeric
#                         try:
#                             v_cast: Any = float(v) if "." in v else int(v)
#                         except Exception:
#                             v_cast = v
#                         g[k] = v_cast
#                     else:
#                         g[k] = v
#                 out.append(g)
#             return out

#         def add_series(x, y, *, name: str = "", kind: str,
#                        yaxis: str = "y",
#                        marker: Optional[Dict[str, Any]] = None,
#                        line: Optional[Dict[str, Any]] = None,
#                        error_y: Optional[Dict[str, Any]] = None,
#                        tr_level_overrides: Optional[Dict[str, Any]] = None):
#             st = {"kind": kind}

#             # Base styling (existing keys preserved)
#             if isinstance(tr_level_overrides, dict):
#                 st.update({k: "" if v is None else str(v) for k, v in tr_level_overrides.items()})

#             sx = _normalize_xy(x)
#             sy = _normalize_xy(y)

#             axis_assign = "secondary" if str(yaxis).lower() not in ("", "y", "y1") else "primary"

#             # Series-level marker defaults → map to shape/shape_size/shape_color
#             def_shape = None
#             def_size = None
#             def_color = None
#             if marker:
#                 m_sym = marker.get("symbol")
#                 m_sz  = marker.get("size")
#                 m_col = marker.get("color")
#                 if not _is_seq(m_sym) and m_sym is not None:
#                     def_shape = str(m_sym)
#                     st["shape"] = def_shape
#                 if not _is_seq(m_sz) and m_sz is not None:
#                     def_size = m_sz
#                     st["shape_size"] = str(m_sz)
#                 if not _is_seq(m_col) and m_col is not None:
#                     def_color = str(m_col)
#                     st["shape_color"] = def_color

#             # Per-index groups (when any of symbol/size/color are arrays)
#             shapes_groups = _build_shape_groups(marker or {}, n=len(sx),
#                                                 def_shape=def_shape, def_size=def_size, def_color=def_color)

#             item: Dict[str, Any] = {
#                 "name": name or "",
#                 "x": sx,
#                 "y": sy,
#                 "style": st,
#                 "axis": axis_assign,
#             }
#             if shapes_groups:
#                 # single or multiple groups supported
#                 item["shapes"] = shapes_groups if len(shapes_groups) > 1 else shapes_groups[0]

#             series.append(item)
#             root_kinds.add(kind)

#         # ---- trace harvesting ----
#         for tr in traces:
#             t = (tr.get("type") or "scatter").lower()
#             name = tr.get("name") or ""

#             # Arrays (decode plotly compressed containers)
#             x = _decode_plotly_array(tr.get("x", []))
#             y = _decode_plotly_array(tr.get("y", []))
#             if not (_is_seq(x) and len(x)):
#                 x = list(x) if x is not None else []
#             if not (_is_seq(y) and len(y)):
#                 y = list(y) if y is not None else []

#             marker = tr.get("marker") or {}
#             line = tr.get("line") or {}
#             error_y = tr.get("error_y") or {}

#             # Prefer trace-level color, then line.color, then marker.color
#             color = tr.get("marker_color") or tr.get("line_color") or line.get("color") or marker.get("color")
#             opacity = tr.get("opacity")
#             if opacity is None:
#                 opacity = marker.get("opacity")

#             extra_base = {
#                 "color": color,
#                 "marker_size": marker.get("size"),
#                 "marker_symbol": marker.get("symbol"),
#                 "opacity": opacity,
#                 "line_width": line.get("width"),
#                 "line_dash": line.get("dash"),
#                 "fill": tr.get("fill"),
#                 "fillcolor": tr.get("fillcolor"),
#                 "y_error_type": (error_y.get("type") if isinstance(error_y, dict) else None),
#                 "y_error_array_len": (
#                     len(error_y.get("array"))
#                     if isinstance(error_y, dict) and _is_seq(error_y.get("array"))
#                     else None
#                 ),
#             }

#             yaxis = tr.get("yaxis") or "y"
            
#             if t == "bar":
#                 if x and y:
#                     # ---- trace-level styling & options for bars ----
#                     if "marker_color" in tr and not extra_base["color"]:
#                         extra_base["color"] = tr["marker_color"]
#                     if "width" in tr:
#                         extra_base["bar_width"] = tr.get("width")

#                     # Orientation (vertical default 'v' vs horizontal 'h')
#                     orientation = (tr.get("orientation") or "v").lower()
#                     extra_base["orientation"] = orientation  # 'v' | 'h'

#                     # Base (baseline) and positional offset for bars
#                     if tr.get("base") is not None:
#                         extra_base["base"] = tr.get("base")  # scalar or array
#                     if tr.get("offset") is not None:
#                         extra_base["offset"] = tr.get("offset")  # scalar or array

#                     # Grouping alignment
#                     if tr.get("offsetgroup") is not None:
#                         extra_base["offsetgroup"] = tr.get("offsetgroup")
#                     if tr.get("alignmentgroup") is not None:
#                         extra_base["alignmentgroup"] = tr.get("alignmentgroup")

#                     # Pattern & gradient fills (best-effort passthrough)
#                     m = marker or {}
#                     if isinstance(m.get("pattern"), dict):
#                         # Include only well-known keys
#                         pat = {k: v for k, v in m["pattern"].items() if k in ("shape", "fgcolor", "size", "solidity", "bgcolor")}
#                         if pat:
#                             extra_base["pattern"] = pat
#                     # Plotly doesn't officially support bar gradients everywhere, but pass through if present
#                     if isinstance(m.get("gradient"), dict):
#                         extra_base["gradient"] = m["gradient"]

#                     # Border (stroke) styling
#                     ml = (m.get("line") or {})
#                     if ml.get("color") is not None:
#                         extra_base["stroke_color"] = ml.get("color")
#                     if ml.get("width") is not None:
#                         extra_base["stroke_width"] = ml.get("width")
#                     if ml.get("dash") is not None:
#                         extra_base["stroke_dash"] = ml.get("dash")

#                     # Data labels (value/category/series name) passthrough
#                     # Consumers can decide how/when to render; we expose both raw text and templates
#                     if tr.get("text") is not None:
#                         extra_base["data_labels_text"] = tr.get("text")  # str | array
#                     if tr.get("texttemplate") is not None:
#                         extra_base["data_labels_template"] = tr.get("texttemplate")
#                     if tr.get("textposition") is not None:
#                         extra_base["data_labels_position"] = tr.get("textposition")
#                     if tr.get("hovertemplate") is not None:
#                         # Provided for parity; consumer may ignore for labels
#                         extra_base["hovertemplate"] = tr.get("hovertemplate")
#                     if tr.get("name"):
#                         extra_base["series_name"] = tr.get("name")

#                     # Error bars on VALUE axis (x for horizontal bars, y for vertical bars)
#                     err_val = tr.get("error_x") if orientation == "h" else tr.get("error_y")
#                     if isinstance(err_val, dict):
#                         extra_base["value_error_type"] = err_val.get("type")
#                         extra_base["value_error_symmetric"] = err_val.get("symmetric", True)
#                         if isinstance(err_val.get("array"), (list, tuple)):
#                             extra_base["value_error_array_len"] = len(err_val.get("array"))
#                         if isinstance(err_val.get("arrayminus"), (list, tuple)):
#                             extra_base["value_error_arrayminus_len"] = len(err_val.get("arrayminus"))

#                     # Invert fill for negative values (hint only; renderer decides how to style)
#                     try:
#                         y_vals = np.asarray(y, dtype=float)
#                         if np.nanmin(y_vals) < 0 and np.nanmax(y_vals) > 0:
#                             extra_base["invert_negative_fill"] = True
#                     except Exception:
#                         pass

#                     # Gap width & series overlap (layout-level in Plotly; expose as hints here)
#                     # barmode handled elsewhere in this adapter; we surface bargap/groupgap/overlay hints
#                     try:
#                         bargap = layout.get("bargap")
#                         bargroupgap = layout.get("bargroupgap")
#                         if bargap is not None:
#                             extra_base["bargap"] = bargap  # 0..1 fraction
#                         if bargroupgap is not None:
#                             extra_base["bargroupgap"] = bargroupgap  # 0..1 fraction
#                         if layout.get("barnorm") is not None:
#                             extra_base["barnorm"] = layout.get("barnorm")  # '', 'fraction', 'percent'
#                         if layout.get("barmode") is not None:
#                             extra_base["barmode_hint"] = layout.get("barmode")  # 'group'|'stack'|'overlay'|'relative'
#                     except Exception:
#                         pass

#                     # Axis scale/min/max; tick interval/format; category order & axis crossing.
#                     # We expose VALUE vs CATEGORY axis props as hints based on orientation.
#                     try:
#                         xax_key = tr.get("xaxis") or "xaxis"
#                         yax_key = tr.get("yaxis") or "yaxis"
#                         xax = layout.get(xax_key, {})
#                         yax = layout.get(yax_key, {})

#                         # Determine which is value vs category
#                         val_ax = xax if orientation == "h" else yax
#                         cat_ax = yax if orientation == "h" else xax

#                         # Value axis: scale + explicit range
#                         if val_ax:
#                             vtype = (val_ax.get("type") or "linear").lower()
#                             extra_base["value_axis_scale"] = "log" if vtype == "log" else "linear"
#                             if isinstance(val_ax.get("range"), (list, tuple)) and len(val_ax["range"]) == 2:
#                                 extra_base["value_axis_min"] = val_ax["range"][0]
#                                 extra_base["value_axis_max"] = val_ax["range"][1]
#                             # Ticks/formatting
#                             if val_ax.get("dtick") is not None:
#                                 extra_base["value_axis_dtick"] = val_ax.get("dtick")
#                             if val_ax.get("tick0") is not None:
#                                 extra_base["value_axis_tick0"] = val_ax.get("tick0")
#                             if val_ax.get("tickformat") is not None:
#                                 extra_base["value_axis_tickformat"] = val_ax.get("tickformat")
#                             if val_ax.get("tickprefix") is not None:
#                                 extra_base["value_axis_tickprefix"] = val_ax.get("tickprefix")
#                             if val_ax.get("ticksuffix") is not None:
#                                 extra_base["value_axis_ticksuffix"] = val_ax.get("ticksuffix")

#                         # Category axis: order & crossing/categorical details
#                         if cat_ax:
#                             if cat_ax.get("categoryorder") is not None:
#                                 extra_base["category_order"] = cat_ax.get("categoryorder")
#                             if isinstance(cat_ax.get("categoryarray"), (list, tuple)):
#                                 extra_base["category_array"] = list(cat_ax.get("categoryarray"))
#                             if cat_ax.get("anchor") is not None:
#                                 extra_base["axis_anchor"] = cat_ax.get("anchor")  # e.g., 'y', 'x2'
#                             if cat_ax.get("side") is not None:
#                                 extra_base["axis_side"] = cat_ax.get("side")      # e.g., 'left', 'right', 'top', 'bottom'
#                             if cat_ax.get("overlaying") is not None:
#                                 extra_base["axis_overlaying"] = cat_ax.get("overlaying")
#                             if cat_ax.get("position") is not None:
#                                 extra_base["axis_position"] = cat_ax.get("position")
#                     except Exception:
#                         pass

#                     # Secondary X-axis assignment hint (useful for horizontal bars)
#                     if tr.get("xaxis") is not None:
#                         extra_base["xaxis_assign"] = tr.get("xaxis")  # 'x' | 'x2' | ...

#                     add_series(x, y, name=name, kind="bar", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides=extra_base)
#             elif t == "histogram":
#                 x_raw = x
#                 if not x_raw:
#                     continue
#                 x_arr = np.asarray(x_raw, dtype=float)
#                 nb = int(tr.get("nbinsx") or tr.get("nbins") or 30)

#                 xbins = tr.get("xbins") or {}
#                 if xbins:
#                     start = float(xbins.get("start")) if xbins.get("start") is not None else float(np.nanmin(x_arr))
#                     end   = float(xbins.get("end"))   if xbins.get("end")   is not None else float(np.nanmax(x_arr))
#                     size  = float(xbins.get("size"))  if xbins.get("size")  is not None else (end - start) / max(nb, 1)
#                     edges = np.arange(start, end + size, size)
#                 else:
#                     edges = np.histogram_bin_edges(x_arr, bins=nb)

#                 counts, edges = np.histogram(
#                     x_arr,
#                     bins=edges,
#                     density=(tr.get("histnorm") in ("probability", "probability density", "density")),
#                 )
#                 centers = (edges[:-1] + edges[1:]) / 2.0
#                 add_series(centers.tolist(), counts.tolist(), name=name, kind="hist", yaxis=yaxis,
#                            marker=marker, line=line, error_y=error_y,
#                            tr_level_overrides=extra_base)

#             elif t in ("scatter", "scattergl"):
#                 if not (x and y):
#                     continue
#                 mode = (tr.get("mode") or "").lower()
#                 if "lines+markers" in mode:
#                     add_series(x, y, name=name, kind="line", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides=extra_base)
#                     add_series(x, y, name=name, kind="scatter", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides={**extra_base, "overlay": "markers"})
#                 elif "lines" in mode:
#                     add_series(x, y, name=name, kind="line", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides=extra_base)
#                 else:
#                     add_series(x, y, name=name, kind="scatter", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides=extra_base)

#             else:
#                 if x and y:
#                     add_series(x, y, name=name, kind="line", yaxis=yaxis,
#                                marker=marker, line=line, error_y=error_y,
#                                tr_level_overrides=extra_base)

#         # ---- Title + axis titles + layout options ----
#         title = ""
#         lt = layout.get("title")
#         if isinstance(lt, dict):
#             title = lt.get("text", "") or ""
#         elif isinstance(lt, str):
#             title = lt

#         xaxis = layout.get("xaxis") or {}
#         x_title = _axis_title_from_layout(xaxis) or layout.get("xaxis_title") or ""

#         yaxis_layout = layout.get("yaxis") or {}
#         y_title = _axis_title_from_layout(yaxis_layout) or layout.get("yaxis_title") or ""

#         y2axis_layout = layout.get("yaxis2") or {}
#         y2_title = _axis_title_from_layout(y2axis_layout) or ""

#         legend = layout.get("legend") or {}
#         legend_orientation = legend.get("orientation") or layout.get("legend_orientation")
#         legend_y = legend.get("y") or layout.get("legend_y")
#         barmode = layout.get("barmode")

#         if not series:
#             return None

#         root_type = "mixed" if len(root_kinds) > 1 else next(iter(root_kinds))

#         spec: Dict[str, Any] = {
#             "type": root_type,
#             "categoryAxis": {
#                 "kind": "date" if _looks_like_dates(series) else "category",
#                 "title": x_title or "",
#             },
#             "valueAxis": {
#                 "kind": "linear",
#                 "title": y_title or "",
#             },
#             "series": series,
#             "layout": {
#                 "title": title,
#                 "legend": any(s.get("name") for s in series),
#             },
#         }

#         if any(s.get("axis") == "secondary" for s in series) or y2_title:
#             spec["valueAxis2"] = {"kind": "linear", "title": y2_title or ""}

#         if legend_orientation:
#             spec["layout"]["legend_orientation"] = str(legend_orientation)
#         if legend_y is not None:
#             spec["layout"]["legend_y"] = str(legend_y)
#         if barmode:
#             spec["layout"]["barmode"] = str(barmode)

#         return spec

#     except Exception:
#         return None





# def _try_chartspec_from_obj(obj: Any) -> Optional[Dict[str, Any]]:
#     # Prefer explicit adapters
#     if _is_plotly_fig(obj):
#         spec = _chartspec_from_plotly(obj)
#         if spec:
#             return spec
#     # Matplotlib figure or axes normalization
#     fig, _ = _mpl_as_figure(obj)
#     if fig is not None:
#         spec = _chartspec_from_matplotlib(fig)
#         if spec:
#             return spec
#     return None

# def _save_plot_image(obj: Any, dst_svg: str) -> str:
#     """Persist as SVG if possible, else PNG; return path actually written."""
#     # Plotly direct path
#     if hasattr(obj, "write_image"):
#         try:
#             obj.write_image(dst_svg, format="svg")
#             return dst_svg
#         except Exception:
#             dst_png = os.path.splitext(dst_svg)[0] + ".png"
#             obj.write_image(dst_png, format="png")
#             return dst_png

#     # Matplotlib normalization
#     fig, _ = _mpl_as_figure(obj)
#     if fig is not None:
#         try:
#             fig.savefig(dst_svg, format="svg")
#             return dst_svg
#         except Exception:
#             dst_png = os.path.splitext(dst_svg)[0] + ".png"
#             fig.savefig(dst_png, format="png", dpi=144)
#             return dst_png

#     # Last resort: if object exposes savefig directly
#     if hasattr(obj, "savefig"):
#         try:
#             obj.savefig(dst_svg, format="svg")
#             return dst_svg
#         except Exception:
#             dst_png = os.path.splitext(dst_svg)[0] + ".png"
#             obj.savefig(dst_png, format="png", dpi=144)
#             return dst_png

#     raise TypeError(f"Object of type {type(obj)} is not a supported plot.")


        # Chart objects
        # elif hasattr(val, "savefig"):  # matplotlib
        #     dst = os.path.join(paths["assets_dir"], f"{key}.svg")
        #     val.savefig(dst, format="svg")
        #     reg(dst, "chart", key)
         # Chart-like objects (matplotlib / plotly)
        # elif hasattr(val, "savefig") or _is_matplotlib_fig(val) or _is_plotly_fig(val):
        #     # 1) Prefer ChartSpec (plot2.0)
        #     spec = None
        #     if _is_plotly_fig(val):
        #         spec = _chartspec_from_plotly(val)
        #     if spec is None and _is_matplotlib_fig(val):
        #         spec = _chartspec_from_matplotlib(val)
        #     # Try to normalize matplotlib objects (Axes -> Figure) if needed
        #     if spec is None:
        #         fig = None
        #         try:
        #             # normalize to a Figure for matplotlib fallbacks
        #             fig, _axes = _mpl_as_figure(val)  # helper: returns (Figure or None, Axes or None)
        #         except Exception:
        #             fig = None
        #         if fig is not None:
        #             spec = _chartspec_from_matplotlib(fig)

        #     if spec:
        #         chart_path = os.path.join(paths["assets_dir"], f"{key}.xml")
        #         xmlp.write_chart_xml(chart_path, spec)
        #         size, sha = _file_info(chart_path)
        #         artifacts.append({
        #             "type": "plot2.0",
        #             "id": key,
        #             "path": chart_path,
        #             "href": _rel_href(chart_path, meta_dir),
        #             "mime": "application/xml",
        #             "bytes": size,
        #             "sha256": sha,
        #         })
        #     else:
        #         # 2) Fallback: persist as image correctly per backend
        #         dst_svg = os.path.join(paths["assets_dir"], f"{key}.svg")
        #         written = None
        #         if _is_plotly_fig(val) or hasattr(val, "write_image"):
        #             try:
        #                 val.write_image(dst_svg, format="svg")
        #                 written = dst_svg
        #             except Exception:
        #                 dst_png = os.path.splitext(dst_svg)[0] + ".png"
        #                 val.write_image(dst_png, format="png")
        #                 written = dst_png
        #         else:
        #             # matplotlib (normalize to Figure)
        #             fig, _axes = _mpl_as_figure(val)
        #             if fig is None and hasattr(val, "savefig"):
        #                 # last resort: object itself exposes savefig
        #                 fig = val
        #             if fig is None:
        #                 raise TypeError(f"Object of type {type(val)} is not a supported plot.")
        #             try:
        #                 fig.savefig(dst_svg, format="svg")
        #                 written = dst_svg
        #             except Exception:
        #                 dst_png = os.path.splitext(dst_svg)[0] + ".png"
        #                 fig.savefig(dst_png, format="png", dpi=144)
        #                 written = dst_png

        #         reg(written, "chart", key)


# def _chartspec_from_matplotlib(fig: Any) -> Optional[Dict[str, Any]]:
#     """
#     Robust Matplotlib → ChartSpec adapter.
#     Supports: line, scatter (PathCollection), bar (Rectangle/BarContainer), histogram (guessed from uniform bins).
#     Emits per-series style hints: {'kind': 'bar'|'hist'|'line'|'scatter'}.
#     Handles multiple Axes; root type becomes 'mixed' if heterogeneous.
#     """
#     try:
#         # Collect across all visible Axes
#         axes = [ax for ax in getattr(fig, "get_axes", lambda: [])() if ax.get_visible()]
#         if not axes:
#             return None

#         series: List[Dict[str, Any]] = []
#         root_kinds: set[str] = set()

#         def add_series(x, y, *, name: str = "", kind: str, extra_style: Optional[Dict[str, Any]] = None):
#             st = {"kind": kind}
#             if isinstance(extra_style, dict):
#                 st.update({k: str(v) for k, v in extra_style.items()})
#             sx = [_to_serializable(xx) for xx in (x if isinstance(x, (list, tuple, np.ndarray)) else [x])]
#             sy = [_to_serializable(yy) for yy in (y if isinstance(y, (list, tuple, np.ndarray)) else [y])]
#             series.append({"name": name or "", "x": sx, "y": sy, "style": st})
#             root_kinds.add(kind)

#         def _is_hist_like(rects: List[Any]) -> bool:
#             # Heuristic: many vertical rectangles of ~uniform width, contiguous or near-contiguous.
#             if len(rects) < 5:
#                 return False
#             xs = []
#             ws = []
#             for r in rects:
#                 try:
#                     x = float(getattr(r, "get_x")())
#                     w = float(getattr(r, "get_width")())
#                     ys = float(getattr(r, "get_y")())
#                 except Exception:
#                     continue
#                 # Histogram bars are usually vertical (y starts at 0 or baseline)
#                 if w <= 0:
#                     continue
#                 xs.append(x)
#                 ws.append(abs(w))
#             if len(xs) < 5:
#                 return False
#             ws = np.asarray(ws, dtype=float)
#             if not np.all(np.isfinite(ws)):
#                 return False
#             # Uniformity: coefficient of variation small
#             if (ws.std() / (ws.mean() + 1e-9)) > 0.15:
#                 return False
#             # Sorted x with near step = width
#             xs = np.sort(np.asarray(xs, dtype=float))
#             steps = np.diff(xs)
#             if len(steps) == 0 or not np.all(np.isfinite(steps)):
#                 return False
#             step_med = np.median(steps)
#             if step_med <= 0:
#                 return False
#             return np.median(np.abs(steps - step_med)) / (step_med + 1e-9) < 0.25

#         for ax in axes:
#             # ----- Line2D (classic plot) -----
#             for line in ax.get_lines():
#                 try:
#                     xdata = list(map(_to_serializable, line.get_xdata(orig=False)))
#                     ydata = list(map(_to_serializable, line.get_ydata(orig=False)))
#                 except Exception:
#                     continue
#                 if not (xdata and ydata):
#                     continue
#                 name = line.get_label() or ""
#                 if name == "_nolegend_":
#                     name = ""
#                 # Determine style hint: line vs scatter vs both
#                 ls = (line.get_linestyle() or "").lower()
#                 mk = (line.get_marker() or "").lower()
#                 if ls not in ("", "none", " ", "none"):  # has a line
#                     add_series(xdata, ydata, name=name, kind="line")
#                     if mk and mk not in ("", "none", " "):
#                         # overlay markers as a separate scatter layer (like Plotly adapter)
#                         add_series(xdata, ydata, name=name, kind="scatter", extra_style={"overlay": "markers"})
#                 else:
#                     # marker-only plot()
#                     add_series(xdata, ydata, name=name, kind="scatter")

#             # ----- True scatter (PathCollection) -----
#             # Matplotlib: ax.collections may include many collection types; guard on get_offsets.
#             for coll in getattr(ax, "collections", []):
#                 get_offsets = getattr(coll, "get_offsets", None)
#                 if not callable(get_offsets):
#                     continue
#                 try:
#                     pts = get_offsets()
#                 except Exception:
#                     continue
#                 if pts is None:
#                     continue
#                 try:
#                     arr = np.asarray(pts, dtype=float)
#                 except Exception:
#                     continue
#                 if arr.ndim != 2 or arr.shape[1] < 2 or arr.size == 0:
#                     continue
#                 x = [_to_serializable(float(v)) for v in arr[:, 0]]
#                 y = [_to_serializable(float(v)) for v in arr[:, 1]]
#                 if x and y:
#                     # Matplotlib PathCollections rarely have a useful label
#                     name = getattr(coll, "get_label", lambda: "")() or ""
#                     if name == "_nolegend_":
#                         name = ""
#                     add_series(x, y, name=name, kind="scatter")

#             # ----- Bars / Histograms (Rectangles) -----
#             # Prefer BarContainers if present; fallback to ax.patches grouping.
#             bar_rects: List[Any] = []
#             try:
#                 for cont in getattr(ax, "containers", []):
#                     # BarContainer exists in mpl; try to detect via attribute shape
#                     if hasattr(cont, "patches") and cont.patches:
#                         bar_rects.extend([p for p in cont.patches if hasattr(p, "get_width") and hasattr(p, "get_height")])
#             except Exception:
#                 pass
#             # Fallback: all Rectangle patches on the axes that look like vertical bars.
#             if not bar_rects:
#                 for p in getattr(ax, "patches", []):
#                     if hasattr(p, "get_width") and hasattr(p, "get_height"):
#                         bar_rects.append(p)

#             # Group rectangles by baseline and orientation; here we keep one simple group per Axes.
#             rects = [r for r in bar_rects if hasattr(r, "get_x") and hasattr(r, "get_y")]
#             if rects:
#                 # Build bar series: x = centers, y = heights (positive or negative)
#                 xs = []
#                 ys = []
#                 for r in rects:
#                     try:
#                         x = float(r.get_x())
#                         y = float(r.get_y())
#                         w = float(r.get_width())
#                         h = float(r.get_height())
#                     except Exception:
#                         continue
#                     # Vertical bars: width > height is not always true; use Rectangle orientation:
#                     # In bar() and hist(), width is along x; use center on x
#                     cx = x + w / 2.0
#                     val = h if y <= 0 else (y + h) - y  # standard height; works for baseline 0
#                     xs.append(_to_serializable(cx))
#                     ys.append(_to_serializable(val))
#                 if xs and ys:
#                     kind = "hist" if _is_hist_like(rects) else "bar"
#                     # Attempt a legend label from the first container/patch that has it
#                     label = ""
#                     try:
#                         for obj in (getattr(ax, "containers", []) or []) + rects:
#                             get_label = getattr(obj, "get_label", None)
#                             if callable(get_label):
#                                 lab = get_label()
#                                 if lab and lab != "_nolegend_":
#                                     label = lab
#                                     break
#                     except Exception:
#                         pass
#                     add_series(xs, ys, name=label, kind=kind)

#         if not series:
#             return None

#         # Title extraction: prefer suptitle if present, else join unique axis titles if multiple
#         title = ""
#         supt = getattr(fig, "_suptitle", None)
#         if hasattr(supt, "get_text"):
#             title = supt.get_text() or ""
#         if not title:
#             titles = []
#             for ax in axes:
#                 t = ax.get_title()
#                 if hasattr(t, "get_text"):
#                     t = t.get_text()
#                 if t:
#                     titles.append(str(t))
#             titles = list(dict.fromkeys(titles))  # unique order-preserving
#             title = " · ".join(titles) if titles else ""

#         root_type = "mixed" if len(root_kinds) > 1 else next(iter(root_kinds))
#         spec = {
#             "type": root_type,
#             "categoryAxis": {"kind": "date" if _looks_like_dates(series) else "category"},
#             "valueAxis": {"kind": "linear"},
#             "series": series,
#             "layout": {
#                 "title": title,
#                 "legend": any(s.get("name") for s in series)
#             },
#         }
#         return spec

#     except Exception:
#         return None


# def _to_serializable(v: Any) -> Any:
#     # Make matplotlib/plotly numpy types, pandas timestamps, etc., JSON/XML friendly.
#     try:
#         import numpy as np  # optional
#         if isinstance(v, (np.integer,)):
#             return int(v)
#         if isinstance(v, (np.floating,)):
#             return float(v)
#         if isinstance(v, (np.bool_,)):
#             return bool(v)
#         if isinstance(v, (np.datetime64,)):
#             # Convert to ISO 8601 string
#             return str(np.datetime_as_string(v, unit="s"))
#     except Exception:
#         pass
#     try:
#         import pandas as pd  # optional
#         if isinstance(v, (pd.Timestamp,)):
#             return v.to_pydatetime().isoformat()
#     except Exception:
#         pass
#     return v

# def _looks_like_dates(series: List[Dict[str, Any]]) -> bool:
#     count, hits = 0, 0
#     for s in series:
#         for xv in s.get("x", []):
#             count += 1
#             if count > 50:  # sample cap
#                 break
#             if isinstance(xv, str) and _is_iso8601_like(xv):
#                 hits += 1
#         if count > 50:
#             break
#     return hits >= max(3, count // 4)


