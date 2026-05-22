# PyExcel

**Status:** v1 (`PyExcel.xlam`, VBA) is the current shipping add-in. v2 (`.xll`, .NET) is in active development.
**Version:** v1 build `20260422_212123` · v2 target `2.0.0-alpha`

> **Contributors:** [`ARCHITECTURE.md`](ARCHITECTURE.md) maps the whole codebase; [`ROADMAP.md`](ROADMAP.md) is the task list and path to a production-grade v2.

A Windows Excel add-in that runs Python scripts against workbook ranges in an isolated project environment. Each run serializes selected ranges to typed XML, spawns the Python interpreter against a user-authored `transform(inputs)` function, polls a `meta.xml` status file for completion, then writes the returned tables, lists, scalars, and figures back to Excel. Runs are non-reactive and self-contained: nothing executes until the user clicks **Run**, and every run is archived to disk.

## Requirements

- Excel for Windows, 2007 or later (the add-in is `.xlam` and uses RibbonX + VBA).
- Python 3.7 or later available on `PATH` at install time — the wizard provisions a project-local `.venv` from this interpreter; no Python is bundled.
- `Trust access to the VBA project object model` enabled in Excel's Trust Center.

## Installation

1. **Unblock the file.** Right‑click `PyExcel.xlam` → *Properties* → check **Unblock** → Apply.
2. **Enable VBA project access.** Excel → *File > Options > Trust Center > Trust Center Settings > Macro Settings* → tick **Trust access to the VBA project object model**.
3. **Load the add-in.** *File > Options > Add-ins*, manage **Excel Add-ins**, **Go…**, **Browse** to `PyExcel.xlam`, tick it.
4. **Run the setup wizard.** On the **Python** ribbon tab, click **Enable PyExcel**. The wizard:
   - prompts for a project name and root folder,
   - converts the host workbook to `.xlsm` (if needed) and pins the project path to it,
   - creates the project tree (`Python/.venv`, `Python/Scripts`, `userScripts`, `Temp/{tables,lists,values,assets}`, `Archive`, `AddIn`),
   - runs `python -m venv` against the system Python to create `.venv`,
   - extracts the embedded runtime (`tools.py`, `xmlParsing.py`, `requirements.txt`, `instructions.txt`) from the add-in's `EmbeddedStore` sheet,
   - upgrades `pip` and installs `Python/requirements.txt` into the venv,
   - verifies that ≥80% of required packages are present and writes a setup log to `Temp/setup_log.txt`.

## Updating

Replace `PyExcel.xlam` with a newer build. On the next workbook activation, the add-in compares the bundled version against the workbook's stored `PyExcel_ProjectVersion` name. If newer, you'll be prompted to update; this performs:

- **SmartClean** — deletes files under add-in-owned folders that aren't listed in the new manifest, while never touching `.venv` or `userScripts`.
- **Re-extract** — writes the new embedded resources to disk.
- **Dependency sync** — runs `pip uninstall -r Uninstall.txt -y` (if non-empty), then `pip install -r requirements.txt --upgrade`, then snapshots the result to `Python/User_Environment_Snapshot.txt`.

To install your own packages into the project venv:

```bash
Python\.venv\Scripts\python.exe -m pip install <package>
```

## Ribbon

A single **Python** tab with five groups:

| Group  | Controls                                                                                          |
| ------ | ------------------------------------------------------------------------------------------------- |
| Main   | Enable PyExcel · Open In Explorer · Read Me                                                       |
| Python | Run · Edit · Script (dropdown of `userScripts/*.py`) · Input · Output · Actions · Add · Edit · Delete Action |
| Import | Import · Source · Destination · Edit                                                              |
| Export | Export · Source · Destination · Edit                                                              |
| Paste  | Paste · Destination · Edit                                                                        |

**Run** executes the script selected in the **Script** combo against the ranges in **Input**, writing results starting at **Output**. **Actions** lets you save a configured (script + input + output) as a reusable preset stored on the workbook.

## Authoring a script

Drop a `.py` file into the project's `userScripts/` folder; it shows up in the **Script** dropdown. Every script defines a `transform(inputs)` function and ends with the CLI shim:

```python
from typing import Dict, Any
import pandas as pd
from tools import run_script_cli, excel_formula

def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:
    sales = inputs["Sales"]            # pandas.DataFrame
    tax   = inputs.get("TaxRate", 0.0) # float

    sales["Total"]      = sales["Quantity"] * sales["Price"]
    sales["FinalPrice"] = [
        excel_formula(f"=F{i+2}*(1+{tax})") for i in range(len(sales))
    ]

    return {
        "ProcessedSales": sales,                            # → Excel table
        "TotalRevenue":   float(sales["Total"].sum()),      # → single cell
        "Clients":        sales["Client"].unique().tolist() # → spill range
    }

if __name__ == "__main__":
    run_script_cli(transform)
```

`run_script_cli` parses `--in`, `--out`, `--meta`, and `--run-id` (all passed by the add-in), runs your transform, writes artifacts, and updates `meta.xml` with `status = done | error | in_progress`. A background heartbeat thread updates the meta timestamp every 10 s so the add-in can distinguish a slow run from a hung one.

### Inputs

The **Input** ribbon field is a semicolon-separated list of Excel ranges, each optionally name‑bound with `{name}=Range`. Unnamed ranges are auto-named `df1`, `df2`, …, `list1`, …, `value1`, …

```
{Sales}=Sheet1!A1:C10; {TaxRate}=Sheet1!E1; {Months}=Sheet1!A1:A12
```

The add-in serializes each range to typed XML based on its shape, and your transform receives a `dict[str, Any]`:

| Range shape              | XML element | Python value in `inputs`                       |
| ------------------------ | ----------- | ---------------------------------------------- |
| Multi-row × multi-column | `<table>`   | `pandas.DataFrame` with inferred column types  |
| Single row or column     | `<list>`    | `list[...]` with a single inferred datatype    |
| Single cell              | `<value>`   | `int` / `float` / `bool` / `str` / `Timestamp` |

Column types and datatypes are detected per range: `int`, `float` (`decimal`), `bool`, `timestamp` (ISO 8601), or `string`; blank cells become `pd.NA`.

### Outputs

Return a `dict[str, Any]`. Each key becomes an artifact id; the value's Python type decides how it's written. A non-dict return (single value, list, or tuple) is wrapped automatically into `result_0`, `result_1`, ….

| Return value                          | Artifact type | Excel rendering                                          |
| ------------------------------------- | ------------- | -------------------------------------------------------- |
| `pandas.DataFrame` / `Series`         | `table`       | Formatted table at the destination                       |
| `dict[str, DataFrame]`                | `table`       | Multiple named tables in one artifact                    |
| `list[DataFrame]` / `tuple[...]`      | `table`       | Stacked tables                                           |
| `list[scalar]` / `tuple[scalar]`      | `list`        | Spill range (vertical or horizontal)                     |
| `int` / `float` / `str`               | `value`       | Single cell                                              |
| `plotly.graph_objects.Figure`         | `plot2.0`     | Native Excel chart (via `PlotlyToExcelXMLConverter`)     |
| `matplotlib.Figure` / `Axes`          | `chart`       | Embedded SVG (PNG fallback at 144 dpi)                   |
| `str` path to `.emf` / `.svg` / `.png` / `.xml` | `chart` / `plot2.0` | Embedded picture, or native chart for `.xml`   |

By default outputs spill from the cell in the **Output** field. To route specific keys to specific ranges, use the same `{name}=Range` syntax in **Output**:

```
{ProcessedSales}=Sheet2!A1; {TotalRevenue}=Summary!B2; {Clients}=Lists!A1
```

### Excel formulas

`excel_formula("=A2*1.1")` returns an `ExcelFormula` dataclass (A1 mode). When the add-in pastes the surrounding DataFrame, it writes the formula text into the cell instead of the literal string, so the formula is live in Excel.

## How a run is executed

1. The **Input** field is parsed into named ranges (`xmlParsing.bas → SerializeRangeToTypedXML`) and written to `Temp/in_<script>_<runid>.xml`.
2. `python.bas → RunPythonJob` shells out:
   ```
   cmd /c "set PYTHONPATH=Python\Scripts;...
           Python\.venv\Scripts\python.exe -u userScripts\<script>.py
               --in <in.xml> --out <out.xml> --meta <meta.xml> --run-id <id>
           2>&1 | powershell Tee-Object <log>"
   ```
3. VBA polls `meta.xml` (60 s for first status, 120 s reactivity window, 300 s absolute), watching the heartbeat timestamp to detect stalls.
4. On `status = done`, `pythonUtils.bas → PasteArtifactsToTargets` reads the manifest and writes each artifact to its mapped destination.
5. The full run — inputs, outputs, meta, log — is copied to `Archive/<timestamp>_<script>/`; the archive keeps the last 10 runs.

## Other ribbon actions

- **Import** loads a CSV/TSV (`FastCSVParse`) or any Excel-compatible workbook (XLSX/XLSM/XLSB/ODS via COM) into the destination range, with sheet‑picker support for multi-sheet sources.
- **Export** writes a source range to CSV or an Excel format.
- **Paste** writes a previously produced artifact (or any compatible XML/image in the project) to a destination range without re-running Python.
- **Actions** are saved (script, input, output) triples persisted on the workbook so a run can be replayed with one click.

## Repository layout

> This shows the **v1** source tree. For the full repository map — including the
> v2 .NET projects and the v2 Python kernel — see [`ARCHITECTURE.md`](ARCHITECTURE.md).

```
PyExcel.xlam            Compiled Excel add-in — load this in Excel
requirements.txt        Python dependencies installed into the project venv
src/
  module/               VBA source kept under version control
    Setup.bas             First-run wizard (folders, venv, extract, pip)
    Update.bas            Version check, SmartClean, dependency sync
    HostManager.bas       Workbook registry, watchdog, ribbon state
    CAppEvents.cls        Application event sink (Open/Activate/SheetActivate)
    python.bas            Run orchestrator (serialize → spawn → poll → paste)
    pythonUtils.bas       Meta/artifact parsing, paste, archive
    xmlParsing.bas        Range → typed XML serialization
    Import.bas / Export.bas / Paste.bas
    chartBuilder.bas      Plotly/Matplotlib artifact → Excel chart/picture
    PathUtils.bas / modDst.bas / modRibbon.bas
    frm*.frm              Dialogs (range picker, edit import/export/paste,
                          export wizard, orientation, progress)
  embedded/             Runtime resources extracted into the project on setup
    tools.py              run_script / run_script_cli, artifact materializer
    xmlParsing.py         read_xml / write_xml, ExcelFormula,
                          PlotlyToExcelXMLConverter
    requirements.txt      Pinned dependency set
    instructions.txt      Pip cheat-sheet shown to users
    Uninstall.txt         Packages to remove on update (optional, empty by default)
  Ribbon/
    RibbonUI.xml          Custom ribbon definition
    customLogo.png        Run-button icon
```

## Runtime project layout

After the wizard runs, the project folder you picked looks like:

```
<project>/
  AddIn/                 Local copy of the add-in and resources
  Python/
    .venv/               Project virtualenv (never touched by updates)
    Scripts/             tools.py, xmlParsing.py (extracted)
    requirements.txt
    User_Environment_Snapshot.txt   (pip freeze after last sync)
  userScripts/           Your transform scripts — picked up by the Run dropdown
  Temp/
    tables/  lists/  values/  assets/    Per-artifact output XML/images
    in_*.xml  out_*.xml  meta_*.xml      Per-run IPC files
    setup_log.txt  pip_install.log
  Archive/
    <yyyymmdd_hhnnss>_<script>/          Inputs, outputs, meta, log (last 10)
```

## Debugging

- `Temp/PyExcel_Debug.log` — one line per VBA module call.
- `Temp/setup_log.txt`, `Temp/pip_install.log` — wizard / pip output.
- `Temp/meta_<script>_<runid>.xml` — last run status, including `stderr` on failure.
- `Archive/<run>/` — full reproducible bundle of any past run.

## Limitations

- Windows‑only (`pywin32`, `WScript.Shell`, `ADODB.Stream`, `MSXML2.DOMDocument`).
- Requires a system Python on `PATH` at install time — none is bundled.
- Whole DataFrames are materialized to XML; very large results (>10⁶ cells) will be slow.
- Formulas are A1 mode only; R1C1 is stubbed but not implemented.

## Contact

- Email: JonatanGani@protonmail.com
- Repository: https://github.com/Jonatan-Gani/PyExcel
