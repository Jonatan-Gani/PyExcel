# PyExcel

**Version:** 20260422_212123

An Excel add-in (`.xlam`) that runs Python scripts against workbook data in an isolated virtual environment. Execution is non-reactive: scripts run on demand, return their results to Excel, and exit. Inputs are mapped from Excel ranges into Pandas/Python objects; outputs are written back as tables, lists, scalars, charts, or images.

## Features

- Run Python scripts from Excel with inputs picked from named ranges or address strings.
- Automatic conversion of Excel ranges to `pandas.DataFrame`, `list`, or scalar values.
- Outputs rendered back as Excel tables, spill ranges, single cells, charts, or images.
- Inject live Excel formulas into output DataFrames via `excel_formula`.
- Self-contained `venv` provisioned by a setup wizard on first run.

## Requirements

- Microsoft Excel for Windows (the add-in is `.xlam` and uses VBA + `pywin32`).
- Python 3.7+ available on the system at install time.

## Installation

1. **Allow VBA project access.** Excel → *File > Options > Trust Center > Trust Center Settings > Macro Settings* → enable **Trust access to the VBA project object model**.
2. **Unblock the file.** Right-click `PyExcel.xlam` → *Properties* → check **Unblock** → Apply.
3. **Load the add-in.** *File > Options > Add-ins*, manage **Excel Add-ins**, click **Go**, **Browse** to `PyExcel.xlam`, and tick it.
4. **Run the setup wizard.** Click **Enable** on the PyExcel ribbon. The wizard prompts for a project folder, creates the working directories (including `userScripts/`), provisions `.venv`, and installs the dependencies listed in `requirements.txt`.

## Updating

Replace `PyExcel.xlam` with the newer build. On next launch the add-in detects the version change, cleans stale resources, re-extracts embedded files, and syncs Python dependencies against the bundled `requirements.txt`.

To add a library yourself, install it into the project venv:

```bash
.venv\Scripts\python.exe -m pip install <package>
```

## Usage

### Inputs

Pass Excel ranges as a semicolon-separated list. Use `{name}=Range` to bind a variable name; otherwise inputs are auto-named (`df1`, `df2`, …, `list1`, …, `value1`, …).

```
{Sales}=Sheet1!A1:C10; {TaxRate}=Sheet1!E1; {Months}=Sheet1!A1:A12
```

| Range shape              | Python type                                  |
| ------------------------ | -------------------------------------------- |
| Multi-row × multi-column | `pandas.DataFrame` (column types inferred)   |
| Single row or column     | `list[str]`                                  |
| Single cell              | `int` / `float` / `bool` / `str` / `Timestamp` |

### Writing a script

Scripts live under `userScripts/` (created by the wizard). Each script defines a `transform(inputs)` function and ends with a CLI entrypoint:

```python
from typing import Dict, Any
import pandas as pd
from tools import run_script_cli, excel_formula

def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:
    sales = inputs["Sales"]           # DataFrame
    tax   = inputs.get("TaxRate", 0)  # scalar

    sales["Total"]      = sales["Quantity"] * sales["Price"]
    sales["FinalPrice"] = [excel_formula(f"=F{i+2}*(1+{tax})") for i in range(len(sales))]

    return {
        "ProcessedSales": sales,                           # → Excel table
        "TotalRevenue":   sales["Total"].sum(),            # → single cell
        "Clients":        sales["Client"].unique().tolist() # → spill range
    }

if __name__ == "__main__":
    run_script_cli(transform)
```

### Outputs

| Return value                          | Excel rendering                              |
| ------------------------------------- | -------------------------------------------- |
| `pandas.DataFrame`                    | Formatted table                              |
| `list` / `tuple`                      | Spill range (lists of DataFrames are stacked)|
| `int` / `float` / `str` / `bool`      | Single cell                                  |
| `matplotlib.Figure` / `plotly` figure | Native Excel chart where possible, else image |
| `str` path to `.png` / `.svg` / `.emf`| Embedded picture                             |

A non-dict return (single value, list, or tuple) is wrapped automatically into keys `result_0`, `result_1`, …

## Project structure

```
PyExcel.xlam           Compiled Excel add-in (load this in Excel)
requirements.txt       Python dependencies installed into the project venv
src/
  module/              VBA source (.bas, .cls, .frm) for version control
  embedded/            Python runtime: tools.py, xmlParsing.py, instructions
  Ribbon/              Ribbon XML and logo
```

`userScripts/`, `.venv/`, and other working directories are created at install time inside the folder you pick in the setup wizard — they are not part of this repository.

## Support

- Email: JonatanGani@protonmail.com
- Repository: https://github.com/Jonatan-Gani/PyExcel
