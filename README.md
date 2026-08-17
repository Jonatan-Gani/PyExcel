# PyExcel

**Status:** PyExcel is being rebuilt as a .NET / Excel-DNA add-in (v2). The
original VBA `.xlam` add-in (v1) has been **removed from the repository** — what
remains is the v2 codebase, which is in **active development** and not yet a
shippable build. See [`ARCHITECTURE.md`](ARCHITECTURE.md) for the codebase map
and [`ROADMAP.md`](ROADMAP.md) for live status and the path to a production build.

**Version:** `2.0.0-alpha`

A Windows Excel add-in that runs user-authored Python `transform()` functions
against workbook ranges in an isolated, project-local Python environment. The
user selects input ranges, picks a script, and clicks **Run**; the add-in
marshals the ranges to a supervised Python kernel, executes the script, and
writes the returned tables, lists, scalars, and charts back into Excel. Runs are
non-reactive: nothing executes until **Run** is clicked.

## Building

You cannot produce the final `.xll` on Linux — the add-in projects target
`net48`. The cross-platform projects and the Python kernel build and test on
Linux/macOS. Full prerequisites, build commands, and the current exit gate are
in [`docs/v2-build.md`](docs/v2-build.md).

```bash
# Cross-platform slice (what CI runs on Linux):
dotnet build src/PyExcel.Common/PyExcel.Common.csproj --framework netstandard2.0 -c Release
dotnet build src/PyExcel.Bridge/PyExcel.Bridge.csproj  --framework netstandard2.0 -c Release
dotnet test  tests/PyExcel.Bridge.Tests/PyExcel.Bridge.Tests.csproj -c Release
pytest tests/
```

```powershell
# Windows — produces src/PyExcel.Addin/bin/Release/net48/PyExcel-AddIn64.xll:
dotnet restore PyExcel.sln
dotnet build PyExcel.sln -c Release --no-restore
```

## Authoring a script

A script defines a `transform(inputs)` function that receives the mapped input
ranges and returns a dict of named results:

```python
from typing import Dict, Any
import pandas as pd

def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:
    sales = inputs["Sales"]            # pandas.DataFrame
    tax   = inputs.get("TaxRate", 0.0) # float

    sales["Total"] = sales["Quantity"] * sales["Price"]

    return {
        "ProcessedSales": sales,                            # → Excel table
        "TotalRevenue":   float(sales["Total"].sum()),      # → single cell
        "Clients":        sales["Client"].unique().tolist() # → spill range
    }
```

> **No interactive input.** Scripts run in a headless kernel with no console
> attached, so `input()` and any read from `sys.stdin` are **disabled** — they
> raise an error immediately rather than hanging the run. Read every value your
> script needs from the `inputs` argument (and pass constants via the action's
> **Keyword args**). `print()` output is captured and shown in the log window.

### Inputs

The **Input** ribbon field is a semicolon-separated list of Excel ranges, each
optionally name-bound with `name=Range`. The name is the key the range appears
under in `inputs`. Unnamed ranges are auto-named by their resolved type —
`df1`, `df2`, …, `list1`, …, `value1`, …

```
Sales=Sheet1!A1:C10; TaxRate=Sheet1!E1; Months=Sheet1!A1:A12
```

Each binding also carries a **Type**, chosen in the action dialog. Left on
`Auto` it resolves from the range's dimensions:

| Range shape              | `Auto` resolves to | Python value in `inputs`                       |
| ------------------------ | ------------------ | ---------------------------------------------- |
| Multi-row × multi-column | DataFrame          | `pandas.DataFrame`; **the first row is the header** |
| Single row or column     | List               | `list[...]`                                    |
| Single cell              | Scalar             | `int` / `float` / `bool` / `str` / `Timestamp` |

Override it to get something else: `DataFrame`, `Series`, `List`, `Tuple`,
`Set`, `Dict`, `NDArray` or `Scalar`. A declared type is also written into the
binding text as `name:type=Range`, so `Prices:ndarray=Sheet1!A1:C10` and the
dialog's Type box are the same setting.

`DataFrame` and `Series` take their names from the first row; `List`, `Tuple`,
`Set`, `NDArray` and `Scalar` treat every cell as data. A `Dict` reads two
columns as key → value pairs, and three or more as column-oriented lists keyed
by the header row.

A range that cannot satisfy its declared type fails the run with a message
naming the binding, the range and its dimensions — rather than surfacing as an
`AttributeError` inside your script:

```
input 'Sales' (Sheet1!A1:B3, 3x2): declared type Series needs a single row or
column, but the range is 3x2. Use DataFrame, or narrow the selection to one
row or column.
```

The full matrix of every declared type against every range shape is in
[`docs/typed-io-contract.md`](docs/typed-io-contract.md).

### Outputs

Return a `dict[str, Any]`. Each key becomes a named result; the value's Python
type decides how it is written. A non-dict return is wrapped automatically.

| Return value                     | Excel rendering                          |
| -------------------------------- | ---------------------------------------- |
| `pandas.DataFrame` / `Series`    | Formatted table at the destination       |
| `list[scalar]` / `tuple[scalar]` | Spill range (vertical or horizontal)     |
| `int` / `float` / `str`          | Single cell                              |
| `plotly.graph_objects.Figure`    | Native Excel chart                       |
| `matplotlib.Figure` / `Axes`     | Embedded image (SVG, PNG fallback)       |

By default outputs spill from the cell in the **Output** field. To route
specific keys to specific ranges, use the same `name=Range` syntax in
**Output** — each returned key is written to the binding of the same name:

```
ProcessedSales=Sheet2!A1; TotalRevenue=Summary!B2; Clients=Lists!A1
```

An Output binding may declare a type too, which the kernel then *enforces*
rather than constructs: if the returned value for that key isn't that type,
the run fails with `output 'ProcessedSales': declared type DataFrame, but
transform() returned list.` Left on `Auto`, nothing is enforced and whatever
comes back is rendered by its shape — the loose behaviour that keeps
iterating on a script comfortable.

If your script returns a key the Output field doesn't mention, the run
reports it instead of discarding the result.

## Ribbon

A single **Python** tab groups the controls:

| Group  | Controls                                                                                |
| ------ | --------------------------------------------------------------------------------------- |
| Main   | Enable (install + enable) · Update · Open In Explorer · Read Me                          |
| Errors | Show Last Error · Copy Last Error                                                        |
| Python | Run · Edit · Script · Input · Output · Actions · Add / Edit / Delete Action              |
| Import | Import · Source · Destination · Edit                                                     |
| Export | Export · Source · Saves to · Edit                                                       |
| Paste  | Paste · Destination · Edit                                                               |

**Run** executes the selected script against the **Input** ranges, writing
results from **Output**. **Actions** save a configured (script + input + output)
as a reusable per-workbook preset.

Each action has a **Keep output window open after a successful run** checkbox
(on by default). When on, a successful run leaves the script's captured
`print()` output on screen so you can read it; turn it off to dismiss the
output once the run succeeds. A *failed* run always keeps its error window open
regardless of this setting.

## Repository layout

See [`ARCHITECTURE.md`](ARCHITECTURE.md) for the full codebase map. In brief:

```
src/PyExcel.*/        v2 .NET projects (C#) — see PyExcel.sln
embedded/pyexcel/     Python kernel package (shipped inside the .xll)
tests/                pytest (kernel) + xUnit (PyExcel.Bridge.Tests)
requirements.txt      Python kernel dependencies installed into the project venv
docs/                 Build and design docs
```

## Logs

When the `.xll` is loaded it writes to `%TEMP%\PyExcel_Debug.log`, format
`[YYYY-MM-DD hh:mm:ss.fff] [LEVEL] message`.

## Limitations

- Windows-only for the Excel add-in (COM + Excel-DNA).
- Requires a system Python available at setup time — the setup wizard
  provisions a project-local virtual environment from it; no Python is bundled.
- Scripts run without an interactive console: `input()` and reads from
  `sys.stdin` are disabled and raise an error. Pass data through the input
  ranges and keyword args instead.

## Contact

- Email: JonatanGani@protonmail.com
- Repository: https://github.com/Jonatan-Gani/PyExcel
