**Version:** 20251222_155035

---

# PyExcel

PyExcel is a powerful Excel Add-in that bridges the gap between Microsoft Excel and Python. It enables users to leverage Python's robust data analysis and visualization libraries (like Pandas, NumPy, and Plotly) directly within their Excel workflows.

## Features

*   **Seamless Integration**: Call Python scripts directly from Excel.
*   **Flexible Inputs**: Automatically detects and converts Excel ranges into Pandas DataFrames, Lists, or Scalar values. Supports named variables for inputs.
*   **Rich Output**: Return modified tables, lists, summary values, or complex objects like Plotly/Matplotlib charts back to Excel.
*   **Excel Formulas**: Inject dynamic Excel formulas into your output dataframes.

## Prerequisites & Setup

### 1. Python Environment
Ensure you have **Python 3.7+** installed. The project uses a virtual environment located in `Python/.venv`.

To install dependencies:
```bash
# From the project root
Python/.venv/Scripts/python.exe -m pip install -r Python/requirements.txt
```

### 2. Excel Configuration
To ensure the Add-in functions correctly, you must configure Excel security settings:

1.  **Enable Trust Access**:
    *   Navigate to: File > Options > Trust Center > Trust Center Settings > Macro Settings.
    *   Check: **Trust access to the VBA project object model**.
2.  **Unblock the Add-in**:
    *   Navigate to the `AddIn/` folder.
    *   Right-click `PyExcel.xlam` (or `PyExcel_Dev.xlam`) and select **Properties**.
    *   Under the "Security" section, check **Unblock** and click **Apply**.

*(Optional for Developers)*: If you are editing the VBA code, ensure the **Microsoft Visual Basic for Applications Extensibility 5.3** reference is enabled in the VBA Editor (Tools > References).

## Update Process

The add-in includes a built-in update manager.

1.  **Automatic Check**: When you open your workbook, the add-in compares its version with the installed version. If a newer version is available, you will be prompted to update.
2.  **Manual Update**: You can manually trigger an update by running the `RunUpdateFromExternalFile` macro.
    *   This will ask you to select the new `.xlam` file.
    *   The system will automatically clean old files, extract new resources, and update Python dependencies.

## Usage

### 1. Defining Inputs in Excel
The system parses input ranges provided from Excel. You can select ranges or pass a semicolon-separated string of addresses.

*   **Variable Naming**: Use `{name}=Range` syntax to assign specific names to your inputs.
    *   Example: `{Sales}=Sheet1!A1:C10; {TaxRate}=Sheet1!E1; {Months}=Sheet1!A1:A12`

#### 📥 Input Type Mapping

| Excel Range Shape | Detected Type | Python Type | Description |
| :--- | :--- | :--- | :--- |
| **Multi-row & Multi-column** | Table | `pandas.DataFrame` | Converted to a DataFrame. Column types (int, float, date) are inferred automatically. |
| **Single Row or Single Column** | List | `list[str]` | Converted to a list of strings. |
| **Single Cell** | Scalar | `int`, `float`, `bool`, `str`, `pd.Timestamp` | Converted to a native Python scalar value. |

> **Note**: If no name is provided, inputs are named automatically: `df1`, `df2`... for tables; `list1`... for lists; `value1`... for scalars.

---

### 2. Writing Python Scripts
Create your Python scripts in the `userScripts/` directory. A script must define a `transform` function that acts as the entry point.

**Function Signature:**
```python
from typing import Dict, Any, Union, List
import pandas as pd

def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:
    # inputs["Sales"] -> pd.DataFrame
    # inputs["TaxRate"] -> float
    ...
    return { "ResultTable": df_result }
```

#### 📤 Output Type Mapping

| Python Return Object | Excel Output | Description |
| :--- | :--- | :--- |
| **`pandas.DataFrame`** | **Excel Table** | Renders as a formatted table. |
| **`list`** or **`tuple`** | **Dynamic List** | Renders as a spill range (vertical or horizontal). Lists of DataFrames are stacked. |
| **Scalar** (`int`, `float`, `str`) | **Single Cell** | Renders as a value in a single cell. |
| **Chart** (`matplotlib` / `plotly`) | **Native Chart / Image** | `plotly` figures render as native Excel charts where possible; otherwise as images. |
| **`str`** (File Path) | **Image** | Path to an image (`.png`, `.svg`) is rendered as an embedded picture. |

---

### 3. Excel Formulas
You can inject formulas into Excel cells using the `excel_formula` wrapper.
```python
from tools import excel_formula

# In a DataFrame column
df["Tax"] = [excel_formula(f"=D{i+2}*0.1") for i in range(len(df))]
```

## Example Script

```python
import pandas as pd
from typing import Dict, Any
from tools import run_script_cli, excel_formula

def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:
    # 1. Access inputs
    sales_df = inputs.get("Sales", pd.DataFrame())
    tax_rate = inputs.get("TaxRate", 0.05) # Default if missing

    # 2. Perform Analysis
    sales_df["Total"] = sales_df["Quantity"] * sales_df["Price"]
    
    # 3. Add Excel Formula (Dynamic reference)
    # This places a formula in Excel, not just a static value
    sales_df["FinalPrice"] = [excel_formula(f"=F{i+2}*(1+{tax_rate})") for i in range(len(sales_df))]

    # 4. Return richer results
    return {
        "ProcessedSales": sales_df,         # Returns a Table
        "TotalRevenue": sales_df["Total"].sum(), # Returns a Scalar Value
        "ClientList": sales_df["Client"].unique().tolist() # Returns a List
    }

if __name__ == "__main__":
    run_script_cli(transform)
```

## Project Structure

*   `AddIn/`: Contains the compiled Excel Add-in files (`.xlam`).
*   `Python/`: The backend Python environment.
    *   `.venv/`: Virtual environment.
    *   `Scripts/`: Core system scripts.
    *   `requirements.txt`: Project dependencies.
*   `src/`: Source code.
    *   `module/`: Raw VBA modules (`.bas`, `.cls`, `.frm`) for version control.
    *   `embedded/`: Helper python scripts embedded or used by the system.
*   `userScripts/`: Recommended location for your custom analysis scripts.

## Contact & Support

For support, feature requests, or bug reports, please contact:

*   **Email**: email@example.com
*   **Git Repository**: https://github.com/username/repo

---
Generated for the PyExcel Project.