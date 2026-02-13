**Version:** 20260213_183611

---

# PyExcel

An Excel add-in that enables execution of Python scripts directly from the workbook within a dedicated, isolated environment. Scripts run non-reactively, ensuring deterministic and reproducible execution. The add-in allows direct use of Python��s data analysis and visualization ecosystem including Pandas, NumPy, Plotly, and sqlite3 while maintaining clear separation between Excel and the Python runtime. This approach provides a controlled, reliable way to integrate advanced analytics and data processing into Excel workflows without compromising stability or speed.

## Features

*   **Seamless Integration**: Call Python scripts directly from Excel.
*   **Flexible Inputs**: Automatically detects and converts Excel ranges into Pandas DataFrames, Lists, or Scalar values. Supports named variables for inputs.
*   **Rich Output**: Return modified tables, lists, summary values, or complex objects like Plotly/Matplotlib charts back to Excel.
*   **Excel Formulas**: Inject dynamic Excel formulas into your output dataframes.

## Prerequisites & Setup

### 1. Python Environment
Ensure you have **Python 3.7+** installed. 

#### We use a Virtual Environment
PyExcel uses a **Virtual Environment (venv)** located in `Python/.venv`. 
*   **Isolation**: It ensures that libraries installed for PyExcel don't interfere with other Python projects on your computer, and vice versa.
*   **Stability**: It guarantees that the exact versions of tools (like Pandas or Plotly) needed by PyExcel are always available, regardless of what else you install on your system.
*   **Portability**: It makes it easier to update or move the project without breaking your global Python setup.

#### Installing Libraries
Because of this isolation, if you want to install a new library to use in your scripts, you must install it to Virtual Environment by using the specific path to the project's Python:

```bash
# From the project root
Python/.venv/Scripts/python.exe -m pip install {LibraryName}
```
*(See `Python/requirements.txt` first to see what's already included)*


### 2. Installation

1.  **Security Configuration**:
    *   **Trust Center**: Go to *File > Options > Trust Center > Trust Center Settings > Macro Settings* and check **Trust access to the VBA project object model**.
    *   **Unblock File**: Navigate to the `AddIn/` folder, right-click `PyExcel.xlam`, select **Properties**, check **Unblock**, and click **Apply**.

2.  **Activation**:
    *   Open Excel and go to *File > Options > Add-ins*.
    *   At the bottom, manage **Excel Add-ins** and click **Go**.
    *   Click **Browse**, select `AddIn/PyExcel.xlam`, and ensure it is checked in the list.

3.  **Initialization**:
    *   Once loaded, click the **Enable** button on the ribbon.
    *   This triggers the **Setup Wizard**, which will:
        *   Ask you to select a location for your project files.
        *   Create the necessary folder structure (`Python/`, `userScripts/`, etc.).
        *   Set up the isolated Python virtual environment.
        *   Install required libraries automatically.

## Update Process

To update PyExcel to a newer version:

1.  **Overwrite**: Simply replace your existing `PyExcel.xlam` file with the new version.
2.  **Automatic Self-Update**:
    *   The next time you open Excel, the Add-in will detect the version change.
    *   It will automatically clean old files, extract new resources, and synchronize Python dependencies from the internal `requirements.txt`.

## Usage

### 1. Defining Inputs in Excel
The system parses input ranges provided from Excel. You can select ranges or pass a semicolon-separated string of addresses.

*   **Variable Naming**: Use `{name}=Range` syntax to assign specific names to your inputs.
    *   Example: `{Sales}=Sheet1!A1:C10; {TaxRate}=Sheet1!E1; {Months}=Sheet1!A1:A12`

#### Input Type Mapping

| Excel Range Shape | Detected Type | Python Type | Description |
| :--- | :--- | :--- | :--- |
| **Multi-row & Multi-column** | Table | `pandas.DataFrame` | Converted to a DataFrame. Column types (int, float, date) are inferred automatically. |
| **Single Row or Single Column** | List | `list[str]` | Converted to a list of strings. |
| **Single Cell** | Scalar | `int`, `float`, `bool`, `str`, `pd.Timestamp` | Converted to a native Python scalar value. |

> **Note**: If no name is provided, inputs are named automatically: `df1`, `df2`... for tables; `list1`... for lists; `value1`... for scalars.

---

### 2. Writing Python Scripts
Add your Python scripts in the `userScripts/` directory. A script must define a `transform` function that acts as the entry point and import the tools package already included in this addin.

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

#### Output Type Mapping

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

*   **Email**: JonatanGani@protonmail.com
*   **Git Repository**: [https://github.com/username/repo](https://github.com/Jonatan-Gani/PyExcel)