# PyExcel.Addin

## Macro
The Excel-DNA add-in entry point and its COM glue: the AutoOpen/AutoClose
lifecycle, the Application event sink, the active-workbook context, and the
read/write of per-workbook state into the workbook's CustomXMLPart (including
v1→v2 migration). This is the Windows-only layer that wires `PyExcel.State` and
`PyExcel.Ribbon` to live Excel.

## Files
### AddIn.cs
The `IExcelAddIn` implementation. Inputs: none. Output: rooted service instances (file
log, event sink) for the add-in lifetime — initialises logging, checks for 64-bit Excel,
wires `ExcelWorkbookContext` and `AppEventSink`, and restores open workbooks on load.

### AppEventSink.cs
Subscribes to Excel Application events (WorkbookOpen/Activate/BeforeSave/BeforeClose,
SheetActivate) and keeps state in sync. Inputs: a `StateService` and `IWorkbookContext`.
Output: state mutations, on-disk structure validation, readiness marking, ribbon
invalidation, and a state flush to the CustomXMLPart on save; every handler is guarded so
exceptions never escape into Excel.

### WorkbookStatePersister.cs
Serialises/deserialises a workbook's PyExcel profile to/from its CustomXMLPart, reading
the current full-profile format and falling back to the legacy state-only part and v1
defined-Names migration. Inputs: an Excel Workbook, a `WorkbookProfileData`, the project
directory, and workbook name/path. Output: XML persisted into the workbook, or a loaded
profile; all COM work is best-effort and logged, never thrown.

### ExcelWorkbookContext.cs
Concrete `IWorkbookContext` reading the active workbook's identity over late-bound COM,
caching last-good values to survive transient COM faults. Inputs: none (reads Excel COM).
Output: `CurrentWorkbookKey` (path or unsaved-synthetic key) and
`CurrentWorkbookDirectory`; null when no workbook is active.

### LegacyStateReader.cs
Windows-only COM reader for v1 PyExcel defined Names (sheet- and workbook-scoped) and
their formulas, decoded via `LegacyFormulaDecoder` without `Evaluate`. Inputs: an Excel
Workbook. Output: a `LegacyWorkbookState`, or null; COM faults are swallowed.

## Subdirectories
None.
