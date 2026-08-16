# PyExcel.Forms

## Macro
The WinForms dialog shells for every Excel-integrated interaction — add/edit
action, Import/Export/Paste editors, sheet/range pickers, the setup wizard, and
the run-progress and error viewers. Each dialog's decision logic is split into a
cross-platform validator/helper (no WinForms) so it is unit-tested on Linux CI;
the WinForms shells themselves are net48-only behind `#if NETFRAMEWORK`.

## Files
### EditActionForm.cs
Dialog to add/edit a ribbon action (name, script, input/output range lists, kwargs, and
a keep-output-window-open checkbox), with a "New script" scaffold button. Inputs:
available scripts, existing action names, an optional action to edit, a range-picker
function, and the userScripts directory. Output: a `RibbonAction?` (null on cancel).

### EditActionValidator.cs
Cross-platform validator that builds the action from the dialog fields (unique name,
present script/ranges, parsed kwargs, keep-output-open flag). Inputs: name, script,
input/output strings, optional kwargs, existing names, original name, and the flag.
Output: an `EditActionValidationResult` (valid flag, error, or built `RibbonAction`).

### EditIoForm.cs
Unified dialog for Import/Export/Paste with parameterised labels and field kinds (range
vs file), validated via `EditIoValidator` on OK. Inputs: title, labels and field kinds,
initial values, workbook dir, and an optional range picker. Output: an
`EditIoValidationResult?` (null on cancel).

### ExportForm.cs
The unified Export dialog, run in two modes: **Edit** (`PromptDefaults`) configures and
returns the workbook's default export recipe; **Export** (`PromptExport`) seeds from those
defaults, lets the user tweak, and returns the recipe to run now (plus a save-as-default
flag). Both edit an `ExportSettings` (source range, folder, base name, file type, optional
unique-name date/time stamp) with a live file-name preview and an overwrite guard; the
composition rules come from the cross-platform `ExportSettingsPlanner`. Inputs: an initial
`ExportSettings`, the workbook dir, and an optional range picker. Output: an `ExportSettings?`
(Edit) or `ExportPromptResult?` (Export), null on cancel.

### ProgressForm.cs
Modeless, non-blocking progress dialog that renders kernel PROGRESS frames on a background
thread and offers Cancel via a `CancellationToken`; implements `IRunProgressSink`. Inputs:
an owner window and a run title. Output: the form itself, exposing `CancellationToken`.

### ErrorDisplayForm.cs
Read-only, resizable modal viewer for the last kernel error traceback (front-most, with a
Copy button) — also reused to show a successful run's captured output. Inputs: an owner
window, a title, and the body text. Output: none (static `Open`).

### SetupForm.cs
First-run setup-wizard shell hosting the headless `SetupService`, streaming venv/pip
output to a live log box. Inputs: an owner window, a project path, and an optional log
sink. Output: a `bool?` (null if dismissed before completion).

### SheetPickerForm.cs
Dialog to pick a sheet when importing a multi-sheet Excel file without a pinned `!Sheet`.
Inputs: an owner window, the sheet names, and an optional preselection. Output: a
`string?` (canonical sheet name, null on cancel).

### SheetPickerValidator.cs
Cross-platform validator ensuring a chosen sheet belongs to the offered list
(case-insensitive, returns canonical casing). Inputs: the selected sheet and the available
sheets. Output: a `SheetPickerValidationResult` (valid flag, error, canonical name).

### RangeNameForm.cs
Dialog to edit a single `RangeBinding` — a range address (with native picker) and optional
name, validated by `RangeAddressValidator`. Inputs: an owner window, an optional initial
binding, and a range picker. Output: a `RangeBinding?` (null on cancel).

### RangeListEditor.cs
Reusable control hosting an ordered list of range bindings with Add/Edit/Remove/Up/Down,
loading from and serialising back to `{name}=range` syntax. Inputs: an optional range
picker. Output: the binding text via `ToBindingText()`.

### RangeAddressValidator.cs
Cross-platform validator ensuring an entry is a single plain range (no semicolons or name
bindings) and parses through `RibbonRangeParser`. Inputs: an address string. Output: a
`RangeAddressValidationResult` (valid flag, error, trimmed address).

### RangePick.cs
Helper to invoke Excel's native range picker from inside a modal: hides the dialog(s) so
the sheet is interactive, runs the picker, then restores. Inputs: a picker function, an
initial address, and the form(s) to hide. Output: the picked address, or null.

### OrientationForm.cs
Dialog asking whether a spilled 1-D result lays out as a row or a column when the target
is a single cell. Inputs: an owner window. Output: a `ListOrientation?` (null on cancel).

### KwargsText.cs
Cross-platform parser/formatter for the kwargs text box (`key=value` per line, all line
endings, interior-whitespace-preserving values, trimmed keys). Inputs: raw text. Output:
an `IReadOnlyDictionary<string,string>?`, a formatted string, or an error string.

### ScriptScaffold.cs
Cross-platform utility that creates a new user script from a starter `transform()`
template, sanitising the file name and resolving collisions with `_1`, `_2`, …. Inputs:
the userScripts dir and a desired name. Output: the safe file name (no directory).

### TextPromptForm.cs
Minimal single-line text prompt (InputBox replacement), e.g. for a new script name.
Inputs: an owner window, a title, a label, and an optional initial value. Output: a
`string?` (trimmed input, null on cancel).

### ScaledForm.cs
Abstract base for every dialog, applying DPI-aware scaling (Win32 `GetDpiForWindow`)
across the control tree and fitting to the screen. Inputs/Output: UNVERIFIED (base class;
scaling applied in place during load).

## Subdirectories
None.
