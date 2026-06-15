# PyExcel.Ribbon

## Macro
The Excel ribbon: the Excel-DNA callback surface that binds UI controls to
`PyExcel.State` and drives the user's actions (Run, Enable/Repair,
Import/Export/Paste, action CRUD, error display). It gates control availability on
the single `ProjectReadiness` verdict and launches the dialogs in `PyExcel.Forms`.

## Files
### PyExcelRibbon.cs
The Excel-DNA ribbon class implementing every callback (`OnRunPython`,
`OnEnablePyExcel`, `OnEditPython`, import/export/paste, action add/edit/delete, error
display). Inputs: the ribbon XML resource and Excel COM via `ExcelDnaUtil` (workbook
context, file dialogs); reads `PyExcelServices` state and health. Output: ribbon control
state and enable/disable gates; invokes `SetupForm` and the various display/picker forms;
mutates state through `StateService`.

## Subdirectories
- **Resources/** — `RibbonUI.xml`, the ribbon layout consumed by Excel-DNA. No
  proprietary scripts (data only), so no nested CLAUDE.md.
