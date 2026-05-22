# PyExcel — Production Readiness Audit

**Date:** 2026-05-22
**Scope:** Full repo (~19k LOC) — v1 add-in (VBA `.xlam`), embedded Python runtime, v2 .NET rewrite.

## Verdict

**Not production ready.** The repository contains two parallel codebases:

- **v1** — the shipping product (`PyExcel.xlam` + `src/module/*.bas` + `src/embedded/*.py`). Functional but carries serious bugs, ~40–70% dead code in several files, no automated tests, debug artifacts everywhere, and silent failure modes.
- **v2** — a .NET / Excel-DNA rewrite that is **~5% complete** (Phase 1 ribbon skeleton + one Python framing module). Every v2 ribbon button is a stub; there is no kernel, no bridge, no run pipeline.

The two cannot both ship. A go/no-go decision on v1-vs-v2 is the first task. The list below is written assuming **v1 must be made shippable** and **v2 must either be finished or removed from the release branch**.

---

## 0. Critical blockers (fix before any release)

- [ ] **Typosquatted dependency.** `requirements.txt:84` lists `panadas==0.2` — a typosquat of `pandas`. This is a supply-chain risk; remove it immediately and audit how it got in.
- [ ] **Two conflicting requirements files.** Root `requirements.txt` (155 pkgs, `numpy==1.26.0`, `pandas==2.1.1`) vs `src/embedded/requirements.txt` (39 pkgs, `numpy==2.3.2`, `pandas==2.3.2`). Decide which is canonical; the README says the embedded one is extracted and installed, so the root file is misleading/dead.
- [ ] **`src/embedded/requirements.txt` is UTF-16 encoded** (BOM + null bytes between every char). `pip` will not reliably parse it. Setup.bas has a `FixRequirementsEncoding` band-aid — fix the *source file* to plain UTF-8 instead.
- [ ] **Personal data shipped to users.** `src/embedded/xmlParsing.py:1037` hardcodes `C:\Users\Jonatan\Documents\Education\Msc\Master Banking and Finance\...`. The `if __name__ == "__main__"` block (lines ~770–1343) is ~570 lines of the author's private coursework, shipped inside the add-in's embedded runtime. Delete it.
- [ ] **Orphaned Python subprocess.** `python.bas RunPythonJob` launches `python.exe` with `sh.Run cmd, 1, False` and discards the handle. On timeout/stall the VBA gives up but **never kills the process** — it keeps running, holding file locks, colliding with the next run. Track and terminate the process.
- [ ] **`tempFiles("out")` is never set.** `python.bas RunPythonJob` (~line 514) interpolates `--out <tempFiles("out")>` but the dict only ever gets an `"in"` key; reading the missing key yields `Empty`, so Python receives an empty `--out` path. Verify the output-file contract end to end.
- [ ] **Decide the fate of v2.** `embedded/pyexcel/kernel/` has only `framing.py`; `src/PyExcel.Addin` + `src/PyExcel.Ribbon` are a Phase-1 skeleton where every button logs a stub. Either finish it or move it off the release branch so it isn't mistaken for working code.

---

## 1. Needs a fix — bugs

### Embedded Python runtime (`src/embedded/`)

- [ ] **Unreachable duplicate branch.** `tools.py:397` and `tools.py:423` are both `elif isinstance(val, str) and os.path.exists(val)`. The second is dead code; the two also disagree on which extensions are valid (`.xml` allowed vs not).
- [ ] **Deprecated API.** `datetime.utcnow()` used in `tools.py:159,176` and `xmlParsing.py:548` — removed-path deprecation in Python 3.12+. Use `datetime.now(timezone.utc)`.
- [ ] **`Any` referenced but not imported.** `xmlParsing.py:66-67` annotate with `Any`; line 16 imports only `Union, Sequence, Mapping, Optional, Dict`. Breaks type-checking / would `NameError` if annotations were evaluated.
- [ ] **Duplicate/shadowing imports.** `xmlParsing.py:16-17` import `Mapping, Sequence` from both `typing` and `collections.abc`.
- [ ] **Inconsistent error handling in `read_xml`.** `xmlParsing.py:268` `<value>` int parsing is `out = int(raw)` with no guard and crashes on bad data, whereas `<list>` parsing (lines 205-225) catches `ValueError`. Make them consistent.
- [ ] **Timestamp heuristic misclassifies.** `read_xml` tries `pd.to_numeric` first for `timestamp` columns (lines 159-172). A column of plain years/IDs (e.g. `2020, 2021`) is silently reinterpreted as Excel serial dates.
- [ ] **Docstring/signature mismatch.** `run_script`/`run_script_cli` in `tools.py` are typed/documented as taking "a list of input DataFrames"; they actually receive a `dict[str, Any]` from `read_xml`.
- [ ] **`bool` classified as `integer`.** `tools.py:279-281` scalar branch — `bool` is a subclass of `int`, so `True/False` outputs are written with `datatype="integer"`.
- [ ] **Locale-sensitive numeric output exists on the VBA side** (see below) but the Python reader assumes `.` decimals — confirm round-trip on non-US locales.

### VBA — run pipeline (`python.bas`, `xmlParsing.bas`, `pythonUtils.bas`)

- [ ] **Command-line quoting/injection.** `python.bas` interpolates `scriptPath`, `metaFile`, `logOut`, `runId` into a `cmd /c "..."` string with a nested PowerShell `Tee-Object`. A project path containing a `'` or `"` breaks the command; nothing is escaped.
- [ ] **Run-id mismatch hangs silently.** `python.bas` poll loop (~line 591): if Python writes a meta whose `run_id` differs/empty, VBA loops to the 300 s `maxWait` and reports a generic timeout.
- [ ] **Heartbeat false stall.** A slow-but-healthy script that doesn't rewrite `<timestamp>` is declared "stalled" at 120 s; conversely a meta read mid-write parses as corrupt → abort. Add atomic-read / retry.
- [ ] **`ws` parameter reassigned in loop.** `RunGenericPythonScript` (`python.bas:77,86`) mutates `ws` while iterating source ranges; the *last* range's sheet becomes the default output sheet — multi-sheet inputs write outputs to the wrong sheet.
- [ ] **`On Error Resume Next` keeps stale `ws`.** `python.bas:76-79` — an invalid `parsedSheet` leaves `ws` pointing at the previous valid sheet, skipping the "sheet not found" check.
- [ ] **`CLng` overflow on large integers.** `xmlParsing.bas:379` single-cell path `CLng(scalarVal)` overflows for integer values > 2,147,483,647 (10-digit IDs) → run fails with "Failed to serialize input."
- [ ] **Locale-dependent number formatting.** `xmlParsing.bas` uses `CStr(CDbl(v))` / `Format$` (lines ~495, 502, 713, 719) — emits `0,5` on comma-decimal locales, producing invalid XML numerics.
- [ ] **Scalar `IsNumeric` misclassification.** `xmlParsing.bas:378` treats text like `"1E5"`, `"$5"`, `" 12 "` as numbers; `"true"/"false"` text never reaches the bool branch.
- [ ] **`PasteTypedXMLToRange` clobbers the last sheet row.** `xmlParsing.bas:~1560` copies formats via `wsDst.Cells(wsDst.rows.count, …)` and only `ClearFormats` afterward — destroys any real formatting on row 1048576.
- [ ] **Multi-chunk extraction gap bug.** Both `Setup.bas ExtractEmbeddedStoreUnified` and `Update.bas ExtractResources` iterate `For i = 1 To chunks.count` over a `Scripting.Dictionary`; any missing/0-based chunk index → silently truncated, corrupt file written with no error. Iterate sorted keys.
- [ ] **`PasteArtifactsToTargets` runs with events/calc live** — slow and can re-trigger `Worksheet_Change` mid-paste. Disable `ScreenUpdating`/`Calculation`/`EnableEvents`.
- [ ] **`Exit Function` abandons remaining artifacts.** `pythonUtils.bas:~971` — cancelling the orientation dialog for one `list` artifact aborts the whole paste loop. Use `GoTo nextItem`.
- [ ] **List paste has no bounds check.** `pythonUtils.bas:~977-989` `Resize` ignores the user's destination size; a long list silently overwrites cells beyond the selection.
- [ ] **`ArchiveFile`/`CleanTempFolder` swallow failures.** Locked files (held by the orphaned python.exe) fail silently; the user is told cleanup succeeded.

### VBA — ribbon & state (`modRibbon.bas`, `HostManager.bas`, `PathUtils.bas`)

- [ ] **Stale `currentSheetName` global.** Several getters (`GetPyOutput`, `GetImportInput`, etc.) read/write the wrong sheet's named ranges after a sheet switch because they don't re-sync from `HostManager_GetCurrentSheet()`. `GetPyInput` syncs; `GetPyOutput` does not — inconsistent.
- [ ] **`SaveSheetValue` can infinite-loop.** `modRibbon.bas:~282-294` — if the inner chunk shrinks to length 0, the outer `pos = pos + Len(chunk)` never advances → Excel hangs.
- [ ] **Watchdog never cancelled.** `HostManager_Watchdog` reschedules via `Application.OnTime` every 10 s and `HostManager_Shutdown` never cancels it; `HostManager_PollStartup` self-reschedules every 0.5 s forever if no user workbook appears.
- [ ] **UNC paths unsupported.** `PathUtils.ResolveProjectPath` only handles drive letters and `*.sharepoint.com`; `\\server\share\…` returns `""` → all updates abort for network-drive users. `MkDirRecursive` also breaks on UNC.
- [ ] **SharePoint resolution is English/default-library only.** `PathUtils` hardcodes `/Documents` (and its 10-char length via `Mid$`) and `OneDriveCommercial` — fails for localized Office, personal OneDrive, or non-default libraries. `DecodeUrlComponent` is single-byte only — mojibake on UTF-8 folder names.
- [ ] **`VersionToNumber` comparison flaws.** `Update.bas:~767` — patch multiplier of 100 collides versions (`1.2.150` vs `1.3.50`); 14-digit timestamp versions always dwarf semver, so a semver build never registers as newer than a timestamp project.
- [ ] **Version-string corruption.** `GetStoredProjectVersion` does `Replace(..., "=", "")` and strips quotes — corrupts any version containing those chars.

### VBA — charts, import/export, forms

- [ ] **frmExportWizard buttons are dead.** Dynamically created `srcEdit_`/`srcRemove_`/`dstEdit_`/`dstRemove_` buttons have **no event wiring**; `SourceEditClicked` etc. are unconnected `Private Sub`s. The wizard cannot edit or remove rows — core feature non-functional.
- [ ] **frmExportWizard uses unqualified `Range()`.** `btnExport_Click:~381` `Range(srcTxt.text)` resolves against the active sheet, but sources are stored as external `[Book]Sheet!` addresses → "Invalid range" for every multi-sheet source.
- [ ] **CSV parsing not RFC-4180 safe.** `Import.bas ReadCSVToArray` splits on `vbNewLine` before quote handling → embedded newlines corrupt rows; opens files as ANSI (`-2`) so UTF-8 is mis-decoded; `.tsv` delimiter is overwritten by `DetectDelimiter`. `chartBuilder.bas CSVToStringArray/CSVToNumericArray` and `Paste.bas TextToColumns` have the same naive-split problem.
- [ ] **CSV export under-quotes.** `Export.bas ExportRangeToCSV:~36` quotes only on `, " \n` — a field with a bare `\r` breaks the file; single-cell export builds a jagged array that the row loop mishandles; error values (`#N/A`) throw.
- [ ] **chartBuilder leaves orphan charts.** `BuildChartFromXML` adds a `ChartObject` then `Exit Sub`s on any invalid/duplicate trace id — an empty chart is left in the workbook. No top-level `On Error` handler.
- [ ] **chartBuilder null-attribute crashes.** `RenderExtrasAnnotations:~1840` and `DrawLineAnnotation:~1937` dereference `getNamedItem("type"/"axis"/"mode").text` assuming the attribute exists — missing attr crashes the build.
- [ ] **chartBuilder marker mapping swapped.** `ShapeToMarkerStyle:~1614` maps `"cross"`→`xlMarkerStyleX` but `"x"`→`xlMarkerStylePlus`.
- [ ] **`SheetPickerForm` lists the wrong workbook.** It enumerates the *active* workbook's sheets even when `Import.bas ReadExcel_COM` passes an external source workbook.
- [ ] **frmEditExport `btnEditOutput_Click:~270`** has no `ListIndex < 0` guard → `List(-1)` throws when nothing is selected.
- [ ] **Verify Import.bas externals compile.** `Import.bas` calls `PrepareOutputRange`, `CaptureRowFormat`, `ApplyFormatToRange`, `ClearExcessRange` — confirm all are defined (most are in `modDst.bas`; `CaptureRowFormat` needs verifying) or the module won't compile.
- [ ] **`PrepareOutputRange` strips user formatting.** `modDst.bas` resets the destination's font/size and clears all formatting unconditionally — no warning, no undo.

---

## 2. Not ready — incomplete, stubs, dead code

- [ ] **v2 is a skeleton.** `PyExcelRibbon.cs` — every `OnAction` except `OnReadMe` is `StubAction(...)`; all getters return defaults; `RibbonEnabled` always returns `false`; `LoadImage` always returns `null`. `AddIn.cs` `AutoOpen` explicitly "Phase 1 stops here". `embedded/pyexcel/kernel/` has only `framing.py` — no supervisor, worker, or `__main__`.
- [ ] **`tools.py` is ~70% dead code.** Of 1748 lines, ~1200 are commented-out legacy implementations (old `read_xml`, `write_xml`, `write_chart_xml`, `_chartspec_from_plotly/matplotlib`, etc.). Delete.
- [ ] **`xmlParsing.py` is ~50% dead code.** ~485 commented lines in the tail + the ~570-line `__main__` coursework block. Delete both.
- [ ] **`xmlParsing.bas` is ~70% dead code.** Lines ~4-205 and ~845-1489 are commented-out old `SerializeRangeToTypedXML` and two prior `PasteTypedXMLToRange` versions.
- [ ] **`chartBuilder.bas` ~1000 lines dead.** Lines ~2-1010 are an entire alternate `cs:`-namespace chart engine, commented out. Plus an explicit no-op "Dummy sentinel" at ~1054.
- [ ] **`frmEditActionOLD.frm/.frx` is superseded** by `frmEditAction.frm` — delete the pair.
- [ ] **Commented-out duplicate class.** `CAppEvents.cls:10-43` is a full commented copy of the class.
- [ ] **Dead first-gen blocks in Edit forms.** `frmEditImport.frm`, `frmEditExport.frm`, `frmEditPaste.frm` each carry a large commented-out original `Initialize`/`btnSave` block.
- [ ] **modRibbon dead handlers.** Two commented-out `OnEnablePyExcel`, one commented `RibbonOnLoad`, a commented duplicate getter block.
- [ ] **python.bas dead code.** ~150 lines of the old `Py()` / `RunPythonJob` commented out.
- [ ] **Duplicated logic across Setup.bas & Update.bas.** `FixRequirementsEncoding`, `Base64ToBinary`, `WriteBinaryFile`, `CreateFoldersRecursive`, `EnsureFolderExists`, progress-bar subs all exist twice and have **already diverged** (different `MSXML2.DOMDocument` versions). Consolidate into one module.
- [ ] **Incomplete A1/R1C1 formula round-trip.** `xmlParsing.bas PasteTypedXMLToRange` reads `mode="a1"`/`a1=` attributes that `SerializeRangeToTypedXML` never emits; README also says R1C1 is "stubbed but not implemented".
- [ ] **`Import.bas ReadExcel_ADO` (~100 lines) is fully written but dead** — routing always uses `ReadExcel_COM`.
- [ ] **Dead helpers:** `Setup.SortVariantNumeric`, `Update.BuildPathKey`, `pythonUtils.IsEmfFile`, `xmlParsing.bas ShuffleArray/JoinCollection/attr$`, `Import.PickFile/UBound2D/Variant2DSize`, `frmExportWizard.SortKeysNumeric` — none called.
- [ ] **`itemType` is a placeholder.** `frmEditAction.frm` hardcodes `"Range"` everywhere — the "Type" column is non-functional.
- [ ] **`pip_install.log` is never written.** `Setup.InstallPipPackages` computes and advertises the log path but never redirects pip output to it — the "Full pip output saved to…" message is false.
- [ ] **Debug artifacts everywhere.** `#Const DEV = True` in `HostManager.bas`; `Debug.Print` banners/timing dumps across `Setup.bas`, `Update.bas`, `modDst.bas` (`ResolveDestinationRange` dumps `Hwnd`/`Application.Ready`), `python.bas` ("Full command: …"), `pythonUtils.bas`, `modRibbon.bas`, `Import.bas`, `Export.bas`, `chartBuilder.bas`, and several forms. A pre-run debug `MsgBox` in `OnRunPython` (`modRibbon.bas:~1494`) forces a click before every run.
- [ ] **Version numbers inconsistent.** README `20260422_212123`, `Directory.Build.props` `2.0.0-alpha`, `pyexcel/__init__.py` `2.0.0a0`. Pick one scheme.
- [ ] **API typo baked in.** `python.bas` public parameter / action key `entreToEnd` (should be "enter").
- [ ] **No CI in repo.** `docs/v2-build.md` describes "what runs in CI" but there is no `.github/workflows/` (or other CI config) present.

---

## 3. Inefficient

- [ ] **Bloated dependency set.** Root `requirements.txt` has 155 packages — `jupyterlab`, `pygame`, `yt-dlp`, `gTTS`, `python-telegram-bot`, `Flask`, `Flask-SocketIO`, `tkcalendar`, `pyodbc`, `mysql-connector`, `psycopg2`, etc. — almost none relevant to an Excel transform runner. Two `git+https` deps (`db_utils`, `glog`) require `git` on the user machine at install time. This makes `pip install` slow, fragile, and large.
- [ ] **Heavy eager imports.** `xmlParsing.py` imports `plotly.graph_objects` and `matplotlib.pyplot` at module scope; `tools.py` imports `matplotlib.figure/axes` at top. Every script run pays this import cost even when no chart is produced. Defer to use sites.
- [ ] **Meta polling re-parses the whole file every 200 ms.** `python.bas` rebuilds an `MSXML2.DOMDocument` each poll for up to 300 s (~1500 parses), then parses again via `ParseMetaXml` on completion.
- [ ] **`ResolveProjectPath()` called 3+ times per run** in `python.bas`.
- [ ] **Table artifacts parsed twice.** `pythonUtils.PasteArtifactsToTargets` reads the file with `ReadTextFromFile`, then `PasteTypedXMLToRange` re-parses the same string into a DOM.
- [ ] **Ribbon callbacks re-hit the filesystem.** `GetScriptFiles`/`GetActionList` re-enumerate the scripts folder / re-parse the Actions named range on every callback; listing N scripts hits the FS N+1 times. `LoadActionsForSheet` re-parses on every refresh instead of using its cache.
- [ ] **Redundant ribbon invalidation.** `HostManager_RibbonRefreshAll` invalidates ~25 controls individually *and then* calls full `.Invalidate`; the watchdog does this every 10 s.
- [ ] **Per-cell ADODB writes.** `xmlParsing.bas SerializeRangeToTypedXML` calls `stream.WriteText` per row/cell — COM overhead scales linearly; build the XML in memory and write once.
- [ ] **Range resolved twice.** `SerializeRangeToTypedXML` resolves each input part once for trimming and again for serialization.
- [ ] **Blocking shell calls.** `Setup`/`Update` run venv + pip via `sh.Run …, True`, freezing Excel's UI thread for minutes with a non-updating progress bar; three separate blocking `cmd` launches for pip upgrade/install/freeze.
- [ ] **Unbounded log growth.** `HostManager.LogToFile` appends to `%TEMP%\PyExcel_Debug.log` with no rotation/size cap; with `#Const DEV = True` it logs on every event.
- [ ] **`ResolveAddressToRange` scans all workbook Names** for every unqualified address instead of a direct lookup.

---

## 4. Not user friendly

- [ ] **Failures are silent.** Pervasive `On Error Resume Next` and `Debug.Print`-only handlers mean paste/serialize/archive/import/export/chart failures produce **no user-visible message** — `pythonUtils.PasteArtifactsToTargets`, `modRibbon` change handlers, `Export.bas`, `chartBuilder` annotation errors, `UpdatePythonDependencies` (reports "successfully updated" even when pip failed), and more.
- [ ] **No progress / no cancel during a run.** The only feedback while Python runs is a console window; Excel's UI is blocked by the poll loop. No status bar, no progress dialog, no cancel button. `ufProgress` exists but isn't used for runs or imports.
- [ ] **Generic, undiagnosable errors.** Serialize failures collapse to "Failed to serialize input." with no range/cell/reason. Timeouts say "Max wait time reached" with no script name, log path, or traceback. `PathUtils` has five distinct failure causes that all surface as one "Could not resolve project path."
- [ ] **Debug `MsgBox` before every run.** `OnRunPython` pops a modal dialog dumping Action/Script/Input/Output that the user must dismiss every single time.
- [ ] **Truncated tracebacks.** Python tracebacks are cut to ~800 chars (`python.bas:~290`) — the actual failing line is often lost.
- [ ] **Off-screen form hack.** `frmEditAction`, `frmEditImport`, `frmEditPaste`, and others move the form to `(-20000,-20000)` to show a range picker; any error in between leaves the form permanently off-screen — only an Excel restart recovers it.
- [ ] **Ribbon editboxes silently revert.** `OnImportInputChange` etc. just `InvalidateControl`, so typing into the ribbon field vanishes with no explanation; the real edit path is a separate dialog.
- [ ] **Auto-prompt on workbook open.** `Update.ShowUpdatePrompt` fires a blocking `MsgBox` ~1 s after a workbook activates, interrupting the user.
- [ ] **No input validation.** `OnRunPython` launches with empty script/input/output fields; the user discovers the problem via a downstream `InputBox` or a crash. Export wizard aborts the whole batch on the first bad range with no indication of which items succeeded.
- [ ] **Destructive paste with no confirmation.** `Paste.bas` and list/table paste overwrite existing destination cells (and formatting) with no prompt or undo affordance.
- [ ] **Multi-sheet pickers are plain `InputBox`es.** `frmEditAction.DoImportFromWorkbook` makes the user retype a sheet name instead of choosing from a list.
- [ ] **Setup/Update freeze with no detail.** Pointless `Application.Wait` pauses; progress bar jumps 0.7→1.0 across the multi-minute pip phase; when Python is missing or is the Windows Store stub, the user gets "venv creation did not complete" with zero diagnostics (stdout/stderr discarded).
- [ ] **80% package threshold treated as success.** `Setup.VerifyPipPackages` accepts a partially broken install as "ready", so the user hits failures later instead of at setup time.

---

## 5. Security

- [ ] **Typosquat package** `panadas==0.2` (see §0).
- [ ] **Command construction from user-controlled paths.** `python.bas` (`cmd /c` + PowerShell `Tee-Object`), `modRibbon` `Shell "explorer.exe …"` / `cmd /c start …` interpolate workbook/project paths without escaping quotes — a path containing `'` or `"` can break or alter the command.
- [ ] **`git+https` dependencies** pull code from external GitHub repos pinned by commit — acceptable if intentional, but document it and confirm the repos are trusted and durable.
- [ ] **Embedded runtime extracted and executed.** Setup decodes base64 from a hidden `EmbeddedStore` sheet and writes/executes Python — ensure integrity (the `meta.xml` artifact `sha256` exists; nothing verifies the *extracted runtime* files).

---

## 6. Testing & process gaps

- [ ] **Zero tests for the shipping product.** The only test file, `tests/kernel/test_framing.py`, covers v2's `framing.py`. There are **no tests** for any VBA module, `tools.py`, or `xmlParsing.py`.
- [ ] **No round-trip tests** for the core contract: Excel range → typed XML → DataFrame → output XML → Excel. This is where most of the bugs above live (locale, types, dates, overflow).
- [ ] **No CI pipeline** present in the repo despite the build doc describing one.
- [ ] **README claims vs reality.** README documents v1 behavior; `docs/v2-build.md` documents an unfinished v2. A user reading the README and loading the repo will be confused about what actually works.

---

## Suggested order of work

1. **§0 blockers** — typosquat, requirements consolidation/encoding, strip personal data, subprocess lifecycle, v2 go/no-go.
2. **§1 run-pipeline bugs** — quoting, run-id/heartbeat, `ws` reassignment, locale numerics, `CLng` overflow, chunk-gap extraction.
3. **§2 dead-code purge** — removes ~3,000+ lines and makes the rest auditable.
4. **§4 user-facing failures** — surface errors, add run progress/cancel, remove the debug `MsgBox`.
5. **§3 efficiency** — dependency slimming, eager-import deferral, polling/ribbon-callback caching.
6. **§6 add tests** — round-trip coverage for the serialization contract before further change.
