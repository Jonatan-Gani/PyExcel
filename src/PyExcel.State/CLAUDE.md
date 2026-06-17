# PyExcel.State

## Macro
Per-workbook state, persistence, and the service locator the ribbon reads. It
holds the in-memory registry of each workbook's configuration (per-sheet
profiles, actions, field bindings), the XML codecs that round-trip that state
through a workbook's CustomXMLPart, the on-disk run archive, the last-error
registry, and v1→v2 migration. No COM here — everything is data and pure logic,
so it builds and tests on Linux; the COM-bound persistence wiring lives in
`PyExcel.Addin`.

## Files
### WorkbookState.cs
Immutable per-workbook snapshot the ribbon renders from. Inputs: workbook key, enabled
flag, sheet name, scripts list, selections, and range bindings from `StateService`.
Output: a read-only `WorkbookState` record with a computed `SelectedAction`.

### SheetProfile.cs
The per-sheet configuration slice (selected script, Run bindings, Import/Export/Paste
fields, actions) that `StateService` projects into a `WorkbookState`. Inputs: none
(immutable record). Output: `SelectedScript`, `PyInput`/`PyOutput`, the I/O field
bindings, `Actions`, `SelectedActionName`, plus `IsConfigured` and `SelectedAction`.

### StateService.cs
In-memory registry mapping a workbook key to per-sheet profiles and workbook-scoped
facts (enabled, project dir, stable identity), serialised by one coarse lock and firing
`StateChanged` on mutation. Inputs: a workbook key, sheet name, `RibbonAction`, and mutator
functions; `SetIdentity` writes the project id + origin path, `Rekey` atomically moves an
entry between keys (a Save As / move / rename). Output: `WorkbookState` projections,
`WorkbookProfileData` snapshots for persistence, and the `StateChanged` event.

### WorkbookProfileData.cs
The full persisted shape of a workbook's project (enabled flag, project dir, stable
identity `ProjectId` / `OriginPath`, per-sheet `SheetProfile` map with default-bucket
inheritance). Inputs: none (immutable record). Output: `Enabled`, `ProjectDir`,
`ProjectId`, `OriginPath`, `Sheets`; an `IsMeaningful` check and a `FromState` converter
from the flat `WorkbookState`.

### WorkbookProfileCodec.cs
XML round-trip for `WorkbookProfileData` (workbook flags + sheet map), self-contained so
the format can evolve independently; serialises only configured sheets in ordinal order.
Inputs: a `WorkbookProfileData`, `XElement`, or XML string. Output: an `XElement` or
string; `TryParseElement`/`TryDeserialize` return a bool with the parsed profile.

### WorkbookStateCodec.cs
XML round-trip for the flat `WorkbookState` (user-editable fields + actions) stored in a
CustomXMLPart, schema version 1, namespace `urn:pyexcel:state:1`. Inputs: a
`WorkbookState`, `XDocument`, or XML string. Output: an `XDocument` or compact string;
`TryDeserialize` returns a bool with the parsed state.

### ProjectProfile.cs
Root record pairing `WorkbookProfileData` (configuration) with `ProjectMetadata`
(provenance) — the complete identity of a PyExcel project embedded in a workbook.
Inputs: none (immutable record). Output: `Profile` and `Metadata`.

### ProjectProfileCodec.cs
XML round-trip for `ProjectProfile` (state + metadata) to/from a human-readable
CustomXMLPart, migrating earlier flat single-state documents forward. Inputs: a
`WorkbookProfileData`, `ProjectMetadata`, XML string, and workbook key. Output: an
`XDocument` or string; `TryDeserialize` returns a bool with the parsed profile/metadata.

### ProjectMetadataFactory.cs
Builds `ProjectMetadata` (environment/provenance) from OS/CLR/machine/Python info and the
venv's `pyvenv.cfg`, preserving the prior `CreatedUtc`. Inputs: project dir, workbook
name/path, and prior metadata. Output: a fresh `ProjectMetadata`.

### ProjectStructureValidator.cs
Fast, file-only check that a project's required directories exist (`.pyexcel-venv`,
`.pyexcel-kernel`, `userScripts`) without spawning Python. Inputs: a project dir. Output:
a `ProjectStructureCheck` record (Ok flag + missing components).

### ProjectReadiness.cs
Enum (NotEnabled / NeedsRepair / Ready) and the classifier that combines the enabled flag
with a structure check to gate ribbon controls. Inputs: an enabled bool and an optional
`ProjectStructureCheck`. Output: a `ProjectReadiness` value.

### WorkbookIdentityReconciler.cs
Pure decision logic for a workbook's stable identity (`WorkbookProfileData.ProjectId` /
`OriginPath`): compares the committed origin against the path the workbook is open at to
classify Unchanged / Moved (origin gone — same project relocated) / Copied (origin still on
disk — a Save-As copy that must become its own project). Inputs: the project id, origin
path, current path, and an origin-exists bool. Output: a `WorkbookIdentityAction`; the COM
sink applies the verdict.

### RibbonRangeParser.cs
Parses ribbon Input/Output text (semicolon-separated `A1:C10` or `name=A1:C10` bindings)
into an ordered list, and `Format()` reverses it. Inputs: a ribbon text string. Output: a
list of `RangeBinding`; throws `FormatException` on malformed syntax or duplicate names.

### ScriptDirectoryWatcher.cs
Watches a directory for `.py` changes and pushes updated script-name lists to a callback,
keeping the ribbon's Script dropdown live. Inputs: a directory path and a callback.
Output: a sorted `IReadOnlyList<string>` of filenames pushed on creation and each change.

### ErrorService.cs
Thread-safe per-workbook last-error registry (workbook-scoped + global slots) that fires
an event after mutation for ribbon invalidation. Inputs: a workbook key (null = global)
and a `KernelErrorRecord`. Output: the last `KernelErrorRecord?`; an `ErrorChanged` event.

### KernelErrorRecord.cs
Immutable error snapshot from the kernel or host (timestamp, source, code, Python type,
message, traceback, script path) with a clipboard formatter. Inputs: none (immutable
record). Output: the fields plus a `FormatForClipboard()` multi-line string.

### RunArchive.cs
On-disk archive of run results (manifest, inputs, output, error) with auto-pruning to a
max-run cap and timestamp-prefixed directory names for chronological sort. Inputs: a root
directory, max-run count, and a `RunArchiveEntry`. Output: the archived directory path;
`List()` returns `ArchivedRun`s newest-first.

### RunArchiveEntry.cs
Record packaging one run for archiving: timestamp, workbook key, script path, function,
source label, duration, input/output Arrow buffers, error, and status. Inputs: none
(immutable record). Output: those fields.

### RunArchiveContext.cs
Opt-in context passed to `PyRun` to enable archiving. Inputs: a `RunArchive`, a source
label, and the active workbook key. Output: a container for those outbound parameters.

### RunArchiveStatus.cs
Enum (Success / Error / Cancelled) distinguishing run outcomes so a replay knows whether
missing output means error or user interruption. Inputs/Output: an enum value.

### ArchivedRun.cs
Record for one row of `RunArchive.List()`, parsed from a manifest. Inputs: none
(immutable record). Output: `Directory`, `RunId`, `Timestamp`, `Status`, `ScriptPath`,
`WorkbookKey`, `Source`, `HasOutput`.

### PyExcelServices.cs
Static service locator for ribbon dependencies (`StateService`, `ErrorService`,
`RunArchive`, `WorkbookContext`, health registry), with safe unconfigured defaults wired
by AutoOpen. Inputs: none (property getters/setters). Output: process-wide service access;
a default `RunArchive` rooted under LocalApplicationData.

### IWorkbookContext.cs
Interface abstracting "which workbook is active" (COM in production, a fake in tests).
Inputs: none (query-only). Output: `CurrentWorkbookKey` and `CurrentWorkbookDirectory`.

### WorkbookKeys.cs
Shared workbook-key derivation: saved workbooks key on `FullName`, unsaved on
`unsaved:{SessionGuid}:{Name}`. Inputs: name, path, full name (COM properties). Output: a
stable workbook-key string.

### LegacyFormulaDecoder.cs
Reverses v1 PyExcel's string-literal formula encoding to recover original values from
Excel Name objects. Inputs: a `RefersTo` formula string. Output: the decoded string, or
null when not a valid v1 formula.

### LegacyStateConverter.cs
Converts v1-era per-sheet Name values into a v2 `WorkbookState` for migration, parsing
v1's delimited action format and legacy flags. Inputs: a `LegacyWorkbookState` (raw v1
strings) and a workbook key. Output: a `WorkbookState`; also exposes the `LegacyNames`
constants and a content check.

### IsExternalInit.cs
Internal compiler polyfill for `IsExternalInit` so records / init-only setters compile on
netstandard2.0 / net48. Inputs/Output: none (compile-time only).
