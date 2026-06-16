# PyExcel.Bridge.Tests

## Macro
The xUnit suite covering the cross-platform C# slice — every pure type in
`PyExcel.Bridge`, `PyExcel.Kernel.Client`, `PyExcel.Excel`, `PyExcel.State`,
`PyExcel.Common`, `PyExcel.Setup`, and the validators in `PyExcel.Forms`. It runs
on Linux CI (`dotnet test`). A few tests spawn a real Python kernel
(`KernelClientTests`, `KernelSupervisorTests`) and need a Python on PATH.

## Files
There are ~57 test classes, each mirroring a single production type — the file name is
`<Type>Tests.cs`. Inputs are the type under test plus crafted fixtures; output is xUnit
assertions (no artifacts). Grouped by area rather than listed individually, because the
map value is "which area is tested where":

- **Wire / framing**: `FramingTests`, `FrameTransportTests`, `CrossLanguageVectorsTests`,
  `ChartWireTests` — frame layout, bounds, and Python↔C# byte-for-byte compatibility.
- **Arrow marshalling**: `ArrowMarshalTests`.
- **Live kernel** (spawns a real subprocess): `KernelClientTests`, `KernelSupervisorTests`
  — handshake, run, ping, cancel, and child-kill behaviour.
- **Charts**: `ChartColorTests`, `ChartSpecParserTests`, `ChartTraceDataTests`.
- **CSV/TSV**: `CsvParserTests`, `CsvWriterTests`, `CsvCellFormatterTests`,
  `CsvCellTypeInferenceTests`.
- **Import/Export/Paste planning + validation**: `ImportPlannerTests`,
  `ExportPlannerTests`, `ExportSettingsTests`, `ExportSettingsPlannerTests`,
  `EditIoValidatorTests`, `PastePlannerTests`, `PastePreflightTests`.
- **State, codecs, persistence**: `WorkbookStateCodecTests`, `WorkbookProfileCodecTests`,
  `ProjectProfileTests`, `StateServiceTests`, `PerSheetStateTests`,
  `WorkbookStatePersistenceTests`, `WorkbookKeysTests`, `LegacyStateConverterTests`,
  `LegacyFormulaDecoderTests`, `ErrorServiceTests`, `RunArchiveTests`,
  `PyExcelServicesTests`, `ScriptDirectoryWatcherTests`.
- **Ribbon ranges / orientation / sheet / progress**: `RibbonRangeParserTests`,
  `RibbonRangeFormatTests`, `RangeAddressValidatorTests`, `OrientationResolverTests`,
  `SheetSelectionTests`, `SheetPickerValidatorTests`, `ProgressModelTests`.
- **Run orchestration**: `PyRunTests`.
- **Setup pipeline**: `SetupServiceTests`, `SetupReportTests`, `SystemPythonProbeTests`,
  `VenvProvisionerTests`, `PipRunnerTests`, `DependencyVerifierTests`,
  `KernelResourceExtractorTests`, `ProjectScaffolderTests`, `ProjectPathResolverTests`,
  `ProjectStructureValidatorTests`, `ProjectReadinessTests`, `ProjectDirectoryTests`,
  `PythonResolverTests`, `ProcessRunnerTests`.
- **Forms validators**: `EditActionValidatorTests`, `KwargsTextTests`,
  `ScriptScaffoldTests`.

## Subdirectories
None.
