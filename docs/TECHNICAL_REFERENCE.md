# XCelReportEngine technical reference

Italian version: [TECHNICAL_REFERENCE-IT.md](TECHNICAL_REFERENCE-IT.md).

XCelReportEngine is an Open XML Excel backend intended to replace the subset of the NI Report Generation Toolkit used by the host applications without requiring Microsoft Office, OLE, or ActiveX.

## Targets

- .NET Framework 4.8 (`net48`) for LabVIEW 2026;
- .NET Framework 4.0 (`net40`) for the experimental LabVIEW 2013 integration;
- AnyCPU;
- Open XML SDK 3.5.1.

## Implemented capabilities

- A LabVIEW-friendly public API and `Int32` session handles.
- XLSX opening and XLTX-to-XLSX materialization.
- Worksheet listing and selection by name or zero-based index.
- Atomic save without closing the session and optional save on close.
- Idempotent locking and unlocking of all worksheets.
- Cell reads as strings and typed writes for strings, `Double`, and Boolean values.
- A1 addresses and zero-based coordinates, including Excel-boundary validation.
- Flat row-major string ranges for LabVIEW interoperability.
- Preservation of existing cell styles while writing.
- Deterministic rejection of formula reads; formulas are not evaluated.
- Embedded images addressed by A1 or zero-based coordinates.
- 96-DPI logical sizing, DrawingML one-cell behavior, and per-worksheet byte deduplication.
- Image scaling, explicit US/metric dimensions, picture selection, and grayscale.
- Alignment, wrapping, rotation, fills, and borders over cells or ranges.
- Stable errors containing a numeric code, name, operation, and inner exception.
- The same 25 automated backend tests on `net40` and `net48`.

## Public API

```text
ReportEngineApi.Create()
ReportEngineApi.ConvertCellIndexToAddress(rowIndex, columnIndex) -> string
ReportEngineApi.ConvertCellAddressToIndex(cellAddress) -> int[2]
OpenWorkbook(sourcePath, outputPath) -> sessionId
ValidateSession(sessionId)
GetWorksheetNames(sessionId) -> string[]
GetActiveWorksheetName(sessionId) -> string
SelectWorksheetByName(sessionId, worksheetName)
SelectWorksheetByIndex(sessionId, worksheetIndex)
LockAllWorksheets(sessionId, password)
UnlockAllWorksheets(sessionId)
ReadCellStringByAddress(sessionId, cellAddress) -> string
ReadCellStringByIndex(sessionId, rowIndex, columnIndex) -> string
WriteCellStringByAddress / WriteCellStringByIndex
WriteCellDoubleByAddress / WriteCellDoubleByIndex
WriteCellBooleanByAddress / WriteCellBooleanByIndex
ReadStringRangeByIndex(...) -> string[]
WriteStringRangeByIndex(..., values)
SetCellAlignmentByAddress / SetCellAlignmentByIndex
SetCellColorAndBorderByAddress / SetCellColorAndBorderByIndex
AppendImage(sessionId, imagePath, rowIndex, columnIndex, cellAddress, alignment, caption) -> pictureIndex
FormatImage(sessionId, measurementSystem, scaleFactor, pictureIndex, height, width, colorType)
SaveWorkbook(sessionId)
CloseWorkbook(sessionId, saveChanges)
```

The signatures deliberately avoid generics, delegates, overloads, and public Open XML types.

## Cell and range conventions

Row and column indexes are zero-based: `(0, 0)` is `A1`. A1 addresses are case-insensitive and accept absolute notation such as `$A$1`.

Reads return an empty string for an absent or empty cell, `TRUE`/`FALSE` for Boolean cells, and the stored invariant-culture representation for numbers. Formula cells raise `FormulaReadNotSupported` because this component does not contain an Excel calculation engine.

String ranges cross the .NET/LabVIEW boundary as one-dimensional row-major arrays. The element at `(row, column)` is stored at `row * columnCount + column`. LabVIEW wrappers use `Reshape Array` to convert to and from 2D arrays. A range read fails as a whole when at least one included cell contains a formula.

## Images

For `AppendImage`, a nonempty `cellAddress` takes precedence over numeric coordinates. `alignment` accepts NI values 0–8 but does not modify the Excel drawing, matching the observed NI backend. `caption` is accepted but ignored. The returned picture index is zero-based. Image bytes are embedded without transcoding and no external relationship to the source path is created.

`FormatImage` uses `pictureIndex = -1` for the last image. Measurement system `0` is US/inches and `1` is metric/centimetres. Negative explicit dimensions preserve the corresponding current dimension. Color type `0` is automatic and `2` is grayscale; unsupported types produce a stable error.

## Cell formatting values

- Horizontal alignment: 0 General, 1 Left, 2 Center, 3 Right, 4 Fill, 5 Justify, 6 CenterContinuous, 7 Distributed.
- Vertical alignment: 0 Bottom, 1 Center, 2 Top, 3 Justify, 4 Distributed.
- Text rotation: Open XML value 0–180.
- Colors: `Int32` RGB values in `0xRRGGBB` format.
- Border style: 0 None, 1 Thin, 2 Medium, 3 Thick, 4 Double, 5 Dashed, 6 Dotted.
- Border edges: bit mask 1 Left, 2 Right, 4 Top, 8 Bottom; 15 selects all edges.

An empty end address means only the start cell. `applyFill` and `applyBorder` allow the two properties to be changed independently. Other style components are preserved.

## Build and verification

```powershell
dotnet restore .net/XCelReportEngine.sln --locked-mode -m:1
dotnet build .net/XCelReportEngine.sln -c Release --no-restore -m:1
dotnet test .net/tests/XCelReportEngine.Tests/XCelReportEngine.Tests.csproj -c Release -f net48 --no-build --no-restore -m:1
.net/tests/XCelReportEngine.Tests/bin/Release/net40/XCelReportEngine.Tests.exe
```

Current verification: both backend targets build with zero warnings and errors; 25 tests pass on each target; generated workbooks are schema-validated. LabVIEW 2026 has manual smoke-test coverage. LabVIEW 2013 wrappers have not been relinked, compiled, packaged, or run in LabVIEW 2013.

Release artifacts are staged under `LV/DotNet` for `net48` and `LV2013/DotNet` for `net40`. Keep `XCelReportEngine.dll`, `DocumentFormat.OpenXml.dll`, and `DocumentFormat.OpenXml.Framework.dll` together.

## LabVIEW API

`XCel Report Engine.lvlib` exposes wrappers for report lifecycle, worksheets, cells, ranges, formatting, and images. `Select Worksheet.vi` and `Write Cell.vi` are polymorphic. `Coordinate2Address.vi` and `Address2Coordinate.vi` are session-independent helpers; the latter returns zero-based row and column coordinates in that order.
