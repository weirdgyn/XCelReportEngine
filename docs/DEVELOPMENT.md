# Development and maintenance

Italian version: [DEVELOPMENT-IT.md](DEVELOPMENT-IT.md).

## Prerequisites

- Windows;
- LabVIEW 2026 for the current LabVIEW integration;
- .NET SDK 10 with the .NET Framework 4.8 targeting pack;
- NuGet access or a populated local package cache;
- the .NET Framework 4.0 reference assemblies for the legacy backend.

LabVIEW 2013 cannot currently be installed and validated alongside the LabVIEW 2026 development environment used by this project. The `LV2013` sources therefore remain experimental even though the shared `net40` backend is built and tested.

## Reproducible build

```powershell
dotnet restore .net/XCelReportEngine.sln --locked-mode -m:1
dotnet build .net/XCelReportEngine.sln -c Release --no-restore -m:1
dotnet test .net/tests/XCelReportEngine.Tests/XCelReportEngine.Tests.csproj -c Release -f net48 --no-build --no-restore -m:1
.net/tests/XCelReportEngine.Tests/bin/Release/net40/XCelReportEngine.Tests.exe
```

The serial build is intentional. In the current environment, a parallel solution build with .NET SDK 10 may terminate without useful diagnostics.

The same 25 source-level test cases run against `net40` and `net48`. The `net40` executable contains a minimal reflection-based test runner because supported versions of `Microsoft.NET.Test.Sdk` and xUnit's Visual Studio adapter require a newer framework. A nonzero exit code indicates failure and is suitable for CI.

The `packages.lock.json` files are versioned. Update dependencies deliberately with `dotnet restore .net/XCelReportEngine.sln --force-evaluate -m:1`.

## .NET and LabVIEW dependency

Release builds stage the assemblies automatically:

```text
net48 -> LV/DotNet/
net40 -> LV2013/DotNet/
```

LabVIEW keeps .NET assemblies loaded until the process exits. Close LabVIEW completely before rebuilding into the Release directory, or use a temporary output directory for .NET-only verification.

The public signatures of `ReportEngineApi` are the interoperability contract. Changing a class, namespace, assembly, signature, or parameter type requires manually updating the LabVIEW Constructor/Invoke Nodes and the `Report Ref.ctl` typedef.

## LabVIEW file rules

- Rename or move `.vi`, `.ctl`, `.lvlib`, and `.lvproj` files only from the LabVIEW IDE.
- Save every caller after changing a typedef.
- Do not commit `.lvlps`, `.aliases`, test output, or user state.
- Verify the LVLIB and relevant LabVIEW tests before committing binary LabVIEW changes.
- List the modified VIs in the commit because Git cannot provide useful textual block-diagram diffs.

## LabVIEW automation status

The repository has no automated LabVIEW build or test pipeline. Reliable automation requires a licensed, installed LabVIEW version plus NI-supported command-line or VI Server orchestration. An MCP server would only expose that orchestration; it would not remove the requirement for the matching LabVIEW runtime/development environment.

LabVIEW 2013 integration must therefore be validated later on a dedicated machine or VM containing LabVIEW 2013. Until then:

- `net40` backend compilation and functional tests are automated;
- `LV2013` relinking, compilation, packaging, and smoke tests are not verified;
- CI must not present the LabVIEW 2013 wrapper as supported.

## Git and pull requests

Before opening a pull request, build and test both .NET targets, run relevant LabVIEW 2026 smoke tests for binary changes, inspect `git status --short`, update the public documentation when the contract changes, and avoid mixing unrelated .NET and LabVIEW work.

GitHub Actions builds both backend targets, runs the `net48` suite with `dotnet test`, and runs the `net40` test executable directly. LabVIEW tests remain manual.

## Unpublished local assets

The following paths are excluded from Git until their contents and Office metadata are sanitized:

```text
LV/Test/
LV/Templates/
LV/Pre-Test/
docs/REPORT_INVENTORY.md
docs/REPORT_LOCKER_ANALYSIS.md
docs/TEMPLATE_ANALYSIS.md
docs/data/workbook_analysis.json
```
