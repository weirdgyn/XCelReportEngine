# XCelReportEngine

An Open XML-based Excel report engine for LabVIEW applications, with no dependency on Microsoft Office, OLE, or ActiveX.

Italian documentation is also available in [README-IT.md](README-IT.md).

## Repository layout

```text
.net/   .NET Framework 4.8/4.0 backend and automated tests
LV/     LabVIEW library, wrappers, controls, and tests
LV2013/ LabVIEW 2013-compatible library, wrappers, and controls
docs/   architecture, development notes, and characterization results
tools/  analysis and test-asset generation utilities
```

The local `LV/Test`, `LV/Templates`, and `LV/Pre-Test` directories are temporarily excluded from the public repository until their assets have been sanitized.

## .NET build and tests

```powershell
dotnet restore .net/XCelReportEngine.sln --locked-mode -m:1
dotnet build .net/XCelReportEngine.sln -c Release --no-restore -m:1
dotnet test .net/tests/XCelReportEngine.Tests/XCelReportEngine.Tests.csproj -c Release -f net48 --no-build --no-restore -m:1
.net/tests/XCelReportEngine.Tests/bin/Release/net40/XCelReportEngine.Tests.exe
```

The same 25 test cases run on both targets. `net48` uses the Visual Studio test adapter; `net40` uses a small self-contained runner because current `Microsoft.NET.Test.Sdk` and xUnit adapters no longer support .NET Framework 4.0.

See [the technical reference](docs/TECHNICAL_REFERENCE.md) for the current feature contract and [the development guide](docs/DEVELOPMENT.md) for build, Git, and LabVIEW instructions.

## License

XCelReportEngine is distributed under the BSD 3-Clause License (`BSD-3-Clause`).

Copyright (c) 2026, Michele Santucci. See [LICENSE.txt](LICENSE.txt) for the full license text.

Bundled third-party components remain subject to their respective licenses.
