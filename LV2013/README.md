# XCelReportEngine for LabVIEW 2013

This directory contains the XCelReportEngine LabVIEW sources exported for LabVIEW 2013 compatibility.

The .NET Framework 4.0 backend is built from the shared .NET sources and staged automatically in `LV2013/DotNet` during a Release build. Generated .NET assemblies are build artifacts and are not tracked by Git.

## Compatibility status

**Experimental — not yet verified with LabVIEW 2013.**

The LabVIEW sources have been exported to the LabVIEW 2013 file format, but they have not yet been opened, relinked, compiled, packaged, or functionally tested with a LabVIEW 2013 installation. The project file, palettes, and VIPM package configuration must also be recreated or validated using LabVIEW 2013.

The exported LabVIEW files currently retain the original absolute development-time reference to the modern assembly under `LV/DotNet`. Before testing or packaging the legacy variant, every .NET constructor, method, and .NET refnum must be relinked and saved against the .NET Framework 4.0 assembly staged in `LV2013/DotNet`.

Do not use this directory as evidence of supported LabVIEW 2013 compatibility until the required integration and smoke tests have been completed.
