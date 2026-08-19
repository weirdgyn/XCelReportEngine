# Changelog

Italian version: [CHANGELOG-IT.md](CHANGELOG-IT.md).

## Unreleased

- Added the canonical product roadmap and repository-wide project instructions.
- Added the initial NI Toolkit parity inventory and current implementation mapping.
- Added the NI-to-XCelReportEngine VI migration map and adopted one-to-one compatibility facades as a roadmap requirement.
- Added the detailed LabVIEW VI naming, composition, reusable subVI, backend-boundary, and LabVIEW 2013 mirroring plan.
- Defined ActiveX, COM, Office automation, and cell calculation as permanent compatibility-scope exclusions.
- Added automated execution of the same 25 backend tests on .NET Framework 4.0 and 4.8.
- Made English the default documentation language and retained Italian documents with an `-IT` suffix.
- Documented the current LabVIEW 2013 build and test limitations.
- Added experimental .NET Framework 4.0/LabVIEW 2013 support.
- Added icons, VIPM packaging configuration, and Antidoc configuration.
- Renamed the component to `XCelReportEngine` and separated `.net`, `LV`, `docs`, and `tools`.
- Added Git configuration for LabVIEW binaries, NuGet lock files, and Windows CI.
- Isolated Excel address parsing and optimized range reads.
- Added coordinate/A1 helpers, LabVIEW wrappers, and polymorphic worksheet/cell VIs.
- Temporarily excluded unsanitized LabVIEW tests, templates, and company-derived analyses.
- Implemented and verified workbook lifecycle, worksheets, cells, ranges, protection, images, alignment, fills, and borders.
