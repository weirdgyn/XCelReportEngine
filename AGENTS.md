# Project agent instructions

These instructions apply to the entire repository. Update this file whenever a new durable project rule is agreed upon.

## Project purpose

XCelReportEngine is a headless Excel report engine for LabVIEW based on Open XML. Its purpose is to replace the eligible general, Microsoft Office-common, and Excel-specific functionality of the NI LabVIEW Report Generation Toolkit without requiring Microsoft Office, OLE, COM, or ActiveX.

The canonical product roadmap is [docs/RoadMap.md](docs/RoadMap.md), and the canonical NI compatibility inventory is [docs/NI_TOOLKIT_PARITY_MATRIX.md](docs/NI_TOOLKIT_PARITY_MATRIX.md). The current public contract is documented in [docs/TECHNICAL_REFERENCE.md](docs/TECHNICAL_REFERENCE.md), architectural decisions in [docs/ARCHITECTURE_DECISIONS.md](docs/ARCHITECTURE_DECISIONS.md), and development procedures in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Non-negotiable architectural rules

- Do not introduce ActiveX, COM, OLE, Microsoft Office Interop, or any dependency on an installed Microsoft Office application.
- Do not implement cell-formula calculation or promise recalculation. Existing formulas and cached values may be preserved, but formula evaluation is outside the product scope.
- Do not execute VBA macros. Macro-enabled Open XML files may eventually be preserved without executing their VBA projects.
- Separate file-model configuration from runtime execution. Page setup and print metadata are eligible; invoking Excel to print, render, refresh, or export is not.
- Prefer deterministic Open XML transformations that work headlessly and produce portable files.
- Preserve existing workbook content, styles, relationships, and unknown package parts unless an operation explicitly changes them.
- Public .NET APIs exposed to LabVIEW must use simple interoperable types. Avoid generics, delegates, overload-dependent designs, and public Open XML SDK types.
- Treat public `ReportEngineApi` signatures, assembly identity, namespaces, and the LabVIEW `Report Ref.ctl` typedef as a compatibility contract.

## Product scope

The active roadmap targets complete coverage of eligible:

- general Report Generation Toolkit operations;
- Microsoft Office-common operations relevant to Excel reports;
- Excel Easy VIs;
- Excel General VIs;
- Excel Format VIs;
- Excel Graphs and Pictures VIs;
- Excel Advanced VIs.

The following are outside the compatibility denominator:

- ActiveX/COM/Office automation operations;
- formula calculation and recalculation;
- operations that require an Excel rendering engine or live Excel process;
- HTML-only report operations;
- VI Documentation operations;
- Word-specific operations.

WordprocessingML/DOCX support is a possible future extension, not part of the current roadmap. Do not rename `XCelReportEngine` or distort the Excel architecture in anticipation of Word support. If Word is pursued later, prefer a separate backend and an extractable shared core while retaining the existing Excel assembly and package identity.

## Compatibility and traceability

- Use the NI LabVIEW Report Generation Toolkit Programming Reference Manual as the behavioral inventory source.
- Record every public NI VI and polymorphic instance in [docs/NI_TOOLKIT_PARITY_MATRIX.md](docs/NI_TOOLKIT_PARITY_MATRIX.md).
- Classify each operation as matched, partially matched, planned, requiring characterization, intentionally different, or excluded with a reason.
- Do not claim percentage coverage until the eligible inventory and denominator are versioned.
- Prefer behavioral clean-room characterization and semantic Open XML comparison. Do not copy or redistribute NI implementation assets.
- Intentional improvements over observed NI bugs must be documented and tested explicitly.

## Target frameworks and LabVIEW versions

- The shared backend targets .NET Framework 4.8 (`net48`) for LabVIEW 2026 and .NET Framework 4.0 (`net40`) for the experimental LabVIEW 2013 integration.
- Backend changes must build for both target frameworks unless the roadmap explicitly changes this requirement.
- The same source-level backend tests should run against both targets whenever technically possible.
- LabVIEW 2013 wrapper compatibility remains experimental until relinking, compilation, packaging, and smoke testing can be performed in a dedicated LabVIEW 2013 environment.
- Do not present a successful `net40` backend build as proof that the LabVIEW 2013 wrapper layer is validated.

## Definition of done for a roadmap function

An eligible function is complete only when the applicable deliverables exist:

1. characterized NI behavior or a documented Open XML contract;
2. backend implementation for `net40` and `net48`;
3. stable errors and argument validation;
4. automated tests on both backend targets;
5. Open XML schema validation and preservation tests;
6. a LabVIEW 2026 public wrapper and relevant polymorphic instances;
7. a corresponding exported LabVIEW 2013 wrapper where feasible, clearly marked unverified until tested;
8. English documentation and an updated compatibility-matrix entry;
9. documentation of every intentional deviation from NI behavior.

## Development rules

- Keep English as the default documentation language. Preserve Italian translations with an `-IT` suffix when translations are maintained.
- Update [docs/CHANGELOG.md](docs/CHANGELOG.md) for notable functional or compatibility changes.
- Add tests before or with behavior changes, including negative cases and package-preservation checks.
- Avoid per-cell Open XML scans and avoid creating duplicate styles for repeated formatting operations.
- Use bulk APIs for tables and ranges to reduce .NET/LabVIEW boundary calls.
- Keep generated assemblies, test workbooks, LabVIEW user state, and unsanitized business templates out of Git.
- Rename or move LabVIEW binary files only from the LabVIEW IDE and save all callers after typedef changes.
- Do not modify LabVIEW binary files without describing which VIs changed and how they were verified.

## Required verification

For backend changes, use the commands documented in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md). At minimum:

- restore in locked mode;
- build the solution in Release configuration with serial MSBuild;
- run the `net48` test suite;
- run the `net40` test executable;
- inspect `git status --short` for generated artifacts;
- run relevant LabVIEW 2026 smoke tests when LabVIEW files or public interop signatures change.
