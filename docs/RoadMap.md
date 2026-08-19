# XCelReportEngine roadmap

## Vision

XCelReportEngine will provide a deterministic, headless, Open XML-based replacement for the eligible general, Microsoft Office-common, and Excel-specific VIs in the NI LabVIEW Report Generation Toolkit.

The engine must generate and modify portable Excel workbooks without requiring Microsoft Excel, Microsoft Office, OLE, COM, ActiveX, or Office Interop. Compatibility is measured against the documented behavior of the NI public VIs, subject to the exclusions below.

This roadmap is evolutionary. Features should be delivered in independently testable increments rather than through a single compatibility rewrite.

## Canonical reference

The behavioral inventory is based on the [NI LabVIEW Report Generation Toolkit Programming Reference Manual](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html), using its public palettes, documented VIs, polymorphic instances, inputs, outputs, defaults, and error behavior.

The project maintains its living compatibility inventory in [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md). The matrix must be complete before publishing a numeric coverage claim. The documentation snapshot date or NI manual revision used for the matrix must be recorded because palette contents can vary by LabVIEW version, target, operating system, and installed licenses.

## Active scope

The target is complete coverage of every technically eligible operation in these groups:

- general Report Generation VIs relevant to Excel;
- Advanced Report Generation VIs relevant to Excel;
- Report Layout VIs relevant to Excel;
- Microsoft Office-common report operations used by the Excel backend;
- Excel Easy Title, Text, Table, and Graph VIs;
- Excel General VIs;
- Excel Format VIs;
- Excel Graphs and Pictures VIs;
- Excel Advanced VIs.

Polymorphic instances count as separate compatibility obligations when their data types, connector panes, defaults, or behavior differ.

## Permanent exclusions

The following functionality is deliberately outside the product and outside the compatibility percentage denominator:

### Office automation

- ActiveX and COM references;
- OLE automation;
- Microsoft Office Interop;
- access to live Excel `Application`, `Workbook`, `Worksheet`, `Range`, `Chart`, or similar automation objects;
- operations that require launching or controlling Microsoft Excel;
- VBA macro execution;
- Excel-driven printing, preview, rendering, PDF export, refresh, or UI automation.

ActiveX independence is a founding requirement of the project. These capabilities must not be reintroduced through optional adapters in the core product.

### Calculation

- formula evaluation;
- workbook or worksheet recalculation;
- dependency-graph calculation;
- functions whose result requires updated formula values from Excel.

Existing formulas and cached values may be preserved. A future feature may write formula text only if its non-calculating contract is explicit and cannot be mistaken for formula evaluation.

### Other report families

- HTML-only report operations;
- VI Documentation operations;
- Word-specific operations.

DOCX/WordprocessingML support remains a possible future, separately scoped product extension. The current roadmap and product name remain Excel-focused.

## Eligible file-model behavior

The following remain eligible even when a related runtime action is excluded:

- page setup, margins, orientation, print areas, print titles, and header/footer metadata;
- native charts stored in the workbook package;
- preservation of formulas and cached values;
- preservation of embedded VBA projects in future `.xlsm`/`.xltm` support, without execution;
- named ranges, hyperlinks, tables, drawings, images, styles, and protection;
- workbook metadata interpreted by Excel when the file is opened later.

The distinction is whether XCelReportEngine edits the Open XML file model or asks Excel to execute an operation.

## Compatibility states

Every inventory entry must use one of these states:

- **Matched** — documented behavior, defaults, outputs, and relevant errors are covered.
- **Partially matched** — a useful subset exists, with explicit missing behavior.
- **Planned** — eligible but not implemented.
- **Requires characterization** — documentation is insufficient to define deterministic behavior.
- **Intentional difference** — implemented behavior deliberately corrects or avoids an NI defect.
- **Excluded: ActiveX/Office automation** — requires a live Office automation model.
- **Excluded: calculation** — requires formula calculation or recalculation.
- **Excluded: other report family** — HTML, VI Documentation, or Word-specific.
- **Obsolete/internal** — not a public compatibility target.

Excluded entries remain in the matrix for traceability but do not reduce the eligible coverage percentage.

## Current baseline

The repository currently provides:

- XLSX opening and XLTX-to-XLSX materialization;
- session-based workbook lifecycle and atomic publishing;
- worksheet listing and selection;
- zero-based coordinate and A1-address conversion;
- string, double, and Boolean cell writes;
- string cell and range reads, with explicit formula-read rejection;
- flat row-major string-table transfer for LabVIEW;
- worksheet lock and unlock;
- cell alignment, text wrapping, rotation, fill, and border formatting;
- embedded image insertion, sizing, selection, grayscale, and media deduplication;
- stable backend error codes;
- LabVIEW 2026 wrappers;
- exported, experimental LabVIEW 2013 wrappers;
- shared automated backend tests on .NET Framework 4.0 and 4.8.

See [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md) for the current public contract.

## Development plan

### Phase 0 — Inventory and measurable parity

1. Extract every public eligible NI VI and polymorphic instance.
2. Complete [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md) with every source URL, category, signature summary, status, backend mapping, wrapper mapping, test, and exclusion.
3. Record the NI documentation revision used as the baseline.
4. Define how generic Report Generation VIs map to Excel behavior.
5. Identify VIs that require empirical characterization rather than documentation-only implementation.
6. Define the first eligible compatibility denominator; publish no percentage before this phase is complete.

### Phase 1 — Reliability and scalability foundation

1. Deduplicate fills, borders, fonts, and cell formats to prevent workbook style growth.
2. Add stress tests for large formatted ranges and repeated operations.
3. Strengthen preservation tests for charts, drawings, relationships, named ranges, formulas, external links, and unknown package parts.
4. Add invalid-package, locked-file, read-only, save-failure, Unicode, culture, and multi-session tests.
5. Review session concurrency and close/save behavior.
6. Establish reusable semantic Open XML comparison utilities for NI golden files.

### Phase 2 — General lifecycle and layout parity

Implement eligible general and layout behavior before duplicating it in Easy VIs:

- create blank reports and reports from templates;
- save, save-as semantics, clear, and close/dispose behavior;
- report type and settings queries where meaningful without Office;
- new report line/page equivalents where they map to Excel layout;
- margins, orientation, paper size, scaling, page breaks, headers, footers, and page numbering metadata;
- append text, numeric and string tables, and images through Excel-compatible generic wrappers;
- deterministic errors for unsupported report types and excluded runtime operations.

### Phase 3 — Excel General parity

- create, copy, rename, select, reorder, hide, show, and delete worksheets;
- retrieve worksheet names, indexes, dimensions, and active worksheet information;
- insert, delete, clear, and search cells or ranges;
- determine the used range and last row/column;
- manage rows and columns, including width, height, visibility, insertion, and deletion;
- merge and unmerge cells;
- freeze and split panes where representable;
- manage named ranges and hyperlinks;
- support bulk typed reads and writes.

### Phase 4 — Excel Format parity

- font name, size, color, bold, italic, underline, and strike-through;
- number formats, dates, times, percentages, currencies, and custom formats;
- horizontal and vertical alignment, wrapping, shrink-to-fit, indentation, and rotation;
- fills and complete border models;
- row height and column width formatting;
- style reuse and deduplication;
- cell protection flags;
- documented auto-format behavior where it has a deterministic Open XML representation.

Exact renderer-dependent AutoFit behavior is not promised. If a deterministic approximation is provided, it must be classified and documented as an intentional difference.

### Phase 5 — Excel Easy VIs

Build Easy VIs from the lower-level services so their complete contracts are covered:

- Excel Easy Title;
- Excel Easy Text;
- Excel Easy Table string and numeric instances;
- row and column headers;
- named-range targeting;
- font input/output behavior;
- iteration semantics;
- next-cell position outputs;
- eligible auto-format options;
- Excel Easy Graph after the native chart foundation exists.

### Phase 6 — Graphs and pictures

- preserve and enumerate existing drawings;
- add native Excel charts with supported data and titles;
- define the supported NI graph-type mapping and record unsupported or intentionally different types;
- position, resize, crop, recolor, and remove pictures;
- support additional documented picture color modes where Open XML provides an equivalent;
- handle images in headers and footers;
- preserve insertion order, identifiers, relationships, and z-order;
- compare generated drawing packages with characterized NI outputs.

### Phase 7 — Excel Advanced parity

Classify the Advanced palette before implementation. Prioritize operations with deterministic Open XML representations, including:

- workbook and worksheet protection options;
- validation, defined names, metadata, and advanced page settings;
- tables and advanced drawing/package operations;
- macro-enabled workbook and template preservation (`.xlsm` and `.xltm`) without VBA execution.

Mark ActiveX references, macro execution, calculation, live refresh, printing, and renderer-dependent operations as permanent exclusions.

### Phase 8 — LabVIEW API parity and packaging

- provide LabVIEW 2026 wrappers for every matched operation;
- reproduce relevant NI names, defaults, dataflow, error-cluster behavior, and polymorphic instances without copying NI implementation assets;
- organize palettes by General, Easy, Format, Graphs and Pictures, and Advanced;
- maintain simple .NET signatures suited to Constructor and Invoke Nodes;
- export corresponding LabVIEW 2013 wrappers where feasible;
- validate VIPM packaging, deployment dependencies, examples, icons, and help integration;
- provide migration examples from NI toolkit VIs.

### Phase 9 — Compatibility closure

1. Run every eligible matrix row against a sanitized NI-generated golden fixture or a documented semantic contract.
2. Validate generated packages with the Open XML validator.
3. Verify representative outputs in supported desktop spreadsheet applications without using them as runtime dependencies.
4. Resolve or document every remaining partial match and intentional difference.
5. Publish eligible coverage metrics by category and overall.
6. Declare 100% only when every eligible row is matched or carries an approved intentional-difference classification.

## Testing strategy

Every new backend behavior should include:

- successful round-trip tests;
- invalid argument and boundary tests;
- `net40` and `net48` execution;
- schema validation;
- preservation checks for unrelated workbook parts;
- idempotency tests where applicable;
- large-range or repeated-operation tests for performance-sensitive features;
- semantic comparison with NI-generated fixtures when NI behavior is the compatibility target.

Tests should compare meaning rather than unstable package details such as relationship IDs, object IDs, random creation identifiers, ZIP entry ordering, or irrelevant serialization differences.

## Definition of done

An eligible roadmap item is complete when:

1. its NI reference and compatibility classification are recorded;
2. its deterministic behavior and intentional differences are documented;
3. its backend implementation supports both target frameworks;
4. stable validation and errors are implemented;
5. automated tests pass on `net40` and `net48`;
6. Open XML schema and preservation checks pass;
7. the LabVIEW 2026 wrapper and required polymorphic instances exist;
8. the LabVIEW 2013 counterpart is exported where feasible and its unverified status is honest;
9. the technical reference, changelog, and compatibility matrix are updated.

## Future Word extension

Word/DOCX support is not part of the active roadmap. If approved later, it should use a separate WordprocessingML backend and should not require renaming or breaking the existing `XCelReportEngine` assembly and LabVIEW package. A shared lifecycle/media core may be extracted only when justified by implemented behavior rather than speculative abstraction.

## Related project documents

- [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md) — implemented API and behavior.
- [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md) — canonical NI VI inventory and coverage status.
- [DEVELOPMENT.md](DEVELOPMENT.md) — build, test, CI, LabVIEW, and Git procedures.
- [ARCHITECTURE_DECISIONS.md](ARCHITECTURE_DECISIONS.md) — accepted architectural decisions.
- [IMAGE_TEST_RESULTS.md](IMAGE_TEST_RESULTS.md) — NI image behavior characterization.
- [CHANGELOG.md](CHANGELOG.md) — notable project changes.
