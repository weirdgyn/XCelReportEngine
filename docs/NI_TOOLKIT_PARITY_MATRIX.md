# NI Report Generation Toolkit parity matrix

## Purpose

This document is the canonical compatibility inventory for XCelReportEngine. It maps public NI LabVIEW Report Generation Toolkit VIs and polymorphic instances to the current backend API, LabVIEW wrappers, tests, exclusions, and remaining work.

Direct user migration mappings, connector compatibility, icons, and temporary multi-VI sequences are maintained separately in [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md).

The inventory source is the [NI LabVIEW Report Generation Toolkit Programming Reference Manual](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html), currently using the documentation revision updated on 2026-02-03.

This matrix is intentionally a living document. The initial entries below cover the public VIs already identified and the XCelReportEngine operations already implemented. Child inventories for the dynamically rendered NI subpalettes must be expanded from their individual NI reference pages before a numeric compatibility percentage is published.

## Status legend

| Status | Meaning | Included in eligible denominator |
|---|---|---:|
| Matched | Relevant documented behavior, defaults, outputs, and errors are covered | Yes |
| Partially matched | A useful subset exists and missing behavior is recorded | Yes |
| Planned | Eligible Open XML behavior is not implemented | Yes |
| Requires characterization | Documentation alone does not define deterministic behavior | Pending decision |
| Intentional difference | XCelReportEngine deliberately differs from NI behavior | Yes |
| Excluded: ActiveX/Office automation | Requires a live Office automation model | No |
| Excluded: calculation | Requires formula calculation or recalculation | No |
| Excluded: other report family | HTML, VI Documentation, or Word-specific | No |
| Obsolete/internal | Not a public compatibility target | No |

## Coverage policy

Do not calculate or publish a compatibility percentage until:

1. every public child VI and polymorphic instance in the eligible NI palettes has an inventory row;
2. every row has an eligibility classification;
3. the documentation revision is fixed;
4. ambiguous functions have been characterized or explicitly deferred.

Excluded rows remain visible for traceability but do not reduce eligible coverage.

## General and Microsoft Office-common operations

This section will be expanded from the Advanced Report Generation and Report Layout palettes. The rows below represent operations already exposed by XCelReportEngine or explicitly classified by project policy.

| NI operation or family | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| New/Create Report from Excel template | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | Eligible | Partially matched | `OpenWorkbook` | `New Report.vi` | Opens XLSX and materializes XLTX; blank-report creation and full NI template/default semantics remain |
| Save Report to File | Same parent reference; individual row pending | Eligible | Partially matched | `SaveWorkbook`, `CloseWorkbook(saveChanges)` | `Save Report.vi` | Atomic XLSX publish is covered; save-as variants and all documented defaults require inventory confirmation |
| Dispose/Close Report | Same parent reference; individual row pending | Eligible | Partially matched | `CloseWorkbook`, `Dispose` | `Close Report.vi` | Session close behavior exists; NI connector/default parity requires confirmation |
| Append Image to Report | Individual NI page to inventory | Eligible | Partially matched | `AppendImage`, `FormatImage` | `Append Image.vi`, `Format Image.vi` | Excel image insertion is covered; generic wrapper signature and all picture modes remain partial |
| Append Text/Table to Excel report | Individual NI pages to inventory | Eligible | Partially matched | cell/range write APIs | `Write Cell*.vi`, `Write String Table.vi` | Basic Excel equivalents exist; generic font, headers, layout, and polymorphic instances remain |
| Report page setup and layout metadata | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | Eligible | Planned | None | None | Margins, orientation, page breaks, paper, headers, footers, and numbering are planned |
| Print/preview through Microsoft Office | Individual NI pages to inventory | Excluded | Excluded: ActiveX/Office automation | None | None | Runtime rendering/printing is a permanent exclusion |
| Export through Excel/Office renderer | Individual NI pages to inventory | Excluded | Excluded: ActiveX/Office automation | None | None | PDF/export through a live Office application is excluded |
| Get Office/ActiveX references | Individual NI pages to inventory | Excluded | Excluded: ActiveX/Office automation | None | None | ActiveX independence is a founding requirement |
| Execute VBA macro | Individual NI pages to inventory | Excluded | Excluded: ActiveX/Office automation | None | None | Macro preservation may be eligible later; execution is excluded |
| Calculate/recalculate formulas | Individual NI pages to inventory | Excluded | Excluded: calculation | None | None | Formula evaluation and recalculation are excluded |
| HTML report operations | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | Excluded | Excluded: other report family | None | None | Not an Excel Open XML compatibility target |
| VI Documentation operations | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | Excluded | Excluded: other report family | None | None | LabVIEW documentation generation is outside scope |
| Word-specific operations | [Word Specific](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/word-mnu.html) | Excluded from active roadmap | Excluded: other report family | None | None | Possible separately scoped future extension |

## Excel Easy VIs

| NI VI / instance | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| Excel Easy Title VI | [NI reference](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-title-vi.html) | Eligible | Partially matched | `WriteCellString*`, formatting APIs | No equivalent Easy wrapper | Title text and position can be composed; named range, complete font contract, defaults, and `font out` remain |
| Excel Easy Text VI | [NI reference](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-text-vi.html) | Eligible | Partially matched | `WriteCellString*`, formatting APIs | `Write Cell String.vi` is lower level | Text write exists; named range, integrated font/format contract, outputs, and defaults remain |
| Excel Easy Table VI | [NI reference](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-table-vi.html) | Eligible | Partially matched | `WriteStringRangeByIndex` | `Write String Table.vi` | Polymorphic inventory and full connector contract remain |
| Excel Easy Table (str) VI | Individual NI reference to inventory | Eligible | Partially matched | `WriteStringRangeByIndex` | `Write String Table.vi` | Basic 2D string transfer exists; row/column headers, auto-format, named range, iteration, font, and next-cell outputs remain |
| Excel Easy Table (num) VI | [NI reference](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-table-num-vi.html) | Eligible | Planned | Only single-cell double writes exist | `Write Cell Double.vi` | Requires bulk numeric range API plus headers, auto-format, named range, iteration, font, and next-cell outputs |
| Excel Easy Graph VI | [NI reference](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-graph-vi.html) | Eligible where native Open XML permits | Planned | None | None | Requires native chart model and NI graph-type mapping; renderer-dependent behavior must be classified |

## Excel General VIs

Parent palette: [Excel General](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelgen-mnu.html).

The NI child list is not yet fully expanded. The following known operations map to the current implementation; each must be replaced or supplemented by an exact public VI row as inventory work proceeds.

| NI operation | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| Get/select worksheet by name | Parent palette; individual page pending | Eligible | Partially matched | `GetWorksheetNames`, `SelectWorksheetByName` | `Get Worksheet Names.vi`, `Select Worksheet by Name.vi` | Core behavior exists; exact NI outputs/defaults require comparison |
| Get/select worksheet by index | Parent palette; individual page pending | Eligible | Partially matched | `SelectWorksheetByIndex` | `Select Worksheet by Index.vi` | Zero-based selection covered; full connector parity pending |
| Get active worksheet name | Parent palette; individual page pending | Eligible | Partially matched | `GetActiveWorksheetName` | `Get Active Worksheet Name.vi` | Current helper may not correspond one-to-one to an NI public VI |
| Get last row / used range | Parent palette; individual page pending | Eligible | Planned | None | None | Required for append semantics |
| Add/copy/rename/delete/reorder worksheet | Parent palette; individual pages pending | Eligible | Planned | None | None | Deterministic Open XML operations |
| Insert/delete/clear/search cells or ranges | Parent palette; individual pages pending | Eligible | Planned | None | None | Search and structure mutation remain |
| Read cell as string | Parent palette; exact NI row pending | Eligible except formula calculation | Partially matched | `ReadCellString*` | `Read Cell String.vi` | Shared, inline, Boolean, numeric, and blank values covered; formulas intentionally rejected |
| Read string range | Parent palette; exact NI row pending | Eligible except formula calculation | Partially matched | `ReadStringRangeByIndex` | `Read String Range.vi` | Flat row-major bridge covered; broader typed instances remain |
| Write string/double/Boolean cell | Parent palette; exact NI rows pending | Eligible | Partially matched | typed `WriteCell*` APIs | polymorphic `Write Cell.vi` | Three types covered; dates, errors, rich text, and other documented types pending |
| Convert zero-based coordinates and A1 address | Project helper; NI equivalent to confirm | Eligible | Matched for project contract | conversion static APIs | `Coordinate2Address.vi`, `Address2Coordinate.vi` | Boundary and normalization tests exist; NI inventory relationship still to confirm |
| Named ranges and hyperlinks | Parent palette; individual pages pending | Eligible | Planned | None | None | Required by Easy VI targeting and general parity |
| Row/column geometry and visibility | Parent palette; individual pages pending | Eligible | Planned | image geometry reads existing dimensions internally | None | No public manipulation API |
| Merge/unmerge and freeze panes | Parent palette; individual pages pending | Eligible | Planned | None | None | Deterministic Open XML operations |

## Excel Format VIs

Parent palette: [Excel Format](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelfrm-mnu.html).

| NI operation | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| Set cell alignment | Parent palette; individual page pending | Eligible | Partially matched | `SetCellAlignmentByAddress/Index` | `Set Cell Alignment.vi` | Horizontal/vertical alignment, wrap, and rotation covered; indentation, shrink-to-fit, and exact NI signature pending |
| Set cell color and border | Parent palette; individual page pending | Eligible | Partially matched | `SetCellColorAndBorderByAddress/Index` | `Set Cell Color and Border.vi` | Fill and basic edge/style model covered; complete NI border model and style deduplication remain |
| Set font | Parent palette; individual page pending | Eligible | Planned | None | None | Name, size, color, bold, italic, underline, and strike-through required |
| Set number format | Parent palette; individual page pending | Eligible | Planned | None | None | Dates, times, percentages, currencies, and custom formats required |
| Set row height / column width | Parent palette; individual pages pending | Eligible | Planned | None | None | Public geometry APIs required |
| AutoFit | Parent palette; individual page pending | Eligible only as deterministic approximation | Requires characterization | None | None | Exact Excel renderer behavior is unavailable; any approximation must be an intentional difference |
| Auto-format/table presets | Parent palette; individual pages pending | Eligible where representable | Requires characterization | None | None | NI preset mapping must be characterized |

## Excel Graphs and Pictures VIs

Parent palette: [Excel Graphs and Pictures](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelpic-mnu.html).

| NI operation | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| Append image | Individual page pending | Eligible | Partially matched | `AppendImage` | `Append Image.vi` | PNG/JPEG/GIF/BMP/TIFF, A1 precedence, insertion order, and deduplication covered |
| Format image | Individual page pending | Eligible | Partially matched | `FormatImage` | `Format Image.vi` | Scaling, US/metric dimensions, index selection, automatic and grayscale covered |
| Image alignment parameter | Append-image reference pending | Eligible | Intentional documented behavior | Accepted and ignored | `Append Image.vi` | Matches observed NI Excel behavior; see `IMAGE_TEST_RESULTS.md` |
| Image caption parameter | Append-image reference pending | Eligible | Intentional documented behavior | Accepted and ignored | `Append Image.vi` | Matches observed NI Excel behavior; accessibility extension would be a deliberate difference |
| Black-and-white / watermark picture modes | Individual pages pending | Eligible where Open XML permits | Planned | None | control values exist but backend rejects unsupported modes | Requires semantic mapping and tests |
| Crop/remove/reposition/enumerate pictures | Individual pages pending | Eligible | Planned | None | None | Full child inventory pending |
| Create native Excel graph | Individual pages pending | Eligible | Planned | None | None | Requires chart parts, data references, layout, and type mapping |
| Format toolkit-created graph | Individual pages pending | Eligible | Planned | None | None | NI states these VIs manipulate toolkit-created graphs, not arbitrary manual graphs |
| Render graph through live Excel | Individual pages pending | Excluded | Excluded: ActiveX/Office automation | None | None | Live rendering is outside scope |

## Excel Advanced VIs

Parent palette: [Excel Advanced](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/exceladv-mnu.html).

This palette requires a complete child inventory before implementation planning. Known current or policy-level mappings are recorded below.

| NI operation or family | NI reference | Eligibility | Status | Backend/API mapping | LabVIEW mapping | Evidence / remaining gap |
|---|---|---|---|---|---|---|
| Lock worksheets | Individual page pending | Eligible | Partially matched | `LockAllWorksheets` | `Lock Report.vi` | Idempotent legacy password protection for all worksheets; per-sheet/options parity pending |
| Unlock worksheets | Individual page pending | Eligible | Partially matched | `UnlockAllWorksheets` | `Unlock Report.vi` | Removes sheet protection; exact NI scope pending |
| Workbook/worksheet protection options | Parent palette; pages pending | Eligible | Planned | basic all-sheet protection only | basic wrappers only | Full option model and workbook protection remain |
| Preserve macro-enabled workbook/template | Relevant pages pending | Eligible | Planned | None | None | Future `.xlsm`/`.xltm` preservation without execution |
| Run macro | Individual page pending | Excluded | Excluded: ActiveX/Office automation | None | None | Permanent exclusion |
| Get Excel ActiveX references | Individual page pending | Excluded | Excluded: ActiveX/Office automation | None | None | Permanent exclusion |
| Refresh/recalculate workbook | Individual pages pending | Excluded | Excluded: calculation | None | None | Permanent exclusion when it requires Excel calculation/runtime |

## Current XCelReportEngine helper and lifecycle API inventory

These APIs are tracked here even when their exact NI counterpart is not yet identified.

| XCelReportEngine API | LabVIEW VI | Current test evidence | NI mapping state |
|---|---|---|---|
| `ReportEngineApi.Create` | constructor in `New Report.vi` | lifecycle tests | Generic infrastructure |
| `OpenWorkbook` | `New Report.vi` | XLSX/XLTX open tests | Partial Create/New Report mapping |
| `ValidateSession` | private validation wrapper | open/closed/unknown tests | Internal compatibility support |
| `SaveWorkbook` | `Save Report.vi` | atomic publish and open-session test | Partial Save Report mapping |
| `CloseWorkbook` | `Close Report.vi` | lifecycle tests | Partial Dispose/Close mapping |
| worksheet name/select APIs | worksheet public VIs | worksheet tests | Excel General mapping pending exact rows |
| typed cell APIs | `Write Cell.vi` instances | typed round-trip tests | Excel General/Easy mapping pending exact rows |
| string range APIs | string table/read VIs | row-major range tests | Excel Easy/General mapping partial |
| alignment API | `Set Cell Alignment.vi` | range/style preservation tests | Excel Format partial |
| fill/border API | `Set Cell Color and Border.vi` | independent fill/border tests | Excel Format partial |
| image APIs | image VIs | geometry, deduplication, scaling, grayscale tests | Graphs and Pictures partial |
| worksheet protection APIs | lock/unlock VIs | idempotency and validation tests | Excel Advanced mapping pending exact rows |

## Inventory completion queue

The next documentation pass must add exact rows, URLs, signatures, and polymorphic instances in this order:

1. general Report Generation VIs relevant to Excel;
2. Advanced Report Generation VIs relevant to Excel;
3. Report Layout VIs relevant to Excel;
4. every Excel General child VI;
5. every Excel Format child VI;
6. every Excel Graphs and Pictures child VI;
7. every Excel Advanced child VI;
8. every polymorphic instance under Excel Easy Table and Excel Easy Graph;
9. exact NI error-code mappings for eligible operations.

Each completed inventory row must identify the connector contract, default values, zero/one-based conventions, named-range precedence, output semantics, deterministic Open XML feasibility, and required characterization fixtures.

## Related documents

- [RoadMap.md](RoadMap.md) — scope, phases, and definition of done.
- [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md) — direct NI-to-XCelReportEngine replacement map.
- [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md) — implemented public contract.
- [IMAGE_TEST_RESULTS.md](IMAGE_TEST_RESULTS.md) — characterized NI image behavior.
- [ARCHITECTURE_DECISIONS.md](ARCHITECTURE_DECISIONS.md) — accepted architecture.
- [DEVELOPMENT.md](DEVELOPMENT.md) — build and verification procedures.
