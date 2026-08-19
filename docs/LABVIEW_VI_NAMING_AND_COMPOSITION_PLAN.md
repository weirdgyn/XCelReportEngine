# LabVIEW VI naming and composition plan

## Purpose

This document defines how the existing XCelReportEngine LabVIEW API should evolve toward direct NI Report Generation Toolkit migration compatibility without turning the .NET backend into a collection of one-method-per-NI-VI special cases.

The intended architecture has three layers:

```text
NI-compatible public facade VIs
            ↓
Reusable XCelReportEngine LabVIEW primitives and subVIs
            ↓
Generic .NET Open XML services
```

The facade optimizes migration. The primitive layer optimizes LabVIEW reuse. The backend exposes file-model operations that are efficient or impractical to implement in G.

## Governing principles

1. Prefer one public compatibility-facade VI for every eligible public NI VI.
2. Match the official NI functional name as closely as practical.
3. Match connector-pane positions, required/recommended/optional terminals, defaults, enum ordering, array dimensions, and error flow wherever types permit.
4. Do not reproduce an NI-specific VI as a dedicated .NET method when it can be composed cleanly in LabVIEW.
5. Add .NET methods for reusable Open XML primitives, bulk data transfer, package traversal, and operations requiring semantic package knowledge.
6. Keep the native XCelReportEngine primitives available for applications that do not need NI compatibility.
7. Do not silently change an existing public VI into a different contract. Use a compatibility facade and a deprecation period where necessary.
8. Maintain the same public facade and primitive names under `LV` and `LV2013`.
9. Rename, move, relink, and save `.vi`, `.ctl`, `.lvlib`, `.lvproj`, and polymorphic members only from the LabVIEW IDE.
10. Never introduce ActiveX, Office automation, formula calculation, or live Excel dependencies to obtain closer naming parity.

## Naming model

### Compatibility facade

Use the documented NI functional name, including the `Excel` prefix and `VI` filename convention where applicable:

```text
Excel Easy Title.vi
Excel Easy Text.vi
Excel Easy Table.vi
Excel Easy Table (str).vi
Excel Easy Table (num).vi
Excel Easy Graph.vi
Excel Set Cell Alignment.vi
Excel Set Cell Color and Border.vi
Excel Format Image.vi
Save Report to File.vi
Dispose Report.vi
```

These VIs belong in the public compatibility palettes and should use project-owned icons carrying a consistent XCelReportEngine compatibility badge.

### Native public primitives

Native primitives should use concise file-model language and must not imply full NI parity. Existing names may remain during the transition. New ambiguous primitives should use an `XRE` prefix or an explicit noun such as Workbook, Worksheet, Cell, Range, Picture, or Chart:

```text
XRE Write Cell.vi
XRE Write String Range.vi
XRE Write Double Range.vi
XRE Apply Cell Font.vi
XRE Resolve Named Range.vi
XRE Append Picture.vi
XRE Save Workbook.vi
XRE Close Workbook.vi
```

The `XRE` prefix is primarily a collision and intent marker. It is not required for existing names that are already unambiguous and stable, but new facade names must never collide with native primitive names.

### Private composition subVIs

Use an action plus a clear `Core`, `Build`, `Resolve`, `Convert`, or `Calculate` suffix. Private subVIs must not contain NI branding in their names:

```text
Resolve Excel Target Core.vi
Build Table With Headers.vi
Calculate Table Bounds.vi
Convert Font Settings.vi
Apply Easy Cell Formatting Core.vi
Map Graph Type Core.vi
Validate Table Dimensions.vi
```

### Polymorphic instances

The public polymorphic VI and instance suffixes should follow the official NI documentation when known. Do not invent alternative abbreviations after the inventory is fixed:

```text
Excel Easy Table.vi
  Excel Easy Table (str).vi
  Excel Easy Table (num).vi
```

Native typed primitives may use explicit CLR/LabVIEW type names when they are not compatibility facades:

```text
XRE Write Cell String.vi
XRE Write Cell Double.vi
XRE Write Cell Boolean.vi
```

## Proposed alignment of existing VIs

No binary rename should be executed until the exact NI row and connector contract are recorded in [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md).

| Current VI | Long-term role | Target name / facade | Action |
|---|---|---|---|
| `New Report.vi` | Public lifecycle facade | Keep provisionally; confirm exact NI Create/New Report name | Align connector and defaults after inventory |
| `Save Report.vi` | Native save primitive | `XRE Save Workbook.vi` | Add `Save Report to File.vi` facade; keep deprecated forwarding wrapper for one compatibility release |
| `Close Report.vi` | Native close primitive | `XRE Close Workbook.vi` | Add `Dispose Report.vi` facade; preserve old wrapper during transition |
| `Lock Report.vi` | Native worksheet protection primitive | `XRE Lock Worksheets.vi` | Add exact NI Excel facade after inventory |
| `Unlock Report.vi` | Native worksheet protection primitive | `XRE Unlock Worksheets.vi` | Add exact NI Excel facade after inventory |
| `Get Worksheet Names.vi` | Native worksheet primitive | `XRE Get Worksheet Names.vi` | Add exact NI facade only if a documented NI counterpart exists |
| `Get Active Worksheet Name.vi` | Native helper | `XRE Get Active Worksheet Name.vi` | Keep as project helper unless NI mapping is found |
| `Select Worksheet.vi` | Native polymorphic primitive | `XRE Select Worksheet.vi` | Compatibility facade should use the exact NI worksheet VI name |
| `Select Worksheet by Name.vi` | Native instance | `XRE Select Worksheet by Name.vi` | Move behind native polymorphic VI or compatibility facade |
| `Select Worksheet by Index.vi` | Native instance | `XRE Select Worksheet by Index.vi` | Move behind native polymorphic VI or compatibility facade |
| `Write Cell.vi` | Native polymorphic primitive | `XRE Write Cell.vi` | Retain as reusable low-level API |
| `Write Cell String.vi` | Native typed instance | `XRE Write Cell String.vi` | Used by Easy Text/Title facades |
| `Write Cell Double.vi` | Native typed instance | `XRE Write Cell Double.vi` | Used by numeric table fallback and direct users |
| `Write Cell Boolean.vi` | Native typed instance | `XRE Write Cell Boolean.vi` | Retain as native typed write |
| `Write String Table.vi` | Native bulk primitive | `XRE Write String Range.vi` | Used by `Excel Easy Table (str).vi` |
| `Read Cell String.vi` | Native read primitive | `XRE Read Cell String.vi` | Add exact NI typed facade after inventory |
| `Read String Range.vi` | Native bulk primitive | `XRE Read String Range.vi` | Add exact NI typed facade after inventory |
| `Append Image.vi` | Native picture primitive | `XRE Append Picture.vi` | Add generic/Excel NI-compatible append facade(s) |
| `Format Image.vi` | Near-compatible Excel operation | `Excel Format Image.vi` facade over `XRE Format Picture.vi` | Split facade from native primitive when connector contract is fixed |
| `Set Cell Alignment.vi` | Native formatting primitive | `XRE Set Cell Alignment.vi` | Add `Excel Set Cell Alignment.vi` facade |
| `Set Cell Color and Border.vi` | Native formatting primitive | `XRE Set Cell Color and Border.vi` | Add `Excel Set Cell Color and Border.vi` facade |
| `Coordinate2Address.vi` | Project helper | `XRE Coordinate to Address.vi` | Keep outside NI coverage percentage |
| `Address2Coordinate.vi` | Project helper | `XRE Address to Coordinate.vi` | Keep outside NI coverage percentage |

Before adopting the `XRE` renames, verify whether existing users require source compatibility. If compatibility must be preserved, leave the current public wrappers in a `Legacy` palette and implement them as thin forwarders.

## Palette structure

Recommended LabVIEW library organization:

```text
XCel Report Engine.lvlib
├── Compatibility
│   ├── General
│   ├── Layout
│   └── Excel
│       ├── Easy
│       ├── General
│       ├── Format
│       ├── Graphs and Pictures
│       └── Advanced
├── Native
│   ├── Workbook
│   ├── Worksheet
│   ├── Cell and Range
│   ├── Formatting
│   ├── Pictures and Charts
│   └── Helpers
├── Controls
├── Private
│   ├── Composition
│   ├── Conversion
│   └── Validation
└── DotNet
```

Only `Compatibility`, `Native`, and supported controls are public. Composition, conversion, validation, and .NET artifacts remain private.

## Reusable LabVIEW subVI backlog

These subVIs are intended to let the project implement complex NI facades without adding super-specific backend calls.

| SubVI | Layer | Responsibility | Backend dependency |
|---|---|---|---|
| `Resolve Excel Target Core.vi` | Private composition | Apply precedence between named range, A1 address, and zero-based coordinates; return normalized bounds | Generic named-range resolver required |
| `Convert Font Settings.vi` | Private conversion | Convert NI-compatible font cluster/enums to native formatting inputs and build `font out` | None |
| `Apply Easy Cell Formatting Core.vi` | Private composition | Apply font, alignment, fill, border, and number format to resolved bounds | Generic formatting primitives |
| `Build Table With Headers.vi` | Private composition | Combine data with optional row and column headers | None; pure G array work |
| `Calculate Table Bounds.vi` | Private calculation | Return start/end, next bottom-left, and next top-right coordinates | None |
| `Validate Table Dimensions.vi` | Private validation | Check header lengths, dimensions, and NI-compatible error cases | None |
| `Map Auto Format Core.vi` | Private conversion | Map NI auto-format enum to a reusable project style description | Generic style application; behavior requires characterization |
| `Map Graph Type Core.vi` | Private conversion | Map NI graph enum to backend chart type and options | Generic chart service |
| `Build Chart Series Core.vi` | Private composition | Convert table data and headers into generic chart-series arrays | Generic chart service |
| `Calculate Iteration Target.vi` | Private calculation | Implement characterized Easy Table iteration placement | None after NI behavior is known |
| `Convert NI Error Core.vi` | Private conversion | Map backend/project errors to the closest documented NI-compatible error behavior | Stable backend errors |

## Generic .NET backend backlog

These methods are justified because they represent reusable Open XML operations or efficient bulk transfers. Names are conceptual until API design review.

| Generic backend capability | Used by | Why it belongs in .NET |
|---|---|---|
| Resolve named range to worksheet and bounds | Easy Title/Text/Table, General search/targeting | Requires workbook defined-name parsing and worksheet resolution |
| Bulk double range read/write | Numeric Easy Table and native bulk API | Avoids one .NET boundary call and XML save per cell |
| Bulk Boolean range read/write | Typed tables and native bulk API | Same performance and consistency reason |
| Apply font to cell/range | Easy Title/Text/Table and Excel Format | Requires style-table manipulation and deduplication |
| Apply number format to cell/range | Numeric tables and Excel Format | Requires style and numbering-format management |
| Deduplicated style repository | Every formatting operation | Prevents Excel style explosion and centralizes style identity |
| Worksheet create/copy/rename/delete/reorder | Excel General facades | Package relationships and workbook metadata must remain consistent |
| Row/column insert/delete/size/visibility | Excel General/Format | Requires coordinated reference and worksheet updates |
| Merge/unmerge and panes | Excel General | Reusable worksheet operations |
| Create/update generic native chart | Easy Graph and graph formatting facades | Chart parts, relationships, caches, and drawing anchors are package-level concerns |
| Picture enumerate/remove/crop/recolor | Graphs and Pictures facades | DrawingML package operation |
| Page setup/header/footer metadata | General/Layout/Excel Advanced facades | SpreadsheetML file-model operation |
| Macro-enabled package preservation | lifecycle facade | Content type and package preservation, never VBA execution |

Do not add backend methods named after `Excel Easy Title`, `Excel Easy Text`, or other facade workflows unless profiling proves that composition cannot meet correctness or performance requirements.

## Detailed composition recipes

### Excel Easy Title VI

Proposed facade sequence:

```text
Validate report reference
→ Resolve Excel Target Core
→ XRE Write Cell String
→ Convert Font Settings
→ XRE Apply Cell Font
→ return font out and report out
```

Required new work:

- NI-compatible font cluster and `font settings source` enum;
- generic named-range resolver;
- range font backend operation with style deduplication;
- characterization of target precedence, empty title, invalid font, and default bold behavior;
- connector-pane and default-value comparison.

No Easy Title-specific backend call is required.

### Excel Easy Text VI

Proposed facade sequence:

```text
Validate report reference
→ Resolve Excel Target Core
→ XRE Write Cell String
→ Convert Font Settings
→ Apply Easy Cell Formatting Core
→ calculate documented output position(s)
→ return font out and report out
```

Required new work:

- exact NI signature inventory;
- named-range targeting;
- font backend primitive;
- mapping of alignment, background, border, or style inputs present in the Excel instance;
- tests for default target and named-range precedence.

### Excel Easy Table (str) VI

Proposed facade sequence:

```text
Validate dimensions
→ Resolve Excel Target Core
→ Build Table With Headers
→ XRE Write String Range
→ Map Auto Format Core
→ Apply Easy Cell Formatting Core
→ Calculate Table Bounds
→ return next-cell coordinates, font out, and report out
```

LabVIEW should perform header composition, bounds calculation, and connector semantics. The backend should perform only bulk write, named-range resolution, and generic formatting.

Required characterization:

- whether row/column headers shift the reported next-cell coordinates;
- behavior when headers are shorter or longer than data dimensions;
- `iteration` placement and overwrite behavior;
- empty arrays and zero-sized dimensions;
- auto-format precedence over explicit font settings.

### Excel Easy Table (num) VI

Use the same composition as the string instance, replacing bulk string write with `XRE Write Double Range`.

An interim LabVIEW loop over `Write Cell Double.vi` is acceptable for characterization but not for the production facade because it causes excessive .NET crossings and repeated worksheet saves.

Required backend support:

- generic bulk double write;
- optional generic typed range read if required by the NI contract;
- number-format application.

### Excel Easy Graph VI

Proposed facade sequence:

```text
Validate data and headers
→ Resolve Excel Target Core
→ Map Graph Type Core
→ Build Chart Series Core
→ XRE Create Native Chart
→ apply generic chart title/legend/axis/layout settings
→ return chart index or documented NI outputs
```

The chart creator must be a generic backend service, not an Easy Graph-specific method. LabVIEW owns NI enum mapping and facade defaults; .NET owns ChartML/DrawingML creation and package relationships.

Required characterization:

- all NI graph enum values and eligible native ChartML mappings;
- XY versus category-series interpretation;
- row/column header semantics;
- data orientation;
- graph placement and sizing;
- behavior for unsupported or renderer-dependent chart types.

### Save Report to File VI

Proposed facade sequence:

```text
Validate report reference and path/defaults
→ choose Save or Save As behavior
→ XRE Save Workbook
→ preserve open/closed state according to NI contract
→ translate error
```

Add a generic backend Save As operation only if the NI contract requires changing the session output path. Do not add printing, PDF export, or Office format conversion.

### Dispose Report VI

Proposed facade sequence:

```text
Validate report reference
→ conditionally save according to NI inputs/defaults
→ XRE Close Workbook
→ invalidate XRE report reference
→ translate error
```

### Append Image to Report / Excel Format Image

Keep append and formatting as separate reusable primitives. Generic compatibility wrappers may compose:

```text
Resolve Excel Target Core
→ XRE Append Picture
→ optional XRE Format Picture
→ return report and picture index if exposed
```

Retain the characterized intentional differences documented in [IMAGE_TEST_RESULTS.md](IMAGE_TEST_RESULTS.md).

## Implementation work packages for the LabVIEW author

### Work package A — Facade infrastructure

1. Create `Compatibility`, `Native`, and `Private/Composition` virtual folders in the LVLIB.
2. Define a project-owned compatibility icon template and badge.
3. Define typedefs for NI-compatible font settings, position, and common enums without copying NI assets.
4. Implement `Resolve Excel Target Core.vi` after the named-range backend exists.
5. Implement common validation, font conversion, bounds calculation, and error conversion subVIs.
6. Add a small facade smoke-test harness that exercises connector defaults and error propagation.

### Work package B — Existing primitive normalization

1. Classify every current public VI as facade, native primitive, helper, or legacy wrapper.
2. Apply approved native names from the mapping table using the LabVIEW 2026 IDE.
3. Repair LVLIB, project, polymorphic VI, palette, VIPB, glyph, and documentation references from the IDE.
4. Retain forwarding wrappers for renamed public VIs when source compatibility is required.
5. Run existing LabVIEW 2026 smoke tests before adding new facades.

### Work package C — Easy Title and Text

1. Add reusable font typedefs and conversion subVIs.
2. Add the generic backend font/range operation and named-range resolver.
3. Implement `Excel Easy Title.vi` solely by composition.
4. Implement `Excel Easy Text.vi` solely by composition.
5. Characterize defaults and compare output workbooks semantically with NI.

### Work package D — Easy Table

1. Implement header building, dimension validation, bounds, and iteration subVIs in G.
2. Add generic bulk double write and generic number formatting to the backend.
3. Normalize the current string table primitive.
4. Implement string and numeric facade instances.
5. Assemble the polymorphic `Excel Easy Table.vi`.
6. Verify large-table performance and style reuse.

### Work package E — Graph foundation

1. Inventory and classify NI graph types.
2. Define generic chart DTO inputs compatible with LabVIEW simple types.
3. Implement generic ChartML creation in .NET.
4. Implement graph-type and series-building subVIs in G.
5. Implement `Excel Easy Graph.vi` as a facade.
6. Add golden package and desktop rendering verification.

### Work package F — Mirroring to LabVIEW 2013

For each completed LabVIEW 2026 work package:

1. stabilize names and connector panes in `LV` first;
2. export/save the affected hierarchy for LabVIEW 2013 from the supported IDE workflow;
3. reproduce the same relative folder, LVLIB, polymorphic, typedef, and icon layout under `LV2013`;
4. relink .NET nodes to the staged `net40` assembly when a LabVIEW 2013 environment becomes available;
5. record exported-but-unverified status explicitly;
6. do not claim LabVIEW 2013 validation until it is built and smoke-tested in LabVIEW 2013.

Do not independently evolve the `LV2013` API. It is a version-compatible export of the canonical LabVIEW 2026 facade and native surface.

## Rename execution checklist

For every approved rename:

1. confirm the exact NI mapping and target name in [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md);
2. identify callers, polymorphic membership, typedef dependencies, glyph, palette, VIPB, Antidoc, project, and LVLIB references;
3. rename or move the VI from LabVIEW 2026, not from the filesystem;
4. save all callers and the owning library/project;
5. reopen the project and resolve all missing dependencies;
6. run mass compile only when appropriate and inspect changes carefully;
7. run relevant LabVIEW smoke tests;
8. verify the .NET build and both backend test targets;
9. update migration map, parity matrix, technical reference, and changelog;
10. export the stabilized hierarchy for LabVIEW 2013 and mark it unverified.

## Definition of done for a compatibility facade

A facade is complete when:

- the official NI VI name, URL, icon reference, connector pane, defaults, and polymorphic instances are inventoried;
- one public XCelReportEngine VI provides the eligible behavior;
- the implementation composes reusable LabVIEW primitives and generic backend operations;
- project-owned icon and help text are present;
- error flow and relevant NI error mapping are tested;
- semantic NI golden comparison exists or an intentional difference is documented;
- LabVIEW 2026 smoke tests pass;
- the corresponding LabVIEW 2013 hierarchy is exported and honestly marked as verified or unverified;
- [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md) and [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md) are updated.

## Related documents

- [RoadMap.md](RoadMap.md)
- [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md)
- [NI_VI_MIGRATION_MAP.md](NI_VI_MIGRATION_MAP.md)
- [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md)
- [DEVELOPMENT.md](DEVELOPMENT.md)
- [IMAGE_TEST_RESULTS.md](IMAGE_TEST_RESULTS.md)
