# NI VI migration map

## Purpose

This document is the user-facing migration map from public NI LabVIEW Report Generation Toolkit VIs to XCelReportEngine VIs.

The implementation and rename procedure behind these mappings is defined in [LABVIEW_VI_NAMING_AND_COMPOSITION_PLAN.md](LABVIEW_VI_NAMING_AND_COMPOSITION_PLAN.md).

The compatibility matrix in [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md) tracks behavioral coverage. This document tracks migration ergonomics: which XCelReportEngine VI replaces each eligible NI VI, whether the replacement is one-to-one, and whether an interim sequence is required.

The preferred outcome for every eligible NI VI is a single public XCelReportEngine compatibility-facade VI. A documented sequence is a temporary migration aid, not the desired final state.

## Icon and naming policy

- The NI VI name is recorded exactly for identification and linked to its official NI documentation page.
- The NI icon column links to the official NI page that displays the original icon. Original NI icon assets must not be copied into this public repository unless redistribution rights are confirmed.
- XCelReportEngine compatibility VIs should use distinct project-owned icons that preserve the functional visual category and include an identifiable XCelReportEngine mark.
- Compatibility-facade VI names should match the NI functional name as closely as practical while avoiding a false claim that the VI is an NI-authored asset.
- Package and palette documentation must state that XCelReportEngine is an independent implementation and is not affiliated with or endorsed by NI.

## Replacement levels

| Level | Meaning |
|---|---|
| 1:1 available | One public XCelReportEngine VI replaces the NI VI |
| 1:1 planned | A compatibility-facade VI will wrap existing or planned lower-level services |
| Sequence | Migration currently requires multiple XCelReportEngine VIs |
| Partial | A replacement exists but does not cover the full NI contract |
| Excluded | ActiveX, Office automation, calculation, or another excluded report family |

## Initial migration table

The table is populated with the NI VIs already identified in the project discussion and the existing XCelReportEngine public surface. It must be expanded together with the canonical parity inventory.

| NI VI | NI icon/reference | Preferred XCelReportEngine replacement | Current replacement or sequence | Level | Connector/default work |
|---|---|---|---|---|---|
| Excel Easy Title VI | [Official NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-title-vi.html) | `Excel Easy Title.vi` compatibility facade | `Write Cell String.vi` → planned font/range formatting | 1:1 planned | Reproduce position, named range, font source/defaults, `font out`, and error flow |
| Excel Easy Text VI | [Official NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-text-vi.html) | `Excel Easy Text.vi` compatibility facade | `Write Cell String.vi` → `Set Cell Alignment.vi` → planned font formatting | 1:1 planned | Reproduce integrated formatting, named range, outputs, defaults, and terminal placement |
| Excel Easy Table VI | [Official NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-table-vi.html) | polymorphic `Excel Easy Table.vi` | `Write String Table.vi` for the string subset | Partial | Add string/numeric instances, headers, named range, iteration, auto-format, font, and next-cell outputs |
| Excel Easy Table (str) VI | [Parent NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-table-vi.html) | `Excel Easy Table (str).vi` instance | `Write String Table.vi` | Partial | Align connector pane and complete headers/output/default semantics |
| Excel Easy Table (num) VI | [Official NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-table-num-vi.html) | `Excel Easy Table (num).vi` instance | Repeated `Write Cell Double.vi` calls; bulk numeric API not yet available | 1:1 planned | Add bulk numeric write and full Easy Table contract |
| Excel Easy Graph VI | [Official NI page and icon](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/vi-lib/addons/office/excel-llb/excel-easy-graph-vi.html) | `Excel Easy Graph.vi` compatibility facade | No replacement | 1:1 planned | Requires native chart service and documented graph-type mapping |
| Create/New Excel Report | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | NI-equivalent create/new compatibility facade | `New Report.vi` | Partial | Exact NI VI identity, report type/template/default behavior, and connector mapping still to inventory |
| Save Report to File VI | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | `Save Report to File.vi` compatibility facade | `Save Report.vi` | Partial | Align path/save semantics, defaults, terminal positions, and error behavior |
| Dispose/Close Report VI | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | NI-equivalent dispose compatibility facade | `Close Report.vi` | Partial | Confirm save/dispose options and connector contract |
| Append Image to Report VI | [Report Generation VIs](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/functions/report-generation-vis.html) | `Append Image to Report.vi` compatibility facade | `Append Image.vi` | Partial | Align MS Office parameter cluster, polymorphic/path instances, defaults, and errors |
| Select worksheet by name | [Excel General palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelgen-mnu.html) | Exact NI-named facade after inventory | `Select Worksheet by Name.vi` | 1:1 candidate | Confirm original NI VI name, outputs, defaults, and connector pane |
| Select worksheet by index | [Excel General palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelgen-mnu.html) | Exact NI-named facade after inventory | `Select Worksheet by Index.vi` | 1:1 candidate | Confirm original NI VI name and connector contract |
| Read Excel text/cell | [Excel General palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelgen-mnu.html) | Exact typed/polymorphic NI facade after inventory | `Read Cell String.vi` or `Read String Range.vi` | Partial | Inventory exact NI instances, data types, formula behavior, and output shapes |
| Write Excel text/cell | [Excel General palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelgen-mnu.html) | Exact typed/polymorphic NI facade after inventory | polymorphic `Write Cell.vi` | Partial | Inventory exact NI instances and align connector/default behavior |
| Excel Set Cell Alignment VI | [Excel Format palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelfrm-mnu.html) | exact NI-named facade | `Set Cell Alignment.vi` | 1:1 candidate | Compare terminals, enums, range addressing, and default values |
| Excel Set Cell Color and Border VI | [Excel Format palette](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelfrm-mnu.html) | exact NI-named facade | `Set Cell Color and Border.vi` | Partial | Align full border model, connector pane, and NI defaults |
| Excel Append/Insert Image VI | [Excel Graphs and Pictures](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelpic-mnu.html) | exact NI-named facade after inventory | `Append Image.vi` | Partial | Inventory exact VI names and separate append from format contracts |
| Excel Format Image VI | [Excel Graphs and Pictures](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/excelpic-mnu.html) | `Excel Format Image.vi` compatibility facade | `Format Image.vi` | Partial | Align connector, enums, picture index, measurement, and color-type contract |
| Excel worksheet lock/protection VI | [Excel Advanced](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/exceladv-mnu.html) | exact NI-named facade after inventory | `Lock Report.vi` | Partial | Current implementation locks every worksheet; inventory exact NI scope/options |
| Excel worksheet unlock VI | [Excel Advanced](https://www.ni.com/docs/en-US/bundle/lvrgt-api-ref/page/menus/categories/programming/ni-report/exceladv-mnu.html) | exact NI-named facade after inventory | `Unlock Report.vi` | Partial | Inventory exact NI scope and outputs |
| Get Excel ActiveX References VI | Official individual page to inventory | None | None | Excluded | Permanent ActiveX/Office automation exclusion |
| Run Macro VI | Official individual page to inventory | None | None | Excluded | Permanent Office automation exclusion; macro preservation is separate |
| Calculate/Recalculate VI family | Official individual pages to inventory | None | None | Excluded | Permanent calculation exclusion |

## Compatibility-facade design rules

1. Prefer one public compatibility VI for every eligible public NI VI.
2. A compatibility VI may internally call multiple lower-level XCelReportEngine VIs or a single expanded backend method.
3. Keep the current lower-level VIs available as the native XCelReportEngine API; the facade is an additional migration layer, not a replacement.
4. Match connector-pane terminal positions and required/recommended/optional status wherever LabVIEW type compatibility permits.
5. Match default values, enum ordering, zero/one-based conventions, output shapes, and error dataflow.
6. Preserve polymorphic organization when NI exposes meaningful typed instances.
7. When the NI report reference type cannot be reused, place the XCelReportEngine report reference on the analogous terminal and document that rewiring is required.
8. If exact behavior is excluded or intentionally different, keep the VI absent or make the limitation explicit; do not silently emulate it with ActiveX or calculation.
9. Use project-owned icons with a consistent compatibility badge and category colors.
10. Maintain this map whenever a facade VI is added, renamed, completed, or superseded.

## Current XCelReportEngine icon catalog

These project-owned icons are already available for current native VIs. Planned compatibility facades should reuse the relevant functional motif while adding a consistent facade badge.

| XCelReportEngine VI family | Current project icon |
|---|---|
| New Report | ![New Report](../LV/Glyph/New%20Report.png) |
| Save Report | ![Save Report](../LV/Glyph/Save%20Report.png) |
| Close Report | ![Close Report](../LV/Glyph/Close%20Report.png) |
| Write Cell | ![Write Cell](../LV/Glyph/Write%20Cell.png) |
| Write Table | ![Write Table](../LV/Glyph/Write%20Table.png) |
| Read Cell String | ![Read Cell String](../LV/Glyph/Read%20Cell%20String.png) |
| Append Image | ![Append Image](../LV/Glyph/Append%20Image.png) |
| Format Image | ![Format Image](../LV/Glyph/Format%20Image.png) |
| Set Cell Alignment | ![Set Cell Alignment](../LV/Glyph/Set%20Cell%20Alignment.png) |
| Set Cell Color and Border | ![Set Cell Color and Border](../LV/Glyph/Set%20Cell%20Color%20and%20Border.png) |

The final migration table should place the official NI visual reference and the corresponding project-owned facade icon on the same row. Until redistribution rights for NI artwork are established, the NI side remains a link to the official page rather than a copied bitmap.

## Migration expectations

The intended migration workflow is manual replacement with minimal rewiring:

1. locate the NI VI in this table;
2. replace it with the listed XCelReportEngine facade VI;
3. reconnect the XCelReportEngine report reference where the NI class wire is not type-compatible;
4. retain matching data, formatting, and error wires where connector compatibility permits;
5. consult documented intentional differences or exclusions.

Binary drop-in replacement is not promised because the NI report class and the XCelReportEngine report reference are different LabVIEW types. The goal is structural and behavioral migration compatibility, not impersonation of NI binaries.

## Completion criteria for this map

This document is complete when every eligible public NI VI and polymorphic instance has:

- exact official name and documentation URL;
- an official icon reference;
- a single preferred XCelReportEngine facade VI;
- a project-owned XCelReportEngine icon;
- connector-pane comparison;
- current replacement level;
- documented sequence for any temporary multi-VI migration;
- link to the behavioral status in [NI_TOOLKIT_PARITY_MATRIX.md](NI_TOOLKIT_PARITY_MATRIX.md).
