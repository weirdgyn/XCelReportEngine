# Architecture decisions

Italian version: [ARCHITECTURE_DECISIONS-IT.md](ARCHITECTURE_DECISIONS-IT.md).

## ADR-001 — .NET target for the LabVIEW backend

Status: accepted, later extended by ADR-002.

### Context

- Current environment: LabVIEW 2026.
- Integration: an in-process assembly loaded through LabVIEW .NET nodes.
- Backend: Open XML SDK.
- Platform: Windows, AnyCPU.

### Decision

The original public assembly and MVP backend target is .NET Framework 4.8 (`net48`, AnyCPU). This matches the CLR used by LabVIEW 2026 without requiring .NET Framework 4.8.1. A classic `.NET 8+` target is unsuitable for direct loading by LabVIEW's .NET Framework nodes and would require a separate process or another interoperability mechanism.

### Consequences

- Deploy `XCelReportEngine.dll`, `DocumentFormat.OpenXml.dll`, and `DocumentFormat.OpenXml.Framework.dll` together.
- Keep the public API limited to simple LabVIEW-compatible types.
- Microsoft Office is not required.

## ADR-002 — Separate legacy backend target

Status: accepted, integration experimental.

### Decision

The shared backend also targets .NET Framework 4.0 (`net40`) for the exported LabVIEW 2013 wrappers. Both backend targets compile from the same source, and the same 25 functional tests run against each target.

### Limitations

Backend compatibility does not prove LabVIEW 2013 integration compatibility. The exported VIs still require relinking, compilation, packaging, and smoke testing in a dedicated LabVIEW 2013 environment. This environment is not currently available alongside the LabVIEW 2026 workstation, so `LV2013` remains experimental.
