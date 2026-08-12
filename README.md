# XCelReportEngine

Report engine Excel basato su Open XML per applicazioni LabVIEW, senza dipendenze da Microsoft Office, OLE o ActiveX.

## Struttura

```text
.net/   backend .NET Framework 4.8 e test automatici
LV/     libreria, wrapper, controlli e test LabVIEW
docs/   architettura, analisi e risultati di caratterizzazione
tools/  strumenti di analisi e generazione asset
```

Le cartelle locali `LV/Test`, `LV/Templates` e `LV/Pre-Test` sono temporaneamente escluse dal repository pubblico finché i relativi asset non saranno sanitizzati.

## Build .NET

```powershell
dotnet restore .net/XCelReportEngine.sln -m:1
dotnet build .net/XCelReportEngine.sln -c Release --no-restore -m:1
dotnet test .net/XCelReportEngine.sln -c Release --no-build --no-restore -m:1
```

La documentazione funzionale corrente è in [docs/TECHNICAL_REFERENCE.md](docs/TECHNICAL_REFERENCE.md). Le istruzioni per sviluppo, Git e LabVIEW sono in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Licenza

XCelReportEngine è distribuito secondo i termini della licenza BSD 3-Clause (`BSD-3-Clause`).

Copyright (c) 2026, Michele Santucci. Consultare [LICENSE.txt](LICENSE.txt) per il testo completo della licenza.

Le componenti di terze parti distribuite con il package rimangono soggette alle rispettive licenze.
