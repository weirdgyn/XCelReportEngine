# Sviluppo e manutenzione

## Prerequisiti

- Windows;
- LabVIEW 2026;
- .NET SDK 10 con targeting pack .NET Framework 4.8;
- accesso ai pacchetti NuGet oppure cache locale già popolata.

## Build riproducibile

```powershell
dotnet restore .net/XCelReportEngine.sln --locked-mode -m:1
dotnet build .net/XCelReportEngine.sln -c Release --no-restore -m:1
dotnet test .net/XCelReportEngine.sln -c Release --no-build --no-restore -m:1
```

La build seriale è intenzionale: nell'ambiente di sviluppo corrente la build parallela della solution con SDK .NET 10 può terminare senza una diagnostica utile.

I file `packages.lock.json` sono versionati. Quando si aggiorna una dipendenza bisogna eseguire consapevolmente:

```powershell
dotnet restore .net/XCelReportEngine.sln --force-evaluate -m:1
```

## Dipendenza tra .NET e LabVIEW

La LVLIB carica gli assembly prodotti in:

```text
.net/src/XCelReportEngine/bin/Release/net48/
```

LabVIEW mantiene gli assembly .NET caricati fino alla chiusura del processo. Prima di ricompilare nella cartella Release è quindi necessario chiudere completamente LabVIEW. In alternativa si può usare una cartella di output temporanea per la sola verifica .NET.

Le firme pubbliche di `ReportEngineApi` costituiscono il contratto di interoperabilità. Una modifica a classe, namespace, assembly, firma o tipo di parametro richiede l'aggiornamento manuale dei Constructor/Invoke Node e del typedef `Report Ref.ctl`.

## Regole per i file LabVIEW

- rinominare o spostare `.vi`, `.ctl`, `.lvlib` e `.lvproj` soltanto dall'IDE LabVIEW;
- salvare tutti i chiamanti dopo una modifica a un typedef;
- non versionare `.lvlps`, `.aliases`, output di test o stato utente;
- verificare la LVLIB e i test LabVIEW prima di creare un commit che modifica file binari;
- descrivere nel commit le VI modificate, perché Git non può produrre un diff testuale utile dei diagrammi.

## Git e pull request

Il branch principale è `main`. Prima di una pull request:

1. eseguire build e test .NET;
2. eseguire gli smoke test LabVIEW pertinenti;
3. controllare `git status --short` per evitare output XLSX o file di stato;
4. aggiornare la documentazione se cambia il contratto pubblico;
5. evitare di mescolare refactoring .NET e modifiche funzionali LabVIEW non correlate.

La workflow GitHub Actions compila e testa la parte .NET su Windows. I test LabVIEW rimangono manuali finché non sarà introdotta una pipeline NI dedicata.

## Asset locali non pubblicati

Le cartelle seguenti sono attualmente escluse da Git:

```text
LV/Test/
LV/Templates/
LV/Pre-Test/
docs/REPORT_INVENTORY_DAS_054_3LIV.md
docs/REPORT_LOCKER_ANALYSIS.md
docs/TEMPLATE_ANALYSIS.md
docs/data/workbook_analysis.json
```

Contengono test, template e risultati di caratterizzazione che devono essere sanitizzati prima della pubblicazione. L'esclusione non cancella i file locali. Per reinserirli sarà necessario rimuovere esplicitamente le relative regole da `.gitignore` dopo la revisione dei contenuti e dei metadati Office.
