# XCelReportEngine

Versione inglese predefinita: [TECHNICAL_REFERENCE.md](TECHNICAL_REFERENCE.md).

Backend Excel Open XML destinato a sostituire il sottoinsieme usato del NI Report Generation Toolkit senza dipendere da Microsoft Office, OLE o ActiveX.

## Target

- .NET Framework 4.8 (`net48`)
- AnyCPU
- LabVIEW 2026
- Open XML SDK 3.5.1

## Stato

Implementato:

- API pubblica semplice per LabVIEW;
- sessioni identificate da `Int32`;
- apertura di XLSX;
- materializzazione XLTX -> XLSX;
- elenco e selezione worksheet per nome/indice zero-based;
- salvataggio atomico senza chiudere la sessione;
- chiusura con salvataggio opzionale;
- Lock/Unlock idempotente di tutti i worksheet;
- lettura di celle come stringa, con indirizzo A1 o coordinate zero-based;
- scrittura tipizzata di stringhe, `Double` e booleani;
- preservazione dello stile preesistente durante la scrittura;
- errore deterministico quando si tenta di leggere una formula;
- inserimento immagini embedded con posizione A1 o coordinate zero-based;
- dimensionamento nativo a 96 DPI logici e anchor DrawingML `oneCell`;
- deduplicazione byte-per-byte delle immagini ripetute nello stesso worksheet;
- allineamento, wrap e rotazione testo su celle o range;
- riempimento e bordi applicabili indipendentemente su celle o range;
- errori con codice stabile, nome, operazione e inner exception;
- gli stessi 25 test automatici su net40 e net48.

## API iniziale

```text
ReportEngineApi.Create()
ReportEngineApi.ConvertCellIndexToAddress(rowIndex, columnIndex) -> string
ReportEngineApi.ConvertCellAddressToIndex(cellAddress) -> int[2]
OpenWorkbook(sourcePath, outputPath) -> sessionId
GetWorksheetNames(sessionId) -> string[]
GetActiveWorksheetName(sessionId) -> string
SelectWorksheetByName(sessionId, worksheetName)
SelectWorksheetByIndex(sessionId, worksheetIndex)
LockAllWorksheets(sessionId, password)
UnlockAllWorksheets(sessionId)
ReadCellStringByAddress(sessionId, cellAddress) -> string
ReadCellStringByIndex(sessionId, rowIndex, columnIndex) -> string
WriteCellStringByAddress(sessionId, cellAddress, value)
WriteCellStringByIndex(sessionId, rowIndex, columnIndex, value)
WriteCellDoubleByAddress(sessionId, cellAddress, value)
WriteCellDoubleByIndex(sessionId, rowIndex, columnIndex, value)
WriteCellBooleanByAddress(sessionId, cellAddress, value)
WriteCellBooleanByIndex(sessionId, rowIndex, columnIndex, value)
ReadStringRangeByIndex(sessionId, startRowIndex, startColumnIndex, rowCount, columnCount) -> string[]
WriteStringRangeByIndex(sessionId, startRowIndex, startColumnIndex, rowCount, columnCount, values)
SetCellAlignmentByAddress(sessionId, startCellAddress, endCellAddress, horizontalAlignment, verticalAlignment, wrapText, textRotation)
SetCellAlignmentByIndex(sessionId, startRowIndex, startColumnIndex, rowCount, columnCount, horizontalAlignment, verticalAlignment, wrapText, textRotation)
SetCellColorAndBorderByAddress(sessionId, startCellAddress, endCellAddress, applyFill, fillColor, applyBorder, borderColor, borderStyle, borderEdges)
SetCellColorAndBorderByIndex(sessionId, startRowIndex, startColumnIndex, rowCount, columnCount, applyFill, fillColor, applyBorder, borderColor, borderStyle, borderEdges)
AppendImage(sessionId, imagePath, rowIndex, columnIndex, cellAddress, alignment, caption) -> pictureIndex
FormatImage(sessionId, measurementSystem, scaleFactor, pictureIndex, height, width, colorType)
SaveWorkbook(sessionId)
CloseWorkbook(sessionId, saveChanges)
```

Le firme evitano generic, delegate, overload e tipi Open XML pubblici.

Gli indici di riga e colonna sono base zero: `(0, 0)` corrisponde ad `A1`. Gli indirizzi A1 non sono case-sensitive e accettano anche la forma assoluta `$A$1`. La lettura restituisce stringa vuota per una cella assente o vuota, `TRUE`/`FALSE` per i booleani e la rappresentazione invariant-culture memorizzata per i numeri. Le formule non vengono valutate e causano `FormulaReadNotSupported`.

I range di stringhe attraversano il confine .NET/LabVIEW come array monodimensionali in ordine row-major. L'elemento `(riga, colonna)` si trova all'indice `riga * columnCount + colonna`. I wrapper LabVIEW convertono tra questo formato e gli array 2D con `Reshape Array`. La lettura fallisce per l'intero range se contiene almeno una formula.

Per `AppendImage`, una `cellAddress` non vuota ha precedenza sulle coordinate numeriche. `alignment` accetta i valori NI 0..8 e, come osservato nel backend Excel NI, non modifica il drawing; anche `caption` viene accettata ma ignorata. Il metodo restituisce il picture index zero-based. Il file viene incorporato senza ricodifica e senza relazione esterna al path sorgente.

Le API di formattazione usano esclusivamente tipi semplici:

- `horizontalAlignment`: 0 General, 1 Left, 2 Center, 3 Right, 4 Fill, 5 Justify, 6 CenterContinuous, 7 Distributed;
- `verticalAlignment`: 0 Bottom, 1 Center, 2 Top, 3 Justify, 4 Distributed;
- `textRotation`: valore OpenXML 0..180;
- colori: `Int32` nel formato RGB `0xRRGGBB`;
- `borderStyle`: 0 None, 1 Thin, 2 Medium, 3 Thick, 4 Double, 5 Dashed, 6 Dotted;
- `borderEdges`: bit mask 1 Left, 2 Right, 4 Top, 8 Bottom; 15 indica tutti i lati.

Un indirizzo finale vuoto indica la sola cella iniziale. `applyFill` e `applyBorder`
consentono di modificare separatamente le due proprietà senza azzerare quella non
richiesta. Tutti gli altri componenti dello stile originale vengono preservati.

## Build e test

```powershell
dotnet build .net/XCelReportEngine.sln -c Release -m:1
dotnet test .net/tests/XCelReportEngine.Tests/XCelReportEngine.Tests.csproj -c Release -f net48 --no-build --no-restore -m:1
.net/tests/XCelReportEngine.Tests/bin/Release/net40/XCelReportEngine.Tests.exe
```

`-m:1` evita un problema osservato nell'SDK .NET 10 installato, che durante la build parallela della solution può terminare senza diagnostica mentre i singoli progetti risultano validi.

Le DLL da distribuire saranno prodotte in `.net\src\XCelReportEngine\bin\Release\net48`.

## Verifiche eseguite

- build Release net48: 0 warning, 0 errori;
- 25 test automatici superati su net40 e 25 su net48;
- validazione schema Open XML degli output di test;
- Lock ripetuto senza duplicazione di `SheetProtection`;
- Unlock e successiva validazione;
- salvataggio atomico mantenendo aperta la sessione;
- conversione fixture XLTX -> XLSX;
- round-trip di shared string, inline string, numero e booleano;
- indirizzi A1 e coordinate zero-based, inclusi controlli sui limiti Excel;
- preservazione dello stile di una cella aggiornata;
- rifiuto esplicito della lettura di formule;
- round-trip di range 2D di stringhe in ordine row-major;
- validazione delle dimensioni e rilevamento delle formule nei range;
- applicazione di allineamento su range con preservazione dei valori;
- applicazione indipendente di riempimento e bordi con preservazione dello stile;
- inserimento di immagini tramite coordinate e indirizzo A1;
- precedenza dell'indirizzo A1, deduplicazione del media e ordine di inserimento;
- smoke test sul template reale `NCH_2022_008_10_LXOD-RS_DRX_DIG.xltx` (11 worksheet, immagini e drawing): apertura, conversione, pubblicazione e riapertura riuscite; content type finale XLSX corretto.
- smoke test LabVIEW del ciclo completo apertura, worksheet, Lock/Unlock, celle, range, immagini, salvataggio e chiusura;
- selezione corretta delle istanze delle VI polimorfiche `Select Worksheet.vi` e `Write Cell.vi`;
- round-trip e gestione errore dei wrapper `Coordinate2Address.vi` e `Address2Coordinate.vi`.

## Output di deployment corrente

```text
XCelReportEngine.dll
DocumentFormat.OpenXml.dll
DocumentFormat.OpenXml.Framework.dll
```

Questi file si trovano insieme nella cartella Release e devono restare insieme nel deployment LabVIEW.

## API LabVIEW

La libreria `XCel Report Engine.lvlib` espone wrapper pubblici per il ciclo di vita del report, worksheet, celle, range e immagini.

Sono disponibili due VI polimorfiche:

```text
Select Worksheet.vi
  Select Worksheet by Name.vi
  Select Worksheet by Index.vi

Write Cell.vi
  Write Cell String.vi
  Write Cell Double.vi
  Write Cell Boolean.vi
```

Gli helper indipendenti da una sessione sono:

```text
Coordinate2Address.vi
Address2Coordinate.vi
```

`Address2Coordinate.vi` restituisce coordinate base zero nell'ordine riga, colonna.
