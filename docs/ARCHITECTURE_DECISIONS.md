# Decisioni architetturali

## ADR-001 — Target .NET del backend LabVIEW

Stato: accettato

### Contesto

- ambiente corrente: LabVIEW 2026;
- applicazioni legacy: alcune in LabVIEW 2013, non considerate requisito di retrocompatibilita' iniziale;
- integrazione prevista: assembly caricato in-process tramite i nodi .NET di LabVIEW;
- backend: Open XML SDK;
- piattaforma: Windows, AnyCPU.

### Decisione

Target dell'assembly pubblico e del backend MVP:

```text
.NET Framework 4.8
TFM: net48
Platform: AnyCPU
```

### Motivazione

LabVIEW 2026 richiede/installa .NET Framework 4.8 e supporta fino a 4.8.1. Usare `net48` evita di richiedere esplicitamente 4.8.1 sulle postazioni e rimane allineato al CLR 4 usato dall'integrazione .NET di LabVIEW.

Open XML SDK supporta .NET Framework e il codice ReportLocker esistente e' interamente portabile a `net48`.

### Alternative non scelte

- `net481`: supportato da LabVIEW 2026 ma non necessario; introdurrebbe un prerequisito aggiuntivo se 4.8.1 non fosse installato.
- `netstandard2.0`: tecnicamente consumabile da .NET Framework 4.8, ma non offre vantaggi per un componente esclusivamente LabVIEW/Windows e puo' complicare risoluzione e deployment delle dipendenze.
- `.NET 8`: non adatto all'assembly caricato direttamente dai classici nodi .NET di LabVIEW, che operano sul CLR/.NET Framework; richiederebbe un processo separato o un diverso meccanismo di interoperabilita'.
- `.NET Framework 3.5/4.0`: servirebbe soprattutto per compatibilita' legacy e imporrebbe vincoli non necessari al nuovo codice.

### Conseguenze

- distribuzione di `ReportEngine.dll`, `DocumentFormat.OpenXml.dll`, `DocumentFormat.OpenXml.Framework.dll` e relative dipendenze accanto all'applicazione/nei path risolti da LabVIEW;
- API pubblica limitata a tipi semplici compatibili con LabVIEW;
- nessun requisito di Office installato;
- nessuna promessa di compatibilita' con LabVIEW 2013 nell'MVP;
- possibile valutazione futura di un build legacy separato soltanto in presenza di un caso economico concreto.
