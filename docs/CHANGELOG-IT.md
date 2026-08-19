# Changelog

Versione inglese predefinita: [CHANGELOG.md](CHANGELOG.md).

Le modifiche rilevanti del progetto saranno documentate in questo file.

## Unreleased

- aggiunte la roadmap canonica del prodotto e le istruzioni generali di progetto valide per l'intero repository;
- aggiunto l'inventario iniziale di parità con il NI Toolkit e la mappatura dell'implementazione corrente;
- definite ActiveX, COM, automazione Office e calcolo delle celle come esclusioni permanenti dal perimetro di compatibilità;
- aggiunta l'esecuzione automatica degli stessi 25 test backend su .NET Framework 4.0 e 4.8;
- resa predefinita la documentazione inglese, conservando le versioni italiane con suffisso `-IT`;
- documentati i limiti correnti di build e test dell'integrazione LabVIEW 2013;

- aggiunto supporto per .net40/LabVIEW 2013 (untested)
- aggiunte icone, packaging configuration (VIPM)
- aggiunta configurazione per antidoc
- rinominato il componente in `XCelReportEngine`;
- separati gli ambienti `.net`, `LV`, `docs` e `tools`;
- aggiunta configurazione Git per i binari LabVIEW;
- aggiunti lock file NuGet e build GitHub Actions su Windows;
- isolato il parsing degli indirizzi Excel nel componente interno `CellAddress`;
- ottimizzata la lettura dei range evitando scansioni XML ripetute;
- aggiunti helper pubblici per convertire coordinate base zero e indirizzi A1;
- aggiunti i wrapper LabVIEW `Coordinate2Address.vi` e `Address2Coordinate.vi`;
- aggiunte le VI polimorfiche LabVIEW `Select Worksheet.vi` e `Write Cell.vi`;
- validate in LabVIEW le VI polimorfiche e le conversioni coordinate/A1, incluso il caso di indirizzo non valido;
- esclusi temporaneamente dal repository pubblico test e template LabVIEW non ancora sanitizzati;
- escluse temporaneamente le analisi derivate da progetti e template aziendali;
- implementati e verificati apertura, worksheet, celle, range, Lock/Unlock e immagini.
- aggiunte API di formattazione OpenXML per allineamento, riempimento e bordi di celle e range.
