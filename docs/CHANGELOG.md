# Changelog

Le modifiche rilevanti del progetto saranno documentate in questo file.

## Unreleased

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
