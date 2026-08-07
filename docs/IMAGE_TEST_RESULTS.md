# Risultati test immagini — NI Report Generation Toolkit

## Riferimenti

- Baseline: `../LV/Pre-Test/IMAGE_POSITION_TEST_BASELINE.xlsx`
- Immagine: `../LV/Pre-Test/prova.png`
- PNG RGB, 320 x 200 pixel, 300 DPI dichiarati, 724 byte

## Interfaccia di Append Image to Report (string wrap).vi

Gli ingressi rilevanti comunicati per la VI sono:

1. `MS Office parameters`: bundle contenente posizione numerica e le stringhe per cella iniziale Excel o bookmark Word;
2. allineamento: enum `LEFT`, `RIGHT`, `TOP`, `TEXTTOP`, `MIDDLE`, `ABSMIDDLE`, `BASELINE`, `BOTTOM`, `ABSBOTTOM`;
3. `report in`;
4. path dell'immagine;
5. caption/stringa alternativa;
6. `error in`.

La VI non espone ingressi per larghezza o altezza. I test di ridimensionamento non appartengono quindi a questa API e dovranno essere eseguiti separatamente su `Excel Format Image.vi`.

Per `Append Image to Report (string wrap).vi` restano da caratterizzare:

- effetto dell'enum di allineamento nel backend Excel;
- uso della stringa cella rispetto alle coordinate numeriche;
- precedenza quando sono valorizzati entrambi i meccanismi di posizione;
- memorizzazione della caption/stringa alternativa nelle proprieta' DrawingML;
- inserimento multiplo e generazione degli identificatori.

## Interfaccia di Excel Format Image.vi

Gli ingressi comunicati sono:

1. measurement system: `US` o `metric`;
2. scale factor: double, default `1`;
3. `report in`;
4. picture index: default `-1`, cioe' ultima immagine del worksheet;
5. height: default `-1`, non modifica l'altezza;
6. width: default `-1`, non modifica la larghezza;
7. color type: `msoPictureAutomatic`, `msoPictureMixed`, `msoPictureGrayscale`, `msoPictureBlack&White`, `msoPictureWatermark`;
8. `error in`.

I test di formattazione devono isolare separatamente scale factor, width, height, measurement system, picture index e color type.

## FMT-01 — scale factor 0,5

File analizzato: `FMT-01.xlsx`

Parametri noti:

- measurement system: metric;
- scale factor: `0,5`;
- picture index: `-1` (ultima immagine);
- height/width: `-1`;
- color type: `msoPictureAutomatic`.

### Risultato

- dimensione originale: circa 320,008 x 200,005 px;
- dimensione risultante: circa 160,004 x 100,003 px;
- rapporto esatto osservato: 0,5 su entrambi gli assi;
- rapporto d'aspetto preservato;
- PNG embedded invariata: 724 byte e stesso SHA-256;
- ridimensionamento applicato esclusivamente ad anchor e trasformazione DrawingML;
- tipo anchor invariato: `twoCellAnchor editAs="oneCell"`;
- posizione presente nel file: `A1`.

Il scale factor e' un moltiplicatore lineare applicato a larghezza e altezza correnti. Non ricodifica il media. La posizione `A1` deve essere confermata come posizione iniziale usata durante il test; in caso affermativo, `Excel Format Image.vi` non ha riposizionato l'oggetto.

### Contratto proposto

- `scaleFactor > 0`: moltiplicare entrambe le dimensioni correnti;
- `scaleFactor = 1`: nessuna variazione di scala;
- preservare il punto di origine dell'anchor;
- ricalcolare cella finale e offset attraversando la geometria del worksheet;
- non modificare la parte media embedded.

Il comportamento per scale factor zero o negativo non e' ancora definito e non va accettato nel nuovo backend senza un requisito applicativo esplicito.

## FMT-02 / FMT-03 / FMT-03m — dimensioni esplicite e sistema di misura

### Risultati osservati

| File | Sistema | Width input | Height input | Risultato fisico |
|---|---|---:|---:|---|
| `FMT-02.xlsx` | US | 160 | -1 | 160 in x 0 |
| `FMT-03.xlsx` | US | -1 | 100 | 0 x 100 in |
| `FMT-03m.xlsx` | metric | -1 | 100 | 0 x circa 100,0125 cm |

In tutti i casi il PNG embedded resta invariato e le dimensioni sono applicate soltanto al drawing.

### Conversione delle unita'

- modalita' US: il valore e' espresso in pollici;
- modalita' metric: il valore e' espresso in centimetri;
- la conversione metrica passa attraverso le unita' Office/punti e introduce un piccolo arrotondamento: input 100 cm -> circa 100,0125 cm nel package.

### Comportamento del sentinel -1

Quando una dimensione viene impostata a un valore positivo e l'altra resta `-1`, il valore negativo non preserva la dimensione corrente: Excel porta la dimensione corrispondente a zero:

- `height=-1` -> altezza zero in FMT-02;
- `width=-1` -> larghezza zero sia in FMT-03 US sia in FMT-03m metric.

Quando invece entrambe le dimensioni restano a `-1`, come verificato in FMT-PI, la VI non modifica la geometria. Il bug e' quindi nella gestione del caso parziale: il ramo di ridimensionamento viene attivato da una dimensione positiva e applica erroneamente anche il valore `-1` dell'altra.

### Decisione proposta per il nuovo backend

Il nuovo backend dovrebbe implementare la semantica dichiarata e utile, non il collasso a zero osservato:

- valore `< 0`: non modificare quella dimensione;
- valore `>= 0`: convertire dalla misura selezionata;
- rifiutare una dimensione esplicita uguale a zero, salvo requisito contrario;
- documentare che la correzione e' una differenza intenzionale rispetto al bug NI osservato.

Per verificare la conversione senza coinvolgere il sentinel serve un test con entrambe le dimensioni positive e realistiche, ad esempio metric `width=4,233`, `height=2,646`.

## FMT-04 — width e height metriche esplicite

File analizzato: `FMT-04.xlsx`

La geometria risultante corrisponde a ingressi metrici `width=160`, `height=100`:

- larghezza: 57.607.200 EMU = 4.536 pt = 160,02 cm = 6.048 px logici;
- altezza: 36.004.500 EMU = 2.835 pt = 100,0125 cm = 3.780 px logici;
- rapporto: 1,6;
- posizione iniziale: `A1`;
- PNG embedded invariata.

### Algoritmo di conversione NI osservato

La VI non converte direttamente centimetri in EMU. Converte prima la misura in punti Office interi e Excel salva poi la trasformazione:

```text
US:     points = round(inches * 72)
metric: points = round(centimeters / 2,54 * 72)
EMU:    points * 12.700
```

Le evidenze sono esatte:

- 160 cm -> circa 4.535,43 pt -> 4.536 pt;
- 100 cm -> circa 2.834,65 pt -> 2.835 pt;
- 160 in -> 11.520 pt;
- 100 in -> 7.200 pt.

Il nuovo backend puo' replicare questa quantizzazione per compatibilita' geometrica 1:1. Una conversione diretta cm->EMU sarebbe piu' precisa, ma produrrebbe differenze fino a circa mezzo punto rispetto ai file NI.

Il test con valori piu' piccoli non e' indispensabile per dedurre la formula, ma puo' essere usato come regression test visivo.

## FMT-GS — msoPictureGrayscale

File analizzato: `FMT-GS.xlsx`

### Risultato

- parte media `image1.png` invariata byte-per-byte;
- stesso hash SHA-256 della sorgente;
- geometria e dimensione invariate;
- stessa relazione embedded/esterna;
- aggiunta di un unico elemento DrawingML: `<a:grayscl/>` come figlio di `a:blip`.

Il grayscale e' quindi un effetto di rendering non distruttivo. Non richiede decodifica o ricodifica dell'immagine.

### Contratto proposto

- `msoPictureGrayscale`: aggiungere `a:grayscl` al blip selezionato;
- `msoPictureAutomatic`: rimuovere le trasformazioni colore applicate dalla funzione;
- mantenere invariata la parte media, consentendo a piu' picture basate sulla stessa PNG di avere effetti colore differenti;
- caratterizzare separatamente Black&White e Watermark soltanto se risultano realmente usati dalle applicazioni.

`msoPictureMixed` rappresenta tipicamente uno stato misto restituito da Office piu' che una trasformazione concreta; il nuovo backend non dovrebbe accettarlo come comando senza un requisito o un test specifico.

## FMT-PI — selezione tramite picture index

File analizzato: `FMT-PI.xlsx`

Configurazione:

- immagine: `apple.png`, PNG 320 x 311 px;
- prima inserzione: `A1`;
- seconda inserzione: `H7`;
- formato della prima immagine: picture index `0`, metric, Automatic;
- formato della seconda immagine: picture index `1`, metric, Grayscale;
- height/width lasciate a `-1` per entrambe.

### Risultato

- una sola `ImagePart` embedded, riutilizzata dai due oggetti;
- dimensione nativa 320 x 311 px preservata per entrambi;
- indice `0` seleziona `Picture 2` in `A1`;
- indice `1` seleziona `Picture 4` in `H7`;
- `a:grayscl` presente soltanto sul blip dell'indice `1`;
- l'effetto colore appartiene al singolo picture/anchor e non alla parte media condivisa.

### Convenzione picture index

- indici zero-based;
- ordine uguale all'ordine di inserimento nel worksheet/drawing;
- `-1` seleziona l'ultima immagine;
- un indice esplicito seleziona un solo oggetto, anche quando piu' oggetti condividono la stessa `ImagePart`.

Il backend deve validare l'intervallo e restituire un errore deterministico per indici minori di `-1` o maggiori/uguali al numero delle immagini.

## IMG-01 — inserimento predefinito

File analizzato: `IMG-01.xlsx`

### Risultato geometrico

- Worksheet modificato: `Sheet1`
- Tipo DrawingML: `xdr:twoCellAnchor`
- Attributo: `editAs="oneCell"`
- Anchor iniziale: colonna 0, riga 0, offset X=0, Y=0 (`A1`)
- Anchor finale: colonna 5, riga 10
- Offset finale X: 76 EMU, circa 0,008 px
- Offset finale Y: 95.298 EMU, circa 10,005 px
- Estensione trasformazione: `cx=3.048.076`, `cy=1.905.048` EMU
- Dimensione risultante: circa 320,008 x 200,005 px
- Rapporto d'aspetto: preservato
- DPI locale: disabilitato (`a14:useLocalDpi val="0"`)

Il toolkit usa quindi la dimensione pixel originale a 96 DPI logici e ignora la densita' dichiarata di 300 DPI. Le differenze fra dimensione nominale e dimensione XML sono inferiori a 0,01 pixel e derivano dall'arrotondamento delle coordinate EMU.

Il comportamento `editAs="oneCell"` significa che l'immagine si sposta con la cella iniziale, ma non viene ridimensionata quando cambiano le dimensioni delle celle attraversate.

### Contenuto immagine

La copia incorporata in `xl/media/image1.png` e' byte-per-byte identica a `prova.png`:

- dimensione: 724 byte;
- SHA-256: `4B2C54ECC20EB1D0158A6B62C953854CBD54D01B0A138F44C7BB35327BDB87EB`.

Il toolkit non ricodifica ne' ridimensiona fisicamente il file PNG.

### Relazioni

Il drawing contiene contemporaneamente:

- relazione embedded verso `../media/image1.png`;
- relazione esterna verso il path assoluto della PNG sorgente;
- `a:blip` con entrambi gli attributi `r:embed` e `r:link`.

La copia embedded rende comunque visualizzabile l'immagine, ma il link assoluto introduce una dipendenza ambientale e rivela il path della macchina di generazione. Il nuovo backend dovrebbe, salvo necessita' di compatibilita' non ancora osservate, creare soltanto la relazione embedded.

### Modifiche collaterali

Excel ha normalizzato il workbook aggiungendo:

- `xl/theme/theme1.xml`;
- `xl/sharedStrings.xml`;
- metadati applicazione/revisione;
- relazione drawing sul worksheet;
- content type PNG e drawing.

Ha inoltre portato l'altezza standard delle righe di `Sheet1` a 14,4 pt. Queste modifiche sono dovute all'apertura/salvataggio tramite Excel e non all'immagine in senso stretto.

## Conclusione provvisoria

Per riprodurre IMG-01 il backend Open XML deve:

1. incorporare il file originale senza ricodifica;
2. creare un `TwoCellAnchor` con `EditAs=OneCell`;
3. usare `A1` e offset iniziali nulli;
4. convertire 320 x 200 px in EMU usando circa 9.525 EMU/px;
5. calcolare la cella finale attraversando larghezze colonna e altezze riga effettive;
6. impostare `useLocalDpi=false`;
7. evitare preferibilmente la relazione esterna al file sorgente.

I test successivi devono determinare se la cella iniziale e le dimensioni possano essere controllate dagli ingressi dei VI oppure derivino dal comportamento di append/formattazione.

## IMG-03 — posizione C5

File analizzato: `IMG-03.xlsx`

Il file e' stato rigenerato dopo la correzione delle coordinate. La prima esecuzione, che produceva `E3`, e' considerata superata.

### Confronto con IMG-01

- Worksheet modificato: `Sheet1`
- Tipo DrawingML invariato: `xdr:twoCellAnchor editAs="oneCell"`
- Ingressi NI: `(riga=4, colonna=2)`
- Anchor iniziale: colonna 2, riga 4, offset X=0, Y=0 (`C5`)
- Anchor finale: colonna 7, riga 14
- Offset finale X: 76 EMU, circa 0,008 px
- Offset finale Y: 76.248 EMU, circa 8,005 px
- Estensione invariata: `cx=3.048.076`, `cy=1.905.048` EMU
- Dimensione invariata: circa 320,008 x 200,005 px
- Media incorporato: lo stesso PNG da 724 byte

La posizione iniziale coincide esattamente con l'angolo superiore sinistro di `C5`: non sono presenti offset residui. La diversa coordinata finale e' soltanto la conseguenza dell'attraversamento delle righe/colonne necessario a rappresentare la stessa dimensione assoluta di 320 x 200 px.

La convenzione dell'API NI e' stata confermata: coordinate numeriche nell'ordine `(riga, colonna)`, entrambe con indice base zero. Pertanto:

- `(0, 0)` corrisponde ad `A1`;
- `(2, 4)` corrisponde ad `E3`, risultato osservato nella prima esecuzione poi sostituita;
- `(4, 2)` corrisponde a `C5`, risultato definitivo di IMG-03.

Questa convenzione coincide direttamente con gli indici `row` e `col` memorizzati negli anchor DrawingML, quindi il backend non deve applicare conversioni `+1/-1` durante la creazione dell'anchor. La conversione in notazione A1 serve soltanto per diagnostica e messaggi destinati all'utente.

## IMG-04 — allineamento RIGHT

File analizzato: `IMG-04.xlsx`

Il test mantiene posizione `C5`, stessa PNG e stessi parametri di IMG-03, modificando l'allineamento in `RIGHT`.

### Confronto con IMG-03

Non sono state osservate differenze funzionali:

- stesso `xdr:twoCellAnchor editAs="oneCell"`;
- stesso anchor iniziale `C5`, offset nulli;
- stesso anchor finale e relativi offset;
- stessa estensione 320 x 200 px;
- stessa trasformazione e posizione assoluta;
- stesso file PNG incorporato e stesso hash;
- stesse relazioni embedded/esterna;
- nessuna proprieta' DrawingML aggiuntiva relativa all'allineamento.

L'unica differenza nel drawing e' il valore casuale `a16:creationId`, rigenerato da Excel e privo di significato geometrico.

### Conclusione

L'enum di allineamento non ha effetto osservabile sul backend Excel di `Append Image to Report (string wrap).vi`, almeno nel confronto fra il valore usato in IMG-03 e `RIGHT`. Il nuovo backend Excel puo' accettare il parametro per compatibilita' del connector pane e ignorarlo. Non e' necessario replicare test analoghi per tutti i valori dell'enum, salvo evidenze applicative contrarie.

## IMG-05 — caption/stringa alternativa

File analizzato: `IMG-05.xlsx`

Il test mantiene posizione, immagine e geometria di IMG-03 e valorizza la caption/stringa alternativa.

### Confronto con IMG-03

- drawing e geometria invariati;
- worksheet invariato;
- shared strings identiche byte-per-byte;
- PNG embedded identica;
- relazioni drawing identiche;
- nessun attributo `descr` o `title` su `xdr:cNvPr`;
- nome oggetto ancora `Picture 2`;
- nessun testo della caption rilevato nelle parti XML o nelle relazioni.

L'unica variazione nel drawing e' il consueto `a16:creationId` casuale.

### Conclusione

La caption/stringa alternativa e' ignorata dal backend Excel della VI e non viene memorizzata come testo alternativo DrawingML. Il nuovo backend puo' ignorare questo ingresso per compatibilita' stretta con il comportamento osservato. In alternativa, una futura estensione potrebbe valorizzare `xdr:cNvPr/@descr` per migliorare accessibilita' e diagnostica, ma sarebbe un comportamento migliorativo e non una replica 1:1.

## IMG-06 — posizione tramite stringa cella

File analizzato: `IMG-06.xlsx`

Configurazione del test:

- coordinate numeriche: `(riga=0, colonna=0)`;
- stringa cella: `H7`.

### Risultato

- anchor iniziale DrawingML: colonna 7, riga 6;
- posizione Excel: `H7`;
- offset iniziali X/Y: zero;
- dimensione e geometria immagine invariate;
- anchor finale: colonna 12, riga 16, con gli offset attesi;
- nessuna proprieta' dell'immagine conserva la stringa cella.

### Conclusione

La stringa cella usa la normale notazione A1 one-based e viene convertita negli indici DrawingML zero-based:

- `H7` -> `(row=6, col=7)`.

Quando la stringa cella e le coordinate numeriche sono entrambe valorizzate, la stringa ha precedenza e sostituisce completamente la posizione numerica. IMG-06 copre quindi anche il previsto test di precedenza con valori discordanti; non e' necessario un IMG-07 separato per questo scopo.

Contratto proposto per il nuovo backend:

1. se la stringa cella e' non vuota, validarla e usarla;
2. altrimenti usare la coppia numerica zero-based;
3. se la stringa non e' un riferimento A1 valido, restituire un errore deterministico senza ripiegare silenziosamente sulle coordinate numeriche.

## IMG-07 — due inserimenti della stessa immagine

File analizzato: `IMG-07.xlsx`

Il workbook contiene due immagini ottenute inserendo due volte `prova.png`, nelle posizioni `H7` e `A1`.

### Media e relazioni

- una sola parte media: `xl/media/image1.png`;
- dimensione: 724 byte;
- hash identico alla PNG sorgente;
- una sola relazione embedded `rId1` verso `../media/image1.png`;
- una sola relazione esterna `rId2` verso il path sorgente;
- entrambi gli oggetti riutilizzano la stessa coppia `r:embed="rId1"` / `r:link="rId2"`.

Il toolkit/Excel deduplica quindi automaticamente inserimenti ripetuti dello stesso file all'interno del medesimo drawing.

### Oggetti DrawingML

Il drawing contiene due `twoCellAnchor editAs="oneCell"`:

1. posizione `H7`, `cNvPr id="3"`, nome `Picture 2`;
2. posizione `A1`, `cNvPr id="5"`, nome `Picture 4`.

Entrambi mantengono la dimensione di circa 320 x 200 px. Gli ID e i nomi sono univoci ma non consecutivi: Excel usa progressioni distinte e lascia intervalli fra gli identificatori visibili. Il nuovo backend deve garantire l'univocita', ma non e' necessario replicare esattamente la numerazione cosmetica se Excel accetta e preserva i valori generati.

La sequenza effettiva delle chiamate e' stata `H7`, poi `A1`, e coincide con l'ordine degli anchor nel package. Excel conserva quindi l'ordine di inserimento. Il nuovo backend puo' riprodurre il comportamento aggiungendo ogni nuovo anchor in coda al drawing; in caso di sovrapposizione, l'ultimo elemento inserito costituisce normalmente il livello visivo superiore.

### Contratto proposto

- calcolare un hash del media o mantenere una cache per sessione;
- riutilizzare `ImagePart` e relazione quando il contenuto e' identico;
- creare comunque un nuovo anchor/picture per ogni chiamata;
- assegnare ID e nomi univoci considerando anche drawing preesistenti;
- preservare l'ordine/z-order aggiungendo i nuovi anchor in coda nell'ordine delle chiamate;
- omettere nel nuovo backend la relazione esterna, salvo scelta esplicita contraria.
