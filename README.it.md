[English](README.md)

# excel-parity-check

È un pezzo di uno strumento che ho costruito al lavoro, riscritto in modo che si possa eseguire. Serve a controllare che un file Excel generato da un'app mostri gli stessi numeri che l'app ha nei suoi dati. Il file viene ricalcolato formula per formula, come fa Excel quando lo apre, e ogni cella viene confrontata con il valore atteso: basta una cella con una differenza che nessuno sa spiegare e il controllo fallisce.

## Il controllo che passava con il file sbagliato

Quello originale è uno strumento interno per i ricavi di un ente di certificazione, che ho costruito da solo, da dipendente, lavorando quasi sempre in sessioni di Claude Code. Qui ne ho riscritto una parte su dati inventati, in modo che si legga e si provi in pochi minuti. Del codice originale, della struttura dei suoi file e dei suoi dati non c'è niente.

Prima, i ricavi si seguivano ogni settimana su grandi cartelle Excel condivise: ogni responsabile d'area scriveva i suoi numeri nel proprio foglio e le formule sommavano tutto tra fogli e file. Adesso i numeri si inseriscono in una griglia web e fa fede il database. I file Excel però si leggono ancora, quindi l'app li genera con la struttura di prima. Ogni settimana bisogna sapere se il file generato dice le stesse cose del database.

La prima versione del controllo confrontava solo le celle di input, cioè i numeri che l'app scrive nel file. Erano giusti, quindi il controllo passava, ma il file era sbagliato: in un foglio la riga del totale del mese non aveva la formula. Gli input c'erano e niente li sommava, quindi tutti i totali costruiti sopra quella riga erano troppo bassi. Nessuno se n'era accorto, perché il controllo guardava solo le celle scritte dall'app.

Non serve nemmeno leggere i valori salvati nel file. Di solito un file generato si porta dietro dei risultati vecchi accanto alle formule e chiede a Excel di ricalcolare quando lo apre, quindi quei numeri non dicono niente su cosa vedrà chi lo apre. Per saperlo bisogna ricalcolare il file da sé. Quando ho aggiunto il ricalcolo sono saltate fuori 48 celle inspiegate in un foglio, mentre i controlli di prima erano tutti verdi.

## Come il file viene ricalcolato e confrontato

```text
 expected.json                   generated .xlsx
 (stands in for the database)    (formulas + stale saved values)
        |                                   |
        |                          read with zipfile + xml
        |                                   |
        |                          formula engine: recalculates
        |                          each cell like Excel on open,
        |                          ignores the saved values
        v                                   v
  expected value for              value the user will see
  each cell (sums computed
  from the source data)
        \                                 /
         \                               /
          compare every cell, found by its row and column labels
                          |
       +------------------+-------------------+
       |                  |                   |
     MATCH              KNOWN            UNEXPLAINED
  within 1 cent    declared cause,      anything else,
                   exact amount         including a cell
                                        we cannot recalculate
                          |
          one UNEXPLAINED cell -> FAIL (exit code 1)
```

Il controllo segue quattro regole:

- Il motore sa fare solo quello che serve ai file: `SUM`, `+ - * /`, parentesi, riferimenti ad altri fogli dello stesso file. Tutto il resto (un'altra funzione, un collegamento a un altro file, un ciclo, un foglio che manca) è un errore, e la cella conta come inspiegata. Leggerla come zero nasconderebbe gli errori che il controllo deve trovare.
- Righe e colonne si trovano dalle etichette. Se un'etichetta cambia o una riga sparisce, le sue celle finiscono tra le inspiegate, mai saltate.
- Una differenza nota deve dire perché c'è e di quanto è. Vale per una cella solo se lo scarto è esattamente quello dichiarato: se nella stessa cella si aggiunge un errore nuovo, gli importi non tornano più e il controllo fallisce. Nello strumento originale l'elenco deve restare corto, e una differenza ci entra solo dopo che qualcuno ne ha trovato la causa.
- I valori attesi non vengono dal file. Il totale atteso lo calcolo dai dati di partenza. Se lo prendessi dal file, una formula sbagliata supererebbe il proprio controllo.

## Provarlo in 5 minuti

Basta Python 3.9 o più recente (l'ho provato con il 3.13), senza installare nient'altro.

```bash
cd excel-parity-check

# run the tests (standard library unittest)
python -m unittest

# a correct file: passes, and shows one declared difference as KNOWN
python -m parity_check examples/good.xlsx

# the real incident, rebuilt: one month total is empty
python -m parity_check examples/empty_subtotal.xlsx

# the same file with the old check, inputs only: green
python -m parity_check examples/empty_subtotal.xlsx --inputs-only
```

Ecco cosa stampa il controllo che fallisce:

```text
UNEXPLAINED (8):
  Seniors!C6       Events / Month total: expected 4,000, file shows 0
  Seniors!E6       Total / Month total: expected 14,000, file shows 8,000
  Seniors!C8       Events / Variance vs budget: expected -1,000, file shows -5,000
  Seniors!E8       Total / Variance vs budget: expected 1,000, file shows -5,000
  Club total!C6    Events / Month total: expected 7,000, file shows 3,000
  Club total!E6    Total / Month total: expected 25,000, file shows 19,000
  Club total!C8    Events / Variance vs budget: expected -2,000, file shows -6,000
  Club total!E8    Total / Variance vs budget: expected 0, file shows -6,000

FAIL: 84 cells checked, 68 match, 8 known, 8 unexplained
```

Il comando esce con 0 se il risultato è PASS, con 1 se è FAIL e con 2 se non riesce a partire, quindi si può usare in uno script o in una CI.

Negli esempi c'è il budget di un mese di un circolo di canottaggio amatoriale che non esiste. I fogli d'area sono due (`Youth`, `Seniors`), più un foglio `Club total` che li somma. In ogni foglio ci sono tre voci di budget e una colonna di totale, quattro righe per le settimane, il totale del mese, il budget e lo scostamento dal budget.

| File | Cosa c'è di sbagliato | Risultato |
|---|---|---|
| `good.xlsx` | niente, a parte un pagamento di uno sponsor registrato in ritardo (dichiarato in `known_differences.json`) | PASS |
| `empty_subtotal.xlsx` | una cella del totale del mese è vuota | FAIL, 8 celle |
| `wrong_reference.xlsx` | una cella del foglio dei totali punta alla settimana sbagliata | FAIL, 2 celle |
| `unsupported_function.xlsx` | un totale del mese usa `ROUND`, che il motore rifiuta | FAIL, 8 celle |

In tutti e quattro i file i valori salvati sono quelli del file corretto: uno strumento che si fidasse dei valori salvati li farebbe passare tutti. Con `python examples/make_examples.py` si ricostruiscono da `expected.json`, e un test verifica che i file ricostruiti siano uguali a quelli nel repository.

```text
parity_check/
  xlsx.py        read and write .xlsx with zipfile and xml.etree
  formula.py     formula engine (SUM, arithmetic, sheet references)
  layout.py      shape of the example sheets, builds a correct copy
  gate.py        expected values, known differences, classification, verdict
  __main__.py    command line
examples/        expected.json, known_differences.json, four .xlsx files
tests/           unittest suite, written before the code
```

## Cosa resta nello strumento interno

Il controllo che uso al lavoro è più grande. Ho tolto tutto quello che descrive lo strumento originale o tocca i suoi dati:

- l'accesso al database e il codice che genera i file originali (qui il database è un file JSON e l'esportazione la fa `layout.py`);
- il modello dei fogli originali: i nomi, le etichette che riconosce, le coordinate, quale riga del database va in quale colonna. La struttura dell'esempio è inventata e molto più ridotta;
- un secondo controllo che confronta i totali dentro il database, prima ancora di aprire il file;
- tutto quello che sta intorno nell'app: login, regole di accesso, la griglia web, l'import dei file vecchi, i job pianificati, il deploy e la configurazione.

Anche il motore è ridotto a quello che serve all'esempio. Quando trascini una formula verso il basso, Excel scrive delle "formule condivise". Il lettore non le espande: segnala quelle celle e il motore le rifiuta.

## Cosa dice la storia dello strumento interno

Li ho contati il 20 novembre 2026 sul repository e sui log del progetto. Lo strumento non è pubblico.

| Cosa | Valore |
|---|---|
| Commit sul ramo principale | 246, dal 12 maggio al 2 novembre 2026 |
| Commit con Claude come co-autore | 191 su 246 |
| File di test nel progetto | 261 |
| Casi di test solo per questo controllo | 97 |
| Celle confrontate in un'esecuzione del controllo su un flusso di dati | circa 147 mila uguali, 0 inspiegate |
| Celle trovate dal ricalcolo quando l'ho aggiunto | 48 inspiegate in un foglio, con i controlli di prima verdi |
| Tempo risparmiato dagli utenti, errori evitati in produzione | non misurato |

## Le regole che ho dato all'agente

Quasi tutto il progetto originale l'ho scritto in sessioni di Claude Code, e all'agente ho dato delle regole da rispettare. Il file di regole dice: prima i test, un risultato tipizzato per ogni azione che cambia i dati, il record di audit scritto nella stessa transazione del database della modifica. Ogni decisione la scrivo prima del codice, con un pre-mortem. Se si rivela sbagliata, aggiungo una voce nuova che la sostituisce e spiega perché.

Per molti task lavorava un subagente nuovo, poi c'erano due revisioni: rispetta la specifica? il codice è fatto bene? Una revisione dell'architettura l'ho fatta fare a un modello diverso e ne ho valutato i punti uno per uno. Gli agenti avevano confini scritti (file permessi e file vietati) e le sessioni in parallelo si passavano il contesto con file condivisi.

Il controllo descritto sopra è nato così. Anche questo estratto l'ho scritto con Claude Code, da zero, partendo dai test.

## Limiti

- Il motore conosce solo `SUM`, `+ - * /` e i riferimenti dentro lo stesso file. Nelle cartelle di lavoro di tutti i giorni serve spesso altro (`IF`, `ROUND`, funzioni di ricerca). La regola non cambia: quello che il motore non sa ricalcolare conta come inspiegato.
- Le formule condivise vengono rifiutate: un file salvato a mano in Excel di solito fallisce, finché il lettore non imparerà a espanderle. I file generati da codice invece di solito scrivono ogni formula per intero.
- La struttura è fissa: etichette nella colonna A, intestazioni nella riga 1. Il controllo originale ha un modello dei suoi fogli più ricco.
- Si controllano solo i valori: formati numerici, stili e righe nascoste no.
- La tolleranza è di un centesimo, in valore assoluto.
- Una cella di testo usata in un calcolo semplice (=A1+1) qui vale 0. Excel mostrerebbe un errore.

## Licenza

MIT, vedi `LICENSE`. Autore: Lorenzo Algeri.
