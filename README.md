# DietaDash

Dashboard locale del registro alimentare, nello stile di TCGDash.
Il database **è** `Dieta_data.xlsx`: la dashboard non tiene una copia propria dei dati.

> **Prima volta?** C'è una guida passo passo per Windows e Mac in
> **[INSTALLAZIONE.md](INSTALLAZIONE.md)**, scritta per chi non ha mai
> installato niente.

## Avvio

Doppio clic su **`avvia.command`** (macOS) o **`avvia.bat`** (Windows).
Oppure da terminale:

```bash
python3 scripts/server.py
```

Si apre da solo `http://localhost:5191`. Serve Python 3 con `openpyxl`
(`avvia.command` lo installa da sé se manca). Per chiudere: `Ctrl+C`.

Se la porta è occupata: `python3 scripts/server.py 5192`.

## Il doppio mandato

I due versi passano entrambi dal workbook, che resta l'unica fonte di verità.

| Verso | Come funziona |
|---|---|
| **Pagina → Excel** | Salvi un giorno, la pagina chiama `POST /api/giorni`, il server riscrive il foglio `Dieta` con `openpyxl`. Prima di ogni scrittura finisce una copia in `data/backup/` (se ne tengono 20). |
| **Excel → pagina** | La pagina interroga `GET /api/revisione` ogni 3 secondi: è il timbro di modifica del file, poche decine di byte. Se cambia, rilegge il registro e si ridisegna. Il controllo si ripete anche quando torni sulla scheda, perché i browser rallentano i timer in secondo piano — ed è proprio lì che la scheda sta mentre modifichi l'Excel. |

Puoi tenere Excel e dashboard aperti insieme. L'unico limite è di Excel, non
della dashboard: **mentre il file è aperto in Excel su Windows il salvataggio
può fallire** perché il file è bloccato. In quel caso la dashboard mostra il
motivo e non perde quello che hai scritto nel modulo: chiudi Excel e risalva.

## Struttura

```
Dieta_data.xlsx      il database (fogli: Dieta, Misure, Config) — non nel repository
avvia.command        avvio con doppio clic, macOS
avvia.bat            avvio con doppio clic, Windows
INSTALLAZIONE.md     guida passo passo per chi parte da zero
scripts/
  dieta_excel.py     unico punto che apre il workbook: lettura, scrittura, backup
  server.py          server locale: pagina statica + API JSON
  esporta_json.py    istantanea per la copia online
web/
  index.html         impalcatura
  style.css          tema scuro, accento arancione
  app.js             dashboard: statistiche, grafici, calendario, inserimento
  dati.json          istantanea pubblicata (senza le misurazioni)
data/backup/         copie automatiche prima di ogni scrittura
```

Il workbook **non è nel repository**: contiene dati personali, misurazioni
comprese. Al primo avvio, se non lo trova, il programma ne crea uno vuoto con
la struttura giusta e puoi cominciare a riempirlo.

## La copia online

Su GitHub Pages non può esistere il server Python, quindi lì la dashboard gira
in **sola lettura** su un'istantanea del registro:

```bash
python3 scripts/esporta_json.py
```

Rigenera `web/dati.json`, poi un `git push` aggiorna il sito. L'istantanea
contiene il registro alimentare ma **non il foglio Misure**: peso e
circonferenze non escono dal tuo computer. Online la scheda *Inserisci*
sparisce e un avviso spiega che per scrivere serve il progetto in locale.

## Il workbook

Foglio **`Dieta`**, una riga per giorno, tabella `DietaDB`:

| Colonna | Contenuto |
|---|---|
| `Data` | AAAA-MM-GG, chiave del giorno |
| `Settimana` | progressivo, blocchi di 7 righe; ricalcolato a ogni scrittura |
| `Giorno` | nome del giorno, derivato dalla data |
| `Colazione` · `Pranzo_Carboidrati` · `Pranzo_Proteine` · `Cena_Carboidrati` · `Cena_Proteine` · `Merenda_Mattina` · `Merenda_Pomeriggio` | testo libero |
| `Allenamento` · `Passi` · `Sgarro` | 0 / 1 |
| `Tipo_Allenamento` | U, L, B o vuoto |
| `Note` | eventi. La parola **Check** marca un controllo |

Foglio **`Config`**: `prossimo_check` (data del prossimo controllo, modificabile
dalla dashboard) e `inizio_dieta`. Foglio **`Misure`**: le misurazioni, lasciate
com'erano.

Puoi aggiungere righe a mano in Excel: la dashboard le legge. Tieni l'ordine
delle colonne, è il contratto fra i due.

## Come sono calcolate le statistiche

- **Medie settimanali** — totale diviso `giorni / 7`. Non conta le settimane di
  calendario ma i giorni effettivamente registrati, così i due buchi dello
  storico (23–29 dicembre, 24–30 marzo) non abbassano le medie.
- **Carboidrati e proteine** — si contano pranzo e cena insieme, escludendo la
  **singola portata** da sgarro, non l'intera giornata: se la cena è stata una
  pizza, il riso del pranzo resta un pasto della dieta. Gli alimenti considerati
  sgarro sono Pizza, Sushi, Ristorante, Kebab, Curdo, Risotto, Piadina.
- **Dall'ultimo check** — dal giorno marcato `Check` nelle note fino a oggi. Per
  spostare il confine basta scrivere `Check` nella nota di un altro giorno.
- **Aderenza** — giorni senza sgarro sul totale dei giorni registrati.
- Il **grafico settimanale** copre gli ultimi 3 mesi e usa solo le settimane
  complete: mezza settimana disegnerebbe un crollo che non c'è stato, e un anno
  intero schiaccerebbe le differenze fra una settimana e l'altra.

## Pagine

- **Panoramica** — riquadri con le medie: in grande il periodo *dall'ultimo
  check*, in piccolo la media *dall'inizio*, e una percentuale che confronta i
  due. Il colore segue il significato, non il segno: meno sgarri è verde anche
  se la variazione è negativa. Poi countdown del prossimo check, andamento
  settimanale degli ultimi 3 mesi, distribuzione di carboidrati e proteine.
- **Calendario** — mese per mese dal più recente: verde dieta seguita, rosso
  sgarro, grigio senza dati, bordo arancione sui check. Clicca un giorno per
  aprirlo nel registro, dove la riga viene evidenziata.
- **Inserisci** — un giorno alla volta, con i valori già usati come
  suggerimenti. Se scrivi un alimento da sgarro la spunta si accende da sola.
- **Registro** — tutto il foglio, con ricerca. Clicca una riga per modificarla.
