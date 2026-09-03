# Come installare DietaDash

Guida passo passo. Non serve saper programmare: si tratta di scaricare una
cartella e fare doppio clic su un file.

In tutto ci vogliono **circa 10 minuti**, e solo la prima volta. Dalla seconda
in poi si apre in 5 secondi.

---

## Cosa stiamo per fare

DietaDash è una pagina che mostra i grafici della tua dieta. I dati veri stanno
in un file Excel, `Dieta_data.xlsx`, che resta **sul tuo computer**: la pagina
lo legge e ci scrive dentro.

Perché una pagina possa leggere un file del computer serve un piccolo
programma che faccia da ponte. Quel programma si chiama **Python**. È gratis,
lo installiamo insieme al passo 2, e serve solo la prima volta.

---

# PARTE 1 — Scaricare il progetto (uguale su Windows e Mac)

### Passo 1.1

Apri questo indirizzo nel browser:

**https://github.com/Zando182/dietadash**

### Passo 1.2

Cerca il bottone verde in alto a destra con scritto **`< > Code`**. Cliccalo.

### Passo 1.3

Si apre un menù. In fondo c'è **`Download ZIP`**. Cliccalo.

Parte il download di un file che si chiama `dietadash-main.zip`.

### Passo 1.4

Vai nella cartella **Download** del tuo computer e trova quel file.

- **Su Windows:** clicca il file con il tasto destro → **Estrai tutto** → **Estrai**.
- **Su Mac:** fai doppio clic sul file. Si apre da solo.

### Passo 1.5

Adesso hai una cartella che si chiama **`dietadash-main`**.

Trascinala dove ti fa comodo, per esempio sulla **Scrivania**. Da adesso questa
è "la cartella del progetto".

> ⚠️ Non spostare i file dentro la cartella e non cambiargli nome: si cercano
> fra loro, e se ne sposti uno gli altri non lo trovano più.

> 📗 **Dov'è il file Excel?** Nella cartella scaricata non c'è: contiene dati
> personali e non viene pubblicato. Non è un problema — la prima volta che
> accendi DietaDash, il programma se ne crea uno vuoto da solo e tu cominci a
> riempirlo.
>
> Se invece stai spostando DietaDash da un computer a un altro e vuoi portarti
> dietro lo storico, copia il tuo `Dieta_data.xlsx` dentro la cartella del
> progetto **prima** di accendere.

---

# PARTE 2 — Installare Python

Qui le strade si dividono. **Salta la parte che non ti riguarda.**

## 🪟 SE HAI WINDOWS

### Passo 2.1 — Controlla se ce l'hai già

1. Premi il tasto **Windows** sulla tastiera (quello con la finestrella).
2. Scrivi `cmd` e premi **Invio**. Si apre una finestra nera.
3. Scrivi dentro questa parola e premi **Invio**:

   ```
   python --version
   ```

**Cosa succede adesso?**

- Se compare qualcosa come `Python 3.12.1` → **ce l'hai già! Salta al Passo 2.5.**
- Se compare un errore, o si apre il Microsoft Store → **continua qui sotto.**

### Passo 2.2 — Scaricare Python

Apri il browser e vai su:

**https://www.python.org/downloads/**

C'è un bottone giallo grande con scritto **`Download Python 3.x.x`**. Cliccalo.

### Passo 2.3 — Installare Python

Apri il file appena scaricato. Si apre una finestra di installazione.

> 🚨 **QUESTO È IL PASSAGGIO PIÙ IMPORTANTE DI TUTTA LA GUIDA.**
>
> In basso nella finestra c'è una casellina con scritto
> **`Add python.exe to PATH`**.
>
> **METTI LA SPUNTA** su quella casellina **prima** di andare avanti.
>
> Se te la dimentichi, il computer non troverà Python e niente funzionerà.
> (Se succede: nessun dramma, disinstalla Python e rifai questo passo.)

Poi clicca **`Install Now`** e aspetta.

### Passo 2.4 — Finito

Quando compare **`Setup was successful`**, clicca **`Close`**.

### Passo 2.5 — Accendere DietaDash

1. Apri la cartella del progetto.
2. Trova il file che si chiama **`avvia.bat`**.
3. **Doppio clic.**

Si apre una finestra nera che scrive qualche riga, e dopo qualche secondo si
apre da solo il browser con la dashboard. 🎉

> Se Windows mostra un cartello blu **"Windows ha protetto il PC"**:
> clicca **`Ulteriori informazioni`** e poi **`Esegui comunque`**.
> Succede con tutti i programmi non famosi, non è un virus.

**La finestra nera deve restare aperta** mentre usi la dashboard. È il ponte
verso l'Excel: se la chiudi, la pagina smette di funzionare.

---

## 🍎 SE HAI UN MAC

### Passo 2.1 — Accendi e basta

Sul Mac di solito Python c'è già. Proviamo direttamente:

1. Apri la cartella del progetto.
2. Trova il file che si chiama **`avvia.command`**.
3. Clicca il file con il **tasto destro** (o Control + clic) e scegli **`Apri`**.

> ⚠️ La prima volta **usa il tasto destro → Apri**, non il doppio clic.
> Il Mac chiederà *«Vuoi davvero aprirlo?»*: rispondi **`Apri`**.
> Serve solo la prima volta; dalle successive basta il doppio clic.

**Cosa succede adesso?**

- Si apre una finestra bianca (il Terminale), scrive qualche riga, e dopo
  qualche secondo si apre il browser con la dashboard → **hai finito! 🎉**
- Compare una finestra che dice *«è necessario installare gli strumenti per
  sviluppatori»* → clicca **`Installa`**, aspetta che finisca, poi rifai il
  Passo 2.1.
- Nella finestra c'è scritto **`Python 3 non trovato`** → vai al Passo 2.2.

### Passo 2.2 — Installare Python (solo se serve)

Apri il browser e vai su:

**https://www.python.org/downloads/**

Clicca il bottone giallo **`Download Python 3.x.x`**. Si scarica un file che
finisce per `.pkg`.

### Passo 2.3 — Installare

Fai doppio clic sul file `.pkg` e clicca sempre **`Continua`**, poi
**`Installa`**. Ti chiederà la password del Mac: è quella con cui accendi il
computer.

Quando ha finito, torna al **Passo 2.1**.

> Se il Mac dice **«avvia.command non può essere aperto perché proviene da uno
> sviluppatore non identificato»**: apri **Impostazioni di Sistema** →
> **Privacy e sicurezza**, scorri in basso e clicca **`Apri comunque`**.

**La finestra del Terminale deve restare aperta** mentre usi la dashboard. È il
ponte verso l'Excel: se la chiudi, la pagina smette di funzionare.

---

# PARTE 3 — Usare la dashboard

Se sei arrivato qui, nel browser vedi una pagina scura con dei riquadri
arancioni. Ci sono quattro schede in alto:

| Scheda | A cosa serve |
|---|---|
| **📊 Panoramica** | I numeri: quanti allenamenti a settimana, quanti passi, quanto stai seguendo la dieta |
| **📅 Calendario** | Un mese alla volta. Verde = dieta seguita, rosso = sgarro |
| **✍️ Inserisci** | Per aggiungere la giornata di oggi |
| **🗂️ Registro** | Tutti i giorni in una tabella. Clicca una riga per correggerla |

### Aggiungere una giornata

1. Clicca **✍️ Inserisci**.
2. I campi sono già riempiti con quello che mangi di solito: cambia solo quello
   che è diverso.
3. Clicca il bottone arancione **`Salva nel workbook`**.

Quello che salvi finisce **davvero dentro il file Excel**, subito.

### Il collegamento va nei due sensi

Puoi anche aprire `Dieta_data.xlsx` con Excel e scrivere lì dentro. Salvi, e
nel giro di pochi secondi la pagina si aggiorna da sola. Non devi ricaricare
niente.

In alto a destra c'è una lucina:

- 🟢 **Excel** → tutto collegato, quello che salvi va nel file.
- 🔴 **scollegato** → hai chiuso la finestra nera. Riaprila con `avvia.bat`
  (Windows) o `avvia.command` (Mac).

### Sei tranquillo: non puoi rompere niente

Ogni volta che salvi, il programma mette da parte una copia del file in
`data/backup`. Se combini un guaio, dentro quella cartella c'è la versione di
prima.

---

# PARTE 4 — Spegnere

1. Chiudi la scheda del browser.
2. Vai nella finestra nera (Windows) o bianca (Mac) e premi insieme
   **`Ctrl`** e **`C`**. Poi chiudila.

Per riaccendere: doppio clic su `avvia.bat` o `avvia.command`. Non devi
reinstallare più niente.

---

# PARTE 5 — Se qualcosa non va

| Cosa vedi | Cosa fare |
|---|---|
| La finestra nera si apre e si chiude subito | Python non è installato bene. Su Windows rifai il **Passo 2.3** ricordandoti la spunta `Add python.exe to PATH`. |
| `Python non trovato` | Non è installato, oppure su Windows manca la spunta del PATH. Rifai la Parte 2. |
| `Porta 5191 occupata` | DietaDash è già acceso da un'altra parte. Cerca fra le finestre aperte, oppure riavvia il computer. |
| `Workbook non trovato` | Manca `Dieta_data.xlsx` nella cartella del progetto, o l'hai spostato. Rimettilo lì. |
| `Il workbook è aperto in Excel` | Chiudi Excel e riprova a salvare. Quello che avevi scritto nel modulo non è perso. |
| Il browser non si apre da solo | Aprilo tu e scrivi nella barra degli indirizzi: `http://localhost:5191` |
| La pagina dice **sola lettura** | Stai guardando la copia online. Per scrivere serve il progetto sul tuo computer: rifai la Parte 1. |

---

## Domande che potresti farti

**Devo essere connesso a internet?**
Solo per scaricare il progetto e Python la prima volta. Dopo funziona tutto
senza internet: la pagina e i dati sono sul tuo computer.

**I miei dati finiscono su internet?**
No. Il file Excel resta sul tuo computer e non lo manda a nessuno. La pagina si
apre da `localhost`, che vuol dire "questo computer e basta".

**Posso spostare la cartella?**
Sì, tutta insieme e dove vuoi. Non spostare i singoli file che ci sono dentro.

**Posso usarlo dal telefono?**
Per inserire i dati no, serve il computer dov'è il file Excel. Per guardare i
grafici sì, se pubblichi la copia online (vedi `README.md`).
