# DR Portal — Manuale della sezione Finanza

*Aggiornato alla versione 1.9.0 (build 2026.08.05).*

La sezione Finanza è riservata alla direzione (e agli amministratori). Mette insieme tre mondi che in azienda vivono separati: **l'estratto conto bancario**, **le fatture** (emesse e ricevute) e **gli incassi/pagamenti registrati sul gestionale di fatturazione**. Ogni scheda guarda questi dati da un'angolazione diversa.

---

## La barra in alto

- **Saldo contabile** — il saldo comunicato dalla banca via collegamento PSD2, aggiornato all'ultima sincronizzazione. È il saldo *contabile* (la banca non espone il disponibile).
- Il numero di **versione** dell'app è in fondo alla pagina: prima di qualunque import controlla che sia l'ultima — un import fatto con una versione vecchia può applicare regole superate.

---

## Le schede, una per una

### 1 · Movimenti — l'estratto conto

L'archivio dei movimenti bancari, riga per riga: data, importo, **saldo progressivo**, tipologia (assegnata in automatico dalla causale), cliente/fornitore, numero fattura citato, note.

- **Filtri**: tipologia, mese, cliente/fornitore, più la **ricerca libera** — che guarda anche dentro la *descrizione* del movimento: se una riga sembra "vuota" ma compare cercando un nome, accanto alla controparte vedi lo stralcio di descrizione che ha fatto scattare la corrispondenza.
- **Importa estratto conto**: carica l'export xlsx dell'home banking. I doppioni vengono riconosciuti e scartati: ricaricare lo stesso file non fa danni.
- **Esporta CSV**: l'elenco filtrato, pronto per Excel.
- **Icona bacchetta** su una riga: precompila una nuova Regola (vedi scheda Regole) partendo da quel movimento.

**Quando usarla**: per cercare un movimento preciso, correggere una classificazione, caricare l'estratto.

### 2 · Overview incassi — il colpo d'occhio mensile

Tabella pivot **cliente × mese** dei soli incassi (o, in modalità Spese, delle uscite per controparte/tipologia). Risponde a "quanto ci ha versato X mese per mese?" guardando **solo la banca** — le fatture qui non c'entrano.

### 3 · Resoconto — banca contro fatturazione

Il controllo del cerchio: **estratto conto − incassato attive + pagato passive = 0**. Se torna, badge verde OK; se no, la Differenza dice di quanto.

- **Riquadro principale**: *Estratto conto* (somma dei movimenti delle controparti scelte) · *Netto fatture* (incassato attive − pagato passive, secondo la fatturazione) · *Differenza*. Sotto, il dettaglio di attive e passive.
- **I tre filtri**:
  - **Cliente** e **Fornitore** scelgono le controparti dei due lati fatture. L'opzione **Nessuno** esclude un lato intero: indispensabile per chi è *solo* fornitore o *solo* cliente, altrimenti il lato lasciato su "Tutte" trascina dentro l'intera azienda.
  - **Estratto conto**: normalmente in automatico ("Segue Cliente/Fornitore"); una selezione esplicita decide direttamente quali movimenti contare.
- **Le due liste In ritardo** (da incassare / da pagare) mostrano le fatture scadute e non saldate delle controparti scelte; le **Fasce di ritardo** le filtrano per anzianità (e non toccano il riquadro dei totali).
- Senza alcun filtro il controllo è su **tutta l'azienda**: qui è normale vedere una differenza, perché in banca arrivano incassi non ancora registrati sul gestionale (e viceversa). Il Resoconto dà il meglio **per singola controparte**.

### 4 · Fatture attive / Fatture passive

L'archivio fatture, alimentato dagli export del gestionale (ZIP XML, report, lista movimenti). Ogni fattura ha **tre letture indipendenti** dello stato d'incasso — nessuna "vince", si confrontano:

1. **Fatturazione** — lo stato scritto sul gestionale (Incassata / Non incassata…).
2. **Incassi** — la somma delle *rate* dal report movimenti del gestionale: è l'unica lettura che quantifica gli **incassi parziali**.
3. **Banca** — cosa risulta dagli abbinamenti con l'estratto conto.

Quando due letture non concordano, la fattura è marcata **discordante**: è un segnale da guardare, non un errore del portale.

Altre cose che la scheda fa:

- **Scadenze e ritardi** — per le attive la scadenza viene *sempre* dai termini di pagamento per cliente (pannello Termini; chi non è in elenco = 30 giorni). Per le passive vale la scadenza scritta in fattura; se manca, i termini del fornitore. I termini possono avere **regole per oggetto**: parole chiave (es. *locazione, affitto*) che cambiano i giorni solo per le fatture il cui oggetto le contiene — così per lo stesso fornitore convivono un termine generico e uno "a vista".
- **Note di credito e storni** — una NC collegata abbatte il residuo della fattura rettificata. Se la NC non dichiara la fattura, il portale la **aggancia da solo** quando esiste *una sola* fattura della stessa controparte con lo stesso identico importo; nei casi ambigui resta il collegamento manuale dal dettaglio della NC.
- **Filtri sulle intestazioni** — ogni colonna ha l'imbuto stile Excel, con ricerca e conteggi; i filtri si combinano a cascata.
- **Griglia di classificazione** (solo passive) — doppio click su *Competenza / Tipologia / Cliente rif.* per correggere al volo; le regole della lista RegoleFatture e lo storico propongono i valori per le fatture nuove.
- **Sollecito** — per un cliente in ritardo genera il testo del sollecito (copia negli appunti o email precompilata all'indirizzo del pannello Termini), incluse NC e fatture loro da compensare.
- **Estratto per cliente** con la riga "Effettivo (dalla banca)" e il solutore **Spiega bonifici** per i pagamenti cumulativi con compensazioni.
- **Scartate** — i documenti rifiutati dallo SdI e i reinvii sono esclusi da ogni conteggio; il chip dedicato li mostra.

**Cosa si importa qui** (vale per entrambe le direzioni, il file viene smistato da solo):

| File | A cosa serve | Quando |
|---|---|---|
| ReportMovimenti dell'anno corrente (xls) | Incassi/pagamenti per rata — aggiorna lo stato di incasso | Ogni giorno o quasi |
| ZIP XML / export fatture | Fatture nuove (e l'oggetto per le regole) | Quando ci sono fatture nuove |
| Report fatture ricevute classificato | Competenza/tipologia/cliente rif. delle passive | Quando l'amministrazione lo aggiorna |
| Foglio contratti (GG pagamento) | Termini di pagamento in blocco | Raramente |

**Regola d'oro degli import**: il ReportMovimenti si esporta sempre **ad anno intero, senza filtri di data** (il gestionale tronca gli export lunghi: mai un export multi-anno). Se compare la domanda di conferma su riduzioni/azzeramenti: OK solo se il file è l'anno completo; **nel dubbio ANNULLA — non si perde mai nulla**.

### 5 · Anomalie

I movimenti che la classificazione automatica non ha saputo interpretare fino in fondo. Si correggono qui (tipologia/controparte) e la correzione può diventare una Regola con la bacchetta.

### 6 · Storico estratti

L'elenco degli import dell'estratto conto, con la possibilità di **annullare un import** intero (a blocchi). Utile se si è caricato un file sbagliato.

### 7 · Regole

Le correzioni permanenti alla classificazione dei **movimenti bancari**: "quando un movimento contiene questo testo, chiamalo così / classificalo così". Valgono per ogni import futuro e, spuntando la casella, anche per l'archivio già caricato. Esempi in uso: `amazon` → AMAZON; descrizione con `commission` → BANCA BPM.

*(Da non confondere con le Regole delle fatture passive — pannello RegoleFatture — che assegnano tipologia di costo e cliente di riferimento alle fatture di un fornitore.)*

---

## I pannelli di impostazione

- **Termini di pagamento** — giorni contrattuali per controparte, separati **Clienti** (quando ci pagano) e **Fornitori** (quando paghiamo noi); campo Email per i solleciti; parole chiave sull'oggetto per le regole speciali; tasto per copiare i termini dei clienti sui fornitori omonimi. Le varianti di nome (CEVA / CEVA LOGISTICS…) sono riconosciute dal nome breve: basta una riga.
- **Regole fatture (passive)** — fornitore → tipologia di costo / cliente di riferimento.
- **Banca (PSD2)** — il collegamento al conto: sincronizzazione movimenti e saldo, consenso da rinnovare ~ogni 89 giorni, URL per l'automazione degli aggiornamenti programmati.

---

## Le routine

**Quotidiana** (5 minuti): ReportMovimenti dell'anno corrente → import; fatture nuove (ZIP XML) se ci sono; sync banca (1 click). Fine.

**Quando qualcosa non torna**: parti dal **Resoconto della controparte** → guarda le tre letture della fattura sospetta in **Fatture** → *Estratto per cliente* e *Spiega bonifici* per capire i pagamenti cumulativi.

---

## Sovrapposizioni note (e come conviverci)

| Sembrano uguali | In realtà |
|---|---|
| Overview incassi ↔ Resoconto | L'Overview guarda **solo la banca** (pivot mensile); il Resoconto **confronta** banca e fatturazione. |
| Ritardi nel Resoconto ↔ chip "In ritardo" in Fatture | Stessi dati: il Resoconto li riassume per controparte con le fasce; Fatture è l'elenco completo con tutti i filtri e il dettaglio. |
| Regole (movimenti) ↔ Regole fatture | Agiscono su archivi diversi: le prime sui movimenti bancari, le seconde sulla classificazione delle fatture passive. |
| Cliente/Fornitore nei Movimenti ↔ nelle Fatture | Nei movimenti è la controparte "vista dalla banca" (una colonna sola, entrambe le direzioni); nelle fatture la direzione è certa. |

Se in futuro una di queste coppie risultasse davvero di troppo (candidata principale: l'Overview incassi, ormai quasi coperta dal Resoconto), si può valutare l'accorpamento — meglio deciderlo dopo un po' di uso quotidiano del Resoconto.
