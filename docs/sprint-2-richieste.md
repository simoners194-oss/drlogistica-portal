# Sprint 2 — Modulo Richieste (SPECIFICA)

> **Stato: SOLO SPECIFICA — non ancora implementato.**
> Questo documento congela la progettazione concordata del modulo Richieste
> (ferie / permessi / straordinari). Serve da riferimento unico per
> l'implementazione futura. Nessun codice del modulo è ancora scritto.
>
> Data di stesura: 2026-07-07 · Contesto: segue lo Sprint 1 (colonne
> `Visibile`/`Autorizza` sulla lista Dipendenti, completato e validato).

---

## 1. Obiettivo

Consentire ai dipendenti di inviare richieste di **Ferie**, **Permesso** e
**Straordinario**, e a chi ha `Autorizza = true` (oggi solo Francesco Romano) di
approvarle/respingerle. Il modulo poggia sulla stessa integrazione SharePoint
reale già in uso (login Codice+PIN, presenze), senza dati mock.

---

## 2. Modello dati — lista SharePoint `Richieste`

Nuova lista SharePoint da creare sul sito DRPORTAL, accanto a `Dipendenti` e
`Timbrature`.

| # | Nome colonna | Tipo SharePoint | Obblig. (SP) | Note |
|---|---|---|---|---|
| 1 | `Title` | Testo (di sistema) | Sì (sistema) | Riusata come **ID leggibile** (es. `REQ-2026-0042`), generato dall'app. Non creare nulla. |
| 2 | `Richiedente` | **Lookup → Dipendenti** | Sì | Collega alla lista Dipendenti (stesso ID interno del resto del portale). Colonna mostrata: `NomeCompleto` o `Title`. |
| 2b | `CodiceRichiedente` | Testo singola riga | No | L'app copia il **codice** (es. `DR001`) dal dipendente alla creazione. Rende il codice visibile/filtrabile/esportabile senza aprire il lookup. **Dato storico congelato.** |
| 3 | `SedeRichiedente` | Scelta (`Fiano Romano`, `San Giuliano`) | No | L'app copia la sede del richiedente all'invio. Per filtro/report veloce senza join. |
| 4 | `TipoRichiesta` | Scelta (`Ferie`, `Permesso`, `Straordinario`) | Sì | **Scelta singola.** Nessun valore "Malattia". |
| 5 | `Modalita` | Scelta (`Preventivo`, `Consuntivo`) | No | Calcolata dall'app, **solo per Straordinario**. Vuota di default. |
| 6 | `DataInizio` | Data (solo data) | Sì | Inizio ferie, o giorno del permesso/straordinario. |
| 7 | `DataFine` | Data (solo data) | Sì | Fine. Per Permesso/Straordinario coincide con `DataInizio`. |
| 8 | `OraInizio` | Testo singola riga | No | Formato `HH:MM`. Usata per Permesso e Straordinario. Testo (non tipo Ora) per semplicità. |
| 9 | `OraFine` | Testo singola riga | No | Come sopra. |
| 10 | `Motivazione` | Testo più righe (plain) | **No a livello SP** | Obbligatoria per Permesso/Straordinario, opzionale per Ferie — regola applicata dall'**app**. |
| 11 | `DurataGiorni` | Numero (0 decimali) | No | Calcolata dall'app (per le Ferie). Per i report. |
| 12 | `DurataOre` | Numero (1–2 decimali) | No | Calcolata dall'app (per Permesso/Straordinario). Per i report. |
| 13 | `Stato` | Scelta (`Bozza`, `Inviata`, `Approvata`, `Respinta`, `Annullata`) | Sì | **Default = `Bozza`.** |
| 14 | `DataInvio` | Data e ora | No | Momento del passaggio `Bozza → Inviata`. Scritta dall'app. Base per anzianità coda e per il calcolo delle 72h dello straordinario. |
| 15 | `Approvatore` | Lookup → Dipendenti | No | Chi ha deciso. Vuota finché non c'è decisione. Coerente con `Richiedente`. |
| 16 | `DataDecisione` | Data e ora | No | Momento dell'approvazione/rifiuto. Scritta dall'app. |
| 17 | `NoteDecisione` | Testo più righe (plain) | **No a livello SP** | Obbligatoria se `Stato = Respinta` — regola applicata dall'**app**. Motivo del rifiuto o nota di approvazione. |
| 18 | `AnnoCompetenza` | Numero (0 decimali) | No | Anno a cui pesa la richiesta per i report. Calcolato dall'app. |
| 19 | `Allegati` | (funzione nativa) | — | Non è una colonna: è la funzione **Allegati** nativa delle liste. Verificare solo che sia abilitata nelle impostazioni lista. |

### Colonne di sistema ereditate (NON crearle)
SharePoint aggiunge da solo `Created`, `Modified`, `Created By`, `Modified By`.
Costituiscono l'**audit trail gratuito** (chi ha creato/modificato ogni
richiesta e quando) — fondamentale per un modulo che tocca i diritti dei
lavoratori.

---

## 3. Trappole di configurazione SharePoint

Sono gli stessi tipi di problema che nello Sprint 1 facevano la differenza tra
"funziona" e "dashboard vuota".

1. **Valori delle scelte esatti.** Su `TipoRichiesta`, `Stato` e `Modalita`
   scrivere i valori **precisi** come in tabella (maiuscola iniziale, nessuno
   spazio). Il codice li confronterà con questi valori esatti: un `"Ferie "`
   con spazio finale o `"In bozza"` invece di `"Bozza"` rompe il filtro.
2. **`Stato` default = `Bozza`.** Impostare il valore predefinito della colonna
   Scelta su `Bozza`, così ogni richiesta nasce in bozza (comportamento corretto
   del workflow).
3. **Niente obbligatorietà SharePoint dove la regola è condizionale.**
   `Motivazione`, `OraInizio`/`OraFine`, `NoteDecisione` **NON** devono essere
   obbligatorie a livello di colonna SP: la loro obbligatorietà dipende dal tipo
   di richiesta o dallo stato. Se rese obbligatorie in SP, bloccano il
   salvataggio delle ferie. La validazione condizionale la fa l'app.

### Nota di design — perché Lookup + codice denormalizzato
`Richiedente`/`Approvatore` sono **Lookup** (non tipo "Persona"), per coerenza
con il modello Codice+PIN attuale (login, presenze e visibilità girano già sulla
lista Dipendenti con gli stessi ID). Il Lookup collega sempre all'**ID interno**
stabile: se un domani il codice `DR001` cambia, le richieste storiche restano
valide. `CodiceRichiedente` è la **denormalizzazione** del codice per
leggibilità e come fotografia storica del momento — stessa logica di
`SedeRichiedente`. Il giorno in cui arriverà l'autenticazione Microsoft reale
(Entra ID), si potrà valutare il passaggio al tipo "Persona".

---

## 4. I tre tipi di richiesta

| | Ferie | Permesso | Straordinario |
|---|---|---|---|
| Periodo | `DataInizio` → `DataFine` (intervallo) | Giorno singolo (`DataFine = DataInizio`) | Giorno singolo |
| Ore | — | `OraInizio`/`OraFine` → `DurataOre` | `OraInizio`/`OraFine` → `DurataOre` |
| Giorni | `DurataGiorni` (calcolata) | — | — |
| `Motivazione` | Opzionale | **Obbligatoria** | **Obbligatoria** |
| `Modalita` | — | — | **`Preventivo` / `Consuntivo`** |

### Regola dello Straordinario (Preventivo / Consuntivo)
- **Preventivo**: straordinario richiesto *prima* di svolgerlo.
- **Consuntivo**: straordinario dichiarato *dopo* averlo svolto.
- **Blocco hard 72h**: l'invio di un **Consuntivo** è **bloccato** se sono
  trascorse più di **72 ore** dallo straordinario. Chi non giustifica entro il
  termine perde la possibilità — nessuna deroga. (Linea rigorosa concordata.)

---

## 5. Regole di validazione (lato app, condizionali)

La validazione condizionale è responsabilità dell'**app**, non di SharePoint:

- `Motivazione` obbligatoria per **Permesso** e **Straordinario**, opzionale per
  **Ferie**.
- `OraInizio`/`OraFine` presenti e valide (`HH:MM`) per **Permesso** e
  **Straordinario**; assenti per Ferie.
- Per **Permesso**/**Straordinario**: `DataFine = DataInizio`.
- `Modalita` valorizzata **solo** per **Straordinario**.
- `NoteDecisione` obbligatoria se `Stato = Respinta`.
- `Consuntivo` straordinario: rifiutato in invio se oltre 72h (vedi §4).
- `DurataGiorni` / `DurataOre` / `AnnoCompetenza` calcolate dall'app, mai inserite
  a mano.

---

## 6. Ciclo di vita e stati

```
Bozza ──(invia)──▶ Inviata ──(approva)──▶ Approvata
                      │
                      └──(respingi)──▶ Respinta   [NoteDecisione obbligatoria]

Bozza / Inviata ──(annulla dal richiedente)──▶ Annullata
```

- **`Bozza → Inviata`**: scrive `DataInvio`, `CodiceRichiedente`,
  `SedeRichiedente`, calcola durate e `AnnoCompetenza`.
- **`Inviata → Approvata/Respinta`**: scrive `Approvatore`, `DataDecisione` e
  (per Respinta) `NoteDecisione`.

---

## 7. Routing di approvazione

- Approvano le richieste **solo** i dipendenti con `Autorizza = true` (oggi solo
  **Francesco Romano**).
- **Caso auto-approvazione** — le richieste **di Francesco** (richiedente =
  unico autorizzatore): **auto-approvazione + notifica**, con questi presidi
  obbligatori per la tracciabilità HR:
  - `Approvatore` e `DataDecisione` vanno **comunque scritti** (Francesco
    approva se stesso, tracciato) — mai lasciati vuoti.
  - `NoteDecisione` riceve un marcatore automatico, es.
    `Auto-approvazione (autorizzatore = richiedente)`, così l'auto-approvazione
    è **filtrabile** in un report di controllo.
  - La "notifica" per ora è il solo log/marcatore (niente email finché non c'è
    MSAL/Graph mail); il gancio è pronto per quando arriverà l'auth Microsoft.
- Il modulo deve interrogare gli **approvatori** con una query propria su
  `Autorizza = true`, **NON** riusando la lista dei visibili: Francesco è
  `Visibile = false` e non compare nelle viste operative filtrate.

---

## 8. Decisioni chiuse (v1)

| Tema | Decisione |
|---|---|
| **Saldo residuo ferie** | **Fuori dal v1** → sprint dedicato successivo (tocca rateo mensile, residui anno precedente, competenza). Le ferie in v1 si inviano/approvano senza controllo automatico del residuo. |
| **Approvazione parziale** (es. 3 giorni su 5) | **No in v1.** Si approva/respinge l'intera richiesta; l'eventuale parziale si gestisce con respingi + reinvio corretto. |
| **Chi approva Francesco** | **Auto-approvazione + notifica** (vedi §7), con audit trail obbligatorio. |

---

## 9. Vincoli di sicurezza (da tenere presente in implementazione)

Dall'audit tecnico: l'app **non ha ancora autenticazione/autorizzazione reali
lato server** (identità in `sessionStorage`, guardie di route solo client). Il
modulo Richieste tocca diritti dei lavoratori, quindi:

- L'autorizzazione ad approvare (`Autorizza = true`) **dovrà** essere verificata
  **lato server** nelle server function, non solo nascondendo bottoni nella UI.
- Finché l'auth server-side non esiste, questo resta un limite noto: da
  chiudere idealmente insieme all'integrazione Entra ID / MSAL prima di
  considerare il modulo pronto per la produzione HR.

---

## 10. Fuori perimetro / rimandato

- Saldo residuo ferie (sprint dedicato).
- Approvazione parziale.
- Notifiche email (richiede MSAL/Graph mail).
- Eventuale `CodiceApprovatore` denormalizzato (oggi l'approvatore è sempre
  Francesco → non serve appesantire).
- Geolocalizzazione delle richieste (non prevista; se mai introdotta, richiede
  informativa GDPR — cfr. audit).
