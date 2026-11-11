// DR Portal — Contenuto del Manuale interno (modulo /manuale).
//
// Ogni voce risponde a UN dubbio concreto: chi cerca trova "cosa deve fare",
// non la teoria. Il testo è volutamente discorsivo, in italiano, scritto come
// lo spiegheresti a voce a un collega. Le `chiavi` sono parole in più che
// aiutano la ricerca (sinonimi, il messaggio d'errore esatto, il gergo).
//
// Per aggiungere una voce basta una entry qui: la pagina si aggiorna da sola.

export interface VoceManuale {
  id: string;
  sezione: string;
  titolo: string;
  /** Paragrafi separati da riga vuota. */
  testo: string;
  chiavi?: string;
}

export const SEZIONI_MANUALE = [
  "In generale",
  "Presenze e timbrature",
  "Richieste",
  "Movimenti",
  "Regole",
  "Anomalie",
  "Resoconto e incassi",
  "Fatture",
  "Pivot",
  "Flussi di cassa",
  "Fornitori",
  "Stipendi",
  "Fiscale",
  "Storico estratti",
  "Amministrazione",
] as const;

export const MANUALE: readonly VoceManuale[] = [
  // --- In generale -----------------------------------------------------------
  {
    id: "gen-vista-direzione",
    sezione: "In generale",
    titolo: "Chi vede i moduli della direzione",
    testo: `I moduli riservati (Finanze con Stipendi e Fiscale, Mezzi, Manuale, Novità riservate) li vedono l'amministratore di sistema (ADM001) e la "vista direzione": Francesco Romano (DR005), ADM002, Diego Gabelli (DR007), Monica Notaro (DR008, dalla 1.91.1), Lucrezia Pratesi (DR000) e Sabrina Guidarelli (DR002). È solo una vista: ferie, permessi, acquisti e supervisione restano governati dalle abilitazioni Autorizza e preposto di sede in anagrafica, e la tab Amministrazione resta di ADM001.

L'elenco sta nel codice (richieste-logic, CODICI_VISTA_DIREZIONE): per aggiungere o togliere qualcuno serve una modifica del portale, non basta l'anagrafica. Chi entra lo fa con la propria utenza di sempre, codice e PIN.`,
    chiavi:
      "vista direzione chi vede finanze mezzi manuale abilitazioni gabelli notaro pratesi guidarelli adm002 dr005 accessi",
  },
  {
    id: "gen-cambia-pin",
    sezione: "In generale",
    titolo: "Ho dimenticato il PIN (o voglio cambiarlo)",
    testo: `Dalla pagina di accesso tocca "PIN dimenticato? Cambialo con l'email". Scrivi il tuo codice dipendente (es. DR034): arriva un codice di verifica di 6 cifre all'email che l'ufficio ha registrato per te. Inserisci le 6 cifre e il PIN nuovo (4-8 cifre, due volte) e da quel momento entri col PIN nuovo.

Il codice vale 15 minuti e funziona una volta sola; se sbagli troppe volte o scade, ne chiedi un altro. Se il portale dice che non c'è un'email registrata, chiedi in ufficio di aggiungerla alla tua scheda: senza email il cambio da soli non può partire e il PIN te lo reimposta l'ufficio.

Per l'ufficio: l'email si registra nella colonna Email dell'anagrafica Dipendenti (anche via import massivo); le email partono dalla casella della segreteria tramite la coda già usata per le comunicazioni.`,
    chiavi:
      "pin dimenticato cambia pin email codice verifica otp reset password accesso non ricordo",
  },
  {
    id: "gen-ricarica",
    sezione: "In generale",
    titolo: "La pagina sembra vecchia o si comporta in modo strano",
    testo: `Dopo un aggiornamento del portale il browser a volte tiene in memoria la versione precedente, e i sintomi sono i più vari: bottoni che non rispondono, colonne fuori posto, un dato appena corretto che sembra ancora sbagliato.

Dalla versione 1.75.2 ogni pagina controlla da sola la versione (ogni 4 minuti e quando torna in primo piano) e si ricarica quando è rimasta indietro: il problema della "timbratrice muta dopo l'aggiornamento" è chiuso lì. Se una pagina ti sembra comunque strana, la ricarica forzata resta la prima mossa: Ctrl+F5 sul computer, trascina la pagina verso il basso sul telefono.`,
    chiavi:
      "ctrl f5 refresh cache aggiorna non funziona bottone bloccato client stale timbratrice muta versione vecchia",
  },
  {
    id: "gen-versione",
    sezione: "In generale",
    titolo: "Capire che versione stai usando",
    testo: `Nel menù laterale, in basso sopra "Connesso a Microsoft 365", c'è scritta la versione (per esempio DR Portal v1.75.0). Se un collega vede una cosa diversa da te, confrontate le versioni: chi ha il numero più vecchio deve solo ricaricare la pagina.

La pagina Novità elenca cosa è cambiato versione per versione: se una schermata ti sembra diversa da ieri, quasi sempre lì c'è la spiegazione.`,
    chiavi: "versione numero release novità changelog aggiornamento",
  },
  {
    id: "gen-lingua",
    sezione: "In generale",
    titolo: "È tutto in inglese",
    testo: `Le bandierine della lingua stanno in alto a destra, vicino a comandi che si usano spesso: capita di toccarle per sbaglio. Non si è rotto niente — clicca la bandiera italiana e torna tutto come prima.`,
    chiavi: "english inglese bandiera lingua traduzione",
  },
  {
    id: "gen-segnalare",
    sezione: "In generale",
    titolo: "Segnalare un problema in modo utile",
    testo: `Una buona segnalazione fa risparmiare ore: scrivi cosa stavi facendo, cosa ti aspettavi di vedere, cosa è successo invece, e allega uno screenshot. Un esempio preciso ("la fattura 220/26 di iMile mi risulta aperta ma su Aruba è pagata") vale più di dieci frasi generiche ("le fatture non tornano").

Prima di segnalare, prova la ricarica forzata (Ctrl+F5): tanti "errori" sono solo la pagina rimasta alla versione precedente.`,
    chiavi: "bug errore sbagliato non torna assistenza aiuto screenshot",
  },

  {
    id: "gen-stampa-export",
    sezione: "In generale",
    titolo: "Stampare o esportare un riepilogo delle ore",
    testo: `Dalla 1.83.0 ogni pagina con un riepilogo delle ore ha gli stessi due tasti: "Esporta CSV" e "Stampa". Li trovi in Le mie ore (il mese del dipendente), La mia sede (Turni del giorno e Ore), Gestione timbrature → Turni del giorno e Rendiconto. Il solo tasto Stampa c'è anche nella Dashboard presenze (accanto ad Aggiorna) e in Supervisione, nelle schede Richieste decise e Rimborsi, che il CSV lo avevano già (dalla 1.83.1).

"Esporta CSV" scarica un file che si apre in Excel con i dati che vedi filtrati in quel momento (giorno, dipendente, ore, timbrature, anomalie…). "Stampa" apre la stampa del browser con la sola pagina dei dati: menù, filtri e bottoni restano fuori, in testa ci sono titolo, sottotitolo e data/ora. Dal telefono la stampa si può salvare come PDF scegliendo "Salva come PDF" come stampante.

Consiglio: prima imposta i filtri (sede, periodo, dipendente), poi esporta o stampa — quello che esce è esattamente quello che vedi.`,
    chiavi:
      "stampa stampare esporta export csv excel pdf riepilogo ore le mie ore sede turni rendiconto",
  },

  // --- Presenze ---------------------------------------------------------------
  {
    id: "pre-riepilogo-presenze",
    sezione: "Presenze e timbrature",
    titolo: "Il riepilogo presenze a matrice (foglio dell'appalto)",
    testo: `Nel Rendiconto, dalla 1.88.0, c'è la scheda "Riepilogo presenze": il foglio presenze che l'ufficio compilava a mano per gli appalti, fatto dal portale. Un giorno per colonna, i dipendenti raggruppati per sede con il numero progressivo; per ognuno la riga Tot (le ore del giorno, attribuite al giorno dell'entrata anche per i notturni) e, se ne ha, la riga Not (le ore cadute nella fascia notturna 22–06); poi la colonna TOT e per ogni sede la riga "TOTALE ORE GIORNALIERE". Quando non si è lavorato compare il codice: FE ferie approvate, ML malattia, PC permesso (se ci sono anche ore lavorate il numero è sottolineato e il permesso sta nel suggerimento), NC turno non chiuso (da sanare in Gestione timbrature), SAB/DOM nei weekend vuoti.

I gruppi seguono la colonna Reparto dell'anagrafica Dipendenti (dalla 1.89.0): dentro una sede si possono distinguere sotto-sedi o reparti, e il gruppo prende il nome "appalto + reparto" come nei fogli dell'ufficio. Per Zingali: Zingali Pavia (Gaggianesi), Zingali Torino (Di Muro, Benanti), Zingali Cerro - Ufficio (Ausilio, Spediacci), Zingali Cerro - Magazzino (tutti gli altri). La Sede resta una sola ("Zingali") e continua a governare preposti e supervisione; il Reparto si assegna da Amministrazione → Appalti dipendenti incollando tre colonne (Nome Cognome, Appalto, Reparto; "-" svuota il reparto), con l'Anteprima prima di scrivere.

Usa gli stessi filtri del Rendiconto: mese intero, settimana fiscale o del mese, giorno singolo; appalto (scegliendo "Zingali" compaiono tutte le sue sedi con le suddivisioni); dipendente. La tendina Sede non c'è più dalla 1.89.3: era quasi un doppione di Appalto, che in più distingue HUB e Distribuzione a Fiano e gli uffici; la Sede resta in anagrafica per preposti, supervisione e timbratrice. "Esporta CSV" e "Stampa" danno la stessa griglia. L'arrotondamento ai 15 minuti (al più vicino, giorno per giorno) è acceso di default, come nei fogli dell'ufficio; togliendo la spunta vedi i minuti esatti. Il Rendiconto classico resta com'era, con le ore esatte.`,
    chiavi:
      "riepilogo presenze foglio appalto matrice giorni colonne tot not notturno 22 06 fe ml pc nc sab dom arrotondamento quarti d'ora 15 minuti monica zingali",
  },
  {
    id: "pre-timbrare",
    sezione: "Presenze e timbrature",
    titolo: "Timbrare: i quattro tasti",
    testo: `Entrata quando inizi, Inizio pausa quando stacchi, Fine pausa quando riprendi, Uscita quando finisci. Puoi fare più pause nello stesso giorno e anche due turni (esci e rientri: il portale ti fa timbrare una nuova entrata).

Un tocco basta. Dalla 1.82.0 la pagina risponde subito: il tasto mostra "Registrazione…", il telefono vibra, l'ultima timbratura e la timeline si aggiornano al momento; la conferma del server arriva un attimo dopo con l'orario ufficiale. Se premi di nuovo mentre gira, il secondo tocco viene ignorato, non registrato due volte. Prima della 1.82.0 il tasto restava uguale per qualche secondo e sembrava non aver preso: era quello a far premere "due o tre volte".

I tasti in basso ("Accesso rapido") portano a Le mie ore, Richieste, Manuale e Cambia PIN.

Alla prima apertura (dalla 1.82.1) la pagina mostra subito l'ultimo stato salvato sul telefono, così i tasti sono pronti mentre il server manda solo le tue timbrature: se nel frattempo un preposto ha corretto la tua giornata, la pagina si allinea da sola appena arriva la risposta. Una timbratura appena fatta non viene mai cancellata da una rilettura: se ti sembra sparita, ricarica la pagina dopo un minuto prima di ripeterla.

Il turno di notte funziona: se entri alle 22 ed esci alle 2 la giornata resta unita e le ore sono contate giuste.

Quando il portale viene aggiornato, una pagina rimasta aperta da prima non può più parlare col server. Dalla 1.89.4 la pagina Presenze se ne accorge da sola: controlla la versione ogni minuto e ogni volta che torna in primo piano, e si ricarica prima del prossimo tocco. Se premi un tasto proprio in quel momento, la timbratura non va persa: resta in coda con l'ora vera della pressione, la pagina si ricarica e la coda parte subito (nel registro compare la nota "Recuperata offline"). Prima della 1.89.4 poteva succedere il contrario: la timbratura sembrava registrata e non arrivava mai (1° ottobre a Zingali). Se una pagina ti sembra "muta", chiudila e riaprila.`,
    chiavi:
      "entrata uscita pausa timbratura turno notturno spezzato aggiornamento pagina vecchia si ricarica da sola coda recuperata offline",
  },
  {
    id: "pre-annulla",
    sezione: "Presenze e timbrature",
    titolo: "Ho sbagliato tasto",
    testo: `Sotto i tasti c'è "Annulla ultima timbratura": la toglie e ti riporta com'eri. Se te ne accorgi dopo ore, non timbrare a caso per compensare: chiedi la correzione della giornata al preposto o all'ufficio, che possono riscriverla con gli orari veri.`,
    chiavi: "annullare sbagliato errore timbratura cancellare",
  },
  {
    id: "pre-devi-entrata",
    sezione: "Presenze e timbrature",
    titolo: '"Devi prima registrare l\'entrata"',
    testo: `Il portale te lo dice quando per lui non sei in servizio. Le cause tipiche sono due: hai dimenticato l'entrata, oppure ieri è rimasto un turno aperto (entrata senza uscita) e dopo 24 ore il portale è ripartito da zero.

In entrambi i casi: timbra l'entrata adesso e segnala la giornata da sistemare al preposto. Caso inverso: se ieri hai dimenticato l'USCITA e oggi l'Entrata risulta bloccata, timbra prima l'Uscita (chiude il turno di ieri) e poi l'Entrata — la giornata di ieri la sistema il preposto con la correzione.`,
    chiavi: "devi prima registrare entrata errore uscita rifiutata non timbra",
  },
  {
    id: "pre-offline",
    sezione: "Presenze e timbrature",
    titolo: "Timbrare senza campo",
    testo: `Se il telefono è senza rete, la timbratura si mette in coda e parte da sola appena la rete torna, con l'orario giusto. Premere il tasto dieci volte non serve: ne basta una.`,
    chiavi: "offline rete internet coda timbratura non parte",
  },
  {
    id: "pre-correzioni",
    sezione: "Presenze e timbrature",
    titolo: "Correggere una giornata (per preposti e operatori)",
    testo: `Le giornate non si aggiustano cancellando timbrature a mano: si usano le correzioni, che riscrivono la giornata intera con gli orari indicati. Così resta traccia di tutto e le ore del Rendiconto tornano da sole.

La giornata è quella del TURNO: per un turno notturno (entrata 14:30, uscita 03:30) si lavora sulla card del giorno dell'entrata, e l'uscita delle 03:30 compare lì come ultimo passo — il portale la salva da solo sul giorno dopo. Basta scrivere gli orari in sequenza: quando un orario "torna indietro" rispetto al precedente, vuol dire che è scattata la mezzanotte.

Dalla versione 1.75.5 vale anche per l'inserimento SINGOLO e per "Aggiungi timbratura mancante": se sulla card del turno scrivi l'ora della notte (es. uscita 02:45), il portale la registra da solo sul giorno dopo e l'avviso verde ti dice su che giorno è finita. Lo fa solo quando l'ora scritta verrebbe prima dell'ultimo passo del turno ancora aperto: se metti già la data del giorno dopo a mano, non cambia nulla.

Prima di inserire, guarda la card: se il dipendente ha timbrato da solo, gli eventi ci sono già e vanno solo sistemati (togliere quello sbagliato, aggiungere quello mancante), non reinseriti da capo — altrimenti nascono doppioni. Fino alla 1.77 una timbratura tra mezzanotte e le 2 finiva sulla card del giorno prima: dalla 1.78.0 sta sul giorno giusto.

Turni spezzati in più pezzi (es. 9-12:30, 14-15:30, 17:30-20:30): nel Turno intero c'è "Aggiungi una pausa" — una riga di pausa per ogni stacco, quante ne servono (dalla 1.78.0). Nella correzione chiesta dal dipendente basta elencare tutti i passi in fila: entrata 09:00, inizio-pausa 12:30, fine-pausa 14:00, inizio-pausa 15:30, fine-pausa 17:30, uscita 20:30.

Regola d'oro: mai cancellare, sempre correggere. Il cestino di SharePoint tiene 93 giorni ed è già servito una volta a recuperare timbrature sparite — meglio non doverci tornare.`,
    chiavi:
      "correzione giornata preposto operatore riscrivere manuale gestione timbrature cancellare",
  },
  {
    id: "pre-anomalie-turni",
    sezione: "Presenze e timbrature",
    titolo: "Le anomalie dei turni (turno o pausa non chiusi)",
    testo: `Il tab Anomalie di Gestione timbrature elenca le giornate passate con un turno o una pausa rimasti aperti: finché non vengono chiusi, le ore di quella giornata non si conteggiano. Il giorno corrente è escluso (il turno può essere ancora in corso); i turni rimasti aperti IERI hanno il loro riquadro in cima alla pagina.

La lista copre gli ultimi 60 giorni: un'anomalia resta lì finché non la si chiude. Un turno o una pausa non chiusi NON si possono scartare (il bottone dice "solo da correggere"): scartarli nasconderebbe un turno incongruente e le ore resterebbero fuori dal conteggio per sempre. Si chiudono in due soli modi: con Correggi inserendo l'orario mancante, oppure — se il turno non c'è mai stato (entrata timbrata per sbaglio) — eliminando l'entrata spuria dai Turni del giorno. Il tasto Scarta resta solo per le informative ("giornata lunga senza stacco"), che non sono errori: scattano quando un tratto di lavoro senza pausa né uscita supera le 7 ore (dalla 1.80.3; prima erano 6, e a Cerro, dove la pausa cade dopo la sesta ora, uscivano ogni giorno). Non lasciarle invecchiare: dopo i 60 giorni spariscono dalla lista ma il buco nelle ore resta — a quel punto si sistema da Turni del giorno scegliendo la data a mano.`,
    chiavi: "anomalie turno non chiuso pausa aperta 60 giorni correggi scarta ore mancanti",
  },

  {
    id: "pre-dispositivo",
    sezione: "Presenze e timbrature",
    titolo: "Da quale dispositivo è partita una timbratura",
    testo: `Dalla versione 1.75.3 ogni timbratura registra anche l'indirizzo IP e il tipo di apparecchio (per esempio "Android · Chrome"). In Gestione timbrature basta fermarsi col mouse su una timbratura per vederli, insieme all'origine Web o Manuale.

Serve per la diagnostica: se qualcuno dice "non mi ha preso la timbratura", si vede subito da dove timbrava e con cosa. Le timbrature più vecchie non hanno il dato — c'è solo da quando la funzione esiste. La posizione GPS non viene raccolta.`,
    chiavi: "ip dispositivo apparecchio browser da dove timbrato diagnostica tooltip",
  },

  // --- Richieste ----------------------------------------------------------------
  {
    id: "ric-inviare",
    sezione: "Richieste",
    titolo: "Ferie, permessi e altre richieste",
    testo: `Dalla pagina Richieste scegli il tipo, indichi le date e invii. La richiesta arriva a chi la deve approvare e tu ne vedi lo stato in ogni momento: in attesa, approvata o rifiutata. Più la mandi in anticipo, più è facile organizzare i turni.`,
    chiavi: "ferie permesso malattia straordinario richiesta approvazione stato",
  },
  {
    id: "ric-orari",
    sezione: "Richieste",
    titolo: "Chiedere la correzione dei propri orari",
    testo: `Se nella tua giornata manca una timbratura o c'è un orario sbagliato, da "Le mie ore" puoi chiedere la correzione indicando gli orari veri. Chi approva riscrive la giornata e tu la ritrovi corretta.`,
    chiavi: "correzione orari le mie ore timbratura mancante sbagliata",
  },

  // --- Movimenti ---------------------------------------------------------------
  {
    id: "mov-fonte",
    sezione: "Movimenti",
    titolo: "Da dove arrivano i movimenti",
    testo: `Il conto BPM è collegato in sola lettura (open banking): i movimenti scendono da soli due volte al giorno. La banca però non manda tutto: i pagamenti cumulativi (per esempio le distinte stipendi) arrivano come un unico addebito senza il dettaglio dei beneficiari, e su tanti bonifici ricevuti la causale è tagliata.

Il dettaglio per beneficiario sta nel report "Esiti pagamenti" di YouBusiness, che la banca conserva circa 90 giorni: va esportato una volta al mese e caricato in Storico estratti, altrimenti si perde.`,
    chiavi:
      "banca bpm sync psd2 open banking causale notprovide distinta beneficiario esiti pagamenti",
  },
  {
    id: "mov-classificazione",
    sezione: "Movimenti",
    titolo: "Come vengono classificati i movimenti",
    testo: `Appena un movimento arriva, il portale prova ad applicare le Regole (vedi la sezione Regole). Se nessuna regola combacia, usa una classificazione di riserva ricavata dal testo della banca: la riconosci perché di solito la riga ha una tipologia generica e le allocazioni vuote.

Una riga "di riserva" non è un errore del portale: è un movimento per cui non esiste ancora una regola. La soluzione stabile è creare la regola e poi premere "Riapplica tutte" — sistemare solo la riga a mano funziona, ma al prossimo movimento uguale sei daccapo.

La controparte di un bonifico in uscita viene letta dal testo della banca dopo "FAVORE": il nome è seguito da code tecniche ("- ADD.TOT" sulla disposizione, "- ADD.SPE" sulla commissione della stessa disposizione) che il portale scarta. Se in Cliente/Fornitore compaiono due voci quasi uguali, tipo "beneficiari vari distinta" e "beneficiari vari distinta add.spe", è una coda non riconosciuta: dalla 1.80.4 tutte le code "ADD.xxx" vengono tolte, e le righe vecchie sono state riunite sotto un nome solo.`,
    chiavi:
      "euristica classificazione automatica tipologia allocazione vuota riserva regola non presa",
  },
  {
    id: "mov-matita",
    sezione: "Movimenti",
    titolo: "Correggere una riga a mano (e cosa comporta)",
    testo: `Con la matita correggi tipologia, sottocategoria, allocazioni e cliente di una singola riga. Da quel momento la riga è marcata come manuale (badge M) e nessuna regola la toccherà mai più: il lavoro fatto a mano vince sempre, anche sul "Riapplica tutte".

Usala per i casi singoli veri. Se ti accorgi che stai correggendo a mano sempre lo stesso tipo di movimento, quello è il segnale che serve una regola.`,
    chiavi: "matita modifica manuale badge M riga correggere movimento",
  },
  {
    id: "mov-filtri",
    sezione: "Movimenti",
    titolo: "Filtri, imbuti e CSV",
    testo: `Sulle intestazioni delle colonne ci sono gli imbuti come in Excel: spunti i valori che vuoi vedere e la tabella si restringe. Sopra ci sono ricerca, intervallo di date e importi, e il filtro "solo non classificate". L'export CSV scarica esattamente quello che stai vedendo, con i filtri applicati.`,
    chiavi: "filtro imbuto excel csv esporta ricerca colonna",
  },
  {
    id: "mov-sbagliato",
    sezione: "Movimenti",
    titolo: "Un movimento è classificato male: cosa fare",
    testo: `Prima domanda: c'è una regola per quella controparte? Cerca nella tab Regole. Se non c'è, creala e premi "Riapplica tutte": si sistemano quella riga e tutte le gemelle, passate e future. Se la regola c'è ma la riga è rimasta diversa, controlla nell'ordine: la riga è manuale (badge M — allora vince lei)? il termine della regola compare davvero nel testo del movimento? il segno della regola è giusto (entrate/uscite)?

Ultimo caso: la regola è nata dopo l'arrivo del movimento. È normale — le regole nuove non toccano il passato finché non premi "Riapplica tutte".`,
    chiavi: "classificato male sbagliato regola non funziona non presa euristica riapplica",
  },

  // --- Regole -------------------------------------------------------------------
  {
    id: "reg-come",
    sezione: "Regole",
    titolo: "Come si scrive una regola",
    testo: `Una regola dice: "se il nome o la descrizione contiene questo testo, classifica così". Puoi mettere più termini separati da virgola — basta che uno combaci. L'ordine delle parole non conta ("Riccardo Carlone" trova anche "Carlone Riccardo") e i termini sotto le tre lettere vengono ignorati, perché troppo generici ("ba" è contenuto ovunque).

Scegli su cosa cercare: nome della controparte, descrizione, o entrambi. Il consiglio pratico: per i fornitori usa il nome; per le operazioni bancarie (commissioni, F24, giroconti) usa un pezzo di descrizione preso pari pari dal movimento vero.`,
    chiavi: "regola nuova creare pattern termine virgola contiene esatto nome descrizione",
  },
  {
    id: "reg-segno",
    sezione: "Regole",
    titolo: "Il segno: perché un incasso non prende mai le regole di spesa",
    testo: `Ogni regola vale solo per le entrate o solo per le uscite. E c'è una regola di ferro incorporata nel portale: un movimento positivo non può mai risultare una spesa. Sugli incassi valgono solo le regole "Solo entrate"; se non ce n'è una che combacia, il movimento diventa semplicemente "Incasso".

È voluto: così un rimborso ricevuto da un fornitore non prende mai la regola di spesa di quel fornitore. Se vuoi che gli incassi di un cliente abbiano tipologia e allocazioni tue, serve una regola con segno "Solo entrate" per quel cliente.`,
    chiavi: "segno entrate uscite incasso positivo bonifico ricevuto invariante rimborso storno",
  },
  {
    id: "reg-precedenza",
    sezione: "Regole",
    titolo: "Se due regole combaciano, chi vince",
    testo: `Vince la più specifica: prima le regole sul nome esatto della controparte, poi quelle "nome contiene", poi quelle sulla descrizione; il jolly * arriva per ultimo e prende quello che resta. A parità di categoria vince la prima in elenco.`,
    chiavi: "precedenza ordine priorità jolly conflitto due regole",
  },
  {
    id: "reg-riapplica",
    sezione: "Regole",
    titolo: '"Riapplica tutte": cosa fa e quando usarla',
    testo: `Rilegge l'intero archivio movimenti e riallinea ogni riga alle regole attuali. Le righe manuali (badge M) restano intatte. Usala ogni volta che crei o correggi delle regole, o quando noti righe "di riserva" che ormai una regola dovrebbe prendere.

Dura qualche minuto e mostra l'avanzamento: lasciala finire. Alla fine dice quante righe ha aggiornato.`,
    chiavi: "riapplica tutte regole archivio aggiorna riclassifica",
  },
  {
    id: "reg-dipendenti",
    sezione: "Regole",
    titolo: "La regola unica dei dipendenti",
    testo: `I bonifici verso i dipendenti si classificano da soli come Pagamento Salario grazie a un'unica regola che confronta il beneficiario con l'anagrafica del portale, tollerando nomi invertiti e piccole differenze. Perché funzioni al meglio, l'anagrafica va tenuta aggiornata — compreso il campo Appalto, che serve anche alle sedi della tab Stipendi.`,
    chiavi: "dipendenti salario bonifico regola anagrafica appalto beneficiario",
  },
  {
    id: "reg-pulizia",
    sezione: "Regole",
    titolo: "Tenere pulite le regole",
    testo: `"Unifica doppie" fonde le regole equivalenti nate per strada. Quando aggiungi termini a una regola esistente è meglio che crearne una nuova quasi uguale. Il campo Note serve a ricordare perché la regola esiste e chi l'ha chiesta: tra sei mesi ringrazierai.`,
    chiavi: "unifica doppie note pulizia doppioni regole gemelle",
  },

  // --- Anomalie -----------------------------------------------------------------
  {
    id: "ano-cosa",
    sezione: "Anomalie",
    titolo: "Cosa sono le anomalie e come si smaltiscono",
    testo: `Le anomalie sono i movimenti su cui il portale non è sicuro: manca la tipologia, o la classificazione è arrivata dalla riserva e nessuna regola l'ha mai confermata. Non sono errori: sono lavoro in attesa.

Il modo più veloce di smaltirle è a gruppi: si raggruppano per controparte o descrizione simile, si decide la classificazione giusta per il gruppo, si crea la regola e si preme "Riapplica tutte". Le anomalie coperte dalle regole nuove spariscono da sole. Quelle informative senza rimedio si possono scartare a mano.`,
    chiavi: "anomalie contatore rosso da verificare smaltire gruppi scartare",
  },

  // --- Resoconto ------------------------------------------------------------------
  {
    id: "res-lettura",
    sezione: "Resoconto e incassi",
    titolo: "Come si legge il Resoconto",
    testo: `Il Resoconto guarda le fatture aperte alla data scelta ("Situazione al…"): quanto c'è da incassare, quanto da pagare, con le fasce di ritardo e la colonna che dice da quanti giorni una fattura è scaduta (o tra quanti scade). Le passive sono col segno meno, così incollando in Excel i totali si sommano da soli.

Sotto ci sono anche gli elenchi completi — tutte le fatture da incassare e da pagare, in ritardo o no. I giorni nella colonna sono rossi quando la scadenza è già passata e verdi quando deve ancora arrivare. Ogni lista ha il suo bottone Esporta CSV, che scarica la lista intera (non solo le 200 righe mostrate a schermo).`,
    chiavi:
      "resoconto situazione ritardi fasce scaduto da incassare da pagare verde rosso giorni colori",
  },
  {
    id: "res-esclusioni",
    sezione: "Resoconto e incassi",
    titolo: 'Il toggle "Applica esclusioni"',
    testo: `Accanto alle fasce c'è l'interruttore che applica le stesse esclusioni configurate nei Flussi di cassa: le controparti escluse spariscono da conteggi e ritardi. Spento di default — se i numeri ti sembrano diversi dal solito, controlla prima di tutto in che posizione è.`,
    chiavi: "esclusioni toggle resoconto numeri diversi controparte esclusa",
  },
  {
    id: "res-compensazione",
    sezione: "Resoconto e incassi",
    titolo: "Compensare attive e passive di una controparte",
    testo: `Nell'elenco del Resoconto puoi spuntare più fatture, anche miste attive e passive: in fondo compare il netto della compensazione ("tanto a incassare meno tanto da pagare = netto"). Utile quando con una controparte si è sia clienti che fornitori.`,
    chiavi: "compensazione spunta netto cliente fornitore saldo",
  },

  {
    id: "res-quattro-numeri",
    sezione: "Resoconto",
    titolo: "I quattro numeri della controparte e il saldo verde/rosso",
    testo: `Dalla 1.102.0 (Simone 09/11/2026, dopo il caso Posta Doc) in cima al Resoconto, sopra lo specchietto Aruba, c'è il riquadro "I quattro numeri della controparte", valido per qualunque filtro o vista salvata: Incassato in banca (i bonifici ricevuti dalle controparti selezionate), Fatturato attivo al netto delle note di credito, Fatturato passivo al netto delle NC, Pagato in banca (i nostri bonifici a loro). In fondo il saldo = attivo − incassato − (passivo − pagato): verde con "ci devono" quando è a nostro favore, rosso con "dobbiamo noi" quando è a loro favore. È la stessa logica con cui si è quadrata Posta Doc (5.222.635,93 − 4.166.947,72 − 371.112,28 + 1.500 = 686.075,93). Banca e Aruba qui non si mischiano: incassato e pagato vengono dai movimenti bancari, le fatture da Aruba; lo specchietto sotto continua a dire se Aruba è aggiornata rispetto alla banca.`,
    chiavi:
      "resoconto quattro numeri incassato in banca fatturato attivo passivo pagato saldo verde rosso ci devono dobbiamo noi controparte postadoc",
  },
  {
    id: "res-saldo-controparte",
    sezione: "Resoconto",
    titolo: "Saldo per controparte: attive e passive affiancate",
    testo: `Dalla 1.101.0 (Simone 04/11/2026) sotto il quadro del Resoconto c'è la tabella "Saldo per controparte": per ogni cliente o fornitore della selezione (o di tutti, se non c'è filtro) le fatture attive e le passive stanno sulla stessa riga: fatture, note di credito, incassato o pagato (anche in compensazione, com'è registrato su Aruba), aperto, e il saldo = da incassare meno da pagare. Serve per chi è cliente e fornitore insieme (Univex, Nolvex, Posta Doc): si vede subito quanto ci deve al netto di quello che gli dobbiamo. Scegliendo un gruppo di controparti nei filtri si ottiene il saldo del gruppo società per società, con il totale in fondo. Nel quadro sopra c'è anche "Passive totali" con pagate/compensate e da pagare, come per le attive. Le righe sono ordinate per saldo; con più di 25 controparti il bottone apre l'elenco completo.`,
    chiavi:
      "resoconto saldo controparte attive passive affiancate compensazione univex nolvex postadoc passive totali quanto ci deve",
  },
  {
    id: "res-estratto-gruppo",
    sezione: "Resoconto e incassi",
    titolo: "L'estratto conto di gruppo (Univex e simili)",
    testo: `Finanze → Estratto conto (dalla 1.91.0). Serve per i clienti che sono un gruppo di società e regolano per compensazione, come Univex: Healthcare e Freight ci pagano, Nolvex e Univex Srl ci fatturano, e i loro bonifici coprono le nostre fatture al netto di quello che dobbiamo. Il portale ricostruisce l'estratto conto come lo vede il cliente: per società e mese di competenza, fatturato (al netto delle note di credito), bonifici ricevuti, compensazioni e aperto, con lo scaduto alla data scelta.

La regola delle compensazioni si imposta una volta per gruppo: le fatture passive di una società che ci fattura si scaricano sulle sue stesse fatture; quelle di chi non ci fattura (Nolvex, Univex Srl) vanno sulla "società che assorbe", scelta in alto (per Univex: Freight). I bonifici si attribuiscono per competenza, dal più vecchio; un bonifico pari esatto a una fattura (non tondo) va su quella fattura.

Accanto c'è la vista Aruba (quello che il portale registra oggi) e la colonna Differenza: quando le compensazioni sono registrate sulla stessa società del cliente è zero; se non lo è, è lì che il cliente e Sabrina non si capiranno. Il controllo mensile fa il conto che ha scoperto il caso del 17 settembre: incassi registrati su Aruba meno bonifici in banca deve fare le passive pagate per compensazione; rosso sopra i 10.000, giallo sopra i 1.000. La previsione mostra l'aperto per mese di scadenza al netto di quanto il cliente tratterrà. Esporta CSV come le altre pagine.

I gruppi sono quelli del Resoconto (Resoconto → Gruppi, membri separati da virgola, match per parole del nome). Le scadenze seguono i termini di pagamento del cliente come nel Resoconto.`,
    chiavi:
      "estratto conto gruppo univex healthcare freight nolvex compensazioni assorbe società competenza bonifici scaduto aruba banca semaforo previsione incassi",
  },
  // --- Fatture -------------------------------------------------------------------
  {
    id: "fat-stati",
    sezione: "Fatture",
    titolo: "Stati e residuo: come leggerli",
    testo: `Le attive sono Incassate o Non incassate, le passive Pagate o Non pagate; il residuo è quanto manca all'appello dopo incassi parziali e note di credito collegate. Una fattura con residuo zero è chiusa anche se formalmente lo stato testuale di Aruba non è ancora arrivato.

Caso particolare già gestito: quando su Aruba una fattura di un consulente risulta pagata ma le rate sommano meno del totale, la differenza è la ritenuta d'acconto (che viaggia con l'F24) e la fattura viene considerata saldata.`,
    chiavi: "incassata pagata residuo parziale nota credito ritenuta acconto saldata",
  },
  {
    id: "fat-aruba",
    sezione: "Fatture",
    titolo: '"Su Aruba risulta pagata ma sul portale no"',
    testo: `Gli stati di pagamento vivono solo sul sito di Aruba, non sulle API: li porta il "giro" che parte dal PC aziendale quattro volte al giorno (e il bottone "Sincronizza da Aruba" per il giro su richiesta). Se il PC è spento, gli stati non viaggiano.

Quindi: primo, Ctrl+F5 (magari è la pagina vecchia); secondo, guarda quando è passato l'ultimo giro; terzo, se serve subito lancia il giro su richiesta. Se dopo un giro completo lo stato ancora non torna, allora sì che è una segnalazione.

Il caso contrario, "su Aruba l'ho riaperta ma sul portale resta incassata", è voluto: il giro non riduce mai un incasso e non riporta mai una fattura indietro, per non farsi ingannare da un report sbagliato. Dalla 1.90.0 la correzione fatta apposta su Aruba si porta nel portale così: in Fatture, matita "Stato incasso" sulla fattura → "Non incassata", che azzera anche l'importo delle rate; poi, se su Aruba restano incassi parziali (acconti, compensazioni), lo script di scarico lanciato a mano come "incassi forza" scrive anche le riduzioni. Il giro delle 7:15 resta prudente come prima.

Dalla 1.96.0 per gli incassi tolti del tutto non serve più la matita: a fine corsa il giro conta le fatture che hanno un incasso nel portale ma non compaiono più nel report di Aruba (nel log del PC: "N con incassi assenti dal report") e "incassi forza" le azzera da solo, come farebbe la matita. Il giro schedulato conta e non tocca; con un blocco perso per strada o più di 40 fatture in una volta non si azzera niente, perché è il report a essere sospetto. Le copie scartate dallo SdI di una fattura reinviata non tengono l'incasso, che sta sulla copia valida; una fattura stornata resta stornata e perde solo l'importo fantasma.

Il giro parte dal PC che ha le due attività pianificate di Windows ("DR Portal - Aruba giornaliero" e "DR Portal - giro su richiesta"). Dalla 1.96.4 spostarle su un altro PC è un'operazione guidata: nel repo, cartella scripts/aruba-prima-nota/trasloco, c'è installa.bat con le istruzioni (LEGGIMI.txt). Controlla Python, Playwright ed Edge, registra le attività per l'utente che lo lancia e fa una corsa di prova; l'unico file da portare a mano è config.json (credenziali Aruba e indirizzo del portale), che non sta nel repo. Le attività partono senza la finestra nera del prompt (tramite wscript). Regola: mai attive su due PC insieme, prima si spengono sul vecchio.`,
    chiavi:
      "aruba pagata portale aperta stato giro pc sincronizza incasso non aggiornato riaperta riduzione forza matita non incassata sparite azzerate corsa report assenti trasloco pc nuovo installa attivita pianificate finestra nera",
  },
  {
    id: "fat-termini-opzioni",
    sezione: "Fatture",
    titolo:
      "Scadenza dal 1° del mese, competenza al mese prima, scadenza in fattura ignorata, arrotondamento al giorno 1-5-10-15, fatture a vista senza regola (accordi per controparte)",
    testo: `Nei Termini di pagamento (Finanze → Regole, scheda Clienti) ogni cliente ha i suoi giorni. Dalla 1.81.0 ci sono due opzioni in più, pensate per l'accordo con Univex (24/09) ma valide per chiunque le abbia:

"Scadenza dal 1° del mese della fattura": se la fattura NON dichiara una scadenza propria, i giorni partono dal primo del mese in cui è stata emessa, non dalla data fattura. Una fattura del 21/07 a 65 giorni scade il 4/09, non il 24/09. Se invece nell'XML c'è una scadenza vera (diversa dalla data fattura), vince quella. Per i clienti senza l'opzione resta la regola di sempre: termini dalla data fattura.

"Competenza = mese precedente": quando il mese non è scritto nella descrizione della fattura, la competenza è il mese prima dell'emissione, qualunque sia il giorno (senza l'opzione vale la regola del giorno 15). Se il mese è scritto in fattura ("competenze giugno"), vince sempre quello.

Le opzioni si vedono come etichette accanto al cliente nell'elenco dei termini ("dal 1° del mese", "mese prec.") e si tolgono rientrando nel termine con la matita. Cambiano solo scadenze, ritardi e mese di competenza: importi e stati d'incasso non si toccano.

Dalla 1.98.0 c'è una terza spunta, per clienti E fornitori: "Ignora la scadenza scritta in fattura". Con la spunta la data dichiarata nell'XML non conta mai e valgono solo i giorni della regola (dalla data fattura, oppure dal 1° del mese se c'è anche quell'opzione). Serve soprattutto ai fornitori, dove di norma vince la scadenza scritta in fattura, e ai clienti con "dal 1° del mese"; per gli altri clienti i termini contrattuali valgono già da soli. Nell'elenco si riconosce dall'etichetta "solo giorni regola". Le scadenze cambiano ovunque: Fatture, Flussi, Cassa proprietà, solleciti.

Dalla 1.106.0 (Simone 11/11/2026) c'è "Arrotonda la scadenza al", con i quattro giorni 1, 5, 10 e 15 da accendere, per clienti E fornitori. Serve per chi paga in giorni fissi del mese: la scadenza, sia quella calcolata dai giorni della regola sia quella scritta in fattura, si sposta IN AVANTI al primo dei giorni accesi, uguale o successivo. Con il solo 10 acceso: una scadenza del 23/11 diventa 10/12, una dell'8/11 diventa 10/11, una del 10/11 resta 10/11. Con più giorni accesi vale il primo che capita: con 1, 5, 10 e 15 una scadenza del 12/11 diventa 15/11 e una del 20/11 diventa 1/12. Nessun giorno acceso = nessun arrotondamento. L'arrotondamento si applica per ultimo, dopo i giorni, il "dal 1° del mese" e la scadenza ignorata. Nell'elenco si riconosce dall'etichetta "arrotonda al 10" (o "arrotonda al 1-5-10-15"); si toglie rientrando nella regola con la matita e spegnendo i giorni. Attenzione in Cassa proprietà: quando una regola cambia le date, le righe degli entranti cambiano scadenza e le scelte fatte su quelle righe (standby, da sollecitare, note) vanno rimesse.

SENZA REGOLA: FATTURE A VISTA (1.108.0, Simone 11/11/2026: "per le fatture a vista, senza data scadenza o regole venga presa la data fattura"). Quando la controparte NON ha una regola nei termini di pagamento, una fattura a vista (scadenza scritta uguale alla data fattura) o senza scadenza scade alla DATA FATTURA: prima il portale le metteva a 30 giorni dalla data. Vale per le attive e per le passive. Restano com'erano: le controparti con una regola (vincono sempre i giorni e le opzioni della regola); le passive senza regola con una scadenza scritta in fattura (vale quella); le attive senza regola con una scadenza scritta diversa dalla data (fino alla 1.108.0 andavano a 30 giorni dalla data fattura; dalla 1.109.0 vale la scadenza scritta, vedi sotto). Effetti da conoscere: le fatture a vista dei clienti senza regola entrano subito tra gli entranti della Cassa proprietà e contano come scadute dal giorno della fattura; le fatture dei fornitori senza scadenza e senza regola risultano da pagare subito, e se il fornitore è in RID o RiBa la Cassa le considera già addebitate dalla banca. Se una controparte paga (o va pagata) più tardi, basta darle una regola con i suoi giorni.

LA SCADENZA SCRITTA IN FATTURA VINCE (1.109.0, Simone 11/11/2026: "quelle che hanno una scadenza scritta seguono la scadenza, a meno che non abbiano una regola che dica di non seguire la scadenza"). Da questa versione la scadenza si decide così, uguale per clienti e fornitori, in ordine: 1) la regola della controparte ha la spunta "Ignora la scadenza scritta in fattura": valgono solo i giorni della regola; 2) in fattura c'è una scadenza scritta: vale quella (per le fatture attive conta solo se è diversa dalla data fattura, perché Aruba scrive la data fattura quando non si indica nulla); 3) c'è una regola: valgono i suoi giorni, dalla data fattura o dal 1° del mese; 4) non c'è nessuna regola: vale la data fattura. In coda si applica l'eventuale arrotondamento al giorno. COSA CAMBIA rispetto a prima: per i CLIENTI con una regola, fino alla 1.108.0 vincevano sempre i giorni della regola, anche se in fattura era scritta un'altra data; ora vince la data scritta. Se per un cliente devono valere sempre e solo i giorni del contratto, basta mettere la spunta "Ignora la scadenza scritta in fattura" sulla sua regola (etichetta "solo giorni regola"). Per i fornitori non cambia nulla: era già così.`,
    chiavi:
      "termini scadenza primo del mese decorrenza competenza mese precedente univex accordo ritardo sollecito ignora scadenza fattura solo giorni regola fornitori xml arrotonda arrotondamento giorno fisso 1 5 10 15 al 10 del mese a vista senza regola senza scadenza data fattura default 30 giorni",
  },
  {
    id: "fat-colonne",
    sezione: "Fatture",
    titolo: "Colonne, filtri ed export",
    testo: `Le tabelle attive e passive hanno le stesse colonne del CSV (regola della casa: ogni colonna nuova va in entrambi), con gli imbuti sulle intestazioni. La classificazione gestionale (tipologia, sottocategoria, allocazioni) si corregge con la matita fattura per fattura, e come sui movimenti il manuale vince campo per campo.`,
    chiavi: "colonne csv export imbuti matita fattura classificazione",
  },

  // --- Pivot ----------------------------------------------------------------------
  {
    id: "piv-uso",
    sezione: "Pivot",
    titolo: "Usare la Pivot: gruppi, colonne e scorrimento",
    testo: `La Pivot raggruppa movimenti e fatture per i campi che scegli tu (allocazioni, tipologia, sottocategoria, cliente, anno, mese, fiscal week), coi mesi in colonna e il totale in fondo. I filtri in alto restringono i dati prima del raggruppamento.

Per comprimere o espandere un gruppo clicca sul suo nome in una riga qualsiasi, oppure sulla riga del suo totale; "Comprimi tutto" ed "Espandi tutto" fanno tutto in un colpo. Per scorrere i mesi usa i bottoni ◀ ▶ sopra la tabella — sono più affidabili delle freccine della barra di Windows. Ogni vista si esporta in CSV così com'è.`,
    chiavi:
      "pivot gruppi comprimere espandere aprire chiudere non si apre frecce scorrere mesi laterali colonne subtotali",
  },

  // --- Flussi ---------------------------------------------------------------------
  {
    id: "flu-ripartizione-appalti",
    sezione: "Flussi di cassa",
    titolo: "Un fornitore su più appalti: la ripartizione in percentuale, anche dai mezzi",
    testo: `Dalla 1.103.0 (Simone 09/11/2026: "assegnare un fornitore a più appalti per dividere meglio le spese") la scheda fornitore (Flussi → clic sul nome, o Cassa proprietà) ha il blocco "Ripartizione tra appalti": una riga per appalto con la sua percentuale, totale 100. Serve per i costi dei mezzi che stanno su più appalti: noleggi, carburante, pedaggi, multe, officine.

"Bilancia dai mezzi" la calcola il portale dal modulo Mezzi: se il fornitore ha contratti di noleggio attivi (il noleggiatore del contratto è lui), ogni mezzo pesa col suo canone sull'appalto del mezzo; se non ha contratti (carburante, pedaggi, multe) pesa il parco attivo per appalto, un peso per mezzo. Con la spunta accesa le quote restano agganciate ai mezzi e si aggiornano da sole quando un mezzo cambia appalto; "Proponi dai mezzi" le scrive una volta sola e poi si correggono a mano. I mezzi resi, dismessi, venduti o rottamati non contano; quelli senza appalto finiscono in "Altri".

Dove si vede: nella tab Fornitori sotto l'appalto ("Zingali 60% · Univex 40%", con "dai mezzi" se automatica) e in Cassa proprietà con "Uscite per cliente", dove la riga del fornitore si divide tra gli appalti secondo le quote (nel dettaglio del gruppo ogni fornitore porta la sua percentuale). Senza ripartizione vale, come prima, l'appalto unico della scheda.`,
    chiavi:
      "fornitore più appalti ripartizione percentuale quote mezzi noleggio carburante pedaggi multe bilancia dai mezzi uscite per cliente dividere spese",
  },
  {
    id: "flu-cassa-proprieta",
    sezione: "Flussi di cassa",
    titolo: "La vista Cassa proprietà (il foglio)",
    testo: `Dalla 1.92.0 i Flussi hanno una seconda vista, "Cassa proprietà", col bottone in alto accanto al titolo; dalla 1.93.0 è IL FOGLIO del commercialista, tale e quale, perché la proprietà legge solo quello: colonne Entrate | Note | Saldo | Uscite | 1-5 del mese dopo | Standby | Note | Note2, e in fondo i suoi totali con le sue etichette. Niente cruscotti né elenchi: tutto sta nel foglio.

Come si legge. ENTRATE a sinistra, in due blocchi col loro totale: "Già entrate" = gli incassi arrivati in banca dal primo del mese (con la data accanto, sono già nel saldo) ed "Entranti" = in corsivo grigio con la scritta "attesa", quello che i clienti devono ancora pagare entro la data limite, una riga per cliente con la somma delle fatture aperte: l'importo si corregge a mano con quello che il cliente annuncia (clic sul numero). SALDO: dal collegamento bancario; se la banca non risponde si scrive nella casella in alto e resta salvato. USCITE in tre colonne: "Uscite" = da pagare ora (scadenze del mese, scadute comprese), "1-5 novembre" = i prossimi giorni fino alla data che scegli (il giorno resta salvato), "Standby" = parcheggiate apposta. Una riga per fornitore, non per fattura (il ×3 dice quante sono; "auto" = RID/RiBa, esce da sola), più gli stipendi netti da pagare per appalto dal file paghe (le spunte Pagato della tab Stipendi li tolgono; gli arretrati si guardano solo fino al mese di paga precedente, più indietro si considera pagato), "Fiscale + BPM" con le scadenze non pagate del consulente (le "previsionali" non entrano) e i costi fissi del mese non ancora usciti. Note = il nome, Note2 = una cella libera dove scrivere ("contestare", "compensazione", "sentire Marelli").

Come si lavora. Per spostare un'uscita di colonna si clicca la cella vuota della colonna dove la si vuole; per correggere un importo (stipendi, fiscale, fissi, attesi) si clicca il numero; la × a bordo riga toglie la riga dal foglio; "+ aggiungi" in fondo a ogni lato aggiunge una riga a mano (divisione utili, compensazioni, incassi annunciati) e la spunta CDC segna una spesa passata dalla carta Amex, che nel foglio compare con "CDC" in Note2. Una riga aggiunta a mano può essere "ogni mese": si ripresenta da sola allo stesso giorno di ogni mese finché non la si elimina — è il modo per tenere nel foglio quello che non ha una fattura nel portale (gli stipendi iMile, che non stanno nel file paghe, i buoni pasto, il commercialista bimestrale). Quello che scade oltre la data limite sta sotto il foglio, "Oltre il 5/11", con i bottoni per portarlo dentro o parcheggiarlo. Le esclusioni dei Flussi valgono anche qui.

Dalla 1.95.0 si distingue a colpo d'occhio cosa è calcolato e cosa no: le righe scritte a mano hanno la tinta viola e la pillola "manuale" (e "ogni mese" se ricorrenti); gli stipendi dei mesi che Lucrezia non ha ancora caricato sono STIMATI sulla media degli ultimi due mesi di netti per appalto, con la pillola "stima", e si correggono a mano (gli appalti chiusi — Postadoc e Ceva dal 30/09/2026, Univex Savona dal 31/10/2026 — non si stimano); le fatture con RID o RiBa hanno il lucchetto e non si possono spostare di colonna, perché l'addebito arriva da solo alla scadenza. Il nome del fornitore si clicca e apre la sua scheda, come nei Flussi classici. La legenda sta sotto il foglio.

Sotto il foglio, dalla 1.94.0, due riquadri. CALENDARIO ENTRATE: le entrate attese in ordine di data con il progressivo; la data di ogni attesa è quella di scadenza finché non la si cambia con un clic sull'etichetta "attesa · gg/mm" (quando il cliente dice quando paga: l'etichetta diventa piena), le date passate sono in rosso. AGENDA · COSE DA FARE: il pallino accanto a ogni nota in Note2 dà una priorità alla nota — un clic bassa (verde), due media (ambra), tre alta (rosso), quattro la toglie — e la nota compare in agenda, ordinata per priorità e data; si aggiungono anche voci libere con priorità e data, da spuntare quando fatte, e la × le elimina.

I totali sono quelli del foglio: "standby incasso" = le entrate attese; "da pagare", "da pagare 1-5/11", "standby da pagare", "totale da pagare" (= da pagare + standby), "cassa+standby incassi-standby da pagare" (= saldo − da pagare − prossimi + attese − standby); sotto, "saldo dopo incassi e pagamenti" (= saldo − da pagare) e accanto lo stesso tolti anche i prossimi giorni. Le entrate arrivate NON si sommano: sono già nel saldo. "Salva fotografia" congela il foglio del giorno con una nota e lo mette fra le linguette sotto il foglio, come i fogli "1210" e "1510" di Excel: si riapre in sola lettura. In alto, sopra il foglio, il saldo in banca grande (si scrive lì se la banca non risponde) e i tre numeri che contano: da pagare, saldo dopo incassi e pagamenti, cassa con gli standby. Esporta CSV con la stessa griglia. I Flussi classici non cambiano.

Dalla 1.97.0 (richieste della proprietà del 02/11/2026) il foglio ha quattro cose in più. REALE: una colonna accanto al Saldo con "saldo + entranti del mese − da pagare − prossimi giorni", cioè quello che resta davvero; sotto il numero c'è il campo "acconti": se un cliente ha già versato un acconto sugli entranti, lo si scrive lì e viene tolto dal totale entrante (è già nel saldo). STANDBY DELLE ENTRATE (in soldoni: quello che non ci pagheranno mai): ogni attesa ha la pillola "⏸ Standby" accanto alla data; messa in standby finisce nel terzo blocco "Standby" della colonna Entrate, in fondo all'elenco degli entranti, resta in vista ma non conta né tra le entranti né nella cassa (↩ la rimette tra le entranti); il totale "entrate in standby" sta nella riga dei totali. STANDBY 2 delle uscite (quello che non pagheremo mai): una quarta colonna, "standby prolungato", per le spese parcheggiate a lungo: non entrano in "totale da pagare" né nella cassa. Le date delle attese portano anche l'anno ("attesa · 16/07/26").

Tre regole nuove sui numeri. Le righe fiscali sono tre: "Fiscale" (F24 e tutte le rateizzazioni, anche quelle dell'IVA), "IVA" (l'IVA corrente, che si parcheggia o si paga) e "Banca" (le rate di finanziamento), ognuna nella sua colonna per scadenza ma una sotto l'altra nel foglio, in quest'ordine. I RID/RiBa già scaduti (e le rate di finanziamento ad addebito già scadute) non stanno più tra le uscite: la banca li ha già presi e il saldo li ha già dentro; sotto il foglio una riga dice quanti sono e per quanto. Una fattura che scade l'ultimo giorno del mese vale come scadenza il primo del mese dopo (quindi va nei "prossimi giorni", non in "da pagare").

Due comodità: il bottone "Uscite per" sopra il foglio raggruppa le spese per tipologia di costo (dalle fatture, poi dalla macrovoce del fornitore) o per cliente/appalto (ClienteRif, poi l'appalto del fornitore, gli stipendi per appalto), in sola lettura con l'elenco dei fornitori nel dettaglio; "fornitore" riporta al foglio normale. Il nome di una riga scritta a mano si cambia con un clic (così "Stipendi iMile" è diventata "Acconto per DR Logistics (Aldo)"). Il foglio scorre dentro la sua cornice con le testate fisse.

Dalla 1.98.0 (richieste di Simone del 04/11/2026) i nomi del foglio sono link, come nella vista classica: il cliente di un'entrata (attesa o già in banca) apre la scheda con le sue fatture aperte (numero, data, scadenza con la regola di fine mese, importo, residuo, stato Aruba, giorni di ritardo); il fornitore di un'uscita apre la scheda fornitore dei Flussi, anche dalla lista "oltre il limite" e dagli addebitati; con "Uscite per" tipologia o cliente, il nome del gruppo apre l'elenco dei fornitori dentro, ognuno cliccabile. Le entrate attese stanno su UNA RIGA PER CLIENTE E SCADENZA: le fatture dello stesso cliente con la stessa data restano insieme, quelle con date diverse vanno su righe diverse (standby, nota e data scelta sulla vecchia riga unica si ereditano). Sulle uscite, accanto al nome del fornitore, il foglietto mostra al passaggio il promemoria scritto nella sua scheda (cosa fa davvero, regole da ricordare); nel dettaglio della riga c'è anche la causale scritta su ogni fattura.

UN "+" PER CLIENTE (1.104.0, Simone 10/11/2026: "unire i clienti, ad esempio mettendo un + per cliente"). Tra gli entranti ogni cliente con più fatture attese sta su una riga sola: totale, "dal" prima scadenza, numero di fatture, e il "+" davanti al nome che apre le singole fatture (una per scadenza, con importo, data attesa, standby e × come prima); il "−" le richiude. Le decisioni (standby, nota, data) si prendono sulle singole fatture; la riga del cliente è solo la somma. L'apertura non si salva: al ricaricamento tutto richiuso.

STAMPA ED EXPORT DELLE VOCI (1.109.0, Simone 11/11/2026). In alto a destra, accanto a "Esporta CSV" e "Stampa" del foglio intero, c'è "Stampa / esporta voci": apre l'elenco delle sette sezioni del foglio, ognuna col numero di voci e il totale: 1) Uscite imminenti (la colonna Uscite), 2) Uscite della seconda colonna (1-5 del mese dopo), 3) Entrati (già in banca), 4) Entranti (dentro anche le "da sollecitare", segnate), 5) Standby entrate, 6) Standby uscite, 7) Standby 2 uscite. Accanto a ogni sezione ci sono due pulsanti: la freccia esporta il CSV (si apre in Excel: sezione, voce, dettaglio, data, importo, tipologia, cliente/appalto, priorità, da sollecitare, nota, e in fondo il totale); la stampante apre la sezione in una finestra sua, pronta da stampare, con intestazione, data, voci in ordine di data e totale (dal riquadro di stampa si sceglie "Salva come PDF" per avere il PDF). In fondo all'elenco, "Tutte le sezioni insieme" fa lo stesso con tutte le sezioni una dopo l'altra. Se i filtri del foglio sono accesi, escono solo le voci filtrate e la stampa lo scrive nell'intestazione: per avere tutto, togliere prima i filtri. Le uscite escono sempre fornitore per fornitore, anche quando il foglio è raggruppato per tipologia o per cliente. Se la finestra di stampa non si apre, il browser sta bloccando i popup: consentirli per il portale.

FILTRI E BARRA IN ALTO (1.107.0, Simone 11/11/2026). Sopra il foglio c'è la barra dei filtri, in due gruppi. ENTRATE: una casella di ricerca (cliente, numero di fattura, importo, nota) e la tendina del blocco (già entrate, entranti, da sollecitare, standby). USCITE: una casella di ricerca (fornitore, voce, causale, importo, nota), i quattro pulsanti delle colonne (Uscite, 1-5 del mese dopo, Standby, Standby 2: si accendono una o più), le tendine della tipologia e del cliente/appalto, quella della priorità e la spunta "solo con nota". I filtri si sommano; più parole nella ricerca devono comparire tutte. Con "Uscite per tipologia" o "per cliente" si filtra prima e si raggruppa poi, quindi i gruppi sommano solo le voci filtrate. Quando un filtro è acceso compare una riga gialla con quante voci restano e quanto sommano (per le uscite anche colonna per colonna) e i blocchi delle entrate mostrano la somma delle voci visibili; "Togli filtri" riporta tutto com'era. I filtri sono solo vista: non si salvano, e i totali in fondo al foglio, il saldo e l'export restano sempre quelli di tutte le voci. Sopra il foglio, sotto i filtri, c'è anche una barra di scorrimento orizzontale: quando il foglio è più largo dello schermo ci si sposta a destra da lì senza scendere fino in fondo; si muove insieme a quella in basso.

DA SOLLECITARE (1.105.0, Simone 10/11/2026). Tra "Entranti" e "Standby" c'è il blocco "Da sollecitare": con il pulsante "📣 Sollecita" sulla singola fattura attesa la si sposta lì, con "↩ Entranti" torna indietro. A differenza dello standby, la fattura da sollecitare CONTA ancora tra gli entranti e nella cassa (è un credito che ci pagheranno): il blocco serve solo a vedere subito chi chiamare, col suo totale; sotto il totale degli entranti c'è "di cui da sollecitare". Il "+" per cliente vale in tutti e tre i blocchi (entranti, da sollecitare, standby).

ENTRATE PER GRUPPO (sempre 1.104.0). Accanto a "Uscite per" c'è "Entrate: per società / per gruppo". Per gruppo, le società che stanno nello stesso gruppo di controparti del Resoconto (Resoconto → Gruppi, oppure una vista salvata: Univex Healthcare, Freight e Srl dentro "Univex") finiscono su una riga sola col nome del gruppo, sempre una riga per scadenza; nel dettaglio si leggono le società e le fatture, e il clic sul nome apre le fatture di tutte. Vale anche per gli incassi già in banca. La scelta resta salvata nel foglio. Le righe unite partono senza standby né note: si decidono da capo. Dalla stessa versione gli acconti scritti nella colonna Reale restano salvati (prima sparivano ricaricando la pagina).

Dalla 1.99.0 (Simone 04/11/2026). COMPENSAZIONI: nella scheda fornitore c'è la spunta "Paga in compensazione" (Ceva, Number1, Univex, Posta Doc…): le fatture di quel fornitore non stanno più tra le uscite ma diventano una riga NEGATIVA tra gli entranti, col suo nome e l'etichetta rossa "compensazione" (il dettaglio elenca le fatture; la riga si parcheggia come le altre attese). Gli entranti netti scendono di conseguenza e il "da pagare" non le conta più; sotto il totale entranti c'è "di cui compensazioni". Un clic su Fiscale, IVA o Banca apre il popup con le scadenze dentro la riga (data, voce, modalità, importo), così il numero si riscontra con lo scadenziario di Sabrina. Nella riga IVA (1.99.1) sta solo l'IVA da pagare, cioè la liquidazione alla sua scadenza naturale (16 del mese dopo per le mensili, 16/05, 20/08, 16/11 e 16/03 per le trimestrali, acconto 27/12, con 45 giorni di tolleranza): le rateizzazioni dell'IVA stanno nel Fiscale. La stima degli stipendi dei mesi non ancora caricati (1.99.1) è per persona: solo chi è nell'ultimo mese caricato, con la media del suo netto sugli ultimi due mesi e l'appalto dell'ultimo mese; il numero accanto alla riga ("N dip.") è quante persone ci sono dentro.

CARBURANTI E ALTRE FATTURE RICORRENTI (1.100.0, Simone 04/11/2026: "arrivano ogni 15 giorni, guarda l'andamento e replicalo"). Nella scheda fornitore c'è il campo "Fatture ricorrenti: stima le prossime ogni N giorni" (ENI e Kuwait = 15). Con quel numero la Cassa, dopo l'ultima fattura arrivata, aggiunge tra le uscite una riga stimata per ogni fattura attesa entro il limite: importo = media delle ultime sei fatture, data = il 15 e l'ultimo del mese (con la cadenza di 15 giorni), scadenza = data più il ritardo tipico delle ultime (per ENI e Kuwait 15 giorni, RID). Le righe sono in corsivo con "stima" nel dettaglio, l'importo si corregge a mano; quando la fattura vera arriva la stima di quella data sparisce da sola.`,
    chiavi:
      "cassa proprietà foglio commercialista link scheda cliente fatture aperte gruppo fornitori promemoria causale fattura una riga per scadenza compensazione compensa ceva number1 univex postadoc scala entranti popup fiscale scadenze dentro la riga fotografia saldo banca standby parcheggiate prossimi giorni 1-5 novembre scadenze imminenti divisione utili amex cdc carta credito stipendi non pagati niente previsioni entrate attese incassi arrivati da pagare totale da pagare calendario entrate data attesa agenda cose da fare priorità alta media bassa ogni mese ricorrente stipendi imile reale acconti standby 2 prolungato entrate in standby fiscale iva banca rid addebitati fine mese primo del mese raggruppa tipologia cliente rinomina acconto dr logistics",
  },
  {
    id: "flu-lettura",
    sezione: "Flussi di cassa",
    titolo: "Come si legge la tabella dei Flussi",
    testo: `Ogni colonna è un mese (o una settimana), la prima colonna è lo Scaduto: quello che era già in ritardo oggi. Le entrate vengono dalle scadenze delle fatture attive aperte, le uscite dalle passive, più le righe che le fatture non conoscono: Stipendi, Consulenze, Costo fiscale rate, Costo fiscale corrente, Altre spese. In fondo il saldo per colonna.

Le voci dei singoli clienti e fornitori sono chiuse: il "+" accanto a Entrate e a ogni blocco delle uscite (RiBa, RID, senza) apre quel gruppo, il "−" lo richiude; "Mostra dettaglio" in alto apre tutto in un colpo. Il CSV esce sempre completo.

Il bottone Cumulato trasforma le colonne in progressivo: ogni mese somma tutto quello che c'è fino a quel punto, Scaduto compreso.`,
    chiavi:
      "flussi cassa colonne mese settimana scaduto saldo cumulato lettura più aprire chiudere voci fornitori gruppo",
  },
  {
    id: "flu-due-tabelle",
    sezione: "Flussi di cassa",
    titolo: "Perché ci sono due tabelle",
    testo: `La tabella principale è la cassa attesa vera: scala gli incassi già arrivati, segue le spunte dei pagamenti e per gli stipendi mostra ciò che è segnato pagato. La seconda ("Solo fatturazioni") è la stessa fotografia ma a fatturato pieno: ogni fattura conta per intero alla sua scadenza, incassate e pagate comprese. Serve a confrontare il fatturato con la cassa: partono dallo stesso Scaduto e divergono sui mesi.`,
    chiavi: "solo fatturazioni due tabelle differenza fatturato pieno cassa",
  },
  {
    id: "flu-voci",
    sezione: "Flussi di cassa",
    titolo: "Le quattro voci e il click sul nome",
    testo: `Stipendi, Consulenze, Costo fiscale rate, Costo fiscale corrente e Altre spese si riempiono da sole: gli stipendi dai netti caricati e dalle spunte Pagato, le Consulenze dai pagamenti con quella tipologia nei movimenti (riga propria dalla 1.86.0, perché non sono costi ricorrenti), le voci fiscali dallo scadenzario, le Altre spese dalla media degli ultimi due mesi di costi generali non fatturati più i costi fissi (affitti) al loro importo. Il simbolo ≈ segnala il valore automatico.

Il mese corrente è il REALE, non una stima (dalla 1.85.0): le voci fiscali contano le scadenze pagate nel mese più quelle da pagare entro il mese; le Altre spese contano quanto è già uscito dalla banca più la media per i giorni che mancano, quindi all'ultimo giorno del mese la cella coincide con la spesa vera. I mesi futuri restano stime (media, scadenze in agenda).

Cliccando sul nome della voce si apre il dettaglio di cosa c'è dentro; nelle Altre spese puoi includere o escludere le singole tipologie con le spunte, e la media si aggiorna. Se scrivi un importo a mano nella cella, il tuo numero vince sull'automatico per quel mese — per tornare all'automatico svuota la cella.`,
    chiavi:
      "voci stipendi fiscale altre spese drill click dettaglio automatico manuale cella mese corrente reale",
  },
  {
    id: "flu-costi-fissi",
    sezione: "Flussi di cassa",
    titolo: "Costi fissi (affitti) nelle Altre spese",
    testo: `Gli affitti e le altre uscite fisse non vanno stimati con la media: dalla 1.85.0 si impostano nel pannello delle Altre spese (clic sul nome della riga, blocco "Costi fissi") con nome, importo mensile, parola con cui riconoscerli nei movimenti bancari (es. "poggi" per l'affitto di Fiano) ed eventuale finestra di mesi. Da lì in poi i loro pagamenti escono dalla media dei costi variabili e la riga li conta al loro importo: nel mese corrente quello davvero uscito dalla banca (con la spunta ✓ e la data), altrimenti quello impostato. Nel pannello vedi per ogni costo fisso se nel mese è già stato pagato.

Impostati il 29/09/2026 dai file di Sabrina: Affitto Fiano 1.200, Zekaj Rovena 2.000, Liza Perlala 2.100 (fino ad aprile 2027), Biagi Mara 950.`,
    chiavi: "costi fissi affitto affitti fiano zekaj perlala biagi importo reale media altre spese",
  },
  {
    id: "flu-anomalie",
    sezione: "Flussi di cassa",
    titolo: "Anomalie: i pagamenti fuori misura non entrano nella media",
    testo: `Le medie di Altre spese e Consulenze servono a stimare i mesi futuri, e un pagamento fuori misura le falserebbe (le disposizioni da 25.000 di settembre non sono la spesa di un mese normale). Dalla 1.86.0 un movimento è "anomalo" quando da solo supera del 25% la spesa di un mese tipico della sua tipologia (la mediana dei totali mensili dei sei mesi precedenti), con un minimo di 1.000 €: una parcella pari al mese tipico è la normale rata mensile, una disposizione da due mesi no. Le anomalie contano nel reale del mese corrente (sono soldi usciti davvero) ma non nella media dei mesi dopo.

Nel pannello delle Altre spese (blocco "Anomalie") e in quello delle Consulenze vedi i movimenti più grandi dei tre mesi con la spunta Anomalia: la regola la mette da sola, tu puoi forzarla in un senso o nell'altro e la scelta resta salvata (il simbolo ✎ segnala le tue).`,
    chiavi:
      "anomalie anomalia media fuori misura una tantum consulenze zekaj soglia mediana forzare spunta",
  },
  {
    id: "flu-stipendi-scaduto",
    sezione: "Flussi di cassa",
    titolo: "Arretrati nello Scaduto: stipendi e scadenze fiscali",
    testo: `Quando in un mese di pagamento arrivato (compreso quello corrente) restano stipendi non segnati pagati, il residuo compare nella colonna Scaduto della riga Stipendi: la cella del mese mostra i pagati, lo Scaduto il resto, e la somma torna al totale. Funziona dove la spunta Pagato è in uso nella tab Stipendi; i mesi vecchi senza spunte non contano, perché sono storia già regolata in banca.

Le due voci fiscali fanno lo stesso (dalla 1.85.0): le scadenze non pagate con data già passata stanno nella colonna Scaduto della loro riga, non nel mese corrente. Clic sul nome della voce per l'elenco.`,
    chiavi: "stipendi scaduto arretrati non pagati residuo scadenze fiscali scadute",
  },
  {
    id: "flu-esclusioni",
    sezione: "Flussi di cassa",
    titolo: "Escludere controparti dai Flussi",
    testo: `Dal pannello Esclusioni togli dalla vista le controparti che non vuoi contare (partite interne, casi particolari), con eventuale finestra di mesi. L'esclusione vale nei Flussi e, se accendi il toggle, anche nel Resoconto. Le stesse righe si possono raggruppare in preset per accenderle e spegnerle in blocco.`,
    chiavi: "esclusioni escludere controparte preset finestra mesi",
  },
  {
    id: "flu-girate",
    sezione: "Flussi di cassa",
    titolo: "Le girate ai fornitori",
    testo: `Una girata dice: "quando incasso da questo cliente, giro una percentuale a questo fornitore". Nei Flussi diventa una riga di uscita con il nome del fornitore. La quota matura a incasso avvenuto — si calcola sugli incassi registrati, note di credito comprese — e resta nello Scaduto finché non è coperta dai bonifici reali verso quel fornitore.`,
    chiavi: "girata quota percentuale fornitore incasso avvenuto dr logistics",
  },
  {
    id: "flu-modalita",
    sezione: "Flussi di cassa",
    titolo: "Uscite divise per RiBa, RID e senza",
    testo: `Dalla 1.84.0 le uscite dei Flussi sono divise in tre blocchi: "di cui con RiBa", "di cui con RID / addebito automatico" e "di cui senza RID né RiBa" (le fatture che partono solo se fai il bonifico). Ogni fattura passiva pesa nel blocco della modalità dichiarata nel suo XML (RiBa = MP12; RID = RID, SDD e domiciliazioni: Kuwait, Consat, Arval, Kion, BNP…), quindi un fornitore con fatture miste — TIM, Jungheinrich — compare in più blocchi. Col dettaglio acceso sotto ogni blocco ci sono i suoi fornitori; il CSV esce con la stessa struttura.

Per forzare un fornitore intero (es. DKV: bonifico in fattura ma addebito automatico in banca) vai nella tab Fornitori e scegli RID, RiBa o Nessuna nella colonna Modalità: "Auto" torna alle fatture.`,
    chiavi:
      "riba rid sdd addebito automatico modalità pagamento uscite divise bonifico forzare fornitore",
  },
  {
    id: "flu-simulazione",
    sezione: "Flussi di cassa",
    titolo: "Se non pago: simulare i mancati pagamenti",
    testo: `Il bottone "Se non pago…" apre tre leve che si sommano: scegli le singole fatture da non pagare (con ricerca, spunte, "Seleziona le visibili"), togli in un colpo tutte quelle senza RID né RiBa, scegli le scadenze fiscali da non pagare. Le due tabelle si ricalcolano subito, il banner giallo dice quante fatture e quanti euro hai tolto e come cambiano uscite scadute e delta saldo, e sotto ogni saldo compare la riga "Delta saldo senza simulazione" per il confronto mese per mese.

È solo una prova: non si salva nulla, non tocca fatture, spunte né scadenze, e sparisce chiudendo la pagina o con "Azzera simulazione". Il CSV esportato durante una simulazione lo dichiara nella prima riga e nel nome del file.`,
    chiavi: "simulazione se non pago non pagare fatture scadenze fiscali confronto saldo prova",
  },

  // --- Fornitori -------------------------------------------------------------------
  {
    id: "for-scheda",
    sezione: "Fornitori",
    titolo: "La scheda Fornitori del flusso di cassa",
    testo: `La tab Fornitori (Finanza) ha una riga per ogni fornitore con almeno una fattura passiva aperta nei Flussi (stesse regole: esclusioni comprese, DR Logistics fuori). Dalle fatture arrivano scaduto, aperto totale (scaduto + a scadere), numero di fatture scadute, aperte e totali in archivio, prima e ultima fattura, prossima scadenza (in rosso se già passata), fatture pagate e importo pagato ad oggi.

Le colonne Modalità, Chi se ne occupa, Macrovoce e Appalto sono tendine (dalla 1.87.0) e si salvano da sole appena scegli: Chi se ne occupa ha le persone di DR (Gabelli, Russo, Pratesi, Guidarelli, Marelli, Spera, Notaro, Carlone), Macrovoce le tipologie dei movimenti bancari (Carburante, Consulenze, Noleggio, Utenze…), Appalto le allocazioni degli appalti usate nelle regole di movimenti e fatture (iMile, Postadoc Hub, Ufficio Fiano Romano…). "Altro…" apre la casella libera per una voce nuova. Nelle tendine vuote la proposta tra parentesi viene dalla classificazione delle fatture. Cerca, ordina (aperto, scaduto, prossima scadenza, nome), Esporta CSV e Stampa come nelle altre pagine.

Gli appalti hanno un nome solo, lo stesso dell'anagrafica dipendenti (CEVA Logistics, iMile, Number1, Postadoc Colleferro (Distribuzione), Postadoc Fiano Romano (Distribuzione), Postadoc Fiano Romano HUB, Ufficio Fiano Romano, Ufficio Milano, Univex Milano, Univex Savona, Zingali, più Widem e "Italia" per i costi generali). La ragione sociale del cliente non è un appalto: Univex Healthcare e Univex Freight fatturano per Milano e per Savona, Postadoc per tutte le sue sedi, quindi sulle loro fatture di addebito l'appalto si sceglie fattura per fattura. Un cliente di riferimento che è il nome del fornitore stesso (NOLVEX sulle fatture di Nolvex) viene ignorato: non entra in tendina né nella proposta.

Nei Flussi di cassa, aperto un gruppo col "+", il clic sul nome del fornitore apre la stessa scheda in un popup: le quattro tendine con Salva, i numeri dalle fatture e l'elenco delle fatture aperte nel flusso con scadenza, residuo e modalità.`,
    chiavi:
      "fornitori scheda chi se ne occupa referente macrovoce appalto scaduto aperto pagate pagato prima ultima fattura prossima scadenza",
  },

  // --- Stipendi --------------------------------------------------------------------
  {
    id: "sti-import",
    sezione: "Stipendi",
    titolo: "Caricare i file paghe",
    testo: `La tab Stipendi si nutre di due famiglie di file: il COSTI del mese (il costo del personale per dipendente, dal consulente paghe) e Stipendi Dr (i netti da bonificare, un foglio per mese).

Dalla versione 1.79.0 il portale li legge DA SOLO da SharePoint (cartella Personale del sito Documenti Condivisi, la stessa che vedi in OneDrive): ogni notte col giro giornaliero, e quando vuoi col bottone "Aggiorna dai file" in cima alla tab. Basta quindi tenere aggiornati i file al loro posto: Personale\\Stipendi Dr.xlsx e Personale\\MENSILITA'\\<MESE ANNO>\\COSTI <MESE> <ANNO>.xlsx. Le correzioni a mano (la M), le spunte Pagato e le mensilità dichiarate non vengono toccate; la riga Stipendi dei Flussi si aggiorna solo se il saldo del mese è cambiato.

I bottoni di import restano per i casi particolari (un file in un'altra cartella, un tracciato per-appalto): sempre con l'anteprima prima di salvare, che dice cosa ha letto e cosa aggiornerà, compresa la riga Stipendi dei Flussi.

Attenzione a non "sporcare" i file: nel file COSTI il portale legge il foglio con l'export del consulente; i fogli "Dettagli1, Dettagli2…" che Excel crea da solo quando si fa doppio clic su una cella della pivot vengono ignorati (dalla 1.79.1), ma altri fogli aggiunti a mano con la stessa struttura verrebbero letti. Le note scritte nella colonna del cognome ("COSTO EXTRA 1", "Nel costo sono inclusi…") non diventano dipendenti: una riga senza nemmeno una voce di indirizzamento (Costi ordinari, Totalizzazioni…) viene scartata (dalla 1.80.1) e conta tra le "righe scartate" dell'anteprima. Meglio comunque lasciare il file com'è arrivato.

Se un file non viene letto, quasi sempre è questione di tracciato: intestazioni diverse dal solito o foglio rinominato. Segnala il file esatto.`,
    chiavi: "import file costi stipendi dr netti anteprima carica mensile",
  },
  {
    id: "sti-pagato",
    sezione: "Stipendi",
    titolo: "La spunta Pagato e la selezione per sede",
    testo: `Nella tabella per dipendente la colonna Pagato segna chi ha ricevuto il bonifico. La tabella dei totali in alto conta solo i segnati pagati; quella sotto tutti. La spunta in testata marca o smarca tutte le righe filtrate: filtra per una sede e con un click segni l'appalto intero.

Le spunte guidano anche i Flussi: il mese mostra i pagati e il residuo va nello Scaduto.`,
    chiavi: "pagato spunta colonna sede appalto selezione multipla bulk",
  },
  {
    id: "sti-modifiche",
    sezione: "Stipendi",
    titolo: "Correggere un valore a mano (la M)",
    testo: `Dalla versione 1.77.0 le celle dei costi (ordinario, straordinario, ferie, mensilità aggiuntive, TFR, totale) e dei netti (stipendio, anticipo, saldo) si correggono col doppio clic: scrivi il numero, Invio salva, Esc annulla. Se lasci la cella vuota, torna il valore del file.

La cella corretta mostra una M: fermandoti col mouse vedi chi l'ha cambiata, quando e quanto valeva nel file. È la traccia: il file paghe resta com'è, la correzione vive a parte e sopravvive anche a un nuovo import dello stesso mese. Le correzioni ai netti contano anche nei Flussi di cassa e nella stima degli stipendi futuri.

Serve anche a riempire un buco: se un dipendente ha la riga dei costi ma manca dal foglio Stipendi Dr del mese, scrivi stipendio e saldo nelle sue celle e la riga nasce da sola. Nel CSV la colonna "Modifiche manuali" elenca i campi corretti.`,
    chiavi:
      "modifica a mano correggere valore M manuale doppio clic traccia chi quando stipendio saldo costo",
  },
  {
    id: "sti-stipendio-vuoto",
    sezione: "Stipendi",
    titolo: "Un dipendente ha il costo ma lo Stipendio vuoto",
    testo: `Il costo viene dal file COSTI del mese, lo stipendio netto dal file Stipendi Dr: le due righe si agganciano per nome. Se lo Stipendio è "—" mentre il costo c'è, quasi sempre il nome è scritto diverso nei due file (un refuso, "Anna Maria" contro "Annamaria", una parola in più, l'ordine invertito).

Dalla versione 1.78.0 l'aggancio è tollerante: ordine delle parole, parole in più e troncamenti si risolvono da soli. I refusi veri no (è successo con "Giaggianesi" al posto di Gaggianesi): si corregge il nome nel file Stipendi Dr e alla lettura successiva torna tutto — le correzioni a mano (la M) e le spunte Pagato restano. In alternativa, col doppio clic scrivi stipendio e saldo direttamente nelle celle.

Poi ci sono i casi LEGITTIMI in cui il netto manca davvero: aspettativa non retribuita, infortunio, malattia, cessato, pagato fuori dal file. Dalla 1.80.0 in quella cella c'è un menù per scegliere il motivo: la riga smette di essere evidenziata e il motivo resta scritto (chi e quando). Le righe senza netto e senza motivo sono "da chiarire": il conteggio sta sopra la tabella ed è la lista da girare a HR. Sotto la tabella compaiono anche i nomi presenti in Stipendi Dr ma senza riga costi (liquidazioni, contanti), con lo stesso menù.

Le righe di Stipendi Dr intestate a una società (Pitagora SpA, Cofidis, Prestitalia, Sigla, UniCredit: le rate delle cessioni del quinto pagate con la distinta) dalla 1.80.2 vengono riconosciute da sole dal nome (SpA, Srl, S.A., "succursale", "finanz…") e stanno in un riquadro a parte "Finanziarie (cessioni del quinto)" col loro totale: non sono dipendenti, non c'è nessun motivo da scegliere, ma contano nel totale dei Flussi come prima. Se una finanziaria nuova non venisse riconosciuta (nome senza sigla societaria), finirebbe tra i "senza riga costi": segnala il nome esatto.`,
    chiavi:
      "stipendio vuoto trattino netto manca nome diverso refuso non combacia abbinamento ricaricare stipendi dr",
  },
  {
    id: "sti-versato",
    sezione: "Stipendi",
    titolo: '"Versato in banca" vuoto o basso',
    testo: `Il Versato si riempie con il dettaglio delle distinte stipendi, che la banca non manda col collegamento automatico: sta nel report "Esiti pagamenti" di YouBusiness, da esportare e caricare in Storico estratti. La banca lo conserva circa 90 giorni — l'export va fatto una volta al mese, altrimenti il dettaglio dei mesi vecchi non è più recuperabile.`,
    chiavi: "versato in banca vuoto esiti pagamenti youbusiness export 90 giorni distinta",
  },
  {
    id: "sti-sedi",
    sezione: "Stipendi",
    titolo: 'Sedi, etichette e la voce "Altri"',
    testo: `Il filtro per sede usa l'etichetta del file paghe quando c'è; quando il file non la porta (è successo con dei mesi arrivati senza colonna appalto), il portale ricava la sede dagli altri mesi della stessa persona o dall'appalto in anagrafica. Chi resta senza nulla finisce nella voce "Altri".

Se vedi troppa gente in "Altri", la cura è compilare il campo Appalto in anagrafica (Amministrazione → Appalti dipendenti): sistemato lì, si sistema ovunque.`,
    chiavi: "sede etichetta altri filtro appalto anagrafica",
  },
  {
    id: "sti-mensilita",
    sezione: "Stipendi",
    titolo: "Le mensilità (12, 13, 14…)",
    testo: `Il numero di mensilità per dipendente viene stimato dai ratei del file paghe, ma si può correggere a mano nella tabella: il valore dichiarato vince sulla stima e sopravvive ai ricaricamenti dei file.`,
    chiavi: "mensilità tredicesima quattordicesima ratei stima",
  },

  // --- Fiscale ---------------------------------------------------------------------
  {
    id: "fis-vista",
    sezione: "Fiscale",
    titolo: "La tab Fiscale in due parole",
    testo: `È lo scadenzario fiscale vivo: ogni riga una scadenza, con voce, importo, data di pagamento e lo stato pagata/da pagare. Le viste pivot mostrano il da pagare e il pagato per mese e per voce, la vista Piani raggruppa le rateazioni (con l'avanzamento rata per rata), e "Da registrare" tiene le posizioni ancora da definire.

Ormai la fonte è il portale: le scadenze si inseriscono, si modificano e si spuntano qui, e ogni tabella si esporta in CSV.`,
    chiavi: "fiscale scadenzario scadenze pivot piani rate da registrare sabrina",
  },
  {
    id: "fis-rate",
    sezione: "Fiscale",
    titolo: "Rate o corrente: chi decide",
    testo: `Ogni scadenza è classificata come rateazione (piani di rientro, avvisi bonari, dilazioni) o costo corrente (l'ordinario del mese: IVA, INPS, F24 correnti). La classificazione è automatica dalle caratteristiche della scadenza e alimenta le due voci dei Flussi — Costo fiscale rate e Costo fiscale corrente. Le scadenze non pagate dei mesi passati scivolano sul mese corrente: non spariscono.`,
    chiavi: "rate corrente rateazione dilazione classificazione costo fiscale flussi",
  },

  // --- Storico estratti ---------------------------------------------------------------
  {
    id: "sto-import",
    sezione: "Storico estratti",
    titolo: "Cosa si carica in Storico estratti",
    testo: `Gli import manuali: estratti conto storici, il report Esiti pagamenti (il dettaglio delle distinte, da esportare ogni mese), archivi di altri conti. Ogni import mostra prima l'anteprima con i doppioni scartati: i duplicati non entrano due volte.

Se un import grosso si interrompe a metà, rilancialo con lo stesso file: quello già scritto viene riconosciuto e salta, entra solo il mancante.`,
    chiavi: "storico estratti import esiti pagamenti estratto conto doppioni interrotto",
  },
  {
    id: "sto-distinte",
    sezione: "Storico estratti",
    titolo: "Distinte e aggancio ai movimenti",
    testo: `Le distinte caricate si agganciano ai movimenti di banca per importo e data, anche quando la banca spezza una distinta in più addebiti: il portale cerca la combinazione che quadra al centesimo e la propone. Il badge con gli omini sul movimento apre il dettaglio dei beneficiari con il riepilogo per appalto.`,
    chiavi: "distinta aggancio tranche spezzata beneficiari appalto badge",
  },

  // --- Amministrazione -----------------------------------------------------------------
  {
    id: "amm-stato",
    sezione: "Amministrazione",
    titolo: "Stato integrazione e log",
    testo: `La pagina Amministrazione mostra lo stato del collegamento a SharePoint, l'esito delle ultime operazioni (il log) e il check di salute con il self-test. Quando qualcosa sembra non aggiornarsi, il log è il primo posto dove guardare: dice cosa è passato, quando, e con che esito.`,
    chiavi: "amministrazione stato integrazione log salute test sharepoint errori",
  },
  {
    id: "amm-banca",
    sezione: "Amministrazione",
    titolo: "Il collegamento banca e il consenso",
    testo: `Il collegamento al conto è in sola lettura e il consenso va rinnovato circa ogni 89 giorni approvando dall'app della banca (la scadenza è scritta nella card). Il rinnovo con Banco BPM richiede DUE giri uguali (login, Autorizza, approvazione nell'app, "Torna al merchant"): il secondo serve a Enable Banking per lo storico oltre 90 giorni e si apre da solo come una nuova pagina di login della banca. Solo dopo il secondo "Torna al merchant" il portale torna con la scelta del conto (scegliere l'IBAN che finisce con 3681) e poi "Sincronizza ora". Fermarsi dopo il primo giro lascia il consenso in banca ma il portale senza sessione (è successo il 28/10/2026). Il codice vale pochi minuti e viene riconosciuto da qualunque finestra del portale, app installata compresa; l'esito del completamento resta scritto nella card. "Sincronizza ora" forza un giro; il "recupero da data" ripesca giorni saltati, per esempio movimenti contabilizzati tardi dalla banca.`,
    chiavi:
      "consenso 89 giorni rinnova sincronizza ora recupero data banca collegamento torna al merchant codice ritorno due giri secondo login",
  },
  {
    id: "amm-anagrafica",
    sezione: "Amministrazione",
    titolo: "Anagrafica dipendenti e appalti",
    testo: `L'import massivo incolla i dati da Excel; quello degli appalti aggancia ogni dipendente alla sua sede (e alimenta la regola salari e le sedi della tab Stipendi). Sempre prima l'anteprima, poi l'import. I PIN si assegnano scrivendo il valore in chiaro nella colonna PIN: il portale lo protegge da solo al primo uso.`,
    chiavi: "anagrafica import dipendenti appalti pin excel incolla",
  },
] as const;
