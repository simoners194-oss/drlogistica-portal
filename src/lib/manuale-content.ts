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
  "Stipendi",
  "Fiscale",
  "Storico estratti",
  "Amministrazione",
] as const;

export const MANUALE: readonly VoceManuale[] = [
  // --- In generale -----------------------------------------------------------
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

  // --- Presenze ---------------------------------------------------------------
  {
    id: "pre-timbrare",
    sezione: "Presenze e timbrature",
    titolo: "Timbrare: i quattro tasti",
    testo: `Entrata quando inizi, Inizio pausa quando stacchi, Fine pausa quando riprendi, Uscita quando finisci. Puoi fare più pause nello stesso giorno e anche due turni (esci e rientri: il portale ti fa timbrare una nuova entrata).

Il turno di notte funziona: se entri alle 22 ed esci alle 2 la giornata resta unita e le ore sono contate giuste.`,
    chiavi: "entrata uscita pausa timbratura turno notturno spezzato",
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

Una riga "di riserva" non è un errore del portale: è un movimento per cui non esiste ancora una regola. La soluzione stabile è creare la regola e poi premere "Riapplica tutte" — sistemare solo la riga a mano funziona, ma al prossimo movimento uguale sei daccapo.`,
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

Quindi: primo, Ctrl+F5 (magari è la pagina vecchia); secondo, guarda quando è passato l'ultimo giro; terzo, se serve subito lancia il giro su richiesta. Se dopo un giro completo lo stato ancora non torna, allora sì che è una segnalazione.`,
    chiavi: "aruba pagata portale aperta stato giro pc sincronizza incasso non aggiornato",
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
    id: "flu-lettura",
    sezione: "Flussi di cassa",
    titolo: "Come si legge la tabella dei Flussi",
    testo: `Ogni colonna è un mese (o una settimana), la prima colonna è lo Scaduto: quello che era già in ritardo oggi. Le entrate vengono dalle scadenze delle fatture attive aperte, le uscite dalle passive, più le righe che le fatture non conoscono: Stipendi, Costo fiscale rate, Costo fiscale corrente, Altre spese. In fondo il saldo per colonna.

Il bottone Cumulato trasforma le colonne in progressivo: ogni mese somma tutto quello che c'è fino a quel punto, Scaduto compreso.`,
    chiavi: "flussi cassa colonne mese settimana scaduto saldo cumulato lettura",
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
    testo: `Stipendi, Costo fiscale rate, Costo fiscale corrente e Altre spese si riempiono da sole: gli stipendi dai netti caricati e dalle spunte Pagato, le voci fiscali dallo scadenzario, le Altre spese dalla media degli ultimi due mesi di costi generali non fatturati. Il simbolo ≈ segnala il valore automatico.

Cliccando sul nome della voce si apre il dettaglio di cosa c'è dentro; nelle Altre spese puoi includere o escludere le singole tipologie con le spunte, e la media si aggiorna. Se scrivi un importo a mano nella cella, il tuo numero vince sull'automatico per quel mese — per tornare all'automatico svuota la cella.`,
    chiavi: "voci stipendi fiscale altre spese drill click dettaglio automatico manuale cella",
  },
  {
    id: "flu-stipendi-scaduto",
    sezione: "Flussi di cassa",
    titolo: "Stipendi arretrati nello Scaduto",
    testo: `Quando in un mese di pagamento arrivato (compreso quello corrente) restano stipendi non segnati pagati, il residuo compare nella colonna Scaduto della riga Stipendi: la cella del mese mostra i pagati, lo Scaduto il resto, e la somma torna al totale. Funziona dove la spunta Pagato è in uso nella tab Stipendi; i mesi vecchi senza spunte non contano, perché sono storia già regolata in banca.`,
    chiavi: "stipendi scaduto arretrati non pagati residuo",
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
    testo: `Il collegamento al conto è in sola lettura e il consenso va rinnovato circa ogni 89 giorni approvando dall'app della banca (la scadenza è scritta nella card). "Sincronizza ora" forza un giro; il "recupero da data" ripesca giorni saltati, per esempio movimenti contabilizzati tardi dalla banca.`,
    chiavi: "consenso 89 giorni rinnova sincronizza ora recupero data banca collegamento",
  },
  {
    id: "amm-anagrafica",
    sezione: "Amministrazione",
    titolo: "Anagrafica dipendenti e appalti",
    testo: `L'import massivo incolla i dati da Excel; quello degli appalti aggancia ogni dipendente alla sua sede (e alimenta la regola salari e le sedi della tab Stipendi). Sempre prima l'anteprima, poi l'import. I PIN si assegnano scrivendo il valore in chiaro nella colonna PIN: il portale lo protegge da solo al primo uso.`,
    chiavi: "anagrafica import dipendenti appalti pin excel incolla",
  },
] as const;
