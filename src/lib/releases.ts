// DR Portal — Registro release (Novità).
//
// Ogni nuova versione aggiunge un blocco in testa a `RELEASES`. La pagina
// Novità visualizza i blocchi in ordine cronologico decrescente (dal più
// recente al più vecchio). Predisposto per l'evoluzione futura: basta
// prependere un nuovo oggetto `Release` per rendere disponibile la
// changelog nella UI, senza modificare la pagina.

export type ReleaseTag = "feature" | "improvement" | "fix" | "security";

/** Destinatari di una novità: la pagina Novità mostra a ogni utente solo le
 *  voci che lo riguardano.
 *  - "tutti"     — visibile a chiunque (default)
 *  - "gestione"  — responsabili, operatori, autorizzatori e amministratori
 *  - "direzione" — solo direzione (DR005) e amministratori */
export type ReleaseAudience = "tutti" | "gestione" | "direzione";

export interface ReleaseEntry {
  title: string;
  description?: string;
  tag?: ReleaseTag;
  audience?: ReleaseAudience;
}

export interface Release {
  version: string;
  /** Data di rilascio in formato ISO (YYYY-MM-DD). */
  date: string;
  /** Nome sintetico della release, opzionale. */
  codename?: string;
  /** A cura di (curatore/autore della release), opzionale. */
  author?: string;
  /** Elenco novità pubblicate con questa versione. */
  entries: ReleaseEntry[];
}

export const RELEASES: readonly Release[] = [
  {
    version: "1.41.0",
    date: "2026-08-13",
    codename: "Le tranche si agganciano da sole",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Distinte pagate in piu' addebiti: aggancio automatico al centesimo",
        description:
          "Quando la banca addebita una distinta in piu' movimenti (es. gli stipendi in 7 addebiti), il portale trova da solo la combinazione di uscite che somma esattamente alla distinta \u2014 al centesimo \u2014 e la aggancia senza bisogno di click: le disposizioni per beneficiario compaiono subito nei Movimenti. La ricerca considera fino a 30 movimenti candidati entro 10 giorni dalla data della distinta. Se la quadratura perfetta non esiste, la riga della distinta mostra la migliore combinazione trovata con lo scarto in euro, per capire al volo cosa manca (un addebito non ancora arrivato dalla banca, una commissione, una distinta diversa).",
      },
    ],
  },
  {
    version: "1.40.0",
    date: "2026-08-13",
    codename: "Regole col segno",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le regole distinguono entrate e uscite, e c'e' il jolly",
        description:
          "Ogni regola puo' valere solo per le entrate o solo per le uscite (campo Segno importo), e il pattern jolly * prende tutto cio' che le regole piu' specifiche non hanno gia' classificato \u2014 il jolly agisce sempre per ultimo. Nella matita del movimento c'e' inoltre \u201cPerche' questa classificazione?\u201d: l'elenco delle regole che colpiscono quel movimento, col termine esatto che ha fatto scattare il match.",
      },
    ],
  },
  {
    version: "1.39.0",
    date: "2026-08-13",
    codename: "La distinta si apre da sola",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "I pagamenti cumulativi si mostrano gia' divisi per beneficiario",
        description:
          "Nei Movimenti (e nelle viste dell'Overview) il movimento agganciato a una distinta non compare piu' come totale unico: al suo posto ci sono le singole disposizioni, una per beneficiario, gia' classificate come Pagamento Salario con l'appalto della persona e il badge della distinta \u2014 dal quale si risale sempre al pagamento cumulativo di provenienza. Se la banca ha addebitato una cifra diversa dalla somma delle disposizioni, una riga di resto fa quadrare i totali al centesimo. L'archivio bancario sottostante resta intatto.",
      },
    ],
  },
  {
    version: "1.38.0",
    date: "2026-08-13",
    codename: "Regole anche per le fatture",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le fatture si classificano con le regole, come i movimenti",
        description:
          "Nell'archivio fatture (attive e passive, separate) c'e' la sezione Regole di classificazione: “se il cliente include … E/O l'oggetto include … allora tipologia, sottocategoria, allocazioni, servizio” — con l'operatore AND/OR a scelta e lo stesso vocabolario a cascata delle regole dei movimenti. Le regole si modificano sul posto e valgono subito su colonne, filtri ed export.",
      },
    ],
  },
  {
    version: "1.37.0",
    date: "2026-08-13",
    codename: "Fatture che si spiegano",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Oggetto fattura e Descrizione in colonna, su attive e passive",
        description:
          "Due colonne nuove nell'archivio fatture: l'Oggetto fattura (la causale scritta in testa al documento) e la Descrizione (le voci delle righe). Filtrabili come tutte le altre, presenti nell'export CSV e alimentate dal sync automatico. Il mese di competenza ha inoltre la regola di ripiego impostabile: giorno 15, mese successivo all'emissione o mese di emissione. E i collegamenti nota di credito registrati su Aruba entrano da soli nell'archivio.",
      },
    ],
  },
  {
    version: "1.36.0",
    date: "2026-08-13",
    codename: "Le NC si collegano da sole",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "I collegamenti nota di credito arrivano dal gestionale Aruba",
        description:
          "Le associazioni NC-fattura gia' registrate su Aruba (la colonna Documenti collegati) vengono estratte automaticamente e portate nell'archivio del portale: le note di credito ancora scollegate ricevono il riferimento alla fattura rettificata, senza mai toccare i collegamenti gia' fatti. Residui, stato \u201cannullata da NC\u201d ed eccedenze si aggiornano di conseguenza. Nella stessa serie: liste lunghe nelle regole spezzate automaticamente, note e ricerca istantanea sulle regole, filtri con intervalli di date e importi, simulazione dei termini sul Resoconto ed export Excel di regole e ritardi.",
      },
    ],
  },
  {
    version: "1.35.0",
    date: "2026-08-12",
    codename: "Una distinta, tante tranche",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le distinte si agganciano anche a piu' movimenti",
        description:
          "Quando la banca addebita una distinta stipendi in piu' tranche (piu' movimenti lo stesso giorno con lo stesso riferimento), ora si agganciano tutti, uno alla volta: lo Storico estratti mostra la copertura in tempo reale \u2014 \u201c3 movimenti \u00b7 84.199 di 138.383\u201d \u2014 e il badge dei beneficiari compare su ogni tranche. I candidati proposti sono ordinati sull'importo che manca.",
      },
    ],
  },
  {
    version: "1.34.0",
    date: "2026-08-12",
    codename: "Anche i vecchi trovano posto",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "direzione",
        title: "Appalto a mano per i beneficiari fuori anagrafica",
        description:
          "Nel dettaglio di una distinta, i beneficiari non riconosciuti (es. ex dipendenti che non si vogliono in anagrafica) hanno una casella per assegnare l'appalto a mano. L'assegnazione si ricorda per nome: le distinte future con lo stesso beneficiario si classificano da sole, e il riepilogo per appalto e il CSV la comprendono. Le distinte senza movimento corrispondente si agganciano a mano scegliendo dai candidati vicini per data (il badge compare anche se la somma non torna al centesimo), e nei Movimenti il filtro \u201cSolo distinte\u201d isola i pagamenti cumulativi con un click.",
      },
    ],
  },
  {
    version: "1.33.0",
    date: "2026-08-12",
    codename: "Nessuno preme piu' niente",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Fatture e banca si sincronizzano da sole a orari fissi",
        description:
          "Il sync delle fatture da Aruba ha ora il suo innesco programmato, come la banca: uno scheduler gratuito chiama il portale al mattino e nel primo pomeriggio e le fatture nuove entrano da sole nell'archivio. Una guardia impedisce le esecuzioni ravvicinate. Il pulsante manuale resta per i giri fuori orario.",
      },
    ],
  },
  {
    version: "1.32.0",
    date: "2026-08-12",
    codename: "La cassa di domani",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Prefatturazione: il pianificato entra nella Previsione",
        description:
          "Nel Resoconto si registrano canoni ricorrenti e fatture pianificate non ancora emesse (controparte, importo, mese di inizio, ricorrenza mensile o una tantum): compaiono nella Previsione come righe dedicate e il saldo mensile le comprende. Quando la fattura vera arriva \u2014 anche dal sync automatico \u2014 la riga pianificata di quel mese si considera coperta da sola: mai doppi conteggi.",
      },
    ],
  },
  {
    version: "1.31.0",
    date: "2026-08-12",
    codename: "Si guarda anche avanti",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Scadenze future e previsione di cassa nel Resoconto",
        description:
          "Nelle tabelle dei ritardi, il filtro \u201cAnche in scadenza entro N giorni\u201d aggiunge ai ritardi le fatture in arrivo (in blu, \u201cfra N gg\u201d): si vede cosa scade dopodomani accanto a cosa e' gia' scaduto. Sotto, la nuova Previsione: incassi attesi, pagamenti attesi e saldo per i prossimi sei mesi, calcolati dalle scadenze delle fatture aperte, con lo scaduto in prima colonna.",
      },
    ],
  },
  {
    version: "1.30.0",
    date: "2026-08-12",
    codename: "Piu' righe, un gesto",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Selezione multipla con azioni in blocco",
        description:
          "Su Movimenti, Fatture e Regole ci sono le caselle di selezione (anche l'intera pagina con un click sull'intestazione). Sui movimenti selezionati si corregge in blocco tipologia, sottocategoria e allocazioni (i campi vuoti non vengono toccati); sulle fatture si classifica in blocco mese di competenza e servizio; le regole selezionate si eliminano insieme. In tutti i casi i valori proposti arrivano dal vocabolario delle regole, a cascata. Nel dettaglio di una distinta, il pulsante Esporta CSV scarica lo spaccato: beneficiari, importi, dipendente riconosciuto e appalto.",
      },
    ],
  },
  {
    version: "1.29.0",
    date: "2026-08-12",
    codename: "Un vocabolario solo",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "tutti",
        title: "Le richieste di acquisto parlano la lingua della Finanza",
        description:
          "Nel modulo acquisti, categoria e dettaglio ora propongono le voci del vocabolario delle regole di classificazione (tipologia e sottocategorie apprese): un solo vocabolario in tutto il portale. L'elenco voci precedente resta conservato come riserva.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Nel Resoconto si vede come si compone ogni importo",
        description:
          "Nelle tabelle dei ritardi c'e' la colonna Composizione: per ogni fattura si legge il totale documento, la nota di credito collegata (con il numero) e il netto \u2014 cos\u00ec un residuo ridotto da una NC non sembra piu' un errore. Nei filtri, intervalli di date e importi anche sulle fatture; nelle regole, note e fusione automatica dei doppioni.",
      },
    ],
  },
  {
    version: "1.28.0",
    date: "2026-08-12",
    codename: "Filtri su misura",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "direzione",
        title: "Range di date e importi, correzione a cascata e scorrimento rapido",
        description:
          "Nei Movimenti e nelle Fatture si filtra per intervallo di date (dal \u2192 al) e di importi (da \u2192 a; un solo estremo vale come maggiore/minore di). La finestra \u201cCorreggi movimento\u201d ora propone menu a discesa a cascata con le voci delle regole apprese, come il modulo delle regole. E il tasto laterale in basso a destra porta in fondo alla pagina con un colpo \u2014 e da la' riporta in cima.",
      },
    ],
  },
  {
    version: "1.27.0",
    date: "2026-08-12",
    codename: "Regole che si tengono in ordine da sole",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "direzione",
        title: "Regole: ricerca istantanea e fusione automatica dei doppioni",
        description:
          "Nell'elenco delle regole basta iniziare a scrivere per trovare subito quella giusta (per nome, termine o categoria), con le sezioni che si aprono da sole sul risultato. E se si inserisce una regola identica a una gia' esistente, i termini nuovi si uniscono automaticamente a quella \u2014 con un avviso \u2014 invece di creare un doppione. Gli elenchi lunghi di nominativi vengono spezzati in regole gemelle senza piu' rifiuti al salvataggio.",
      },
    ],
  },
  {
    version: "1.26.0",
    date: "2026-08-11",
    codename: "Regole senza doppioni",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "direzione",
        title: "Le regole doppie si unificano con un click",
        description:
          "Nella scheda Regole, il pulsante \u201cUnifica doppie\u201d fonde le regole equivalenti \u2014 stessi esiti e stesso criterio di ricerca \u2014 in una sola con l'unione dei termini, elimina le copie e ripulisce i termini ripetuti dentro le singole regole. Il comportamento di classificazione non cambia: cambia solo l'ordine in casa. Le fatture, intanto, si sincronizzano da Aruba senza piu' caricamenti manuali.",
      },
    ],
  },
  {
    version: "1.25.0",
    date: "2026-08-11",
    codename: "Le fatture arrivano da sole",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Sincronizzazione fatture direttamente da Aruba",
        description:
          "Con il collegamento Premium attivo, il pulsante \u201cSincronizza da Aruba\u201d scarica dalle API le fatture nuove \u2014 emesse e ricevute \u2014 e le importa nell'archivio con la stessa pipeline e la stessa protezione anti-doppioni dei caricamenti manuali. La finestra riparte dall'ultima sincronizzazione (con margine di sicurezza) e si puo' allargare indicando i giorni indietro. I caricamenti manuali di XML restano disponibili, ma non sono piu' necessari per l'ordinaria amministrazione.",
      },
    ],
  },
  {
    version: "1.24.0",
    date: "2026-08-11",
    codename: "I costi trovano il loro appalto",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le distinte stipendi si leggono per appalto",
        description:
          "Nel dettaglio di un pagamento cumulativo, sotto l'elenco dei beneficiari, compare il riepilogo per appalto: ogni nome viene riconosciuto in anagrafica dipendenti e il suo importo finisce sull'appalto di assegnazione \u2014 con il conteggio dei non riconosciuti in chiaro. Il costo del personale del mese si legge cos\u00ec gi\u00e0 allocato.",
      },
      {
        tag: "feature",
        audience: "direzione",
        title: "Filtri su tutte le voci dei movimenti e nuovi resoconti",
        description:
          "Nei Movimenti si filtra anche per sottocategoria e per allocazione primaria e secondaria (con la voce \u201c(vuota)\u201d per stanare il non classificato) e ogni colonna della tabella ha il suo imbuto stile Excel, con spunte sui valori e ricerca; il totale in alto segue i filtri. Nell'Overview due nuove viste: \u201cPer regole\u201d (tipologia e sottocategoria per mese o anno) e \u201cPer appalto\u201d (le allocazioni per mese o anno).",
      },
    ],
  },
  {
    version: "1.23.0",
    date: "2026-08-11",
    codename: "Dentro i pagamenti cumulativi",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le distinte di pagamento si aprono e mostrano i beneficiari",
        description:
          "I pagamenti cumulativi (distinta stipendi, ritiro effetti RiBa) non sono piu' scatole chiuse: caricando il report \u201cEsiti pagamenti\u201d dell'home banking, il portale archivia ogni disposizione e la aggancia da sola al movimento bancario corrispondente \u2014 stessa somma, data vicina. Sul movimento compare l'icona con il numero dei beneficiari: un clic apre l'elenco completo con gli importi. I file gia' caricati non si duplicano.",
      },
    ],
  },
  {
    version: "1.22.0",
    date: "2026-08-10",
    codename: "Mese e servizio anche sulle attive",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Le fatture attive hanno mese di competenza e servizio",
        description:
          "Come nel prospetto della direzione, anche l'archivio delle fatture emesse mostra il mese di riferimento e il servizio (appalto) di ogni fattura: colonne filtrabili, modifica col doppio click sulla cella o dal dettaglio, ed export CSV completo. Il mese di competenza \u2014 in entrambe le direzioni \u2014 ora si legge anche dalla causale della fattura (\u201ccompetenze luglio\u201d), prima della regola del giorno 15. Sul Resoconto c'e' la nuova riga con il totale delle attive (incassate, da incassare, di cui in ritardo) e i residui a confronto.",
      },
    ],
  },
  {
    version: "1.21.0",
    date: "2026-08-10",
    codename: "Menu a cascata",
    author: "Simone Russo",
    entries: [
      {
        tag: "improvement",
        audience: "direzione",
        title: "Le regole si compilano scegliendo da menu a discesa",
        description:
          "Tipologia, sottocategoria e allocazioni nel modulo delle regole sono ora veri menu a discesa: si aprono con un click e propongono solo le voci gia' usate nelle regole apprese, in ordine alfabetico e a cascata (scelta la tipologia, le sottocategorie si restringono a quelle coerenti). La voce “Nuova voce…” apre il campo libero per aggiungerne una. Sullo specchietto del Resoconto compare inoltre il sottotitolo “Controllo aggiornamento Aruba”.",
      },
    ],
  },
  {
    version: "1.19.0",
    date: "2026-08-07",
    codename: "Salari sull'appalto giusto",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "I salari finiscono sull'appalto del dipendente",
        description:
          "In anagrafica dipendenti c'è il campo Appalto: quando il riconoscimento automatico classifica un salario, assegna anche l'allocazione — secondaria = appalto del dipendente, primaria Appalto o Costi generali (per gli uffici) — senza mai sovrascrivere scelte già fatte. L'export CSV dei movimenti riporta ora anche sottocategoria, allocazioni e conto.",
      },
    ],
  },
  {
    version: "1.18.0",
    date: "2026-08-07",
    codename: "Un archivio, più conti",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Ogni movimento sa a quale conto appartiene",
        description:
          "Nell'archivio convivono più conti correnti: ora i movimenti del collegamento bancario sono etichettati automaticamente, i lotti caricati a mano si assegnano con un click dallo Storico estratti, e nei Movimenti c'è il filtro per conto — così il totale filtrato si confronta con il saldo del conto giusto.",
      },
    ],
  },
  {
    version: "1.17.0",
    date: "2026-08-07",
    codename: "Regole più potenti",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Riconoscimento automatico degli stipendi",
        description:
          "I bonifici in uscita verso i dipendenti vengono classificati da soli come Pagamento Salario, con il nome corretto dall'anagrafica — anche quando la banca lo riporta troncato o invertito. Un pulsante applica la regola a tutto l'archivio.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Regole raggruppate e ricerca multipla",
        description:
          "L'elenco delle regole è raggruppato per tipologia e si apre al tocco. La ricerca dei movimenti e le regole accettano più termini separati da virgola, e sopra l'elenco compare il totale di ciò che è filtrato, con entrate e uscite separate. Le voci scritte a mano nelle regole restano tra i suggerimenti.",
      },
    ],
  },
  {
    version: "1.15.0",
    date: "2026-08-06",
    codename: "Classificazione estesa",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Sottocategoria e allocazioni sui movimenti",
        description:
          "Ogni movimento può avere una sottocategoria (es. Pernottamento, Pasto, Trasporto) e due allocazioni: primaria (Costi generali / Appalto) e secondaria (per ufficio o commessa). Si impostano dalle regole o correggendo il singolo movimento; scrivendo un valore nuovo, la voce nasce senza configurazioni.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Il dovuto delle fatture passive è il netto a pagare",
        description:
          "Quando una fattura ricevuta dichiara un netto a pagare diverso dal totale (ritenute, bolli), l'importo in evidenza è quello da pagare davvero, con il totale documento riportato sotto. Scadenze e ritardi seguono lo stesso criterio.",
      },
    ],
  },
  {
    version: "1.13.5",
    date: "2026-08-05",
    codename: "Resoconto su misura",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Viste salvate nel Resoconto",
        description:
          "La configurazione dei filtri (clienti, fornitori, estratto conto) si salva con un nome e si richiama con un click. Aggiunte l'opzione «Nessuno» per escludere un lato del confronto, il filtro diretto sull'estratto conto e il sollecito direttamente dagli specchietti dei ritardi.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Storni riconosciuti e ritardo in colonna",
        description:
          "Le note di credito di storno senza riferimento si agganciano da sole alla fattura di pari importo della stessa controparte. Negli elenchi delle fatture c'è la colonna «Ritardo [gg]», filtrabile come le altre.",
      },
      {
        tag: "improvement",
        audience: "gestione",
        title: "Anomalie scartabili e richieste con nome",
        description:
          "Le anomalie delle timbrature — in particolare le informative — si possono scartare dopo la verifica, e restano tracciate. Nella coda delle richieste compare nome e cognome accanto al codice.",
      },
      {
        tag: "fix",
        audience: "tutti",
        title: "Orari sempre in ora italiana",
        description:
          "Tutti gli orari mostrati dal portale sono nel fuso italiano, qualunque sia l'impostazione del telefono: l'ora delle timbrature coincide ora in ogni schermata.",
      },
    ],
  },
  {
    version: "1.9.0",
    date: "2026-08-05",
    codename: "Resoconto e regole di pagamento",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Resoconto per controparte",
        description:
          "Nuova scheda in Finanza che confronta l'estratto conto con il netto delle fatture (incassato attive meno pagato passive) per una o più controparti, con esito verde quando i conti tornano. Filtri per cliente, fornitore ed estratto conto, opzione «Nessuno» per escludere un lato del confronto e ritardi in evidenza nelle due direzioni, con fasce di anzianità.",
      },
      {
        tag: "feature",
        audience: "direzione",
        title: "Termini di pagamento con regole per oggetto",
        description:
          "I giorni di pagamento possono ora dipendere dall'oggetto della fattura: ad esempio, per lo stesso fornitore, le fatture di locazione a vista e tutte le altre coi termini ordinari. Le regole si impostano dal pannello Termini con parole chiave.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Storni riconosciuti automaticamente",
        description:
          "Una nota di credito che non dichiara la fattura rettificata viene ora agganciata da sola quando l'importo corrisponde esattamente a una sola fattura della stessa controparte: la fattura stornata esce dai ritardi senza interventi manuali.",
      },
      {
        tag: "improvement",
        audience: "direzione",
        title: "Fatture più rapide e filtri con ricerca",
        description:
          "L'apertura della pagina Fatture è molto più veloce e tutti i menu a tendina dei filtri hanno un campo di ricerca con completamento. La ricerca dei movimenti mostra il motivo per cui una riga corrisponde, anche quando il testo è nella descrizione.",
      },
      {
        tag: "improvement",
        audience: "gestione",
        title: "Assenze visibili nel resoconto del giorno",
        description:
          "Il resoconto giornaliero delle timbrature segnala con un'etichetta chi è in malattia o in ferie: l'assenza spiegata non compare più come «nessuna timbratura».",
      },
      {
        tag: "improvement",
        audience: "gestione",
        title: "Rendiconto per giorno singolo e filtro sede",
        description:
          "Il rendiconto può essere consultato anche per un giorno preciso, e la vista di sede permette a chi supervisiona più sedi di filtrarle una alla volta.",
      },
    ],
  },
  {
    version: "1.8.0",
    date: "2026-07-28",
    codename: "Scadenzario e riconciliazione",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Fatture emesse e ricevute con scadenzario",
        description:
          "La sezione Finanza gestisce ora le fatture elettroniche: caricando l'archivio XML esportato dal servizio di fatturazione, il portale distingue automaticamente le fatture emesse dalle ricevute, legge le scadenze di pagamento indicate in fattura e presenta crediti e debiti — aperti, in ritardo e saldati — riepilogati per cliente e per fornitore.",
      },
      {
        tag: "feature",
        audience: "direzione",
        title: "Riconciliazione con il conto corrente",
        description:
          "Incassi e pagamenti del conto vengono abbinati alle fatture: in automatico quando la causale cita il numero o l'importo corrisponde, con l'imputazione a scalare per acconti e saldi cumulativi, oppure manualmente per i casi particolari. Ogni abbinamento resta tracciato e può essere rimosso.",
      },
      {
        tag: "feature",
        audience: "direzione",
        title: "Regole di classificazione personalizzate",
        description:
          "Le correzioni apportate ai movimenti bancari — tipologia e controparte — possono essere salvate come regole permanenti: si applicano automaticamente agli estratti conto futuri e, su richiesta, all'archivio esistente. L'overview delle spese è ora filtrabile per tipologia e per anno.",
      },
      {
        tag: "feature",
        title: "Annulla ultima timbratura",
        description:
          "In caso di tocco sbagliato, entro 5 minuti è possibile annullare la propria ultima timbratura: i pulsanti corretti tornano subito disponibili.",
      },
      {
        tag: "feature",
        audience: "gestione",
        title: "Resoconto giornaliero delle timbrature",
        description:
          "In Gestione timbrature si selezionano sede e giorno per vedere tutti i dipendenti con le rispettive timbrature — con evidenza di chi non ha registrato nulla — e correggere direttamente: si elimina l'evento errato e si inseriscono quelli mancanti.",
      },
      {
        tag: "improvement",
        audience: "gestione",
        title: "Dashboard più rapida",
        description:
          "Toccando una sede nella sintesi, la dashboard si filtra e si posiziona direttamente sul riepilogo di quella sede.",
      },
      {
        tag: "fix",
        title: "Traduzioni corrette su Chrome",
        description:
          "Risolte le etichette errate mostrate da alcuni telefoni, dovute alla traduzione automatica di Chrome: la lingua della pagina è ora dichiarata correttamente. Se il problema persiste su un dispositivo, impostare una volta «Non tradurre mai questo sito».",
      },
    ],
  },
  {
    version: "1.7.0",
    date: "2026-07-22",
    codename: "Finanza",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        audience: "direzione",
        title: "Finanza — estratto conto per la direzione",
        description:
          "Nuova sezione riservata alla direzione: import dell'estratto conto bancario da Excel con classificazione automatica dei movimenti (tipologia, cliente, riferimenti fattura), overview degli incassi per cliente e mese, pagina anomalie per sanare a mano i casi dubbi. I ricaricamenti scartano automaticamente i movimenti già importati.",
      },
    ],
  },
  {
    version: "1.6.0",
    date: "2026-07-17",
    codename: "English / Italiano",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        title: "Portale bilingue inglese/italiano",
        description:
          "Tutta l'interfaccia è disponibile in inglese e in italiano: si cambia lingua con le bandierine in alto a destra e la scelta viene ricordata sul dispositivo. Lingua predefinita: inglese.",
      },
    ],
  },
  {
    version: "1.5.0",
    date: "2026-07-16",
    codename: "Procurement e buste paga",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        title: "Procurement — richieste di acquisto",
        description:
          "Le risorse delle sedi storiche inviano richieste di acquisto con voce e dettaglio; l'approvazione è riservata alla direzione, con coda dedicata ed esportazione Excel.",
      },
      {
        tag: "feature",
        title: "Buste paga automatiche",
        description:
          "Caricamento di tutti i cedolini in un colpo solo: l'abbinamento al dipendente avviene dal codice fiscale nel nome del file, con anteprima. Ognuno riceve il documento e una notifica.",
      },
      {
        tag: "feature",
        title: "Notifiche push sul telefono",
        description:
          "Le comunicazioni e i nuovi documenti personali arrivano come notifica sul telefono, anche ad app chiusa. Attivazione con un tocco dalla pagina Comunicazioni.",
      },
      {
        tag: "feature",
        title: "Rendiconto per settimana",
        description:
          "Il rendiconto è ora filtrabile anche per settimana fiscale dell'anno o per settimana del mese (lun-dom), oltre che per mese.",
      },
      {
        tag: "improvement",
        title: "Voci di spesa dettagliate e residui a colpo d'occhio",
        description:
          "Rimborsi e acquisti usano voci macro con sotto-dettaglio gestibili dall'azienda; gli approvatori vedono ferie e permessi residui accanto a ogni richiesta.",
      },
      {
        tag: "improvement",
        title: "Resta connesso",
        description:
          "L'accesso resta attivo alla chiusura dell'app: niente più login a ogni apertura (fino a 30 giorni o al logout).",
      },
      {
        tag: "security",
        title: "PIN protetti",
        description:
          "I PIN non sono più leggibili da chi accede agli archivi: vengono conservati in forma cifrata, con blocco automatico dopo troppi tentativi errati.",
      },
    ],
  },
  {
    version: "1.4.0",
    date: "2026-07-15",
    codename: "Documenti e Comunicazioni",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        title: "Documenti dipendente",
        description:
          "Contratti, buste paga, DPI, certificati corsi e altri documenti, personali o generali per sede. Ogni dipendente ritrova i propri direttamente nel portale.",
      },
      {
        tag: "feature",
        title: "Comunicazioni interne",
        description:
          "Bacheca per riunioni e avvisi, per tutte le sedi o una specifica, con allegato e presa visione: il dipendente conferma la lettura e chi pubblica vede chi ha letto.",
      },
      {
        tag: "feature",
        title: "Avvisi in tempo reale",
        description:
          "Un promemoria compare quando ci sono comunicazioni da leggere o nuove novità dell'applicazione.",
      },
      {
        tag: "improvement",
        title: "Etichetta straordinario",
        description:
          'Nella dashboard la voce "Oltre 8 ore" è ora "In straordinario", più coerente con il conteggio a ore settimanali.',
      },
    ],
  },
  {
    version: "1.3.0",
    date: "2026-07-15",
    codename: "Rimborsi e nuove sedi",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
        title: "Rimborsi spese con giustificativo",
        description:
          "Nuovo tipo di richiesta per i rimborsi: alleghi la foto o il PDF dello scontrino/fattura e l'approvatore vede importo, tipologia e documento direttamente nella coda di approvazione.",
      },
      {
        tag: "feature",
        title: "Richieste decise sempre consultabili",
        description:
          "Nuova vista delle richieste approvate e rifiutate, con il giustificativo sempre a portata di clic e filtri per stato, sede, periodo e dipendente. Visibile anche all'operatore.",
      },
      {
        tag: "feature",
        title: "Supporto a più sedi",
        description:
          "Il portale gestisce ora un numero qualsiasi di sedi che timbrano: presenze, correzioni e filtri si adattano automaticamente alle sedi caricate.",
      },
      {
        tag: "feature",
        title: "Import massivo dei dipendenti",
        description:
          "Caricamento in blocco dell'anagrafica da Excel/CSV con anteprima di controllo prima della scrittura, direttamente da Amministrazione.",
      },
      {
        tag: "improvement",
        title: "Supervisione per sede e nuovo Inquadramento",
        description:
          "Ogni supervisore vede e autorizza le richieste di propria competenza; aggiunto il campo Inquadramento nell'anagrafica dipendente.",
      },
    ],
  },
  {
    version: "1.2.0",
    date: "2026-07-13",
    codename: "Rendiconto",
    entries: [
      {
        tag: "feature",
        title: "Rendiconto mensile",
        description:
          "Riepilogo ore per dipendente: ore lavorate, straordinari (calcolati dalle timbrature e autorizzati dalle richieste), permessi, ferie e malattie, con filtri per mese, sede e dipendente.",
      },
    ],
  },
  {
    version: "1.1.0",
    date: "2026-07-13",
    codename: "Richieste e Supervisione",
    entries: [
      {
        tag: "feature",
        title: "Modulo Richieste",
        description:
          "Invio di ferie, permessi, straordinari, smart working, malattia e reperibilità, con flusso di approvazione e comunicazioni.",
      },
      {
        tag: "feature",
        title: "Approvazioni",
        description:
          "Coda di approvazione per gli autorizzatori: approva o respingi (con motivazione) le richieste inviate.",
      },
      {
        tag: "feature",
        title: "Gestione timbrature (operatore)",
        description:
          "Inserimento manuale di timbrature e turni interi, con filtro per sede. Le manuali sono tracciate per la supervisione.",
      },
      {
        tag: "feature",
        title: "Rilevazione anomalie",
        description:
          "Segnalazione automatica delle giornate con turno o pausa non chiusi, con correzione rapida.",
      },
      {
        tag: "feature",
        title: "Supervisione",
        description:
          "Report delle richieste approvate (filtri per sede, periodo e dipendente) e vista delle timbrature manuali.",
      },
      {
        tag: "improvement",
        title: "Visibilità e ruoli",
        description:
          "Nuovi attributi dipendente (visibilità, autorizzazione, operatore, ore settimanali) per un controllo più fine.",
      },
      {
        tag: "fix",
        title: "Integrazione più robusta",
        description:
          "Ritentativi automatici sugli errori temporanei del connettore, per evitare schermate vuote.",
      },
    ],
  },
  {
    version: "1.0.0",
    date: "2026-07-03",
    codename: "Messa in esercizio",
    entries: [
      {
        tag: "feature",
        title: "Nuovo modulo Presenze",
        description:
          "Timbrature con macchina a stati, timeline giornaliera e chiusura automatica della giornata.",
      },
      {
        tag: "feature",
        title: "Dashboard Responsabili",
        description: "Vista live in sola lettura di tutte le sedi con KPI e dettaglio dipendente.",
      },
      {
        tag: "feature",
        title: "Login con Codice e PIN",
        description:
          "Autenticazione unificata per Dipendenti, Responsabili e Amministratori di sistema.",
      },
      {
        tag: "feature",
        title: "Integrazione Microsoft SharePoint",
        description: "Dati letti e scritti direttamente sulle liste Dipendenti e Timbrature.",
      },
      {
        tag: "feature",
        title: "Dashboard live Fiano Romano e San Giuliano",
        description: "Monitoraggio in tempo reale delle sedi operative DR Logistica.",
      },
    ],
  },
] as const;

export const TAG_LABEL: Record<ReleaseTag, string> = {
  feature: "Nuova funzionalità",
  improvement: "Miglioramento",
  fix: "Correzione",
  security: "Sicurezza",
};
