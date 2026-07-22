// DR Portal — Registro release (Novità).
//
// Ogni nuova versione aggiunge un blocco in testa a `RELEASES`. La pagina
// Novità visualizza i blocchi in ordine cronologico decrescente (dal più
// recente al più vecchio). Predisposto per l'evoluzione futura: basta
// prependere un nuovo oggetto `Release` per rendere disponibile la
// changelog nella UI, senza modificare la pagina.

export type ReleaseTag = "feature" | "improvement" | "fix" | "security";

export interface ReleaseEntry {
  title: string;
  description?: string;
  tag?: ReleaseTag;
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
    version: "1.7.0",
    date: "2026-07-22",
    codename: "Finanza",
    author: "Simone Russo",
    entries: [
      {
        tag: "feature",
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
