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
  /** Elenco novità pubblicate con questa versione. */
  entries: ReleaseEntry[];
}

export const RELEASES: readonly Release[] = [
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
