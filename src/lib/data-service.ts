// DR Portal — data service
// -----------------------------------------------------------------------------
// Questo modulo è l'unico punto di accesso ai dati di presenza.
// Attualmente restituisce dati mock in memoria; è già strutturato per essere
// sostituito da un client SharePoint / Microsoft Graph senza toccare le UI.
//
// Per collegare SharePoint in futuro basterà:
//   1. Implementare le stesse funzioni (getDipendenti, getSedi, timbra…)
//      leggendo/scrivendo su liste SharePoint via Microsoft Graph.
//   2. Sostituire l'export `dataService` con l'implementazione reale
//      (per esempio in base a import.meta.env.VITE_DR_DATA_SOURCE).
// -----------------------------------------------------------------------------

import { SEDI, type Dipendente, type SedeId, type Timbratura } from "./mock-data";
import { computeOreOggi } from "./presenze-logic";
import {
  spCreateTimbratura,
  spGetDiagnostics,
  spGetSnapshot,
  spRunSelfTest,
  type SpDiagnostics,
} from "./sharepoint.functions";
import type { SpDipendente, SpTimbratura } from "./sharepoint.server";
import { setSpStatus } from "./use-sp-status";

export interface DataService {
  getSedi(): Promise<typeof SEDI>;
  getDipendenti(): Promise<Dipendente[]>;
  getDipendente(id: string): Promise<Dipendente | undefined>;
  timbra(dipendenteId: string, tipo: Timbratura["tipo"]): Promise<Dipendente>;
}

// ---------------------------------------------------------------------------
// Diagnostica integrazione — usata dalla pagina Amministrazione.
// ---------------------------------------------------------------------------

export interface IntegrationStatus {
  mode: "sharepoint";
  dipendentiCaricati: number;
  ultimoAggiornamento: Date | null;
  ultimoErrore: string | null;
  log: IntegrationLogEntry[];
  diagnostics: SpDiagnostics | null;
}

export interface IntegrationLogEntry {
  ts: Date;
  level: "info" | "warn" | "error";
  operation: string;
  message: string;
}

const integrationStatus: IntegrationStatus = {
  mode: "sharepoint",
  dipendentiCaricati: 0,
  ultimoAggiornamento: null,
  ultimoErrore: null,
  log: [],
  diagnostics: null,
};

export function getIntegrationStatus(): IntegrationStatus {
  return {
    ...integrationStatus,
    log: (integrationStatus.diagnostics?.log ?? []).map((e) => ({
      ts: new Date(e.ts),
      level: e.level,
      operation: e.operation,
      message: e.message + (e.durataMs ? ` (${e.durataMs}ms)` : ""),
    })),
  };
}

export async function refreshIntegrationDiagnostics(force = false) {
  try {
    const d = (await spGetDiagnostics({ data: { force } })) as SpDiagnostics;
    integrationStatus.diagnostics = d;
    if (d.error) integrationStatus.ultimoErrore = d.error;
    else integrationStatus.ultimoErrore = null;
  } catch (err) {
    integrationStatus.ultimoErrore = err instanceof Error ? err.message : String(err);
  }
  return integrationStatus.diagnostics;
}

export async function runSpSelfTest() {
  const result = await spRunSelfTest();
  // Refresh diagnostics after test to pick up newly logged events on server.
  await refreshIntegrationDiagnostics(false);
  return result;
}

function markSuccess(count: number, operation = "getDipendenti") {
  integrationStatus.dipendentiCaricati = count;
  integrationStatus.ultimoAggiornamento = new Date();
  integrationStatus.ultimoErrore = null;
  setSpStatus("online");
  void operation;
}

function markError(err: unknown, operation: string) {
  const msg = err instanceof Error ? err.message : String(err ?? "Errore sconosciuto");
  integrationStatus.ultimoErrore = msg;
  integrationStatus.ultimoAggiornamento = new Date();
  setSpStatus("offline", msg);
  void operation;
}

// ---------------------------------------------------------------------------
// Implementazione SHAREPOINT — via server functions e Lovable Connector Gateway.
// Se una chiamata fallisce (config incompleta, credenziali mancanti, errore di
// rete o SharePoint down) si ripiega automaticamente sul mock service in modo
// che l'app resti sempre funzionante. L'errore viene tracciato in
// integrationStatus e mostrato in Amministrazione.
// ---------------------------------------------------------------------------

function mergeDipendentiTimbrature(dips: SpDipendente[], tims: SpTimbratura[]): Dipendente[] {
  const byEmp = new Map<string, SpTimbratura[]>();
  for (const t of tims) {
    const arr = byEmp.get(t.dipendenteId) ?? [];
    arr.push(t);
    byEmp.set(t.dipendenteId, arr);
  }
  return dips.map((d) => {
    const eventi = (byEmp.get(d.id) ?? []).sort((a, b) => a.dataOra.localeCompare(b.dataOra));
    const entrata = eventi.find((e) => e.evento === "entrata");
    const last = eventi[eventi.length - 1];
    let stato: Dipendente["stato"] = "non-timbrato";
    if (last) {
      stato =
        last.evento === "entrata" || last.evento === "fine-pausa"
          ? "presente"
          : last.evento === "inizio-pausa"
            ? "pausa"
            : "uscito";
    }
    const ultimaTimbratura: Timbratura | undefined = last
      ? { tipo: last.evento, ora: last.dataOra }
      : undefined;
    const eventiOggi: Timbratura[] = eventi.map((e) => ({
      tipo: e.evento,
      ora: e.dataOra,
    }));
    const ore = computeOreOggi(eventiOggi);
    return {
      id: d.id,
      nome: d.nome,
      cognome: d.cognome,
      ruolo: d.ruolo || "Dipendente",
      sede: d.sede === "tutte" ? "" : d.sede,
      orarioAtteso: "09:00",
      stato,
      entrataOra: entrata?.dataOra,
      ultimaTimbratura,
      eventiOggi,
      oreLavorateMinuti: ore.oreLavorateMinuti,
      pausaMinuti: ore.pausaMinuti,
      oltreOrarioMinuti: ore.oltreOrarioMinuti,
      straordinariMinuti: ore.oltreOrarioMinuti,
    };
  });
}

// Snapshot filtrato (solo visibili) — alimenta dashboard, elenchi e conteggi.
let cachedSnapshot: Dipendente[] = [];
// Snapshot completo (inclusi i nascosti) — usato SOLO per l'auto-lettura del
// proprio record in getDipendente, mai per viste aggregate.
let cachedFull: Dipendente[] = [];

const sharepointDataService: DataService = {
  async getSedi() {
    return SEDI;
  },
  async getDipendenti() {
    try {
      const snap = (await spGetSnapshot()) as {
        dipendenti: SpDipendente[];
        timbrature: SpTimbratura[];
      };
      // Filtro di VISIBILITÀ (unico punto di verità per tutte le viste
      // operative: dashboard, elenco sede, conteggi, dettaglio, statistiche
      // e — in futuro — report). Escludiamo i dipendenti con visibile=false
      // PRIMA del merge, così ogni aggregato a valle è automaticamente
      // corretto senza toccare le route. Il filtro non tocca l'accesso:
      // l'autenticazione dipende solo da `attivo` (gestita altrove).
      const visibili = snap.dipendenti.filter((d) => d.visibile);
      const list = mergeDipendentiTimbrature(visibili, snap.timbrature);
      // Cache dello snapshot COMPLETO (non filtrato) per l'auto-lettura del
      // proprio record: un utente nascosto deve poter vedere le proprie
      // presenze anche se non compare nelle viste operative (regola 2).
      cachedFull = mergeDipendentiTimbrature(snap.dipendenti, snap.timbrature);
      cachedSnapshot = list;
      markSuccess(list.length, "getDipendenti");
      return list;
    } catch (err) {
      markError(err, "getDipendenti");
      return [];
    }
  },
  async getDipendente(id) {
    // Auto-lettura per id: cerca nello snapshot COMPLETO (include i nascosti),
    // così chi ha visibile=false carica comunque le proprie presenze.
    const hitFull = cachedFull.find((d) => d.id === id);
    if (hitFull) return { ...hitFull };
    const hit = cachedSnapshot.find((d) => d.id === id);
    if (hit) return { ...hit };
    // Cache fredda: popola entrambe le cache e ricerca nella completa.
    await sharepointDataService.getDipendenti();
    return cachedFull.find((d) => d.id === id);
  },
  async timbra(id, tipo) {
    await spCreateTimbratura({ data: { dipendenteId: id, evento: tipo, origine: "Web" } });
    // Refresh dello snapshot, poi cerca nella cache COMPLETA: un utente
    // nascosto (visibile=false) timbra le proprie presenze e deve comunque
    // ricevere il proprio record aggiornato, pur non essendo nella lista
    // filtrata.
    await sharepointDataService.getDipendenti();
    const updated = cachedFull.find((d) => d.id === id);
    if (updated) return updated;
    throw new Error("Timbratura salvata ma dipendente non trovato dopo il refresh.");
  },
};

// L'app usa esclusivamente dati reali da SharePoint. In caso di errore
// vengono restituiti array vuoti e l'errore è visibile in Amministrazione.
export const dataService: DataService = sharepointDataService;

// La variabile `Timbratura` era importata per la vecchia implementazione mock:
// la manteniamo referenziata per evitare warning di unused import in TS strict.
void ({} as Timbratura);

// Filtra per sede senza duplicare la logica nelle pagine.
export function bySede(list: Dipendente[], sede: SedeId) {
  return list.filter((d) => d.sede === sede);
}

// Statistiche aggregate riusate sia dalla dashboard sia dai widget live.
export function aggregate(list: Dipendente[]) {
  const attivi = list.length;
  const presenti = list.filter((d) => d.stato === "presente").length;
  const pausa = list.filter((d) => d.stato === "pausa").length;
  const usciti = list.filter((d) => d.stato === "uscito").length;
  const assenti = list.filter((d) => d.stato === "non-timbrato").length;
  const oltre = list.filter((d) => (d.oltreOrarioMinuti ?? 0) > 0).length;
  const ritardi = list.filter((d) => (d.ritardoMinuti ?? 0) > 0).length;
  return { attivi, presenti, pausa, usciti, assenti, oltre, ritardi };
}

// Stato "visivo" richiesto dalla dashboard live:
//   verde = presente · giallo = in pausa · rosso = assente · blu = oltre orario
export type DisplayStato = "presente" | "pausa" | "assente" | "oltre";

export function displayStato(d: Dipendente): DisplayStato {
  if (d.stato === "pausa") return "pausa";
  if (d.stato === "non-timbrato" || d.stato === "uscito") return "assente";
  if ((d.oltreOrarioMinuti ?? 0) > 0) return "oltre";
  return "presente";
}

export const DISPLAY_LABEL: Record<DisplayStato, string> = {
  presente: "Presente",
  pausa: "In pausa",
  assente: "Assente",
  oltre: "Oltre orario",
};

export const DISPLAY_DOT: Record<DisplayStato, string> = {
  presente: "bg-status-present",
  pausa: "bg-status-break",
  assente: "bg-status-absent",
  oltre: "bg-status-out",
};

// Calcola le ore lavorate oggi. Semplice: from entrataOra a "ora"
// (in un client reale useremo la somma degli intervalli tra entrate/uscite).
export function oreLavorateOggi(d: Dipendente, now = new Date()): string {
  if (!d.entrataOra) return "0h 00m";
  const start = new Date(d.entrataOra).getTime();
  const end =
    d.stato === "uscito" && d.ultimaTimbratura
      ? new Date(d.ultimaTimbratura.ora).getTime()
      : now.getTime();
  const min = Math.max(0, Math.floor((end - start) / 60000));
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}h ${m.toString().padStart(2, "0")}m`;
}
