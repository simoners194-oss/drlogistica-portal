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

import {
  DIPENDENTI as MOCK_DIPENDENTI,
  SEDI,
  type Dipendente,
  type SedeId,
  type Timbratura,
} from "./mock-data";
import {
  currentIntegrationMode,
  isSharePointConfigured,
  missingSharePointFields,
  sharepointConfig,
  type IntegrationMode,
} from "./sharepoint-config";
import {
  spCreateTimbratura,
  spGetSnapshot,
} from "./sharepoint.functions";
import type { SpDipendente, SpTimbratura } from "./sharepoint.server";

// Stato in memoria (mock). In produzione verrà rimpiazzato da chiamate a
// SharePoint. Cloniamo l'array per poter mutare le timbrature durante la
// sessione.
let state: Dipendente[] = MOCK_DIPENDENTI.map((d) => ({ ...d }));

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
  mode: IntegrationMode;
  dipendentiCaricati: number;
  ultimoAggiornamento: Date | null;
  ultimoErrore: string | null;
  campiMancanti: string[];
  fallbackAttivo: boolean;
  log: IntegrationLogEntry[];
}

export interface IntegrationLogEntry {
  ts: Date;
  level: "info" | "warn" | "error";
  operation: string;
  message: string;
}

const integrationStatus: IntegrationStatus = {
  mode: currentIntegrationMode(),
  dipendentiCaricati: 0,
  ultimoAggiornamento: null,
  ultimoErrore: null,
  campiMancanti: missingSharePointFields(),
  fallbackAttivo: false,
  log: [],
};

export function getIntegrationStatus(): IntegrationStatus {
  return {
    ...integrationStatus,
    campiMancanti: [...integrationStatus.campiMancanti],
    log: [...integrationStatus.log],
  };
}

function log(level: IntegrationLogEntry["level"], operation: string, message: string) {
  integrationStatus.log = [
    { ts: new Date(), level, operation, message },
    ...integrationStatus.log,
  ].slice(0, 30);
}

function markSuccess(count: number, operation = "getDipendenti") {
  integrationStatus.dipendentiCaricati = count;
  integrationStatus.ultimoAggiornamento = new Date();
  integrationStatus.ultimoErrore = null;
  integrationStatus.fallbackAttivo = false;
  log("info", operation, `OK — ${count} elementi`);
}

function markError(err: unknown, operation: string, fallback: boolean) {
  const msg = err instanceof Error ? err.message : String(err ?? "Errore sconosciuto");
  integrationStatus.ultimoErrore = msg;
  integrationStatus.ultimoAggiornamento = new Date();
  integrationStatus.fallbackAttivo = fallback;
  log("error", operation, msg);
}

// ---------------------------------------------------------------------------
// Implementazione MOCK (attuale)
// ---------------------------------------------------------------------------

const mockDataService: DataService = {
  async getSedi() {
    return SEDI;
  },
  async getDipendenti() {
    const list = state.map((d) => ({ ...d }));
    markSuccess(list.length);
    return list;
  },
  async getDipendente(id) {
    const d = state.find((x) => x.id === id);
    return d ? { ...d } : undefined;
  },
  async timbra(dipendenteId, tipo) {
    const idx = state.findIndex((d) => d.id === dipendenteId);
    if (idx < 0) throw new Error("Dipendente non trovato");
    const nuova: Timbratura = { tipo, ora: new Date().toISOString() };
    const prev = state[idx];
    const stato =
      tipo === "entrata" || tipo === "fine-pausa"
        ? "presente"
        : tipo === "inizio-pausa"
          ? "pausa"
          : "uscito";
    const entrataOra =
      tipo === "entrata" ? nuova.ora : prev.entrataOra;
    state[idx] = { ...prev, stato, ultimaTimbratura: nuova, entrataOra };
    return { ...state[idx] };
  },
};

// ---------------------------------------------------------------------------
// Implementazione SHAREPOINT — via server functions e Lovable Connector Gateway.
// Se una chiamata fallisce (config incompleta, credenziali mancanti, errore di
// rete o SharePoint down) si ripiega automaticamente sul mock service in modo
// che l'app resti sempre funzionante. L'errore viene tracciato in
// integrationStatus e mostrato in Amministrazione.
// ---------------------------------------------------------------------------

function mergeDipendentiTimbrature(
  dips: SpDipendente[],
  tims: SpTimbratura[],
): Dipendente[] {
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
    return {
      id: d.id,
      nome: d.nome,
      cognome: d.cognome,
      ruolo: d.ruolo || "—",
      sede: d.sede,
      orarioAtteso: "09:00",
      stato,
      entrataOra: entrata?.dataOra,
      ultimaTimbratura,
    };
  });
}

let cachedSnapshot: Dipendente[] = [];

const sharepointDataService: DataService = {
  async getSedi() {
    return mockDataService.getSedi();
  },
  async getDipendenti() {
    try {
      void sharepointConfig;
      const snap = await spGetSnapshot();
      const list = mergeDipendentiTimbrature(snap.dipendenti, snap.timbrature);
      cachedSnapshot = list;
      markSuccess(list.length, "getDipendenti");
      return list;
    } catch (err) {
      markError(err, "getDipendenti", true);
      return mockDataService.getDipendenti();
    }
  },
  async getDipendente(id) {
    const hit = cachedSnapshot.find((d) => d.id === id);
    if (hit) return { ...hit };
    try {
      const list = await sharepointDataService.getDipendenti();
      return list.find((d) => d.id === id);
    } catch {
      return mockDataService.getDipendente(id);
    }
  },
  async timbra(id, tipo) {
    try {
      await spCreateTimbratura({ data: { dipendenteId: id, evento: tipo, origine: "web" } });
      log("info", "timbra", `Registrata "${tipo}" per dipendente ${id}`);
      const list = await sharepointDataService.getDipendenti();
      const updated = list.find((d) => d.id === id);
      if (updated) return updated;
      throw new Error("Timbratura salvata ma dipendente non trovato dopo il refresh.");
    } catch (err) {
      markError(err, "timbra", true);
      return mockDataService.timbra(id, tipo);
    }
  },
};

// Selettore: usa SharePoint solo se completamente configurato.
export const dataService: DataService = isSharePointConfigured()
  ? sharepointDataService
  : mockDataService;

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
  const straordinari = list.filter(
    (d) => (d.straordinariMinuti ?? 0) > 0 && d.stato !== "uscito",
  ).length;
  const ritardi = list.filter((d) => (d.ritardoMinuti ?? 0) > 0).length;
  return { attivi, presenti, pausa, usciti, assenti, straordinari, ritardi };
}

// Stato "visivo" richiesto dalla dashboard live:
//   verde = presente · giallo = in pausa · rosso = assente · blu = oltre orario
export type DisplayStato = "presente" | "pausa" | "assente" | "oltre";

export function displayStato(d: Dipendente): DisplayStato {
  if (d.stato === "pausa") return "pausa";
  if (d.stato === "non-timbrato" || d.stato === "uscito") return "assente";
  if ((d.straordinariMinuti ?? 0) > 0) return "oltre";
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