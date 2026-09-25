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
import { computeOreOggi, tagliaEventiVisibili } from "./presenze-logic";
import {
  spCreateTimbratura,
  spGetDiagnostics,
  spGetMioStato,
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
    // Lo snapshot copre ~36h: si mostrano gli eventi di OGGI più quelli del
    // turno ancora in corso aperto ieri sera (notturno). I giorni precedenti
    // già chiusi — o un turno dimenticato oltre il tetto — restano fuori,
    // così lo stato del nuovo giorno riparte da "non timbrato".
    const stream = (byEmp.get(d.id) ?? [])
      .sort((a, b) => a.dataOra.localeCompare(b.dataOra))
      .map((e) => ({ evento: e.evento, ora: e.dataOra }));
    const visibili = tagliaEventiVisibili(stream);
    const eventiOggi: Timbratura[] = visibili.map((e) => ({ tipo: e.evento, ora: e.ora }));
    const entrata = eventiOggi.find((e) => e.tipo === "entrata");
    const last = eventiOggi[eventiOggi.length - 1];
    let stato: Dipendente["stato"] = "non-timbrato";
    if (last) {
      stato =
        last.tipo === "entrata" || last.tipo === "fine-pausa"
          ? "presente"
          : last.tipo === "inizio-pausa"
            ? "pausa"
            : "uscito";
    }
    const ultimaTimbratura: Timbratura | undefined = last;
    const ore = computeOreOggi(eventiOggi);
    return {
      id: d.id,
      nome: d.nome,
      cognome: d.cognome,
      ruolo: d.ruolo || "Dipendente",
      sede: d.sede === "tutte" ? "" : d.sede,
      orarioAtteso: "09:00",
      stato,
      entrataOra: entrata?.ora,
      ultimaTimbratura,
      eventiOggi,
      oreLavorateMinuti: ore.oreLavorateMinuti,
      pausaMinuti: ore.pausaMinuti,
      oltreOrarioMinuti: ore.oltreOrarioMinuti,
      straordinariMinuti: ore.oltreOrarioMinuti,
    };
  });
}

/** Applica una timbratura appena premuta al record LOCALE del dipendente,
 *  con le stesse regole del merge (stato, ultima, ore): la timbratrice mostra
 *  subito l'esito, senza aspettare il giro completo dello snapshot. */
export function applicaEventoLocale(
  d: Dipendente,
  tipo: Timbratura["tipo"],
  ora: string,
): Dipendente {
  return ricalcolaDaEventi(d, [...(d.eventiOggi ?? []), { tipo, ora }]);
}

/** Ricostruisce il record dal flusso di eventi di oggi (stesse regole del
 *  merge dello snapshot): stato, entrata, ultima timbratura, ore. */
export function ricalcolaDaEventi(d: Dipendente, eventiOggi: Timbratura[]): Dipendente {
  const ordinati = [...eventiOggi].sort((a, b) => a.ora.localeCompare(b.ora));
  const entrata = ordinati.find((e) => e.tipo === "entrata");
  const last = ordinati[ordinati.length - 1];
  const stato: Dipendente["stato"] = !last
    ? "non-timbrato"
    : last.tipo === "entrata" || last.tipo === "fine-pausa"
      ? "presente"
      : last.tipo === "inizio-pausa"
        ? "pausa"
        : "uscito";
  const ore = computeOreOggi(ordinati);
  return {
    ...d,
    stato,
    entrataOra: entrata?.ora,
    ultimaTimbratura: last,
    eventiOggi: ordinati,
    oreLavorateMinuti: ore.oreLavorateMinuti,
    pausaMinuti: ore.pausaMinuti,
    oltreOrarioMinuti: ore.oltreOrarioMinuti,
    straordinariMinuti: ore.oltreOrarioMinuti,
  };
}

/** Una rilettura dal server NON deve cancellare una timbratura appena
 *  registrata: Graph, sulle query filtrate della lista, può non restituire
 *  ancora la riga creata pochi secondi prima (ADM001 25/09: "sono entrato
 *  ma Inizio pausa dice che non sono entrato"). Gli eventi locali più
 *  recenti dell'ultimo del server, registrati da meno di 10 minuti e non
 *  già presenti, restano nel record. */
export function unisciEventiRecenti(
  locale: Dipendente | null | undefined,
  server: Dipendente,
  adesso = Date.now(),
): Dipendente {
  const loc = locale?.eventiOggi ?? [];
  if (!loc.length) return server;
  const srv = server.eventiOggi ?? [];
  const ultimoServer = srv[srv.length - 1]?.ora ?? "";
  const vicino = (a: string, b: string) =>
    Math.abs(new Date(a).getTime() - new Date(b).getTime()) < 120_000;
  const extra = loc.filter(
    (e) =>
      e.ora > ultimoServer &&
      adesso - new Date(e.ora).getTime() < 10 * 60_000 &&
      !srv.some((s) => s.tipo === e.tipo && vicino(s.ora, e.ora)),
  );
  return extra.length ? ricalcolaDaEventi(server, [...srv, ...extra]) : server;
}

// --- Timbratrice: il proprio stato, leggero e con memoria locale --------------
// Primo caricamento (Posta Doc, 25/09): la pagina aspettava lo snapshot di
// TUTTI i dipendenti e TUTTE le timbrature. Ora: (1) l'ultimo stato salvato
// sul dispositivo compare subito, (2) il server manda solo il record del
// chiamante e le sue timbrature, (3) la pagina si allinea appena arrivano.
const MIO_CACHE_KEY = "dr.presenze.mio";
const MIO_CACHE_ORE = 14;

export function leggiMioCache(id: string): Dipendente | null {
  try {
    const raw = localStorage.getItem(MIO_CACHE_KEY);
    if (!raw) return null;
    const j = JSON.parse(raw) as { id?: string; at?: number; d?: Dipendente };
    if (!j || j.id !== id || !j.d || typeof j.at !== "number") return null;
    if (Date.now() - j.at > MIO_CACHE_ORE * 3600_000) return null;
    return j.d;
  } catch {
    return null;
  }
}

export function salvaMioCache(d: Dipendente): void {
  try {
    localStorage.setItem(MIO_CACHE_KEY, JSON.stringify({ id: d.id, at: Date.now(), d }));
  } catch {
    /* storage pieno o bloccato: la pagina funziona lo stesso */
  }
}

export async function getMioStato(id: string): Promise<Dipendente | undefined> {
  const snap = (await spGetMioStato()) as {
    dipendente: SpDipendente | null;
    timbrature: SpTimbratura[];
  };
  if (!snap.dipendente || snap.dipendente.id !== id) return undefined;
  const [d] = mergeDipendentiTimbrature([snap.dipendente], snap.timbrature);
  if (d) {
    salvaMioCache(d);
    setSpStatus("online");
  }
  return d;
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
