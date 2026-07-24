// DR Portal — SharePoint gateway helpers (server-only)
// -----------------------------------------------------------------------------
// Wrapper attorno al Lovable Connector Gateway per Microsoft SharePoint.
// Server-only (import bloccato da suffisso .server.ts).
//
// Regole di produzione:
// - Discovery preferisce il sito canonico DR (drlogisticaroma / DRPORTAL).
// - Fallback: elenca i siti e sceglie l'unico che contiene entrambe le liste;
//   se ne trova più di uno, errore chiaro (mai scelta arbitraria).
// - Cache discovery con TTL (default 1h). "Force" bypassa e invalida.
// - Sulle chiamate successive si usano SEMPRE gli ID (siteId, listId) e i
//   nomi INTERNI delle colonne, mai i display name. In caso di 404 la cache
//   viene invalidata e la chiamata riprovata una sola volta.
// - Logging strutturato in memoria (ultimi 100 eventi) accessibile via server fn.
// - Nessun token/credenziale finisce nei log.
// -----------------------------------------------------------------------------

import {
  validateRichiesta,
  validateDecisione,
  computeDurataGiorni,
  computeDurataOre,
  computeAnnoCompetenza,
  isAutoApprovazione,
  supervisionaSede,
  isSupervisoreGlobale,
  isSedeStorica,
  richiedeApprovazione,
  misuraInGiorni,
  isRimborso,
  formatTitle,
  canDecide,
  canCancel,
  parseStato,
  NOTA_AUTO_APPROVAZIONE,
  type TipoRichiesta,
  type ModalitaStraordinario,
  type TipoAcquisto,
  type DecisioneRichiesta,
} from "./richieste-logic";
import { anomalieDelGiorno, UNDO_TIMBRATURA_MINUTI, type TipoAnomalia } from "./presenze-logic";
import {
  chiaveMovimento,
  classificaMovimento,
  applicaRegole,
  matchRegola,
  LEGACY_IMPORT_ID,
  type MovimentoRaw,
  type RegolaFinanza,
} from "./finanza-logic";
import type { FatturaRaw, TerminePagamento, AbbinamentoIncasso } from "./fatture-logic";
export { LEGACY_IMPORT_ID };
export type { RegolaFinanza, FatturaRaw, TerminePagamento, AbbinamentoIncasso };
import { normalizeRuolo } from "./session";
import {
  generateVapidKeys,
  sendWebPush,
  type PushSubscriptionData,
  type VapidKeys,
} from "./webpush.server";
import {
  oreLavorateGiorno,
  isoDow,
  lunediDellaSettimana,
  orePrevisteSettimana,
  straordinarioSettimana,
  ymd,
  round2,
} from "./rendiconto-logic";

const GATEWAY_BASE = "https://connector-gateway.lovable.dev/microsoft_sharepoint";
const CACHE_TTL_MS = Number(process.env.SP_CACHE_TTL_MS ?? 60 * 60 * 1000);
const TARGET_HOST = "drlogisticaroma.sharepoint.com";
const TARGET_SITE_PATH = "DRPORTAL";

// Display name attesi (usati per la risoluzione internalName in getListColumns).
// Le operazioni Graph usano poi gli internalName risolti, non queste stringhe.
export const SP_DISPLAY = {
  dipendenti: {
    Nome: "Nome",
    Cognome: "Cognome",
    NomeCompleto: "NomeCompleto",
    Email: "Email",
    Sede: "Sede",
    Attivo: "Attivo",
    Ruolo: "Responsabile",
    Codice: "Codice",
    PIN: "PIN",
    Visibile: "Visibile",
    Autorizza: "Autorizza",
    Operatore: "Operatore",
    OreSettimanali: "OreSettimanali",
    Inquadramento: "Inquadramento",
    GiorniFerieAnnui: "GiorniFerieAnnui",
    OrePermessiAnnui: "OrePermessiAnnui",
    CF: "CF",
  },
  timbrature: {
    Dipendente: "Dipendente",
    Evento: "Evento",
    DataOra: "DataOra",
    Origine: "Dispositivo",
    Esito: "Esito",
    Note: "Note",
    // NB: la posizione/geolocalizzazione NON è tra le colonne attese: non viene
    // raccolta (implicazioni GDPR / Art. 4 Statuto dei Lavoratori). Il codice
    // mantiene comunque il "gancio" opzionale (F.Posizione) se un domani la si
    // introdurrà con base giuridica e informativa: basterà riaggiungerla qui.
  },
  // Modulo Richieste (Sprint 2). Lista OPZIONALE: la sua assenza non deve
  // rompere la discovery di Dipendenti/Timbrature (vedi discoverSharePoint).
  richieste: {
    Richiedente: "Richiedente",
    CodiceRichiedente: "CodiceRichiedente",
    SedeRichiedente: "SedeRichiedente",
    TipoRichiesta: "TipoRichiesta",
    // Nome colonna SharePoint con accento (grafia italiana corretta). La chiave
    // logica resta "Modalita" (senza accento) per comodità nel codice.
    Modalita: "Modalità",
    DataInizio: "DataInizio",
    DataFine: "DataFine",
    OraInizio: "OraInizio",
    OraFine: "OraFine",
    Motivazione: "Motivazione",
    DurataGiorni: "DurataGiorni",
    DurataOre: "DurataOre",
    Stato: "Stato",
    DataInvio: "DataInvio",
    Approvatore: "Approvatore",
    DataDecisione: "DataDecisione",
    NoteDecisione: "NoteDecisione",
    ProtocolloINPS: "ProtocolloINPS",
    Importo: "Importo",
    TipologiaAcquisto: "TipologiaAcquisto",
    Giustificativo: "Giustificativo",
    AnnoCompetenza: "AnnoCompetenza",
  },
  // Modulo Documenti (Sprint 4). Lista OPZIONALE (discovery soft).
  documenti: {
    Categoria: "Categoria",
    Titolo: "Titolo",
    Ambito: "Ambito",
    DestinatarioId: "DestinatarioId",
    CodiceDestinatario: "CodiceDestinatario",
    SedeDestinatario: "SedeDestinatario",
    File: "File",
    NomeFile: "NomeFile",
    DataDocumento: "DataDocumento",
    CaricatoDa: "CaricatoDa",
  },
  // Modulo Comunicazioni interne (Sprint 4). Lista OPZIONALE.
  comunicazioni: {
    Titolo: "Titolo",
    Testo: "Testo",
    // Nomi colonna adattati per non collidere con campi già esistenti sulla
    // lista: "Tipo"→"Tipologia", "Autore"→"AutoreComunicazione".
    Tipo: "Tipologia",
    Sede: "Sede",
    DataComunicazione: "DataComunicazione",
    Autore: "AutoreComunicazione",
    Allegato: "Allegato",
    RichiedePresaVisione: "RichiedePresaVisione",
  },
  // Prese visione delle comunicazioni (ricevute di lettura). Lista OPZIONALE.
  preseVisione: {
    ComunicazioneId: "ComunicazioneId",
    DipendenteId: "DipendenteId",
    CodiceDipendente: "CodiceDipendente",
    DataLettura: "DataLettura",
  },
  // Sottoscrizioni Web Push (notifiche telefono). Lista OPZIONALE. Contiene
  // anche la riga speciale "__vapid__" con le chiavi applicative.
  pushSubscriptions: {
    Endpoint: "Endpoint",
    P256dh: "P256dh",
    Auth: "Auth",
    DipendenteId: "DipendenteId",
    Sede: "Sede",
  },
  // Voci di spesa (macro → dettaglio) per rimborsi e acquisti. Lista OPZIONALE
  // gestita direttamente dall'azienda: aggiungere una voce = aggiungere una riga.
  voci: {
    Ambito: "Ambito",
    Macro: "Macro",
    Dettaglio: "Dettaglio",
  },
  // Coda email in uscita (outbox): il portale accoda, un flusso Power Automate
  // invia e aggiorna lo Stato. Lista OPZIONALE.
  codaEmail: {
    Destinatari: "Destinatari",
    Oggetto: "Oggetto",
    Corpo: "Corpo",
    Allegato: "Allegato",
    Stato: "Stato",
    Mittente: "Mittente",
  },
  // Movimenti bancari (sezione Finanza, solo direttore DR005). Lista OPZIONALE.
  // Title = chiave di deduplicazione (calcolata dai campi grezzi, vedi
  // finanza-logic.ts). I campi grezzi (DataContabile..Descrizione) sono
  // IMMUTABILI dopo l'import; la sanatura tocca solo Tipologia/Cliente/
  // NrFattura/Note/DaVerificare.
  movimenti: {
    DataContabile: "DataContabile",
    DataValuta: "DataValuta",
    Importo: "Importo",
    Divisa: "Divisa",
    Causale: "Causale",
    Descrizione: "Descrizione",
    Tipologia: "Tipologia",
    Cliente: "Cliente",
    NrFattura: "NrFattura",
    Note: "Note",
    DaVerificare: "DaVerificare",
    // Lotto di import (per lo storico e l'annullamento di un import intero).
    ImportId: "ImportId",
  },
  // Regole apprese della sezione Finanza (correzioni permanenti del direttore:
  // "questo cliente/pattern → questa tipologia e/o questo nome"). OPZIONALE.
  regoleFinanza: {
    Pattern: "Pattern",
    CampoMatch: "CampoMatch",
    ModoMatch: "ModoMatch",
    Tipologia: "Tipologia",
    ClienteNuovo: "ClienteNuovo",
  },
  // Fatture emesse (sezione Finanza → Fatture). Title = NOME FILE SdI (chiave
  // univoca, mai il numero). v1 alimentata dall'export Aruba; predisposta per
  // il sync via API Aruba v2 (stessa chiave). OPZIONALE.
  fatture: {
    Numero: "Numero",
    IdSdi: "IdSdi",
    DataInvio: "DataInvio",
    DataDocumento: "DataDocumento",
    TipoDocumento: "TipoDocumento",
    Cliente: "Cliente",
    PIVA: "PIVA",
    MetodoPagamento: "MetodoPagamento",
    Imponibile: "Imponibile",
    Iva: "Iva",
    TotaleDocumento: "TotaleDocumento",
    NettoAPagare: "NettoAPagare",
    StatoSdI: "StatoSdI",
  },
  // Termini di pagamento per cliente (giorni). Gestita dal direttore. OPZIONALE.
  terminiPagamento: {
    Cliente: "Cliente",
    Giorni: "Giorni",
    Descrizione: "Descrizione",
  },
  // Abbinamenti fattura ↔ movimento bancario (n:n con importo allocato).
  // Chiavi NATURALI (nome file + chiave movimento): sopravvivono a
  // annulla/reimport di entrambe le sorgenti. OPZIONALE.
  abbinamenti: {
    FatturaFile: "FatturaFile",
    MovimentoChiave: "MovimentoChiave",
    Importo: "Importo",
    Origine: "Origine",
  },
  // Collegamento Aruba Fatturazione Elettronica (una sola riga di config).
  // La password è CIFRATA AES-GCM con chiave derivata dal segreto server:
  // chi legge la lista vede solo testo cifrato. OPZIONALE.
  arubaConfig: {
    Username: "Username",
    PasswordCifrata: "PasswordCifrata",
    UltimaSync: "UltimaSync",
  },
  // Richieste di acquisto (modulo Procurement). Lista OPZIONALE.
  acquisti: {
    Richiedente: "Richiedente",
    CodiceRichiedente: "CodiceRichiedente",
    SedeRichiedente: "SedeRichiedente",
    Macro: "Macro",
    Dettaglio: "Dettaglio",
    Descrizione: "Descrizione",
    Importo: "Importo",
    Stato: "Stato",
    DataRichiesta: "DataRichiesta",
    Approvatore: "Approvatore",
    DataDecisione: "DataDecisione",
    NoteDecisione: "NoteDecisione",
  },
} as const;

const REQUIRED_DIP_KEYS = [
  "Nome",
  "Cognome",
  "NomeCompleto",
  "Email",
  "Sede",
  "Attivo",
  "Ruolo",
] as const;
const REQUIRED_TIM_KEYS = ["Dipendente", "Evento", "DataOra", "Origine", "Esito", "Note"] as const;

// Nomi delle liste SharePoint da individuare (case-insensitive, tolleranti a
// varianti singolare/plurale).
const LIST_NAMES = {
  dipendenti: ["Dipendenti", "Dipendente"],
  timbrature: ["Timbrature", "Timbratura"],
  richieste: ["Richieste", "Richiesta"],
  // NB: NON usare "Documenti": in SharePoint italiano è il nome della libreria
  // documenti di default → collisione. La lista metadati è "DocumentiDipendenti".
  documenti: ["DocumentiDipendenti", "DocumentoDipendente"],
  comunicazioni: ["Comunicazioni", "Comunicazione"],
  preseVisione: ["PreseVisione", "PresaVisione", "PreseVisioni"],
  pushSubscriptions: ["PushSubscriptions", "PushSubscription"],
  voci: ["Voci", "Voce", "VociSpesa"],
  acquisti: ["RichiesteAcquisto", "RichiestaAcquisto", "Acquisti"],
  codaEmail: ["CodaEmail", "Coda Email", "EmailQueue"],
  movimenti: ["MovimentiBancari", "MovimentoBancario", "Movimenti"],
  regoleFinanza: ["RegoleFinanza", "RegolaFinanza", "RegoleBanca"],
  fatture: ["FattureEmesse", "FatturaEmessa", "Fatture"],
  terminiPagamento: ["TerminiPagamento", "TerminiDiPagamento", "TerminePagamento"],
  abbinamenti: ["AbbinamentiIncassi", "AbbinamentoIncasso", "Abbinamenti"],
  arubaConfig: ["ArubaConfig", "ConfigurazioneAruba"],
} as const;

// ---------------------------------------------------------------------------
// Logging diagnostico strutturato (ultimi 100 eventi, in-memory)
// ---------------------------------------------------------------------------
export interface SpLogEvent {
  ts: string;
  level: "info" | "warn" | "error";
  operation: string;
  message: string;
  durataMs?: number;
  details?: string;
}
const LOG_MAX = 100;
const spLog: SpLogEvent[] = [];

function sanitize(msg: string): string {
  return msg
    .replace(/Bearer\s+[A-Za-z0-9._-]+/gi, "Bearer ***")
    .replace(/(api[-_ ]?key["']?\s*[:=]\s*["']?)[^"'\s,}]+/gi, "$1***")
    .replace(/(access_token["']?\s*[:=]\s*["']?)[^"'\s,}]+/gi, "$1***")
    .slice(0, 500);
}

export function logSp(
  level: SpLogEvent["level"],
  operation: string,
  message: string,
  extra?: { durataMs?: number; details?: Record<string, unknown> },
) {
  spLog.unshift({
    ts: new Date().toISOString(),
    level,
    operation,
    message: sanitize(message),
    durataMs: extra?.durataMs,
    details: extra?.details ? sanitize(JSON.stringify(extra.details)) : undefined,
  });
  if (spLog.length > LOG_MAX) spLog.length = LOG_MAX;
}

export function getSpLog(): SpLogEvent[] {
  return [...spLog];
}

// ---------------------------------------------------------------------------
// Tipi discovery
// ---------------------------------------------------------------------------
export interface SpDiscovered {
  siteId: string;
  siteName: string;
  siteWebUrl: string;
  listDipendenti: string;
  listDipendentiName: string;
  listTimbrature: string;
  listTimbratureName: string;
  // Lista Richieste — OPZIONALE (modulo Sprint 2). null se non presente.
  listRichieste: string | null;
  listRichiesteName: string | null;
  // Mappa "chiave logica" -> internalName reale su SharePoint.
  dipendentiFields: Record<string, string>;
  timbratureFields: Record<string, string>;
  richiesteFields: Record<string, string>;
  dipendentiMissing: string[];
  timbratureMissing: string[];
  richiesteMissing: string[];
  // Liste Sprint 4 (Documenti / Comunicazioni / PreseVisione) — OPZIONALI.
  listDocumenti: string | null;
  listDocumentiName: string | null;
  documentiFields: Record<string, string>;
  documentiMissing: string[];
  listComunicazioni: string | null;
  listComunicazioniName: string | null;
  comunicazioniFields: Record<string, string>;
  comunicazioniMissing: string[];
  listPreseVisione: string | null;
  listPreseVisioneName: string | null;
  preseVisioneFields: Record<string, string>;
  preseVisioneMissing: string[];
  listPushSubscriptions: string | null;
  listPushSubscriptionsName: string | null;
  pushSubscriptionsFields: Record<string, string>;
  pushSubscriptionsMissing: string[];
  listVoci: string | null;
  listVociName: string | null;
  vociFields: Record<string, string>;
  vociMissing: string[];
  listAcquisti: string | null;
  listAcquistiName: string | null;
  acquistiFields: Record<string, string>;
  acquistiMissing: string[];
  listCodaEmail: string | null;
  listCodaEmailName: string | null;
  codaEmailFields: Record<string, string>;
  codaEmailMissing: string[];
  listMovimenti: string | null;
  listMovimentiName: string | null;
  movimentiFields: Record<string, string>;
  movimentiMissing: string[];
  listRegoleFinanza: string | null;
  listRegoleFinanzaName: string | null;
  regoleFinanzaFields: Record<string, string>;
  regoleFinanzaMissing: string[];
  listFatture: string | null;
  listFattureName: string | null;
  fattureFields: Record<string, string>;
  fattureMissing: string[];
  listTermini: string | null;
  listTerminiName: string | null;
  terminiFields: Record<string, string>;
  terminiMissing: string[];
  listAbbinamenti: string | null;
  listAbbinamentiName: string | null;
  abbinamentiFields: Record<string, string>;
  abbinamentiMissing: string[];
  listArubaConfig: string | null;
  listArubaConfigName: string | null;
  arubaConfigFields: Record<string, string>;
  arubaConfigMissing: string[];
  cachedAt: string;
  expiresAt: string;
}

let discoveredCache: SpDiscovered | null = null;
let lastGraphResponseMs = 0;

export function clearSpDiscoveryCache() {
  discoveredCache = null;
  logSp("info", "cache", "Cache discovery invalidata");
}

export function getSpDiscoveryCached(): SpDiscovered | null {
  if (discoveredCache && new Date(discoveredCache.expiresAt).getTime() < Date.now()) {
    discoveredCache = null;
  }
  return discoveredCache;
}

async function gatewayFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const lovableKey = process.env.LOVABLE_API_KEY;
  const spKey = process.env.MICROSOFT_SHAREPOINT_API_KEY;
  if (!lovableKey || !spKey) {
    throw new Error(
      "Credenziali SharePoint non disponibili sul server (LOVABLE_API_KEY / MICROSOFT_SHAREPOINT_API_KEY).",
    );
  }
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${lovableKey}`);
  headers.set("X-Connection-Api-Key", spKey);
  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  const started = Date.now();
  const res = await fetch(`${GATEWAY_BASE}${path}`, { ...init, headers });
  lastGraphResponseMs = Date.now() - started;
  return res;
}

export class SpHttpError extends Error {
  constructor(
    public status: number,
    message: string,
    public path: string,
  ) {
    super(message);
  }
}

async function gatewayJson<T = unknown>(path: string, init: RequestInit = {}): Promise<T> {
  // Retry transitorio su 5xx/429: il gateway a volte risponde 503
  // "upstream connect error" per pochi secondi. Ritentiamo con backoff
  // esponenziale prima di propagare l'errore alla UI.
  // SOLO GET (idempotenti): le scritture (POST/PATCH/DELETE) NON si ritentano,
  // per non rischiare timbrature/richieste duplicate su un 503 tardivo.
  const idempotent = (init.method ?? "GET").toUpperCase() === "GET";
  const maxAttempts = idempotent ? 3 : 1;
  let lastStatus = 0;
  let lastBody = "";
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let res: Response;
    try {
      res = await gatewayFetch(path, init);
    } catch (err) {
      if (attempt === maxAttempts) throw err;
      await new Promise((r) => setTimeout(r, 300 * attempt));
      continue;
    }
    if (res.ok) return (await res.json()) as T;
    lastStatus = res.status;
    lastBody = await res.text().catch(() => "");
    const retriable =
      res.status === 429 || res.status === 502 || res.status === 503 || res.status === 504;
    if (!retriable || attempt === maxAttempts) break;
    logSp(
      "warn",
      "gateway",
      `Retry ${attempt}/${maxAttempts - 1} dopo ${res.status} su ${path.split("?")[0]}`,
    );
    await new Promise((r) => setTimeout(r, 400 * attempt));
  }
  throw new SpHttpError(
    lastStatus,
    `SharePoint ${init.method ?? "GET"} ${path.split("?")[0]} → ${lastStatus} ${sanitize(lastBody)}`,
    path,
  );
}

// ---------------------------------------------------------------------------
// Auto-discovery del sito e delle liste
// ---------------------------------------------------------------------------

interface GraphSite {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
}
interface GraphList {
  id: string;
  name?: string;
  displayName?: string;
  list?: { hidden?: boolean; template?: string };
}
interface GraphColumn {
  id?: string;
  name?: string; // internal name
  displayName?: string;
  hidden?: boolean;
  readOnly?: boolean;
}

function matchListName(list: GraphList, targets: readonly string[]): boolean {
  const candidates = [list.displayName, list.name].filter(Boolean).map((s) => s!.toLowerCase());
  return targets.some((t) => candidates.includes(t.toLowerCase()));
}

async function getListColumns(siteId: string, listId: string): Promise<GraphColumn[]> {
  const res = await gatewayJson<{ value: GraphColumn[] }>(
    `/sites/${siteId}/lists/${listId}/columns?$select=id,name,displayName,hidden,readOnly`,
  );
  return res.value ?? [];
}

function resolveInternalNames(
  columns: GraphColumn[],
  desired: Record<string, string>,
): { map: Record<string, string>; missing: string[] } {
  // Escludi colonne di sistema (hidden o readOnly) come LinkFilename2/LinkTitle
  // che condividono display name "Nome"/"Titolo" con colonne custom reali.
  const usable = columns.filter((c) => !c.hidden && !c.readOnly);
  const byDisplay = new Map<string, GraphColumn>();
  const byName = new Map<string, GraphColumn>();
  for (const c of usable) {
    if (c.displayName) byDisplay.set(c.displayName.toLowerCase(), c);
    if (c.name) byName.set(c.name.toLowerCase(), c);
  }
  const map: Record<string, string> = {};
  const missing: string[] = [];
  for (const [logical, display] of Object.entries(desired)) {
    const hit =
      byDisplay.get(display.toLowerCase()) ??
      byName.get(display.toLowerCase()) ??
      byDisplay.get(logical.toLowerCase()) ??
      byName.get(logical.toLowerCase());
    if (hit?.name) map[logical] = hit.name;
    else missing.push(display);
  }
  return { map, missing };
}

async function tryResolveTargetSite(): Promise<GraphSite | null> {
  // 1) Tentativo diretto per path canonico DRPORTAL.
  try {
    const site = await gatewayJson<GraphSite>(`/sites/${TARGET_HOST}:/sites/${TARGET_SITE_PATH}`);
    if (site?.id) {
      logSp("info", "discover.site", `Sito canonico trovato: ${site.displayName ?? site.name}`);
      return site;
    }
  } catch (err) {
    logSp(
      "warn",
      "discover.site",
      `Path canonico non risolto (${TARGET_HOST}/sites/${TARGET_SITE_PATH}): ${
        err instanceof Error ? err.message : String(err)
      }`,
    );
  }
  return null;
}

export async function discoverSharePoint(force = false): Promise<SpDiscovered> {
  const cached = getSpDiscoveryCached();
  if (!force && cached) return cached;
  if (force) discoveredCache = null;

  const started = Date.now();

  // 1) Sito canonico DR.
  let targetSite = await tryResolveTargetSite();

  // 2) Fallback: scansione siti + filtro per liste attese.
  if (!targetSite) {
    const sitesRes = await gatewayJson<{ value: GraphSite[] }>(`/sites?search=*`);
    const sites = sitesRes.value ?? [];
    if (sites.length === 0) {
      throw new Error(
        "Microsoft Graph non ha restituito nessun sito SharePoint accessibile. Verifica il permesso Sites.Read.All sul connettore.",
      );
    }
    const candidates: GraphSite[] = [];
    for (const site of sites) {
      try {
        const listsRes = await gatewayJson<{ value: GraphList[] }>(
          `/sites/${site.id}/lists?$select=id,name,displayName,list`,
        );
        const lists = (listsRes.value ?? []).filter((l) => !l.list?.hidden);
        const dip = lists.find((l) => matchListName(l, LIST_NAMES.dipendenti));
        const tim = lists.find((l) => matchListName(l, LIST_NAMES.timbrature));
        if (dip && tim) candidates.push(site);
      } catch {
        /* sito non ispezionabile — ignorato */
      }
    }
    if (candidates.length === 0) {
      throw new Error(
        `Nessun sito SharePoint contiene entrambe le liste "Dipendenti" e "Timbrature". Atteso ${TARGET_HOST}/sites/${TARGET_SITE_PATH}.`,
      );
    }
    if (candidates.length > 1) {
      const names = candidates.map((s) => s.displayName || s.name).join(", ");
      throw new Error(
        `Discovery ambigua: trovati ${candidates.length} siti candidati (${names}). Specificare il sito canonico ${TARGET_HOST}/sites/${TARGET_SITE_PATH}.`,
      );
    }
    targetSite = candidates[0];
  }

  // 3) Discovery liste sul sito scelto.
  const listsRes = await gatewayJson<{ value: GraphList[] }>(
    `/sites/${targetSite.id}/lists?$select=id,name,displayName,list`,
  );
  const lists = (listsRes.value ?? []).filter((l) => !l.list?.hidden);
  const dip = lists.find((l) => matchListName(l, LIST_NAMES.dipendenti));
  const tim = lists.find((l) => matchListName(l, LIST_NAMES.timbrature));
  if (!dip || !tim) {
    throw new Error(
      `Sito "${targetSite.displayName ?? targetSite.name}" trovato ma manca ${!dip ? '"Dipendenti"' : ""}${!dip && !tim ? " e " : ""}${!tim ? '"Timbrature"' : ""}.`,
    );
  }

  // 4) Risoluzione internal name colonne.
  const [dipCols, timCols] = await Promise.all([
    getListColumns(targetSite.id, dip.id),
    getListColumns(targetSite.id, tim.id),
  ]);
  const dipRes = resolveInternalNames(dipCols, SP_DISPLAY.dipendenti);
  const timRes = resolveInternalNames(timCols, SP_DISPLAY.timbrature);
  // Colonne Dipendenti facoltative: la loro assenza NON segna la salute rossa
  // (il codice ha default: Inquadramento="" e GiorniFerieAnnui=26).
  const OPTIONAL_DIP = new Set(["Inquadramento", "GiorniFerieAnnui", "OrePermessiAnnui", "CF"]);
  dipRes.missing = dipRes.missing.filter((m) => !OPTIONAL_DIP.has(m));

  // 5) Discovery SOFT della lista Richieste (Sprint 2): se assente o non
  // ispezionabile, si prosegue senza — le presenze non devono dipenderne.
  const rich = lists.find((l) => matchListName(l, LIST_NAMES.richieste));
  let listRichieste: string | null = null;
  let listRichiesteName: string | null = null;
  let richiesteFields: Record<string, string> = {};
  let richiesteMissing: string[] = [];
  if (rich) {
    listRichieste = rich.id;
    listRichiesteName = rich.displayName || rich.name || rich.id;
    try {
      const richCols = await getListColumns(targetSite.id, rich.id);
      const richRes = resolveInternalNames(richCols, SP_DISPLAY.richieste);
      richiesteFields = richRes.map;
      richiesteMissing = richRes.missing;
    } catch (err) {
      logSp(
        "warn",
        "discover.richieste",
        `Colonne Richieste non risolte: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  // 6) Discovery SOFT delle liste Sprint 4 (Documenti/Comunicazioni/PreseVisione).
  const softList = async (names: readonly string[], display: Record<string, string>) => {
    const l = lists.find((x) => matchListName(x, names));
    if (!l) return { id: null, name: null, fields: {} as Record<string, string>, missing: [] };
    try {
      const cols = await getListColumns(targetSite.id, l.id);
      const res = resolveInternalNames(cols, display);
      return {
        id: l.id,
        name: l.displayName || l.name || l.id,
        fields: res.map,
        missing: res.missing,
      };
    } catch (err) {
      logSp(
        "warn",
        "discover.softlist",
        `Colonne non risolte per ${l.displayName || l.name}: ${err instanceof Error ? err.message : String(err)}`,
      );
      return { id: l.id, name: l.displayName || l.name || l.id, fields: {}, missing: [] };
    }
  };
  const docs = await softList(LIST_NAMES.documenti, SP_DISPLAY.documenti);
  const coms = await softList(LIST_NAMES.comunicazioni, SP_DISPLAY.comunicazioni);
  const pv = await softList(LIST_NAMES.preseVisione, SP_DISPLAY.preseVisione);
  const push = await softList(LIST_NAMES.pushSubscriptions, SP_DISPLAY.pushSubscriptions);
  const voci = await softList(LIST_NAMES.voci, SP_DISPLAY.voci);
  const acq = await softList(LIST_NAMES.acquisti, SP_DISPLAY.acquisti);
  const coda = await softList(LIST_NAMES.codaEmail, SP_DISPLAY.codaEmail);
  const mov = await softList(LIST_NAMES.movimenti, SP_DISPLAY.movimenti);
  const reg = await softList(LIST_NAMES.regoleFinanza, SP_DISPLAY.regoleFinanza);
  const fat = await softList(LIST_NAMES.fatture, SP_DISPLAY.fatture);
  const trm = await softList(LIST_NAMES.terminiPagamento, SP_DISPLAY.terminiPagamento);
  const abb = await softList(LIST_NAMES.abbinamenti, SP_DISPLAY.abbinamenti);
  const aru = await softList(LIST_NAMES.arubaConfig, SP_DISPLAY.arubaConfig);

  const now = Date.now();
  discoveredCache = {
    siteId: targetSite.id,
    siteName: targetSite.displayName || targetSite.name || targetSite.id,
    siteWebUrl: targetSite.webUrl ?? "",
    listDipendenti: dip.id,
    listDipendentiName: dip.displayName || dip.name || dip.id,
    listTimbrature: tim.id,
    listTimbratureName: tim.displayName || tim.name || tim.id,
    listRichieste,
    listRichiesteName,
    dipendentiFields: dipRes.map,
    timbratureFields: timRes.map,
    richiesteFields,
    dipendentiMissing: dipRes.missing,
    timbratureMissing: timRes.missing,
    richiesteMissing,
    listDocumenti: docs.id,
    listDocumentiName: docs.name,
    documentiFields: docs.fields,
    documentiMissing: docs.missing,
    listComunicazioni: coms.id,
    listComunicazioniName: coms.name,
    comunicazioniFields: coms.fields,
    comunicazioniMissing: coms.missing,
    listPreseVisione: pv.id,
    listPreseVisioneName: pv.name,
    preseVisioneFields: pv.fields,
    preseVisioneMissing: pv.missing,
    listPushSubscriptions: push.id,
    listPushSubscriptionsName: push.name,
    pushSubscriptionsFields: push.fields,
    pushSubscriptionsMissing: push.missing,
    listVoci: voci.id,
    listVociName: voci.name,
    vociFields: voci.fields,
    vociMissing: voci.missing,
    listAcquisti: acq.id,
    listAcquistiName: acq.name,
    acquistiFields: acq.fields,
    acquistiMissing: acq.missing,
    listCodaEmail: coda.id,
    listCodaEmailName: coda.name,
    codaEmailFields: coda.fields,
    codaEmailMissing: coda.missing,
    listMovimenti: mov.id,
    listMovimentiName: mov.name,
    movimentiFields: mov.fields,
    movimentiMissing: mov.missing,
    listRegoleFinanza: reg.id,
    listRegoleFinanzaName: reg.name,
    regoleFinanzaFields: reg.fields,
    regoleFinanzaMissing: reg.missing,
    listFatture: fat.id,
    listFattureName: fat.name,
    fattureFields: fat.fields,
    fattureMissing: fat.missing,
    listTermini: trm.id,
    listTerminiName: trm.name,
    terminiFields: trm.fields,
    terminiMissing: trm.missing,
    listAbbinamenti: abb.id,
    listAbbinamentiName: abb.name,
    abbinamentiFields: abb.fields,
    abbinamentiMissing: abb.missing,
    listArubaConfig: aru.id,
    listArubaConfigName: aru.name,
    arubaConfigFields: aru.fields,
    arubaConfigMissing: aru.missing,
    cachedAt: new Date(now).toISOString(),
    expiresAt: new Date(now + CACHE_TTL_MS).toISOString(),
  };
  logSp(
    "info",
    "discover",
    `Discovery OK — sito ${discoveredCache.siteName}, liste ${discoveredCache.listDipendentiName}+${discoveredCache.listTimbratureName}`,
    { durataMs: Date.now() - started },
  );
  if (dipRes.missing.length || timRes.missing.length) {
    logSp(
      "warn",
      "discover.columns",
      `Colonne mancanti — Dipendenti: [${dipRes.missing.join(", ") || "-"}] · Timbrature: [${timRes.missing.join(", ") || "-"}]`,
    );
  }
  return discoveredCache;
}

// Retry helper: su 404 invalida cache e riprova UNA sola volta.
async function withDiscoveryRetry<T>(op: () => Promise<T>): Promise<T> {
  try {
    return await op();
  } catch (err) {
    if (err instanceof SpHttpError && err.status === 404) {
      logSp("warn", "retry", `404 su ${err.path.split("?")[0]}, invalido cache e ritento`);
      discoveredCache = null;
      return await op();
    }
    throw err;
  }
}

// ---------------------------------------------------------------------------
// Tipi Graph minimali
// ---------------------------------------------------------------------------
interface GraphListItem<F = Record<string, unknown>> {
  id: string;
  fields?: F;
}
interface GraphListResponse<F> {
  value: GraphListItem<F>[];
}

type SedeRaw = string | undefined | null;
// Conserva il NOME reale della sede (com'è su SharePoint), così le sedi nuove
// non vengono più schiacciate su un id fisso. Solo il valore speciale "tutte"
// (admin senza sede operativa) viene normalizzato.
function normalizeSede(v: SedeRaw): string {
  const s = (v ?? "").toString().trim();
  if (!s) return "";
  const low = s.toLowerCase();
  if (low === "tutte" || low === "all" || low === "*") return "tutte";
  return s;
}

function requireField(
  map: Record<string, string>,
  key: string,
  list: "Dipendenti" | "Timbrature" | "Richieste",
): string {
  const v = map[key];
  if (!v)
    throw new Error(`Colonna obbligatoria "${key}" mancante nella lista ${list} su SharePoint.`);
  return v;
}

// ---------------------------------------------------------------------------
// Dipendenti
// ---------------------------------------------------------------------------
export interface SpDipendente {
  id: string;
  nome: string;
  cognome: string;
  nomeCompleto: string;
  email: string;
  sede: string;
  attivo: boolean;
  ruolo: string;
  // Visibilità nelle viste operative (dashboard, elenchi, conteggi, report).
  // Fail-open: se la colonna manca o è vuota il dipendente è considerato
  // VISIBILE, così una dimenticanza di backfill non svuota la dashboard.
  // NB: `visibile` NON governa l'accesso — l'autenticazione dipende solo da
  // `attivo`. Sono due assi ortogonali.
  visibile: boolean;
  // Flag per la futura autorizzazione di ferie/permessi/straordinari
  // (modulo Richieste, non ancora implementato). Default: false.
  autorizza: boolean;
  // Operatore/back-office (DR000 Lucrezia): può inserire/correggere timbrature
  // manuali. Default false. L'autorizzazione effettiva è ri-verificata sul
  // server nelle operazioni sensibili, non solo qui.
  operatore: boolean;
  // Ore contrattuali settimanali (full-time e part-time). null se non impostate.
  // Usate da rilevazione anomalie e rendiconto.
  oreSettimanali: number | null;
  // Inquadramento contrattuale (es. livello/qualifica CCNL). Puramente
  // informativo per ora: nessuna logica lo usa. Colonna SharePoint OPZIONALE —
  // se assente/vuota → "" (non entra nei controlli di salute).
  inquadramento: string;
  // Giorni di ferie spettanti nell'anno (per il saldo residuo). null se non
  // impostato → si usa il default DEFAULT_FERIE_ANNUE.
  giorniFerieAnnui: number | null;
  // Ore di permesso annue (ROL/ex festività). null se non impostate → il
  // residuo permessi non viene mostrato (nessun default inventato).
  orePermessiAnnui: number | null;
  // Codice fiscale (per l'abbinamento automatico delle buste paga). Colonna
  // OPZIONALE; vuoto se non impostato.
  cf: string;
  // Codice dipendente (es. DR005). Serve al client per il gating dei moduli
  // riservati al direttore; l'autorizzazione effettiva è comunque ri-verificata
  // server-side sul record SharePoint.
  codice: string;
}

// Parsing tollerante di un campo booleano SharePoint (Sì/No).
// `undefined` (colonna assente/mai valorizzata) → valore di default fornito.
function parseSpBool(raw: unknown, whenMissing: boolean): boolean {
  if (raw === undefined || raw === null || raw === "") return whenMissing;
  return Boolean(raw);
}

// Parsing tollerante di un campo numerico SharePoint. Vuoto/assente → default.
function parseSpNumber(raw: unknown, whenMissing: number | null): number | null {
  if (raw === undefined || raw === null || raw === "") return whenMissing;
  const n = Number(raw);
  return Number.isFinite(n) ? n : whenMissing;
}

export async function fetchDipendenti(): Promise<SpDipendente[]> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  const F = cfg.dipendentiFields;
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items?expand=fields&$top=999`,
    ),
  );
  const out = res.value
    .map((it) => {
      const f = it.fields ?? {};
      const nome = String(f[F.Nome ?? ""] ?? "").trim();
      const cognome = String(f[F.Cognome ?? ""] ?? "").trim();
      const nomeCompleto = String(f[F.NomeCompleto ?? ""] ?? `${nome} ${cognome}`).trim();
      const rawAttivo = F.Attivo ? f[F.Attivo] : undefined;
      const attivo = rawAttivo === undefined ? true : Boolean(rawAttivo);
      return {
        id: String(it.id),
        nome,
        cognome,
        nomeCompleto,
        email: String(f[F.Email ?? ""] ?? "").trim(),
        sede: normalizeSede((F.Sede ? f[F.Sede] : undefined) as SedeRaw),
        attivo,
        ruolo: String(f[F.Ruolo ?? ""] ?? "").trim(),
        // Fail-open sulla visibilità; autorizza/operatore default false.
        visibile: parseSpBool(F.Visibile ? f[F.Visibile] : undefined, true),
        autorizza: parseSpBool(F.Autorizza ? f[F.Autorizza] : undefined, false),
        operatore: parseSpBool(F.Operatore ? f[F.Operatore] : undefined, false),
        oreSettimanali: parseSpNumber(F.OreSettimanali ? f[F.OreSettimanali] : undefined, null),
        inquadramento: String(f[F.Inquadramento ?? ""] ?? "").trim(),
        giorniFerieAnnui: parseSpNumber(
          F.GiorniFerieAnnui ? f[F.GiorniFerieAnnui] : undefined,
          null,
        ),
        orePermessiAnnui: parseSpNumber(
          F.OrePermessiAnnui ? f[F.OrePermessiAnnui] : undefined,
          null,
        ),
        cf: String(f[F.CF ?? ""] ?? "")
          .trim()
          .toUpperCase(),
        codice: normalizeCodice(F.Codice ? f[F.Codice] : ""),
      };
    })
    .filter((d) => d.attivo);
  logSp("info", "fetch.dipendenti", `${out.length} dipendenti attivi`, {
    durataMs: Date.now() - started,
  });
  return out;
}

// ---------------------------------------------------------------------------
// Import massivo Dipendenti (admin) — incolla CSV/TSV dal pannello
// Amministrazione. Riusa le credenziali server del portale (nessun secret da
// reperire) e lo stesso gateway. Con dryRun non scrive nulla: solo anteprima.
// ---------------------------------------------------------------------------
const IMPORT_BOOL_COLS = new Set(["attivo", "visibile", "autorizza", "operatore"]);
const IMPORT_NUM_COLS = new Set(["oresettimanali"]);
const IMPORT_MAX_ROWS = 500;

type ImportFieldValue = string | number | boolean;
export interface ImportRowResult {
  label: string;
  ok: boolean;
  error?: string;
  preview?: Record<string, ImportFieldValue>;
}
export interface ImportDipendentiResult {
  dryRun: boolean;
  matchedColumns: string[];
  missingColumns: string[];
  // Colonne SCRIVIBILI realmente esposte da SharePoint (per diagnosticare i
  // mismatch di intestazione: nome diverso, colonna calcolata/nascosta, ecc.).
  availableColumns: string[];
  totalRows: number;
  created: number;
  failed: number;
  rows: ImportRowResult[];
}

// Rileva il separatore dall'intestazione: TAB (incolla da Excel), punto e
// virgola (default CSV Excel italiano) o virgola.
function detectDelim(text: string): "," | "\t" | ";" {
  const nl = text.indexOf("\n");
  const firstLine = nl >= 0 ? text.slice(0, nl) : text;
  if (firstLine.includes("\t")) return "\t";
  if (firstLine.includes(";")) return ";";
  return ",";
}

// Parser tollerante ai campi tra virgolette. Separatore singolo (`,`, `;` o TAB).
function parseDelimited(text: string, delim: "," | "\t" | ";"): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') inQ = false;
      else field += c;
    } else if (c === '"') inQ = true;
    else if (c === delim) {
      row.push(field);
      field = "";
    } else if (c === "\r") {
      /* skip */
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += c;
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((r) => r.some((x) => x.trim() !== ""));
}

// Normalizza intestazione/nome colonna per il confronto: converte gli spazi
// unicode invisibili (NBSP, zero-width, BOM…) in spazio normale, collassa gli
// spazi e ignora le maiuscole. Evita i falsi "colonna mancante" causati da
// caratteri non visibili incollati da Excel/Word.
function normKey(s: string): string {
  let out = "";
  for (const ch of s) {
    const code = ch.codePointAt(0) ?? 0;
    // Spazi/caratteri invisibili unicode (NBSP, zero-width, BOM…) → spazio.
    const invisible =
      code === 0x00a0 ||
      code === 0x00ad ||
      code === 0x200b ||
      code === 0x2060 ||
      code === 0xfeff ||
      code === 0x202f ||
      code === 0x205f ||
      code === 0x3000 ||
      (code >= 0x2000 && code <= 0x200a);
    out += invisible ? " " : ch;
  }
  return out.replace(/\s+/g, " ").trim().toLowerCase();
}

export async function importDipendenti(
  csvText: string,
  dryRun: boolean,
): Promise<ImportDipendentiResult> {
  const cfg = await discoverSharePoint();
  const listId = cfg.listDipendenti;

  // Colonne reali della lista (display + internal), escluse read-only/nascoste.
  const colsRes = await withDiscoveryRetry(() =>
    gatewayJson<{
      value?: Array<{
        name?: string;
        displayName?: string;
        readOnly?: boolean;
        hidden?: boolean;
      }>;
    }>(`/sites/${cfg.siteId}/lists/${listId}/columns?$select=name,displayName,readOnly,hidden`),
  );
  const internalByLabel = new Map<string, string>();
  const availableSet = new Set<string>();
  for (const c of colsRes.value ?? []) {
    if (c.hidden || c.readOnly) continue;
    if (c.displayName && c.name) internalByLabel.set(normKey(c.displayName), c.name);
    if (c.name) internalByLabel.set(normKey(c.name), c.name);
    if (c.displayName) availableSet.add(c.displayName);
    else if (c.name) availableSet.add(c.name);
  }
  const availableColumns = [...availableSet].sort((a, b) => a.localeCompare(b));

  const delim = detectDelim(csvText);
  const rows = parseDelimited(csvText, delim);
  if (rows.length < 2)
    throw new Error("Il testo deve avere l'intestazione + almeno una riga di dati.");
  const header = rows[0].map((h) => h.trim());
  const dataRows = rows.slice(1);
  if (dataRows.length > IMPORT_MAX_ROWS)
    throw new Error(`Troppe righe (${dataRows.length}); massimo ${IMPORT_MAX_ROWS} per import.`);

  const matchedColumns: string[] = [];
  const missingColumns: string[] = [];
  for (const h of header) {
    if (internalByLabel.has(normKey(h))) matchedColumns.push(h);
    else missingColumns.push(h);
  }
  // Se un'intestazione non corrisponde a una colonna, NON importo nulla: c'è il
  // rischio di disallineamento. L'admin corregge le intestazioni e riprova.
  if (missingColumns.length) {
    return {
      dryRun,
      matchedColumns,
      missingColumns,
      availableColumns,
      totalRows: dataRows.length,
      created: 0,
      failed: 0,
      rows: [],
    };
  }

  const codiceInt = internalByLabel.get("codice");
  const cognomeInt = internalByLabel.get("cognome");
  const nomeInt = internalByLabel.get("nome");

  const buildFields = (r: string[]): Record<string, ImportFieldValue> => {
    const fields: Record<string, ImportFieldValue> = {};
    header.forEach((h, i) => {
      const key = normKey(h);
      const internal = internalByLabel.get(key)!;
      const raw = (r[i] ?? "").trim();
      if (IMPORT_BOOL_COLS.has(key)) {
        fields[internal] = /^(s[iì]|true|1|x)$/i.test(raw);
      } else if (IMPORT_NUM_COLS.has(key)) {
        if (raw === "") return;
        const n = Number(raw.replace(",", "."));
        if (Number.isFinite(n)) fields[internal] = n;
      } else if (raw !== "") {
        fields[internal] = raw;
      }
    });
    return fields;
  };
  const labelOf = (fields: Record<string, ImportFieldValue>): string =>
    [codiceInt && fields[codiceInt], cognomeInt && fields[cognomeInt], nomeInt && fields[nomeInt]]
      .filter(Boolean)
      .join(" ") || "(riga)";

  const outRows: ImportRowResult[] = [];
  let created = 0;
  let failed = 0;

  for (const r of dataRows) {
    const fields = buildFields(r);
    const label = labelOf(fields);
    if (dryRun) {
      outRows.push({ label, ok: true, preview: fields });
      continue;
    }
    try {
      await withDiscoveryRetry(() =>
        gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
          method: "POST",
          body: JSON.stringify({ fields }),
        }),
      );
      outRows.push({ label, ok: true });
      created++;
    } catch (err) {
      outRows.push({ label, ok: false, error: err instanceof Error ? err.message : String(err) });
      failed++;
    }
  }

  logSp(
    "info",
    "import.dipendenti",
    dryRun
      ? `Anteprima import: ${dataRows.length} righe`
      : `Import dipendenti: creati ${created}, errori ${failed}`,
  );

  return {
    dryRun,
    matchedColumns,
    missingColumns,
    availableColumns,
    totalRows: dataRows.length,
    created,
    failed,
    rows: outRows,
  };
}

// ---------------------------------------------------------------------------
// Login locale (Codice + PIN) — verifica lato server contro la lista
// SharePoint "Dipendenti". Non espone MAI il PIN al client.
// ---------------------------------------------------------------------------
export interface LoginResult {
  ok: boolean;
  dipendente?: SpDipendente;
  error?: string;
}

function normalizeCodice(v: unknown): string {
  return String(v ?? "")
    .trim()
    .toUpperCase();
}

function normalizePin(v: unknown): string {
  return String(v ?? "").trim();
}

// --- S3: protezione PIN (hash) + rate limiting ------------------------------
// Il PIN è salvato come `sha256$<salt>$<hash>` dove hash = SHA-256(salt:pin:pepper).
// Il "pepper" è un segreto solo-server: chi legge la lista SharePoint non può
// forzare i PIN (4 cifre) senza di esso. I PIN in chiaro (legacy) vengono
// migrati ad hash al primo login riuscito, senza azioni manuali.
const PIN_HASH_PREFIX = "sha256$";
function pinPepper(): string {
  return process.env.SESSION_SECRET || process.env.MICROSOFT_SHAREPOINT_API_KEY || "";
}
function bytesToHex(b: Uint8Array): string {
  return Array.from(b)
    .map((x) => x.toString(16).padStart(2, "0"))
    .join("");
}
async function hashPin(pin: string, saltHex: string): Promise<string> {
  const data = new TextEncoder().encode(`${saltHex}:${pin}:${pinPepper()}`);
  const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
  return bytesToHex(new Uint8Array(await crypto.subtle.digest("SHA-256", buf)));
}
async function makeStoredPin(pin: string): Promise<string> {
  const salt = bytesToHex(crypto.getRandomValues(new Uint8Array(8)));
  return `${PIN_HASH_PREFIX}${salt}$${await hashPin(pin, salt)}`;
}
async function verifyStoredPin(input: string, stored: string): Promise<boolean> {
  if (stored.startsWith(PIN_HASH_PREFIX)) {
    const parts = stored.split("$");
    if (parts.length !== 3) return false;
    return (await hashPin(input, parts[1])) === parts[2];
  }
  return stored === input; // legacy: PIN ancora in chiaro
}

// Rate limiting per codice: max 5 tentativi falliti in 10 minuti. In-memory
// (best-effort sugli isolate dei Workers), sufficiente a fermare i tentativi
// manuali ripetuti.
const loginAttempts = new Map<string, { count: number; resetAt: number }>();
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_WINDOW_MS = 10 * 60 * 1000;

// Protezione massiva (admin): converte in hash tutti i PIN ancora in chiaro.
export async function protectAllPins(): Promise<{ protetti: number; giaProtetti: number }> {
  const cfg = await discoverSharePoint();
  const F = cfg.dipendentiFields;
  const pinField = F.PIN;
  if (!pinField) throw new Error('Colonna "PIN" non trovata su Dipendenti.');
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items?expand=fields&$top=999`,
    ),
  );
  let protetti = 0;
  let giaProtetti = 0;
  for (const it of res.value) {
    const raw = normalizePin((it.fields ?? {})[pinField]);
    if (!raw) continue;
    if (raw.startsWith(PIN_HASH_PREFIX)) {
      giaProtetti++;
      continue;
    }
    await gatewayJson(`/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items/${it.id}/fields`, {
      method: "PATCH",
      body: JSON.stringify({ [pinField]: await makeStoredPin(raw) }),
    });
    protetti++;
  }
  logSp("info", "pin.protect", `PIN protetti: ${protetti} (già protetti: ${giaProtetti})`);
  return { protetti, giaProtetti };
}

export async function loginByCodicePin(
  codiceInput: string,
  pinInput: string,
): Promise<LoginResult> {
  const started = Date.now();
  const codice = normalizeCodice(codiceInput);
  const pin = normalizePin(pinInput);
  if (!codice || !pin) {
    return { ok: false, error: "Codice o PIN non validi." };
  }
  const cfg = await discoverSharePoint();
  const F = cfg.dipendentiFields;
  const codiceField = F.Codice;
  const pinField = F.PIN;
  if (!codiceField || !pinField) {
    logSp(
      "error",
      "login",
      `Colonne login mancanti su Dipendenti (Codice=${!!codiceField}, PIN=${!!pinField}).`,
    );
    return {
      ok: false,
      error:
        'Login non configurato: aggiungere le colonne "Codice" e "PIN" alla lista SharePoint "Dipendenti".',
    };
  }
  // Rate limiting: blocca il codice dopo troppi tentativi falliti ravvicinati.
  const now = Date.now();
  const rl = loginAttempts.get(codice);
  if (rl && rl.resetAt > now && rl.count >= LOGIN_MAX_ATTEMPTS) {
    const min = Math.max(1, Math.ceil((rl.resetAt - now) / 60000));
    logSp("warn", "login", `Rate limit attivo per codice="${codice}"`);
    return { ok: false, error: `Troppi tentativi falliti. Riprova tra ${min} minuti.` };
  }

  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items?expand=fields&$top=999`,
    ),
  );
  const attivoField = F.Attivo;
  const candidato = res.value.find((it) => {
    const f = it.fields ?? {};
    const attivo = attivoField ? Boolean(f[attivoField]) : true;
    return attivo && normalizeCodice(f[codiceField]) === codice;
  });
  const storedPin = candidato ? normalizePin((candidato.fields ?? {})[pinField]) : "";
  const pinOk = candidato && storedPin ? await verifyStoredPin(pin, storedPin) : false;
  if (!candidato || !pinOk) {
    const cur = loginAttempts.get(codice);
    if (cur && cur.resetAt > now) cur.count++;
    else loginAttempts.set(codice, { count: 1, resetAt: now + LOGIN_WINDOW_MS });
    logSp("warn", "login", `Tentativo fallito per codice="${codice}"`, {
      durataMs: Date.now() - started,
    });
    return { ok: false, error: "Codice o PIN non validi." };
  }
  loginAttempts.delete(codice);
  const match = candidato;

  // Migrazione trasparente S3: PIN ancora in chiaro → sostituito con l'hash
  // al primo login riuscito (best-effort, non blocca l'accesso).
  if (!storedPin.startsWith(PIN_HASH_PREFIX)) {
    try {
      await gatewayJson(
        `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items/${match.id}/fields`,
        { method: "PATCH", body: JSON.stringify({ [pinField]: await makeStoredPin(pin) }) },
      );
      logSp("info", "login", `PIN protetto (hash) per ${codice}`);
    } catch {
      /* riproverà al prossimo login */
    }
  }
  const f = match.fields ?? {};
  const nome = String(f[F.Nome ?? ""] ?? "").trim();
  const cognome = String(f[F.Cognome ?? ""] ?? "").trim();
  const dipendente: SpDipendente = {
    id: String(match.id),
    nome,
    cognome,
    nomeCompleto: String(f[F.NomeCompleto ?? ""] ?? `${nome} ${cognome}`).trim(),
    email: String(f[F.Email ?? ""] ?? "").trim(),
    sede: normalizeSede((F.Sede ? f[F.Sede] : undefined) as SedeRaw),
    attivo: true,
    ruolo: String(f[F.Ruolo ?? ""] ?? "").trim(),
    // Popolati per coerenza del modello. NON influenzano l'esito del login:
    // un utente con visibile=false può comunque autenticarsi (regola 2).
    visibile: parseSpBool(F.Visibile ? f[F.Visibile] : undefined, true),
    autorizza: parseSpBool(F.Autorizza ? f[F.Autorizza] : undefined, false),
    operatore: parseSpBool(F.Operatore ? f[F.Operatore] : undefined, false),
    oreSettimanali: parseSpNumber(F.OreSettimanali ? f[F.OreSettimanali] : undefined, null),
    inquadramento: String(f[F.Inquadramento ?? ""] ?? "").trim(),
    giorniFerieAnnui: parseSpNumber(F.GiorniFerieAnnui ? f[F.GiorniFerieAnnui] : undefined, null),
    orePermessiAnnui: parseSpNumber(F.OrePermessiAnnui ? f[F.OrePermessiAnnui] : undefined, null),
    cf: String(f[F.CF ?? ""] ?? "")
      .trim()
      .toUpperCase(),
    codice,
  };
  logSp("info", "login", `Login ok per ${codice} (id=${dipendente.id})`, {
    durataMs: Date.now() - started,
  });
  return { ok: true, dipendente };
}

// ---------------------------------------------------------------------------
// Timbrature
// ---------------------------------------------------------------------------
export type EventoTimbratura = "entrata" | "inizio-pausa" | "fine-pausa" | "uscita";

export interface SpTimbratura {
  id: string;
  dipendenteId: string;
  evento: EventoTimbratura;
  dataOra: string; // ISO
  origine?: string;
  posizione?: string;
  esito?: string;
  note?: string;
}

function todayIsoStart(): string {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d.toISOString();
}

function parseEvento(v: unknown): EventoTimbratura | null {
  const s = String(v ?? "")
    .toLowerCase()
    .replace(/\s+/g, "-");
  if (s === "entrata" || s === "inizio-pausa" || s === "fine-pausa" || s === "uscita") return s;
  return null;
}

function eventoToSharePoint(e: EventoTimbratura): string {
  return {
    entrata: "Entrata",
    "inizio-pausa": "Inizio Pausa",
    "fine-pausa": "Fine Pausa",
    uscita: "Uscita",
  }[e];
}

// Costruisce il nome del lookup id partendo dall'internal name della colonna.
// Esempio: internal "Dipendente" -> "DipendenteLookupId"; internal "Dipendente0"
// -> "Dipendente0LookupId".
function lookupIdFieldName(internal: string): string {
  return `${internal}LookupId`;
}

// Legge le timbrature con DataOra >= fromISO (ordinate crescenti).
async function fetchTimbratureDaISO(fromISO: string): Promise<SpTimbratura[]> {
  const cfg = await discoverSharePoint();
  const F = cfg.timbratureFields;
  const dataOraField = requireField(F, "DataOra", "Timbrature");
  const eventoField = requireField(F, "Evento", "Timbrature");
  const dipendenteField = requireField(F, "Dipendente", "Timbrature");
  const lookupId = lookupIdFieldName(dipendenteField);

  const filter = encodeURIComponent(`fields/${dataOraField} ge '${fromISO}'`);
  const basePath = `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items?expand=fields&$top=999`;
  const filteredPath = `${basePath}&$orderby=fields/${dataOraField} asc&$filter=${filter}`;
  let res: GraphListResponse<Record<string, unknown>>;
  try {
    res = await withDiscoveryRetry(() =>
      gatewayJson<GraphListResponse<Record<string, unknown>>>(filteredPath),
    );
  } catch {
    res = await withDiscoveryRetry(() =>
      gatewayJson<GraphListResponse<Record<string, unknown>>>(basePath),
    );
  }
  const startMs = new Date(fromISO).getTime();
  return res.value
    .map((it): SpTimbratura | null => {
      const f = it.fields ?? {};
      const evento = parseEvento(f[eventoField]);
      const dataOra = String(f[dataOraField] ?? "");
      const dipRaw = f[lookupId];
      return evento && dataOra && dipRaw != null
        ? {
            id: String(it.id),
            dipendenteId: String(dipRaw),
            evento,
            dataOra,
            origine: F.Origine ? (f[F.Origine] as string | undefined) : undefined,
            posizione: F.Posizione ? (f[F.Posizione] as string | undefined) : undefined,
            esito: F.Esito ? (f[F.Esito] as string | undefined) : undefined,
            note: F.Note ? (f[F.Note] as string | undefined) : undefined,
          }
        : null;
    })
    .filter((x): x is SpTimbratura => x !== null && new Date(x.dataOra).getTime() >= startMs)
    .sort((a, b) => a.dataOra.localeCompare(b.dataOra));
}

export async function fetchTimbratureOggi(): Promise<SpTimbratura[]> {
  const started = Date.now();
  const out = await fetchTimbratureDaISO(todayIsoStart());
  logSp("info", "fetch.timbrature", `${out.length} timbrature oggi`, {
    durataMs: Date.now() - started,
  });
  return out;
}

// ---------------------------------------------------------------------------
// Anomalie giornaliere (Sprint 3, on-read) — vista operatore.
// ---------------------------------------------------------------------------
export interface AnomaliaItem {
  dipendenteId: string;
  nomeCompleto: string;
  sede: string; // nome sede (come su SharePoint)
  data: string; // YYYY-MM-DD
  tipo: TipoAnomalia;
}

export async function computeAnomalie(giorni = 14): Promise<AnomaliaItem[]> {
  const started = Date.now();
  // Finestra: dagli ultimi `giorni` fino a IERI (oggi è in corso → escluso).
  const from = new Date();
  from.setHours(0, 0, 0, 0);
  from.setDate(from.getDate() - giorni);
  const [tims, dips] = await Promise.all([
    fetchTimbratureDaISO(from.toISOString()),
    fetchDipendenti(),
  ]);
  const byId = new Map(dips.map((d) => [d.id, d]));
  const todayStr = new Date().toISOString().slice(0, 10);

  // Raggruppa per dipendente + giorno (esclude oggi/futuro).
  const groups = new Map<string, { dipId: string; giorno: string; eventi: EventoTimbratura[] }>();
  for (const t of tims) {
    const giorno = t.dataOra.slice(0, 10);
    if (giorno >= todayStr) continue;
    const key = `${t.dipendenteId}|${giorno}`;
    let g = groups.get(key);
    if (!g) {
      g = { dipId: t.dipendenteId, giorno, eventi: [] };
      groups.set(key, g);
    }
    g.eventi.push(t.evento);
  }

  const out: AnomaliaItem[] = [];
  for (const g of groups.values()) {
    const dip = byId.get(g.dipId);
    const ore = dip?.oreSettimanali ?? null;
    const rilevaPausa = !(ore != null && ore <= 16);
    for (const tipo of anomalieDelGiorno(g.eventi, { rilevaPausa })) {
      out.push({
        dipendenteId: g.dipId,
        nomeCompleto: dip ? dip.nomeCompleto || `${dip.nome} ${dip.cognome}` : `#${g.dipId}`,
        sede: dip?.sede ?? "",
        data: g.giorno,
        tipo,
      });
    }
  }
  out.sort((a, b) =>
    a.data === b.data ? a.nomeCompleto.localeCompare(b.nomeCompleto) : b.data.localeCompare(a.data),
  );
  logSp("info", "anomalie", `${out.length} anomalie (ultimi ${giorni}g)`, {
    durataMs: Date.now() - started,
  });
  return out;
}

// ---------------------------------------------------------------------------
// Supervisore (Sprint 3, DR005/Francesco): timbrature manuali per visione.
// ---------------------------------------------------------------------------
export interface TimbraturaManualeItem {
  id: string;
  dipendenteId: string;
  nomeCompleto: string;
  sede: string; // id sede
  evento: EventoTimbratura;
  dataOra: string; // ISO
  note?: string;
}

export async function fetchTimbratureManuali(giorni = 30): Promise<TimbraturaManualeItem[]> {
  const from = new Date();
  from.setHours(0, 0, 0, 0);
  from.setDate(from.getDate() - giorni);
  const [tims, dips] = await Promise.all([
    fetchTimbratureDaISO(from.toISOString()),
    fetchDipendenti(),
  ]);
  const byId = new Map(dips.map((d) => [d.id, d]));
  return tims
    .filter((t) => (t.origine ?? "").toLowerCase() === "manuale")
    .map((t) => {
      const d = byId.get(t.dipendenteId);
      return {
        id: t.id,
        dipendenteId: t.dipendenteId,
        nomeCompleto: d ? d.nomeCompleto || `${d.nome} ${d.cognome}` : `#${t.dipendenteId}`,
        sede: d?.sede ?? "",
        evento: t.evento,
        dataOra: t.dataOra,
        note: t.note,
      };
    })
    .sort((a, b) => b.dataOra.localeCompare(a.dataOra));
}

// ---------------------------------------------------------------------------
// Rendiconto mensile (riscontro settimanale a monte ore).
// ---------------------------------------------------------------------------
export interface RendicontoRiga {
  dipendenteId: string;
  nomeCompleto: string;
  sede: string;
  oreSettimanali: number | null;
  oreLavorate: number; // effettive dal timbrature (giorni chiusi del mese)
  straordinarioCalcolato: number; // dalle timbrature (settimane con lunedì nel mese)
  straordinarioAutorizzato: number; // da richieste Straordinario approvate (mese)
  permessiOre: number;
  ferieGiorni: number;
  malattiaGiorni: number;
  giorniNonChiusi: number; // giornate con turno aperto (ore non calcolabili)
}

// Giorni di ferie annui di default, se il dipendente non ha la colonna
// GiorniFerieAnnui valorizzata su SharePoint.
export const DEFAULT_FERIE_ANNUE = 26;

export interface SaldoFerieRiga {
  dipendenteId: string;
  nomeCompleto: string;
  sede: string;
  spettanti: number;
  godute: number;
  residui: number;
  // Permessi (ore). null se OrePermessiAnnui non è impostato per il dipendente.
  permessiSpettantiOre: number | null;
  permessiGoduteOre: number;
  permessiResiduiOre: number | null;
}

// Saldo ferie e permessi per l'anno: spettanti (da colonna o default) meno
// quanto approvato nell'anno.
export async function computeSaldoFerie(anno: number): Promise<SaldoFerieRiga[]> {
  const [dipendenti, richieste] = await Promise.all([
    fetchDipendenti(),
    fetchRichieste({ stato: "Approvata" }),
  ]);
  const goduteById = new Map<string, number>();
  const permessiById = new Map<string, number>();
  for (const r of richieste) {
    if (Number((r.dataInizio || "").slice(0, 4)) !== anno) continue;
    if (r.tipo === "Ferie") {
      const gg =
        r.durataGiorni && r.durataGiorni > 0
          ? r.durataGiorni
          : computeDurataGiorni(
              r.dataInizio.slice(0, 10),
              (r.dataFine || r.dataInizio).slice(0, 10),
            );
      goduteById.set(r.richiedenteId, (goduteById.get(r.richiedenteId) ?? 0) + gg);
    } else if (r.tipo === "Permesso") {
      const ore =
        r.durataOre && r.durataOre > 0
          ? r.durataOre
          : r.oraInizio && r.oraFine
            ? computeDurataOre(r.oraInizio, r.oraFine)
            : 0;
      permessiById.set(r.richiedenteId, (permessiById.get(r.richiedenteId) ?? 0) + ore);
    }
  }
  return dipendenti
    .filter((d) => d.visibile)
    .map((d) => {
      const spettanti = d.giorniFerieAnnui ?? DEFAULT_FERIE_ANNUE;
      const godute = goduteById.get(d.id) ?? 0;
      const permessiGoduteOre = round2(permessiById.get(d.id) ?? 0);
      const permessiSpettantiOre = d.orePermessiAnnui;
      return {
        dipendenteId: d.id,
        nomeCompleto: d.nomeCompleto || `${d.cognome} ${d.nome}`,
        sede: d.sede,
        spettanti,
        godute,
        residui: spettanti - godute,
        permessiSpettantiOre,
        permessiGoduteOre,
        permessiResiduiOre:
          permessiSpettantiOre != null ? round2(permessiSpettantiOre - permessiGoduteOre) : null,
      };
    })
    .sort((a, b) => a.nomeCompleto.localeCompare(b.nomeCompleto));
}

function eachDay(fromStr: string, toStr: string): string[] {
  const out: string[] = [];
  const d = new Date(`${fromStr}T00:00:00`);
  const end = new Date(`${toStr}T00:00:00`).getTime();
  while (d.getTime() <= end) {
    out.push(ymd(d));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

export async function computeRendiconto(anno: number, mese: number): Promise<RendicontoRiga[]> {
  const monthStart = new Date(anno, mese - 1, 1);
  const monthEnd = new Date(anno, mese, 0);
  return computeRendicontoPeriodo(ymd(monthStart), ymd(monthEnd));
}

// Rendiconto su un periodo arbitrario (mese, settimana fiscale o settimana del
// mese). Le settimane contano nel periodo in cui cade il loro LUNEDÌ.
export async function computeRendicontoPeriodo(
  monthStartStr: string,
  monthEndStr: string,
): Promise<RendicontoRiga[]> {
  const monthStart = new Date(`${monthStartStr}T00:00:00`);
  const monthEnd = new Date(`${monthEndStr}T00:00:00`);
  // Estendi alle settimane complete (per lo straordinario settimanale).
  const from = new Date(monthStart);
  from.setDate(from.getDate() - ((from.getDay() === 0 ? 7 : from.getDay()) - 1));
  const to = new Date(monthEnd);
  to.setDate(to.getDate() + (7 - (to.getDay() === 0 ? 7 : to.getDay())));
  const fromStr = ymd(from);
  const toStr = ymd(to);

  const [tims, richieste, dips] = await Promise.all([
    fetchTimbratureDaISO(
      new Date(from.getFullYear(), from.getMonth(), from.getDate()).toISOString(),
    ),
    fetchRichieste({}),
    fetchDipendenti(),
  ]);

  // Eventi per dipendente+giorno (giorno = data locale).
  const eventiByDipDay = new Map<string, { evento: EventoTimbratura; ora: string }[]>();
  for (const t of tims) {
    const giorno = ymd(new Date(t.dataOra));
    if (giorno < fromStr || giorno > toStr) continue;
    const key = `${t.dipendenteId}|${giorno}`;
    const arr = eventiByDipDay.get(key) ?? [];
    arr.push({ evento: t.evento, ora: t.dataOra });
    eventiByDipDay.set(key, arr);
  }

  // Assenze/ore da richieste (per dipendente+giorno).
  const ferie = new Set<string>();
  const malattia = new Set<string>();
  const permessoOre = new Map<string, number>();
  const straordAut = new Map<string, number>();
  for (const r of richieste) {
    const di = r.dataInizio.slice(0, 10);
    const df = (r.dataFine || r.dataInizio).slice(0, 10);
    if (r.tipo === "Ferie" && r.stato === "Approvata") {
      for (const g of eachDay(di, df)) ferie.add(`${r.richiedenteId}|${g}`);
    } else if (r.tipo === "Malattia" && (r.stato === "Comunicata" || r.stato === "Approvata")) {
      for (const g of eachDay(di, df)) malattia.add(`${r.richiedenteId}|${g}`);
    } else if (r.tipo === "Permesso" && r.stato === "Approvata") {
      const k = `${r.richiedenteId}|${di}`;
      permessoOre.set(k, (permessoOre.get(k) ?? 0) + (r.durataOre ?? 0));
    } else if (r.tipo === "Straordinario" && r.stato === "Approvata") {
      const k = `${r.richiedenteId}|${di}`;
      straordAut.set(k, (straordAut.get(k) ?? 0) + (r.durataOre ?? 0));
    }
  }

  const inMonth = (g: string) => g >= monthStartStr && g <= monthEndStr;
  const out: RendicontoRiga[] = [];
  for (const d of dips) {
    const dipId = d.id;
    let oreLavorate = 0;
    let giorniNonChiusi = 0;
    let straordinarioCalcolato = 0;
    let straordinarioAutorizzato = 0;
    let permessi = 0;
    let ferieGiorni = 0;
    let malattiaGiorni = 0;

    // Ore lavorate per giorno su tutto il range esteso (servono al calcolo
    // settimanale); nel totale mensile contano solo i giorni del mese.
    const oreGiorno = new Map<string, number>();
    for (const g of eachDay(fromStr, toStr)) {
      const ev = eventiByDipDay.get(`${dipId}|${g}`);
      if (!ev || ev.length === 0) continue;
      const ore = oreLavorateGiorno(ev);
      if (ore == null) {
        if (inMonth(g)) giorniNonChiusi++;
        continue;
      }
      oreGiorno.set(g, ore);
      if (inMonth(g)) oreLavorate += ore;
    }

    // Metriche mensili (calendario) dalle richieste.
    for (const g of eachDay(monthStartStr, monthEndStr)) {
      if (ferie.has(`${dipId}|${g}`)) ferieGiorni++;
      if (malattia.has(`${dipId}|${g}`)) malattiaGiorni++;
      permessi += permessoOre.get(`${dipId}|${g}`) ?? 0;
      straordinarioAutorizzato += straordAut.get(`${dipId}|${g}`) ?? 0;
    }

    // Straordinario calcolato: settimane il cui lunedì cade nel mese.
    if (d.oreSettimanali != null) {
      const weeks = new Map<string, string[]>();
      for (const g of eachDay(fromStr, toStr)) {
        const lun = lunediDellaSettimana(g);
        const arr = weeks.get(lun) ?? [];
        arr.push(g);
        weeks.set(lun, arr);
      }
      for (const [lun, giorni] of weeks) {
        if (!inMonth(lun)) continue;
        let lunSab = 0;
        let dom = 0;
        let assenze = 0;
        let permW = 0;
        for (const g of giorni) {
          const ore = oreGiorno.get(g) ?? 0;
          if (isoDow(g) === 7) dom += ore;
          else lunSab += ore;
          if (ferie.has(`${dipId}|${g}`) || malattia.has(`${dipId}|${g}`)) assenze++;
          permW += permessoOre.get(`${dipId}|${g}`) ?? 0;
        }
        const prev = orePrevisteSettimana(d.oreSettimanali, assenze, permW);
        straordinarioCalcolato += straordinarioSettimana(lunSab, dom, prev);
      }
    }

    out.push({
      dipendenteId: dipId,
      nomeCompleto: d.nomeCompleto || `${d.cognome} ${d.nome}`,
      sede: d.sede,
      oreSettimanali: d.oreSettimanali,
      oreLavorate: round2(oreLavorate),
      straordinarioCalcolato: round2(straordinarioCalcolato),
      straordinarioAutorizzato: round2(straordinarioAutorizzato),
      permessiOre: round2(permessi),
      ferieGiorni,
      malattiaGiorni,
      giorniNonChiusi,
    });
  }
  out.sort((a, b) => a.nomeCompleto.localeCompare(b.nomeCompleto));
  logSp("info", "rendiconto", `Rendiconto ${monthStartStr}→${monthEndStr}: ${out.length} righe`);
  return out;
}

export interface CreateTimbraturaInput {
  dipendenteId: string;
  evento: EventoTimbratura;
  origine?: string;
  posizione?: string;
  esito?: string;
  note?: string;
}

export async function createTimbratura(input: CreateTimbraturaInput): Promise<SpTimbratura> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  const F = cfg.timbratureFields;
  const dipendenteField = requireField(F, "Dipendente", "Timbrature");
  const eventoField = requireField(F, "Evento", "Timbrature");
  const dataOraField = requireField(F, "DataOra", "Timbrature");
  const dipInt = Number(input.dipendenteId);
  if (!Number.isFinite(dipInt))
    throw new Error("dipendenteId non valido per SharePoint (atteso ID intero della lista).");

  // Validazione macchina a stati lato server: rifiuta transizioni non valide
  // anche se il client fosse aggirato. Legge le timbrature odierne del
  // dipendente e determina l'ultimo evento.
  const oggi = await fetchTimbratureOggi();
  const eventiDip = oggi
    .filter((t) => t.dipendenteId === input.dipendenteId)
    .sort((a, b) => a.dataOra.localeCompare(b.dataOra));
  const last = eventiDip.length ? eventiDip[eventiDip.length - 1].evento : null;
  const allowed = nextAllowedSp(last);
  if (!allowed.includes(input.evento)) {
    logSp(
      "warn",
      "create.timbratura",
      `Transizione non ammessa per dip=${input.dipendenteId}: ${last ?? "nessuna"} → ${input.evento}`,
    );
    throw new Error(
      last === "uscita"
        ? "La giornata lavorativa è già stata chiusa. Per eventuali correzioni contatta il tuo responsabile."
        : "Timbratura non consentita in questo momento.",
    );
  }

  const dataOra = new Date().toISOString();
  const fields: Record<string, unknown> = {
    [lookupIdFieldName(dipendenteField)]: dipInt,
    [eventoField]: eventoToSharePoint(input.evento),
    [dataOraField]: dataOra,
  };
  if (F.Origine)
    fields[F.Origine] = (input.origine ?? "Web").replace(/^\w/, (c) => c.toUpperCase());
  if (F.Esito) fields[F.Esito] = input.esito ?? "Accettata";
  if (F.Posizione && input.posizione) fields[F.Posizione] = input.posizione;
  if (F.Note && input.note) fields[F.Note] = input.note;

  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  logSp("info", "create.timbratura", `Nuova timbratura #${created.id} (${input.evento})`, {
    durataMs: Date.now() - started,
  });
  return {
    id: String(created.id),
    dipendenteId: String(input.dipendenteId),
    evento: input.evento,
    dataOra,
    origine: input.origine ?? "web",
    esito: input.esito ?? "ok",
    posizione: input.posizione,
    note: input.note,
  };
}

// Macchina a stati identica a src/lib/presenze-logic.ts. Duplicata qui
// perché sharepoint.server.ts non può importare moduli client-safe che
// verrebbero comunque bundlati insieme; la logica è banale e stabile.
function nextAllowedSp(last: EventoTimbratura | null): EventoTimbratura[] {
  switch (last) {
    case null:
      return ["entrata"];
    case "entrata":
    case "fine-pausa":
      return ["inizio-pausa", "uscita"];
    case "inizio-pausa":
      return ["fine-pausa"];
    case "uscita":
      return [];
  }
}

export async function deleteTimbratura(id: string): Promise<void> {
  const cfg = await discoverSharePoint();
  const res = await gatewayFetch(`/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items/${id}`, {
    method: "DELETE",
  });
  if (!res.ok && res.status !== 204) {
    throw new SpHttpError(res.status, `DELETE timbratura ${id} → ${res.status}`, "delete");
  }
  logSp("info", "delete.timbratura", `Rimossa timbratura #${id}`);
}

// Annulla l'ULTIMA timbratura di oggi del dipendente, se registrata da meno di
// UNDO_TIMBRATURA_MINUTI ("ho premuto il tasto sbagliato"). La finestra è
// verificata QUI sul dato reale, non sull'orologio del client.
export async function annullaUltimaTimbratura(dipendenteId: string): Promise<SpTimbratura> {
  const oggi = await fetchTimbratureOggi();
  const mie = oggi.filter((t) => t.dipendenteId === dipendenteId);
  const ultima = mie[mie.length - 1];
  if (!ultima) throw new Error("Nessuna timbratura da annullare oggi.");
  const etaMs = Date.now() - new Date(ultima.dataOra).getTime();
  if (etaMs > UNDO_TIMBRATURA_MINUTI * 60_000)
    throw new Error(
      `L'annullamento è possibile solo entro ${UNDO_TIMBRATURA_MINUTI} minuti dalla timbratura. Rivolgiti all'operatore per la correzione.`,
    );
  await deleteTimbratura(ultima.id);
  logSp(
    "info",
    "undo.timbratura",
    `Annullata ${ultima.evento} di ${dipendenteId} (${ultima.dataOra})`,
  );
  return ultima;
}

// Timbrature di UN dipendente in UN giorno (vista correzione operatore).
export async function fetchTimbratureGiorno(
  dipendenteId: string,
  dataISO: string, // YYYY-MM-DD
): Promise<SpTimbratura[]> {
  const dayStart = new Date(`${dataISO}T00:00:00`).toISOString();
  const dayEnd = new Date(`${dataISO}T23:59:59.999`).toISOString();
  const tutte = await fetchTimbratureDaISO(dayStart);
  return tutte.filter((t) => t.dipendenteId === dipendenteId && t.dataOra <= dayEnd);
}

// Eliminazione di una timbratura da parte dell'OPERATORE (correzione errori:
// es. "uscita" premuta al posto di "inizio pausa"). Il flag Operatore è
// ri-verificato sul record SharePoint, come per gli inserimenti manuali.
export async function deleteTimbraturaOperatore(
  operatoreId: string,
  timbraturaId: string,
): Promise<void> {
  const cfg = await discoverSharePoint();
  await assertOperatore(cfg, operatoreId);
  await deleteTimbratura(timbraturaId);
  logSp(
    "info",
    "delete.timbratura.operatore",
    `Operatore ${operatoreId} ha rimosso la timbratura #${timbraturaId}`,
  );
}

// Inserimento MANUALE di una timbratura (operatore DR000). A differenza di
// createTimbratura NON applica la macchina a stati (le correzioni possono
// inserire eventi fuori ordine o nel passato) e marca Origine=Manuale, così le
// manuali sono filtrabili (tab supervisore DR005). Autorizzazione server-side:
// solo un dipendente con Operatore=true può inserirle.
export interface CreateTimbraturaManualeInput {
  operatoreId: string;
  dipendenteId: string;
  evento: EventoTimbratura;
  dataOra: string; // ISO datetime
  note?: string;
}

export interface CreateTurnoManualeInput {
  operatoreId: string;
  dipendenteId: string;
  entrata: string; // ISO datetime
  uscita: string; // ISO datetime
  inizioPausa?: string; // ISO datetime (opzionale)
  finePausa?: string; // ISO datetime (opzionale)
  note?: string;
}

// Verifica che l'operatore abbia la capability Operatore su SharePoint.
async function assertOperatore(cfg: SpDiscovered, operatoreId: string): Promise<void> {
  const DF = cfg.dipendentiFields;
  const opFields = await fetchDipendenteFields(cfg, operatoreId);
  const operatore = DF.Operatore ? Boolean(opFields[DF.Operatore]) : false;
  if (!operatore) {
    logSp("warn", "create.manuale", `Tentativo non autorizzato da id=${operatoreId}`);
    throw new Error("Non sei autorizzato a inserire timbrature manuali.");
  }
}

// Inserisce UNA timbratura manuale (Origine=Manuale). Nessuna macchina a stati.
async function insertManuale(
  cfg: SpDiscovered,
  dipInt: number,
  evento: EventoTimbratura,
  dataOraISO: string,
  note?: string,
): Promise<SpTimbratura> {
  const F = cfg.timbratureFields;
  const dipendenteField = requireField(F, "Dipendente", "Timbrature");
  const eventoField = requireField(F, "Evento", "Timbrature");
  const dataOraField = requireField(F, "DataOra", "Timbrature");
  const fields: Record<string, unknown> = {
    [lookupIdFieldName(dipendenteField)]: dipInt,
    [eventoField]: eventoToSharePoint(evento),
    [dataOraField]: dataOraISO,
  };
  if (F.Origine) fields[F.Origine] = "Manuale";
  if (F.Esito) fields[F.Esito] = "Accettata";
  if (F.Note && note) fields[F.Note] = note.trim();
  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  return {
    id: String(created.id),
    dipendenteId: String(dipInt),
    evento,
    dataOra: dataOraISO,
    origine: "Manuale",
    esito: "Accettata",
    note,
  };
}

export async function createTimbraturaManuale(
  input: CreateTimbraturaManualeInput,
): Promise<SpTimbratura> {
  const cfg = await discoverSharePoint();
  await assertOperatore(cfg, input.operatoreId);
  const dipInt = Number(input.dipendenteId);
  if (!Number.isFinite(dipInt)) throw new Error("dipendenteId non valido.");
  const evento = parseEvento(input.evento);
  if (!evento) throw new Error("Evento non valido.");
  const when = new Date(input.dataOra);
  if (Number.isNaN(when.getTime())) throw new Error("Data/ora non valida.");
  const t = await insertManuale(cfg, dipInt, evento, when.toISOString(), input.note);
  logSp(
    "info",
    "create.manuale",
    `Timbratura manuale #${t.id} (${evento}) dip=${input.dipendenteId} op=${input.operatoreId}`,
  );
  return t;
}

// Inserisce un TURNO INTERO in un colpo: entrata, [pausa], uscita (tutte manuali).
// Utile quando il dipendente non ha potuto timbrare l'intera giornata.
export async function createTurnoManuale(input: CreateTurnoManualeInput): Promise<SpTimbratura[]> {
  const cfg = await discoverSharePoint();
  await assertOperatore(cfg, input.operatoreId);
  const dipInt = Number(input.dipendenteId);
  if (!Number.isFinite(dipInt)) throw new Error("dipendenteId non valido.");

  const ms = (iso: string) => new Date(iso).getTime();
  const entrata = ms(input.entrata);
  const uscita = ms(input.uscita);
  if (Number.isNaN(entrata) || Number.isNaN(uscita)) throw new Error("Orari del turno non validi.");
  if (uscita <= entrata) throw new Error("L'uscita deve essere successiva all'entrata.");

  const eventi: { evento: EventoTimbratura; iso: string }[] = [
    { evento: "entrata", iso: input.entrata },
  ];
  if (input.inizioPausa || input.finePausa) {
    if (!input.inizioPausa || !input.finePausa)
      throw new Error("Per la pausa servono sia l'inizio sia la fine.");
    const ip = ms(input.inizioPausa);
    const fp = ms(input.finePausa);
    if (Number.isNaN(ip) || Number.isNaN(fp)) throw new Error("Orari della pausa non validi.");
    if (!(entrata < ip && ip < fp && fp < uscita))
      throw new Error("La pausa deve essere compresa tra entrata e uscita (inizio prima di fine).");
    eventi.push({ evento: "inizio-pausa", iso: input.inizioPausa });
    eventi.push({ evento: "fine-pausa", iso: input.finePausa });
  }
  eventi.push({ evento: "uscita", iso: input.uscita });
  eventi.sort((a, b) => ms(a.iso) - ms(b.iso));

  const out: SpTimbratura[] = [];
  for (const e of eventi) {
    out.push(await insertManuale(cfg, dipInt, e.evento, new Date(e.iso).toISOString(), input.note));
  }
  logSp(
    "info",
    "create.turno",
    `Turno manuale (${out.length} eventi) dip=${input.dipendenteId} op=${input.operatoreId}`,
  );
  return out;
}

// ---------------------------------------------------------------------------
// Richieste (ferie / permessi / straordinari) — modulo Sprint 2
// ---------------------------------------------------------------------------
export interface SpRichiesta {
  id: string;
  title: string;
  richiedenteId: string;
  codiceRichiedente: string;
  sedeRichiedente: string;
  tipo: string;
  modalita?: string;
  dataInizio: string; // ISO (data)
  dataFine: string; // ISO (data)
  oraInizio?: string;
  oraFine?: string;
  motivazione?: string;
  durataGiorni?: number;
  durataOre?: number;
  stato: string;
  dataInvio?: string;
  approvatoreId?: string;
  dataDecisione?: string;
  noteDecisione?: string;
  protocolloInps?: string;
  importo?: number;
  tipoAcquisto?: string;
  giustificativo?: string;
  annoCompetenza?: number;
  createdAt?: string;
}

export interface CreateRichiestaInput {
  richiedenteId: string;
  tipo: TipoRichiesta;
  dataInizio: string; // "YYYY-MM-DD"
  dataFine: string; // "YYYY-MM-DD"
  oraInizio?: string;
  oraFine?: string;
  motivazione?: string;
  modalita?: ModalitaStraordinario;
  protocolloInps?: string; // solo Malattia (facoltativo)
  importo?: number; // solo Rimborso spese
  tipoAcquisto?: TipoAcquisto; // solo Rimborso spese
  giustificativo?: string; // solo Rimborso spese (link/URL documento)
  submit?: boolean; // true → Inviata/Comunicata (con eventuale auto-approvazione)
}

export interface RichiesteFilter {
  richiedenteId?: string;
  stato?: string;
}

export interface DecideRichiestaInput {
  richiestaId: string;
  approvatoreId: string;
  decisione: DecisioneRichiesta;
  noteDecisione?: string;
}

function numOrUndef(v: unknown): number | undefined {
  if (v === undefined || v === null || v === "") return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function requireRichiesteList(cfg: SpDiscovered): string {
  if (!cfg.listRichieste)
    throw new Error(
      'Lista "Richieste" non trovata su SharePoint. Crearla sul sito DRPORTAL o verificarne il nome.',
    );
  return cfg.listRichieste;
}

// Legge i soli fields di un item Dipendenti per id (Codice/Sede/Autorizza).
async function fetchDipendenteFields(
  cfg: SpDiscovered,
  id: string,
): Promise<Record<string, unknown>> {
  const it = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items/${id}?expand=fields`,
    ),
  );
  return it.fields ?? {};
}

function mapRichiesta(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): SpRichiesta {
  const F = cfg.richiesteFields;
  const f = it.fields ?? {};
  const richLookup = F.Richiedente ? f[lookupIdFieldName(F.Richiedente)] : undefined;
  const appLookup = F.Approvatore ? f[lookupIdFieldName(F.Approvatore)] : undefined;
  return {
    id: String(it.id),
    title: String(f["Title"] ?? ""),
    richiedenteId: richLookup != null ? String(richLookup) : "",
    codiceRichiedente: F.CodiceRichiedente ? String(f[F.CodiceRichiedente] ?? "") : "",
    sedeRichiedente: F.SedeRichiedente ? String(f[F.SedeRichiedente] ?? "") : "",
    tipo: F.TipoRichiesta ? String(f[F.TipoRichiesta] ?? "") : "",
    modalita: F.Modalita ? (f[F.Modalita] as string | undefined) : undefined,
    dataInizio: F.DataInizio ? String(f[F.DataInizio] ?? "") : "",
    dataFine: F.DataFine ? String(f[F.DataFine] ?? "") : "",
    oraInizio: F.OraInizio ? (f[F.OraInizio] as string | undefined) : undefined,
    oraFine: F.OraFine ? (f[F.OraFine] as string | undefined) : undefined,
    motivazione: F.Motivazione ? (f[F.Motivazione] as string | undefined) : undefined,
    durataGiorni: F.DurataGiorni ? numOrUndef(f[F.DurataGiorni]) : undefined,
    durataOre: F.DurataOre ? numOrUndef(f[F.DurataOre]) : undefined,
    stato: F.Stato ? String(f[F.Stato] ?? "") : "",
    dataInvio: F.DataInvio ? (f[F.DataInvio] as string | undefined) : undefined,
    approvatoreId: appLookup != null ? String(appLookup) : undefined,
    dataDecisione: F.DataDecisione ? (f[F.DataDecisione] as string | undefined) : undefined,
    noteDecisione: F.NoteDecisione ? (f[F.NoteDecisione] as string | undefined) : undefined,
    protocolloInps: F.ProtocolloINPS ? (f[F.ProtocolloINPS] as string | undefined) : undefined,
    importo: F.Importo ? numOrUndef(f[F.Importo]) : undefined,
    tipoAcquisto: F.TipologiaAcquisto ? (f[F.TipologiaAcquisto] as string | undefined) : undefined,
    giustificativo: F.Giustificativo ? (f[F.Giustificativo] as string | undefined) : undefined,
    annoCompetenza: F.AnnoCompetenza ? numOrUndef(f[F.AnnoCompetenza]) : undefined,
    createdAt: (f["Created"] as string | undefined) ?? undefined,
  };
}

async function fetchRichiestaById(cfg: SpDiscovered, id: string): Promise<SpRichiesta> {
  const it = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${requireRichiesteList(cfg)}/items/${id}?expand=fields`,
    ),
  );
  return mapRichiesta(cfg, it);
}

export async function fetchRichieste(filter: RichiesteFilter = {}): Promise<SpRichiesta[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listRichieste) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listRichieste}/items?expand=fields&$top=999`,
    ),
  );
  let out = res.value.map((it) => mapRichiesta(cfg, it));
  if (filter.richiedenteId) out = out.filter((r) => r.richiedenteId === filter.richiedenteId);
  if (filter.stato) out = out.filter((r) => r.stato === filter.stato);
  out.sort((a, b) => (b.createdAt ?? "").localeCompare(a.createdAt ?? ""));
  return out;
}

// Vista supervisore: come fetchRichieste ma limitata alle sole richieste delle
// sedi di competenza dell'autorizzatore (identificato per Codice). DR005 copre
// le sedi storiche (Fiano Romano / San Giuliano); DR000 tutte le altre. L'admin
// non passa di qui — usa fetchRichieste e vede tutto.
export async function fetchRichiestePerSupervisore(
  supervisoreId: string,
  stato?: string,
): Promise<SpRichiesta[]> {
  const cfg = await discoverSharePoint();
  const DF = cfg.dipendentiFields;
  const dip = await fetchDipendenteFields(cfg, supervisoreId);
  const codice = DF.Codice ? String(dip[DF.Codice] ?? "").trim() : "";
  const all = await fetchRichieste({ stato });
  // DR005 è onnisciente: vede le richieste di tutte le sedi (come l'admin).
  if (isSupervisoreGlobale(codice)) return all;
  return all.filter((r) => supervisionaSede(codice, r.sedeRichiedente));
}

export async function createRichiesta(input: CreateRichiestaInput): Promise<SpRichiesta> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  const listId = requireRichiesteList(cfg);
  const F = cfg.richiesteFields;
  const richiedenteField = requireField(F, "Richiedente", "Richieste");
  const tipoField = requireField(F, "TipoRichiesta", "Richieste");
  const dataInizioField = requireField(F, "DataInizio", "Richieste");
  const dataFineField = requireField(F, "DataFine", "Richieste");
  const statoField = requireField(F, "Stato", "Richieste");

  const richiedenteNum = Number(input.richiedenteId);
  if (!Number.isFinite(richiedenteNum)) throw new Error("richiedenteId non valido.");

  // Re-validazione lato server (mai fidarsi del client).
  const v = validateRichiesta(input);
  if (!v.ok) throw new Error(v.errors.join(" "));

  // Denormalizzazione codice/sede + routing auto-approvazione: legge il record
  // del richiedente da SharePoint (autorevole).
  const DF = cfg.dipendentiFields;
  const dipFields = await fetchDipendenteFields(cfg, input.richiedenteId);
  const codice = DF.Codice ? String(dipFields[DF.Codice] ?? "").trim() : "";
  const sedeRaw = DF.Sede ? String(dipFields[DF.Sede] ?? "").trim() : "";
  const autorizza = DF.Autorizza ? Boolean(dipFields[DF.Autorizza]) : false;

  const submit = Boolean(input.submit);
  const anno = computeAnnoCompetenza(input.dataInizio);
  const approva = richiedeApprovazione(input.tipo);
  // Stato iniziale: Bozza se non inviata; all'invio → Inviata (tipi con
  // approvazione) oppure Comunicata (tipi senza approvazione, es. Malattia).
  const statoIniziale = submit ? (approva ? "Inviata" : "Comunicata") : "Bozza";

  const fields: Record<string, unknown> = {
    // Title placeholder: viene sovrascritto subito dopo con REQ-<anno>-<id>.
    Title: formatTitle(anno, "TMP"),
    [lookupIdFieldName(richiedenteField)]: richiedenteNum,
    [tipoField]: input.tipo,
    [dataInizioField]: `${input.dataInizio}T00:00:00Z`,
    [dataFineField]: `${input.dataFine}T00:00:00Z`,
    [statoField]: statoIniziale,
  };
  if (F.CodiceRichiedente && codice) fields[F.CodiceRichiedente] = codice;
  if (F.SedeRichiedente && sedeRaw) fields[F.SedeRichiedente] = sedeRaw;
  if (F.Motivazione && input.motivazione) fields[F.Motivazione] = input.motivazione.trim();
  if (F.AnnoCompetenza) fields[F.AnnoCompetenza] = anno;
  if (isRimborso(input.tipo)) {
    if (F.Importo && input.importo != null) fields[F.Importo] = input.importo;
    if (F.TipologiaAcquisto && input.tipoAcquisto) fields[F.TipologiaAcquisto] = input.tipoAcquisto;
    if (F.Giustificativo && input.giustificativo)
      fields[F.Giustificativo] = input.giustificativo.trim();
  } else if (misuraInGiorni(input.tipo)) {
    if (F.DurataGiorni)
      fields[F.DurataGiorni] = computeDurataGiorni(input.dataInizio, input.dataFine);
  } else {
    if (F.OraInizio && input.oraInizio) fields[F.OraInizio] = input.oraInizio;
    if (F.OraFine && input.oraFine) fields[F.OraFine] = input.oraFine;
    if (F.DurataOre && input.oraInizio && input.oraFine)
      fields[F.DurataOre] = computeDurataOre(input.oraInizio, input.oraFine);
  }
  if (input.tipo === "Straordinario" && F.Modalita && input.modalita)
    fields[F.Modalita] = input.modalita;
  if (input.tipo === "Malattia" && F.ProtocolloINPS && input.protocolloInps)
    fields[F.ProtocolloINPS] = input.protocolloInps.trim();
  if (submit && F.DataInvio) fields[F.DataInvio] = new Date().toISOString();

  // Auto-approvazione: SOLO tipi con approvazione, richiedente autorizzato che
  // invia una propria richiesta (oggi Francesco). Traccia approvatore/data/nota
  // per l'audit. I tipi senza approvazione (Malattia) non passano di qui.
  const auto = submit && approva && isAutoApprovazione(input.richiedenteId, autorizza);
  if (auto) {
    fields[statoField] = "Approvata";
    if (F.Approvatore) fields[lookupIdFieldName(F.Approvatore)] = richiedenteNum;
    if (F.DataDecisione) fields[F.DataDecisione] = new Date().toISOString();
    if (F.NoteDecisione) fields[F.NoteDecisione] = NOTA_AUTO_APPROVAZIONE;
  }

  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );

  // PATCH del Title leggibile REQ-<anno>-<idNativo>.
  const title = formatTitle(anno, created.id);
  try {
    await gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${created.id}/fields`, {
      method: "PATCH",
      body: JSON.stringify({ Title: title }),
    });
  } catch (err) {
    logSp(
      "warn",
      "create.richiesta",
      `Title non aggiornato per #${created.id}: ${err instanceof Error ? err.message : String(err)}`,
    );
  }

  logSp(
    "info",
    "create.richiesta",
    `Richiesta ${title} (${input.tipo}, ${String(fields[statoField])})`,
    { durataMs: Date.now() - started },
  );
  return fetchRichiestaById(cfg, created.id);
}

export async function decideRichiesta(input: DecideRichiestaInput): Promise<SpRichiesta> {
  const cfg = await discoverSharePoint();
  requireRichiesteList(cfg);
  const F = cfg.richiesteFields;
  const statoField = requireField(F, "Stato", "Richieste");

  // Autorizzazione SERVER-SIDE: l'approvatore deve avere Autorizza=true su SP.
  // Non nascondere il bottone nella UI non basta; qui è dove conta davvero.
  const DF = cfg.dipendentiFields;
  const dipFields = await fetchDipendenteFields(cfg, input.approvatoreId);
  const autorizza = DF.Autorizza ? Boolean(dipFields[DF.Autorizza]) : false;
  if (!autorizza) {
    logSp("warn", "decide.richiesta", `Tentativo non autorizzato da id=${input.approvatoreId}`);
    throw new Error("Non sei autorizzato ad approvare o respingere richieste.");
  }

  const vd = validateDecisione(input.decisione, input.noteDecisione);
  if (!vd.ok) throw new Error(vd.errors.join(" "));

  // Re-check dello stato per evitare doppia decisione concorrente (TOCTOU).
  const current = await fetchRichiestaById(cfg, input.richiestaId);
  if (!canDecide(parseStato(current.stato))) {
    throw new Error(`Richiesta non decidibile nello stato "${current.stato}".`);
  }

  // Competenza per sede: l'autorizzatore può decidere SOLO sulle richieste delle
  // sedi che supervisiona (DR005 = sedi storiche, DR000 = tutte le altre).
  // L'amministratore di sistema può decidere ovunque.
  const ruoloApprover = DF.Ruolo ? normalizeRuolo(String(dipFields[DF.Ruolo] ?? "")) : "dipendente";
  if (ruoloApprover !== "amministratore_sistema") {
    const codiceApprover = DF.Codice ? String(dipFields[DF.Codice] ?? "").trim() : "";
    if (!supervisionaSede(codiceApprover, current.sedeRichiedente)) {
      logSp(
        "warn",
        "decide.richiesta",
        `Sede non di competenza per id=${input.approvatoreId} (sede="${current.sedeRichiedente}")`,
      );
      throw new Error("Questa richiesta è di competenza di un altro supervisore.");
    }
  }

  const approvatoreNum = Number(input.approvatoreId);
  const fields: Record<string, unknown> = { [statoField]: input.decisione };
  if (F.Approvatore && Number.isFinite(approvatoreNum))
    fields[lookupIdFieldName(F.Approvatore)] = approvatoreNum;
  if (F.DataDecisione) fields[F.DataDecisione] = new Date().toISOString();
  if (F.NoteDecisione && input.noteDecisione) fields[F.NoteDecisione] = input.noteDecisione.trim();

  await withDiscoveryRetry(() =>
    gatewayJson(
      `/sites/${cfg.siteId}/lists/${cfg.listRichieste}/items/${input.richiestaId}/fields`,
      {
        method: "PATCH",
        body: JSON.stringify(fields),
      },
    ),
  );
  logSp("info", "decide.richiesta", `Richiesta #${input.richiestaId} → ${input.decisione}`);
  return fetchRichiestaById(cfg, input.richiestaId);
}

export async function cancelRichiesta(inp: {
  richiestaId: string;
  richiedenteId: string;
}): Promise<SpRichiesta> {
  const cfg = await discoverSharePoint();
  requireRichiesteList(cfg);
  const F = cfg.richiesteFields;
  const statoField = requireField(F, "Stato", "Richieste");
  const current = await fetchRichiestaById(cfg, inp.richiestaId);
  if (current.richiedenteId !== inp.richiedenteId)
    throw new Error("Non puoi annullare una richiesta non tua.");
  if (!canCancel(parseStato(current.stato)))
    throw new Error(`Richiesta non annullabile nello stato "${current.stato}".`);
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${cfg.listRichieste}/items/${inp.richiestaId}/fields`, {
      method: "PATCH",
      body: JSON.stringify({ [statoField]: "Annullata" }),
    }),
  );
  logSp("info", "cancel.richiesta", `Richiesta #${inp.richiestaId} annullata`);
  return fetchRichiestaById(cfg, inp.richiestaId);
}

export async function deleteRichiesta(id: string): Promise<void> {
  const cfg = await discoverSharePoint();
  const listId = requireRichiesteList(cfg);
  const res = await gatewayFetch(`/sites/${cfg.siteId}/lists/${listId}/items/${id}`, {
    method: "DELETE",
  });
  if (!res.ok && res.status !== 204) {
    throw new SpHttpError(res.status, `DELETE richiesta ${id} → ${res.status}`, "delete");
  }
  logSp("info", "delete.richiesta", `Rimossa richiesta #${id}`);
}

// ---------------------------------------------------------------------------
// Upload giustificativi di spesa (libreria documenti del sito)
// ---------------------------------------------------------------------------
// Microsoft Graph non espone gli allegati delle liste tramite il gateway, per
// cui i giustificativi vengono caricati nel drive del sito (verificato dal
// self-test). Il file finisce in /Rimborsi/<anno>/<timestamp>-<nome> e viene
// restituito il webUrl da salvare nel campo "Giustificativo" della richiesta.
function base64ToBytes(b64: string): Uint8Array {
  // Accetta sia base64 puro sia data URL ("data:...;base64,XXXX").
  const clean = b64.includes(",") ? b64.slice(b64.indexOf(",") + 1) : b64;
  const bin = atob(clean);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

export interface UploadGiustificativoResult {
  webUrl: string;
  fileName: string;
}

// Upload generico di un file sulla libreria documenti del sito, in una
// sottocartella (`subfolder`). Ritorna il webUrl da salvare come riferimento.
// Limite 8 MB. Usato da giustificativi, documenti dipendente e allegati.
export async function uploadFileToLibrary(
  subfolder: string,
  filename: string,
  contentBase64: string,
): Promise<UploadGiustificativoResult> {
  const cfg = await discoverSharePoint();
  const bytes = base64ToBytes(contentBase64);
  if (bytes.length === 0) throw new Error("Il file caricato è vuoto.");
  if (bytes.length > 8 * 1024 * 1024) {
    throw new Error("File troppo grande: il limite è 8 MB.");
  }
  // Nome file sicuro: solo caratteri innocui, coda limitata a 80 char.
  const safe = (filename || "documento").replace(/[^A-Za-z0-9._-]/g, "_").slice(-80) || "documento";
  const folder = (subfolder || "Documenti").replace(/[^A-Za-z0-9._/-]/g, "_");
  const anno = new Date().getFullYear();
  const path = `${folder}/${anno}/${Date.now()}-${safe}`;
  const created = await withDiscoveryRetry(() =>
    gatewayJson<{ webUrl?: string; name?: string }>(
      `/sites/${cfg.siteId}/drive/root:/${path}:/content`,
      {
        method: "PUT",
        headers: { "Content-Type": "application/octet-stream" },
        // runtime (Workers/Node) accetta Uint8Array come body; il cast evita
        // l'incompatibilità del tipo BodyInit su lib DOM recenti.
        body: bytes as unknown as BodyInit,
      },
    ),
  );
  logSp("info", "upload.file", `Caricato ${path} (${bytes.length} byte)`);
  return { webUrl: created.webUrl ?? "", fileName: created.name ?? safe };
}

export async function uploadGiustificativo(
  filename: string,
  contentBase64: string,
): Promise<UploadGiustificativoResult> {
  return uploadFileToLibrary("Rimborsi", filename, contentBase64);
}

// ---------------------------------------------------------------------------
// Coda email (outbox) — il portale accoda, Power Automate invia
// ---------------------------------------------------------------------------
// Il connettore non ha Mail.Send (probe → 403): l'invio è delegato a un flusso
// Power Automate che osserva la lista CodaEmail (Stato="Da inviare"), spedisce
// e aggiorna lo Stato. Il portale si limita ad accodare.

export function parseEmails(raw: string): string[] {
  return (raw ?? "")
    .split(/[,;]/)
    .map((s) => s.trim())
    .filter((s) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s));
}

// Mittente di default: casella della Segreteria. Chi accoda può indicarne un
// altro (es. l'email di chi pubblica una comunicazione); il flusso Power
// Automate usa il campo Mittente come "Da (Invia come)".
export const EMAIL_MITTENTE_DEFAULT = "segreteria@drlogistica.it";

// Codice dipendente (es. DR005) dal record SharePoint ("" se assente).
// Verifica AUTOREVOLE per i moduli riservati al direttore: non ci si fida del
// codice nella sessione client, si rilegge il record.
export async function getCodiceDipendente(dipendenteId: string): Promise<string> {
  try {
    const cfg = await discoverSharePoint();
    const DF = cfg.dipendentiFields;
    if (!DF.Codice) return "";
    const f = await fetchDipendenteFields(cfg, dipendenteId);
    return normalizeCodice(f[DF.Codice]);
  } catch {
    return "";
  }
}

// Email di un dipendente dal suo record SharePoint ("" se assente).
export async function getEmailDipendente(dipendenteId: string): Promise<string> {
  try {
    const cfg = await discoverSharePoint();
    const DF = cfg.dipendentiFields;
    if (!DF.Email) return "";
    const f = await fetchDipendenteFields(cfg, dipendenteId);
    return String(f[DF.Email] ?? "").trim();
  } catch {
    return "";
  }
}

export async function enqueueEmail(msg: {
  destinatari: string[];
  oggetto: string;
  corpo: string;
  allegato?: string;
  mittente?: string;
}): Promise<boolean> {
  const cfg = await discoverSharePoint();
  if (!cfg.listCodaEmail || msg.destinatari.length === 0) return false;
  const F = cfg.codaEmailFields;
  const fields: Record<string, unknown> = { Title: msg.oggetto.slice(0, 200) };
  if (F.Destinatari) fields[F.Destinatari] = msg.destinatari.join("; ");
  if (F.Oggetto) fields[F.Oggetto] = msg.oggetto;
  if (F.Corpo) fields[F.Corpo] = msg.corpo;
  if (F.Allegato && msg.allegato) fields[F.Allegato] = msg.allegato;
  if (F.Stato) fields[F.Stato] = "Da inviare";
  if (F.Mittente) fields[F.Mittente] = msg.mittente?.trim() || EMAIL_MITTENTE_DEFAULT;
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${cfg.listCodaEmail}/items`, {
      method: "POST",
      body: JSON.stringify({ fields }),
    }),
  );
  logSp(
    "info",
    "email.enqueue",
    `Email in coda: "${msg.oggetto}" → ${msg.destinatari.length} destinatari`,
  );
  return true;
}

// ---------------------------------------------------------------------------
// Modulo Documenti dipendente (Sprint 4)
// ---------------------------------------------------------------------------
function requireDocumentiList(cfg: SpDiscovered): string {
  if (!cfg.listDocumenti)
    throw new Error('Lista "Documenti" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  return cfg.listDocumenti;
}

export type DocumentoCategoria = "Contratto" | "Busta paga" | "DPI" | "Certificato corso" | "Altro";
export type DocumentoAmbito = "Personale" | "Generale";

export interface SpDocumento {
  id: string;
  categoria: string;
  titolo: string;
  ambito: string;
  destinatarioId: string;
  codiceDestinatario: string;
  sedeDestinatario: string;
  file: string;
  nomeFile: string;
  dataDocumento: string;
  caricatoDa: string;
  createdAt?: string;
}
export interface CreateDocumentoInput {
  categoria: DocumentoCategoria;
  titolo: string;
  ambito: DocumentoAmbito;
  destinatarioId?: string; // per Ambito=Personale
  sedeDestinatario?: string; // per Ambito=Generale ("Tutte" o nome sede)
  file: string; // webUrl del file caricato
  nomeFile?: string;
  caricatoDa: string; // codice di chi carica
}

function mapDocumento(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): SpDocumento {
  const F = cfg.documentiFields;
  const f = it.fields ?? {};
  return {
    id: String(it.id),
    categoria: F.Categoria ? String(f[F.Categoria] ?? "") : "",
    titolo: String((F.Titolo ? f[F.Titolo] : undefined) ?? f["Title"] ?? ""),
    ambito: F.Ambito ? String(f[F.Ambito] ?? "") : "",
    destinatarioId: F.DestinatarioId ? String(f[F.DestinatarioId] ?? "") : "",
    codiceDestinatario: F.CodiceDestinatario ? String(f[F.CodiceDestinatario] ?? "") : "",
    sedeDestinatario: F.SedeDestinatario ? String(f[F.SedeDestinatario] ?? "") : "",
    file: F.File ? String(f[F.File] ?? "") : "",
    nomeFile: F.NomeFile ? String(f[F.NomeFile] ?? "") : "",
    dataDocumento: F.DataDocumento ? String(f[F.DataDocumento] ?? "") : "",
    caricatoDa: F.CaricatoDa ? String(f[F.CaricatoDa] ?? "") : "",
    createdAt: (f["Created"] as string | undefined) ?? undefined,
  };
}

async function fetchAllDocumenti(cfg: SpDiscovered): Promise<SpDocumento[]> {
  if (!cfg.listDocumenti) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDocumenti}/items?expand=fields&$top=999`,
    ),
  );
  return res.value
    .map((it) => mapDocumento(cfg, it))
    .sort((a, b) =>
      (b.dataDocumento || b.createdAt || "").localeCompare(a.dataDocumento || a.createdAt || ""),
    );
}

export async function fetchDocumentiAll(): Promise<SpDocumento[]> {
  return fetchAllDocumenti(await discoverSharePoint());
}

export async function fetchDocumentiForUser(userId: string, sede: string): Promise<SpDocumento[]> {
  const all = await fetchAllDocumenti(await discoverSharePoint());
  const sedeLow = (sede || "").trim().toLowerCase();
  return all.filter((d) => {
    if (d.ambito === "Personale") return d.destinatarioId === userId;
    // Generale: destinato a tutte le sedi o alla sede dell'utente.
    const s = (d.sedeDestinatario || "").trim().toLowerCase();
    return s === "" || s === "tutte" || s === sedeLow;
  });
}

export async function createDocumento(input: CreateDocumentoInput): Promise<SpDocumento> {
  const cfg = await discoverSharePoint();
  const listId = requireDocumentiList(cfg);
  const F = cfg.documentiFields;
  const fields: Record<string, unknown> = {};
  if (F.Categoria) fields[F.Categoria] = input.categoria;
  if (F.Titolo) fields[F.Titolo] = input.titolo;
  if (F.Ambito) fields[F.Ambito] = input.ambito;
  if (F.File) fields[F.File] = input.file;
  if (F.NomeFile && input.nomeFile) fields[F.NomeFile] = input.nomeFile;
  if (F.DataDocumento) fields[F.DataDocumento] = new Date().toISOString();
  if (F.CaricatoDa) fields[F.CaricatoDa] = input.caricatoDa;
  fields["Title"] = input.titolo || input.categoria || "Documento";
  let emailDestinatario = "";
  if (input.ambito === "Personale" && input.destinatarioId) {
    if (F.DestinatarioId) fields[F.DestinatarioId] = input.destinatarioId;
    try {
      const DF = cfg.dipendentiFields;
      const dip = await fetchDipendenteFields(cfg, input.destinatarioId);
      if (F.CodiceDestinatario && DF.Codice)
        fields[F.CodiceDestinatario] = String(dip[DF.Codice] ?? "");
      if (DF.Email) emailDestinatario = String(dip[DF.Email] ?? "").trim();
    } catch {
      /* denormalizzazione best-effort */
    }
  } else if (input.ambito === "Generale") {
    if (F.SedeDestinatario) fields[F.SedeDestinatario] = input.sedeDestinatario || "Tutte";
  }
  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  logSp("info", "create.documento", `Documento "${input.titolo}" (${input.categoria})`);
  // Documento personale → email al destinatario via coda (best-effort).
  // Le buste paga usano il formato richiesto dalla Segreteria; gli altri
  // documenti un testo generico.
  if (emailDestinatario && parseEmails(emailDestinatario).length) {
    const isBustaPaga = input.categoria === "Busta paga";
    // "Busta paga Giugno 2026" → periodo = titolo senza il prefisso categoria.
    const periodo = input.titolo.replace(/^busta paga\s*/i, "").trim();
    await enqueueEmail({
      destinatari: parseEmails(emailDestinatario),
      oggetto: isBustaPaga
        ? `Busta Paga ${periodo}`
        : `[DR Portal] Nuovo documento: ${input.titolo}`,
      corpo: isBustaPaga
        ? `Buongiorno,\nin allegato il link alla busta paga relativa a ${periodo}:\n${input.file}\n\nLa ritrovi anche nella tua area personale: https://portal.drlogistica.it/documenti\n\nSaluti\nSegreteria DR`
        : `È disponibile un nuovo documento nella tua area personale del portale DR Logistica.\n\nCategoria: ${input.categoria}\nTitolo: ${input.titolo}\n\nConsultalo su https://portal.drlogistica.it/documenti`,
      allegato: input.file,
    }).catch(() => {});
  }
  return mapDocumento(cfg, { id: created.id, fields });
}

// ---------------------------------------------------------------------------
// Modulo Comunicazioni interne + Prese visione (Sprint 4)
// ---------------------------------------------------------------------------
function requireComunicazioniList(cfg: SpDiscovered): string {
  if (!cfg.listComunicazioni)
    throw new Error('Lista "Comunicazioni" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  return cfg.listComunicazioni;
}

export type ComunicazioneTipo = "Riunione" | "Comunicazione";

export interface SpComunicazione {
  id: string;
  titolo: string;
  testo: string;
  tipo: string;
  sede: string;
  dataComunicazione: string;
  autore: string;
  allegato: string;
  richiedePresaVisione: boolean;
  createdAt?: string;
}
export interface CreateComunicazioneInput {
  titolo: string;
  testo: string;
  tipo: ComunicazioneTipo;
  sede: string; // "Tutte" o nome sede
  autore: string; // codice
  allegato?: string;
  richiedePresaVisione: boolean;
}
export interface SpPresaVisione {
  id: string;
  comunicazioneId: string;
  dipendenteId: string;
  codiceDipendente: string;
  dataLettura: string;
}

function mapComunicazione(
  cfg: SpDiscovered,
  it: GraphListItem<Record<string, unknown>>,
): SpComunicazione {
  const F = cfg.comunicazioniFields;
  const f = it.fields ?? {};
  return {
    id: String(it.id),
    titolo: String((F.Titolo ? f[F.Titolo] : undefined) ?? f["Title"] ?? ""),
    testo: F.Testo ? String(f[F.Testo] ?? "") : "",
    tipo: F.Tipo ? String(f[F.Tipo] ?? "") : "",
    sede: F.Sede ? String(f[F.Sede] ?? "") : "",
    dataComunicazione: F.DataComunicazione ? String(f[F.DataComunicazione] ?? "") : "",
    autore: F.Autore ? String(f[F.Autore] ?? "") : "",
    allegato: F.Allegato ? String(f[F.Allegato] ?? "") : "",
    richiedePresaVisione: parseSpBool(
      F.RichiedePresaVisione ? f[F.RichiedePresaVisione] : undefined,
      false,
    ),
    createdAt: (f["Created"] as string | undefined) ?? undefined,
  };
}

async function fetchAllComunicazioni(cfg: SpDiscovered): Promise<SpComunicazione[]> {
  if (!cfg.listComunicazioni) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listComunicazioni}/items?expand=fields&$top=999`,
    ),
  );
  return res.value
    .map((it) => mapComunicazione(cfg, it))
    .sort((a, b) =>
      (b.dataComunicazione || b.createdAt || "").localeCompare(
        a.dataComunicazione || a.createdAt || "",
      ),
    );
}

export async function fetchComunicazioniAll(): Promise<SpComunicazione[]> {
  return fetchAllComunicazioni(await discoverSharePoint());
}

export async function fetchComunicazioniForUser(sede: string): Promise<SpComunicazione[]> {
  const all = await fetchAllComunicazioni(await discoverSharePoint());
  const sedeLow = (sede || "").trim().toLowerCase();
  return all.filter((c) => {
    const s = (c.sede || "").trim().toLowerCase();
    return s === "" || s === "tutte" || s === sedeLow;
  });
}

export async function createComunicazione(
  input: CreateComunicazioneInput,
): Promise<SpComunicazione> {
  const cfg = await discoverSharePoint();
  const listId = requireComunicazioniList(cfg);
  const F = cfg.comunicazioniFields;
  const fields: Record<string, unknown> = {};
  fields["Title"] = input.titolo || "Comunicazione";
  if (F.Titolo) fields[F.Titolo] = input.titolo;
  if (F.Testo) fields[F.Testo] = input.testo;
  if (F.Tipo) fields[F.Tipo] = input.tipo;
  if (F.Sede) fields[F.Sede] = input.sede || "Tutte";
  if (F.DataComunicazione) fields[F.DataComunicazione] = new Date().toISOString();
  if (F.Autore) fields[F.Autore] = input.autore;
  if (F.Allegato && input.allegato) fields[F.Allegato] = input.allegato;
  if (F.RichiedePresaVisione) fields[F.RichiedePresaVisione] = input.richiedePresaVisione;
  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  logSp("info", "create.comunicazione", `Comunicazione "${input.titolo}" (${input.tipo})`);
  return mapComunicazione(cfg, { id: created.id, fields });
}

function mapPresaVisione(
  cfg: SpDiscovered,
  it: GraphListItem<Record<string, unknown>>,
): SpPresaVisione {
  const F = cfg.preseVisioneFields;
  const f = it.fields ?? {};
  return {
    id: String(it.id),
    comunicazioneId: F.ComunicazioneId ? String(f[F.ComunicazioneId] ?? "") : "",
    dipendenteId: F.DipendenteId ? String(f[F.DipendenteId] ?? "") : "",
    codiceDipendente: F.CodiceDipendente ? String(f[F.CodiceDipendente] ?? "") : "",
    dataLettura: F.DataLettura ? String(f[F.DataLettura] ?? "") : "",
  };
}

async function fetchAllPreseVisione(cfg: SpDiscovered): Promise<SpPresaVisione[]> {
  if (!cfg.listPreseVisione) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listPreseVisione}/items?expand=fields&$top=999`,
    ),
  );
  return res.value.map((it) => mapPresaVisione(cfg, it));
}

// Prese visione di una comunicazione (chi l'ha letta).
export async function fetchPreseVisione(comunicazioneId: string): Promise<SpPresaVisione[]> {
  const all = await fetchAllPreseVisione(await discoverSharePoint());
  return all.filter((p) => p.comunicazioneId === comunicazioneId);
}

// Id delle comunicazioni già confermate da un dipendente.
export async function fetchPreseVisioneForUser(dipendenteId: string): Promise<string[]> {
  const all = await fetchAllPreseVisione(await discoverSharePoint());
  return all.filter((p) => p.dipendenteId === dipendenteId).map((p) => p.comunicazioneId);
}

// Registra la presa visione (idempotente: non duplica se già presente).
export async function markPresaVisione(
  comunicazioneId: string,
  dipendenteId: string,
  codiceDipendente: string,
): Promise<void> {
  const cfg = await discoverSharePoint();
  if (!cfg.listPreseVisione)
    throw new Error('Lista "PreseVisione" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  const esistenti = await fetchAllPreseVisione(cfg);
  if (
    esistenti.some((p) => p.comunicazioneId === comunicazioneId && p.dipendenteId === dipendenteId)
  )
    return;
  const F = cfg.preseVisioneFields;
  const fields: Record<string, unknown> = { Title: `PV-${comunicazioneId}-${dipendenteId}` };
  if (F.ComunicazioneId) fields[F.ComunicazioneId] = comunicazioneId;
  if (F.DipendenteId) fields[F.DipendenteId] = dipendenteId;
  if (F.CodiceDipendente) fields[F.CodiceDipendente] = codiceDipendente;
  if (F.DataLettura) fields[F.DataLettura] = new Date().toISOString();
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${cfg.listPreseVisione}/items`, {
      method: "POST",
      body: JSON.stringify({ fields }),
    }),
  );
  logSp("info", "presa.visione", `Comunicazione #${comunicazioneId} letta da #${dipendenteId}`);
}

// ---------------------------------------------------------------------------
// Voci di spesa (macro → dettaglio) — lista gestita dall'azienda
// ---------------------------------------------------------------------------
export interface SpVoce {
  macro: string;
  dettaglio: string;
}

// Voci per ambito ("Rimborso" | "Acquisto"), raggruppate macro → dettagli.
export async function fetchVoci(ambito: string): Promise<SpVoce[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listVoci) return [];
  const F = cfg.vociFields;
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listVoci}/items?expand=fields&$top=999`,
    ),
  );
  const amb = ambito.trim().toLowerCase();
  return res.value
    .map((it) => {
      const f = it.fields ?? {};
      return {
        ambito: F.Ambito ? String(f[F.Ambito] ?? "").trim() : "",
        macro: F.Macro ? String(f[F.Macro] ?? "").trim() : "",
        dettaglio: F.Dettaglio ? String(f[F.Dettaglio] ?? "").trim() : "",
      };
    })
    .filter((v) => v.macro && v.ambito.toLowerCase() === amb)
    .map(({ macro, dettaglio }) => ({ macro, dettaglio }))
    .sort((a, b) => a.macro.localeCompare(b.macro) || a.dettaglio.localeCompare(b.dettaglio));
}

// ---------------------------------------------------------------------------
// Procurement — richieste di acquisto (solo sedi storiche, approva DR005)
// ---------------------------------------------------------------------------
function requireAcquistiList(cfg: SpDiscovered): string {
  if (!cfg.listAcquisti)
    throw new Error(
      'Lista "RichiesteAcquisto" non trovata su SharePoint. Crearla sul sito DRPORTAL.',
    );
  return cfg.listAcquisti;
}

export interface SpAcquisto {
  id: string;
  title: string;
  richiedenteId: string;
  codiceRichiedente: string;
  sedeRichiedente: string;
  macro: string;
  dettaglio: string;
  descrizione: string;
  importo?: number;
  stato: string;
  dataRichiesta: string;
  approvatoreId?: string;
  dataDecisione?: string;
  noteDecisione?: string;
  createdAt?: string;
}
export interface CreateAcquistoInput {
  macro: string;
  dettaglio: string;
  descrizione: string;
  importo?: number;
}

function mapAcquisto(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): SpAcquisto {
  const F = cfg.acquistiFields;
  const f = it.fields ?? {};
  const richLookup = F.Richiedente ? f[lookupIdFieldName(F.Richiedente)] : undefined;
  const appLookup = F.Approvatore ? f[lookupIdFieldName(F.Approvatore)] : undefined;
  return {
    id: String(it.id),
    title: String(f["Title"] ?? ""),
    richiedenteId: richLookup != null ? String(richLookup) : "",
    codiceRichiedente: F.CodiceRichiedente ? String(f[F.CodiceRichiedente] ?? "") : "",
    sedeRichiedente: F.SedeRichiedente ? String(f[F.SedeRichiedente] ?? "") : "",
    macro: F.Macro ? String(f[F.Macro] ?? "") : "",
    dettaglio: F.Dettaglio ? String(f[F.Dettaglio] ?? "") : "",
    descrizione: F.Descrizione ? String(f[F.Descrizione] ?? "") : "",
    importo: F.Importo ? numOrUndef(f[F.Importo]) : undefined,
    stato: F.Stato ? String(f[F.Stato] ?? "") : "",
    dataRichiesta: F.DataRichiesta ? String(f[F.DataRichiesta] ?? "") : "",
    approvatoreId: appLookup != null ? String(appLookup) : undefined,
    dataDecisione: F.DataDecisione ? (f[F.DataDecisione] as string | undefined) : undefined,
    noteDecisione: F.NoteDecisione ? (f[F.NoteDecisione] as string | undefined) : undefined,
    createdAt: (f["Created"] as string | undefined) ?? undefined,
  };
}

export async function fetchAcquisti(filter: {
  richiedenteId?: string;
  stato?: string;
}): Promise<SpAcquisto[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listAcquisti) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listAcquisti}/items?expand=fields&$top=999`,
    ),
  );
  let out = res.value.map((it) => mapAcquisto(cfg, it));
  if (filter.richiedenteId) out = out.filter((r) => r.richiedenteId === filter.richiedenteId);
  if (filter.stato) out = out.filter((r) => r.stato === filter.stato);
  out.sort((a, b) => (b.createdAt ?? "").localeCompare(a.createdAt ?? ""));
  return out;
}

export async function createAcquisto(
  richiedenteId: string,
  input: CreateAcquistoInput,
): Promise<SpAcquisto> {
  const cfg = await discoverSharePoint();
  const listId = requireAcquistiList(cfg);
  const F = cfg.acquistiFields;
  if (!input.macro.trim()) throw new Error("Seleziona la voce di acquisto.");
  if (!input.descrizione.trim()) throw new Error("Descrivi cosa serve acquistare.");

  // Denormalizzazione codice/sede dal record autorevole del richiedente e
  // enforcement: il Procurement è attivo SOLO per le sedi storiche.
  const DF = cfg.dipendentiFields;
  const dipFields = await fetchDipendenteFields(cfg, richiedenteId);
  const codice = DF.Codice ? String(dipFields[DF.Codice] ?? "").trim() : "";
  const sedeRaw = DF.Sede ? String(dipFields[DF.Sede] ?? "").trim() : "";
  if (!isSedeStorica(sedeRaw))
    throw new Error("Le richieste di acquisto sono attive solo per le sedi storiche.");

  const richiedenteNum = Number(richiedenteId);
  const fields: Record<string, unknown> = { Title: "ACQ-TMP" };
  if (F.Richiedente && Number.isFinite(richiedenteNum))
    fields[lookupIdFieldName(F.Richiedente)] = richiedenteNum;
  if (F.CodiceRichiedente && codice) fields[F.CodiceRichiedente] = codice;
  if (F.SedeRichiedente && sedeRaw) fields[F.SedeRichiedente] = sedeRaw;
  if (F.Macro) fields[F.Macro] = input.macro.trim();
  if (F.Dettaglio && input.dettaglio) fields[F.Dettaglio] = input.dettaglio.trim();
  if (F.Descrizione) fields[F.Descrizione] = input.descrizione.trim();
  if (F.Importo && input.importo != null) fields[F.Importo] = input.importo;
  if (F.Stato) fields[F.Stato] = "Inviata";
  if (F.DataRichiesta) fields[F.DataRichiesta] = new Date().toISOString();

  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  const anno = new Date().getFullYear();
  try {
    await gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${created.id}/fields`, {
      method: "PATCH",
      body: JSON.stringify({ Title: `ACQ-${anno}-${created.id}` }),
    });
  } catch {
    /* solo titolo, non bloccante */
  }
  logSp("info", "create.acquisto", `Acquisto ACQ-${anno}-${created.id} (${input.macro})`);
  return mapAcquisto(cfg, { id: created.id, fields });
}

// Approvazione/rifiuto: SOLO il supervisore globale DR005 (o l'admin).
export async function decideAcquisto(input: {
  acquistoId: string;
  approvatoreId: string;
  decisione: "Approvata" | "Respinta";
  noteDecisione?: string;
}): Promise<SpAcquisto> {
  const cfg = await discoverSharePoint();
  const listId = requireAcquistiList(cfg);
  const F = cfg.acquistiFields;

  const DF = cfg.dipendentiFields;
  const dipFields = await fetchDipendenteFields(cfg, input.approvatoreId);
  const autorizza = DF.Autorizza ? Boolean(dipFields[DF.Autorizza]) : false;
  const codice = DF.Codice ? String(dipFields[DF.Codice] ?? "").trim() : "";
  const ruolo = DF.Ruolo ? normalizeRuolo(String(dipFields[DF.Ruolo] ?? "")) : "dipendente";
  const isAdminRole = ruolo === "amministratore_sistema";
  if (!isAdminRole && !(autorizza && isSupervisoreGlobale(codice))) {
    logSp("warn", "decide.acquisto", `Tentativo non autorizzato da id=${input.approvatoreId}`);
    throw new Error("Le richieste di acquisto possono essere approvate solo da DR005.");
  }
  if (input.decisione === "Respinta" && !input.noteDecisione?.trim())
    throw new Error("Il motivo del rifiuto è obbligatorio.");

  const current = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items/${input.acquistoId}?expand=fields`,
    ),
  ).then((it) => mapAcquisto(cfg, it));
  if (current.stato !== "Inviata")
    throw new Error(`Richiesta non decidibile nello stato "${current.stato}".`);

  const approvatoreNum = Number(input.approvatoreId);
  const fields: Record<string, unknown> = {};
  if (F.Stato) fields[F.Stato] = input.decisione;
  if (F.Approvatore && Number.isFinite(approvatoreNum))
    fields[lookupIdFieldName(F.Approvatore)] = approvatoreNum;
  if (F.DataDecisione) fields[F.DataDecisione] = new Date().toISOString();
  if (F.NoteDecisione && input.noteDecisione) fields[F.NoteDecisione] = input.noteDecisione.trim();
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${input.acquistoId}/fields`, {
      method: "PATCH",
      body: JSON.stringify(fields),
    }),
  );
  logSp("info", "decide.acquisto", `Acquisto #${input.acquistoId} → ${input.decisione}`);
  return { ...current, stato: input.decisione };
}

// ---------------------------------------------------------------------------
// Finanza (direttore DR005) — movimenti bancari su lista MovimentiBancari
// ---------------------------------------------------------------------------
// L'import avviene a blocchi dal client (che ha già parsato l'xlsx): il server
// ri-classifica ogni riga con finanza-logic (unica fonte di verità), ricalcola
// la chiave dai campi grezzi e scarta i doppioni già presenti in lista. La
// sanatura manuale (updateMovimento) non tocca MAI i campi grezzi, quindi un
// ricaricamento successivo riconosce comunque il movimento originale.

function requireMovimentiList(cfg: SpDiscovered): string {
  if (!cfg.listMovimenti)
    throw new Error(
      'Lista "MovimentiBancari" non trovata su SharePoint. Crearla sul sito DRPORTAL.',
    );
  return cfg.listMovimenti;
}

export interface SpMovimento {
  id: string;
  chiave: string; // Title
  dataContabile: string; // YYYY-MM-DD
  dataValuta: string; // YYYY-MM-DD
  importo: number;
  divisa: string;
  causale: string;
  descrizione: string;
  tipologia: string;
  cliente: string;
  nrFattura: string;
  note: string;
  daVerificare: boolean;
  importId: string;
}

function mapMovimento(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): SpMovimento {
  const F = cfg.movimentiFields;
  const f = it.fields ?? {};
  const iso = (v: unknown) => String(v ?? "").slice(0, 10);
  return {
    id: String(it.id),
    chiave: String(f["Title"] ?? ""),
    dataContabile: F.DataContabile ? iso(f[F.DataContabile]) : "",
    dataValuta: F.DataValuta ? iso(f[F.DataValuta]) : "",
    importo: F.Importo ? (numOrUndef(f[F.Importo]) ?? 0) : 0,
    divisa: F.Divisa ? String(f[F.Divisa] ?? "EUR") : "EUR",
    causale: F.Causale ? String(f[F.Causale] ?? "") : "",
    descrizione: F.Descrizione ? String(f[F.Descrizione] ?? "") : "",
    tipologia: F.Tipologia ? String(f[F.Tipologia] ?? "") : "",
    cliente: F.Cliente ? String(f[F.Cliente] ?? "") : "",
    nrFattura: F.NrFattura ? String(f[F.NrFattura] ?? "") : "",
    note: F.Note ? String(f[F.Note] ?? "") : "",
    daVerificare: parseSpBool(F.DaVerificare ? f[F.DaVerificare] : undefined, false),
    importId: F.ImportId ? String(f[F.ImportId] ?? "") : "",
  };
}

// La lista supera facilmente i 999 item ($top massimo): si seguono le pagine
// @odata.nextLink convertendo l'URL Graph assoluto nel path del gateway.
async function fetchMovimentiPages(
  firstPath: string,
): Promise<GraphListItem<Record<string, unknown>>[]> {
  const out: GraphListItem<Record<string, unknown>>[] = [];
  let path: string | null = firstPath;
  let guard = 0;
  while (path && guard < 30) {
    guard++;
    const res: GraphListResponse<Record<string, unknown>> & { "@odata.nextLink"?: string } =
      await withDiscoveryRetry(() =>
        gatewayJson<GraphListResponse<Record<string, unknown>> & { "@odata.nextLink"?: string }>(
          path as string,
        ),
      );
    out.push(...(res.value ?? []));
    const next = res["@odata.nextLink"];
    if (next) {
      const u = new URL(next);
      path = u.pathname.replace(/^\/(?:v1\.0|beta)/, "") + u.search;
    } else {
      path = null;
    }
  }
  return out;
}

export interface MovimentiFilter {
  /** Data contabile minima (YYYY-MM-DD, inclusa). */
  from?: string;
  /** Data contabile massima (YYYY-MM-DD, inclusa). */
  to?: string;
  soloDaVerificare?: boolean;
}

export async function fetchMovimenti(filter: MovimentiFilter = {}): Promise<SpMovimento[]> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  if (!cfg.listMovimenti) return [];
  const items = await fetchMovimentiPages(
    `/sites/${cfg.siteId}/lists/${cfg.listMovimenti}/items?expand=fields&$top=999`,
  );
  let out = items.map((it) => mapMovimento(cfg, it));
  if (filter.from) out = out.filter((m) => m.dataContabile >= filter.from!);
  if (filter.to) out = out.filter((m) => m.dataContabile <= filter.to!);
  if (filter.soloDaVerificare) out = out.filter((m) => m.daVerificare);
  out.sort((a, b) => b.dataContabile.localeCompare(a.dataContabile) || b.id.localeCompare(a.id));
  logSp("info", "fetch.movimenti", `${out.length} movimenti`, { durataMs: Date.now() - started });
  return out;
}

/** Chiavi (Title) di tutti i movimenti già in lista, per la deduplicazione. */
export async function fetchMovimentiChiavi(): Promise<string[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listMovimenti) return [];
  const items = await fetchMovimentiPages(
    `/sites/${cfg.siteId}/lists/${cfg.listMovimenti}/items?expand=fields(select=Title)&$top=999`,
  );
  return items.map((it) => String(it.fields?.["Title"] ?? "")).filter(Boolean);
}

export type ImportMovimentoRow = MovimentoRaw & { occ: number };
export interface ImportMovimentiResult {
  ricevuti: number;
  importati: number;
  doppioni: number;
  anomalie: number;
  errori: string[];
}

const IMPORT_MOV_MAX_ROWS = 150; // per invocazione: il client spezza in blocchi

export async function importMovimenti(
  rows: ImportMovimentoRow[],
  importId: string,
): Promise<ImportMovimentiResult> {
  const cfg = await discoverSharePoint();
  const listId = requireMovimentiList(cfg);
  const F = cfg.movimentiFields;
  if (rows.length > IMPORT_MOV_MAX_ROWS)
    throw new Error(`Troppi movimenti in un blocco (max ${IMPORT_MOV_MAX_ROWS}).`);

  const esistenti = new Set(await fetchMovimentiChiavi());
  const result: ImportMovimentiResult = {
    ricevuti: rows.length,
    importati: 0,
    doppioni: 0,
    anomalie: 0,
    errori: [],
  };

  // Chiave e classificazione ricalcolate QUI dai campi grezzi: il client può
  // aver già fatto lo stesso lavoro per l'anteprima, ma la verità è del server.
  // Le regole apprese si applicano DOPO la classificazione automatica.
  const regole = await fetchRegoleFinanza().catch(() => [] as RegolaFinanza[]);
  const daScrivere: { fields: Record<string, unknown>; chiave: string }[] = [];
  for (const r of rows) {
    const chiave = chiaveMovimento(r, r.occ);
    if (esistenti.has(chiave)) {
      result.doppioni++;
      continue;
    }
    esistenti.add(chiave); // dedup anche dentro il blocco
    const c = applicaRegole({ ...classificaMovimento(r), descrizione: r.descrizione }, regole);
    if (c.daVerificare) result.anomalie++;
    const fields: Record<string, unknown> = { Title: chiave };
    if (F.DataContabile) fields[F.DataContabile] = `${r.dataContabile}T00:00:00Z`;
    if (F.DataValuta) fields[F.DataValuta] = `${r.dataValuta}T00:00:00Z`;
    if (F.Importo) fields[F.Importo] = r.importo;
    if (F.Divisa) fields[F.Divisa] = r.divisa;
    if (F.Causale) fields[F.Causale] = r.causale;
    if (F.Descrizione) fields[F.Descrizione] = r.descrizione;
    if (F.Tipologia) fields[F.Tipologia] = c.tipologia;
    if (F.Cliente && c.cliente) fields[F.Cliente] = c.cliente;
    if (F.NrFattura && c.nrFattura) fields[F.NrFattura] = c.nrFattura;
    if (F.DaVerificare) fields[F.DaVerificare] = c.daVerificare;
    if (F.ImportId && importId) fields[F.ImportId] = importId;
    daScrivere.push({ fields, chiave });
  }

  // Scritture in parallelo controllato (batch da 4) per contenere la durata.
  const BATCH = 4;
  for (let i = 0; i < daScrivere.length; i += BATCH) {
    const batch = daScrivere.slice(i, i + BATCH);
    const esiti = await Promise.allSettled(
      batch.map((b) =>
        gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
          method: "POST",
          body: JSON.stringify({ fields: b.fields }),
        }),
      ),
    );
    esiti.forEach((e, j) => {
      if (e.status === "fulfilled") result.importati++;
      else
        result.errori.push(
          `${batch[j].chiave.slice(0, 40)}…: ${
            e.reason instanceof Error ? e.reason.message : String(e.reason)
          }`,
        );
    });
  }
  logSp(
    "info",
    "import.movimenti",
    `Import movimenti: ${result.importati} scritti, ${result.doppioni} doppioni, ${result.errori.length} errori`,
  );
  return result;
}

export interface UpdateMovimentoInput {
  movimentoId: string;
  tipologia?: string;
  cliente?: string;
  nrFattura?: string;
  note?: string;
  daVerificare?: boolean;
}

// Sanatura manuale: aggiorna SOLO i campi classificati (mai i grezzi né il
// Title/chiave — la memoria dell'input originale resta intatta).
export async function updateMovimento(input: UpdateMovimentoInput): Promise<SpMovimento> {
  const cfg = await discoverSharePoint();
  const listId = requireMovimentiList(cfg);
  const F = cfg.movimentiFields;
  const fields: Record<string, unknown> = {};
  if (F.Tipologia && input.tipologia !== undefined) fields[F.Tipologia] = input.tipologia;
  if (F.Cliente && input.cliente !== undefined) fields[F.Cliente] = input.cliente;
  if (F.NrFattura && input.nrFattura !== undefined) fields[F.NrFattura] = input.nrFattura;
  if (F.Note && input.note !== undefined) fields[F.Note] = input.note;
  if (F.DaVerificare && input.daVerificare !== undefined)
    fields[F.DaVerificare] = input.daVerificare;
  if (Object.keys(fields).length === 0) throw new Error("Nessun campo da aggiornare.");
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${input.movimentoId}/fields`, {
      method: "PATCH",
      body: JSON.stringify(fields),
    }),
  );
  const updated = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items/${input.movimentoId}?expand=fields`,
    ),
  );
  logSp("info", "update.movimento", `Movimento #${input.movimentoId} sanato`);
  return mapMovimento(cfg, updated);
}

// --- Storico import + annullamento ------------------------------------------
// Lo storico è derivato dai movimenti stessi (raggruppati per ImportId): non
// serve una lista dedicata. Le righe importate prima dell'introduzione della
// colonna ImportId formano il gruppo "" (annullabile col valore speciale
// LEGACY_IMPORT_ID, definito in finanza-logic per essere usabile dal client).

export interface ImportStoricoRiga {
  importId: string; // "" per il gruppo legacy
  movimenti: number;
  anomalie: number;
  dal: string;
  al: string;
  totale: number;
}

export async function fetchImportStorico(): Promise<ImportStoricoRiga[]> {
  const all = await fetchMovimenti();
  const gruppi = new Map<string, ImportStoricoRiga>();
  for (const m of all) {
    const g = gruppi.get(m.importId) ?? {
      importId: m.importId,
      movimenti: 0,
      anomalie: 0,
      dal: m.dataContabile,
      al: m.dataContabile,
      totale: 0,
    };
    g.movimenti++;
    if (m.daVerificare) g.anomalie++;
    if (m.dataContabile && (!g.dal || m.dataContabile < g.dal)) g.dal = m.dataContabile;
    if (m.dataContabile > g.al) g.al = m.dataContabile;
    g.totale += m.importo;
    gruppi.set(m.importId, g);
  }
  // Più recenti in alto (l'ImportId inizia con un timestamp ISO); legacy in coda.
  return [...gruppi.values()].sort((a, b) => b.importId.localeCompare(a.importId));
}

// Cancella (a blocchi) tutti i movimenti di un import. Il client ripete la
// chiamata finché `rimanenti` non arriva a 0 — così nessuna invocazione supera
// i limiti di durata/subrequest del runtime.
const ANNULLA_MAX_PER_CALL = 120;

export async function annullaImport(
  importId: string,
): Promise<{ eliminati: number; rimanenti: number }> {
  const cfg = await discoverSharePoint();
  const listId = requireMovimentiList(cfg);
  const target = importId === LEGACY_IMPORT_ID ? "" : importId;
  const all = await fetchMovimenti();
  const ids = all.filter((m) => m.importId === target).map((m) => m.id);
  const daEliminare = ids.slice(0, ANNULLA_MAX_PER_CALL);
  let eliminati = 0;
  const BATCH = 4;
  for (let i = 0; i < daEliminare.length; i += BATCH) {
    // DELETE risponde 204 senza corpo: gatewayFetch diretto (come deleteTimbratura).
    const esiti = await Promise.allSettled(
      daEliminare.slice(i, i + BATCH).map(async (id) => {
        const res = await gatewayFetch(`/sites/${cfg.siteId}/lists/${listId}/items/${id}`, {
          method: "DELETE",
        });
        if (!res.ok && res.status !== 204)
          throw new SpHttpError(res.status, `DELETE movimento ${id} → ${res.status}`, "delete");
      }),
    );
    eliminati += esiti.filter((e) => e.status === "fulfilled").length;
  }
  const rimanenti = ids.length - eliminati;
  logSp(
    "info",
    "annulla.import",
    `Annullamento import "${importId}": ${eliminati} eliminati, ${rimanenti} rimanenti`,
  );
  return { eliminati, rimanenti };
}

// --- Regole apprese (lista RegoleFinanza) -----------------------------------
// Il direttore insegna al sistema le correzioni permanenti (es. un bonifico a
// un professionista → Consulenze, "kuwait" → Carburante). Le regole si applicano a
// ogni import futuro e, a richiesta, retroattivamente all'archivio.

function requireRegoleList(cfg: SpDiscovered): string {
  if (!cfg.listRegoleFinanza)
    throw new Error('Lista "RegoleFinanza" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  return cfg.listRegoleFinanza;
}

function mapRegola(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): RegolaFinanza {
  const F = cfg.regoleFinanzaFields;
  const f = it.fields ?? {};
  const campo = String(F.CampoMatch ? (f[F.CampoMatch] ?? "") : "").toLowerCase();
  const modo = String(F.ModoMatch ? (f[F.ModoMatch] ?? "") : "").toLowerCase();
  return {
    id: String(it.id),
    pattern: F.Pattern ? String(f[F.Pattern] ?? "") : "",
    campo: campo === "descrizione" ? "descrizione" : "cliente",
    modo: modo === "contiene" ? "contiene" : "esatto",
    tipologia: F.Tipologia ? String(f[F.Tipologia] ?? "").trim() || undefined : undefined,
    cliente: F.ClienteNuovo ? String(f[F.ClienteNuovo] ?? "").trim() || undefined : undefined,
  };
}

export async function fetchRegoleFinanza(): Promise<RegolaFinanza[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listRegoleFinanza) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listRegoleFinanza}/items?expand=fields&$top=999`,
    ),
  );
  return res.value.map((it) => mapRegola(cfg, it)).filter((r) => r.pattern.trim());
}

export async function createRegolaFinanza(input: RegolaFinanza): Promise<RegolaFinanza> {
  const cfg = await discoverSharePoint();
  const listId = requireRegoleList(cfg);
  const F = cfg.regoleFinanzaFields;
  if (!input.pattern.trim()) throw new Error("Il pattern della regola è obbligatorio.");
  if (!input.tipologia?.trim() && !input.cliente?.trim())
    throw new Error("La regola deve cambiare almeno la tipologia o il nome della controparte.");
  const fields: Record<string, unknown> = { Title: input.pattern.trim().slice(0, 120) };
  if (F.Pattern) fields[F.Pattern] = input.pattern.trim();
  if (F.CampoMatch) fields[F.CampoMatch] = input.campo;
  if (F.ModoMatch) fields[F.ModoMatch] = input.modo;
  if (F.Tipologia && input.tipologia?.trim()) fields[F.Tipologia] = input.tipologia.trim();
  if (F.ClienteNuovo && input.cliente?.trim()) fields[F.ClienteNuovo] = input.cliente.trim();
  const created = await withDiscoveryRetry(() =>
    gatewayJson<GraphListItem<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${listId}/items`,
      { method: "POST", body: JSON.stringify({ fields }) },
    ),
  );
  logSp("info", "create.regola", `Regola finanza "${input.pattern}" (#${created.id})`);
  return { ...input, id: String(created.id) };
}

export async function deleteRegolaFinanza(id: string): Promise<void> {
  const cfg = await discoverSharePoint();
  const listId = requireRegoleList(cfg);
  const res = await gatewayFetch(`/sites/${cfg.siteId}/lists/${listId}/items/${id}`, {
    method: "DELETE",
  });
  if (!res.ok && res.status !== 204)
    throw new SpHttpError(res.status, `DELETE regola ${id} → ${res.status}`, "delete");
  logSp("info", "delete.regola", `Rimossa regola finanza #${id}`);
}

// Applica una regola ai movimenti GIÀ in archivio, a blocchi (il client ripete
// finché rimanenti=0). Aggiorna solo i campi classificati; i grezzi e la
// chiave restano intatti (la dedup non cambia).
const APPLICA_MAX_PER_CALL = 100;

export async function applicaRegolaAiMovimenti(
  regola: RegolaFinanza,
): Promise<{ aggiornati: number; rimanenti: number }> {
  const cfg = await discoverSharePoint();
  const listId = requireMovimentiList(cfg);
  const F = cfg.movimentiFields;
  const all = await fetchMovimenti();
  const target = all.filter((m) => {
    if (!matchRegola(m, regola)) return false;
    const cambiaTip = Boolean(regola.tipologia?.trim()) && m.tipologia !== regola.tipologia?.trim();
    const cambiaCli = Boolean(regola.cliente?.trim()) && m.cliente !== regola.cliente?.trim();
    const togliFlag = Boolean(regola.tipologia?.trim()) && m.daVerificare;
    return cambiaTip || cambiaCli || togliFlag;
  });
  const batch = target.slice(0, APPLICA_MAX_PER_CALL);
  let aggiornati = 0;
  const BATCH = 4;
  for (let i = 0; i < batch.length; i += BATCH) {
    const esiti = await Promise.allSettled(
      batch.slice(i, i + BATCH).map((m) => {
        const fields: Record<string, unknown> = {};
        if (F.Tipologia && regola.tipologia?.trim()) fields[F.Tipologia] = regola.tipologia.trim();
        if (F.Cliente && regola.cliente?.trim()) fields[F.Cliente] = regola.cliente.trim();
        if (F.DaVerificare && regola.tipologia?.trim()) fields[F.DaVerificare] = false;
        return gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${m.id}/fields`, {
          method: "PATCH",
          body: JSON.stringify(fields),
        });
      }),
    );
    aggiornati += esiti.filter((e) => e.status === "fulfilled").length;
  }
  const rimanenti = target.length - aggiornati;
  logSp(
    "info",
    "applica.regola",
    `Regola "${regola.pattern}": ${aggiornati} movimenti aggiornati, ${rimanenti} rimanenti`,
  );
  return { aggiornati, rimanenti };
}

// ---------------------------------------------------------------------------
// Finanza → Fatture emesse + termini di pagamento + abbinamenti incassi
// ---------------------------------------------------------------------------
// v1: la lista è alimentata dall'export xlsx del pannello Aruba (import con
// dedup per NOME FILE SdI = Title). Predisposta per il sync via API Aruba v2:
// stessa chiave, stesso modello. Il reimport aggiorna lo Stato SdI se cambiato.

function requireFattureList(cfg: SpDiscovered): string {
  if (!cfg.listFatture)
    throw new Error('Lista "FattureEmesse" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  return cfg.listFatture;
}

export interface SpFattura extends FatturaRaw {
  id: string;
}

function mapFattura(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): SpFattura {
  const F = cfg.fattureFields;
  const f = it.fields ?? {};
  const iso = (v: unknown) => String(v ?? "").slice(0, 10);
  return {
    id: String(it.id),
    nomeFile: String(f["Title"] ?? ""),
    numero: F.Numero ? String(f[F.Numero] ?? "") : "",
    idSdi: F.IdSdi ? String(f[F.IdSdi] ?? "") : "",
    dataInvio: F.DataInvio ? iso(f[F.DataInvio]) : "",
    dataDocumento: F.DataDocumento ? iso(f[F.DataDocumento]) : "",
    tipoDocumento: F.TipoDocumento ? String(f[F.TipoDocumento] ?? "") : "",
    cliente: F.Cliente ? String(f[F.Cliente] ?? "") : "",
    piva: F.PIVA ? String(f[F.PIVA] ?? "") : "",
    metodoPagamento: F.MetodoPagamento ? String(f[F.MetodoPagamento] ?? "") : "",
    imponibile: F.Imponibile ? (numOrUndef(f[F.Imponibile]) ?? 0) : 0,
    iva: F.Iva ? (numOrUndef(f[F.Iva]) ?? 0) : 0,
    totale: F.TotaleDocumento ? (numOrUndef(f[F.TotaleDocumento]) ?? 0) : 0,
    netto: F.NettoAPagare ? (numOrUndef(f[F.NettoAPagare]) ?? 0) : 0,
    statoSdI: F.StatoSdI ? String(f[F.StatoSdI] ?? "") : "",
  };
}

export async function fetchFatture(): Promise<SpFattura[]> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  if (!cfg.listFatture) return [];
  const items = await fetchMovimentiPages(
    `/sites/${cfg.siteId}/lists/${cfg.listFatture}/items?expand=fields&$top=999`,
  );
  const out = items.map((it) => mapFattura(cfg, it));
  out.sort((a, b) => b.dataDocumento.localeCompare(a.dataDocumento));
  logSp("info", "fetch.fatture", `${out.length} fatture`, { durataMs: Date.now() - started });
  return out;
}

export interface ImportFattureResult {
  ricevute: number;
  importate: number;
  aggiornate: number; // Stato SdI cambiato su fatture già presenti
  doppioni: number;
  errori: string[];
}

const IMPORT_FAT_MAX_ROWS = 150;

export async function importFatture(rows: FatturaRaw[]): Promise<ImportFattureResult> {
  const cfg = await discoverSharePoint();
  const listId = requireFattureList(cfg);
  const F = cfg.fattureFields;
  if (rows.length > IMPORT_FAT_MAX_ROWS)
    throw new Error(`Troppe fatture in un blocco (max ${IMPORT_FAT_MAX_ROWS}).`);

  const esistenti = new Map((await fetchFatture()).map((f) => [f.nomeFile, f]));
  const result: ImportFattureResult = {
    ricevute: rows.length,
    importate: 0,
    aggiornate: 0,
    doppioni: 0,
    errori: [],
  };
  const ops: (() => Promise<"nuova" | "aggiornata">)[] = [];
  for (const r of rows) {
    if (!r.nomeFile.trim()) continue;
    const prev = esistenti.get(r.nomeFile);
    if (prev) {
      // Già in archivio: si aggiorna SOLO lo Stato SdI se cambiato (i dati
      // contabili della fattura sono immutabili per definizione).
      if (F.StatoSdI && r.statoSdI && r.statoSdI !== prev.statoSdI) {
        ops.push(async () => {
          await gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${prev.id}/fields`, {
            method: "PATCH",
            body: JSON.stringify({ [F.StatoSdI]: r.statoSdI }),
          });
          return "aggiornata";
        });
      } else {
        result.doppioni++;
      }
      continue;
    }
    const fields: Record<string, unknown> = { Title: r.nomeFile.trim() };
    if (F.Numero) fields[F.Numero] = r.numero;
    if (F.IdSdi) fields[F.IdSdi] = r.idSdi;
    if (F.DataInvio && r.dataInvio) fields[F.DataInvio] = `${r.dataInvio}T00:00:00Z`;
    if (F.DataDocumento && r.dataDocumento)
      fields[F.DataDocumento] = `${r.dataDocumento}T00:00:00Z`;
    if (F.TipoDocumento) fields[F.TipoDocumento] = r.tipoDocumento;
    if (F.Cliente) fields[F.Cliente] = r.cliente;
    if (F.PIVA) fields[F.PIVA] = r.piva;
    if (F.MetodoPagamento) fields[F.MetodoPagamento] = r.metodoPagamento;
    if (F.Imponibile) fields[F.Imponibile] = r.imponibile;
    if (F.Iva) fields[F.Iva] = r.iva;
    if (F.TotaleDocumento) fields[F.TotaleDocumento] = r.totale;
    if (F.NettoAPagare) fields[F.NettoAPagare] = r.netto;
    if (F.StatoSdI) fields[F.StatoSdI] = r.statoSdI;
    ops.push(async () => {
      await gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields }),
      });
      return "nuova";
    });
  }
  const BATCH = 4;
  for (let i = 0; i < ops.length; i += BATCH) {
    const esiti = await Promise.allSettled(ops.slice(i, i + BATCH).map((op) => op()));
    for (const e of esiti) {
      if (e.status === "fulfilled") {
        if (e.value === "nuova") result.importate++;
        else result.aggiornate++;
      } else {
        result.errori.push(e.reason instanceof Error ? e.reason.message : String(e.reason));
      }
    }
  }
  logSp(
    "info",
    "import.fatture",
    `Import fatture: ${result.importate} nuove, ${result.aggiornate} aggiornate, ${result.doppioni} doppioni`,
  );
  return result;
}

export async function fetchTerminiPagamento(): Promise<TerminePagamento[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listTermini) return [];
  const F = cfg.terminiFields;
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listTermini}/items?expand=fields&$top=999`,
    ),
  );
  return res.value
    .map((it) => {
      const f = it.fields ?? {};
      return {
        cliente: F.Cliente ? String(f[F.Cliente] ?? "").trim() : "",
        giorni: F.Giorni ? (numOrUndef(f[F.Giorni]) ?? 0) : 0,
        descrizione: F.Descrizione ? String(f[F.Descrizione] ?? "").trim() || undefined : undefined,
      };
    })
    .filter((t) => t.cliente && t.giorni > 0);
}

function requireAbbinamentiList(cfg: SpDiscovered): string {
  if (!cfg.listAbbinamenti)
    throw new Error(
      'Lista "AbbinamentiIncassi" non trovata su SharePoint. Crearla sul sito DRPORTAL.',
    );
  return cfg.listAbbinamenti;
}

function mapAbbinamento(
  cfg: SpDiscovered,
  it: GraphListItem<Record<string, unknown>>,
): AbbinamentoIncasso {
  const F = cfg.abbinamentiFields;
  const f = it.fields ?? {};
  return {
    id: String(it.id),
    fatturaFile: F.FatturaFile ? String(f[F.FatturaFile] ?? "") : "",
    movimentoChiave: F.MovimentoChiave ? String(f[F.MovimentoChiave] ?? "") : "",
    importo: F.Importo ? (numOrUndef(f[F.Importo]) ?? 0) : 0,
    origine: (F.Origine ? String(f[F.Origine] ?? "") : "") === "Manuale" ? "Manuale" : "Auto",
  };
}

export async function fetchAbbinamenti(): Promise<AbbinamentoIncasso[]> {
  const cfg = await discoverSharePoint();
  if (!cfg.listAbbinamenti) return [];
  const items = await fetchMovimentiPages(
    `/sites/${cfg.siteId}/lists/${cfg.listAbbinamenti}/items?expand=fields&$top=999`,
  );
  return items.map((it) => mapAbbinamento(cfg, it)).filter((a) => a.fatturaFile && a.importo > 0);
}

const ABB_MAX_PER_CALL = 150;

// Registra un blocco di abbinamenti (auto o manuali). Idempotente: la coppia
// fattura|movimento già presente viene scartata.
export async function createAbbinamenti(
  rows: AbbinamentoIncasso[],
): Promise<{ creati: number; scartati: number; errori: string[] }> {
  const cfg = await discoverSharePoint();
  const listId = requireAbbinamentiList(cfg);
  const F = cfg.abbinamentiFields;
  if (rows.length > ABB_MAX_PER_CALL)
    throw new Error(`Troppi abbinamenti in un blocco (max ${ABB_MAX_PER_CALL}).`);
  const esistenti = new Set(
    (await fetchAbbinamenti()).map((a) => `${a.fatturaFile}|${a.movimentoChiave}`),
  );
  const result = { creati: 0, scartati: 0, errori: [] as string[] };
  const daScrivere = rows.filter((r) => {
    const k = `${r.fatturaFile}|${r.movimentoChiave}`;
    if (esistenti.has(k) || !r.fatturaFile || !r.movimentoChiave || !(r.importo > 0)) {
      result.scartati++;
      return false;
    }
    esistenti.add(k);
    return true;
  });
  const BATCH = 4;
  for (let i = 0; i < daScrivere.length; i += BATCH) {
    const esiti = await Promise.allSettled(
      daScrivere.slice(i, i + BATCH).map((r) => {
        const fields: Record<string, unknown> = { Title: "ABB" };
        if (F.FatturaFile) fields[F.FatturaFile] = r.fatturaFile;
        if (F.MovimentoChiave) fields[F.MovimentoChiave] = r.movimentoChiave;
        if (F.Importo) fields[F.Importo] = r.importo;
        if (F.Origine) fields[F.Origine] = r.origine;
        return gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
          method: "POST",
          body: JSON.stringify({ fields }),
        });
      }),
    );
    esiti.forEach((e) => {
      if (e.status === "fulfilled") result.creati++;
      else result.errori.push(e.reason instanceof Error ? e.reason.message : String(e.reason));
    });
  }
  logSp(
    "info",
    "create.abbinamenti",
    `Abbinamenti: ${result.creati} creati, ${result.scartati} scartati`,
  );
  return result;
}

export async function deleteAbbinamento(id: string): Promise<void> {
  const cfg = await discoverSharePoint();
  const listId = requireAbbinamentiList(cfg);
  const res = await gatewayFetch(`/sites/${cfg.siteId}/lists/${listId}/items/${id}`, {
    method: "DELETE",
  });
  if (!res.ok && res.status !== 204)
    throw new SpHttpError(res.status, `DELETE abbinamento ${id} → ${res.status}`, "delete");
  logSp("info", "delete.abbinamento", `Rimosso abbinamento #${id}`);
}

// ---------------------------------------------------------------------------
// Collegamento Aruba — credenziali su lista ArubaConfig, password CIFRATA
// ---------------------------------------------------------------------------
// La password serve in chiaro per il signin, quindi non si può hashare come i
// PIN: si cifra AES-GCM con chiave derivata (SHA-256) dal segreto solo-server
// (stesso pepper dei PIN). Chi legge la lista vede "aes1$<iv>$<ct>" in base64.
// NB: se il segreto server viene ruotato, le credenziali vanno reinserite.

const ARUBA_CIPHER_PREFIX = "aes1$";

async function arubaKey(): Promise<CryptoKey> {
  const pepper = pinPepper();
  if (!pepper) throw new Error("Segreto server assente: impossibile cifrare le credenziali.");
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`aruba:${pepper}`));
  return crypto.subtle.importKey("raw", digest, { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);
}
function b64encode(b: Uint8Array): string {
  let s = "";
  for (const x of b) s += String.fromCharCode(x);
  return btoa(s);
}
function b64decode(s: string): Uint8Array {
  const raw = atob(s);
  const out = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
  return out;
}

async function cifraSegreto(testo: string): Promise<string> {
  const key = await arubaKey();
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    key,
    new TextEncoder().encode(testo),
  );
  return `${ARUBA_CIPHER_PREFIX}${b64encode(iv)}$${b64encode(new Uint8Array(ct))}`;
}

async function decifraSegreto(cifrato: string): Promise<string> {
  if (!cifrato.startsWith(ARUBA_CIPHER_PREFIX)) throw new Error("Formato credenziale non valido.");
  const [ivB64, ctB64] = cifrato.slice(ARUBA_CIPHER_PREFIX.length).split("$");
  const key = await arubaKey();
  const pt = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: b64decode(ivB64) as unknown as BufferSource },
    key,
    b64decode(ctB64) as unknown as BufferSource,
  );
  return new TextDecoder().decode(pt);
}

function requireArubaList(cfg: SpDiscovered): string {
  if (!cfg.listArubaConfig)
    throw new Error('Lista "ArubaConfig" non trovata su SharePoint. Crearla sul sito DRPORTAL.');
  return cfg.listArubaConfig;
}

async function fetchArubaRow(
  cfg: SpDiscovered,
): Promise<GraphListItem<Record<string, unknown>> | null> {
  if (!cfg.listArubaConfig) return null;
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listArubaConfig}/items?expand=fields&$top=10`,
    ),
  );
  return res.value[0] ?? null;
}

export interface ArubaStato {
  listaPresente: boolean;
  configurato: boolean;
  /** Username mascherato (mai integrale verso il client). */
  username: string;
  ultimaSync: string | null;
}

export async function getArubaStato(): Promise<ArubaStato> {
  const cfg = await discoverSharePoint();
  if (!cfg.listArubaConfig)
    return { listaPresente: false, configurato: false, username: "", ultimaSync: null };
  const F = cfg.arubaConfigFields;
  const row = await fetchArubaRow(cfg);
  const f = row?.fields ?? {};
  const username = F.Username ? String(f[F.Username] ?? "") : "";
  const pw = F.PasswordCifrata ? String(f[F.PasswordCifrata] ?? "") : "";
  const mascherato = username
    ? `${username.slice(0, 3)}${"•".repeat(Math.max(0, username.length - 3))}`
    : "";
  return {
    listaPresente: true,
    configurato: Boolean(username && pw.startsWith(ARUBA_CIPHER_PREFIX)),
    username: mascherato,
    ultimaSync: F.UltimaSync ? ((f[F.UltimaSync] as string | undefined) ?? null) : null,
  };
}

export async function saveArubaCredenziali(username: string, password: string): Promise<void> {
  const cfg = await discoverSharePoint();
  const listId = requireArubaList(cfg);
  const F = cfg.arubaConfigFields;
  const fields: Record<string, unknown> = { Title: "__aruba__" };
  if (F.Username) fields[F.Username] = username.trim();
  if (F.PasswordCifrata) fields[F.PasswordCifrata] = await cifraSegreto(password);
  const row = await fetchArubaRow(cfg);
  if (row) {
    await withDiscoveryRetry(() =>
      gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${row.id}/fields`, {
        method: "PATCH",
        body: JSON.stringify(fields),
      }),
    );
  } else {
    await withDiscoveryRetry(() =>
      gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields }),
      }),
    );
  }
  logSp("info", "aruba.config", "Credenziali Aruba aggiornate (password cifrata)");
}

/** Credenziali in chiaro — SOLO per il client API server-side. Mai loggarle. */
export async function getArubaCredenziali(): Promise<{
  username: string;
  password: string;
} | null> {
  const cfg = await discoverSharePoint();
  if (!cfg.listArubaConfig) return null;
  const F = cfg.arubaConfigFields;
  const row = await fetchArubaRow(cfg);
  if (!row) return null;
  const f = row.fields ?? {};
  const username = F.Username ? String(f[F.Username] ?? "").trim() : "";
  const cifrata = F.PasswordCifrata ? String(f[F.PasswordCifrata] ?? "") : "";
  if (!username || !cifrata.startsWith(ARUBA_CIPHER_PREFIX)) return null;
  try {
    return { username, password: await decifraSegreto(cifrata) };
  } catch {
    throw new Error(
      "Impossibile decifrare le credenziali Aruba (segreto server cambiato?). Reinserirle.",
    );
  }
}

// ---------------------------------------------------------------------------
// Web Push — storage subscription + chiavi VAPID su lista PushSubscriptions
// ---------------------------------------------------------------------------
// La riga speciale Title="__vapid__" contiene le chiavi applicative (pubblica
// nel campo P256dh, privata JWK nel campo Endpoint multiriga). Niente secret
// esterni: tutto vive su SharePoint, protetto dalle credenziali del gateway.
const VAPID_ROW_TITLE = "__vapid__";

function requirePushList(cfg: SpDiscovered): string {
  if (!cfg.listPushSubscriptions)
    throw new Error(
      'Lista "PushSubscriptions" non trovata su SharePoint. Crearla sul sito DRPORTAL.',
    );
  return cfg.listPushSubscriptions;
}

interface PushRow {
  id: string;
  title: string;
  endpoint: string;
  p256dh: string;
  auth: string;
  dipendenteId: string;
  sede: string;
}

// Le colonne "più righe di testo" possono essere in modalità testo FORMATTATO:
// SharePoint avvolge il valore in HTML (<div>…</div>) e corromperebbe endpoint
// e chiavi. Rimuovi i tag e decodifica le entità minime, difensivamente.
function stripHtml(v: string): string {
  return v
    .replace(/<[^>]*>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .trim();
}

function mapPushRow(cfg: SpDiscovered, it: GraphListItem<Record<string, unknown>>): PushRow {
  const F = cfg.pushSubscriptionsFields;
  const f = it.fields ?? {};
  return {
    id: String(it.id),
    title: String(f["Title"] ?? ""),
    endpoint: stripHtml(String((F.Endpoint ? f[F.Endpoint] : "") ?? "")),
    p256dh: stripHtml(String((F.P256dh ? f[F.P256dh] : "") ?? "")),
    auth: stripHtml(String((F.Auth ? f[F.Auth] : "") ?? "")),
    dipendenteId: F.DipendenteId ? String(f[F.DipendenteId] ?? "") : "",
    sede: F.Sede ? String(f[F.Sede] ?? "") : "",
  };
}

async function fetchPushRows(cfg: SpDiscovered): Promise<PushRow[]> {
  if (!cfg.listPushSubscriptions) return [];
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listPushSubscriptions}/items?expand=fields&$top=999`,
    ),
  );
  return res.value.map((it) => mapPushRow(cfg, it));
}

// Chiavi VAPID: lette dalla riga __vapid__, generate e salvate al primo uso.
export async function getVapidKeys(): Promise<VapidKeys> {
  const cfg = await discoverSharePoint();
  const listId = requirePushList(cfg);
  const rows = await fetchPushRows(cfg);
  const existing = rows.find((r) => r.title === VAPID_ROW_TITLE);
  if (existing && existing.p256dh && existing.endpoint) {
    return { publicKey: existing.p256dh, privateJwk: existing.endpoint };
  }
  const keys = await generateVapidKeys();
  const F = cfg.pushSubscriptionsFields;
  const fields: Record<string, unknown> = { Title: VAPID_ROW_TITLE };
  if (F.P256dh) fields[F.P256dh] = keys.publicKey;
  if (F.Endpoint) fields[F.Endpoint] = keys.privateJwk;
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
      method: "POST",
      body: JSON.stringify({ fields }),
    }),
  );
  logSp("info", "push.vapid", "Chiavi VAPID generate e salvate");
  return keys;
}

export async function getVapidPublicKey(): Promise<string> {
  return (await getVapidKeys()).publicKey;
}

// Registra (o aggiorna) la subscription del dispositivo di un dipendente.
export async function savePushSubscription(
  dipendenteId: string,
  sede: string,
  sub: PushSubscriptionData,
): Promise<void> {
  const cfg = await discoverSharePoint();
  const listId = requirePushList(cfg);
  const F = cfg.pushSubscriptionsFields;
  const rows = await fetchPushRows(cfg);
  const dup = rows.find((r) => r.endpoint === sub.endpoint && r.title !== VAPID_ROW_TITLE);
  const fields: Record<string, unknown> = { Title: `push-${dipendenteId}` };
  if (F.Endpoint) fields[F.Endpoint] = sub.endpoint;
  if (F.P256dh) fields[F.P256dh] = sub.p256dh;
  if (F.Auth) fields[F.Auth] = sub.auth;
  if (F.DipendenteId) fields[F.DipendenteId] = dipendenteId;
  if (F.Sede) fields[F.Sede] = sede;
  if (dup) {
    await withDiscoveryRetry(() =>
      gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items/${dup.id}/fields`, {
        method: "PATCH",
        body: JSON.stringify(fields),
      }),
    );
  } else {
    await withDiscoveryRetry(() =>
      gatewayJson(`/sites/${cfg.siteId}/lists/${listId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields }),
      }),
    );
  }
  logSp("info", "push.subscribe", `Subscription registrata per #${dipendenteId}`);
}

async function deletePushRow(cfg: SpDiscovered, id: string): Promise<void> {
  try {
    await gatewayFetch(`/sites/${cfg.siteId}/lists/${cfg.listPushSubscriptions}/items/${id}`, {
      method: "DELETE",
    });
  } catch {
    /* best-effort */
  }
}

// Notifica push ai dispositivi di UN dipendente (es. nuovo documento
// personale). Best-effort: nessun errore propagato.
export async function sendPushToDipendente(
  dipendenteId: string,
  payload: { title: string; body: string; url: string },
): Promise<void> {
  try {
    const cfg = await discoverSharePoint();
    if (!cfg.listPushSubscriptions) return;
    const keys = await getVapidKeys();
    const rows = (await fetchPushRows(cfg)).filter(
      (r) => r.title !== VAPID_ROW_TITLE && r.dipendenteId === dipendenteId,
    );
    for (const r of rows) {
      const res = await sendWebPush(
        { endpoint: r.endpoint, p256dh: r.p256dh, auth: r.auth },
        payload,
        keys,
      ).catch(() => null);
      if (res?.gone) await deletePushRow(cfg, r.id);
    }
  } catch {
    /* best-effort */
  }
}

// Invia una notifica push a tutti i dispositivi registrati della sede
// destinataria ("Tutte" → tutti). Best-effort: gli errori non bloccano la
// pubblicazione; le subscription morte (404/410) vengono ripulite.
export async function sendPushToSede(
  sedeDestinataria: string,
  payload: { title: string; body: string; url: string },
): Promise<{ sent: number; failed: number; errori: string[]; dispositivi: number }> {
  const cfg = await discoverSharePoint();
  if (!cfg.listPushSubscriptions)
    return { sent: 0, failed: 0, errori: ["Lista PushSubscriptions non trovata"], dispositivi: 0 };
  let keys: VapidKeys;
  try {
    keys = await getVapidKeys();
  } catch (err) {
    return {
      sent: 0,
      failed: 0,
      errori: [err instanceof Error ? err.message : String(err)],
      dispositivi: 0,
    };
  }
  const rows = (await fetchPushRows(cfg)).filter((r) => r.title !== VAPID_ROW_TITLE);
  const sedeLow = (sedeDestinataria || "Tutte").trim().toLowerCase();
  const target = rows.filter((r) => {
    if (sedeLow === "" || sedeLow === "tutte") return true;
    const s = (r.sede || "").trim().toLowerCase();
    return s === "" || s === "tutte" || s === sedeLow;
  });
  let sent = 0;
  let failed = 0;
  const errori: string[] = [];
  // Massimo 100 invii per pubblicazione (backstop). Sequenziale a gruppi di 5.
  const MAX = 100;
  const batch = target.slice(0, MAX);
  for (let i = 0; i < batch.length; i += 5) {
    const results = await Promise.allSettled(
      batch.slice(i, i + 5).map(async (r) => {
        const res = await sendWebPush(
          { endpoint: r.endpoint, p256dh: r.p256dh, auth: r.auth },
          payload,
          keys,
        );
        if (res.gone) await deletePushRow(cfg, r.id);
        if (!res.ok) throw new Error(`HTTP ${res.status} da ${new URL(r.endpoint).host}`);
        return true;
      }),
    );
    for (const r of results) {
      if (r.status === "fulfilled") sent++;
      else {
        failed++;
        if (errori.length < 3)
          errori.push(r.reason instanceof Error ? r.reason.message : String(r.reason));
      }
    }
  }
  logSp(
    failed > 0 ? "warn" : "info",
    "push.send",
    `Push "${payload.title}": inviate ${sent}, fallite ${failed}${errori.length ? ` — ${errori.join(" · ")}` : ""}`,
  );
  return { sent, failed, errori, dispositivi: target.length };
}

// ---------------------------------------------------------------------------
// Health / diagnostics
// ---------------------------------------------------------------------------
export interface SpHealth {
  graphOk: boolean;
  tokenOk: boolean;
  permissionsOk: boolean;
  siteFound: boolean;
  siteId: string | null;
  siteName: string | null;
  siteWebUrl: string | null;
  dipendentiListFound: boolean;
  dipendentiListId: string | null;
  timbratureListFound: boolean;
  timbratureListId: string | null;
  dipendentiColumnsOk: boolean;
  dipendentiMissing: string[];
  timbratureColumnsOk: boolean;
  timbratureMissing: string[];
  cacheExpiresAt: string | null;
  graphResponseMs: number;
  error: string | null;
}

export async function computeHealth(): Promise<SpHealth> {
  const empty: SpHealth = {
    graphOk: false,
    tokenOk: false,
    permissionsOk: false,
    siteFound: false,
    siteId: null,
    siteName: null,
    siteWebUrl: null,
    dipendentiListFound: false,
    dipendentiListId: null,
    timbratureListFound: false,
    timbratureListId: null,
    dipendentiColumnsOk: false,
    dipendentiMissing: [],
    timbratureColumnsOk: false,
    timbratureMissing: [],
    cacheExpiresAt: null,
    graphResponseMs: 0,
    error: null,
  };
  try {
    // Ping veloce: /sites/root verifica raggiungibilità Graph + token.
    const started = Date.now();
    const ping = await gatewayFetch(`/sites/root?$select=id`);
    empty.graphResponseMs = Date.now() - started;
    empty.graphOk = ping.status < 500;
    empty.tokenOk = ping.status !== 401;
    empty.permissionsOk = ping.status !== 403;
    if (!ping.ok) {
      empty.error = `Graph /sites/root → ${ping.status}`;
      return empty;
    }
    const disc = await discoverSharePoint();
    empty.siteFound = true;
    empty.siteId = disc.siteId;
    empty.siteName = disc.siteName;
    empty.siteWebUrl = disc.siteWebUrl;
    empty.dipendentiListFound = Boolean(disc.listDipendenti);
    empty.dipendentiListId = disc.listDipendenti;
    empty.timbratureListFound = Boolean(disc.listTimbrature);
    empty.timbratureListId = disc.listTimbrature;
    empty.dipendentiMissing = disc.dipendentiMissing;
    empty.timbratureMissing = disc.timbratureMissing;
    empty.dipendentiColumnsOk = disc.dipendentiMissing.length === 0;
    empty.timbratureColumnsOk = disc.timbratureMissing.length === 0;
    empty.cacheExpiresAt = disc.expiresAt;
    empty.graphResponseMs = lastGraphResponseMs;
  } catch (err) {
    empty.error = err instanceof Error ? err.message : String(err);
  }
  return empty;
}

// ---------------------------------------------------------------------------
// Self-test integrazione end-to-end (con rollback).
// ---------------------------------------------------------------------------
export interface SpSelfTestCheck {
  key: string;
  label: string;
  ok: boolean;
  durataMs?: number;
  message?: string;
}
export interface SpSelfTestResult {
  score: number; // 0-100
  checks: SpSelfTestCheck[];
  ranAt: string;
}

export async function runSelfTest(): Promise<SpSelfTestResult> {
  const checks: SpSelfTestCheck[] = [];
  const push = (c: SpSelfTestCheck) => checks.push(c);

  async function step(key: string, label: string, fn: () => Promise<string | void>) {
    const t = Date.now();
    try {
      const msg = await fn();
      push({ key, label, ok: true, durataMs: Date.now() - t, message: msg || undefined });
    } catch (err) {
      push({
        key,
        label,
        ok: false,
        durataMs: Date.now() - t,
        message: err instanceof Error ? err.message : String(err),
      });
    }
  }

  await step("graph", "Connessione Graph", async () => {
    const r = await gatewayFetch(`/sites/root?$select=id`);
    if (!r.ok) throw new Error(`Graph /sites/root → ${r.status}`);
  });
  await step("token", "Token valido", async () => {
    const r = await gatewayFetch(`/sites/root?$select=id`);
    if (r.status === 401) throw new Error("Token non valido (401)");
  });
  await step("permissions", "Permessi Sites.Read/ReadWrite", async () => {
    const r = await gatewayFetch(`/sites?search=*`);
    if (r.status === 403) throw new Error("Permessi insufficienti (403)");
    if (!r.ok) throw new Error(`Graph /sites → ${r.status}`);
  });

  let disc: SpDiscovered | null = null;
  await step("site", "Discovery sito DRPORTAL", async () => {
    disc = await discoverSharePoint(true);
    return disc.siteName;
  });
  await step("lists", "Discovery liste", async () => {
    if (!disc) throw new Error("Sito non trovato");
    return `${disc.listDipendentiName} · ${disc.listTimbratureName}`;
  });
  await step("columns", "Colonne obbligatorie", async () => {
    if (!disc) throw new Error("Discovery non completata");
    if (disc.dipendentiMissing.length || disc.timbratureMissing.length) {
      throw new Error(
        `Mancanti — Dip: [${disc.dipendentiMissing.join(",")}] Tim: [${disc.timbratureMissing.join(",")}]`,
      );
    }
  });

  let dipList: SpDipendente[] = [];
  let firstDip: SpDipendente | null = null;
  await step("read.dipendenti", "Lettura dipendenti", async () => {
    const list = await fetchDipendenti();
    if (list.length === 0) throw new Error("Nessun dipendente restituito");
    dipList = list;
    firstDip = list[0];
    return `${list.length} record`;
  });
  const timbratiOggi = new Set<string>();
  await step("read.timbrature", "Lettura timbrature oggi", async () => {
    const list = await fetchTimbratureOggi();
    for (const t of list) timbratiOggi.add(t.dipendenteId);
    return `${list.length} record`;
  });

  let testId: string | null = null;
  await step("write.timbratura", "Scrittura timbratura di test", async () => {
    // Sceglie un dipendente SENZA timbrature oggi (stato null → entrata
    // ammessa), così la prova di scrittura non urta la macchina a stati.
    const testDip = dipList.find((d) => !timbratiOggi.has(d.id)) ?? firstDip;
    if (!testDip) throw new Error("Nessun dipendente per il test");
    const t = await createTimbratura({
      dipendenteId: testDip.id,
      evento: "entrata",
      origine: "SelfTest",
      esito: "Test",
      note: "self-test integrazione (rollback automatico)",
    });
    testId = t.id;
    return `#${t.id}`;
  });
  await step("rollback.timbratura", "Rollback timbratura di test", async () => {
    if (!testId) throw new Error("Nessun record da eliminare");
    await deleteTimbratura(testId);
  });

  // Richieste (Sprint 2) — lista opzionale: se assente questi check falliscono
  // in modo informativo senza compromettere il resto del self-test.
  await step("list.richieste", "Discovery lista Richieste", async () => {
    if (!disc) throw new Error("Discovery non completata");
    if (!disc.listRichieste) throw new Error("Lista 'Richieste' non trovata sul sito");
    return disc.listRichiesteName ?? undefined;
  });
  await step("columns.richieste", "Colonne Richieste", async () => {
    if (!disc?.listRichieste) throw new Error("Lista Richieste assente");
    if (disc.richiesteMissing.length)
      throw new Error(`Mancanti — [${disc.richiesteMissing.join(", ")}]`);
  });
  let testRichId: string | null = null;
  await step("write.richiesta", "Scrittura richiesta di test", async () => {
    if (!disc?.listRichieste) throw new Error("Lista Richieste assente");
    if (!firstDip) throw new Error("Nessun dipendente per il test");
    const today = new Date().toISOString().slice(0, 10);
    const r = await createRichiesta({
      richiedenteId: firstDip.id,
      tipo: "Ferie",
      dataInizio: today,
      dataFine: today,
      submit: false,
    });
    testRichId = r.id;
    return r.title || `#${r.id}`;
  });
  await step("rollback.richiesta", "Rollback richiesta di test", async () => {
    if (!testRichId) throw new Error("Nessuna richiesta da eliminare");
    await deleteRichiesta(testRichId);
  });

  // Liste Sprint 4 (opzionali): visibilità su Documenti/Comunicazioni/Push.
  await step("list.documenti", "Lista DocumentiDipendenti", async () => {
    if (!disc?.listDocumenti) throw new Error("Lista 'DocumentiDipendenti' non trovata");
    if (disc.documentiMissing.length)
      throw new Error(`Colonne mancanti — [${disc.documentiMissing.join(", ")}]`);
    return disc.listDocumentiName ?? undefined;
  });
  await step("list.comunicazioni", "Lista Comunicazioni", async () => {
    if (!disc?.listComunicazioni) throw new Error("Lista 'Comunicazioni' non trovata");
    if (disc.comunicazioniMissing.length)
      throw new Error(`Colonne mancanti — [${disc.comunicazioniMissing.join(", ")}]`);
    return disc.listComunicazioniName ?? undefined;
  });
  await step("list.presevisione", "Lista PreseVisione", async () => {
    if (!disc?.listPreseVisione) throw new Error("Lista 'PreseVisione' non trovata");
    if (disc.preseVisioneMissing.length)
      throw new Error(`Colonne mancanti — [${disc.preseVisioneMissing.join(", ")}]`);
    return disc.listPreseVisioneName ?? undefined;
  });
  // Email: il connettore NON ha Mail.Send (probe storico → 403). L'invio è
  // delegato a un flusso Power Automate sulla lista CodaEmail: qui si verifica
  // che la lista esista con le colonne attese.
  await step("email.coda", "Coda email (lista CodaEmail + Power Automate)", async () => {
    if (!disc?.listCodaEmail)
      throw new Error(
        "Lista 'CodaEmail' non trovata — creare la lista e il flusso Power Automate per l'invio",
      );
    if (disc.codaEmailMissing.length)
      throw new Error(`Colonne mancanti — [${disc.codaEmailMissing.join(", ")}]`);
    return disc.listCodaEmailName ?? undefined;
  });

  // Finanza (direttore): lista movimenti bancari — opzionale.
  await step("list.movimenti", "Lista MovimentiBancari (Finanza)", async () => {
    if (!disc?.listMovimenti)
      throw new Error("Lista 'MovimentiBancari' non trovata — richiesta dalla sezione Finanza");
    if (disc.movimentiMissing.length)
      throw new Error(`Colonne mancanti — [${disc.movimentiMissing.join(", ")}]`);
    return disc.listMovimentiName ?? undefined;
  });
  await step("list.regolefinanza", "Lista RegoleFinanza (Finanza)", async () => {
    if (!disc?.listRegoleFinanza)
      throw new Error("Lista 'RegoleFinanza' non trovata — richiesta dalle regole apprese");
    if (disc.regoleFinanzaMissing.length)
      throw new Error(`Colonne mancanti — [${disc.regoleFinanzaMissing.join(", ")}]`);
    return disc.listRegoleFinanzaName ?? undefined;
  });
  await step("aruba.config", "Collegamento Aruba (lista ArubaConfig)", async () => {
    if (!disc?.listArubaConfig)
      throw new Error("Lista 'ArubaConfig' non trovata — richiesta dal collegamento Aruba");
    if (disc.arubaConfigMissing.length)
      throw new Error(`Colonne mancanti — [${disc.arubaConfigMissing.join(", ")}]`);
    const stato = await getArubaStato();
    return stato.configurato
      ? `credenziali configurate (${stato.username})`
      : "lista pronta — credenziali da inserire dalla pagina Finanza → Fatture";
  });
  await step("list.fatture", "Liste Fatture (FattureEmesse/Termini/Abbinamenti)", async () => {
    if (!disc) throw new Error("Discovery non completata");
    const mancano: string[] = [];
    if (!disc.listFatture) mancano.push("FattureEmesse");
    if (!disc.listTermini) mancano.push("TerminiPagamento");
    if (!disc.listAbbinamenti) mancano.push("AbbinamentiIncassi");
    if (mancano.length) throw new Error(`Liste mancanti: ${mancano.join(", ")}`);
    const colonne = [
      ...disc.fattureMissing.map((c) => `FattureEmesse.${c}`),
      ...disc.terminiMissing.map((c) => `TerminiPagamento.${c}`),
      ...disc.abbinamentiMissing.map((c) => `AbbinamentiIncassi.${c}`),
    ];
    if (colonne.length) throw new Error(`Colonne mancanti — [${colonne.join(", ")}]`);
    return `${disc.listFattureName} · ${disc.listTerminiName} · ${disc.listAbbinamentiName}`;
  });

  await step("push.ready", "Notifiche push (lista + chiavi VAPID)", async () => {
    if (!disc?.listPushSubscriptions) throw new Error("Lista 'PushSubscriptions' non trovata");
    if (disc.pushSubscriptionsMissing.length)
      throw new Error(`Colonne mancanti — [${disc.pushSubscriptionsMissing.join(", ")}]`);
    // Genera (o legge) le chiavi VAPID: verifica end-to-end scrittura+crypto.
    const pub = await getVapidPublicKey();
    return `chiave pubblica ${pub.slice(0, 12)}…`;
  });

  await step("latency", "Tempo risposta Graph", async () => {
    return `${lastGraphResponseMs} ms`;
  });

  const ok = checks.filter((c) => c.ok).length;
  const score = Math.round((ok / checks.length) * 100);
  logSp("info", "selfTest", `Self-test completato: ${score}/100 (${ok}/${checks.length})`);
  return { score, checks, ranAt: new Date().toISOString() };
}

// Timestamp ultima sincronizzazione presenze — aggiornato dal data layer.
let lastSyncAt: string | null = null;
export function markSync() {
  lastSyncAt = new Date().toISOString();
}
export function getLastSyncAt() {
  return lastSyncAt;
}
export function getLastGraphResponseMs() {
  return lastGraphResponseMs;
}
