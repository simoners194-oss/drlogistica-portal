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
  },
  timbrature: {
    Dipendente: "Dipendente",
    Evento: "Evento",
    DataOra: "DataOra",
    Origine: "Dispositivo",
    Posizione: "GeoLoc",
    Esito: "Esito",
    Note: "Note",
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
const REQUIRED_TIM_KEYS = [
  "Dipendente",
  "Evento",
  "DataOra",
  "Origine",
  "Esito",
  "Note",
  "Posizione",
] as const;

// Nomi delle liste SharePoint da individuare (case-insensitive, tolleranti a
// varianti singolare/plurale).
const LIST_NAMES = {
  dipendenti: ["Dipendenti", "Dipendente"],
  timbrature: ["Timbrature", "Timbratura"],
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
    .replace(/Bearer\s+[A-Za-z0-9._\-]+/gi, "Bearer ***")
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
  // Mappa "chiave logica" -> internalName reale su SharePoint.
  dipendentiFields: Record<string, string>;
  timbratureFields: Record<string, string>;
  dipendentiMissing: string[];
  timbratureMissing: string[];
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
  constructor(public status: number, message: string, public path: string) {
    super(message);
  }
}

async function gatewayJson<T = unknown>(path: string, init: RequestInit = {}): Promise<T> {
  const res = await gatewayFetch(path, init);
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new SpHttpError(
      res.status,
      `SharePoint ${init.method ?? "GET"} ${path.split("?")[0]} → ${res.status} ${sanitize(body)}`,
      path,
    );
  }
  return (await res.json()) as T;
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
    const site = await gatewayJson<GraphSite>(
      `/sites/${TARGET_HOST}:/sites/${TARGET_SITE_PATH}`,
    );
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

  const now = Date.now();
  discoveredCache = {
    siteId: targetSite.id,
    siteName: targetSite.displayName || targetSite.name || targetSite.id,
    siteWebUrl: targetSite.webUrl ?? "",
    listDipendenti: dip.id,
    listDipendentiName: dip.displayName || dip.name || dip.id,
    listTimbrature: tim.id,
    listTimbratureName: tim.displayName || tim.name || tim.id,
    dipendentiFields: dipRes.map,
    timbratureFields: timRes.map,
    dipendentiMissing: dipRes.missing,
    timbratureMissing: timRes.missing,
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
function normalizeSede(v: SedeRaw): "roma" | "san-giuliano" | "tutte" {
  const s = (v ?? "").toString().trim().toLowerCase().replace(/\s+/g, "-");
  if (s === "tutte" || s === "all" || s === "*") return "tutte";
  if (s.startsWith("san")) return "san-giuliano";
  return "roma";
}

function requireField(map: Record<string, string>, key: string, list: "Dipendenti" | "Timbrature"): string {
  const v = map[key];
  if (!v) throw new Error(`Colonna obbligatoria "${key}" mancante nella lista ${list} su SharePoint.`);
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
  sede: "roma" | "san-giuliano" | "tutte";
  attivo: boolean;
  ruolo: string;
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
      };
    })
    .filter((d) => d.attivo);
  logSp("info", "fetch.dipendenti", `${out.length} dipendenti attivi`, {
    durataMs: Date.now() - started,
  });
  return out;
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
  return String(v ?? "").trim().toUpperCase();
}

function normalizePin(v: unknown): string {
  return String(v ?? "").trim();
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
  const res = await withDiscoveryRetry(() =>
    gatewayJson<GraphListResponse<Record<string, unknown>>>(
      `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items?expand=fields&$top=999`,
    ),
  );
  const attivoField = F.Attivo;
  const match = res.value.find((it) => {
    const f = it.fields ?? {};
    const c = normalizeCodice(f[codiceField]);
    const p = normalizePin(f[pinField]);
    const attivo = attivoField ? Boolean(f[attivoField]) : true;
    return attivo && c === codice && p === pin;
  });
  if (!match) {
    logSp("warn", "login", `Tentativo fallito per codice="${codice}"`, {
      durataMs: Date.now() - started,
    });
    return { ok: false, error: "Codice o PIN non validi." };
  }
  const f = match.fields ?? {};
  const nome = String(f[F.Nome ?? ""] ?? "").trim();
  const cognome = String(f[F.Cognome ?? ""] ?? "").trim();
  const dipendente: SpDipendente = {
    id: String(match.id),
    nome,
    cognome,
    nomeCompleto:
      String(f[F.NomeCompleto ?? ""] ?? `${nome} ${cognome}`).trim(),
    email: String(f[F.Email ?? ""] ?? "").trim(),
    sede: normalizeSede((F.Sede ? f[F.Sede] : undefined) as SedeRaw),
    attivo: true,
    ruolo: String(f[F.Ruolo ?? ""] ?? "").trim(),
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
  const s = String(v ?? "").toLowerCase().replace(/\s+/g, "-");
  if (s === "entrata" || s === "inizio-pausa" || s === "fine-pausa" || s === "uscita") return s;
  return null;
}

function eventoToSharePoint(e: EventoTimbratura): string {
  return { entrata: "Entrata", "inizio-pausa": "Inizio Pausa", "fine-pausa": "Fine Pausa", uscita: "Uscita" }[e];
}

// Costruisce il nome del lookup id partendo dall'internal name della colonna.
// Esempio: internal "Dipendente" -> "DipendenteLookupId"; internal "Dipendente0"
// -> "Dipendente0LookupId".
function lookupIdFieldName(internal: string): string {
  return `${internal}LookupId`;
}

export async function fetchTimbratureOggi(): Promise<SpTimbratura[]> {
  const started = Date.now();
  const cfg = await discoverSharePoint();
  const F = cfg.timbratureFields;
  const dataOraField = requireField(F, "DataOra", "Timbrature");
  const eventoField = requireField(F, "Evento", "Timbrature");
  const dipendenteField = requireField(F, "Dipendente", "Timbrature");
  const lookupId = lookupIdFieldName(dipendenteField);

  const filter = encodeURIComponent(`fields/${dataOraField} ge '${todayIsoStart()}'`);
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
  const startMs = new Date(todayIsoStart()).getTime();
  const out = res.value
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
    .filter(
      (x): x is SpTimbratura => x !== null && new Date(x.dataOra).getTime() >= startMs,
    )
    .sort((a, b) => a.dataOra.localeCompare(b.dataOra));
  logSp("info", "fetch.timbrature", `${out.length} timbrature oggi`, {
    durataMs: Date.now() - started,
  });
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

  let firstDip: SpDipendente | null = null;
  await step("read.dipendenti", "Lettura dipendenti", async () => {
    const list = await fetchDipendenti();
    if (list.length === 0) throw new Error("Nessun dipendente restituito");
    firstDip = list[0];
    return `${list.length} record`;
  });
  await step("read.timbrature", "Lettura timbrature oggi", async () => {
    const list = await fetchTimbratureOggi();
    return `${list.length} record`;
  });

  let testId: string | null = null;
  await step("write.timbratura", "Scrittura timbratura di test", async () => {
    if (!firstDip) throw new Error("Nessun dipendente per il test");
    const t = await createTimbratura({
      dipendenteId: firstDip.id,
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