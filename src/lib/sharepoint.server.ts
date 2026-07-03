// DR Portal — SharePoint gateway helpers (server-only)
// -----------------------------------------------------------------------------
// Wrapper attorno al Lovable Connector Gateway per Microsoft SharePoint.
// Non importare mai questo file da codice client: viene bloccato dal bundler
// grazie al suffisso `.server.ts` ed è pensato per essere usato solo dalle
// server functions in `sharepoint.functions.ts`.
//
// Il sito SharePoint e le due liste ("Dipendenti", "Timbrature") vengono
// scoperti in automatico via Microsoft Graph — nessuna variabile d'ambiente
// richiesta. La configurazione discovered viene tenuta in cache in memoria
// per l'intera vita del worker.
// -----------------------------------------------------------------------------

const GATEWAY_BASE = "https://connector-gateway.lovable.dev/microsoft_sharepoint";

// Nomi effettivi delle colonne SharePoint (schema reale del tenant DR).
export const SP_FIELDS = {
  dipendenti: {
    Nome: "Nome",
    Cognome: "Cognome",
    NomeCompleto: "NomeCompleto",
    Email: "Email",
    Sede: "Sede",
    Attivo: "Attivo",
    // Colonna "Responsabile" della lista SharePoint (valori: Dipendente / Responsabile).
    Ruolo: "Responsabile",
  },
  timbrature: {
    DipendenteLookupId: "DipendenteLookupId",
    Evento: "Evento",
    DataOra: "DataOra",
    Origine: "Dispositivo",
    Posizione: "GeoLoc",
    Esito: "Esito",
    Note: "Note",
  },
} as const;

// Nomi delle liste SharePoint da individuare (case-insensitive, tolleranti a
// varianti singolare/plurale).
const LIST_NAMES = {
  dipendenti: ["Dipendenti", "Dipendente"],
  timbrature: ["Timbrature", "Timbratura"],
} as const;

export interface SpDiscovered {
  siteId: string;
  siteName: string;
  siteWebUrl: string;
  listDipendenti: string;
  listDipendentiName: string;
  listTimbrature: string;
  listTimbratureName: string;
}

let discoveredCache: SpDiscovered | null = null;

export function clearSpDiscoveryCache() {
  discoveredCache = null;
}

export function getSpDiscoveryCached(): SpDiscovered | null {
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
  const res = await fetch(`${GATEWAY_BASE}${path}`, { ...init, headers });
  return res;
}

async function gatewayJson<T = unknown>(path: string, init: RequestInit = {}): Promise<T> {
  const res = await gatewayFetch(path, init);
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`SharePoint ${init.method ?? "GET"} ${path} → ${res.status} ${body.slice(0, 300)}`);
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

function matchListName(list: GraphList, targets: readonly string[]): boolean {
  const candidates = [list.displayName, list.name].filter(Boolean).map((s) => s!.toLowerCase());
  return targets.some((t) => candidates.includes(t.toLowerCase()));
}

export async function discoverSharePoint(force = false): Promise<SpDiscovered> {
  if (!force && discoveredCache) return discoveredCache;

  // 1. Elenca i siti (Graph usa il parametro `search` diretto, non `$search`).
  const sitesRes = await gatewayJson<{ value: GraphSite[] }>(`/sites?search=*`);
  const sites = sitesRes.value ?? [];
  if (sites.length === 0) {
    throw new Error(
      "Microsoft Graph non ha restituito nessun sito SharePoint accessibile. Verifica che al connettore sia stato concesso il permesso Sites.Read.All (o Sites.ReadWrite.All).",
    );
  }

  // 2. Per ogni sito, cerca le due liste. Il primo che le contiene entrambe vince.
  const errors: string[] = [];
  for (const site of sites) {
    try {
      const listsRes = await gatewayJson<{ value: GraphList[] }>(
        `/sites/${site.id}/lists?$select=id,name,displayName,list`,
      );
      const lists = (listsRes.value ?? []).filter((l) => !l.list?.hidden);
      const dip = lists.find((l) => matchListName(l, LIST_NAMES.dipendenti));
      const tim = lists.find((l) => matchListName(l, LIST_NAMES.timbrature));
      if (dip && tim) {
        discoveredCache = {
          siteId: site.id,
          siteName: site.displayName || site.name || site.id,
          siteWebUrl: site.webUrl ?? "",
          listDipendenti: dip.id,
          listDipendentiName: dip.displayName || dip.name || dip.id,
          listTimbrature: tim.id,
          listTimbratureName: tim.displayName || tim.name || tim.id,
        };
        return discoveredCache;
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errors.push(`${site.displayName || site.name}: ${msg}`);
    }
  }

  const sitesList = sites.map((s) => s.displayName || s.name).filter(Boolean).join(", ");
  throw new Error(
    `Nessun sito SharePoint contiene entrambe le liste "Dipendenti" e "Timbrature". Siti esaminati: ${sitesList}.${
      errors.length ? ` Errori: ${errors.join(" | ")}` : ""
    }`,
  );
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
function normalizeSede(v: SedeRaw): "roma" | "san-giuliano" {
  const s = (v ?? "").toString().trim().toLowerCase().replace(/\s+/g, "-");
  if (s.startsWith("san")) return "san-giuliano";
  return "roma";
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
  sede: "roma" | "san-giuliano";
  attivo: boolean;
  ruolo: string;
}

export async function fetchDipendenti(): Promise<SpDipendente[]> {
  const cfg = await discoverSharePoint();
  const F = SP_FIELDS.dipendenti;
  const res = await gatewayJson<GraphListResponse<Record<string, unknown>>>(
    `/sites/${cfg.siteId}/lists/${cfg.listDipendenti}/items?expand=fields&$top=999`,
  );
  return res.value
    .map((it) => {
      const f = it.fields ?? {};
      const nome = String(f[F.Nome] ?? "").trim();
      const cognome = String(f[F.Cognome] ?? "").trim();
      const nomeCompleto = String(f[F.NomeCompleto] ?? `${nome} ${cognome}`).trim();
      const rawAttivo = f[F.Attivo];
      const attivo = rawAttivo === undefined ? true : Boolean(rawAttivo);
      return {
        id: String(it.id),
        nome,
        cognome,
        nomeCompleto,
        email: String(f[F.Email] ?? "").trim(),
        sede: normalizeSede(f[F.Sede] as SedeRaw),
        attivo,
        ruolo: String(f[F.Ruolo] ?? "").trim(),
      };
    })
    .filter((d) => d.attivo);
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
  // Colonna Evento su SharePoint accetta esattamente questi choice.
  return { entrata: "Entrata", "inizio-pausa": "Inizio Pausa", "fine-pausa": "Fine Pausa", uscita: "Uscita" }[e];
}

export async function fetchTimbratureOggi(): Promise<SpTimbratura[]> {
  const cfg = await discoverSharePoint();
  const F = SP_FIELDS.timbrature;
  // Filtro server-side su DataOra >= inizio giornata. Se il campo non è
  // indicizzato SharePoint può rifiutare il filtro: in tal caso ripieghiamo
  // sul filtraggio client-side.
  const filter = encodeURIComponent(`fields/${F.DataOra} ge '${todayIsoStart()}'`);
  let path = `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items?expand=fields&$top=999&$orderby=fields/${F.DataOra} asc&$filter=${filter}`;
  let res: GraphListResponse<Record<string, unknown>>;
  try {
    res = await gatewayJson<GraphListResponse<Record<string, unknown>>>(path);
  } catch {
    path = `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items?expand=fields&$top=999`;
    res = await gatewayJson<GraphListResponse<Record<string, unknown>>>(path);
  }
  const startMs = new Date(todayIsoStart()).getTime();
  return res.value
    .map((it): SpTimbratura | null => {
      const f = it.fields ?? {};
      const evento = parseEvento(f[F.Evento]);
      const dataOra = String(f[F.DataOra] ?? "");
      const dipRaw = f[F.DipendenteLookupId];
      return evento && dataOra && dipRaw != null
        ? {
            id: String(it.id),
            dipendenteId: String(dipRaw),
            evento,
            dataOra,
            origine: f[F.Origine] ? String(f[F.Origine]) : undefined,
            posizione: f[F.Posizione] ? String(f[F.Posizione]) : undefined,
            esito: f[F.Esito] ? String(f[F.Esito]) : undefined,
            note: f[F.Note] ? String(f[F.Note]) : undefined,
          }
        : null;
    })
    .filter(
      (x): x is SpTimbratura => x !== null && new Date(x.dataOra).getTime() >= startMs,
    )
    .sort((a, b) => a.dataOra.localeCompare(b.dataOra));
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
  const cfg = await discoverSharePoint();
  const F = SP_FIELDS.timbrature;
  const dipInt = Number(input.dipendenteId);
  if (!Number.isFinite(dipInt)) throw new Error("dipendenteId non valido per SharePoint (atteso ID intero della lista).");
  const dataOra = new Date().toISOString();
  const fields: Record<string, unknown> = {
    [F.DipendenteLookupId]: dipInt,
    [F.Evento]: eventoToSharePoint(input.evento),
    [F.DataOra]: dataOra,
    [F.Origine]: (input.origine ?? "Web").replace(/^\w/, (c) => c.toUpperCase()),
    [F.Esito]: input.esito ?? "Accettata",
  };
  if (input.posizione) fields[F.Posizione] = input.posizione;
  if (input.note) fields[F.Note] = input.note;

  const created = await gatewayJson<GraphListItem<Record<string, unknown>>>(
    `/sites/${cfg.siteId}/lists/${cfg.listTimbrature}/items`,
    { method: "POST", body: JSON.stringify({ fields }) },
  );
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