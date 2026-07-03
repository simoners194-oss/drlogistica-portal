// DR Portal — SharePoint gateway helpers (server-only)
// -----------------------------------------------------------------------------
// Wrapper attorno al Lovable Connector Gateway per Microsoft SharePoint.
// Non importare mai questo file da codice client: viene bloccato dal bundler
// grazie al suffisso `.server.ts` ed è pensato per essere usato solo dalle
// server functions in `sharepoint.functions.ts`.
// -----------------------------------------------------------------------------

const GATEWAY_BASE = "https://connector-gateway.lovable.dev/microsoft_sharepoint";

// Nomi interni SharePoint dei campi delle due liste. Se in fase di deploy
// i clienti hanno rinominato le colonne, basta impostare le variabili
// VITE_SP_FIELD_* nel .env per sovrascrivere i default.
function env(key: string, fallback: string): string {
  try {
    const v = (import.meta as unknown as { env?: Record<string, string> }).env?.[key];
    const s = typeof v === "string" ? v.trim() : "";
    return s || fallback;
  } catch {
    return fallback;
  }
}

export const SP_FIELDS = {
  dipendenti: {
    Nome: env("VITE_SP_F_DIP_NOME", "Nome"),
    Cognome: env("VITE_SP_F_DIP_COGNOME", "Cognome"),
    NomeCompleto: env("VITE_SP_F_DIP_NOME_COMPLETO", "NomeCompleto"),
    Email: env("VITE_SP_F_DIP_EMAIL", "Email"),
    Sede: env("VITE_SP_F_DIP_SEDE", "Sede"),
    Attivo: env("VITE_SP_F_DIP_ATTIVO", "Attivo"),
    Ruolo: env("VITE_SP_F_DIP_RUOLO", "Ruolo"),
  },
  timbrature: {
    DipendenteLookupId: env("VITE_SP_F_TIM_DIPENDENTE_ID", "DipendenteLookupId"),
    Evento: env("VITE_SP_F_TIM_EVENTO", "Evento"),
    DataOra: env("VITE_SP_F_TIM_DATAORA", "DataOra"),
    Origine: env("VITE_SP_F_TIM_ORIGINE", "Origine"),
    Posizione: env("VITE_SP_F_TIM_POSIZIONE", "Posizione"),
    Esito: env("VITE_SP_F_TIM_ESITO", "Esito"),
    Note: env("VITE_SP_F_TIM_NOTE", "Note"),
  },
} as const;

export interface SpConfig {
  siteId: string;
  listDipendenti: string;
  listTimbrature: string;
}

export function loadSpConfig(): SpConfig {
  return {
    siteId: env("VITE_SP_SITE_ID", ""),
    listDipendenti: env("VITE_SP_LIST_DIPENDENTI_ID", ""),
    listTimbrature: env("VITE_SP_LIST_TIMBRATURE_ID", ""),
  };
}

export function isSpReady(cfg: SpConfig = loadSpConfig()): boolean {
  return Boolean(cfg.siteId && cfg.listDipendenti && cfg.listTimbrature);
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
  const cfg = loadSpConfig();
  if (!isSpReady(cfg)) throw new Error("Configurazione SharePoint incompleta.");
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
      const attivo = Boolean(f[F.Attivo]);
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
    .filter((d) => d.attivo !== false);
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

export async function fetchTimbratureOggi(): Promise<SpTimbratura[]> {
  const cfg = loadSpConfig();
  if (!isSpReady(cfg)) throw new Error("Configurazione SharePoint incompleta.");
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
    .map((it) => {
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
    .filter((x): x is SpTimbratura => x !== null && new Date(x.dataOra).getTime() >= startMs)
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
  const cfg = loadSpConfig();
  if (!isSpReady(cfg)) throw new Error("Configurazione SharePoint incompleta.");
  const F = SP_FIELDS.timbrature;
  const dipInt = Number(input.dipendenteId);
  if (!Number.isFinite(dipInt)) throw new Error("dipendenteId non valido per SharePoint (atteso ID intero della lista).");
  const dataOra = new Date().toISOString();
  const fields: Record<string, unknown> = {
    [F.DipendenteLookupId]: dipInt,
    [F.Evento]: input.evento,
    [F.DataOra]: dataOra,
    [F.Origine]: input.origine ?? "web",
    [F.Esito]: input.esito ?? "ok",
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