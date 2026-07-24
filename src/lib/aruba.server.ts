// DR Portal — Client API Aruba Fatturazione Elettronica v2 (server-only).
// -----------------------------------------------------------------------------
// SOLA LETTURA, ambiente di PRODUZIONE (le GET non toccano lo SdI; l'invio
// fatture NON è implementato e resterà dietro un flag esplicito).
// Riferimento: https://fatturazioneelettronica.aruba.it/apidoc/v2/docs.html
//   - autenticazione: POST {AUTH}/auth/signin (x-www-form-urlencoded),
//     access_token valido 30 minuti (Bearer).
//   - ricerca emesse:  GET {WS}/api/v2/invoices-out  (paginata, size max 100)
//   - ricerca ricevute: GET {WS}/api/v2/invoices-in
// I NOMI dei parametri data e i campi della risposta variano per versione:
// il probe (arubaProvaConnessione) restituisce la forma REALE della risposta
// così la mappatura del sync si finalizza sul dato vero, senza inventare.
// Le credenziali vivono su SharePoint cifrate (vedi sharepoint.server.ts);
// qui non si logga MAI né password né token.

import { getArubaCredenziali } from "./sharepoint.server";

const ARUBA_AUTH_BASE = "https://auth.fatturazioneelettronica.aruba.it";
const ARUBA_WS_BASE = "https://ws.fatturazioneelettronica.aruba.it";
const TIMEOUT_MS = 20_000;

// Token in-memory (il Worker è effimero: alla peggio si rifà il signin).
let tokenCache: { accessToken: string; scadeA: number } | null = null;

async function fetchConTimeout(url: string, init: RequestInit): Promise<Response> {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  try {
    return await fetch(url, { ...init, signal: ctl.signal });
  } finally {
    clearTimeout(timer);
  }
}

export class ArubaError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
  }
}

async function arubaSignin(force = false): Promise<string> {
  if (!force && tokenCache && tokenCache.scadeA > Date.now()) return tokenCache.accessToken;
  const cred = await getArubaCredenziali();
  if (!cred) throw new ArubaError(0, "Credenziali Aruba non configurate.");
  const body = new URLSearchParams({
    grant_type: "password",
    username: cred.username,
    password: cred.password,
  });
  let res: Response;
  try {
    res = await fetchConTimeout(`${ARUBA_AUTH_BASE}/auth/signin`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });
  } catch (err) {
    throw new ArubaError(
      0,
      err instanceof Error && err.name === "AbortError"
        ? "Timeout di connessione al servizio di autenticazione Aruba."
        : "Servizio di autenticazione Aruba non raggiungibile.",
    );
  }
  if (!res.ok) {
    // Mai riportare il corpo integrale: può contenere echi dei parametri.
    throw new ArubaError(
      res.status,
      res.status === 401 || res.status === 400
        ? "Autenticazione Aruba rifiutata: verifica username e password del servizio."
        : `Autenticazione Aruba fallita (HTTP ${res.status}).`,
    );
  }
  const data = (await res.json()) as { access_token?: string; expires_in?: number };
  if (!data.access_token) throw new ArubaError(500, "Risposta di autenticazione Aruba non valida.");
  // Margine di 60s sulla scadenza dichiarata (default 30 minuti).
  const ttlMs = (Number(data.expires_in) > 0 ? Number(data.expires_in) : 1800) * 1000;
  tokenCache = { accessToken: data.access_token, scadeA: Date.now() + ttlMs - 60_000 };
  return tokenCache.accessToken;
}

async function arubaGet(path: string, params: Record<string, string>): Promise<unknown> {
  const qs = new URLSearchParams(params).toString();
  const url = `${ARUBA_WS_BASE}${path}${qs ? `?${qs}` : ""}`;
  let token = await arubaSignin();
  for (let attempt = 1; attempt <= 2; attempt++) {
    let res: Response;
    try {
      res = await fetchConTimeout(url, { headers: { Authorization: `Bearer ${token}` } });
    } catch (err) {
      throw new ArubaError(
        0,
        err instanceof Error && err.name === "AbortError"
          ? "Timeout della richiesta ad Aruba."
          : "Servizio Aruba non raggiungibile.",
      );
    }
    if (res.status === 401 && attempt === 1) {
      // Token scaduto/invalidato: un solo nuovo signin, mai loop.
      token = await arubaSignin(true);
      continue;
    }
    if (res.status === 429)
      throw new ArubaError(429, "Limite richieste Aruba raggiunto: riprovare tra qualche minuto.");
    if (!res.ok) {
      const snippet = (await res.text().catch(() => "")).slice(0, 300);
      throw new ArubaError(res.status, `Aruba GET ${path} → HTTP ${res.status} ${snippet}`);
    }
    return res.json();
  }
  throw new ArubaError(401, "Autenticazione Aruba non valida dopo il rinnovo del token.");
}

// --- Probe di connessione ----------------------------------------------------
// Verifica end-to-end: signin + una ricerca minima sulle fatture emesse degli
// ultimi 2 giorni (finestra piccola: alcuni piani limitano l'ampiezza del
// range — DA VERIFICARE NELLA DOCUMENTAZIONE ARUBA il massimo consentito).
// Restituisce la FORMA della risposta (nomi dei campi + valori troncati del
// primo elemento) per finalizzare la mappatura del sync sul dato reale.

export interface ArubaProbeResult {
  ok: boolean;
  messaggio: string;
  /** Chiavi di primo livello della risposta (es. content/totalElements/...). */
  chiaveRisposta?: string[];
  /** Numero di elementi trovati nella finestra di prova. */
  elementi?: number;
  /** Campi del primo elemento con valore troncato (per la mappatura). */
  campiEsempio?: Record<string, string>;
}

function truncVal(v: unknown): string {
  const s = typeof v === "object" ? JSON.stringify(v) : String(v ?? "");
  return s.length > 60 ? `${s.slice(0, 60)}…` : s;
}

export async function arubaProvaConnessione(): Promise<ArubaProbeResult> {
  await arubaSignin(true); // verifica esplicita delle credenziali
  const now = new Date();
  const start = new Date(now.getTime() - 2 * 86400000);
  const iso = (d: Date) => d.toISOString().slice(0, 19);
  // Parametri come da documentazione v2 (creationStartDate/creationEndDate,
  // page/size). In caso di errore il messaggio HTTP viene riportato integro
  // (troncato) proprio per correggere rapidamente nomi/formati.
  const raw = await arubaGet("/api/v2/invoices-out", {
    creationStartDate: iso(start),
    creationEndDate: iso(now),
    page: "1",
    size: "5",
  });
  const obj = (raw ?? {}) as Record<string, unknown>;
  const lista = Array.isArray(obj["content"])
    ? (obj["content"] as unknown[])
    : Array.isArray(raw)
      ? (raw as unknown[])
      : [];
  const primo = (lista[0] ?? null) as Record<string, unknown> | null;
  return {
    ok: true,
    messaggio: `Connessione OK — ${lista.length} fatture nella finestra di prova (2 giorni).`,
    chiaveRisposta: Object.keys(obj),
    elementi: lista.length,
    campiEsempio: primo
      ? Object.fromEntries(Object.entries(primo).map(([k, v]) => [k, truncVal(v)]))
      : undefined,
  };
}

/** Ricerca grezza fatture emesse (per il sync, dopo la verifica del probe). */
export async function arubaSearchInvoicesOut(params: {
  startISO: string;
  endISO: string;
  page: number;
  size: number;
}): Promise<unknown> {
  return arubaGet("/api/v2/invoices-out", {
    creationStartDate: params.startISO,
    creationEndDate: params.endISO,
    page: String(params.page),
    size: String(Math.min(params.size, 100)),
  });
}
