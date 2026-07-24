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

import {
  getArubaCredenziali,
  getArubaTokenCacheRaw,
  saveArubaTokenCacheRaw,
} from "./sharepoint.server";

const ARUBA_AUTH_BASE = "https://auth.fatturazioneelettronica.aruba.it";
const ARUBA_WS_BASE = "https://ws.fatturazioneelettronica.aruba.it";
const TIMEOUT_MS = 20_000;
// Le fetch dei Worker partono senza User-Agent e i WAF trattano il traffico
// anonimo in modo aggressivo (rate limit/429): ci si identifica sempre.
const UA = "DRPortal/1.7 (portal.drlogistica.it; integrazione fatturazione)";

// Aruba LIMITA i signin ripetuti (HTTP 429): il token va RIUTILIZZATO.
// Strategia a tre livelli: memoria del Worker → cache persistita (cifrata,
// lista ArubaConfig, sopravvive ai riavvii) → refresh_token → signin pieno.
interface TokenSet {
  access: string;
  refresh: string;
  accessScadeA: number; // epoch ms
  refreshScadeA: number; // epoch ms
}
let tokenMem: TokenSet | null = null;

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

async function authRequest(body: URLSearchParams): Promise<TokenSet> {
  let res: Response;
  try {
    res = await fetchConTimeout(`${ARUBA_AUTH_BASE}/auth/signin`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Accept: "application/json",
        "User-Agent": UA,
      },
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
    if (res.status === 429)
      throw new ArubaError(
        429,
        "L'ACCOUNT Aruba è in raffreddamento per troppi tentativi di accesso (HTTP 429). Ogni nuovo tentativo, anche fallito, riazzera il timer: NON riprovare per almeno 2 ore, poi UN solo tentativo.",
      );
    throw new ArubaError(
      res.status,
      res.status === 401 || res.status === 400
        ? "Autenticazione Aruba rifiutata: verifica username e password del servizio."
        : `Autenticazione Aruba fallita (HTTP ${res.status}).`,
    );
  }
  const data = (await res.json()) as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
  };
  if (!data.access_token) throw new ArubaError(500, "Risposta di autenticazione Aruba non valida.");
  // Margini prudenti: access 30' (dichiarato), refresh 60' da documentazione.
  const accessTtl = (Number(data.expires_in) > 0 ? Number(data.expires_in) : 1800) * 1000;
  return {
    access: data.access_token,
    refresh: data.refresh_token ?? "",
    accessScadeA: Date.now() + accessTtl - 60_000,
    refreshScadeA: Date.now() + 60 * 60_000 - 60_000,
  };
}

async function salvaToken(ts: TokenSet): Promise<void> {
  tokenMem = ts;
  await saveArubaTokenCacheRaw(JSON.stringify(ts)); // best-effort, cifrata
}

async function arubaSignin(force = false): Promise<string> {
  const now = Date.now();
  if (!force) {
    // 1) memoria del Worker
    if (tokenMem && tokenMem.accessScadeA > now) return tokenMem.access;
    // 2) cache persistita (sopravvive ai riavvii del Worker)
    if (!tokenMem) {
      try {
        const raw = await getArubaTokenCacheRaw();
        if (raw) {
          const ts = JSON.parse(raw) as TokenSet;
          if (ts?.access && ts.accessScadeA > now) {
            tokenMem = ts;
            return ts.access;
          }
          if (ts?.refresh && ts.refreshScadeA > now) tokenMem = ts; // per il refresh sotto
        }
      } catch {
        /* cache illeggibile: si prosegue col signin */
      }
    }
    // 3) refresh del token (più leggero e non soggetto al limite dei signin)
    if (tokenMem?.refresh && tokenMem.refreshScadeA > now) {
      try {
        const ts = await authRequest(
          new URLSearchParams({ grant_type: "refresh_token", refresh_token: tokenMem.refresh }),
        );
        await salvaToken(ts);
        return ts.access;
      } catch {
        /* refresh fallito: si ricade sul signin pieno */
      }
    }
  }
  // 4) signin pieno con le credenziali — UN SOLO tentativo. Il limite di
  // Aruba si è rivelato PER ACCOUNT (anti brute-force): riprovare in
  // automatico non aiuta, anzi riazzera il raffreddamento. Diagnosi 24/07:
  // stesso IP, utente finto → 400, utente reale → 429.
  const cred = await getArubaCredenziali();
  if (!cred) throw new ArubaError(0, "Credenziali Aruba non configurate.");
  const ts = await authRequest(
    new URLSearchParams({
      grant_type: "password",
      username: cred.username,
      password: cred.password,
    }),
  );
  await salvaToken(ts);
  return ts.access;
}

async function arubaGet(path: string, params: Record<string, string>): Promise<unknown> {
  const qs = new URLSearchParams(params).toString();
  const url = `${ARUBA_WS_BASE}${path}${qs ? `?${qs}` : ""}`;
  let token = await arubaSignin();
  for (let attempt = 1; attempt <= 2; attempt++) {
    let res: Response;
    try {
      res = await fetchConTimeout(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "User-Agent": UA,
        },
      });
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
  // RIUSA il token quando c'è (Aruba limita i signin ripetuti — 429): la
  // verifica delle credenziali avviene comunque al primo accesso reale.
  await arubaSignin();
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
