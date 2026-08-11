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
  saveArubaUltimaSync,
  getArubaStato,
  fetchFatture,
  importFatture,
  type DirezioneFattura,
} from "./sharepoint.server";
import { parseFatturaPA, normalizzaNomeFile, type FatturaRaw } from "./fatture-logic";
import { unzipSync } from "fflate";

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

// P.IVA del cedente (DR Logistica): obbligatoria sulla ricerca v2 insieme
// a senderCountry ('senderVatcode' con la c minuscola, come da errore API).
const ARUBA_PIVA = "16935881009";

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
    // senderCountry e' OBBLIGATORIO sulla v2 (HTTP 400 senza): il cedente
    // siamo noi, quindi Italia.
    senderCountry: "IT",
    senderVatcode: ARUBA_PIVA,
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

// --- Probe di DOWNLOAD ------------------------------------------------------
// La ricerca da' id e filename ma non il contenuto: qui si tentano gli
// endpoint di dettaglio candidati sulla prima fattura trovata e si riporta
// cosa risponde ciascuno (status, content-type, chiavi/anteprima) — cosi' il
// sync si costruisce sull'endpoint che esiste davvero, senza inventare.

export interface ArubaDownloadProbe {
  ok: boolean;
  messaggio: string;
  filename?: string;
  tentativi: {
    percorso: string;
    status: number;
    contentType?: string;
    chiavi?: string[];
    anteprima?: string;
  }[];
}

async function arubaGetRaw(
  path: string,
  params: Record<string, string>,
): Promise<{ status: number; contentType: string; testo: string }> {
  const qs = new URLSearchParams(params).toString();
  const url = `${ARUBA_WS_BASE}${path}${qs ? `?${qs}` : ""}`;
  let token = await arubaSignin();
  for (let attempt = 1; attempt <= 2; attempt++) {
    const res = await fetchConTimeout(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "*/*", "User-Agent": UA },
    });
    if (res.status === 401 && attempt === 1) {
      token = await arubaSignin(true);
      continue;
    }
    const contentType = res.headers.get("content-type") ?? "";
    const testo = await res.text().catch(() => "");
    return { status: res.status, contentType, testo };
  }
  return { status: 401, contentType: "", testo: "" };
}

export async function arubaProvaDownload(filenameRichiesto?: string): Promise<ArubaDownloadProbe> {
  await arubaSignin();
  let filename = (filenameRichiesto ?? "").trim();
  if (!filename) {
    const now = new Date();
    const start = new Date(now.getTime() - 7 * 86400000);
    const iso = (d: Date) => d.toISOString().slice(0, 19);
    const raw = await arubaGet("/api/v2/invoices-out", {
      ...paramsRicerca("out"),
      creationStartDate: iso(start),
      creationEndDate: iso(now),
      page: "1",
      size: "1",
    });
    const obj = (raw ?? {}) as Record<string, unknown>;
    const lista = Array.isArray(obj["content"]) ? (obj["content"] as unknown[]) : [];
    filename = String(((lista[0] ?? {}) as Record<string, unknown>)["filename"] ?? "");
    if (!filename)
      return {
        ok: false,
        messaggio: "Nessuna fattura negli ultimi 7 giorni: riprovare quando ce n'e' una.",
        tentativi: [],
      };
  }
  const tentativi: ArubaDownloadProbe["tentativi"] = [];
  for (const docType of ["out", "in"] as const) {
    const percorso = `/services/invoice/${docType}/getByFilename`;
    try {
      const r = await arubaGetRaw(percorso, { filename });
      const voce: ArubaDownloadProbe["tentativi"][number] = {
        percorso,
        status: r.status,
        contentType: r.contentType,
      };
      if (r.status >= 200 && r.status < 300) {
        // Verifica END-TO-END con la stessa strada del sync: estrazione
        // dell'XML e parse — cosi' un lotto "non parsato" si spiega qui.
        // Il nome file RESTITUITO puo' non coincidere con quello chiesto
        // (match lasco di Aruba senza estensione .p7m): va detto subito.
        let nota = "";
        try {
          const j = JSON.parse(r.testo) as Record<string, unknown>;
          const fnVero = String(j["filename"] ?? "");
          if (fnVero && !fnVero.startsWith(filename.replace(/\.p7m$/i, "")))
            nota = `ATTENZIONE: Aruba ha risposto con UN ALTRO file (${fnVero}) — `;
          if (!nota) {
            const campi = ["unsignedFile", "file"]
              .map((k2) => {
                const val = j[k2];
                return `${k2}=${typeof val === "string" && val ? val.length : 0}`;
              })
              .join(" ");
            nota = `${campi} — `;
          }
        } catch {
          /* diagnostica best-effort */
        }
        const xml = await arubaScaricaXml(docType, filename);
        if (!xml) {
          // Primi byte del contenuto grezzo: dicono COSA c'e' davvero
          // (p7m DER = 30 82…, ZIP = 50 4b, XML = 3c 3f…).
          let grezzo = "";
          try {
            const j = JSON.parse(r.testo) as Record<string, unknown>;
            const b64 =
              (typeof j["unsignedFile"] === "string" && (j["unsignedFile"] as string)) ||
              (typeof j["file"] === "string" && (j["file"] as string)) ||
              "";
            const bytes = b64 ? decodeBase64(b64.slice(0, 4000)) : null;
            if (bytes) {
              const hex = [...bytes.slice(0, 24)]
                .map((x) => x.toString(16).padStart(2, "0"))
                .join(" ");
              const testo = new TextDecoder("utf-8", { fatal: false })
                .decode(bytes.slice(0, 160))
                .replace(/[^\x20-\x7e]/g, ".");
              grezzo = ` — grezzo: [${hex}] "${testo}"`;
            }
          } catch {
            /* best-effort */
          }
          voce.anteprima = `${nota}estrazione FALLITA${grezzo}`;
        } else {
          const parsed = parseFatturaPA(xml, filename);
          voce.anteprima =
            nota +
            `estratti ${xml.length} caratteri; parse: ${parsed.rows.length} righe` +
            (parsed.scartati.length ? `, scarti [${parsed.scartati.join(", ")}]` : "") +
            `; inizio: ${xml.slice(0, 100)}`;
        }
        tentativi.push(voce);
        // Risposta con un file DIVERSO da quello chiesto (match lasco di
        // Aruba): non e' un successo — si prova anche l'altro canale.
        if (nota.startsWith("ATTENZIONE")) continue;
        return {
          ok: true,
          messaggio: `Dettaglio trovato (${docType}) per ${filename}.`,
          filename,
          tentativi,
        };
      }
      voce.anteprima = r.testo.slice(0, 160);
      tentativi.push(voce);
    } catch (err) {
      tentativi.push({
        percorso,
        status: 0,
        anteprima: err instanceof Error ? err.message : String(err),
      });
    }
  }
  return { ok: false, messaggio: `Nessun dettaglio per ${filename}.`, filename, tentativi };
}

// --- Probe INCASSI ----------------------------------------------------------
// Mette a nudo il dettaglio completo di UNA fattura (senza i campi-file):
// se Aruba espone lo stato incassato/pagato, sta nel campo invoices[] —
// si verifica su una fattura che l'utente SA essere segnata come incassata.

export interface ArubaIncassiProbe {
  ok: boolean;
  messaggio: string;
  filename?: string;
  /** JSON leggibile del dettaglio, senza file/unsignedFile (troncato). */
  dettaglio?: string;
}

export async function arubaProvaIncassi(filenameRichiesto?: string): Promise<ArubaIncassiProbe> {
  await arubaSignin();
  let filename = (filenameRichiesto ?? "").trim();
  if (!filename) {
    const now = new Date();
    const start = new Date(now.getTime() - 9 * 86400000);
    const iso = (d: Date) => d.toISOString().slice(0, 19);
    const raw = await arubaGet("/api/v2/invoices-out", {
      ...paramsRicerca("out"),
      creationStartDate: iso(start),
      creationEndDate: iso(now),
      page: "1",
      size: "1",
    });
    const obj = (raw ?? {}) as Record<string, unknown>;
    const lista = Array.isArray(obj["content"]) ? (obj["content"] as unknown[]) : [];
    filename = String(((lista[0] ?? {}) as Record<string, unknown>)["filename"] ?? "");
    if (!filename) return { ok: false, messaggio: "Nessuna fattura emessa negli ultimi 9 giorni." };
  }
  // La direzione del nome file non si conosce a priori: prima out, poi in.
  for (const docType of ["out", "in"] as const) {
    const r = await arubaGetRaw(`/services/invoice/${docType}/getByFilename`, { filename });
    if (r.status < 200 || r.status >= 300) continue;
    try {
      const j = JSON.parse(r.testo) as Record<string, unknown>;
      for (const k of ["file", "unsignedFile"]) {
        const val = j[k];
        j[k] = typeof val === "string" && val ? `(presente, ${val.length} caratteri)` : "(assente)";
      }
      return {
        ok: true,
        messaggio: `Dettaglio completo (${docType}): cercare i campi di incasso/pagamento dentro invoices[].`,
        filename,
        dettaglio: JSON.stringify(j, null, 2).slice(0, 6000),
      };
    } catch {
      return { ok: false, messaggio: "Risposta non JSON.", filename };
    }
  }
  return {
    ok: false,
    messaggio: `Nessun dettaglio trovato per ${filename} (ne' tra le emesse ne' tra le ricevute).`,
    filename,
  };
}

// --- SYNC FATTURE -----------------------------------------------------------
// Stessa pipeline dei caricamenti manuali (parseFatturaPA + importFatture,
// chiave anti-doppioni = nome file SdI normalizzato): qui cambia solo la
// SORGENTE — gli XML arrivano dall'API invece che dagli ZIP. La finestra di
// ricerca usa creationDate (la data di CARICAMENTO su Aruba, non quella del
// documento: cosi' si prendono anche le fatture caricate in ritardo).

export interface ArubaSyncEsito {
  direzione: DirezioneFattura;
  lotti: number;
  daScaricare: number;
  importate: number;
  aggiornate: number;
  errori: string[];
  /** false se il giro NON ha esaurito i nuovi (limite 429 o tetto download):
   *  in quel caso UltimaSync non avanza e la finestra ricopre i mancanti. */
  completo: boolean;
}

export interface ArubaSyncResult {
  finestraDa: string;
  finestraA: string;
  esiti: ArubaSyncEsito[];
}

const SYNC_MAX_PAGINE = 30;
const SYNC_MAX_DOWNLOAD = 300; // per giro: il resto al giro successivo

function paramsRicerca(docType: "out" | "in"): Record<string, string> {
  // Emesse: il cedente siamo noi. Ricevute: il cessionario siamo noi
  // (nomi parametro simmetrici: se l'API ne pretende altri, l'errore
  // integrale arrivera' nel pannello come per senderVatcode).
  return docType === "out"
    ? { senderCountry: "IT", senderVatcode: ARUBA_PIVA }
    : { receiverCountry: "IT", receiverVatcode: ARUBA_PIVA };
}

// Aruba limita ogni ricerca a 10 GIORNI di finestra (HTTP 400 oltre):
// la finestra chiesta si affetta in tranche da 9 giorni e si cuce il totale.
const SYNC_FETTA_GIORNI = 9;

async function arubaListaLotti(
  docType: "out" | "in",
  start: Date,
  end: Date,
): Promise<{ filename: string }[]> {
  const iso = (d: Date) => d.toISOString().slice(0, 19);
  const out: { filename: string }[] = [];
  const visti = new Set<string>();
  for (
    let da = new Date(start);
    da < end;
    da = new Date(da.getTime() + SYNC_FETTA_GIORNI * 86400000)
  ) {
    const a = new Date(Math.min(da.getTime() + SYNC_FETTA_GIORNI * 86400000, end.getTime()));
    for (let page = 1; page <= SYNC_MAX_PAGINE; page++) {
      const raw = await arubaGet(`/api/v2/invoices-${docType}`, {
        ...paramsRicerca(docType),
        creationStartDate: iso(da),
        creationEndDate: iso(a),
        page: String(page),
        size: "100",
      });
      const obj = (raw ?? {}) as Record<string, unknown>;
      const content = Array.isArray(obj["content"]) ? (obj["content"] as unknown[]) : [];
      for (const el of content) {
        const fn = String((el as Record<string, unknown>)["filename"] ?? "").trim();
        if (fn && !visti.has(fn)) {
          visti.add(fn);
          out.push({ filename: fn });
        }
      }
      if (content.length < 100) break;
    }
  }
  return out;
}

function decodeBase64(b64: string): Uint8Array | null {
  try {
    const bin = atob(b64.replace(/[\s]/g, "").replace(/-/g, "+").replace(/_/g, "/"));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  } catch {
    return null;
  }
}

/** Scarica l'XML di una fattura: preferisce unsignedFile (gia' senza firma);
 *  dal p7m l'XML si ritaglia dalla busta firmata. */
async function arubaScaricaXml(docType: "out" | "in", filename: string): Promise<string | null> {
  const r = await arubaGetRaw(`/services/invoice/${docType}/getByFilename`, { filename });
  if (r.status < 200 || r.status >= 300)
    throw new ArubaError(r.status, `dettaglio ${filename} → HTTP ${r.status}`);
  let j: Record<string, unknown>;
  try {
    j = JSON.parse(r.testo) as Record<string, unknown>;
  } catch {
    return null;
  }
  const b64 =
    (typeof j["unsignedFile"] === "string" && (j["unsignedFile"] as string)) ||
    (typeof j["file"] === "string" && (j["file"] as string)) ||
    "";
  if (!b64) return null;
  let bytes = decodeBase64(b64);
  if (!bytes) return null;
  // Alcune risposte impacchettano il file in uno ZIP: si apre e si prende
  // il primo XML/p7m che contiene.
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    try {
      const dentro = unzipSync(bytes);
      const nomi = Object.keys(dentro).filter((k) => /\.(xml|p7m)$/i.test(k));
      if (!nomi.length) return null;
      bytes = dentro[nomi[0]];
    } catch {
      return null;
    }
  }
  let xml = new TextDecoder("utf-8").decode(bytes).replace(/^﻿/, "");
  if (!xml.trimStart().startsWith("<")) {
    // Busta p7m (o payload sporco): si ritaglia dall'inizio dell'XML —
    // dichiarazione <?xml oppure direttamente il tag FatturaElettronica
    // (alcuni XML non hanno la dichiarazione), fino al tag di chiusura.
    const i0xml = xml.indexOf("<?xml");
    const mTag = /<[A-Za-z0-9]*:?FatturaElettronica[\s>]/.exec(xml);
    const i0 = i0xml >= 0 ? i0xml : (mTag?.index ?? -1);
    const fine = "FatturaElettronica>";
    const i1 = xml.lastIndexOf(fine);
    if (i0 < 0 || i1 <= i0) return null;
    // La firma p7m spezza il contenuto in blocchi DER e i byte di servizio
    // cadono in mezzo ai tag: spazzati via i caratteri di controllo, i tag
    // spezzati si ricongiungono (caso IT01879020517A2026_f4GEF).
    xml = xml
      .slice(i0, i1 + fine.length)
      // eslint-disable-next-line no-control-regex
      .replace(/[ --]/g, "");
  }
  return xml;
}

export async function arubaSyncFatture(giorniIndietro?: number): Promise<ArubaSyncResult> {
  await arubaSignin();
  const now = new Date();
  // Finestra: dall'ultimo sync (con 3 giorni di margine) o, in mancanza,
  // dai giorni chiesti (default 7, max 90).
  let da: Date;
  const giorni = Math.min(Math.max(giorniIndietro ?? 0, 0), 90);
  if (giorni > 0) {
    da = new Date(now.getTime() - giorni * 86400000);
  } else {
    const stato = await getArubaStato();
    da = stato.ultimaSync
      ? new Date(new Date(stato.ultimaSync).getTime() - 3 * 86400000)
      : new Date(now.getTime() - 7 * 86400000);
  }
  const iso = (d: Date) => d.toISOString().slice(0, 19);
  const esiti: ArubaSyncEsito[] = [];
  for (const [docType, direzione] of [
    ["out", "Emessa"],
    ["in", "Ricevuta"],
  ] as const) {
    const esito: ArubaSyncEsito = {
      direzione,
      lotti: 0,
      daScaricare: 0,
      importate: 0,
      aggiornate: 0,
      errori: [],
      completo: true,
    };
    try {
      const lotti = await arubaListaLotti(docType, da, now);
      esito.lotti = lotti.length;
      const presenti = new Set((await fetchFatture(direzione)).map((f) => f.nomeFile));
      const nuovi = lotti.filter((l) => !presenti.has(normalizzaNomeFile(l.filename)));
      esito.daScaricare = nuovi.length;
      const righe: FatturaRaw[] = [];
      // Aruba limita anche i DOWNLOAD (429): passo cadenzato, un solo retry
      // dopo pausa lunga, e al secondo 429 ci si ferma — i rimanenti al
      // prossimo giro (la dedup riparte da dove si era arrivati).
      const attesa = (ms: number) => new Promise((ok) => setTimeout(ok, ms));
      let fermatoPerLimite = false;
      for (const lotto of nuovi.slice(0, SYNC_MAX_DOWNLOAD)) {
        try {
          let xml: string | null = null;
          try {
            xml = await arubaScaricaXml(docType, lotto.filename);
          } catch (err) {
            if (err instanceof ArubaError && err.status === 429) {
              await attesa(12_000);
              xml = await arubaScaricaXml(docType, lotto.filename);
            } else throw err;
          }
          if (!xml) {
            esito.errori.push(`${lotto.filename}: XML non estraibile`);
            continue;
          }
          const parsed = parseFatturaPA(xml, lotto.filename);
          // Il parser deduce la direzione dalla P.IVA: si tengono solo le
          // righe coerenti con la lista di destinazione.
          righe.push(...parsed.rows.filter((r) => r.direzione === direzione));
          for (const sc of parsed.scartati) esito.errori.push(`${sc}: non parsato`);
          await attesa(400);
        } catch (err) {
          if (err instanceof ArubaError && err.status === 429) {
            fermatoPerLimite = true;
            break;
          }
          esito.errori.push(err instanceof Error ? err.message : String(err));
          if (esito.errori.length > 20) break;
        }
      }
      if (fermatoPerLimite)
        esito.errori.push(
          `Limite richieste Aruba: scaricate ${righe.length} su ${esito.daScaricare} — ripremere Sincronizza tra qualche minuto per il resto.`,
        );
      if (fermatoPerLimite || nuovi.length > SYNC_MAX_DOWNLOAD) esito.completo = false;
      for (let i = 0; i < righe.length; i += 100) {
        const res = await importFatture(righe.slice(i, i + 100), direzione);
        esito.importate += res.importate;
        esito.aggiornate += res.aggiornate;
        esito.errori.push(...res.errori);
      }
    } catch (err) {
      esito.errori.push(err instanceof Error ? err.message : String(err));
      esito.completo = false;
    }
    esiti.push(esito);
  }
  // UltimaSync avanza SOLO se tutte le direzioni hanno esaurito i nuovi:
  // un giro fermato dal limite non deve restringere la finestra successiva
  // (successe: 26 ricevute scavalcate — mai piu').
  if (esiti.every((e) => e.completo)) await saveArubaUltimaSync(now.toISOString());
  return { finestraDa: iso(da), finestraA: iso(now), esiti };
}

/** Ricerca grezza fatture emesse (per il sync, dopo la verifica del probe). */
export async function arubaSearchInvoicesOut(params: {
  startISO: string;
  endISO: string;
  page: number;
  size: number;
}): Promise<unknown> {
  return arubaGet("/api/v2/invoices-out", {
    senderCountry: "IT",
    senderVatcode: ARUBA_PIVA,
    creationStartDate: params.startISO,
    creationEndDate: params.endISO,
    page: String(params.page),
    size: String(Math.min(params.size, 100)),
  });
}
