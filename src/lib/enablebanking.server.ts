// DR Portal — Collegamento banca via Enable Banking (PSD2/open banking).
// -----------------------------------------------------------------------------
// SOLA LETTURA: il portale legge i movimenti del conto aziendale Banco BPM
// autorizzato dalla direzione (SCA con l'app della banca, consenso ~89 giorni).
// Nessuna operazione dispositiva esiste né può esistere qui.
//
// Autenticazione: ogni richiesta porta un JWT RS256 firmato con la chiave
// privata dell'applicazione registrata su enablebanking.com (kid = app id).
// La chiave privata vive CIFRATA su SharePoint (vedi sharepoint.server.ts) e
// arriva qui già in chiaro, solo server-side. Mai loggarla.
//
// Verificato sul campo (probe locali del 28/07/2026, conto reale):
//  - la banca serve ~90 giorni di storico, qualunque date_from si chieda;
//  - entry_reference è univoco (1454/1454) → chiave di deduplicazione;
//  - bank_transaction_code.code è la causale ABI (16G, 480, 260, …) già
//    compatibile con la mappa tipologie di finanza-logic;
//  - remittance_information[0] è la stessa descrizione dell'export xlsx.

import type { MovimentoRaw } from "./finanza-logic";

const EB_API = "https://api.enablebanking.com";
// Deve coincidere ESATTAMENTE con un redirect URL registrato nell'app
// Enable Banking (Control Panel → Applications).
export const EB_REDIRECT_URL = "https://portal.drlogistica.it/";

export interface EbCredenziali {
  appId: string;
  /** Chiave privata PEM (PKCS#8, "BEGIN PRIVATE KEY"). */
  privateKeyPem: string;
}

// --- JWT RS256 con WebCrypto (niente node:crypto: gira su Cloudflare Worker) --

function b64urlBytes(b: Uint8Array): string {
  let s = "";
  for (const x of b) s += String.fromCharCode(x);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function b64urlJson(obj: unknown): string {
  return b64urlBytes(new TextEncoder().encode(JSON.stringify(obj)));
}

function pemToPkcs8(pem: string): Uint8Array {
  const body = pem.replace(/-----(BEGIN|END)[A-Z ]*KEY-----/g, "").replace(/\s+/g, "");
  if (!body) throw new Error("Chiave privata vuota o non in formato PEM.");
  const raw = atob(body);
  const out = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
  return out;
}

/** Importa la chiave PEM (PKCS#8). Usata anche come VALIDAZIONE al salvataggio. */
export async function ebImportaChiave(privateKeyPem: string): Promise<CryptoKey> {
  if (!/BEGIN PRIVATE KEY/.test(privateKeyPem))
    throw new Error(
      'Formato chiave non riconosciuto: attesa una chiave PKCS#8 ("-----BEGIN PRIVATE KEY-----").',
    );
  return crypto.subtle.importKey(
    "pkcs8",
    pemToPkcs8(privateKeyPem) as unknown as BufferSource,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );
}

async function ebJwt(cred: EbCredenziali): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const base = `${b64urlJson({ typ: "JWT", alg: "RS256", kid: cred.appId })}.${b64urlJson({
    iss: "enablebanking.com",
    aud: "api.enablebanking.com",
    iat: now,
    exp: now + 3600,
  })}`;
  const key = await ebImportaChiave(cred.privateKeyPem);
  const sig = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(base));
  return `${base}.${b64urlBytes(new Uint8Array(sig))}`;
}

async function ebCall<T>(
  cred: EbCredenziali,
  method: "GET" | "POST",
  path: string,
  body?: unknown,
): Promise<T> {
  const res = await fetch(`${EB_API}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${await ebJwt(cred)}`,
      Accept: "application/json",
      ...(body ? { "Content-Type": "application/json" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  if (!res.ok) {
    // Il corpo dell'errore Enable Banking è descrittivo e non contiene segreti.
    throw new Error(`Enable Banking ${method} ${path} → HTTP ${res.status}: ${text.slice(0, 300)}`);
  }
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error(`Enable Banking ${method} ${path}: risposta non JSON.`);
  }
}

// --- Flusso di collegamento ---------------------------------------------------

export interface EbConto {
  uid: string;
  iban: string;
  nome: string;
}

/** Anagrafica dell'applicazione registrata: prova di collegamento SENZA
 *  passare dalla banca (isola i problemi di app id / chiave privata). */
export async function ebApplicazione(
  cred: EbCredenziali,
): Promise<{ nome: string; ambiente: string; redirect: string[] }> {
  const app = await ebCall<{ name?: string; environment?: string; redirect_urls?: string[] }>(
    cred,
    "GET",
    "/application",
  );
  return {
    nome: String(app.name ?? cred.appId),
    ambiente: String(app.environment ?? "?"),
    redirect: app.redirect_urls ?? [],
  };
}

/** Avvia l'autorizzazione: restituisce l'URL della banca da aprire nel browser. */
export async function ebAvviaAuth(cred: EbCredenziali): Promise<{ url: string }> {
  const validUntil = new Date(Date.now() + 90 * 86400000).toISOString();
  const res = await ebCall<{ url: string }>(cred, "POST", "/auth", {
    access: { valid_until: validUntil },
    aspsp: { name: "Banco BPM", country: "IT" },
    state: crypto.randomUUID(),
    redirect_url: EB_REDIRECT_URL,
    psu_type: "business",
  });
  return { url: res.url };
}

interface EbSessionResponse {
  session_id?: string;
  accounts?: {
    uid?: string;
    account_id?: { iban?: string } | string;
    iban?: string;
    name?: string;
    product?: string;
  }[];
  access?: { valid_until?: string };
}

/** Completa il collegamento con il `code` della redirect: sessione + conti. */
export async function ebCreaSessione(
  cred: EbCredenziali,
  code: string,
): Promise<{ conti: EbConto[]; consensoScade: string }> {
  const s = await ebCall<EbSessionResponse>(cred, "POST", "/sessions", { code });
  const conti = (s.accounts ?? []).map((a): EbConto => {
    const iban = typeof a.account_id === "object" ? (a.account_id?.iban ?? "") : (a.iban ?? "");
    return {
      uid: String(a.uid ?? (typeof a.account_id === "string" ? a.account_id : "")),
      iban: String(iban ?? ""),
      nome: String(a.name ?? a.product ?? ""),
    };
  });
  const consensoScade = s.access?.valid_until ?? new Date(Date.now() + 89 * 86400000).toISOString();
  return { conti, consensoScade };
}

// --- Transazioni --------------------------------------------------------------

export interface EbTransazione {
  entry_reference?: string;
  transaction_amount?: { currency?: string; amount?: string };
  credit_debit_indicator?: string; // CRDT | DBIT
  status?: string; // BOOK | PDNG
  booking_date?: string;
  value_date?: string;
  bank_transaction_code?: { code?: string };
  remittance_information?: string[];
}

export async function ebTransazioni(
  cred: EbCredenziali,
  contoUid: string,
  dateFrom: string,
  continuationKey?: string,
): Promise<{ transazioni: EbTransazione[]; continuation: string | null }> {
  const params = new URLSearchParams({ date_from: dateFrom });
  if (continuationKey) params.set("continuation_key", continuationKey);
  const res = await ebCall<{ transactions?: EbTransazione[]; continuation_key?: string }>(
    cred,
    "GET",
    `/accounts/${encodeURIComponent(contoUid)}/transactions?${params.toString()}`,
  );
  return { transazioni: res.transactions ?? [], continuation: res.continuation_key ?? null };
}

/** Prefisso della chiave dei movimenti arrivati dall'API (Title in lista). */
export const EB_CHIAVE_PREFIX = "EB|";

/**
 * Converte una transazione contabilizzata (BOOK) nella coppia {raw, chiave}
 * dell'import movimenti. Restituisce null per le transazioni pendenti o senza
 * riferimento univoco (verranno riprese al giro successivo, una volta BOOK).
 */
export function ebMappaMovimento(t: EbTransazione): { raw: MovimentoRaw; chiave: string } | null {
  if (t.status !== "BOOK") return null;
  const ref = (t.entry_reference ?? "").trim();
  const booking = (t.booking_date ?? "").slice(0, 10);
  const amount = Number(t.transaction_amount?.amount ?? "");
  if (!ref || !booking || !Number.isFinite(amount)) return null;
  const segno = t.credit_debit_indicator === "DBIT" ? -1 : 1;
  return {
    chiave: `${EB_CHIAVE_PREFIX}${ref}`.slice(0, 250),
    raw: {
      dataContabile: booking,
      dataValuta: (t.value_date ?? booking).slice(0, 10),
      importo: Math.round(amount * 100 * segno) / 100,
      divisa: t.transaction_amount?.currency ?? "EUR",
      causale: (t.bank_transaction_code?.code ?? "").trim(),
      descrizione: (t.remittance_information ?? []).join(" ").trim(),
    },
  };
}
