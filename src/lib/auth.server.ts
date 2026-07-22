// DR Portal — Sessione server firmata (server-only).
// -----------------------------------------------------------------------------
// Emette e verifica un cookie di sessione httpOnly FIRMATO (HMAC-SHA256), così
// l'identità dell'utente vive sul server e non è più falsificabile dal client.
// Il segreto di firma è in `process.env.SESSION_SECRET` (da impostare nei
// secret del progetto). Se manca, la sessione server è disattivata in modo
// sicuro: setSessionCookie non scrive nulla e readSessionUser ritorna null
// (l'app continua a funzionare col comportamento precedente).
// -----------------------------------------------------------------------------

import { getCookie, setCookie, deleteCookie } from "@tanstack/react-start/server";
import type { Ruolo, SessionSede } from "./session";

const COOKIE_NAME = "dr:session";
// 30 giorni: la sessione deve sopravvivere alla chiusura dell'app sul telefono
// (tap su notifica push → dentro senza login). Firma HMAC + httpOnly restano.
const MAX_AGE_S = 60 * 60 * 24 * 30;

export interface ServerSessionUser {
  id: string;
  nome: string;
  cognome: string;
  sede: SessionSede;
  ruolo: Ruolo;
  autorizza: boolean;
  operatore: boolean;
  // Codice dipendente (es. DR005) — usato per i moduli riservati al direttore.
  // Sessioni emesse prima dell'introduzione del campo → undefined (re-login).
  codice?: string;
}

// --- base64url su byte/stringhe (unicode-safe) ---
function bytesToB64url(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function strToB64url(str: string): string {
  return bytesToB64url(new TextEncoder().encode(str));
}
function b64urlToStr(b64: string): string {
  const s = atob(b64.replace(/-/g, "+").replace(/_/g, "/"));
  const bytes = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

async function hmac(secret: string, data: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(data));
  return bytesToB64url(new Uint8Array(sig));
}

// Confronto a tempo costante (evita timing attack sulla firma).
function safeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// Segreto di firma. Preferisce SESSION_SECRET (dedicato); se assente, ricade su
// un segreto server-only già iniettato dal connettore (mai esposto al client,
// stabile). Così la sessione funziona senza configurazione manuale; per la
// produzione conviene comunque impostare un SESSION_SECRET dedicato.
function getSecret(): string | null {
  const s =
    process.env.SESSION_SECRET ||
    process.env.MICROSOFT_SHAREPOINT_API_KEY ||
    process.env.LOVABLE_API_KEY;
  return s && s.length >= 16 ? s : null;
}

async function sign(user: ServerSessionUser): Promise<string | null> {
  const secret = getSecret();
  if (!secret) return null;
  const payload = strToB64url(JSON.stringify({ u: user, exp: Date.now() + MAX_AGE_S * 1000 }));
  const sig = await hmac(secret, payload);
  return `${payload}.${sig}`;
}

async function verify(token: string): Promise<ServerSessionUser | null> {
  const secret = getSecret();
  if (!secret) return null;
  const [payload, sig] = token.split(".");
  if (!payload || !sig) return null;
  const expected = await hmac(secret, payload);
  if (!safeEqual(sig, expected)) return null;
  try {
    const data = JSON.parse(b64urlToStr(payload)) as { u?: ServerSessionUser; exp?: number };
    if (!data.exp || data.exp < Date.now() || !data.u?.id) return null;
    return data.u;
  } catch {
    return null;
  }
}

// Emette il cookie di sessione firmato (no-op se manca il segreto).
export async function setSessionCookie(user: ServerSessionUser): Promise<boolean> {
  const token = await sign(user);
  if (!token) return false;
  setCookie(COOKIE_NAME, token, {
    httpOnly: true,
    secure: true,
    sameSite: "lax",
    path: "/",
    maxAge: MAX_AGE_S,
  });
  return true;
}

// Legge e verifica la sessione dal cookie della richiesta corrente.
export async function readSessionUser(): Promise<ServerSessionUser | null> {
  const token = getCookie(COOKIE_NAME);
  if (!token) return null;
  return verify(token);
}

export function clearSessionCookie(): void {
  deleteCookie(COOKIE_NAME, { path: "/" });
}

// Indica se il segreto di firma è configurato (per diagnostica).
export function sessionSecretConfigured(): boolean {
  return getSecret() != null;
}
