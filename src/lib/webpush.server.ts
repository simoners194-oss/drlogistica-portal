// DR Portal — Web Push (server-only, zero dipendenze).
// -----------------------------------------------------------------------------
// Implementazione del protocollo Web Push con la sola Web Crypto API, quindi
// compatibile con Cloudflare Workers (dove gira il server):
//   - VAPID (RFC 8292): JWT ES256 firmato con la chiave privata applicativa.
//   - Cifratura payload (RFC 8291) con content coding aes128gcm (RFC 8188).
// Nessuna libreria esterna: tutto via crypto.subtle.
// -----------------------------------------------------------------------------

// --- base64url helpers -------------------------------------------------------
function b64urlEncode(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function b64urlDecode(b64: string): Uint8Array {
  const pad = b64.length % 4 === 0 ? "" : "=".repeat(4 - (b64.length % 4));
  const s = atob(b64.replace(/-/g, "+").replace(/_/g, "/") + pad);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
function utf8(s: string): Uint8Array {
  return new TextEncoder().encode(s);
}
function concatBytes(...parts: Uint8Array[]): Uint8Array {
  const len = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(len);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
// Vista ArrayBuffer "pura" dei byte (evita problemi di tipo con SharedArrayBuffer).
function ab(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

// --- VAPID keys --------------------------------------------------------------
export interface VapidKeys {
  /** Chiave pubblica P-256 in formato raw (65 byte) base64url — va al client. */
  publicKey: string;
  /** Chiave privata come JWK JSON (resta sul server). */
  privateJwk: string;
}

export async function generateVapidKeys(): Promise<VapidKeys> {
  const pair = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, true, [
    "sign",
  ]);
  const rawPub = new Uint8Array(await crypto.subtle.exportKey("raw", pair.publicKey));
  const jwkPriv = await crypto.subtle.exportKey("jwk", pair.privateKey);
  return { publicKey: b64urlEncode(rawPub), privateJwk: JSON.stringify(jwkPriv) };
}

// --- VAPID JWT (ES256) -------------------------------------------------------
async function vapidAuthHeader(
  endpoint: string,
  keys: VapidKeys,
  subject: string,
): Promise<string> {
  const aud = new URL(endpoint).origin;
  const header = b64urlEncode(utf8(JSON.stringify({ typ: "JWT", alg: "ES256" })));
  const claims = b64urlEncode(
    utf8(
      JSON.stringify({
        aud,
        exp: Math.floor(Date.now() / 1000) + 12 * 60 * 60,
        sub: subject,
      }),
    ),
  );
  const signingInput = `${header}.${claims}`;
  const privateKey = await crypto.subtle.importKey(
    "jwk",
    JSON.parse(keys.privateJwk) as JsonWebKey,
    { name: "ECDSA", namedCurve: "P-256" },
    false,
    ["sign"],
  );
  // Web Crypto restituisce la firma ECDSA già in formato r||s (64 byte): è
  // esattamente il formato JWS richiesto.
  const sig = new Uint8Array(
    await crypto.subtle.sign(
      { name: "ECDSA", hash: "SHA-256" },
      privateKey,
      ab(utf8(signingInput)),
    ),
  );
  return `vapid t=${signingInput}.${b64urlEncode(sig)}, k=${keys.publicKey}`;
}

// --- HKDF --------------------------------------------------------------------
async function hkdf(
  salt: Uint8Array,
  ikm: Uint8Array,
  info: Uint8Array,
  length: number,
): Promise<Uint8Array> {
  const key = await crypto.subtle.importKey("raw", ab(ikm), "HKDF", false, ["deriveBits"]);
  const bits = await crypto.subtle.deriveBits(
    { name: "HKDF", hash: "SHA-256", salt: ab(salt), info: ab(info) },
    key,
    length * 8,
  );
  return new Uint8Array(bits);
}

// --- RFC 8291: cifratura del payload (aes128gcm) ------------------------------
async function encryptPayload(
  p256dhB64: string,
  authB64: string,
  payload: Uint8Array,
): Promise<Uint8Array> {
  const uaPublicRaw = b64urlDecode(p256dhB64); // 65 byte, raw P-256
  const authSecret = b64urlDecode(authB64); // 16 byte

  // Coppia effimera del server applicativo (ECDH).
  const asPair = await crypto.subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, true, [
    "deriveBits",
  ]);
  const asPublicRaw = new Uint8Array(await crypto.subtle.exportKey("raw", asPair.publicKey));
  const uaPublicKey = await crypto.subtle.importKey(
    "raw",
    ab(uaPublicRaw),
    { name: "ECDH", namedCurve: "P-256" },
    false,
    [],
  );
  const ecdhSecret = new Uint8Array(
    await crypto.subtle.deriveBits({ name: "ECDH", public: uaPublicKey }, asPair.privateKey, 256),
  );

  // IKM = HKDF(auth_secret, ecdh_secret, "WebPush: info" || 0x00 || ua_pub || as_pub, 32)
  const keyInfo = concatBytes(utf8("WebPush: info"), new Uint8Array([0]), uaPublicRaw, asPublicRaw);
  const ikm = await hkdf(authSecret, ecdhSecret, keyInfo, 32);

  const salt = crypto.getRandomValues(new Uint8Array(16));
  const cek = await hkdf(
    salt,
    ikm,
    concatBytes(utf8("Content-Encoding: aes128gcm"), new Uint8Array([0])),
    16,
  );
  const nonce = await hkdf(
    salt,
    ikm,
    concatBytes(utf8("Content-Encoding: nonce"), new Uint8Array([0])),
    12,
  );

  // Record singolo: payload || 0x02 (delimitatore ultimo record, RFC 8188).
  const record = concatBytes(payload, new Uint8Array([2]));
  const aesKey = await crypto.subtle.importKey("raw", ab(cek), "AES-GCM", false, ["encrypt"]);
  const ciphertext = new Uint8Array(
    await crypto.subtle.encrypt({ name: "AES-GCM", iv: ab(nonce) }, aesKey, ab(record)),
  );

  // Header aes128gcm: salt(16) || rs(4) || idlen(1) || keyid(as_public, 65).
  const rs = new Uint8Array(4);
  new DataView(rs.buffer).setUint32(0, 4096);
  const idlen = new Uint8Array([asPublicRaw.length]);
  return concatBytes(salt, rs, idlen, asPublicRaw, ciphertext);
}

// --- Invio -------------------------------------------------------------------
export interface PushSubscriptionData {
  endpoint: string;
  p256dh: string;
  auth: string;
}
export interface PushSendResult {
  ok: boolean;
  status: number;
  /** true se la subscription è morta (404/410) e va eliminata. */
  gone: boolean;
}

export async function sendWebPush(
  sub: PushSubscriptionData,
  payloadJson: unknown,
  keys: VapidKeys,
  subject = "mailto:portal@drlogistica.it",
): Promise<PushSendResult> {
  const body = await encryptPayload(sub.p256dh, sub.auth, utf8(JSON.stringify(payloadJson)));
  const auth = await vapidAuthHeader(sub.endpoint, keys, subject);
  const res = await fetch(sub.endpoint, {
    method: "POST",
    headers: {
      Authorization: auth,
      "Content-Encoding": "aes128gcm",
      "Content-Type": "application/octet-stream",
      TTL: "86400",
      Urgency: "normal",
    },
    body: ab(body) as BodyInit,
  });
  // Le push service rispondono 201 Created; 404/410 = subscription scaduta.
  return {
    ok: res.status >= 200 && res.status < 300,
    status: res.status,
    gone: res.status === 404 || res.status === 410,
  };
}
