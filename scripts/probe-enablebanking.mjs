// DR Portal — PROBE Enable Banking (PSD2/open banking) — SOLO LETTURA.
// -----------------------------------------------------------------------------
// Si lancia in locale (Node >= 18), NON dal portale: serve a verificare, prima
// di scrivere codice nel Worker, che il collegamento al conto Banco BPM
// funzioni e che i dati abbiano la forma attesa.
//
// Prerequisiti (una tantum, dal Control Panel enablebanking.com):
//   1. account su https://enablebanking.com/sign-in/ (si crea al primo accesso)
//   2. registrare una API application; il BROWSER genera e scarica la chiave
//      privata `<app-id>.pem` (custodirla: non va mai condivisa né committata)
//   3. tra i redirect URL dell'app inserire: https://portal.drlogistica.it/
//
// Uso:
//   node scripts/probe-enablebanking.mjs --app <application-id> --key <percorso .pem>
//
// Flusso: JWT firmato RS256 → elenco banche IT → avvio autorizzazione (URL da
// aprire nel browser; SCA con l'app YouApp di Banco BPM) → incolli il `code`
// della redirect → sessione → conti → prima pagina di transazioni.
// Nessun dato viene inviato altrove; il JSON grezzo resta sul tuo PC.

import { readFileSync, writeFileSync } from "node:fs";
import { createSign, randomUUID } from "node:crypto";
import { createInterface } from "node:readline/promises";

const API = "https://api.enablebanking.com";

function arg(name, fallback = null) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}
const APP_ID = arg("app");
const KEY_PATH = arg("key");
const COUNTRY = arg("country", "IT");
if (!APP_ID || !KEY_PATH) {
  console.error("Uso: node scripts/probe-enablebanking.mjs --app <application-id> --key <file.pem>");
  process.exit(1);
}
const PRIVATE_KEY = readFileSync(KEY_PATH, "utf-8");

function b64url(objOrBuf) {
  const buf = Buffer.isBuffer(objOrBuf) ? objOrBuf : Buffer.from(JSON.stringify(objOrBuf));
  return buf.toString("base64url");
}
function jwt() {
  const now = Math.floor(Date.now() / 1000);
  const header = { typ: "JWT", alg: "RS256", kid: APP_ID };
  const payload = { iss: "enablebanking.com", aud: "api.enablebanking.com", iat: now, exp: now + 3600 };
  const base = `${b64url(header)}.${b64url(payload)}`;
  const signer = createSign("RSA-SHA256");
  signer.update(base);
  return `${base}.${signer.sign(PRIVATE_KEY).toString("base64url")}`;
}

async function call(method, path, body) {
  const res = await fetch(`${API}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${jwt()}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch {
    /* corpo non JSON */
  }
  if (!res.ok) {
    // Il corpo integrale dell'errore È l'informazione diagnostica che cerchiamo.
    console.error(`\n✗ ${method} ${path} → HTTP ${res.status}`);
    console.error(text.slice(0, 1500));
    process.exit(1);
  }
  return json;
}

const rl = createInterface({ input: process.stdin, output: process.stdout });

console.log("— PROBE Enable Banking (sola lettura) —\n");

// 1) verifica applicazione
const app = await call("GET", "/application");
console.log(`Applicazione: ${app.name ?? APP_ID} · ambiente: ${app.environment ?? "?"} · redirect: ${(app.redirect_urls ?? []).join(", ")}`);

// 2) banche disponibili in Italia
const aspsps = await call("GET", `/aspsps?country=${COUNTRY}`);
const lista = aspsps.aspsps ?? aspsps ?? [];
const bpm = lista.filter((a) => /bpm/i.test(a.name));
console.log(`\nBanche ${COUNTRY} disponibili: ${lista.length}. Corrispondenze "BPM":`);
for (const a of bpm) console.log(`  - ${a.name} (${a.country})${a.psu_types ? ` · tipi utente: ${a.psu_types.join("/")}` : ""}`);
if (!bpm.length) {
  console.log("  (nessuna — elenco completo nei primi 30 nomi:)");
  for (const a of lista.slice(0, 30)) console.log("  -", a.name);
}

const nomeBanca = (await rl.question(`\nNome esatto della banca da collegare [${bpm[0]?.name ?? "Banco BPM"}]: `)).trim() || bpm[0]?.name || "Banco BPM";
const psuType = (await rl.question("Tipo utenza (business/personal) [business]: ")).trim() || "business";
const redirect = (app.redirect_urls ?? [])[0] ?? "https://portal.drlogistica.it/";

// 3) avvio autorizzazione (consenso: 90 giorni)
const validUntil = new Date(Date.now() + 90 * 86400000).toISOString();
const auth = await call("POST", "/auth", {
  access: { valid_until: validUntil },
  aspsp: { name: nomeBanca, country: COUNTRY },
  state: randomUUID(),
  redirect_url: redirect,
  psu_type: psuType,
});
console.log("\n➜ Apri questo URL nel browser e autorizza con l'app della banca (YouApp):\n");
console.log(auth.url);
console.log(`\nA fine autorizzazione verrai rimandato su ${redirect}?code=...`);
const code = (await rl.question("\nIncolla qui il valore del parametro 'code': ")).trim();

// 4) sessione + conti
const session = await call("POST", "/sessions", { code });
console.log(`\nSessione creata. Conti autorizzati: ${session.accounts?.length ?? 0}`);
(session.accounts ?? []).forEach((a, i) =>
  console.log(`  [${i}] uid=${a.uid ?? a.account_id ?? "?"} · ${a.account_id?.iban ?? a.iban ?? ""} · ${a.name ?? ""}`),
);
const idx = Number((await rl.question("Indice del conto da leggere [0]: ")).trim() || "0");
const conto = (session.accounts ?? [])[idx];
const uid = conto?.uid ?? conto?.account_id;

// 5) transazioni (prima pagina)
const tx = await call("GET", `/accounts/${uid}/transactions`);
const rows = tx.transactions ?? [];
console.log(`\nTransazioni ricevute: ${rows.length}${tx.continuation_key ? " (altre pagine disponibili)" : ""}`);
if (rows[0]) {
  console.log("\nCAMPI della prima transazione (per la mappatura):");
  console.log(Object.keys(rows[0]).join(", "));
  console.log("\nCampione (3 righe):");
  for (const r of rows.slice(0, 3)) console.log(JSON.stringify(r).slice(0, 300));
}
writeFileSync("enablebanking-probe.json", JSON.stringify({ conti: session.accounts, transazioni: rows.slice(0, 50) }, null, 2));
console.log("\nJSON grezzo salvato in enablebanking-probe.json (resta sul tuo PC).");
rl.close();
