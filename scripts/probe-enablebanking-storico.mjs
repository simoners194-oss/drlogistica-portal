// DR Portal — PROBE 2 Enable Banking: PROFONDITÀ DELLO STORICO — SOLO LETTURA.
// -----------------------------------------------------------------------------
// Verifica fin dove Banco BPM serve le transazioni via PSD2 (obiettivo: 2022).
// Riusa la sessione già autorizzata (89 giorni): legge l'uid del conto da
// enablebanking-probe.json, nessuna nuova SCA richiesta.
//
// Uso (dalla cartella del repo):
//   node scripts/probe-enablebanking-storico.mjs --app <application-id> --key <percorso .pem>
//
// Opzionale: --uid <uid conto> per scegliere un conto diverso da quello …3681.
// Scarica TUTTE le pagine dal 2022-01-01 (o dalla data più antica concessa),
// stampa il conteggio per anno e salva il grezzo in enablebanking-storico.json
// (file locale, già nel .gitignore).

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { createSign } from "node:crypto";

const API = "https://api.enablebanking.com";

function arg(name, fallback = null) {
  const i = process.argv.indexOf(`--${name}`);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}
const APP_ID = arg("app");
const KEY_PATH = arg("key");
if (!APP_ID || !KEY_PATH) {
  console.error("Uso: node scripts/probe-enablebanking-storico.mjs --app <application-id> --key <file.pem>");
  process.exit(1);
}
const PRIVATE_KEY = readFileSync(KEY_PATH, "utf-8");

function b64url(objOrBuf) {
  const buf = Buffer.isBuffer(objOrBuf) ? objOrBuf : Buffer.from(JSON.stringify(objOrBuf));
  return buf.toString("base64url");
}
function jwt() {
  const now = Math.floor(Date.now() / 1000);
  const base = `${b64url({ typ: "JWT", alg: "RS256", kid: APP_ID })}.${b64url({
    iss: "enablebanking.com",
    aud: "api.enablebanking.com",
    iat: now,
    exp: now + 3600,
  })}`;
  const signer = createSign("RSA-SHA256");
  signer.update(base);
  return `${base}.${signer.sign(PRIVATE_KEY).toString("base64url")}`;
}

async function call(path) {
  const res = await fetch(`${API}${path}`, {
    headers: { Authorization: `Bearer ${jwt()}`, Accept: "application/json" },
  });
  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch {
    /* corpo non JSON */
  }
  return { status: res.status, ok: res.ok, json, text };
}

// Conto: da --uid, altrimenti quello con IBAN che finisce per 3681 dal probe 1.
let uid = arg("uid");
if (!uid) {
  if (!existsSync("enablebanking-probe.json")) {
    console.error("enablebanking-probe.json non trovato: lancia dalla cartella del repo, o passa --uid <uid>.");
    process.exit(1);
  }
  const conti = JSON.parse(readFileSync("enablebanking-probe.json", "utf-8")).conti ?? [];
  const principale = conti.find((c) => String(c.account_id?.iban ?? c.iban ?? "").endsWith("3681")) ?? conti[0];
  uid = principale?.uid ?? principale?.account_id;
  console.log(`Conto: ${principale?.account_id?.iban ?? principale?.iban ?? "?"} (uid ${uid})`);
}

console.log("— PROBE storico Enable Banking (sola lettura) —\n");

// Trova la data di partenza più antica accettata: prova 2022, poi via via più vicino.
const TENTATIVI = ["2022-01-01", "2023-01-01", "2024-01-01", "2025-01-01", null]; // null = default banca
let dateFrom = null;
let prima = null;
for (const d of TENTATIVI) {
  const q = d ? `?date_from=${d}` : "";
  process.stdout.write(`Provo date_from=${d ?? "(default banca)"} … `);
  const r = await call(`/accounts/${uid}/transactions${q}`);
  if (r.ok) {
    console.log(`OK (${r.json.transactions?.length ?? 0} nella prima pagina)`);
    dateFrom = d;
    prima = r.json;
    break;
  }
  console.log(`HTTP ${r.status}`);
  console.log(`   ↳ ${r.text.slice(0, 400)}`);
}
if (!prima) {
  console.error("\nNessuna variante accettata: la sessione potrebbe essere scaduta (rifare il probe 1).");
  process.exit(1);
}

// Scarica tutte le pagine.
const tutte = [...(prima.transactions ?? [])];
let cont = prima.continuation_key ?? null;
let pagina = 1;
while (cont && pagina < 400) {
  pagina++;
  const q = `${dateFrom ? `date_from=${dateFrom}&` : ""}continuation_key=${encodeURIComponent(cont)}`;
  const r = await call(`/accounts/${uid}/transactions?${q}`);
  if (!r.ok) {
    console.error(`\n✗ pagina ${pagina}: HTTP ${r.status}\n${r.text.slice(0, 400)}`);
    break;
  }
  tutte.push(...(r.json.transactions ?? []));
  cont = r.json.continuation_key ?? null;
  process.stdout.write(`\rPagine scaricate: ${pagina} · transazioni: ${tutte.length}   `);
}
console.log(`\n\nTotale transazioni: ${tutte.length} (richieste da ${dateFrom ?? "default banca"})`);

// Conteggio per anno + estremi.
const perAnno = {};
let min = null;
let max = null;
for (const t of tutte) {
  const d = t.booking_date ?? t.value_date ?? "";
  const anno = d.slice(0, 4) || "?";
  perAnno[anno] = (perAnno[anno] ?? 0) + 1;
  if (!min || d < min) min = d;
  if (!max || d > max) max = d;
}
console.log(`Periodo coperto: ${min} → ${max}\n\nMovimenti per anno:`);
for (const [anno, n] of Object.entries(perAnno).sort()) console.log(`  ${anno}: ${n}`);

writeFileSync(
  "enablebanking-storico.json",
  JSON.stringify({ uid, dateFrom, totale: tutte.length, perAnno, min, max, transazioni: tutte }, null, 2),
);
console.log("\nJSON grezzo salvato in enablebanking-storico.json (resta sul tuo PC).");
