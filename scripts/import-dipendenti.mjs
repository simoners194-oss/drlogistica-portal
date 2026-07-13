// DR Portal — Import massivo Dipendenti da CSV su SharePoint.
// -----------------------------------------------------------------------------
// Carica in blocco righe di dipendenti nella lista SharePoint "Dipendenti"
// passando dallo STESSO gateway del portale (Lovable Connector → Microsoft
// Graph). Utile per volumi grandi o import ripetibili; per pochi record
// conviene "Modifica in visualizzazione griglia" + incolla da Excel.
//
// USO
//   node scripts/import-dipendenti.mjs <file.csv> [--dry-run]
//
// CREDENZIALI (variabili d'ambiente, le stesse del server — prendile dai
// secret del progetto Lovable):
//   LOVABLE_API_KEY=...
//   MICROSOFT_SHAREPOINT_API_KEY=...
//
// FORMATO CSV (intestazione sulla prima riga; i nomi devono combaciare con i
// display name delle colonne su SharePoint). Esempio:
//   Nome,Cognome,Email,Sede,Attivo,Responsabile,Codice,PIN,Visibile,Autorizza,Operatore,OreSettimanali
//   Mario,Rossi,mario@dr.it,Fiano Romano,Sì,Dipendente,DR010,1234,Sì,No,No,40
//   Lucia,Verdi,lucia@dr.it,San Giuliano,Sì,Dipendente,DR011,5678,Sì,No,No,16
//
// NOTE
// - Le colonne booleane (Attivo/Visibile/Autorizza/Operatore) accettano
//   Sì/No/true/false/1/0. OreSettimanali è numerico. Il resto è testo.
// - `--dry-run` mostra cosa verrebbe creato SENZA scrivere nulla.
// - Non aggiorna record esistenti: crea sempre nuovi item (attenzione ai
//   doppioni sul Codice).
// -----------------------------------------------------------------------------

import { readFileSync } from "node:fs";

const GATEWAY = "https://connector-gateway.lovable.dev/microsoft_sharepoint";
const HOST = "drlogisticaroma.sharepoint.com";
const SITE_PATH = "DRPORTAL";

const BOOL_COLS = new Set(["Attivo", "Visibile", "Autorizza", "Operatore"]);
const NUM_COLS = new Set(["OreSettimanali"]);

const LOVABLE_KEY = process.env.LOVABLE_API_KEY;
const SP_KEY = process.env.MICROSOFT_SHAREPOINT_API_KEY;

const [csvPath, ...flags] = process.argv.slice(2);
const dryRun = flags.includes("--dry-run");

function fail(msg) {
  console.error(`\n❌ ${msg}\n`);
  process.exit(1);
}

if (!csvPath) fail("Manca il percorso del CSV. Uso: node scripts/import-dipendenti.mjs <file.csv> [--dry-run]");
if (!LOVABLE_KEY || !SP_KEY) fail("Servono le env LOVABLE_API_KEY e MICROSOFT_SHAREPOINT_API_KEY.");

async function gj(path, init = {}) {
  const headers = {
    Authorization: `Bearer ${LOVABLE_KEY}`,
    "X-Connection-Api-Key": SP_KEY,
    ...(init.body ? { "Content-Type": "application/json" } : {}),
  };
  const res = await fetch(`${GATEWAY}${path}`, { ...init, headers });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`${init.method ?? "GET"} ${path.split("?")[0]} → ${res.status} ${body.slice(0, 300)}`);
  }
  return res.json();
}

// Parser CSV minimale con supporto ai campi tra virgolette.
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') {
        inQ = false;
      } else field += c;
    } else if (c === '"') inQ = true;
    else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\r") {
      /* skip */
    } else if (c === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += c;
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((r) => r.some((x) => x.trim() !== ""));
}

function coerce(col, raw) {
  const v = (raw ?? "").trim();
  if (BOOL_COLS.has(col)) return /^(s[iì]|true|1|x)$/i.test(v);
  if (NUM_COLS.has(col)) {
    if (v === "") return undefined;
    const n = Number(v.replace(",", "."));
    return Number.isFinite(n) ? n : undefined;
  }
  return v === "" ? undefined : v;
}

async function main() {
  console.log(`\n📂 Leggo ${csvPath}${dryRun ? "  (DRY-RUN)" : ""}`);
  const rows = parseCsv(readFileSync(csvPath, "utf8"));
  if (rows.length < 2) fail("Il CSV non ha righe di dati (serve intestazione + almeno una riga).");
  const header = rows[0].map((h) => h.trim());
  const records = rows.slice(1);

  console.log("🔎 Discovery sito/lista/colonne…");
  const site = await gj(`/sites/${HOST}:/sites/${SITE_PATH}`);
  const listsRes = await gj(`/sites/${site.id}/lists?$select=id,name,displayName`);
  const dip = (listsRes.value ?? []).find((l) =>
    /dipendent/i.test(l.displayName || l.name || ""),
  );
  if (!dip) fail('Lista "Dipendenti" non trovata sul sito.');
  const colsRes = await gj(
    `/sites/${site.id}/lists/${dip.id}/columns?$select=name,displayName,readOnly,hidden`,
  );
  const internalByDisplay = new Map();
  for (const c of colsRes.value ?? []) {
    if (c.hidden || c.readOnly) continue;
    if (c.displayName) internalByDisplay.set(c.displayName.toLowerCase(), c.name);
    if (c.name) internalByDisplay.set(c.name.toLowerCase(), c.name);
  }

  // Mappa colonna CSV -> internalName SharePoint.
  const missing = header.filter((h) => !internalByDisplay.has(h.toLowerCase()));
  if (missing.length) fail(`Colonne CSV non trovate su SharePoint: ${missing.join(", ")}`);

  let ok = 0;
  let ko = 0;
  for (let r = 0; r < records.length; r++) {
    const row = records[r];
    const fields = {};
    header.forEach((h, i) => {
      const val = coerce(h, row[i]);
      if (val !== undefined) fields[internalByDisplay.get(h.toLowerCase())] = val;
    });
    const label = `${fields.Codice ?? ""} ${fields.Cognome ?? ""} ${fields.Nome ?? ""}`.trim();
    if (dryRun) {
      console.log(`  • [dry] ${label}: ${JSON.stringify(fields)}`);
      ok++;
      continue;
    }
    try {
      await gj(`/sites/${site.id}/lists/${dip.id}/items`, {
        method: "POST",
        body: JSON.stringify({ fields }),
      });
      console.log(`  ✅ ${label}`);
      ok++;
    } catch (err) {
      console.error(`  ❌ ${label}: ${err.message}`);
      ko++;
    }
  }
  console.log(`\nFatto. Creati: ${ok}${ko ? ` · Errori: ${ko}` : ""}${dryRun ? " (nessuna scrittura)" : ""}\n`);
}

main().catch((err) => fail(err.message));
