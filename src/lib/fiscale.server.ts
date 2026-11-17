// DR Portal — Finanza → Flussi: storage dello scadenziario fiscale.
// Snapshot JSON unico nella libreria del sito (`FinanzaData/fiscale.json`),
// stesso pattern di stipendi.json: il file di Sabrina è un database
// cumulativo, quindi ogni import SOSTITUISCE tutto.

import * as XLSX from "xlsx";
import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
  TARGET_HOST,
} from "./sharepoint.server";
import {
  classificaScadenza,
  emptyFiscaleDb,
  importFiscaleSospetto,
  parseScadenzario,
  scegliFileScadenzario,
  unisciImportFiscale,
  type DaRateizzareFiscale,
  type FiscaleDb,
  type LetturaFiscale,
  type ScadenzaFiscale,
} from "./fiscale-logic";

const DB_PATH = "FinanzaData/fiscale.json";

interface DriveItemMeta {
  id?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function loadFiscaleDb(): Promise<FiscaleDb> {
  const cfg = await discoverSharePoint();
  let meta: DriveItemMeta;
  try {
    meta = await withDiscoveryRetry(() =>
      gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}?$select=id`),
    );
  } catch (err) {
    if (err instanceof SpHttpError && err.status === 404) return emptyFiscaleDb();
    throw err;
  }
  const metaFull = await withDiscoveryRetry(() =>
    gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/items/${meta.id}`),
  );
  const url = metaFull["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("fiscale.json: downloadUrl non disponibile da Graph.");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`fiscale.json: download fallito (${res.status}).`);
  const raw = (await res.json()) as Partial<FiscaleDb>;
  const db: FiscaleDb = {
    ...emptyFiscaleDb(),
    ...raw,
    scadenze: Array.isArray(raw.scadenze) ? raw.scadenze : [],
    daRateizzare: Array.isArray(raw.daRateizzare) ? raw.daRateizzare : [],
  };
  // Id anche in LETTURA (deterministici: dipendono solo dall'ordine nel
  // json): senza, un archivio scritto prima degli id arriverebbe al client
  // tutto senza matita/checkbox finché qualcuno non fa la prima scrittura.
  assegnaIdMancanti(db);
  return db;
}

/** Id progressivi stabili ("S1", "S2", …) per le modifiche dal portale. */
function assegnaIdMancanti(db: FiscaleDb): void {
  let max = 0;
  for (const s of db.scadenze) {
    const m = /^S(\d+)$/.exec(s.id ?? "");
    if (m) max = Math.max(max, Number(m[1]));
  }
  for (const s of db.scadenze) if (!s.id) s.id = `S${++max}`;
}

async function salvaFiscaleDb(db: FiscaleDb, utente: string): Promise<FiscaleDb> {
  const cfg = await discoverSharePoint();
  assegnaIdMancanti(db);
  db.versione = (db.versione ?? 0) + 1;
  db.aggiornatoIl = new Date().toISOString();
  db.aggiornatoDa = utente;
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(db),
    }),
  );
  logSp(
    "info",
    "fiscale.save",
    `Scadenziario fiscale salvato (v${db.versione}, ${db.scadenze.length} scadenze)`,
  );
  return db;
}

/** Import del file caricato a mano: sostituisce le righe DA FILE, conserva
 *  quelle create o modificate sul portale (vedi `unisciImportFiscale`). */
export async function replaceFiscale(
  scadenze: ScadenzaFiscale[],
  daRateizzare: DaRateizzareFiscale[],
  fonte: string,
  utente: string,
): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  return salvaFiscaleDb(unisciImportFiscale(db, scadenze, daRateizzare, fonte).db, utente);
}

// ---------------------------------------------------------------------------
// LETTURA AUTOMATICA DALLA CARTELLA (1.120.0, Simone 17/11/2026: "sì, fallo
// automatico"). Lo scadenzario di Sabrina sta sul sito SharePoint
// "DocumentiCondivisi", cartella "Fiscale" (lo stesso sito delle paghe e dei
// costi appalti). Il portale guarda la data di modifica del file: se è quella
// dell'ultima lettura non fa nulla; se è cambiata lo scarica e lo passa allo
// stesso lettore del caricamento a mano. Una lettura che non riesce non tocca
// l'archivio: restituisce l'archivio di prima con il motivo in `ultimaLettura`.
// ---------------------------------------------------------------------------
const SITO_DOCUMENTI = "DocumentiCondivisi";
export const CARTELLA_FISCALE = "Fiscale";

interface FiglioCartella {
  id: string;
  name: string;
  file?: unknown;
  lastModifiedDateTime?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function leggiFiscaleDaCartella(utente: string, forza = false): Promise<FiscaleDb> {
  const il = new Date().toISOString();
  const db = await loadFiscaleDb();
  // Senza scrivere: l'esito di questa lettura viaggia con la risposta.
  const comEra = (l: Omit<LetturaFiscale, "il" | "da">): FiscaleDb => ({
    ...db,
    ultimaLettura: { il, da: utente, ...l },
  });
  let scelto: FiglioCartella | undefined;
  try {
    const sito = await withDiscoveryRetry(() =>
      gatewayJson<{ id: string }>(`/sites/${TARGET_HOST}:/sites/${SITO_DOCUMENTI}`),
    );
    const figli = await withDiscoveryRetry(() =>
      gatewayJson<{ value: FiglioCartella[] }>(
        `/sites/${sito.id}/drive/root:/${encodeURIComponent(CARTELLA_FISCALE)}:/children?$select=id,name,file,lastModifiedDateTime&$top=200`,
      ),
    );
    scelto = scegliFileScadenzario((figli.value ?? []).filter((f) => f.file));
    if (!scelto)
      return comEra({
        esito: "assente",
        messaggio: `nella cartella ${CARTELLA_FISCALE} non c'è un file "SCADENZARIO FISCALE"`,
      });
    const modificato = scelto.lastModifiedDateTime ?? "";
    if (!forza && modificato && db.fonteModificato === modificato && db.fonteFile === scelto.name)
      return comEra({ esito: "uguale", file: scelto.name, modificato });
    const pieno = await withDiscoveryRetry(() =>
      gatewayJson<FiglioCartella>(`/sites/${sito.id}/drive/items/${scelto!.id}`),
    );
    const url = pieno["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("downloadUrl non disponibile da Graph");
    const res = await fetch(url);
    if (!res.ok) throw new Error(`download fallito (${res.status})`);
    const wb = XLSX.read(await res.arrayBuffer(), { cellDates: false });
    const letto = parseScadenzario(
      wb.SheetNames.map((nome) => ({
        nome,
        matrix: XLSX.utils.sheet_to_json(wb.Sheets[nome], {
          header: 1,
          raw: true,
          defval: null,
        }) as unknown[][],
      })),
    );
    if (!letto || !letto.scadenze.length)
      throw new Error("nessun foglio con la tabella delle scadenze");
    const sospetto = importFiscaleSospetto(db, letto.scadenze);
    if (sospetto) throw new Error(sospetto);
    const { db: unito, cambiate } = unisciImportFiscale(
      db,
      letto.scadenze,
      letto.daRateizzare,
      scelto.name,
    );
    const salvato = await salvaFiscaleDb(
      {
        ...unito,
        fonteModificato: modificato,
        ultimaLettura: {
          il,
          da: utente,
          esito: "letto",
          file: scelto.name,
          modificato,
          cambiate,
          messaggio: `foglio ${letto.foglio}`,
        },
      },
      utente,
    );
    logSp(
      "info",
      "fiscale.cartella",
      `Scadenziario letto dalla cartella ${CARTELLA_FISCALE}: ${scelto.name} (${letto.scadenze.length} scadenze, ${cambiate} cambiate)`,
    );
    return salvato;
  } catch (e) {
    const messaggio = e instanceof Error ? e.message : String(e);
    logSp(
      "warn",
      "fiscale.cartella",
      `Lettura dalla cartella ${CARTELLA_FISCALE} fallita: ${messaggio}`,
    );
    return comEra({
      esito: "errore",
      file: scelto?.name,
      modificato: scelto?.lastModifiedDateTime,
      messaggio,
    });
  }
}

/** Crea o aggiorna una scadenza dal portale; la categoria si ricalcola. */
export async function upsertScadenzaFiscale(
  scadenza: ScadenzaFiscale,
  utente: string,
): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  scadenza.origine = "portale";
  scadenza.categoria = classificaScadenza(
    scadenza.voce,
    scadenza.anno,
    scadenza.dataPagamento,
    scadenza.modalita ?? "",
    scadenza.periodo ?? "",
  );
  if (scadenza.id) {
    const i = db.scadenze.findIndex((s) => s.id === scadenza.id);
    // Id sconosciuto = il client guarda un archivio superato (import o
    // modifica altrui nel frattempo): MAI pushare una copia — errore chiaro.
    if (i < 0)
      throw new Error(
        "Scadenza non trovata: l'archivio è cambiato nel frattempo. Ricarica la pagina e riprova.",
      );
    const prima = db.scadenze[i];
    // Il form non rimanda i campi che non mostra: si conservano.
    scadenza.chiaveFile = prima.chiaveFile;
    if (scadenza.voceOld == null) scadenza.voceOld = prima.voceOld;
    db.scadenze[i] = scadenza;
  } else {
    db.scadenze.push(scadenza);
  }
  return salvaFiscaleDb(db, utente);
}

/** Sostituisce l'elenco "da registrare in futuro" (importi noti senza data,
 *  fuori dal cash flow) — compilabile dal portale come il resto. */
export async function setDaRateizzareFiscale(
  lista: DaRateizzareFiscale[],
  utente: string,
): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  db.daRateizzare = lista;
  db.daRateizzarePortale = true;
  return salvaFiscaleDb(db, utente);
}

export async function eliminaScadenzaFiscale(id: string, utente: string): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  const riga = db.scadenze.find((s) => s.id === id);
  if (riga?.chiaveFile && !(db.tombstones ?? []).includes(riga.chiaveFile))
    db.tombstones = [...(db.tombstones ?? []), riga.chiaveFile];
  db.scadenze = db.scadenze.filter((s) => s.id !== id);
  return salvaFiscaleDb(db, utente);
}
