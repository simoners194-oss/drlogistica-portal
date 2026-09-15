// DR Portal — Finanza → Flussi: storage dello scadenziario fiscale.
// Snapshot JSON unico nella libreria del sito (`FinanzaData/fiscale.json`),
// stesso pattern di stipendi.json: il file di Sabrina è un database
// cumulativo, quindi ogni import SOSTITUISCE tutto.

import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
} from "./sharepoint.server";
import {
  chiaveScadenzaFile,
  classificaScadenza,
  emptyFiscaleDb,
  type DaRateizzareFiscale,
  type FiscaleDb,
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

/** Import del file: sostituisce le righe DA FILE, conserva quelle create o
 *  modificate sul portale (origine "portale") — l'Excel è in dismissione.
 *  La chiaveFile evita i doppioni: una riga del file già modificata sul
 *  portale NON viene reimportata (vince la versione portale), una eliminata
 *  dal portale NON risorge (tombstone). Gli id delle righe da file restano
 *  STABILI tra un import e l'altro (match per chiaveFile). LIMITE NOTO: se
 *  la STESSA riga viene cambiata sia sul portale sia nell'Excel, la chiave
 *  diverge e compaiono entrambe — da sanare a mano (transizione breve). */
export async function replaceFiscale(
  scadenze: ScadenzaFiscale[],
  daRateizzare: DaRateizzareFiscale[],
  fonte: string,
  utente: string,
): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  const portale = db.scadenze.filter((s) => s.origine === "portale");
  const chiaviPortale = new Set(portale.map((s) => s.chiaveFile).filter(Boolean));
  const tombstones = new Set(db.tombstones ?? []);
  const idVecchi = new Map(
    db.scadenze
      .filter((s) => s.origine !== "portale")
      .map((s) => [s.chiaveFile ?? chiaveScadenzaFile(s), s.id] as const),
  );
  const nuove: ScadenzaFiscale[] = [];
  for (const s of scadenze) {
    s.origine = "file";
    s.chiaveFile = chiaveScadenzaFile(s);
    if (chiaviPortale.has(s.chiaveFile)) continue;
    if (tombstones.has(s.chiaveFile)) continue;
    s.id = idVecchi.get(s.chiaveFile);
    nuove.push(s);
  }
  db.scadenze = [...nuove, ...portale];
  db.daRateizzare = daRateizzare;
  db.fonteFile = fonte;
  return salvaFiscaleDb(db, utente);
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

export async function eliminaScadenzaFiscale(id: string, utente: string): Promise<FiscaleDb> {
  const db = await loadFiscaleDb();
  const riga = db.scadenze.find((s) => s.id === id);
  if (riga?.chiaveFile && !(db.tombstones ?? []).includes(riga.chiaveFile))
    db.tombstones = [...(db.tombstones ?? []), riga.chiaveFile];
  db.scadenze = db.scadenze.filter((s) => s.id !== id);
  return salvaFiscaleDb(db, utente);
}
