// DR Portal — Finanza → Stipendi: storage lato server.
// -----------------------------------------------------------------------------
// Snapshot JSON unico nella libreria del sito (`FinanzaData/stipendi.json`),
// stesso pattern del modulo Mezzi: niente liste nuove da creare a mano,
// dataset piccolo (una manciata di mesi × ~100 dipendenti), lettura atomica.
// Concorrenza: read-modify-write per mutazione; ultimo-che-scrive-vince,
// accettato (scrive di fatto una persona al mese).

import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
} from "./sharepoint.server";
import { emptyStipendiDb, type StipendiDb, type StipendiMese } from "./stipendi-logic";

const DB_PATH = "FinanzaData/stipendi.json";

interface DriveItemMeta {
  id?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function loadStipendiDb(): Promise<StipendiDb> {
  const cfg = await discoverSharePoint();
  let meta: DriveItemMeta;
  try {
    meta = await withDiscoveryRetry(() =>
      gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}?$select=id`),
    );
  } catch (err) {
    if (err instanceof SpHttpError && err.status === 404) return emptyStipendiDb();
    throw err;
  }
  const metaFull = await withDiscoveryRetry(() =>
    gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/items/${meta.id}`),
  );
  const url = metaFull["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("stipendi.json: downloadUrl non disponibile da Graph.");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`stipendi.json: download fallito (${res.status}).`);
  const raw = (await res.json()) as Partial<StipendiDb>;
  return { ...emptyStipendiDb(), ...raw, mesi: Array.isArray(raw.mesi) ? raw.mesi : [] };
}

async function saveStipendiDb(db: StipendiDb, utente: string): Promise<StipendiDb> {
  const cfg = await discoverSharePoint();
  db.versione = (db.versione ?? 0) + 1;
  db.aggiornatoIl = new Date().toISOString();
  db.aggiornatoDa = utente;
  const body = JSON.stringify(db);
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body,
    }),
  );
  logSp(
    "info",
    "stipendi.save",
    `Snapshot stipendi salvato (v${db.versione}, ${db.mesi.length} mesi)`,
  );
  return db;
}

/** Upsert di un mese (sostituisce il mese omonimo se già caricato). */
export async function upsertStipendiMese(mese: StipendiMese, utente: string): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  mese.caricatoIl = new Date().toISOString();
  mese.caricatoDa = utente;
  const i = db.mesi.findIndex((m) => m.mese === mese.mese);
  if (i >= 0) db.mesi[i] = mese;
  else db.mesi.push(mese);
  db.mesi.sort((a, b) => (a.mese < b.mese ? -1 : 1));
  return saveStipendiDb(db, utente);
}

export async function deleteStipendiMese(mese: string, utente: string): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  db.mesi = db.mesi.filter((m) => m.mese !== mese);
  return saveStipendiDb(db, utente);
}
