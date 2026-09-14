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
  return {
    ...emptyFiscaleDb(),
    ...raw,
    scadenze: Array.isArray(raw.scadenze) ? raw.scadenze : [],
    daRateizzare: Array.isArray(raw.daRateizzare) ? raw.daRateizzare : [],
  };
}

/** Sostituisce l'intero scadenziario (il file di Sabrina è cumulativo). */
export async function replaceFiscale(
  scadenze: ScadenzaFiscale[],
  daRateizzare: DaRateizzareFiscale[],
  fonte: string,
  utente: string,
): Promise<FiscaleDb> {
  const cfg = await discoverSharePoint();
  const db = await loadFiscaleDb();
  db.scadenze = scadenze;
  db.daRateizzare = daRateizzare;
  db.fonteFile = fonte;
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
