// DR Portal — Finanze → "Costi appalti": storage.
// Un solo JSON nella libreria del sito (`FinanzaData/costi-appalti.json`),
// stesso pattern di cassa.json: le settimane caricate dal file dei costi.
// Il browser manda UNA settimana (o chiede di toglierne una); il server
// rilegge il file, la unisce e salva: così due persone che caricano settimane
// diverse non si pestano i piedi.
import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
} from "./sharepoint.server";
import {
  emptyCostiAppaltiDb,
  unisciSettimana,
  type CostiAppaltiDb,
  type SettimanaCosti,
} from "./costi-appalti-logic";

const DB_PATH = "FinanzaData/costi-appalti.json";

interface DriveItemMeta {
  id?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function loadCostiAppaltiDb(): Promise<CostiAppaltiDb> {
  const cfg = await discoverSharePoint();
  let meta: DriveItemMeta;
  try {
    meta = await withDiscoveryRetry(() =>
      gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}?$select=id`),
    );
  } catch (err) {
    // Il file nasce al primo caricamento.
    if (err instanceof SpHttpError && err.status === 404) return emptyCostiAppaltiDb();
    throw err;
  }
  const metaFull = await withDiscoveryRetry(() =>
    gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/items/${meta.id}`),
  );
  const url = metaFull["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("costi-appalti.json: downloadUrl non disponibile da Graph.");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`costi-appalti.json: download fallito (${res.status}).`);
  const raw = (await res.json()) as Partial<CostiAppaltiDb>;
  return {
    ...emptyCostiAppaltiDb(),
    ...raw,
    versione: Number(raw.versione) || 0,
    settimane: Array.isArray(raw.settimane) ? raw.settimane : [],
  };
}

async function scrivi(db: CostiAppaltiDb, utente: string, cosa: string): Promise<CostiAppaltiDb> {
  const cfg = await discoverSharePoint();
  const nuovo: CostiAppaltiDb = {
    ...db,
    versione: (db.versione ?? 0) + 1,
    aggiornatoIl: new Date().toISOString(),
    aggiornatoDa: utente,
  };
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(nuovo),
    }),
  );
  logSp(
    "info",
    "costi-appalti.save",
    `Costi appalti: ${cosa} (v${nuovo.versione}, ${nuovo.settimane.length} settimane)`,
  );
  return nuovo;
}

/** Aggiunge una settimana; se c'è già (stessa data d'inizio) la sostituisce. */
export async function salvaSettimanaCosti(
  settimana: SettimanaCosti,
  utente: string,
): Promise<CostiAppaltiDb> {
  const attuale = await loadCostiAppaltiDb();
  const s: SettimanaCosti = {
    ...settimana,
    caricataIl: new Date().toISOString(),
    caricataDa: utente,
  };
  return scrivi(unisciSettimana(attuale, s), utente, `caricata la settimana ${s.dal}`);
}

export async function eliminaSettimanaCosti(id: string, utente: string): Promise<CostiAppaltiDb> {
  const attuale = await loadCostiAppaltiDb();
  if (!attuale.settimane.some((s) => s.id === id)) return attuale;
  return scrivi(
    { ...attuale, settimane: attuale.settimane.filter((s) => s.id !== id) },
    utente,
    `tolta la settimana ${id}`,
  );
}
