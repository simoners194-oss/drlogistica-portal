// DR Portal — Finanza → Flussi → "Cassa proprietà": storage.
// Un solo JSON nella libreria del sito (`FinanzaData/cassa.json`), stesso
// pattern di fiscale.json: stati per riga, righe manuali, fotografie
// salvate, impostazioni. Il client manda l'intero db: la versione serve a
// non sovrascrivere un salvataggio fatto da un'altra scheda nel frattempo.
import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
} from "./sharepoint.server";
import { emptyCassaDb, type CassaDb } from "./cassa-logic";

const DB_PATH = "FinanzaData/cassa.json";

interface DriveItemMeta {
  id?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function loadCassaDb(): Promise<CassaDb> {
  const cfg = await discoverSharePoint();
  let meta: DriveItemMeta;
  try {
    meta = await withDiscoveryRetry(() =>
      gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}?$select=id`),
    );
  } catch (err) {
    if (err instanceof SpHttpError && err.status === 404) return emptyCassaDb();
    throw err;
  }
  const metaFull = await withDiscoveryRetry(() =>
    gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/items/${meta.id}`),
  );
  const url = metaFull["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("cassa.json: downloadUrl non disponibile da Graph.");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`cassa.json: download fallito (${res.status}).`);
  const raw = (await res.json()) as Partial<CassaDb>;
  return {
    ...emptyCassaDb(),
    ...raw,
    giornoLimite: Number(raw.giornoLimite) > 0 ? Number(raw.giornoLimite) : 5,
    stati: raw.stati && typeof raw.stati === "object" ? raw.stati : {},
    manuali: Array.isArray(raw.manuali) ? raw.manuali : [],
    fotografie: Array.isArray(raw.fotografie) ? raw.fotografie : [],
  };
}

/** Salva l'intero db. Se sul server c'è una versione più nuova di quella da
 *  cui il client è partito, rifiuta: chi salva deve prima ricaricare. */
export async function salvaCassaDb(db: CassaDb, utente: string): Promise<CassaDb> {
  const cfg = await discoverSharePoint();
  const attuale = await loadCassaDb();
  if ((attuale.versione ?? 0) > (db.versione ?? 0))
    throw new Error(
      "La cassa è stata salvata da un'altra scheda nel frattempo: ricarica la pagina e rifai la modifica.",
    );
  const nuovo: CassaDb = {
    ...db,
    versione: (attuale.versione ?? 0) + 1,
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
    "cassa.save",
    `Cassa proprietà salvata (v${nuovo.versione}, ${Object.keys(nuovo.stati).length} stati, ${nuovo.manuali.length} manuali, ${nuovo.fotografie.length} fotografie)`,
  );
  return nuovo;
}
