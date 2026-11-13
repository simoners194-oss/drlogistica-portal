// DR Portal — Finanze → "Costi appalti": storage.
// Un solo JSON nella libreria del sito (`FinanzaData/costi-appalti.json`),
// stesso pattern di cassa.json: le settimane caricate dal file dei costi.
// Il browser manda UNA settimana (o chiede di toglierne una); il server
// rilegge il file, la unisce e salva: così due persone che caricano settimane
// diverse non si pestano i piedi.
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
  emptyCostiAppaltiDb,
  fileCostiSettimana,
  leggiFoglioCosti,
  leggiPeriodo,
  righeDaCelle,
  unisciSettimana,
  validaSettimana,
  type AppaltoCosti,
  type CostiAppaltiDb,
  type LetturaCartellaCosti,
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

// ---------------------------------------------------------------------------
// LETTURA AUTOMATICA DALLA CARTELLA (1.117.0, Simone 13/11/2026: "non si può
// automatizzare? … sono tutti qui: Documenti Condivisi - Documenti\Operativo\
// Analisi"). I file settimanali ("Costi WK45.xlsx") stanno sul sito SharePoint
// "DocumentiCondivisi", lo stesso da cui si leggono già le paghe: il portale
// guarda la cartella, scarica SOLO i file nuovi o cambiati (dalla data di
// modifica) e li passa allo stesso lettore del caricamento a mano. Una
// settimana corretta nel portale non viene rimpiazzata dal file.
// ---------------------------------------------------------------------------
const SITO_DOCUMENTI = "DocumentiCondivisi";
export const CARTELLA_COSTI = "Operativo/Analisi";

interface FiglioCartella {
  id: string;
  name: string;
  file?: unknown;
  lastModifiedDateTime?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}
const encPath = (p: string) => p.split("/").map(encodeURIComponent).join("/");

export async function leggiCartellaCosti(utente: string): Promise<CostiAppaltiDb> {
  const ora = new Date().toISOString();
  const sito = await withDiscoveryRetry(() =>
    gatewayJson<{ id: string }>(`/sites/${TARGET_HOST}:/sites/${SITO_DOCUMENTI}`),
  );
  const figli = await withDiscoveryRetry(() =>
    gatewayJson<{ value: FiglioCartella[] }>(
      `/sites/${sito.id}/drive/root:/${encPath(CARTELLA_COSTI)}:/children?$select=id,name,file,lastModifiedDateTime&$top=200`,
    ),
  );
  const file = (figli.value ?? [])
    .filter((f) => f.file && fileCostiSettimana(f.name))
    .sort((a, b) => a.name.localeCompare(b.name, "it", { numeric: true }));
  let db = await loadCostiAppaltiDb();
  const esito: LetturaCartellaCosti = {
    il: ora,
    da: utente,
    trovati: file.length,
    letti: [],
    saltati: [],
    errori: [],
  };
  for (const f of file) {
    const modificato = f.lastModifiedDateTime ?? "";
    const gia = db.settimane.find((s) => s.file === f.name);
    if (gia?.origine === "portale") {
      esito.saltati.push(f.name);
      continue;
    }
    if (gia && gia.fileModificato === modificato) continue;
    try {
      const pieno = await withDiscoveryRetry(() =>
        gatewayJson<FiglioCartella>(`/sites/${sito.id}/drive/items/${f.id}`),
      );
      const url = pieno["@microsoft.graph.downloadUrl"];
      if (!url) throw new Error("downloadUrl non disponibile da Graph");
      const res = await fetch(url);
      if (!res.ok) throw new Error(`download fallito (${res.status})`);
      const wb = XLSX.read(await res.arrayBuffer());
      const appalti: AppaltoCosti[] = [];
      let etichetta = "";
      for (const nome of wb.SheetNames) {
        const letto = leggiFoglioCosti(
          nome,
          righeDaCelle(wb.Sheets[nome] as unknown as Record<string, unknown>),
        );
        if (!letto) continue;
        appalti.push(letto.appalto);
        etichetta ||= letto.etichetta;
      }
      if (!appalti.length) throw new Error("nessun foglio di appalto");
      const periodo = leggiPeriodo(etichetta, f.name, ora.slice(0, 10));
      if (!periodo) throw new Error("non capisco di quale settimana è");
      const s: SettimanaCosti = {
        ...validaSettimana({ ...periodo, etichetta, file: f.name, appalti }),
        caricataIl: ora,
        caricataDa: `cartella ${CARTELLA_COSTI}`,
        fileModificato: modificato,
      };
      if (db.settimane.find((x) => x.id === s.id)?.origine === "portale") {
        esito.saltati.push(f.name);
        continue;
      }
      // Un file rinominato: la sua settimana di prima non resta doppia.
      db = { ...db, settimane: db.settimane.filter((x) => x.file !== f.name) };
      db = unisciSettimana(db, s);
      esito.letti.push(f.name);
    } catch (e) {
      esito.errori.push(`${f.name}: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
  return scrivi(
    { ...db, ultimaLettura: esito },
    utente,
    `letta la cartella ${CARTELLA_COSTI} (${esito.trovati} file, ${esito.letti.length} nuovi o cambiati, ${esito.errori.length} errori)`,
  );
}
