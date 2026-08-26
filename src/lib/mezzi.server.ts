// DR Portal — modulo Mezzi: storage e operazioni lato server.
// -----------------------------------------------------------------------------
// Lo stato del modulo vive in UN file JSON nella libreria del sito SharePoint
// (`MezziData/mezzi-db.json`), letto/scritto via Connector Gateway. Scelta
// deliberata: niente nuove liste da creare a mano sul sito, snapshot atomico
// per la UI, dataset piccolo (decine di mezzi, centinaia di righe).
// Concorrenza: mutazioni read-modify-write lato server (finestra di conflitto
// di ~1-2s con 3 utenti totali); il campo `versione` consente in futuro un
// controllo ottimistico vero. Ultimo-che-scrive-vince, accettato per la v1.
// -----------------------------------------------------------------------------

import {
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  enqueueEmail,
  logSp,
  SpHttpError,
} from "./sharepoint.server";
import {
  emptyMezziDb,
  giorniA,
  PARAMETRI_DEFAULT,
  type MezziDb,
  type Scadenza,
} from "./mezzi-types";

const DB_PATH = "MezziData/mezzi-db.json";

// ---------------------------------------------------------------------------
// Lettura / scrittura snapshot
// ---------------------------------------------------------------------------

interface DriveItemMeta {
  id?: string;
  size?: number;
  ["@microsoft.graph.downloadUrl"]?: string;
}

export async function loadMezziDb(): Promise<MezziDb> {
  const cfg = await discoverSharePoint();
  let meta: DriveItemMeta;
  try {
    meta = await withDiscoveryRetry(() =>
      gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}?$select=id,size`),
    );
  } catch (err) {
    if (err instanceof SpHttpError && err.status === 404) {
      logSp("info", "mezzi.load", "mezzi-db.json assente: parto da database vuoto");
      return emptyMezziDb();
    }
    throw err;
  }
  // Il contenuto si scarica dall'URL pre-firmato (evita il 302 del gateway
  // su GET :/content). L'URL non richiede autenticazione ed è a scadenza breve.
  const metaFull = await withDiscoveryRetry(() =>
    gatewayJson<DriveItemMeta>(`/sites/${cfg.siteId}/drive/items/${meta.id}`),
  );
  const url = metaFull["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error("mezzi-db.json: downloadUrl non disponibile da Graph.");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`mezzi-db.json: download fallito (${res.status}).`);
  const raw = (await res.json()) as Partial<MezziDb>;
  // Merge difensivo con lo scheletro: campi mancanti → default.
  const base = emptyMezziDb();
  return {
    ...base,
    ...raw,
    parametri: { ...PARAMETRI_DEFAULT, ...(raw.parametri ?? {}) },
  } as MezziDb;
}

export async function saveMezziDb(db: MezziDb, utente?: string): Promise<MezziDb> {
  const cfg = await discoverSharePoint();
  db.versione = (db.versione ?? 0) + 1;
  db.aggiornatoIl = new Date().toISOString();
  if (utente) db.aggiornatoDa = utente;
  const body = JSON.stringify(db);
  await withDiscoveryRetry(() =>
    gatewayJson(`/sites/${cfg.siteId}/drive/root:/${DB_PATH}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body,
    }),
  );
  logSp("info", "mezzi.save", `Snapshot mezzi salvato (v${db.versione}, ${body.length} byte)`);
  return db;
}

/** Mutazione read-modify-write: carica lo snapshot, applica `fn`, salva. */
export async function mutateMezziDb(
  utente: string,
  fn: (db: MezziDb) => void | Promise<void>,
): Promise<MezziDb> {
  const db = await loadMezziDb();
  await fn(db);
  return saveMezziDb(db, utente);
}

/** Import completo (seed iniziale o ripristino): sostituisce lo snapshot. */
export async function importMezziDb(json: string, utente: string): Promise<MezziDb> {
  let parsed: Partial<MezziDb>;
  try {
    parsed = JSON.parse(json) as Partial<MezziDb>;
  } catch {
    throw new Error("JSON non valido.");
  }
  if (!Array.isArray(parsed.mezzi)) {
    throw new Error('Il JSON non sembra un database Mezzi (manca l\'array "mezzi").');
  }
  const attuale = await loadMezziDb().catch(() => emptyMezziDb());
  const base = emptyMezziDb();
  const db: MezziDb = {
    ...base,
    ...parsed,
    versione: attuale.versione, // saveMezziDb incrementa
    parametri: { ...PARAMETRI_DEFAULT, ...(parsed.parametri ?? {}) },
  } as MezziDb;
  logSp(
    "info",
    "mezzi.import",
    `Import snapshot: ${db.mezzi.length} mezzi, ${db.scadenze.length} scadenze`,
  );
  return saveMezziDb(db, utente);
}

// ---------------------------------------------------------------------------
// Cron scadenze: email 90/60/30 (+ scadute) via Coda Email
// ---------------------------------------------------------------------------

export interface CronMezziResult {
  controllate: number;
  notifiche: number;
  emailInviata: boolean;
  dettagli: string[];
}

/** Etichetta leggibile del tipo scadenza (per le email). */
const TIPO_LABEL: Record<string, string> = {
  revisione: "Revisione",
  assicurazione: "Assicurazione",
  bollo: "Bollo",
  atp: "Attestato ATP",
  tagliando: "Tagliando",
  fine_contratto: "Fine contratto noleggio",
  disdetta: "Finestra disdetta contratto",
  ztl: "Permesso ZTL",
  patente: "Patente/CQC",
  altro: "Scadenza",
};

function fmtIt(iso: string): string {
  const [y, m, g] = iso.slice(0, 10).split("-");
  return `${g}/${m}/${y}`;
}

/**
 * Scansione scadenze: per ogni soglia configurata (default 90/60/30 giorni)
 * invia UNA email cumulativa con le scadenze appena entrate nella soglia,
 * e marca l'invio sulla scadenza (niente doppioni ai run successivi).
 * Le scadenze già superate e mai segnalate finiscono nel blocco "SCADUTE".
 */
export async function cronScadenzeMezzi(): Promise<CronMezziResult> {
  const db = await loadMezziDb();
  const soglie = [...(db.parametri.soglieAlertGiorni ?? [90, 60, 30])].sort((a, b) => b - a);
  const destinatari = db.parametri.emailAlert ?? [];
  const targaDi = new Map(db.mezzi.map((m) => [m.id, m.targa]));

  const daNotificare: { s: Scadenza; soglia: number; giorni: number }[] = [];
  let controllate = 0;
  for (const s of db.scadenze) {
    if (s.chiusa) continue;
    controllate++;
    const giorni = giorniA(s.scadenza);
    const inviati = s.alertInviati ?? {};
    if (giorni < 0) {
      if (!inviati["0"]) daNotificare.push({ s, soglia: 0, giorni });
      continue;
    }
    // La soglia più stretta già raggiunta e non ancora notificata.
    const soglia = soglie.filter((g) => giorni <= g).pop();
    if (soglia !== undefined && !inviati[String(soglia)]) {
      daNotificare.push({ s, soglia, giorni });
    }
  }

  const dettagli: string[] = [];
  let emailInviata = false;
  if (daNotificare.length > 0 && destinatari.length > 0) {
    const oggi = new Date().toISOString();
    const blocchi = new Map<number, string[]>();
    for (const { s, soglia, giorni } of daNotificare) {
      const targa = s.mezzoId ? (targaDi.get(s.mezzoId) ?? s.mezzoId) : "—";
      const label = TIPO_LABEL[s.tipo] ?? s.tipo;
      const riga = `- ${targa} · ${label}${s.descrizione ? ` (${s.descrizione})` : ""} → scadenza ${fmtIt(s.scadenza)} (${giorni < 0 ? `SCADUTA da ${-giorni} giorni` : `tra ${giorni} giorni`})`;
      if (!blocchi.has(soglia)) blocchi.set(soglia, []);
      blocchi.get(soglia)!.push(riga);
      s.alertInviati = { ...(s.alertInviati ?? {}), [String(soglia)]: oggi };
      dettagli.push(riga);
    }
    const parti: string[] = [
      "Promemoria automatico scadenze parco mezzi (DR Portal — modulo Mezzi).",
      "",
    ];
    for (const soglia of [...blocchi.keys()].sort((a, b) => a - b)) {
      parti.push(soglia === 0 ? "SCADUTE — intervenire subito:" : `Entro ${soglia} giorni:`);
      parti.push(...blocchi.get(soglia)!, "");
    }
    parti.push("Dettaglio e gestione: https://portal.drlogistica.it/mezzi (scheda Scadenze).");
    emailInviata = await enqueueEmail({
      destinatari,
      oggetto: `Mezzi: ${daNotificare.length} scadenz${daNotificare.length === 1 ? "a" : "e"} da gestire`,
      corpo: parti.join("\n"),
    });
    await saveMezziDb(db, "cron");
  }

  return { controllate, notifiche: daNotificare.length, emailInviata, dettagli };
}
