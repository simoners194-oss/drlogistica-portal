// DR Portal — Finanza → Stipendi: storage lato server.
// -----------------------------------------------------------------------------
// Snapshot JSON unico nella libreria del sito (`FinanzaData/stipendi.json`),
// stesso pattern del modulo Mezzi: niente liste nuove da creare a mano,
// dataset piccolo (una manciata di mesi × ~100 dipendenti), lettura atomica.
// Concorrenza: read-modify-write per mutazione; ultimo-che-scrive-vince,
// accettato (scrive di fatto una persona al mese).

import * as XLSX from "xlsx";
import {
  cronToken,
  discoverSharePoint,
  gatewayJson,
  withDiscoveryRetry,
  logSp,
  SpHttpError,
  TARGET_HOST,
  tokenUguale,
  upsertFlussoCassa,
} from "./sharepoint.server";
import {
  applicaModifiche,
  chiaveNome,
  emptyStipendiDb,
  meseSuccessivo,
  parseCostiFile,
  parseStipendiDr,
  valoreDaFile,
  type AnagraficaDipendente,
  type CampoModificabile,
  type ModificaManuale,
  type MotivoNetto,
  type MotivoNettoTipo,
  type NettiMese,
  type StipendiDb,
  type StipendiMese,
} from "./stipendi-logic";

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

/** Sostituisce l'anagrafica contrattuale (dalla "Mappatura Dipendenti"),
 *  conservando gli override manuali delle mensilità già impostati. */
export async function replaceStipendiAnagrafica(
  anagrafica: AnagraficaDipendente[],
  fonte: string,
  utente: string,
): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const override = new Map<string, number>();
  for (const a of db.anagrafica ?? [])
    if (a.mensilita != null) override.set(chiaveNome(a.nome), a.mensilita);
  const chiaviNuove = new Set<string>();
  for (const a of anagrafica) {
    const k = chiaveNome(a.nome);
    chiaviNuove.add(k);
    const m = override.get(k);
    if (m != null) a.mensilita = m;
  }
  // Override di dipendenti FUORI dal nuovo file (presenti solo nei file
  // paghe): non vanno persi al re-import — restano come righe sintetiche.
  for (const a of db.anagrafica ?? []) {
    if (a.mensilita != null && !chiaviNuove.has(chiaveNome(a.nome)))
      anagrafica.push({ nome: a.nome, mensilita: a.mensilita });
  }
  db.anagrafica = anagrafica;
  db.anagraficaFonte = fonte;
  return saveStipendiDb(db, utente);
}

/** Imposta (o toglie, con null) le mensilità dichiarate di un dipendente —
 *  es. "io SONO 12" di Simone contro la stima 13 dai ratei paghe. */
export async function setStipendiMensilita(
  nome: string,
  mensilita: number | null,
  utente: string,
): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const k = chiaveNome(nome);
  const lista = db.anagrafica ?? [];
  const entry = lista.find((a) => chiaveNome(a.nome) === k);
  if (entry) {
    if (mensilita == null) delete entry.mensilita;
    else entry.mensilita = mensilita;
  } else if (mensilita != null) {
    lista.push({ nome, mensilita });
  }
  db.anagrafica = lista;
  return saveStipendiDb(db, utente);
}

/** Segna/toglie il PAGATO manuale per una lista di dipendenti (chiavi nome)
 *  su un mese di competenza — un solo giro per la selezione multipla. */
export async function setPagatiStipendi(
  mese: string,
  aggiungi: string[],
  togli: string[],
  utente: string,
): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const mappa = { ...(db.pagatiPerMese ?? {}) };
  const set = new Set(mappa[mese] ?? []);
  for (const k of aggiungi) set.add(k);
  for (const k of togli) set.delete(k);
  if (set.size > 0) mappa[mese] = [...set].sort();
  else delete mappa[mese];
  db.pagatiPerMese = mappa;
  return saveStipendiDb(db, utente);
}

/** Riga Stipendi dei Flussi (per MESE DI PAGAMENTO = competenza+1):
 *  totale = somma saldi del mese; pagati = somma saldi dei soli flaggati
 *  "pagato sì"; flaggati = quanti nomi sono spuntati (0 = nessuno usa
 *  ancora la spunta per quel mese e la tabella reale usa il totale). */
export async function flussiStipendi(): Promise<{
  stima: { media: number; mesi: string[] } | null;
  perMese: Record<string, { totale: number; pagati: number; flaggati: number }>;
}> {
  // Le correzioni a mano contano anche qui: i Flussi devono vedere gli
  // stessi numeri della tabella Stipendi.
  const db = applicaModifiche(await loadStipendiDb());
  const perMese: Record<string, { totale: number; pagati: number; flaggati: number }> = {};
  for (const n of db.netti ?? []) {
    const chiaviPagati = new Set(db.pagatiPerMese?.[n.mese] ?? []);
    let pagati = 0;
    // flaggati = SOLO le spunte che matchano una riga dei netti: una spunta
    // su un nome senza saldo (riga COSTI fuori da "Stipendi Dr") non deve
    // far scattare il regime "solo pagati" azzerando la riga dei Flussi.
    let flaggati = 0;
    for (const d of n.dipendenti)
      if (chiaviPagati.has(chiaveNome(d.nome))) {
        pagati += d.saldo;
        flaggati++;
      }
    const [anno, mm] = n.mese.split("-").map(Number);
    const mesePag = mm === 12 ? `${anno + 1}-01` : `${anno}-${String(mm + 1).padStart(2, "0")}`;
    perMese[mesePag] = {
      totale: Math.round(n.totaleSaldo * 100) / 100,
      pagati: Math.round(pagati * 100) / 100,
      flaggati,
    };
  }
  return { stima: await stimaStipendiMensileDa(db), perMese };
}

/** Stima dello stipendio mensile FUTURO per i Flussi di cassa (richiesta
 *  Simone 14/09): media del NETTO DOVUTO degli ultimi 2 mesi caricati da
 *  "Stipendi Dr.xlsx". Si usa il netto dovuto e non il saldo perché i saldi
 *  recenti sono abbattuti dagli anticipi già versati. */
export async function stimaStipendiMensile(): Promise<{ media: number; mesi: string[] } | null> {
  return stimaStipendiMensileDa(applicaModifiche(await loadStipendiDb()));
}

/** Correzione A MANO di un valore (costi o netti) per dipendente e mese:
 *  valore=null toglie la correzione e si torna al file. La riga conserva
 *  chi, quando e il valore che aveva il file ("prima"), che la tabella
 *  mostra sulla "M". Vive fuori dai mesi importati: il re-import non la tocca. */
export async function setModificaStipendio(
  input: {
    mese: string;
    chiave: string;
    nome: string;
    campo: CampoModificabile;
    valore: number | null;
  },
  utente: string,
): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const lista = [...(db.modifiche ?? [])];
  const i = lista.findIndex(
    (x) => x.mese === input.mese && x.chiave === input.chiave && x.campo === input.campo,
  );
  if (input.valore == null) {
    if (i < 0) return db; // niente da togliere, niente da salvare
    lista.splice(i, 1);
  } else {
    const prima = i >= 0 ? lista[i].prima : valoreDaFile(db, input.mese, input.chiave, input.campo);
    const rec: ModificaManuale = {
      mese: input.mese,
      chiave: input.chiave,
      nome: input.nome,
      campo: input.campo,
      valore: Math.round(input.valore * 100) / 100,
      prima,
      da: utente,
      il: new Date().toISOString(),
    };
    if (i >= 0) lista[i] = rec;
    else lista.push(rec);
  }
  db.modifiche = lista;
  logSp(
    "info",
    "stipendi.modifica",
    `${input.valore == null ? "Tolta" : "Salvata"} correzione ${input.campo} ${input.mese} ${input.nome} → ${input.valore ?? "file"} (${utente})`,
  );
  return saveStipendiDb(db, utente);
}

function stimaStipendiMensileDa(db: StipendiDb): { media: number; mesi: string[] } | null {
  const netti = (db.netti ?? [])
    .filter((n) => n.totaleStipendio > 0)
    .sort((a, b) => (a.mese < b.mese ? -1 : 1));
  const ultimi = netti.slice(-2);
  if (ultimi.length === 0) return null;
  const media =
    Math.round((ultimi.reduce((s, n) => s + n.totaleStipendio, 0) / ultimi.length) * 100) / 100;
  return { media, mesi: ultimi.map((n) => n.mese) };
}

// ---------------------------------------------------------------------------
// LETTURA AUTOMATICA DAI FILE (1.79.0, domanda Simone 21/09 "ma non li
// prende da solo da OneDrive?"). I file paghe stanno sul sito SharePoint
// "DocumentiCondivisi" (quello sincronizzato su OneDrive come "Documenti
// Condivisi - Documenti"), cartella Personale: Stipendi Dr.xlsx (netti, un
// foglio per mese) e MENSILITA'/<MESE ANNO>/COSTI <MESE> <ANNO>.xlsx. Il
// gateway Graph arriva a ogni sito del tenant: si scaricano e si passano
// agli STESSI parser dell'import manuale. Le correzioni a mano (M), le
// spunte Pagato e le mensilità dichiarate vivono a parte e restano.
// ---------------------------------------------------------------------------
const SITO_DOCUMENTI = "DocumentiCondivisi";
const CARTELLA_PERSONALE = "Personale";
const FILE_NETTI = "Stipendi Dr.xlsx";
const CARTELLA_MENSILITA = "MENSILITA'";
// Tracciato per-appalto di gennaio–maggio 2026 ("0X - Costi Personale
// <Mese> 2026.xlsx", un foglio per appalto, mese dal nome del file).
const CARTELLA_GEN_MAG = "COSTI PERSONALE GENNAIO - MAGGIO";

let sitoDocumentiCache: { id: string; at: number } | null = null;
async function sitoDocumentiId(): Promise<string> {
  if (sitoDocumentiCache && Date.now() - sitoDocumentiCache.at < 3600_000)
    return sitoDocumentiCache.id;
  const s = await withDiscoveryRetry(() =>
    gatewayJson<{ id: string }>(`/sites/${TARGET_HOST}:/sites/${SITO_DOCUMENTI}`),
  );
  sitoDocumentiCache = { id: s.id, at: Date.now() };
  return s.id;
}

interface DriveChild {
  id: string;
  name: string;
  folder?: unknown;
  file?: unknown;
  lastModifiedDateTime?: string;
  ["@microsoft.graph.downloadUrl"]?: string;
}
// Ogni segmento del percorso va codificato a parte (spazi, apostrofo di
// MENSILITA', parentesi): Graph vuole "root:/Personale/MENSILITA'/…:/children".
const encPath = (p: string) => p.split("/").map(encodeURIComponent).join("/");

async function figliDrive(siteId: string, path: string): Promise<DriveChild[]> {
  const res = await withDiscoveryRetry(() =>
    gatewayJson<{ value: DriveChild[] }>(
      `/sites/${siteId}/drive/root:/${encPath(path)}:/children?$select=id,name,folder,file,lastModifiedDateTime&$top=200`,
    ),
  );
  return res.value ?? [];
}

type Fogli = { nome: string; matrix: unknown[][] }[];
async function scaricaFogli(
  siteId: string,
  path: string,
): Promise<{ nome: string; fogli: Fogli; modificatoIl: string }> {
  const meta = await withDiscoveryRetry(() =>
    gatewayJson<DriveChild>(
      `/sites/${siteId}/drive/root:/${encPath(path)}?$select=id,name,lastModifiedDateTime`,
    ),
  );
  const full = await withDiscoveryRetry(() =>
    gatewayJson<DriveChild>(`/sites/${siteId}/drive/items/${meta.id}`),
  );
  const url = full["@microsoft.graph.downloadUrl"];
  if (!url) throw new Error(`${meta.name}: downloadUrl non disponibile da Graph.`);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${meta.name}: download fallito (${res.status}).`);
  const wb = XLSX.read(await res.arrayBuffer(), { cellDates: false });
  const fogli: Fogli = wb.SheetNames.map((nome) => ({
    nome,
    matrix: XLSX.utils.sheet_to_json(wb.Sheets[nome], {
      header: 1,
      raw: true,
      defval: null,
    }) as unknown[][],
  }));
  return { nome: meta.name, fogli, modificatoIl: meta.lastModifiedDateTime ?? "" };
}

export interface SyncStipendiEsito {
  nettiMesi: string[];
  costiMesi: string[];
  saltati: string[];
  errori: string[];
  flussiAggiornati: number;
  durataMs: number;
}

/** Legge Stipendi Dr.xlsx e i COSTI mensili da SharePoint e aggiorna lo
 *  snapshot. Idempotente: ogni mese trovato sostituisce l'omonimo (come
 *  l'import manuale); i mesi assenti dai file restano com'erano. La riga
 *  "Stipendi" dei Flussi si tocca SOLO se il saldo del mese è cambiato. */
export async function syncStipendiDaSharePoint(
  utente: string,
): Promise<{ db: StipendiDb; esito: SyncStipendiEsito }> {
  const started = Date.now();
  const esito: SyncStipendiEsito = {
    nettiMesi: [],
    costiMesi: [],
    saltati: [],
    errori: [],
    flussiAggiornati: 0,
    durataMs: 0,
  };
  const siteId = await sitoDocumentiId();
  const db = await loadStipendiDb();
  const ora = new Date().toISOString();
  const r2 = (n: number) => Math.round(n * 100) / 100;

  // --- NETTI (Stipendi Dr.xlsx) -------------------------------------------
  try {
    const f = await scaricaFogli(siteId, `${CARTELLA_PERSONALE}/${FILE_NETTI}`);
    const mesi = parseStipendiDr(f.fogli, f.nome);
    if (!mesi) esito.errori.push(`${FILE_NETTI}: nessun foglio riconosciuto`);
    else {
      const netti = [...(db.netti ?? [])];
      for (const m of mesi) {
        const i = netti.findIndex((x) => x.mese === m.mese);
        const saldoPrima = i >= 0 ? netti[i].totaleSaldo : null;
        m.caricatoIl = ora;
        m.caricatoDa = utente;
        if (i >= 0) netti[i] = m;
        else netti.push(m);
        esito.nettiMesi.push(m.mese);
        // Riga Stipendi dei Flussi (mese di PAGAMENTO = competenza+1), come
        // l'import manuale con la spunta accesa — ma solo se il saldo cambia.
        if (
          Math.abs(m.totaleSaldo) >= 0.005 &&
          (saldoPrima == null || Math.abs(r2(saldoPrima) - r2(m.totaleSaldo)) >= 0.005)
        ) {
          try {
            await upsertFlussoCassa({
              nome: "Stipendi",
              genere: "voce",
              mese: meseSuccessivo(m.mese),
              importo: -m.totaleSaldo,
              note: `Netto stipendi da versare, competenza ${m.mese} (${f.nome}, lettura automatica)`,
            });
            esito.flussiAggiornati++;
          } catch (err) {
            esito.errori.push(
              `Flussi ${m.mese}: ${err instanceof Error ? err.message : String(err)}`,
            );
          }
        }
      }
      netti.sort((a, b) => (a.mese < b.mese ? -1 : 1));
      db.netti = netti;
    }
  } catch (err) {
    esito.errori.push(`${FILE_NETTI}: ${err instanceof Error ? err.message : String(err)}`);
  }

  // --- COSTI: un file per cartella-mese (MENSILITA'/<MESE ANNO>/COSTI *.xlsx)
  //     più la cartella storica gennaio–maggio (tutti i file COSTI dentro).
  const leggiCosti = async (percorso: string, file: DriveChild) => {
    try {
      const f = await scaricaFogli(siteId, `${percorso}/${file.name}`);
      const res = parseCostiFile(f.fogli, file.name);
      if (!res || res.dipendenti.length === 0 || !res.mese) {
        esito.saltati.push(`${file.name}: tracciato non riconosciuto o mese assente`);
        return;
      }
      const mese: StipendiMese = {
        mese: res.mese,
        fonteFile: file.name,
        caricatoIl: ora,
        caricatoDa: utente,
        dipendenti: res.dipendenti,
      };
      const i = db.mesi.findIndex((m) => m.mese === res.mese);
      if (i >= 0) db.mesi[i] = mese;
      else db.mesi.push(mese);
      esito.costiMesi.push(res.mese);
    } catch (err) {
      esito.errori.push(`${file.name}: ${err instanceof Error ? err.message : String(err)}`);
    }
  };
  const fileCosti = (l: DriveChild[]) => l.filter((x) => x.file && /costi.*\.xlsx$/i.test(x.name));
  try {
    const cartelle = (await figliDrive(siteId, `${CARTELLA_PERSONALE}/${CARTELLA_MENSILITA}`))
      .filter((c) => c.folder)
      .sort((a, b) => a.name.localeCompare(b.name));
    for (const c of cartelle) {
      const percorso = `${CARTELLA_PERSONALE}/${CARTELLA_MENSILITA}/${c.name}`;
      let files: DriveChild[];
      try {
        files = fileCosti(await figliDrive(siteId, percorso));
      } catch (err) {
        esito.errori.push(`${c.name}: ${err instanceof Error ? err.message : String(err)}`);
        continue;
      }
      if (files.length === 0) {
        esito.saltati.push(`${c.name}: nessun file COSTI`);
        continue;
      }
      for (const file of files) await leggiCosti(percorso, file);
    }
  } catch (err) {
    esito.errori.push(`${CARTELLA_MENSILITA}: ${err instanceof Error ? err.message : String(err)}`);
  }
  try {
    const percorso = `${CARTELLA_PERSONALE}/${CARTELLA_GEN_MAG}`;
    const files = fileCosti(await figliDrive(siteId, percorso)).sort((a, b) =>
      a.name.localeCompare(b.name),
    );
    for (const file of files) await leggiCosti(percorso, file);
  } catch (err) {
    // Cartella storica: se un giorno sparisce non è un errore della lettura.
    esito.saltati.push(`${CARTELLA_GEN_MAG}: ${err instanceof Error ? err.message : String(err)}`);
  }
  db.mesi.sort((a, b) => (a.mese < b.mese ? -1 : 1));

  esito.durataMs = Date.now() - started;
  db.ultimaSync = {
    il: ora,
    da: utente,
    nettiMesi: esito.nettiMesi,
    costiMesi: esito.costiMesi,
    errori: esito.errori,
  };
  const salvato = await saveStipendiDb(db, utente);
  logSp(
    esito.errori.length ? "warn" : "info",
    "stipendi.sync",
    `Lettura file paghe: netti ${esito.nettiMesi.length} mesi, costi ${esito.costiMesi.length} mesi, flussi ${esito.flussiAggiornati}, saltati ${esito.saltati.length}, errori ${esito.errori.length}`,
    { durataMs: esito.durataMs },
  );
  return { db: salvato, esito };
}

/** Motivo per cui un dipendente non ha il netto del mese (o ha il netto senza
 *  costo): null lo toglie. Traccia chi/quando. */
export async function setMotivoNetto(
  input: {
    mese: string;
    chiave: string;
    nome: string;
    motivo: MotivoNettoTipo | null;
    nota?: string;
  },
  utente: string,
): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const lista = [...(db.motiviNetto ?? [])];
  const i = lista.findIndex((x) => x.mese === input.mese && x.chiave === input.chiave);
  if (input.motivo == null) {
    if (i < 0) return db;
    lista.splice(i, 1);
  } else {
    const rec: MotivoNetto = {
      mese: input.mese,
      chiave: input.chiave,
      nome: input.nome,
      motivo: input.motivo,
      nota: input.nota?.trim() || undefined,
      da: utente,
      il: new Date().toISOString(),
    };
    if (i >= 0) lista[i] = rec;
    else lista.push(rec);
  }
  db.motiviNetto = lista;
  logSp(
    "info",
    "stipendi.motivo",
    `${input.motivo == null ? "Tolto" : "Salvato"} motivo ${input.mese} ${input.nome} → ${input.motivo ?? "nessuno"} (${utente})`,
  );
  return saveStipendiDb(db, utente);
}

/** Innesco programmato (/cron-stipendi): vale il token "stipendi" oppure
 *  quello "fatture" già in mano al PC del giro giornaliero. */
export async function syncStipendiCron(token: string): Promise<SyncStipendiEsito> {
  const ok =
    tokenUguale(token, await cronToken("stipendi")) ||
    tokenUguale(token, await cronToken("fatture"));
  if (!ok) {
    logSp("warn", "stipendi.sync", "Chiamata programmata con token non valido");
    throw new Error("Token non valido.");
  }
  return (await syncStipendiDaSharePoint("cron")).esito;
}

/** Upsert dei NETTI da "Stipendi Dr.xlsx": ogni mese presente nel file
 *  sostituisce l'omonimo già caricato, gli altri restano. */
export async function upsertStipendiNetti(mesi: NettiMese[], utente: string): Promise<StipendiDb> {
  const db = await loadStipendiDb();
  const netti = [...(db.netti ?? [])];
  const ora = new Date().toISOString();
  for (const m of mesi) {
    m.caricatoIl = ora;
    m.caricatoDa = utente;
    const i = netti.findIndex((x) => x.mese === m.mese);
    if (i >= 0) netti[i] = m;
    else netti.push(m);
  }
  netti.sort((a, b) => (a.mese < b.mese ? -1 : 1));
  db.netti = netti;
  return saveStipendiDb(db, utente);
}
