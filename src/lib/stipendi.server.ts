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
import {
  chiaveNome,
  emptyStipendiDb,
  type AnagraficaDipendente,
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
  const db = await loadStipendiDb();
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
  return stimaStipendiMensileDa(await loadStipendiDb());
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
