// DR Portal — Finanze → "Costi appalti" (Simone 13/11/2026: "dobbiamo
// integrare un nuovo tab con il foglio in allegato, che posso filtrare i vari
// clienti per vedere i costi").
//
// Il foglio è il file settimanale dei costi ("Costi WK45.xlsx"): un foglio di
// Excel per appalto (Postadoc, iMile, Zingali…) più il foglio "Totale". Ogni
// foglio ha tre colonne (VOCE | SETTIMANA (02/11 - 08/11) | Descrizione) e le
// righe: i ricavi, la riga RICAVI, i costi voce per voce (fornitore, costi
// accessori, addebiti, fuel, stipendi, noleggi, pedaggi), la riga COSTI, il
// MOL 1 e la sua percentuale.
//
// Qui c'è solo logica pura (niente React, niente Excel): la lettura di un
// foglio già ridotto a righe [voce, importo, nota], i totali RICALCOLATI dalle
// voci (quelli scritti nel file servono solo da controllo), la settimana e i
// conti per la vista. Il file si legge nel browser, si salva il risultato.

export interface VoceCosti {
  /** Come è scritta nel file ("FORNITORE", "COSTI ACCESSORI"…). */
  voce: string;
  importo: number;
  /** La colonna Descrizione ("dato di settembre", "10%", "0,85 € per km"). */
  nota?: string;
  /** Da dove viene l'importo, nelle settimane preparate dal portale (assente
   *  = letto dal file): "paghe" = dal costo delle paghe del portale;
   *  "mensile" = un importo mensile riportato ai giorni della settimana;
   *  "regola" = calcolato (margine fisso, tariffa al km); "copiata" = uguale
   *  alla settimana prima, DA CONFERMARE; "mano" = scritto a mano. */
  fonte?: FonteVoce;
  /** L'importo mensile dietro la cifra della settimana (paghe e mensili). */
  mensile?: number;
}

export type FonteVoce = "paghe" | "mensile" | "regola" | "copiata" | "mano";
export const FONTI_VOCE: readonly FonteVoce[] = ["paghe", "mensile", "regola", "copiata", "mano"];

export interface AppaltoCosti {
  /** Il nome del foglio: "Postadoc", "Univex Milano"… */
  nome: string;
  /** Riga GG: giorni della settimana lavorati. */
  giorni?: number;
  ricavi: VoceCosti[];
  costi: VoceCosti[];
  /** Chilometri della settimana, per gli appalti a tariffa al km (Widem). */
  km?: number;
  /** I km sono quelli della settimana prima: da confermare. */
  kmCopiati?: boolean;
  /** L'etichetta delle paghe del portale da cui vengono gli stipendi di
   *  questo appalto ("ZINGALI"); "" = nessuna, scelta a mano. */
  paghe?: string;
  /** I totali scritti nel file: solo per il controllo al caricamento. */
  fileRicavi?: number;
  fileCosti?: number;
  fileMol?: number;
}

export interface SettimanaCosti {
  /** La data d'inizio (YYYY-MM-DD): una settimana, un file. */
  id: string;
  /** Numero della settimana nell'anno (ISO): 45. */
  wk: number;
  dal: string;
  al: string;
  /** L'intestazione del file, tale e quale: "SETTIMANA (02/11 - 08/11)". */
  etichetta: string;
  file: string;
  caricataIl: string;
  caricataDa: string;
  /** "portale" = preparata o corretta nel portale; assente = letta dal file. */
  origine?: "portale";
  /** Data di modifica del file letto dalla cartella condivisa: finché non
   *  cambia, il file non si rilegge. */
  fileModificato?: string;
  appalti: AppaltoCosti[];
}

export interface CostiAppaltiDb {
  versione: number;
  /** Dalla più recente. */
  settimane: SettimanaCosti[];
  /** L'ultima lettura della cartella condivisa (Operativo/Analisi). */
  ultimaLettura?: LetturaCartellaCosti;
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export const emptyCostiAppaltiDb = (): CostiAppaltiDb => ({ versione: 0, settimane: [] });

/** Com'è andata una lettura della cartella dei file settimanali. */
export interface LetturaCartellaCosti {
  il: string;
  da: string;
  /** File trovati nella cartella. */
  trovati: number;
  /** File letti perché nuovi o cambiati. */
  letti: string[];
  /** File non riletti perché corretti nel portale. */
  saltati: string[];
  errori: string[];
}

/** I file settimanali dei costi si riconoscono dal nome: "Costi WK45.xlsx". */
export const fileCostiSettimana = (nome: string) =>
  /^costi\s*wk\s*\d{1,2}\.xls[xm]?$/i.test(nome.trim());

const r2 = (n: number) => Math.round(n * 100) / 100;
// Le voci tengono quattro decimali: nel file sono frazioni (un mese diviso
// per le settimane) e arrotondarle al centesimo una per una sposta i totali
// di un centesimo rispetto a Excel. Al centesimo si arrotondano i totali.
const r4 = (n: number) => Math.round(n * 10000) / 10000;

/** Etichetta di una riga: maiuscola, senza spazi doppi né in coda. */
export const normVoce = (s: unknown) =>
  String(s ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

/** Nome di un appalto per i confronti ("Number1" = "NUMBER 1"). */
export const chiaveAppalto = (s: string) =>
  s
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, "");

/** Una riga del foglio già letta: etichetta, importo (null se vuoto), nota. */
export type RigaFoglioCosti = readonly [voce: unknown, importo: unknown, nota?: unknown];

const numero = (v: unknown): number | null => {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  const s = v.replace(/[€\s]/g, "");
  if (!s || s === "-") return null;
  // "1.234,56" all'italiana oppure "1,234.56" all'inglese: l'ultimo
  // separatore è quello dei decimali.
  const virgola = s.lastIndexOf(",");
  const punto = s.lastIndexOf(".");
  const n = Number(virgola > punto ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
};

/** Dalle celle di un foglio letto con SheetJS (oggetto "A1" → cella) alle
 *  righe [voce, importo, nota] delle prime tre colonne. La nota numerica (il
 *  "10%" accanto al fornitore) si prende com'è scritta nel foglio. */
export function righeDaCelle(ws: Record<string, unknown>, maxRighe = 120): RigaFoglioCosti[] {
  const cella = (col: string, r: number) =>
    ws[`${col}${r}`] as { v?: unknown; w?: string; t?: string } | undefined;
  // La riga d'intestazione: la colonna B dice "SETTIMANA…" (nel foglio
  // Postadoc del tracciato vecchio la prima cella è "POSTADOC TOTALE", non
  // "VOCE"). La descrizione sta nella colonna intitolata "Descrizione": la
  // terza nel tracciato di oggi, la quinta in quello vecchio (WK40), dove la
  // terza è l'incidenza.
  let testata = 0;
  for (let r = 1; r <= 6 && !testata; r++)
    if (/settimana/i.test(String(cella("B", r)?.v ?? "")) || normVoce(cella("A", r)?.v) === "VOCE")
      testata = r;
  let colNota = "C";
  if (testata) {
    const trovata = ["C", "D", "E", "F", "G"].find((col) =>
      /^descrizione$/i.test(String(cella(col, testata)?.v ?? "").trim()),
    );
    if (trovata) colNota = trovata;
  }
  const out: RigaFoglioCosti[] = [];
  for (let r = 1; r <= maxRighe; r++) {
    const a = cella("A", r);
    const b = cella("B", r);
    const c = cella(colNota, r);
    const nota = c ? (c.t === "n" ? String(c.w ?? c.v ?? "") : String(c.v ?? "")) : "";
    out.push([r === testata ? "VOCE" : (a?.v ?? ""), b?.v ?? null, r === testata ? "" : nota]);
  }
  return out;
}

/** Legge un foglio di appalto. `null` se non è un foglio di appalto (il
 *  foglio "Totale", o un foglio senza le righe RICAVI e COSTI). */
export function leggiFoglioCosti(
  nome: string,
  righe: readonly RigaFoglioCosti[],
): { appalto: AppaltoCosti; etichetta: string } | null {
  const testata = righe.findIndex((r) => normVoce(r?.[0]) === "VOCE");
  if (testata < 0) return null;
  const etichetta = String(righe[testata]?.[1] ?? "")
    .replace(/\s+/g, " ")
    .trim();
  const appalto: AppaltoCosti = { nome: nome.trim(), ricavi: [], costi: [] };
  // "ricavi" finché non arriva la riga RICAVI, poi "costi" fino a COSTI.
  let fase: "ricavi" | "costi" | "fine" = "ricavi";
  let visteRicavi = false;
  let visteCosti = false;
  for (const r of righe.slice(testata + 1)) {
    const voce = normVoce(r?.[0]);
    const importo = numero(r?.[1]);
    const nota = String(r?.[2] ?? "")
      .replace(/\s+/g, " ")
      .trim();
    if (!voce) continue;
    if (voce === "WK") continue;
    if (voce === "GG") {
      if (importo != null && importo > 0 && importo <= 31) appalto.giorni = importo;
      continue;
    }
    if (voce === "RICAVI") {
      appalto.fileRicavi = importo == null ? undefined : r2(importo);
      fase = "costi";
      visteRicavi = true;
      continue;
    }
    if (voce === "COSTI") {
      appalto.fileCosti = importo == null ? undefined : r2(importo);
      fase = "fine";
      visteCosti = true;
      continue;
    }
    if (/^MOL\b/.test(voce)) {
      if (appalto.fileMol === undefined && importo != null) appalto.fileMol = r2(importo);
      continue;
    }
    if (fase === "fine") continue;
    // Le voci vuote restano (a zero): il foglio ha sempre le stesse righe.
    const riga: VoceCosti = { voce, importo: r4(importo ?? 0) };
    if (nota) riga.nota = nota.slice(0, 200);
    (fase === "ricavi" ? appalto.ricavi : appalto.costi).push(riga);
  }
  if (!visteRicavi || !visteCosti) return null;
  // Il foglio "Totale" ha le righe RICAVI e COSTI ma nessuna voce sotto: non
  // è un appalto (i totali si ricalcolano dai fogli degli appalti).
  if (appalto.ricavi.length + appalto.costi.length === 0) return null;
  if (/^(totale|total|riepilogo)$/i.test(appalto.nome)) return null;
  return { appalto, etichetta };
}

export interface TotaliAppalto {
  ricavi: number;
  costi: number;
  mol: number;
  /** MOL / ricavi (0,10 = 10%); null se non ci sono ricavi. */
  margine: number | null;
}

export function totaliVoci(ricavi: number, costi: number): TotaliAppalto {
  const mol = r2(ricavi - costi);
  return {
    ricavi: r2(ricavi),
    costi: r2(costi),
    mol,
    margine: ricavi > 0.005 ? mol / ricavi : null,
  };
}

/** I totali di un appalto, RICALCOLATI dalle sue voci. */
export const totaliAppalto = (a: AppaltoCosti): TotaliAppalto =>
  totaliVoci(
    a.ricavi.reduce((s, v) => s + v.importo, 0),
    a.costi.reduce((s, v) => s + v.importo, 0),
  );

/** I totali di più appalti insieme (il "Totale DR" del file). */
export function totaliInsieme(appalti: readonly AppaltoCosti[]): TotaliAppalto {
  let ricavi = 0;
  let costi = 0;
  for (const a of appalti) {
    const t = totaliAppalto(a);
    ricavi += t.ricavi;
    costi += t.costi;
  }
  return totaliVoci(ricavi, costi);
}

/** Dove i totali scritti nel file non tornano con la somma delle voci (più
 *  di cinque centesimi): si mostra al caricamento, prima di salvare. */
export function controlliAppalto(a: AppaltoCosti): string[] {
  const t = totaliAppalto(a);
  const out: string[] = [];
  const diverso = (x: number | undefined, y: number) => x !== undefined && Math.abs(x - y) > 0.05;
  if (diverso(a.fileRicavi, t.ricavi)) out.push("ricavi");
  if (diverso(a.fileCosti, t.costi)) out.push("costi");
  if (diverso(a.fileMol, t.mol)) out.push("mol");
  return out;
}

/** Le voci di un lato (ricavi o costi) di più appalti, nell'ordine in cui
 *  compaiono nei fogli. */
export function vociDi(appalti: readonly AppaltoCosti[], lato: "ricavi" | "costi"): string[] {
  const out: string[] = [];
  for (const a of appalti) for (const v of a[lato]) if (!out.includes(v.voce)) out.push(v.voce);
  return out;
}

export const voceDi = (a: AppaltoCosti, lato: "ricavi" | "costi", voce: string) =>
  a[lato].find((v) => v.voce === voce);

// --- la settimana -------------------------------------------------------------

const iso = (d: Date) => d.toISOString().slice(0, 10);
const giorno = (anno: number, mese: number, g: number) => new Date(Date.UTC(anno, mese - 1, g));

/** Settimana ISO (lunedì-domenica) di una data YYYY-MM-DD. */
export function settimanaIso(data: string): { anno: number; wk: number } {
  const d = new Date(`${data}T00:00:00Z`);
  const g = (d.getUTCDay() + 6) % 7; // lunedì = 0
  d.setUTCDate(d.getUTCDate() - g + 3); // il giovedì della settimana
  const primoGiovedi = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const wk =
    1 +
    Math.round(
      ((d.getTime() - primoGiovedi.getTime()) / 86400000 -
        3 +
        ((primoGiovedi.getUTCDay() + 6) % 7)) /
        7,
    );
  return { anno: d.getUTCFullYear(), wk };
}

/** Dall'intestazione "SETTIMANA (02/11 - 08/11)" e dal nome del file
 *  ("Costi WK45.xlsx") alle date. L'anno non è scritto: si prende quello in
 *  cui la settimana ha il numero del nome del file; senza numero, la data più
 *  vicina a oggi che non sia nel futuro lontano. */
export function leggiPeriodo(
  etichetta: string,
  nomeFile: string,
  oggi: string,
): { dal: string; al: string; wk: number } | null {
  const m = etichetta.match(
    /(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[-–—a]+\s*(\d{1,2})\s*[/.-]\s*(\d{1,2})/i,
  );
  if (!m) {
    // Intestazione senza date (il tracciato vecchio dice solo "Settimana"):
    // resta il numero nel nome del file, lunedì-domenica di quella settimana.
    const wk = Number(nomeFile.match(/w(?:ee)?k\s*[-_]?\s*(\d{1,2})/i)?.[1] ?? NaN);
    if (!(wk >= 1 && wk <= 53)) return null;
    const lunedi = (anno: number) => {
      // Il 4 gennaio sta sempre nella settimana 1.
      const g4 = giorno(anno, 1, 4);
      const primoLunedi = g4.getTime() - ((g4.getUTCDay() + 6) % 7) * 86400000;
      return iso(new Date(primoLunedi + (wk - 1) * 7 * 86400000));
    };
    const annoQui = Number(oggi.slice(0, 4));
    const entro = iso(new Date(Date.parse(`${oggi}T00:00:00Z`) + 45 * 86400000));
    const dalWk = lunedi(annoQui) <= entro ? lunedi(annoQui) : lunedi(annoQui - 1);
    return {
      dal: dalWk,
      al: iso(new Date(Date.parse(`${dalWk}T00:00:00Z`) + 6 * 86400000)),
      wk: settimanaIso(dalWk).wk,
    };
  }
  const [g1, m1, g2, m2] = [m[1], m[2], m[3], m[4]].map(Number);
  if (m1 < 1 || m1 > 12 || m2 < 1 || m2 > 12 || g1 < 1 || g1 > 31 || g2 < 1 || g2 > 31) return null;
  const wkFile = Number(nomeFile.match(/w(?:ee)?k\s*[-_]?\s*(\d{1,2})/i)?.[1] ?? NaN);
  const annoOggi = Number(oggi.slice(0, 4));
  const candidati = [annoOggi, annoOggi - 1, annoOggi + 1];
  const limite = iso(new Date(Date.parse(`${oggi}T00:00:00Z`) + 45 * 86400000));
  let anno = candidati.find((a) => settimanaIso(iso(giorno(a, m1, g1))).wk === wkFile);
  if (anno === undefined)
    anno = candidati.find((a) => iso(giorno(a, m1, g1)) <= limite) ?? annoOggi;
  const dal = iso(giorno(anno, m1, g1));
  // La settimana a cavallo dell'anno finisce nell'anno dopo.
  const al = iso(giorno(m2 < m1 ? anno + 1 : anno, m2, g2));
  return { dal, al, wk: settimanaIso(dal).wk };
}

const DATA = /^\d{4}-\d{2}-\d{2}$/;

/** Ripulisce una settimana arrivata dal browser prima di salvarla. */
export function validaSettimana(input: unknown): SettimanaCosti {
  const o = (input ?? {}) as Partial<SettimanaCosti>;
  const testo = (v: unknown, max: number) =>
    String(v ?? "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, max);
  const dal = testo(o.dal, 10);
  const al = testo(o.al, 10);
  if (!DATA.test(dal) || !DATA.test(al)) throw new Error("Date della settimana non valide.");
  if (al < dal) throw new Error("La settimana finisce prima di cominciare.");
  const voci = (l: unknown): VoceCosti[] =>
    (Array.isArray(l) ? l : []).slice(0, 60).flatMap((v): VoceCosti[] => {
      const voce = normVoce((v as VoceCosti)?.voce).slice(0, 60);
      const importo = Number((v as VoceCosti)?.importo);
      if (!voce || !Number.isFinite(importo) || Math.abs(importo) > 1e9) return [];
      const riga: VoceCosti = { voce, importo: r4(importo) };
      const nota = testo((v as VoceCosti)?.nota, 200);
      if (nota) riga.nota = nota;
      const fonte = (v as VoceCosti)?.fonte;
      if (fonte && FONTI_VOCE.includes(fonte)) riga.fonte = fonte;
      const mensile = Number((v as VoceCosti)?.mensile);
      if ((v as VoceCosti)?.mensile != null && Number.isFinite(mensile) && Math.abs(mensile) < 1e9)
        riga.mensile = r4(mensile);
      return [riga];
    });
  const visti = new Set<string>();
  const appalti: AppaltoCosti[] = (Array.isArray(o.appalti) ? o.appalti : [])
    .slice(0, 80)
    .flatMap((a): AppaltoCosti[] => {
      const nome = testo(a?.nome, 60);
      const k = chiaveAppalto(nome);
      if (!k || visti.has(k)) return [];
      visti.add(k);
      const giorni = Number(a?.giorni);
      const km = Number(a?.km);
      const riga: AppaltoCosti = {
        nome,
        giorni: giorni > 0 && giorni <= 31 ? giorni : undefined,
        ricavi: voci(a?.ricavi),
        costi: voci(a?.costi),
      };
      if (a?.km != null && Number.isFinite(km) && km >= 0 && km < 1e8) riga.km = r4(km);
      if (a?.kmCopiati) riga.kmCopiati = true;
      if (a?.paghe != null) riga.paghe = testo(a.paghe, 80);
      return [riga];
    });
  if (!appalti.length) throw new Error("Nel file non c'è nessun foglio di appalto.");
  return {
    id: dal,
    wk: settimanaIso(dal).wk,
    dal,
    al,
    etichetta: testo(o.etichetta, 80),
    file: testo(o.file, 120),
    caricataIl: "",
    caricataDa: "",
    origine: o.origine === "portale" ? "portale" : undefined,
    appalti,
  };
}

/** Mette (o rimette) una settimana nel db: stessa data d'inizio = sostituita. */
export function unisciSettimana(db: CostiAppaltiDb, s: SettimanaCosti): CostiAppaltiDb {
  const settimane = [s, ...db.settimane.filter((x) => x.id !== s.id)]
    .sort((a, b) => b.dal.localeCompare(a.dal))
    .slice(0, 260);
  return { ...db, settimane };
}

/** Tutti gli appalti comparsi nelle settimane, col nome più recente. */
export function appaltiDi(settimane: readonly SettimanaCosti[]): string[] {
  const per = new Map<string, string>();
  for (const s of settimane)
    for (const a of s.appalti)
      if (!per.has(chiaveAppalto(a.nome))) per.set(chiaveAppalto(a.nome), a.nome);
  return [...per.values()];
}

// --- la settimana preparata dal portale --------------------------------------
// Simone 13/11/2026: "non si può automatizzare? prendere i dati da ciò che
// abbiamo e prevedere il caricamento dei soli dati mancanti". Dai quattro
// file (WK41, 43, 44, 45) il foglio è fatto così, per ogni appalto:
//  - voci MENSILI (stipendi, noleggi, costi accessori, addebiti): un importo
//    del mese riportato ai giorni della settimana (mese × giorni / giorni del
//    mese: per questo a ottobre e a novembre la stessa voce cambia di poco);
//  - voci a REGOLA: il FORNITORE con una percentuale nella descrizione
//    ("10%") è quello che resta dei ricavi tolti il margine e gli altri
//    costi; con "0,85 € per km" ricavo e fornitore vengono dai chilometri;
//  - voci della SETTIMANA (ricavi, fuel, pedaggi, fornitori a consuntivo):
//    cambiano ogni volta, sono i soli dati da inserire.
// Gli stipendi il portale li ha già (il costo delle paghe per appalto): il
// file li ricopiava a mano ("dato di settembre").

const VOCI_MENSILI = /^(STIPENDI|NOLEGGI|COSTI ACCESSORI|ADDEBITI)$/;
/** Una voce che nel foglio è un importo mensile riportato alla settimana. */
export const voceMensile = (voce: string) => VOCI_MENSILI.test(voce);

/** Giorni del mese di una data YYYY-MM-DD. */
export const giorniDelMese = (data: string) =>
  new Date(Date.UTC(Number(data.slice(0, 4)), Number(data.slice(5, 7)), 0)).getUTCDate();
/** Giorni di una settimana, estremi compresi. */
export const giorniSettimana = (dal: string, al: string) =>
  Math.round((Date.parse(`${al}T00:00:00Z`) - Date.parse(`${dal}T00:00:00Z`)) / 86400000) + 1;
export const piuGiorni = (data: string, n: number) =>
  iso(new Date(Date.parse(`${data}T00:00:00Z`) + n * 86400000));

const PERCENTO = /^(\d{1,2}(?:[.,]\d+)?)\s*%$/;
const AL_KM = /(\d+(?:[.,]\d+)?)\s*€?\s*(?:per|al|a|\/)\s*km/i;
const dec = (x: string) => Number(x.replace(",", "."));

export type RegolaAppalto =
  { tipo: "margine"; margine: number } | { tipo: "km"; ricavoKm: number; fornitoreKm: number };

/** La regola scritta nel foglio, nella descrizione del FORNITORE: "10%" =
 *  margine fisso; "0,85 € per km" (con "0,95 € per km" sui ricavi) = al km. */
export function regolaDi(a: AppaltoCosti): RegolaAppalto | null {
  const nota = (a.costi.find((v) => v.voce === "FORNITORE")?.nota ?? "").trim();
  const p = nota.match(PERCENTO);
  if (p) return { tipo: "margine", margine: dec(p[1]) / 100 };
  const f = nota.match(AL_KM);
  const r = (a.ricavi.find((v) => v.voce === "RICAVO SERVIZI")?.nota ?? "").match(AL_KM);
  if (f && r) return { tipo: "km", ricavoKm: dec(r[1]), fornitoreKm: dec(f[1]) };
  return null;
}

/** Ricalcola le voci a regola di un appalto dopo che è cambiato qualcosa. */
export function ricalcola(a: AppaltoCosti): AppaltoCosti {
  const regola = regolaDi(a);
  if (!regola) return a;
  if (regola.tipo === "km") {
    if (a.km == null) return a;
    const km = a.km;
    return {
      ...a,
      ricavi: a.ricavi.map((v) =>
        v.voce === "RICAVO SERVIZI"
          ? { ...v, importo: r4(km * regola.ricavoKm), fonte: "regola", mensile: undefined }
          : v,
      ),
      costi: a.costi.map((v) =>
        v.voce === "FORNITORE"
          ? { ...v, importo: r4(km * regola.fornitoreKm), fonte: "regola", mensile: undefined }
          : v,
      ),
    };
  }
  const ricavi = a.ricavi.reduce((t, v) => t + v.importo, 0);
  const altri = a.costi.reduce((t, v) => t + (v.voce === "FORNITORE" ? 0 : v.importo), 0);
  return {
    ...a,
    costi: a.costi.map((v) =>
      v.voce === "FORNITORE"
        ? {
            ...v,
            importo: r4(ricavi * (1 - regola.margine) - altri),
            fonte: "regola",
            mensile: undefined,
          }
        : v,
    ),
  };
}

/** Il costo delle paghe del portale, mese per mese, per etichetta (l'appalto
 *  come lo scrive il consulente: "ZINGALI", "UNIVEX SAVONA"). */
export interface PagheMese {
  mese: string;
  etichette: { nome: string; dipendenti: number; costo: number }[];
}

const MESI_NOME = [
  "gennaio",
  "febbraio",
  "marzo",
  "aprile",
  "maggio",
  "giugno",
  "luglio",
  "agosto",
  "settembre",
  "ottobre",
  "novembre",
  "dicembre",
];
export const nomeMese = (mese: string) =>
  `${MESI_NOME[Number(mese.slice(5, 7)) - 1] ?? mese} ${mese.slice(0, 4)}`;

/** Le paghe di un'etichetta: il mese più recente che ce l'ha, non dopo il
 *  mese della settimana. */
export function pagheDi(
  paghe: readonly PagheMese[],
  etichetta: string,
  entroMese: string,
): { mese: string; costo: number; dipendenti: number } | null {
  const k = chiaveAppalto(etichetta);
  if (!k) return null;
  const mesi = [...paghe]
    .filter((m) => m.mese <= entroMese)
    .sort((a, b) => b.mese.localeCompare(a.mese));
  for (const m of mesi) {
    const e = m.etichette.find((x) => chiaveAppalto(x.nome) === k);
    if (e) return { mese: m.mese, costo: e.costo, dipendenti: e.dipendenti };
  }
  return null;
}

/** L'etichetta delle paghe con lo stesso nome dell'appalto, se c'è. */
export function etichettaPagheDi(nomeAppalto: string, paghe: readonly PagheMese[]): string {
  const k = chiaveAppalto(nomeAppalto);
  for (const m of [...paghe].sort((a, b) => b.mese.localeCompare(a.mese))) {
    const e = m.etichette.find((x) => chiaveAppalto(x.nome) === k);
    if (e) return e.nome;
  }
  return "";
}

const conVoce = (
  l: readonly VoceCosti[],
  voce: string,
  cambia: (v: VoceCosti) => VoceCosti,
): VoceCosti[] =>
  l.some((v) => v.voce === voce)
    ? l.map((v) => (v.voce === voce ? cambia(v) : v))
    : [...l, cambia({ voce, importo: 0 })];

/** Sceglie da quale etichetta delle paghe vengono gli stipendi dell'appalto
 *  ("" = nessuna: lo stipendio torna una voce mensile qualsiasi). */
export function impostaPaghe(
  a: AppaltoCosti,
  etichetta: string,
  paghe: readonly PagheMese[],
  dal: string,
  al: string,
): AppaltoCosti {
  const p = etichetta ? pagheDi(paghe, etichetta, dal.slice(0, 7)) : null;
  if (!p) {
    return ricalcola({
      ...a,
      paghe: "",
      costi: a.costi.map((v) =>
        v.voce === "STIPENDI" && v.fonte === "paghe" ? { ...v, fonte: "mano" } : v,
      ),
    });
  }
  const importo = r4((p.costo * giorniSettimana(dal, al)) / giorniDelMese(dal));
  return ricalcola({
    ...a,
    paghe: etichetta,
    costi: conVoce(a.costi, "STIPENDI", (v) => ({
      ...v,
      importo,
      mensile: r4(p.costo),
      fonte: "paghe",
      nota: `paghe di ${nomeMese(p.mese)} · ${p.dipendenti} dip.`,
    })),
  });
}

/** Scrive a mano un importo (ricavo o costo) e ricalcola le regole. */
export function impostaVoce(
  a: AppaltoCosti,
  lato: "ricavi" | "costi",
  voce: string,
  importo: number,
): AppaltoCosti {
  const cambia = (v: VoceCosti): VoceCosti => ({
    ...v,
    importo: r4(importo),
    fonte: "mano",
    mensile: undefined,
  });
  return ricalcola({ ...a, [lato]: conVoce(a[lato], voce, cambia) });
}

/** Cambia la descrizione di una voce (può accendere o spegnere una regola). */
export function impostaNota(
  a: AppaltoCosti,
  lato: "ricavi" | "costi",
  voce: string,
  nota: string,
): AppaltoCosti {
  const pulita = nota.replace(/\s+/g, " ").trim().slice(0, 200);
  return ricalcola({
    ...a,
    [lato]: conVoce(a[lato], voce, (v) => ({ ...v, nota: pulita || undefined })),
  });
}

/** I chilometri della settimana, per gli appalti a tariffa al km. */
export function impostaKm(a: AppaltoCosti, km: number): AppaltoCosti {
  return ricalcola({ ...a, km: r4(Math.max(0, km)), kmCopiati: undefined });
}

/** Prepara la settimana dopo `prec` con quello che il portale sa già:
 *  le voci mensili riportate ai giorni della settimana nuova, gli stipendi
 *  dalle paghe, le regole ricalcolate. Quello che cambia ogni settimana
 *  (ricavi, fuel, pedaggi, fornitori a consuntivo, km) resta uguale alla
 *  settimana prima e segnato "copiata": sono i soli dati da inserire. */
export function preparaSettimana(
  prec: SettimanaCosti,
  paghe: readonly PagheMese[],
): { dal: string; al: string; appalti: AppaltoCosti[] } {
  const giorniPrec = giorniSettimana(prec.dal, prec.al);
  const dal = piuGiorni(prec.dal, giorniPrec);
  const al = piuGiorni(prec.al, giorniPrec);
  const giorni = giorniSettimana(dal, al);
  const appalti = prec.appalti.map((p): AppaltoCosti => {
    const porta = (v: VoceCosti): VoceCosti => {
      if (voceMensile(v.voce) && Math.abs(v.importo) > 0.00005) {
        const mensile = v.mensile ?? r4((v.importo * giorniDelMese(prec.dal)) / giorniPrec);
        return {
          ...v,
          mensile,
          importo: r4((mensile * giorni) / giorniDelMese(dal)),
          fonte: "mensile",
        };
      }
      return {
        voce: v.voce,
        importo: v.importo,
        nota: v.nota,
        fonte: Math.abs(v.importo) > 0.00005 ? "copiata" : undefined,
      };
    };
    let a: AppaltoCosti = {
      nome: p.nome,
      giorni,
      ricavi: p.ricavi.map(porta),
      costi: p.costi.map(porta),
      km: p.km,
      paghe: p.paghe,
    };
    const regola = regolaDi(a);
    if (regola?.tipo === "km") {
      // I km della settimana prima (dai ricavi, se non erano scritti).
      const ricavo = a.ricavi.find((v) => v.voce === "RICAVO SERVIZI")?.importo ?? 0;
      a.km = a.km ?? (regola.ricavoKm > 0 ? Math.round(ricavo / regola.ricavoKm) : 0);
      a.kmCopiati = true;
    }
    // Stipendi dalle paghe: l'etichetta scelta la volta prima; la prima
    // volta quella con lo stesso nome dell'appalto, ma solo se nel foglio gli
    // stipendi c'erano (dove il foglio li tiene a zero non si aggiungono).
    const stipendi = p.costi.find((v) => v.voce === "STIPENDI")?.importo ?? 0;
    const etichetta =
      a.paghe ?? (Math.abs(stipendi) > 0.005 ? etichettaPagheDi(p.nome, paghe) : "");
    if (etichetta) a = impostaPaghe(a, etichetta, paghe, dal, al);
    return ricalcola(a);
  });
  return { dal, al, appalti };
}

/** Quanti dati della settimana restano da confermare (copiati dalla prima). */
export function daConfermare(appalti: readonly AppaltoCosti[]): number {
  let n = 0;
  for (const a of appalti) {
    if (a.kmCopiati) n++;
    for (const v of [...a.ricavi, ...a.costi]) if (v.fonte === "copiata") n++;
  }
  return n;
}
