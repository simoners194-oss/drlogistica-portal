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
}

export interface AppaltoCosti {
  /** Il nome del foglio: "Postadoc", "Univex Milano"… */
  nome: string;
  /** Riga GG: giorni della settimana lavorati. */
  giorni?: number;
  ricavi: VoceCosti[];
  costi: VoceCosti[];
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
  appalti: AppaltoCosti[];
}

export interface CostiAppaltiDb {
  versione: number;
  /** Dalla più recente. */
  settimane: SettimanaCosti[];
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export const emptyCostiAppaltiDb = (): CostiAppaltiDb => ({ versione: 0, settimane: [] });

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
  const out: RigaFoglioCosti[] = [];
  for (let r = 1; r <= maxRighe; r++) {
    const a = cella("A", r);
    const b = cella("B", r);
    const c = cella("C", r);
    const nota = c ? (c.t === "n" ? String(c.w ?? c.v ?? "") : String(c.v ?? "")) : "";
    out.push([a?.v ?? "", b?.v ?? null, nota]);
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
  if (!m) return null;
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
      const nota = testo((v as VoceCosti)?.nota, 200);
      return [nota ? { voce, importo: r4(importo), nota } : { voce, importo: r4(importo) }];
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
      return [
        {
          nome,
          giorni: giorni > 0 && giorni <= 31 ? giorni : undefined,
          ricavi: voci(a?.ricavi),
          costi: voci(a?.costi),
        },
      ];
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
