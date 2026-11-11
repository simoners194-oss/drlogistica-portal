// DR Portal — Finanza → Flussi → Cassa proprietà: i FILTRI del foglio
// (Simone 11/11/2026: "metti filtri su tutte le voci di cassa proprietà").
//
// Restringono le voci VISIBILI: non si salvano, non cambiano i totali del
// foglio né l'export. Le uscite si filtrano PRIMA del raggruppamento per
// tipologia o cliente, così i gruppi sommano solo le voci filtrate.
// Logica pura (niente React): la vista la usa su entrate e uscite.
import type { PrioritaCassa, SecchioCassa } from "./cassa-logic";

export type BloccoEntrate = "gia" | "entranti" | "sollecitare" | "standby";

export interface FiltriFoglio {
  /** Entrate: testo cercato (cliente, fattura, importo, nota). */
  eTesto: string;
  /** Entrate: un blocco solo; vuoto = tutti. */
  eBlocco: "" | BloccoEntrate;
  /** Uscite: testo cercato (fornitore, voce, causale, importo, nota). */
  uTesto: string;
  /** Colonne delle uscite da mostrare; vuoto = tutte. */
  uColonne: SecchioCassa[];
  uTipologia: string;
  uCliente: string;
  /** "con" = una priorità qualsiasi. */
  uPriorita: "" | PrioritaCassa | "con";
  /** Solo le uscite con una nota scritta. */
  uNota: boolean;
}

export const FILTRI_VUOTI: FiltriFoglio = {
  eTesto: "",
  eBlocco: "",
  uTesto: "",
  uColonne: [],
  uTipologia: "",
  uCliente: "",
  uPriorita: "",
  uNota: false,
};

/** Valore del filtro per le voci senza tipologia o senza cliente. */
export const SENZA = "—";
export const COLONNE_USCITE = ["ora", "prossimi", "standby", "standby2"] as const;

/** Quello che serve di una riga del foglio per filtrarla. */
export interface VoceFiltrabile {
  nome: string;
  dettaglio: string;
  nota: string;
  importo: number;
  secchio: SecchioCassa;
  attesa: boolean;
  parcheggiata: boolean;
  daSollecitare: boolean;
  priorita?: PrioritaCassa;
  tipologia?: string;
  cliente?: string;
  quote?: readonly { appalto: string; quota: number }[];
  membriClienti?: readonly string[];
  membri?: readonly { nome: string }[];
}

const fmtIt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

export const sommaCelle = (l: readonly { importo: number }[]) =>
  Math.round(l.reduce((s, c) => s + c.importo, 0) * 100) / 100;

/** Minuscolo e senza accenti: "Società" si trova scrivendo "societa". */
export const normTesto = (s: string) => s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");

/** Tutte le parole cercate devono comparire nella voce: nome, dettaglio,
 *  nota, società o fornitori dentro la riga, importo (scritto all'italiana
 *  o col punto decimale). */
export function contieneTesto(c: VoceFiltrabile, testo: string): boolean {
  const parole = normTesto(testo).split(/\s+/).filter(Boolean);
  if (!parole.length) return true;
  const dove = normTesto(
    [
      c.nome,
      c.dettaglio,
      c.nota,
      ...(c.membriClienti ?? []),
      ...(c.membri ?? []).map((m) => m.nome),
      fmtIt(c.importo),
      String(c.importo),
    ].join(" "),
  );
  return parole.every((p) => dove.includes(p));
}

/** Il blocco delle entrate in cui sta la riga. */
export function bloccoEntrata(c: {
  attesa: boolean;
  parcheggiata: boolean;
  daSollecitare: boolean;
}): BloccoEntrate {
  if (!c.attesa) return "gia";
  if (c.parcheggiata) return "standby";
  return c.daSollecitare ? "sollecitare" : "entranti";
}

export const tipologiaDi = (u: { tipologia?: string }) => u.tipologia?.trim() || SENZA;

/** Clienti/appalti di un'uscita: tutti quelli della ripartizione, se c'è. */
export const clientiDi = (u: {
  cliente?: string;
  quote?: readonly { appalto: string; quota: number }[];
}): string[] =>
  u.quote && u.quote.length > 1
    ? u.quote.map((q) => q.appalto.trim() || SENZA)
    : [u.cliente?.trim() || SENZA];

export const filtroEntrateAttivo = (f: FiltriFoglio) => Boolean(f.eTesto.trim() || f.eBlocco);

export const filtroUsciteAttivo = (f: FiltriFoglio) =>
  Boolean(
    f.uTesto.trim() || f.uColonne.length || f.uTipologia || f.uCliente || f.uPriorita || f.uNota,
  );

export function filtraEntrate<T extends VoceFiltrabile>(
  entrate: readonly T[],
  f: FiltriFoglio,
): T[] {
  if (!filtroEntrateAttivo(f)) return [...entrate];
  return entrate.filter(
    (c) => (!f.eBlocco || bloccoEntrata(c) === f.eBlocco) && contieneTesto(c, f.eTesto),
  );
}

export function filtraUscite<T extends VoceFiltrabile>(uscite: readonly T[], f: FiltriFoglio): T[] {
  if (!filtroUsciteAttivo(f)) return [...uscite];
  return uscite.filter(
    (u) =>
      (!f.uColonne.length || f.uColonne.includes(u.secchio)) &&
      (!f.uTipologia || tipologiaDi(u) === f.uTipologia) &&
      (!f.uCliente || clientiDi(u).includes(f.uCliente)) &&
      (f.uPriorita === "con" ? Boolean(u.priorita) : !f.uPriorita || u.priorita === f.uPriorita) &&
      (!f.uNota || Boolean(u.nota.trim())) &&
      contieneTesto(u, f.uTesto),
  );
}

/** Valori distinti per una tendina, in ordine alfabetico, "senza" in fondo. */
export const distinti = (valori: readonly string[]): string[] =>
  [...new Set(valori)].sort((a, b) =>
    a === SENZA ? 1 : b === SENZA ? -1 : a.localeCompare(b, "it"),
  );
