// DR Portal — Regole di business per il Rendiconto (riscontro settimanale).
// Logica PURA e testabile. Il modello concordato:
// - riscontro SETTIMANALE (Lun–Dom) a monte ore = OreSettimanali;
// - sabato riempie le ore mancanti; oltre il previsto = straordinario;
// - domenica = SEMPRE straordinario;
// - un giorno di ferie/malattia riduce il previsto di OreSettimanali/5;
// - un permesso (ore) riduce il previsto delle sue ore;
// - smart working si timbra come un giorno normale.

import { MAX_TURNO_ORE, giornoLocale, type EventoTimbratura } from "./presenze-logic";

export function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

// Data locale in formato YYYY-MM-DD (evita gli shift di fuso di toISOString).
export function ymd(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const g = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${g}`;
}

// Giorno della settimana ISO: 1=Lun … 7=Dom.
export function isoDow(dateStr: string): number {
  const d = new Date(`${dateStr}T00:00:00`).getDay(); // 0=Dom..6=Sab
  return d === 0 ? 7 : d;
}

// Lunedì (YYYY-MM-DD) della settimana che contiene la data.
export function lunediDellaSettimana(dateStr: string): string {
  const d = new Date(`${dateStr}T00:00:00`);
  const dow = d.getDay() === 0 ? 7 : d.getDay();
  d.setDate(d.getDate() - (dow - 1));
  return ymd(d);
}

// --- Ore notturne (riepilogo presenze, v1.88.0) ------------------------------
// Il foglio di Monica ("PRESENZE Zingali SUDDIVISO") ha per ogni dipendente
// la riga "Tot" e la riga "Not": le ore cadute nella fascia notturna del
// CCNL logistica, 22:00–06:00 ora di Roma. Si calcolano sui segmenti in
// servizio (le DataOra sono UTC: serve lo scostamento di Roma, +60 o +120).
export const NOTTE_INIZIO_MIN = 22 * 60;
export const NOTTE_FINE_MIN = 6 * 60;
const fmtRomaParti = new Intl.DateTimeFormat("en-US", {
  timeZone: "Europe/Rome",
  hour12: false,
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
  second: "2-digit",
});

/** Scostamento in minuti tra l'ora di Roma e l'UTC in quell'istante. */
export function offsetRomaMinuti(ms: number): number {
  const p: Record<string, string> = {};
  for (const x of fmtRomaParti.formatToParts(new Date(ms))) p[x.type] = x.value;
  const locale = Date.UTC(
    Number(p.year),
    Number(p.month) - 1,
    Number(p.day),
    Number(p.hour) % 24, // alcuni motori scrivono "24" a mezzanotte
    Number(p.minute),
    Number(p.second),
  );
  return Math.round((locale - Math.floor(ms / 1000) * 1000) / 60000);
}

const sovrapposizione = (a: number, b: number, c: number, d: number) =>
  Math.max(0, Math.min(b, d) - Math.max(a, c));

/** Minuti dell'intervallo [daMs, aMs) che cadono nella fascia 22:00–06:00
 *  di Roma, anche a cavallo di mezzanotte. Se dentro l'intervallo cambia
 *  l'ora legale, si spezza a metà e si somma. */
export function minutiNotturni(daMs: number, aMs: number): number {
  if (aMs <= daMs) return 0;
  const off = offsetRomaMinuti(daMs);
  if (off !== offsetRomaMinuti(aMs - 1)) {
    if (aMs - daMs < 120_000) return 0;
    const meta = daMs + Math.floor((aMs - daMs) / 2);
    return minutiNotturni(daMs, meta) + minutiNotturni(meta, aMs);
  }
  // Minuti "locali" dall'epoca: le finestre notturne sono fisse per giorno.
  const a = daMs / 60000 + off;
  const b = aMs / 60000 + off;
  let tot = 0;
  for (let g = Math.floor(a / 1440); g <= Math.floor((b - 1) / 1440); g++) {
    const base = g * 1440;
    tot += sovrapposizione(a, b, base, base + NOTTE_FINE_MIN);
    tot += sovrapposizione(a, b, base + NOTTE_INIZIO_MIN, base + 1440);
  }
  return Math.round(tot * 100) / 100;
}

export interface OreDaTurni {
  /** Ore (decimali) per GIORNO DI INIZIO TURNO: il notturno 22→02 conta
   *  tutto sul giorno dell'entrata. */
  oreGiorno: Map<string, number>;
  /** Ore (decimali) nella fascia notturna 22–06, stesso giorno di attribuzione. */
  notteGiorno: Map<string, number>;
  /** Giorni con un turno dimenticato (aperto oltre MAX_TURNO_ORE): ore non
   *  calcolabili in modo attendibile → da sanare prima del rendiconto. */
  giorniNonChiusi: Set<string>;
}

// Ore lavorate dal FLUSSO COMPLETO degli eventi di un dipendente, a segmenti
// in servizio (aperti da entrata/fine-pausa, chiusi da uscita/inizio-pausa),
// anche a cavallo di mezzanotte. Un turno ancora in corso (entro il tetto)
// non produce né ore né anomalia: si valuterà quando chiude.
export function orePerGiornoDaTurni(
  eventi: { evento: EventoTimbratura; ora: string }[],
  now = new Date(),
): OreDaTurni {
  const maxMs = MAX_TURNO_ORE * 3600_000;
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const oreMs = new Map<string, number>();
  const notteMin = new Map<string, number>();
  const giorniNonChiusi = new Set<string>();
  let apertura: { ms: number; giorno: string } | null = null;
  // Giorno del TURNO in corso (quello dell'entrata): anche i segmenti dopo
  // una pausa che scavalca la mezzanotte restano attribuiti lì.
  let giornoTurno: string | null = null;
  for (const e of sorted) {
    const ms = new Date(e.ora).getTime();
    if (e.evento === "entrata" || e.evento === "fine-pausa") {
      if (e.evento === "entrata" && giornoTurno == null) giornoTurno = giornoLocale(e.ora);
      if (apertura == null) {
        apertura = { ms, giorno: giornoTurno ?? giornoLocale(e.ora) };
      } else if (ms - apertura.ms > maxMs) {
        giorniNonChiusi.add(apertura.giorno);
        giornoTurno = giornoLocale(e.ora);
        apertura = { ms, giorno: giornoTurno };
      }
    } else if (apertura != null) {
      const durata = ms - apertura.ms;
      if (durata > maxMs) giorniNonChiusi.add(apertura.giorno);
      else {
        oreMs.set(apertura.giorno, (oreMs.get(apertura.giorno) ?? 0) + Math.max(0, durata));
        notteMin.set(
          apertura.giorno,
          (notteMin.get(apertura.giorno) ?? 0) + minutiNotturni(apertura.ms, ms),
        );
      }
      apertura = null;
      if (e.evento === "uscita") giornoTurno = null; // il turno si chiude qui
    }
  }
  if (apertura != null && now.getTime() - apertura.ms > maxMs) giorniNonChiusi.add(apertura.giorno);
  const oreGiorno = new Map<string, number>();
  for (const [g, ms] of oreMs) oreGiorno.set(g, round2(ms / 3600000));
  const notteGiorno = new Map<string, number>();
  for (const [g, min] of notteMin) if (min > 0) notteGiorno.set(g, round2(min / 60));
  return { oreGiorno, notteGiorno, giorniNonChiusi };
}

/** Gruppo del riepilogo presenze (Simone 01/10): la colonna Reparto
 *  dell'anagrafica divide una sede in sotto-sedi/reparti ("Pavia", "Torino",
 *  "Cerro - Ufficio", "Cerro - Magazzino"); l'etichetta è "<appalto> <reparto>"
 *  come nei fogli dell'ufficio (ZINGALI PAVIA, ZINGALI CERRO - UFFICIO).
 *  Senza reparto resta la sede. */
export function etichettaGruppo(sede: string, appalto: string, reparto: string): string {
  const r = (reparto ?? "").trim();
  if (!r) return (sede ?? "").trim();
  return `${(appalto ?? "").trim() || (sede ?? "").trim()} ${r}`.trim();
}

/** Arrotondamento ai 15 minuti (al più vicino): 8,37 → 8,25; 8,38 → 8,5.
 *  È la convenzione dei fogli presenze di Monica (quarti d'ora). */
export function arrotondaQuarto(ore: number): number {
  return Math.round(ore * 4) / 4;
}

// Ore previste della settimana: monte ore contrattuale meno le assenze
// giustificate. `giorniAssenza` = ferie + malattia nella settimana;
// `orePermesso` = ore di permesso approvate nella settimana.
export function orePrevisteSettimana(
  oreSettimanali: number,
  giorniAssenza: number,
  orePermesso: number,
): number {
  const perGiorno = oreSettimanali / 5;
  return Math.max(0, round2(oreSettimanali - giorniAssenza * perGiorno - orePermesso));
}

// Straordinario della settimana: ore Lun–Sab oltre il previsto, più tutte le
// ore di domenica (sempre straordinario).
export function straordinarioSettimana(
  oreLunSab: number,
  oreDomenica: number,
  orePreviste: number,
): number {
  const extra = Math.max(0, oreLunSab - orePreviste);
  return round2(extra + oreDomenica);
}
