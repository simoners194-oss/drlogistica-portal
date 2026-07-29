// DR Portal — Regole di business per il Rendiconto (riscontro settimanale).
// Logica PURA e testabile. Il modello concordato:
// - riscontro SETTIMANALE (Lun–Dom) a monte ore = OreSettimanali;
// - sabato riempie le ore mancanti; oltre il previsto = straordinario;
// - domenica = SEMPRE straordinario;
// - un giorno di ferie/malattia riduce il previsto di OreSettimanali/5;
// - un permesso (ore) riduce il previsto delle sue ore;
// - smart working si timbra come un giorno normale.

import { MAX_TURNO_ORE, type EventoTimbratura } from "./presenze-logic";

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

export interface OreDaTurni {
  /** Ore (decimali) per GIORNO DI INIZIO TURNO: il notturno 22→02 conta
   *  tutto sul giorno dell'entrata. */
  oreGiorno: Map<string, number>;
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
  const giorniNonChiusi = new Set<string>();
  let apertura: { ms: number; giorno: string } | null = null;
  // Giorno del TURNO in corso (quello dell'entrata): anche i segmenti dopo
  // una pausa che scavalca la mezzanotte restano attribuiti lì.
  let giornoTurno: string | null = null;
  for (const e of sorted) {
    const ms = new Date(e.ora).getTime();
    if (e.evento === "entrata" || e.evento === "fine-pausa") {
      if (e.evento === "entrata" && giornoTurno == null) giornoTurno = e.ora.slice(0, 10);
      if (apertura == null) {
        apertura = { ms, giorno: giornoTurno ?? e.ora.slice(0, 10) };
      } else if (ms - apertura.ms > maxMs) {
        giorniNonChiusi.add(apertura.giorno);
        giornoTurno = e.ora.slice(0, 10);
        apertura = { ms, giorno: giornoTurno };
      }
    } else if (apertura != null) {
      const durata = ms - apertura.ms;
      if (durata > maxMs) giorniNonChiusi.add(apertura.giorno);
      else oreMs.set(apertura.giorno, (oreMs.get(apertura.giorno) ?? 0) + Math.max(0, durata));
      apertura = null;
      if (e.evento === "uscita") giornoTurno = null; // il turno si chiude qui
    }
  }
  if (apertura != null && now.getTime() - apertura.ms > maxMs) giorniNonChiusi.add(apertura.giorno);
  const oreGiorno = new Map<string, number>();
  for (const [g, ms] of oreMs) oreGiorno.set(g, round2(ms / 3600000));
  return { oreGiorno, giorniNonChiusi };
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
