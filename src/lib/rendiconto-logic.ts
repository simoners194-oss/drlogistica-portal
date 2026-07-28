// DR Portal — Regole di business per il Rendiconto (riscontro settimanale).
// Logica PURA e testabile. Il modello concordato:
// - riscontro SETTIMANALE (Lun–Dom) a monte ore = OreSettimanali;
// - sabato riempie le ore mancanti; oltre il previsto = straordinario;
// - domenica = SEMPRE straordinario;
// - un giorno di ferie/malattia riduce il previsto di OreSettimanali/5;
// - un permesso (ore) riduce il previsto delle sue ore;
// - smart working si timbra come un giorno normale.

import type { EventoTimbratura } from "./presenze-logic";

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

// Ore lavorate (decimali) di una giornata dai suoi eventi, sommando i
// SEGMENTI in servizio (aperti da entrata/fine-pausa, chiusi da
// uscita/inizio-pausa): copre sia il modello a due tasti con più turni al
// giorno sia i dati storici con le pause a eventi. Ritorna null se la
// giornata non è chiusa (l'ultimo evento lascia il dipendente in servizio):
// non calcolabile in modo attendibile → va corretta prima del rendiconto.
export function oreLavorateGiorno(
  eventi: { evento: EventoTimbratura; ora: string }[],
): number | null {
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  if (!sorted.length) return null;
  const last = sorted[sorted.length - 1].evento;
  if (last === "entrata" || last === "fine-pausa") return null; // giornata aperta
  let lavoroMs = 0;
  let dentroDa: number | null = null;
  for (const e of sorted) {
    const ms = new Date(e.ora).getTime();
    if (e.evento === "entrata" || e.evento === "fine-pausa") {
      if (dentroDa == null) dentroDa = ms;
    } else if (dentroDa != null) {
      lavoroMs += Math.max(0, ms - dentroDa);
      dentroDa = null;
    }
  }
  return round2(lavoroMs / 3600000);
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
