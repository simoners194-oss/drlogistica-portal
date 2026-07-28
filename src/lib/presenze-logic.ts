// DR Portal — Regole di business per il Modulo Presenze.
// Centralizza la macchina a stati delle timbrature e il calcolo delle ore
// lavorate, così client e server (SharePoint) applicano le stesse regole.

import type { Timbratura } from "./mock-data";

export type EventoTimbratura = Timbratura["tipo"];

// Tutti i tipi di evento ESISTENTI (i dati storici contengono anche le pause).
export const EVENTI: EventoTimbratura[] = ["entrata", "inizio-pausa", "fine-pausa", "uscita"];
// Eventi che si possono ANCORA registrare: modello semplificato a due tasti.
// La pausa non è più un evento: chi stacca preme Uscita e al rientro Entrata.
export const EVENTI_ATTIVI: EventoTimbratura[] = ["entrata", "uscita"];

// Finestra (minuti) entro cui il dipendente può annullare da solo l'ULTIMA
// timbratura di oggi ("ho premuto il tasto sbagliato"). Oltre, serve
// l'operatore (Gestione timbrature → Turni del giorno).
export const UNDO_TIMBRATURA_MINUTI = 5;

// Macchina a stati: dato l'ultimo evento (o null se nessuna timbratura oggi),
// quali eventi sono ammessi. Più turni Entrata→Uscita nello stesso giorno
// sono LEGITTIMI (la pausa è un turno che si chiude e uno che si riapre).
export function nextAllowedEvents(last: EventoTimbratura | null): EventoTimbratura[] {
  if (last === null || last === "uscita") return ["entrata"];
  // In servizio (entrata/fine-pausa) o pausa legacy aperta: si può solo uscire
  // (l'uscita chiude anche una vecchia pausa rimasta aperta).
  return ["uscita"];
}

export function isTransitionAllowed(
  evento: EventoTimbratura,
  last: EventoTimbratura | null,
): boolean {
  return nextAllowedEvents(last).includes(evento);
}

// Messaggio esplicativo per il pulsante disabilitato.
export function reasonNotAllowed(
  evento: EventoTimbratura,
  last: EventoTimbratura | null,
): string | null {
  if (isTransitionAllowed(evento, last)) return null;
  if (evento === "entrata") return "Sei già in servizio: per staccare premi Uscita.";
  if (evento === "uscita")
    return last === null
      ? "Devi prima registrare l'entrata."
      : "Sei fuori servizio: al rientro premi Entrata.";
  return BLOCK_MESSAGE;
}

export const BLOCK_MESSAGE = "Timbratura non consentita in questo momento.";

// ---------------------------------------------------------------------------
// Calcolo ore lavorate
// ---------------------------------------------------------------------------

export interface OreOggi {
  entrataOra: string | null; // ISO — prima entrata del giorno
  uscitaOra: string | null; // ISO — ultima uscita, se ora si è fuori servizio
  /** Minuti di stacco DENTRO la giornata (tra un'uscita e il rientro
   *  successivo, o le vecchie pause a eventi). */
  pausaMinuti: number;
  oreLavorateMinuti: number;
  oltreOrarioMinuti: number; // solo se supera 8h
  /** Vecchia pausa a eventi ancora aperta (solo dati storici). */
  inPausa: boolean;
  /** Fuori servizio ADESSO (ultimo evento = uscita). Non è definitivo: una
   *  nuova Entrata apre un altro turno nello stesso giorno. */
  chiusa: boolean;
}

export const SOGLIA_ORE_MIN = 8 * 60;

// Calcolo a SEGMENTI: la giornata è una sequenza di turni "in servizio"
// (aperti da entrata/fine-pausa, chiusi da uscita/inizio-pausa). Le ore
// lavorate sono la somma dei segmenti; gli stacchi tra un segmento e il
// successivo sono la "pausa". Copre sia il modello nuovo a due tasti sia i
// dati storici con gli eventi pausa.
export function computeOreOggi(events: Timbratura[], now = new Date()): OreOggi {
  const sorted = [...events].sort((a, b) => a.ora.localeCompare(b.ora));
  let lavoroMs = 0;
  let pausaMs = 0;
  let dentroDa: number | null = null; // inizio del segmento in corso
  let fuoriDa: number | null = null; // inizio dello stacco in corso
  let entrataOra: string | null = null;
  let uscitaOra: string | null = null;
  for (const e of sorted) {
    const ms = new Date(e.ora).getTime();
    if (e.tipo === "entrata" || e.tipo === "fine-pausa") {
      if (dentroDa == null) {
        dentroDa = ms;
        if (fuoriDa != null) {
          pausaMs += Math.max(0, ms - fuoriDa); // lo stacco si chiude: era pausa
          fuoriDa = null;
        }
      }
      if (entrataOra == null && e.tipo === "entrata") entrataOra = e.ora;
    } else {
      // uscita | inizio-pausa: chiude il segmento in corso
      if (dentroDa != null) {
        lavoroMs += Math.max(0, ms - dentroDa);
        dentroDa = null;
        fuoriDa = ms;
      }
      if (e.tipo === "uscita") uscitaOra = e.ora;
    }
  }
  const last = sorted.length ? sorted[sorted.length - 1].tipo : null;
  if (dentroDa != null) lavoroMs += Math.max(0, now.getTime() - dentroDa); // turno in corso
  const inPausa = last === "inizio-pausa"; // solo legacy
  if (inPausa && fuoriDa != null) pausaMs += Math.max(0, now.getTime() - fuoriDa);
  const chiusa = last === "uscita";
  const pausaMinuti = Math.floor(pausaMs / 60000);
  const oreLavorateMinuti = Math.max(0, Math.floor(lavoroMs / 60000));
  const oltreOrarioMinuti = Math.max(0, oreLavorateMinuti - SOGLIA_ORE_MIN);
  return {
    entrataOra,
    uscitaOra: chiusa ? uscitaOra : null,
    pausaMinuti,
    oreLavorateMinuti,
    oltreOrarioMinuti,
    inPausa,
    chiusa,
  };
}

export function formatDurata(minuti: number): string {
  const h = Math.floor(minuti / 60);
  const m = Math.max(0, minuti % 60);
  return `${h}h ${m.toString().padStart(2, "0")}m`;
}

export function lastEvento(events: Timbratura[]): EventoTimbratura | null {
  if (events.length === 0) return null;
  const sorted = [...events].sort((a, b) => a.ora.localeCompare(b.ora));
  return sorted[sorted.length - 1].tipo;
}

// ---------------------------------------------------------------------------
// Rilevazione anomalie giornaliere (Sprint 3, on-read)
// ---------------------------------------------------------------------------
export type TipoAnomalia = "turno-non-chiuso" | "pausa-non-chiusa" | "senza-stacco";

export const LABEL_ANOMALIA: Record<TipoAnomalia, string> = {
  "turno-non-chiuso": "Turno non chiuso (manca l'uscita)",
  "pausa-non-chiusa": "Pausa non chiusa (manca la fine pausa)",
  "senza-stacco": "Giornata lunga senza stacco (informativa)",
};

// Sopra questa durata continuativa, una giornata senza alcuno stacco viene
// segnalata come anomalia INFORMATIVA (nessun blocco): con i due tasti la
// pausa non è più dichiarata, si vede solo come uscita+rientro.
export const SENZA_STACCO_MIN_ORE = 6;

// Rileva le anomalie di una giornata CONCLUSA dai suoi eventi (con orario):
// - turno non chiuso: l'ultimo evento lascia il dipendente in servizio;
// - pausa non chiusa: più inizio-pausa che fine-pausa (solo dati storici);
// - senza stacco (informativa): giornata chiusa, un unico turno continuativo
//   oltre SENZA_STACCO_MIN_ORE ore, nessuna uscita/rientro né pausa in mezzo.
// `rilevaPausa=false` per chi non ha diritto alla pausa (es. part-time ≤16h).
export function anomalieDelGiorno(
  eventi: { evento: EventoTimbratura; ora: string }[],
  opts: { rilevaPausa: boolean },
): TipoAnomalia[] {
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const tipi = sorted.map((e) => e.evento);
  const last = tipi.length ? tipi[tipi.length - 1] : null;
  const out: TipoAnomalia[] = [];
  if (last === "entrata" || last === "fine-pausa") out.push("turno-non-chiuso");
  const ip = tipi.filter((e) => e === "inizio-pausa").length;
  const fp = tipi.filter((e) => e === "fine-pausa").length;
  if (opts.rilevaPausa && ip > fp) out.push("pausa-non-chiusa");
  if (
    opts.rilevaPausa &&
    last === "uscita" &&
    tipi.filter((e) => e === "entrata").length === 1 &&
    tipi.filter((e) => e === "uscita").length === 1 &&
    ip === 0
  ) {
    const entrata = sorted.find((e) => e.evento === "entrata");
    const uscita = sorted.find((e) => e.evento === "uscita");
    if (entrata && uscita) {
      const oreContinuative =
        (new Date(uscita.ora).getTime() - new Date(entrata.ora).getTime()) / 3600000;
      if (oreContinuative > SENZA_STACCO_MIN_ORE) out.push("senza-stacco");
    }
  }
  return out;
}
