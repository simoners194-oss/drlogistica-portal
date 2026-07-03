// DR Portal — Regole di business per il Modulo Presenze.
// Centralizza la macchina a stati delle timbrature e il calcolo delle ore
// lavorate, così client e server (SharePoint) applicano le stesse regole.

import type { Timbratura } from "./mock-data";

export type EventoTimbratura = Timbratura["tipo"];

export const EVENTI: EventoTimbratura[] = [
  "entrata",
  "inizio-pausa",
  "fine-pausa",
  "uscita",
];

// Macchina a stati: dato l'ultimo evento (o null se nessuna timbratura oggi),
// quali eventi sono ammessi.
export function nextAllowedEvents(last: EventoTimbratura | null): EventoTimbratura[] {
  switch (last) {
    case null:
      return ["entrata"];
    case "entrata":
    case "fine-pausa":
      return ["inizio-pausa", "uscita"];
    case "inizio-pausa":
      return ["fine-pausa"];
    case "uscita":
      return [];
  }
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
  if (last === "uscita") return GIORNATA_CHIUSA_MESSAGE;
  switch (evento) {
    case "entrata":
      return last === null
        ? "Timbratura non consentita in questo momento."
        : "Entrata già registrata oggi.";
    case "inizio-pausa":
      return last === null
        ? "Devi prima registrare l'entrata."
        : last === "inizio-pausa"
          ? "Pausa già in corso: registra la fine pausa."
          : "Disponibile solo dopo un'entrata o una fine pausa.";
    case "fine-pausa":
      return last === null
        ? "Devi prima registrare l'entrata."
        : "Disponibile solo se hai una pausa in corso.";
    case "uscita":
      return last === null
        ? "Devi prima registrare l'entrata."
        : "Chiudi prima la pausa in corso, poi registra l'uscita.";
  }
  return "Timbratura non consentita in questo momento.";
}

export const BLOCK_MESSAGE = "Timbratura non consentita in questo momento.";

// Messaggio ufficiale mostrato quando la giornata lavorativa è già stata
// chiusa con l'Uscita: qualunque ulteriore timbratura è vietata al
// dipendente e va gestita dal modulo amministrativo.
export const GIORNATA_CHIUSA_MESSAGE =
  "La giornata lavorativa è già stata chiusa. Per eventuali correzioni contatta il tuo responsabile.";

// ---------------------------------------------------------------------------
// Calcolo ore lavorate
// ---------------------------------------------------------------------------

export interface OreOggi {
  entrataOra: string | null; // ISO
  uscitaOra: string | null; // ISO se giornata chiusa
  pausaMinuti: number;
  oreLavorateMinuti: number;
  oltreOrarioMinuti: number; // solo se supera 8h
  inPausa: boolean;
  chiusa: boolean;
}

export const SOGLIA_ORE_MIN = 8 * 60;

// Ordina gli eventi per data ora crescente. Ipotesi: al massimo una entrata
// e una uscita per giornata; più pause sono ammesse.
export function computeOreOggi(events: Timbratura[], now = new Date()): OreOggi {
  const sorted = [...events].sort((a, b) => a.ora.localeCompare(b.ora));
  const entrata = sorted.find((e) => e.tipo === "entrata");
  const uscita = sorted.find((e) => e.tipo === "uscita");
  if (!entrata) {
    return {
      entrataOra: null,
      uscitaOra: null,
      pausaMinuti: 0,
      oreLavorateMinuti: 0,
      oltreOrarioMinuti: 0,
      inPausa: false,
      chiusa: false,
    };
  }
  const startMs = new Date(entrata.ora).getTime();
  const endMs = uscita ? new Date(uscita.ora).getTime() : now.getTime();
  // Somma degli intervalli di pausa (chiusi con fine-pausa; se aperti, fino a
  // "now" o all'uscita).
  let pausaMs = 0;
  let inPausa = false;
  let inizioPausaMs: number | null = null;
  for (const e of sorted) {
    if (e.tipo === "inizio-pausa") {
      inizioPausaMs = new Date(e.ora).getTime();
    } else if (e.tipo === "fine-pausa" && inizioPausaMs != null) {
      pausaMs += Math.max(0, new Date(e.ora).getTime() - inizioPausaMs);
      inizioPausaMs = null;
    }
  }
  if (inizioPausaMs != null) {
    // pausa ancora aperta
    pausaMs += Math.max(0, endMs - inizioPausaMs);
    inPausa = true;
  }
  const pausaMinuti = Math.floor(pausaMs / 60000);
  const totMinuti = Math.max(0, Math.floor((endMs - startMs) / 60000));
  const oreLavorateMinuti = Math.max(0, totMinuti - pausaMinuti);
  const oltreOrarioMinuti = Math.max(0, oreLavorateMinuti - SOGLIA_ORE_MIN);
  return {
    entrataOra: entrata.ora,
    uscitaOra: uscita?.ora ?? null,
    pausaMinuti,
    oreLavorateMinuti,
    oltreOrarioMinuti,
    inPausa,
    chiusa: Boolean(uscita),
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