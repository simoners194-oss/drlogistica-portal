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

// Durata massima di un TURNO (ore). Il turno vive anche a cavallo di
// mezzanotte: l'uscita chiude l'entrata aperta anche se il giorno è cambiato.
// Oltre il tetto, il turno si considera dimenticato: l'Entrata torna
// disponibile e il turno aperto finisce nelle anomalie da sanare.
export const MAX_TURNO_ORE = 13;

/** Evento con orario, indipendente dal nome del campo client/server. */
export interface EventoConOra {
  evento: EventoTimbratura;
  ora: string; // ISO
}

const MAX_TURNO_MS = () => MAX_TURNO_ORE * 3600_000;

/** Ora (ISO) di apertura del turno ANCORA in corso, oppure null se il
 *  dipendente è fuori servizio o il turno aperto è scaduto (> MAX_TURNO_ORE).
 *  Il flusso può attraversare la mezzanotte: si passa una finestra di ~36h. */
export function aperturaTurnoCorrente(eventi: EventoConOra[], now = new Date()): string | null {
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const last = sorted[sorted.length - 1];
  if (!last || last.evento === "uscita") return null;
  // L'apertura è la prima entrata del blocco dopo l'ultima uscita.
  let apertura: string | null = null;
  for (const e of sorted) {
    if (e.evento === "uscita") apertura = null;
    else if (e.evento === "entrata" && apertura == null) apertura = e.ora;
  }
  const rif = apertura ?? last.ora; // dati sporchi senza entrata: ultimo evento
  if (now.getTime() - new Date(rif).getTime() > MAX_TURNO_MS()) return null;
  return rif;
}

/** Ultimo evento RILEVANTE per la macchina a stati: come l'ultimo del flusso,
 *  ma un turno aperto oltre il tetto conta come "nessuna timbratura" (si può
 *  ricominciare con una nuova Entrata; il turno scaduto va sanato a parte). */
export function ultimoEventoEffettivo(
  eventi: EventoConOra[],
  now = new Date(),
): EventoTimbratura | null {
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const last = sorted[sorted.length - 1];
  if (!last) return null;
  if (last.evento === "uscita") return "uscita";
  return aperturaTurnoCorrente(sorted, now) != null ? last.evento : null;
}

/** Eventi per la vista "oggi": quelli del giorno corrente PIÙ, se il turno in
 *  corso è iniziato prima di mezzanotte, quelli dall'apertura del turno.
 *  I giorni precedenti già chiusi restano fuori. */
export function tagliaEventiVisibili(eventi: EventoConOra[], now = new Date()): EventoConOra[] {
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const inizioOggi = new Date(now);
  inizioOggi.setHours(0, 0, 0, 0);
  const apertura = aperturaTurnoCorrente(sorted, now);
  const cutoffMs = Math.min(
    inizioOggi.getTime(),
    apertura ? new Date(apertura).getTime() : Number.POSITIVE_INFINITY,
  );
  return sorted.filter((e) => new Date(e.ora).getTime() >= cutoffMs);
}

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

// Sopra questa durata continuativa, un turno senza alcuno stacco viene
// segnalato come anomalia INFORMATIVA (nessun blocco): con i due tasti la
// pausa non è più dichiarata, si vede solo come uscita+rientro.
export const SENZA_STACCO_MIN_ORE = 6;

export interface AnomaliaGiorno {
  giorno: string; // YYYY-MM-DD di ATTRIBUZIONE (giorno di inizio del turno)
  tipo: TipoAnomalia;
}

// Analisi A TURNI del flusso eventi di UN dipendente (più giorni, in qualsiasi
// ordine). I turni vivono anche a cavallo di mezzanotte, quindi le anomalie
// NON si possono calcolare giorno per giorno:
// - turno non chiuso: un segmento in servizio resta aperto oltre MAX_TURNO_ORE
//   (attribuito al giorno di apertura). Un turno in corso da meno del tetto
//   NON è un'anomalia, anche se iniziato ieri sera.
// - senza stacco (informativa): segmento chiuso più lungo di
//   SENZA_STACCO_MIN_ORE senza alcuna uscita/rientro in mezzo.
// - pausa non chiusa: inizio-pausa mai chiuso (solo dati storici).
// `rilevaPausa=false` per chi non ha diritto alla pausa (es. part-time ≤16h).
export function anomalieDaStream(
  eventi: EventoConOra[],
  opts: { rilevaPausa: boolean; now?: Date },
): AnomaliaGiorno[] {
  const now = opts.now ?? new Date();
  const maxMs = MAX_TURNO_ORE * 3600_000;
  const staccoMs = SENZA_STACCO_MIN_ORE * 3600_000;
  const sorted = [...eventi].sort((a, b) => a.ora.localeCompare(b.ora));
  const t = (e: EventoConOra) => new Date(e.ora).getTime();
  const giornoDi = (e: EventoConOra) => e.ora.slice(0, 10);
  const out: AnomaliaGiorno[] = [];
  let apertura: EventoConOra | null = null; // inizio del segmento in servizio
  let pausaAperta: EventoConOra | null = null; // solo eventi legacy

  for (const e of sorted) {
    const ms = t(e);
    if (e.evento === "entrata" || e.evento === "fine-pausa") {
      if (e.evento === "fine-pausa") pausaAperta = null;
      if (e.evento === "entrata" && pausaAperta && ms - t(pausaAperta) > maxMs) {
        if (opts.rilevaPausa) out.push({ giorno: giornoDi(pausaAperta), tipo: "pausa-non-chiusa" });
        pausaAperta = null;
      }
      if (apertura == null) {
        apertura = e;
      } else if (ms - t(apertura) > maxMs) {
        // il vecchio segmento era stato dimenticato: si riparte da qui
        out.push({ giorno: giornoDi(apertura), tipo: "turno-non-chiuso" });
        apertura = e;
      }
    } else {
      // uscita | inizio-pausa: chiude il segmento in servizio
      if (e.evento === "uscita") pausaAperta = null;
      if (e.evento === "inizio-pausa") pausaAperta = e;
      if (apertura != null) {
        const durata = ms - t(apertura);
        if (durata > maxMs) out.push({ giorno: giornoDi(apertura), tipo: "turno-non-chiuso" });
        else if (opts.rilevaPausa && durata > staccoMs)
          out.push({ giorno: giornoDi(apertura), tipo: "senza-stacco" });
        apertura = null;
      }
    }
  }
  if (apertura != null && now.getTime() - t(apertura) > maxMs)
    out.push({ giorno: giornoDi(apertura), tipo: "turno-non-chiuso" });
  if (opts.rilevaPausa && pausaAperta != null && now.getTime() - t(pausaAperta) > maxMs)
    out.push({ giorno: giornoDi(pausaAperta), tipo: "pausa-non-chiusa" });

  // Dedup (stesso giorno + stesso tipo può ripetersi con dati sporchi).
  const seen = new Set<string>();
  return out.filter((a) => {
    const k = `${a.giorno}|${a.tipo}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}
