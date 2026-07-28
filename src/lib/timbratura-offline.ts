// DR Portal — Coda OFFLINE delle timbrature (client).
// Se al momento della pressione non c'è rete, l'evento viene salvato sul
// dispositivo con l'ORA REALE della pressione e re-inviato automaticamente
// appena la connessione torna (evento 'online' + tentativi periodici).
// Il server accetta l'orario del client solo in una finestra prudente
// (mai nel futuro, max 12 ore indietro) e la timbratura resta tracciata
// con una nota "recuperata offline".

import type { EventoTimbratura } from "./presenze-logic";

export interface TimbraturaInCoda {
  evento: EventoTimbratura;
  /** ISO dell'istante in cui il dipendente ha premuto il tasto. */
  dataOra: string;
}

const KEY = "dr:timbrature:coda";

export function leggiCoda(): TimbraturaInCoda[] {
  try {
    const raw = window.localStorage.getItem(KEY);
    const list = raw ? (JSON.parse(raw) as TimbraturaInCoda[]) : [];
    return Array.isArray(list) ? list.filter((t) => t?.evento && t?.dataOra) : [];
  } catch {
    return [];
  }
}

export function salvaCoda(list: TimbraturaInCoda[]): void {
  try {
    if (list.length === 0) window.localStorage.removeItem(KEY);
    else window.localStorage.setItem(KEY, JSON.stringify(list));
  } catch {
    /* localStorage non disponibile: niente coda */
  }
}

export function accoda(evento: EventoTimbratura, dataOra: string): TimbraturaInCoda[] {
  const coda = [...leggiCoda(), { evento, dataOra }];
  salvaCoda(coda);
  return coda;
}

/** Errore di RETE (nessuna risposta) vs rifiuto del server (risposta vera). */
export function isErroreRete(err: unknown): boolean {
  if (typeof navigator !== "undefined" && navigator.onLine === false) return true;
  const m = (err instanceof Error ? err.message : String(err)).toLowerCase();
  return /failed to fetch|networkerror|network error|load failed|fetch failed|err_internet|timeout/i.test(
    m,
  );
}

export interface EsitoSvuotamento {
  inviate: number;
  /** Rifiutate dal server (es. sequenza non valida): NON si riprovano. */
  scartate: number;
  rimaste: number;
}

// Invia la coda in ordine. Al primo errore di RETE si ferma (si riproverà);
// un rifiuto del server scarta il singolo elemento e prosegue.
export async function svuotaCoda(
  invia: (t: TimbraturaInCoda) => Promise<void>,
): Promise<EsitoSvuotamento> {
  const coda = leggiCoda();
  const esito: EsitoSvuotamento = { inviate: 0, scartate: 0, rimaste: 0 };
  const rimaste: TimbraturaInCoda[] = [];
  for (let i = 0; i < coda.length; i++) {
    try {
      await invia(coda[i]);
      esito.inviate++;
    } catch (err) {
      if (isErroreRete(err)) {
        rimaste.push(...coda.slice(i));
        break;
      }
      esito.scartate++;
    }
  }
  salvaCoda(rimaste);
  esito.rimaste = rimaste.length;
  return esito;
}
