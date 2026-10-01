// DR Portal — sentinella d'aggiornamento (client).
//
// A ogni publish gli indirizzi interni delle chiamate server cambiano: una
// scheda rimasta aperta con la versione VECCHIA resta muta — i bottoni
// chiamano endpoint che non esistono più (caso "timbratrice" di Cerro,
// 16/09: uscite perse subito dopo la publish delle 15). La rotta /versione
// ha un indirizzo FISSO e dichiara la versione pubblicata: se è diversa
// dalla propria, la pagina si ricarica da sola — una volta sola per
// versione, mai in loop.
//
// 01/10/2026 (Zingali "non riescono a timbrare", sei publish in un giorno):
// l'hosting risponde a un indirizzo server vecchio con 200 e un file
// JavaScript, NON con un errore. Il client di TanStack restituisce allora la
// Response grezza come "risultato": la timbratura sembrava registrata
// (toast verde), il server non la riceveva mai e la pagina non si ricaricava
// mai, perché nessuno lanciava un errore. Da qui: `rispostaStantia` smaschera
// quel risultato (middleware globale in start.ts + controlli nei punti
// critici) e `avviaSentinella` controlla la versione da sola ogni minuto e
// quando la pagina torna in primo piano, PRIMA che qualcuno prema un tasto.

import { APP_INFO } from "./version";

/** Messaggio dell'errore "pagina vecchia": NON deve somigliare a un errore di
 *  rete (isErroreRete), così la timbratura va in coda e la pagina si ricarica. */
export const MSG_VERSIONE_VECCHIA =
  "Portale aggiornato: la pagina si ricarica da sola, la timbratura resta in coda.";

/** true quando una chiamata server ha restituito la Response grezza (o
 *  nulla) invece del risultato: succede SOLO su una scheda vecchia dopo una
 *  publish, perché l'hosting serve un file JavaScript (200) all'indirizzo
 *  non più esistente. Un risultato vero è sempre un valore serializzato. */
export function rispostaStantia(x: unknown): boolean {
  return typeof Response !== "undefined" && x instanceof Response;
}

export async function versioneViva(): Promise<string | null> {
  try {
    const r = await fetch(`/versione?ts=${Date.now()}`, { cache: "no-store" });
    if (!r.ok) return null;
    return (await r.text()).match(/data-versione="([^"]+)"/)?.[1] ?? null;
  } catch {
    return null; // offline o server irraggiungibile: si riprova al giro dopo
  }
}

const K_RICARICA = "dr:ricaricaPer";
const K_FORZATA = "dr:ricaricaForzataPer";

/** true = versione nuova rilevata e ricarica avviata. `prima` viene eseguito
 *  subito PRIMA del reload (es. mettere in coda la timbratura appena persa).
 *  `forza`: la pagina è SICURAMENTE vecchia (risposta stantia): se la
 *  ricarica normale per questa versione è già stata fatta e non è bastata
 *  (HTML servito da una cache), se ne fa una seconda con l'indirizzo
 *  "sporcato" per saltare la cache — una sola, poi ci si arrende. */
// Un solo controllo alla volta: il middleware globale e il tasto che ha
// fallito chiamano entrambi la sentinella nello stesso istante; il secondo
// si accoda al primo invece di fare una seconda ricarica (quella "sporcata").
let controlloInCorso: Promise<boolean> | null = null;

export function ricaricaSeAggiornato(prima?: () => void, forza = false): Promise<boolean> {
  if (controlloInCorso) {
    prima?.();
    return controlloInCorso;
  }
  controlloInCorso = ricaricaSeAggiornatoOra(prima, forza).finally(() => {
    controlloInCorso = null;
  });
  return controlloInCorso;
}

async function ricaricaSeAggiornatoOra(
  prima?: () => void,
  forza: boolean = false,
): Promise<boolean> {
  const viva = await versioneViva();
  if (!viva) return false;
  if (viva === APP_INFO.version && !forza) return false;
  const chiave = viva === APP_INFO.version ? `${viva}!` : viva;
  let giaFatta = false;
  let giaForzata = false;
  try {
    giaFatta = window.sessionStorage.getItem(K_RICARICA) === chiave;
    giaForzata = window.sessionStorage.getItem(K_FORZATA) === chiave;
  } catch {
    /* senza sessionStorage si ricarica comunque, una volta */
  }
  if (giaFatta && (!forza || giaForzata)) return false;
  try {
    window.sessionStorage.setItem(giaFatta ? K_FORZATA : K_RICARICA, chiave);
  } catch {
    /* ignorato */
  }
  prima?.();
  if (giaFatta) {
    // Seconda ricarica: indirizzo sporcato, la cache dell'HTML non vale.
    const url = new URL(window.location.href);
    url.searchParams.set("r", String(Date.now()));
    window.location.replace(url.toString());
  } else {
    window.location.reload();
  }
  return true;
}

/** Controllo periodico della versione (ogni `intervalloMs`, e appena la
 *  pagina torna visibile): se è uscita una versione nuova la pagina si
 *  ricarica da sola quando `puoRicaricare()` lo consente (niente tocco in
 *  corso, niente coda in invio). Restituisce la funzione per fermarla. */
export function avviaSentinella(opts: {
  intervalloMs?: number;
  puoRicaricare?: () => boolean;
}): () => void {
  if (typeof window === "undefined") return () => {};
  let inCorso = false;
  const controlla = async () => {
    if (inCorso || document.visibilityState === "hidden") return;
    if (opts.puoRicaricare && !opts.puoRicaricare()) return;
    inCorso = true;
    try {
      await ricaricaSeAggiornato();
    } finally {
      inCorso = false;
    }
  };
  const iv = window.setInterval(controlla, opts.intervalloMs ?? 60_000);
  const onVisibile = () => {
    if (document.visibilityState === "visible") void controlla();
  };
  document.addEventListener("visibilitychange", onVisibile);
  window.addEventListener("focus", onVisibile);
  return () => {
    window.clearInterval(iv);
    document.removeEventListener("visibilitychange", onVisibile);
    window.removeEventListener("focus", onVisibile);
  };
}
