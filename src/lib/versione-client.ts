// DR Portal — sentinella d'aggiornamento (client).
//
// A ogni publish gli indirizzi interni delle chiamate server cambiano: una
// scheda rimasta aperta con la versione VECCHIA resta muta — i bottoni
// chiamano endpoint che non esistono più (caso "timbratrice" di Cerro,
// 16/09: uscite perse subito dopo la publish delle 15). La rotta /versione
// ha un indirizzo FISSO e dichiara la versione pubblicata: se è diversa
// dalla propria, la pagina si ricarica da sola — una volta sola per
// versione, mai in loop.

import { APP_INFO } from "./version";

export async function versioneViva(): Promise<string | null> {
  try {
    const r = await fetch(`/versione?ts=${Date.now()}`, { cache: "no-store" });
    if (!r.ok) return null;
    return (await r.text()).match(/data-versione="([^"]+)"/)?.[1] ?? null;
  } catch {
    return null; // offline o server irraggiungibile: si riprova al giro dopo
  }
}

/** true = versione nuova rilevata e ricarica avviata. `prima` viene eseguito
 *  subito PRIMA del reload (es. mettere in coda la timbratura appena persa). */
export async function ricaricaSeAggiornato(prima?: () => void): Promise<boolean> {
  const viva = await versioneViva();
  if (!viva || viva === APP_INFO.version) return false;
  const K = "dr:ricaricaPer";
  try {
    // Anti-loop: una sola ricarica per versione vista. Se dopo il reload la
    // versione risulta ancora diversa (deploy in propagazione), non si
    // ricarica di nuovo per lo stesso valore.
    if (window.sessionStorage.getItem(K) === viva) return false;
    window.sessionStorage.setItem(K, viva);
  } catch {
    /* senza sessionStorage si ricarica comunque */
  }
  prima?.();
  window.location.reload();
  return true;
}
