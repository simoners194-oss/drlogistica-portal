// DR Portal — Ritorno dall'autorizzazione bancaria (Enable Banking).
// La pagina di login mette da parte il ?code= (o l'errore) della redirect e
// il pannello Banca lo consuma: localStorage, così vale per TUTTE le finestre
// del profilo (app installata compresa). Il codice è monouso e scade in pochi
// minuti: oltre RITORNO_MAX_ETA_MS si scarta e si segnala.

export const EB_RITORNO_KEY = "dr:eb:ritorno";
export const RITORNO_MAX_ETA_MS = 15 * 60 * 1000;

export interface EbRitorno {
  code?: string;
  errore?: string;
  ts: number;
}

/** Legge e CANCELLA il ritorno parcheggiato (monouso). Compatibile con la
 *  vecchia chiave di sessionStorage (prima della 1.96.2). */
export function prelevaRitornoBanca(): EbRitorno | null {
  let r: EbRitorno | null = null;
  try {
    const raw = window.localStorage.getItem(EB_RITORNO_KEY);
    if (raw) {
      const p = JSON.parse(raw) as Partial<EbRitorno>;
      r = { code: p.code || undefined, errore: p.errore || undefined, ts: Number(p.ts) || 0 };
    }
    window.localStorage.removeItem(EB_RITORNO_KEY);
  } catch {
    /* localStorage non disponibile */
  }
  if (!r) {
    try {
      const code = window.sessionStorage.getItem("dr:eb:code");
      if (code) r = { code, ts: Date.now() };
      window.sessionStorage.removeItem("dr:eb:code");
    } catch {
      /* sessionStorage non disponibile */
    }
  }
  return r;
}
