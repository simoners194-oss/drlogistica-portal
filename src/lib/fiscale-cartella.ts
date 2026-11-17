// DR Portal — lo scadenzario fiscale si rilegge da solo dalla cartella
// condivisa "Fiscale" (1.120.0, Simone 17/11/2026: "sì, fallo automatico").
// Lato browser: quando si apre una pagina che usa il fiscale (Flussi, Cassa
// proprietà, scheda Fiscale) si chiede al server di guardare la cartella.
// Guardare costa due letture leggere su Microsoft 365: per non ripeterle a
// ogni cambio di scheda, dallo stesso browser si guarda al più una volta ogni
// tre minuti. Il pulsante "Rileggi dalla cartella" guarda sempre.
import { spFiscaleLeggiCartella } from "./fiscale.functions";
import type { FiscaleDb, LetturaFiscale } from "./fiscale-logic";

const CHIAVE = "dr:fiscale:cartella";
const MINUTI_TRA_CONTROLLI = 3;

/** Restituisce l'archivio con l'esito in `ultimaLettura`, oppure null se il
 *  controllo è stato saltato perché appena fatto. */
export async function controllaFiscaleCartella(forza = false): Promise<FiscaleDb | null> {
  if (!forza) {
    try {
      const ultimo = Number(window.localStorage.getItem(CHIAVE) ?? 0);
      if (Date.now() - ultimo < MINUTI_TRA_CONTROLLI * 60000) return null;
    } catch {
      // senza memoria del browser si controlla e basta
    }
  }
  const db = await spFiscaleLeggiCartella({ data: { forza } });
  try {
    // Dopo un errore si riprova alla prossima apertura, senza aspettare.
    if (db.ultimaLettura?.esito !== "errore")
      window.localStorage.setItem(CHIAVE, String(Date.now()));
  } catch {
    // vedi sopra
  }
  return db;
}

/** "17/11 10:24" da una data ISO; "" se non è una data. */
export function quandoBreve(iso: string | undefined): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleString("it-IT", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

/** Da quale file viene lo scadenzario che si sta guardando, in una riga. */
export function fonteFiscaleBreve(
  db: Pick<FiscaleDb, "fonteFile" | "fonteModificato" | "aggiornatoIl"> | null | undefined,
  etichette: { salvato: string; letto: string },
): string {
  if (!db?.fonteFile) return "";
  const salvato = quandoBreve(db.fonteModificato);
  const letto = quandoBreve(db.aggiornatoIl);
  return [
    db.fonteFile,
    salvato ? `${etichette.salvato} ${salvato}` : "",
    letto ? `${etichette.letto} ${letto}` : "",
  ]
    .filter(Boolean)
    .join(" · ");
}

export type { LetturaFiscale };
