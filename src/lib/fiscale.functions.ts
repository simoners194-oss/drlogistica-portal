// DR Portal — Finanza → Flussi: server functions dello scadenziario fiscale.
// Stessa riservatezza degli stipendi: amministratore di sistema e vista
// direzione soltanto.

import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import {
  eliminaScadenzaFiscale,
  loadFiscaleDb,
  replaceFiscale,
  setDaRateizzareFiscale,
  upsertScadenzaFiscale,
} from "./fiscale.server";
import type { DaRateizzareFiscale, FiscaleDb, ScadenzaFiscale } from "./fiscale-logic";

async function utenteFiscale(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  const ok = me.ruolo === "amministratore_sistema" || haVistaDirezione(me.codice ?? "");
  if (!ok) throw new Error("Sezione riservata alla direzione.");
  return me;
}

export const spFiscaleGet = createServerFn({ method: "GET" }).handler(
  async (): Promise<FiscaleDb> => {
    await utenteFiscale();
    return loadFiscaleDb();
  },
);

export const spFiscaleSalva = createServerFn({ method: "POST" })
  .inputValidator(
    (input: {
      scadenze: ScadenzaFiscale[];
      daRateizzare: DaRateizzareFiscale[];
      fonte: string;
    }) => {
      if (!Array.isArray(input?.scadenze) || input.scadenze.length === 0)
        throw new Error("Nessuna scadenza nel file.");
      for (const s of input.scadenze) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(s.dataPagamento ?? ""))
          throw new Error("Data pagamento non valida nel file.");
        if (!Number.isFinite(s.importo)) throw new Error("Importo non valido nel file.");
      }
      return input;
    },
  )
  .handler(async ({ data }): Promise<FiscaleDb> => {
    const me = await utenteFiscale();
    return replaceFiscale(
      data.scadenze,
      Array.isArray(data.daRateizzare) ? data.daRateizzare : [],
      data.fonte,
      me.codice || me.id,
    );
  });

/** Crea/modifica una scadenza direttamente dal portale (tab Fiscale). */
export const spFiscaleUpsertScadenza = createServerFn({ method: "POST" })
  .inputValidator((input: { scadenza: ScadenzaFiscale }) => {
    const s = input?.scadenza;
    if (!s?.voce?.trim()) throw new Error("Voce mancante.");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s.dataPagamento ?? ""))
      throw new Error("Data pagamento non valida.");
    if (!Number.isFinite(s.importo) || s.importo <= 0)
      throw new Error("Importo non valido (serve un numero maggiore di zero).");
    if (s.anno != null && (!Number.isInteger(s.anno) || s.anno < 2000 || s.anno > 2100))
      throw new Error("Anno di competenza non valido.");
    return input;
  })
  .handler(async ({ data }): Promise<FiscaleDb> => {
    const me = await utenteFiscale();
    const s = data.scadenza;
    s.voce = s.voce.trim().toUpperCase();
    return upsertScadenzaFiscale(s, me.codice || me.id);
  });

/** Elenco "da registrare in futuro" (senza data, fuori dal cash flow). */
export const spFiscaleDaRateizzare = createServerFn({ method: "POST" })
  .inputValidator((input: { lista: DaRateizzareFiscale[] }) => {
    if (!Array.isArray(input?.lista)) throw new Error("Elenco non valido.");
    for (const d of input.lista) {
      if (!d?.voce?.trim()) throw new Error("Voce mancante in una riga.");
      if (!Number.isFinite(d.importo) || d.importo <= 0)
        throw new Error("Importo non valido in una riga.");
    }
    return input;
  })
  .handler(async ({ data }): Promise<FiscaleDb> => {
    const me = await utenteFiscale();
    return setDaRateizzareFiscale(data.lista, me.codice || me.id);
  });

export const spFiscaleEliminaScadenza = createServerFn({ method: "POST" })
  .inputValidator((input: { id: string }) => {
    if (!input?.id?.trim()) throw new Error("Id mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<FiscaleDb> => {
    const me = await utenteFiscale();
    return eliminaScadenzaFiscale(data.id.trim(), me.codice || me.id);
  });
