// DR Portal — Finanza → Flussi: server functions dello scadenziario fiscale.
// Stessa riservatezza degli stipendi: amministratore di sistema e vista
// direzione soltanto.

import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import { loadFiscaleDb, replaceFiscale } from "./fiscale.server";
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
