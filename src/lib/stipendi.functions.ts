// DR Portal — Finanza → Stipendi: server functions.
// Dati sensibili (costo del personale per dipendente): accesso riservato
// all'amministratore di sistema e alla vista direzione (DR005, DR007,
// ADM002) — un responsabile "semplice" non li vede.

import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import { deleteStipendiMese, loadStipendiDb, upsertStipendiMese } from "./stipendi.server";
import type { StipendiDb, StipendiMese } from "./stipendi-logic";

async function utenteStipendi(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  const ok = me.ruolo === "amministratore_sistema" || haVistaDirezione(me.codice ?? "");
  if (!ok) throw new Error("Sezione riservata alla direzione.");
  return me;
}

const firma = (me: ServerSessionUser) => me.codice || me.id;

export const spStipendiGet = createServerFn({ method: "GET" }).handler(
  async (): Promise<StipendiDb> => {
    await utenteStipendi();
    return loadStipendiDb();
  },
);

export const spStipendiSalvaMese = createServerFn({ method: "POST" })
  .inputValidator((input: { mese: StipendiMese }) => {
    const m = input?.mese;
    if (!m || !/^\d{4}-\d{2}$/.test(m.mese ?? "")) throw new Error("Mese non valido.");
    if (!Array.isArray(m.dipendenti) || m.dipendenti.length === 0)
      throw new Error("Nessun dipendente nel file.");
    return input;
  })
  .handler(async ({ data }): Promise<StipendiDb> => {
    const me = await utenteStipendi();
    return upsertStipendiMese(data.mese, firma(me));
  });

export const spStipendiEliminaMese = createServerFn({ method: "POST" })
  .inputValidator((input: { mese: string }) => {
    if (!/^\d{4}-\d{2}$/.test(input?.mese ?? "")) throw new Error("Mese non valido.");
    return input;
  })
  .handler(async ({ data }): Promise<StipendiDb> => {
    const me = await utenteStipendi();
    return deleteStipendiMese(data.mese, firma(me));
  });
