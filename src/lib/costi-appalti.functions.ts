// DR Portal — Finanze → "Costi appalti": server functions.
// Ricavi, costi e margini per cliente: stessa riservatezza della Cassa
// proprietà (amministratore di sistema e vista direzione soltanto).
import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import {
  eliminaSettimanaCosti,
  leggiCartellaCosti,
  loadCostiAppaltiDb,
  salvaSettimanaCosti,
} from "./costi-appalti.server";
import { validaSettimana, type CostiAppaltiDb, type SettimanaCosti } from "./costi-appalti-logic";

async function utenteCosti(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  const ok = me.ruolo === "amministratore_sistema" || haVistaDirezione(me.codice ?? "");
  if (!ok) throw new Error("Sezione riservata alla direzione.");
  return me;
}

export const spCostiAppaltiGet = createServerFn({ method: "GET" }).handler(
  async (): Promise<CostiAppaltiDb> => {
    await utenteCosti();
    return loadCostiAppaltiDb();
  },
);

/** Carica una settimana letta dal file (sostituisce quella con la stessa
 *  data d'inizio, se c'è già). */
export const spCostiAppaltiSalva = createServerFn({ method: "POST" })
  .inputValidator((input: { settimana: SettimanaCosti }) => ({
    settimana: validaSettimana(input?.settimana),
  }))
  .handler(async ({ data }): Promise<CostiAppaltiDb> => {
    const me = await utenteCosti();
    return salvaSettimanaCosti(data.settimana, `${me.nome} ${me.cognome}`.trim() || me.id);
  });

export const spCostiAppaltiElimina = createServerFn({ method: "POST" })
  .inputValidator((input: { id: string }) => {
    const id = String(input?.id ?? "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(id)) throw new Error("Settimana non valida.");
    return { id };
  })
  .handler(async ({ data }): Promise<CostiAppaltiDb> => {
    const me = await utenteCosti();
    return eliminaSettimanaCosti(data.id, `${me.nome} ${me.cognome}`.trim() || me.id);
  });

/** Guarda la cartella condivisa dei file settimanali e legge quelli nuovi o
 *  cambiati. */
export const spCostiAppaltiLeggiCartella = createServerFn({ method: "POST" }).handler(
  async (): Promise<CostiAppaltiDb> => {
    const me = await utenteCosti();
    return leggiCartellaCosti(`${me.nome} ${me.cognome}`.trim() || me.id);
  },
);
