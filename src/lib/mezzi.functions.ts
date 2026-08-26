// DR Portal — modulo Mezzi: server functions (unico entry point per il client).
// Le implementazioni di storage vivono in mezzi.server.ts (escluso dal bundle
// client). Autorizzazioni: lettura e scrittura riservate a responsabili,
// amministratore di sistema e vista direzione (Diego/il proprietario);
// i dipendenti non vedono il modulo.

import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import {
  cronScadenzeMezzi,
  importMezziDb,
  loadMezziDb,
  mutateMezziDb,
  type CronMezziResult,
} from "./mezzi.server";
import { verificaTokenCronFatture } from "./sharepoint.server";
import {
  affidamentoCorrente,
  autistaAllaData,
  normalizzaTarga,
  nuovoId,
  type Affidamento,
  type Contratto,
  type InterventoOfficina,
  type LetturaKm,
  type MezziDb,
  type Mezzo,
  type Multa,
  type ParametriMezzi,
  type PermessoZtl,
  type Rifornimento,
  type Scadenza,
} from "./mezzi-types";

async function utenteMezzi(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  const abilitato =
    me.ruolo === "amministratore_sistema" ||
    me.ruolo === "responsabile" ||
    haVistaDirezione(me.codice ?? "");
  if (!abilitato) throw new Error("Non sei autorizzato al modulo Mezzi.");
  return me;
}

function firma(me: ServerSessionUser): string {
  return me.codice || `${me.nome} ${me.cognome}`.trim() || me.id;
}

function upsert<T extends { id: string }>(arr: T[], item: T): void {
  const i = arr.findIndex((x) => x.id === item.id);
  if (i >= 0) arr[i] = item;
  else arr.push(item);
}

// ---------------------------------------------------------------------------

export const spMezziGetDb = createServerFn({ method: "GET" }).handler(
  async (): Promise<MezziDb> => {
    await utenteMezzi();
    return loadMezziDb();
  },
);

export const spMezziSalvaMezzo = createServerFn({ method: "POST" })
  .inputValidator((input: { mezzo: Mezzo }) => {
    if (!input?.mezzo?.targa?.trim()) throw new Error("Targa mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const mezzo = { ...data.mezzo };
    mezzo.targa = mezzo.targa.trim().toUpperCase();
    if (!mezzo.id) mezzo.id = normalizzaTarga(mezzo.targa);
    return mutateMezziDb(firma(me), (db) => upsert(db.mezzi, mezzo));
  });

export const spMezziSalvaAffidamento = createServerFn({ method: "POST" })
  .inputValidator((input: { affidamento: Affidamento }) => {
    const a = input?.affidamento;
    if (!a?.mezzoId) throw new Error("Mezzo mancante.");
    if (!a?.autistaNome?.trim()) throw new Error("Autista mancante.");
    if (!a?.dal) throw new Error("Data di inizio mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const a = { ...data.affidamento };
    if (!a.id) a.id = nuovoId("aff");
    return mutateMezziDb(firma(me), (db) => {
      // Nuovo affidamento aperto → chiude l'eventuale precedente in corso.
      if (!a.al) {
        const corrente = affidamentoCorrente(db.affidamenti, a.mezzoId);
        if (corrente && corrente.id !== a.id) {
          corrente.al = a.dal;
        }
      }
      upsert(db.affidamenti, a);
      // Denormalizza appalto/autista correnti sul mezzo.
      const m = db.mezzi.find((x) => x.id === a.mezzoId);
      if (m && !a.al) {
        if (a.appalto) m.appalto = a.appalto;
      }
    });
  });

export const spMezziSalvaScadenza = createServerFn({ method: "POST" })
  .inputValidator((input: { scadenza: Scadenza }) => {
    const s = input?.scadenza;
    if (!s?.tipo) throw new Error("Tipo scadenza mancante.");
    if (!s?.scadenza) throw new Error("Data scadenza mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const s = { ...data.scadenza };
    if (!s.id) s.id = nuovoId("scad");
    return mutateMezziDb(firma(me), (db) => {
      const prima = db.scadenze.find((x) => x.id === s.id);
      // Rinnovo (data cambiata) → riparte il ciclo di alert.
      if (prima && prima.scadenza !== s.scadenza) s.alertInviati = {};
      upsert(db.scadenze, s);
    });
  });

export const spMezziSalvaContratto = createServerFn({ method: "POST" })
  .inputValidator((input: { contratto: Contratto }) => {
    if (!input?.contratto?.mezzoId) throw new Error("Mezzo mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const c = { ...data.contratto };
    if (!c.id) c.id = nuovoId("ctr");
    if (c.attivo === undefined) c.attivo = true;
    return mutateMezziDb(firma(me), (db) => upsert(db.contratti, c));
  });

export const spMezziSalvaMulta = createServerFn({ method: "POST" })
  .inputValidator((input: { multa: Multa }) => {
    const m = input?.multa;
    if (!m?.mezzoId) throw new Error("Mezzo mancante.");
    if (!m?.dataInfrazione) throw new Error("Data/ora infrazione mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const m = { ...data.multa };
    if (!m.id) m.id = nuovoId("mul");
    return mutateMezziDb(firma(me), (db) => {
      // Match automatico dell'autista dallo storico affidamenti se non indicato.
      if (!m.autistaNome) {
        const aff = autistaAllaData(db.affidamenti, m.mezzoId, m.dataInfrazione);
        if (aff) {
          m.autistaNome = aff.autistaNome;
          m.autistaCodice = aff.autistaCodice;
        }
      }
      const prima = db.multe.find((x) => x.id === m.id);
      if (!prima || prima.stato !== m.stato) {
        m.storico = [
          ...(prima?.storico ?? m.storico ?? []),
          { data: new Date().toISOString(), stato: m.stato, utente: firma(me) },
        ];
      }
      upsert(db.multe, m);
    });
  });

export const spMezziSalvaZtl = createServerFn({ method: "POST" })
  .inputValidator((input: { permesso: PermessoZtl }) => {
    if (!input?.permesso?.comune?.trim()) throw new Error("Comune mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const p = { ...data.permesso };
    if (!p.id) p.id = nuovoId("ztl");
    return mutateMezziDb(firma(me), (db) => upsert(db.ztl, p));
  });

export const spMezziSalvaIntervento = createServerFn({ method: "POST" })
  .inputValidator((input: { intervento: InterventoOfficina }) => {
    const i = input?.intervento;
    if (!i?.mezzoId) throw new Error("Mezzo mancante.");
    if (!i?.data) throw new Error("Data mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const i = { ...data.intervento };
    if (!i.id) i.id = nuovoId("off");
    return mutateMezziDb(firma(me), (db) => upsert(db.officina, i));
  });

export const spMezziSalvaLetturaKm = createServerFn({ method: "POST" })
  .inputValidator((input: { lettura: LetturaKm }) => {
    const l = input?.lettura;
    if (!l?.mezzoId) throw new Error("Mezzo mancante.");
    if (!l?.data) throw new Error("Data mancante.");
    if (!(Number(l.km) > 0)) throw new Error("Km non validi.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const l = { ...data.lettura, km: Number(data.lettura.km) };
    if (!l.id) l.id = nuovoId("km");
    return mutateMezziDb(firma(me), (db) => {
      upsert(db.km, l);
      const m = db.mezzi.find((x) => x.id === l.mezzoId);
      if (m && (!m.kmAggiornatiAl || m.kmAggiornatiAl <= l.data)) {
        m.kmAttuali = l.km;
        m.kmAggiornatiAl = l.data;
      }
    });
  });

export const spMezziSalvaRifornimento = createServerFn({ method: "POST" })
  .inputValidator((input: { rifornimento: Rifornimento }) => {
    const r = input?.rifornimento;
    if (!r?.mezzoId) throw new Error("Mezzo mancante.");
    if (!r?.data) throw new Error("Data mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    const r = { ...data.rifornimento };
    if (!r.id) r.id = nuovoId("rif");
    return mutateMezziDb(firma(me), (db) => upsert(db.carburante, r));
  });

export const spMezziSalvaParametri = createServerFn({ method: "POST" })
  .inputValidator((input: { parametri: ParametriMezzi }) => {
    if (!input?.parametri) throw new Error("Parametri mancanti.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    return mutateMezziDb(firma(me), (db) => {
      db.parametri = data.parametri;
    });
  });

/** Cancellazione generica di un record da una collezione del modulo. */
export type MezziCollezione =
  | "mezzi"
  | "affidamenti"
  | "scadenze"
  | "contratti"
  | "multe"
  | "ztl"
  | "officina"
  | "km"
  | "carburante";

export const spMezziElimina = createServerFn({ method: "POST" })
  .inputValidator((input: { collezione: MezziCollezione; id: string }) => {
    if (!input?.collezione || !input?.id) throw new Error("Parametri mancanti.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    return mutateMezziDb(firma(me), (db) => {
      const arr = db[data.collezione] as { id: string }[];
      const i = arr.findIndex((x) => x.id === data.id);
      if (i >= 0) arr.splice(i, 1);
      // Eliminare un mezzo rimuove anche i suoi record collegati.
      if (data.collezione === "mezzi") {
        db.affidamenti = db.affidamenti.filter((x) => x.mezzoId !== data.id);
        db.scadenze = db.scadenze.filter((x) => x.mezzoId !== data.id);
        db.contratti = db.contratti.filter((x) => x.mezzoId !== data.id);
        db.multe = db.multe.filter((x) => x.mezzoId !== data.id);
        db.ztl = db.ztl.filter((x) => x.mezzoId !== data.id);
        db.officina = db.officina.filter((x) => x.mezzoId !== data.id);
        db.km = db.km.filter((x) => x.mezzoId !== data.id);
        db.carburante = db.carburante.filter((x) => x.mezzoId !== data.id);
      }
    });
  });

export const spMezziImportaDb = createServerFn({ method: "POST" })
  .inputValidator((input: { json: string }) => {
    if (!input?.json?.trim()) throw new Error("JSON mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<MezziDb> => {
    const me = await utenteMezzi();
    if (me.ruolo !== "amministratore_sistema") {
      throw new Error("Solo l'amministratore di sistema può importare il database.");
    }
    return importMezziDb(data.json, firma(me));
  });

/** Cron scadenze (chiamato da Power Automate con lo stesso token dei cron esistenti). */
export const spMezziCron = createServerFn({ method: "POST" })
  .inputValidator((input: { token: string }) => {
    if (!input?.token) throw new Error("Token mancante.");
    return input;
  })
  .handler(async ({ data }): Promise<CronMezziResult> => {
    await verificaTokenCronFatture(data.token);
    return cronScadenzeMezzi();
  });
