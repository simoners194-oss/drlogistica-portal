// DR Portal — Server functions per SharePoint (via Lovable Connector Gateway).
// Sono l'unico entry point che il codice client usa per parlare con
// Microsoft Graph. Le implementazioni vere vivono in `sharepoint.server.ts`
// (bloccato dal bundle client dal suffisso .server.ts).

import { createServerFn } from "@tanstack/react-start";
import { normalizeRuolo } from "./session";
import {
  setSessionCookie,
  readSessionUser,
  clearSessionCookie,
  sessionSecretConfigured,
  type ServerSessionUser,
} from "./auth.server";
import {
  cancelRichiesta,
  clearSpDiscoveryCache,
  computeAnomalie,
  computeHealth,
  computeRendiconto,
  createRichiesta,
  createTimbratura,
  createTimbraturaManuale,
  createTurnoManuale,
  decideRichiesta,
  discoverSharePoint,
  fetchDipendenti,
  fetchRichieste,
  fetchRichiestePerSupervisore,
  fetchTimbratureManuali,
  importDipendenti,
  type ImportDipendentiResult,
  fetchTimbratureOggi,
  getLastSyncAt,
  getSpLog,
  loginByCodicePin,
  markSync,
  runSelfTest,
  uploadGiustificativo,
  type UploadGiustificativoResult,
  type CreateRichiestaInput,
  type CreateTimbraturaInput,
  type CreateTimbraturaManualeInput,
  type CreateTurnoManualeInput,
  type AnomaliaItem,
  type TimbraturaManualeItem,
  type RendicontoRiga,
  type DecideRichiestaInput,
  type EventoTimbratura,
  type LoginResult,
  type RichiesteFilter,
  type SpHealth,
  type SpLogEvent,
  type SpRichiesta,
  type SpSelfTestResult,
  type SpDipendente,
  type SpDiscovered,
  type SpTimbratura,
} from "./sharepoint.server";

// --- S1b: identità e autorizzazione dalla SESSIONE SERVER (cookie firmato) ---
// L'attore di ogni operazione è preso dal cookie, MAI dal payload del client.
async function currentUser(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  return me;
}
function assertCap(ok: boolean): void {
  if (!ok) throw new Error("Non sei autorizzato per questa operazione.");
}
const isAdmin = (me: ServerSessionUser) => me.ruolo === "amministratore_sistema";

export interface SpDiagnostics {
  hasLovableKey: boolean;
  hasConnectionKey: boolean;
  discovered: SpDiscovered | null;
  error: string | null;
  health: SpHealth | null;
  log: SpLogEvent[];
  lastSyncAt: string | null;
}

export const spGetDiagnostics = createServerFn({ method: "GET" })
  .inputValidator((input?: { force?: boolean }) => ({ force: Boolean(input?.force) }))
  .handler(async ({ data }): Promise<SpDiagnostics> => {
    assertCap(isAdmin(await currentUser()));
    const hasLovableKey = Boolean(process.env.LOVABLE_API_KEY);
    const hasConnectionKey = Boolean(process.env.MICROSOFT_SHAREPOINT_API_KEY);
    if (!hasLovableKey || !hasConnectionKey) {
      return {
        hasLovableKey,
        hasConnectionKey,
        discovered: null,
        error:
          "Credenziali del connettore SharePoint mancanti sul server (LOVABLE_API_KEY / MICROSOFT_SHAREPOINT_API_KEY).",
        health: null,
        log: getSpLog(),
        lastSyncAt: getLastSyncAt(),
      };
    }
    try {
      if (data.force) clearSpDiscoveryCache();
      const discovered = await discoverSharePoint(data.force);
      const health = await computeHealth();
      return {
        hasLovableKey,
        hasConnectionKey,
        discovered,
        error: null,
        health,
        log: getSpLog(),
        lastSyncAt: getLastSyncAt(),
      };
    } catch (err) {
      const health = await computeHealth().catch(() => null);
      return {
        hasLovableKey,
        hasConnectionKey,
        discovered: null,
        error: err instanceof Error ? err.message : String(err),
        health,
        log: getSpLog(),
        lastSyncAt: getLastSyncAt(),
      };
    }
  });

export interface SpSnapshot {
  dipendenti: SpDipendente[];
  timbrature: SpTimbratura[];
}

export const spGetSnapshot = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpSnapshot> => {
    await currentUser(); // richiede una sessione valida
    // Garantisce discovery prima delle chiamate in parallelo (evita doppia
    // esecuzione della discovery quando la cache è fredda).
    await discoverSharePoint();
    const [dipendenti, timbrature] = await Promise.all([fetchDipendenti(), fetchTimbratureOggi()]);
    markSync();
    return { dipendenti, timbrature };
  },
);

export const spCreateTimbratura = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTimbraturaInput): CreateTimbraturaInput => {
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    const validi: EventoTimbratura[] = ["entrata", "inizio-pausa", "fine-pausa", "uscita"];
    if (!validi.includes(input.evento)) throw new Error("evento non valido");
    return input;
  })
  .handler(async ({ data }) => {
    const me = await currentUser();
    // Timbra SEMPRE per sé stesso: id dalla sessione, non dal client.
    return createTimbratura({ ...data, dipendenteId: me.id });
  });

export const spRunSelfTest = createServerFn({ method: "POST" }).handler(
  async (): Promise<SpSelfTestResult> => {
    assertCap(isAdmin(await currentUser()));
    return runSelfTest();
  },
);

export const spLogin = createServerFn({ method: "POST" })
  .inputValidator((input: { codice: string; pin: string }) => {
    if (typeof input?.codice !== "string" || typeof input?.pin !== "string") {
      throw new Error("Codice o PIN non validi.");
    }
    return { codice: input.codice, pin: input.pin };
  })
  .handler(async ({ data }): Promise<LoginResult> => {
    const res = await loginByCodicePin(data.codice, data.pin);
    if (res.ok && res.dipendente) {
      const d = res.dipendente;
      // S1: emette la sessione server firmata (no-op se manca SESSION_SECRET).
      await setSessionCookie({
        id: d.id,
        nome: d.nome,
        cognome: d.cognome,
        sede: d.sede,
        ruolo: normalizeRuolo(d.ruolo),
        autorizza: d.autorizza,
        operatore: d.operatore,
      });
    }
    return res;
  });

// Identità dalla SESSIONE SERVER (cookie firmato). `sessionePronta` indica se
// il segreto di firma è configurato. Serve alla verifica della S1 e, in
// seguito, all'enforcement lato server.
export interface WhoAmI {
  user: ServerSessionUser | null;
  sessionePronta: boolean;
}

export const spWhoAmI = createServerFn({ method: "GET" }).handler(async (): Promise<WhoAmI> => ({
  user: await readSessionUser(),
  sessionePronta: sessionSecretConfigured(),
}));

export const spLogout = createServerFn({ method: "POST" }).handler(
  async (): Promise<{ ok: true }> => {
    clearSessionCookie();
    return { ok: true };
  },
);

// ---------------------------------------------------------------------------
// Richieste (Sprint 2)
// ---------------------------------------------------------------------------

export const spGetRichieste = createServerFn({ method: "GET" })
  .inputValidator((input?: RichiesteFilter): RichiesteFilter => ({
    richiedenteId: input?.richiedenteId ? String(input.richiedenteId) : undefined,
    stato: input?.stato ? String(input.stato) : undefined,
  }))
  .handler(async ({ data }): Promise<SpRichiesta[]> => {
    const me = await currentUser();
    if (data.richiedenteId) {
      // Vista personale: forzata al proprio id (non si leggono richieste altrui).
      return fetchRichieste({ richiedenteId: me.id, stato: data.stato });
    }
    // Coda approvatore (richieste DA DECIDERE, stato "Inviata"): solo
    // autorizzatori; scope per sede di competenza (DR005 globale, altri solo le
    // proprie sedi).
    if (data.stato === "Inviata") {
      assertCap(me.autorizza || isAdmin(me));
      if (isAdmin(me)) return fetchRichieste({ stato: data.stato });
      return fetchRichiestePerSupervisore(me.id, data.stato);
    }
    // Report richieste già decise (Approvata/Respinta) e altre viste
    // privilegiate: visibili a autorizzatori, OPERATORE e admin, SENZA scope per
    // sede — l'operatore DR000 deve vedere TUTTE le approvate.
    assertCap(me.autorizza || me.operatore || isAdmin(me));
    return fetchRichieste({ stato: data.stato });
  });

export const spCreateRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: CreateRichiestaInput): CreateRichiestaInput => {
    if (!input?.richiedenteId) throw new Error("richiedenteId mancante");
    if (!input?.tipo) throw new Error("tipo mancante");
    return input;
  })
  .handler(async ({ data }): Promise<SpRichiesta> => {
    const me = await currentUser();
    return createRichiesta({ ...data, richiedenteId: me.id });
  });

// Upload del giustificativo di spesa (rimborsi). Richiede solo una sessione
// valida: ogni dipendente può caricare il proprio documento. Ritorna il webUrl
// da salvare nel campo "Giustificativo" della richiesta.
export const spUploadGiustificativo = createServerFn({ method: "POST" })
  .inputValidator((input: { filename: string; contentBase64: string }) => {
    if (!input?.contentBase64) throw new Error("Contenuto file mancante");
    return {
      filename: String(input.filename ?? "documento"),
      contentBase64: String(input.contentBase64),
    };
  })
  .handler(async ({ data }): Promise<UploadGiustificativoResult> => {
    await currentUser();
    return uploadGiustificativo(data.filename, data.contentBase64);
  });

export const spDecideRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: DecideRichiestaInput): DecideRichiestaInput => {
    if (!input?.richiestaId) throw new Error("richiestaId mancante");
    if (!input?.approvatoreId) throw new Error("approvatoreId mancante");
    if (input.decisione !== "Approvata" && input.decisione !== "Respinta")
      throw new Error("decisione non valida");
    return input;
  })
  .handler(async ({ data }): Promise<SpRichiesta> => {
    const me = await currentUser();
    // L'approvatore è chi ha la sessione; il server ri-verifica autorizza su SP.
    return decideRichiesta({ ...data, approvatoreId: me.id });
  });

export const spCancelRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: { richiestaId: string; richiedenteId: string }) => {
    if (!input?.richiestaId) throw new Error("richiestaId mancante");
    if (!input?.richiedenteId) throw new Error("richiedenteId mancante");
    return {
      richiestaId: String(input.richiestaId),
      richiedenteId: String(input.richiedenteId),
    };
  })
  .handler(async ({ data }): Promise<SpRichiesta> => {
    const me = await currentUser();
    return cancelRichiesta({ richiestaId: data.richiestaId, richiedenteId: me.id });
  });

// ---------------------------------------------------------------------------
// Operatore (Sprint 3): elenco dipendenti + timbrature manuali
// ---------------------------------------------------------------------------

export const spGetDipendenti = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpDipendente[]> => {
    const me = await currentUser();
    assertCap(me.operatore || me.autorizza || me.ruolo === "responsabile" || isAdmin(me));
    return fetchDipendenti();
  },
);

// Import massivo Dipendenti da CSV/TSV incollato — SOLO amministratore.
// dryRun=true restituisce l'anteprima senza scrivere nulla.
export const spImportDipendenti = createServerFn({ method: "POST" })
  .inputValidator((input: { csv: string; dryRun?: boolean }) => {
    if (!input?.csv || typeof input.csv !== "string") throw new Error("Testo CSV mancante");
    return { csv: input.csv, dryRun: Boolean(input.dryRun) };
  })
  .handler(async ({ data }): Promise<ImportDipendentiResult> => {
    const me = await currentUser();
    assertCap(isAdmin(me));
    return importDipendenti(data.csv, data.dryRun);
  });

export const spCreateTimbraturaManuale = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTimbraturaManualeInput): CreateTimbraturaManualeInput => {
    if (!input?.operatoreId) throw new Error("operatoreId mancante");
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    if (!input?.evento) throw new Error("evento mancante");
    if (!input?.dataOra) throw new Error("dataOra mancante");
    return input;
  })
  .handler(async ({ data }): Promise<SpTimbratura> => {
    const me = await currentUser();
    // operatoreId dalla sessione; il server ri-verifica il flag Operatore su SP.
    return createTimbraturaManuale({ ...data, operatoreId: me.id });
  });

export const spCreateTurnoManuale = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTurnoManualeInput): CreateTurnoManualeInput => {
    if (!input?.operatoreId) throw new Error("operatoreId mancante");
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    if (!input?.entrata || !input?.uscita) throw new Error("entrata/uscita mancanti");
    return input;
  })
  .handler(async ({ data }): Promise<SpTimbratura[]> => {
    const me = await currentUser();
    return createTurnoManuale({ ...data, operatoreId: me.id });
  });

export const spGetAnomalie = createServerFn({ method: "GET" })
  .inputValidator((input?: { giorni?: number }) => ({
    giorni: input?.giorni && input.giorni > 0 ? Math.floor(input.giorni) : 14,
  }))
  .handler(async ({ data }): Promise<AnomaliaItem[]> => {
    const me = await currentUser();
    assertCap(me.operatore || isAdmin(me));
    return computeAnomalie(data.giorni);
  });

export const spGetTimbratureManuali = createServerFn({ method: "GET" })
  .inputValidator((input?: { giorni?: number }) => ({
    giorni: input?.giorni && input.giorni > 0 ? Math.floor(input.giorni) : 30,
  }))
  .handler(async ({ data }): Promise<TimbraturaManualeItem[]> => {
    const me = await currentUser();
    assertCap(me.autorizza || me.operatore || isAdmin(me));
    return fetchTimbratureManuali(data.giorni);
  });

export const spGetRendiconto = createServerFn({ method: "GET" })
  .inputValidator((input: { anno: number; mese: number }) => {
    const anno = Number(input?.anno);
    const mese = Number(input?.mese);
    if (!Number.isFinite(anno) || !Number.isFinite(mese) || mese < 1 || mese > 12)
      throw new Error("anno/mese non validi");
    return { anno, mese };
  })
  .handler(async ({ data }): Promise<RendicontoRiga[]> => {
    const me = await currentUser();
    assertCap(me.operatore || me.autorizza || me.ruolo === "responsabile" || isAdmin(me));
    return computeRendiconto(data.anno, data.mese);
  });
