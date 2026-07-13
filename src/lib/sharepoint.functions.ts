// DR Portal — Server functions per SharePoint (via Lovable Connector Gateway).
// Sono l'unico entry point che il codice client usa per parlare con
// Microsoft Graph. Le implementazioni vere vivono in `sharepoint.server.ts`
// (bloccato dal bundle client dal suffisso .server.ts).

import { createServerFn } from "@tanstack/react-start";
import {
  cancelRichiesta,
  clearSpDiscoveryCache,
  computeHealth,
  createRichiesta,
  createTimbratura,
  createTimbraturaManuale,
  createTurnoManuale,
  decideRichiesta,
  discoverSharePoint,
  fetchDipendenti,
  fetchRichieste,
  fetchTimbratureOggi,
  getLastSyncAt,
  getSpLog,
  loginByCodicePin,
  markSync,
  runSelfTest,
  type CreateRichiestaInput,
  type CreateTimbraturaInput,
  type CreateTimbraturaManualeInput,
  type CreateTurnoManualeInput,
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
  .handler(async ({ data }) => createTimbratura(data));

export const spRunSelfTest = createServerFn({ method: "POST" }).handler(
  async (): Promise<SpSelfTestResult> => runSelfTest(),
);

export const spLogin = createServerFn({ method: "POST" })
  .inputValidator((input: { codice: string; pin: string }) => {
    if (typeof input?.codice !== "string" || typeof input?.pin !== "string") {
      throw new Error("Codice o PIN non validi.");
    }
    return { codice: input.codice, pin: input.pin };
  })
  .handler(async ({ data }): Promise<LoginResult> => loginByCodicePin(data.codice, data.pin));

// ---------------------------------------------------------------------------
// Richieste (Sprint 2)
// ---------------------------------------------------------------------------

export const spGetRichieste = createServerFn({ method: "GET" })
  .inputValidator((input?: RichiesteFilter): RichiesteFilter => ({
    richiedenteId: input?.richiedenteId ? String(input.richiedenteId) : undefined,
    stato: input?.stato ? String(input.stato) : undefined,
  }))
  .handler(async ({ data }): Promise<SpRichiesta[]> => fetchRichieste(data));

export const spCreateRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: CreateRichiestaInput): CreateRichiestaInput => {
    if (!input?.richiedenteId) throw new Error("richiedenteId mancante");
    if (!input?.tipo) throw new Error("tipo mancante");
    return input;
  })
  .handler(async ({ data }): Promise<SpRichiesta> => createRichiesta(data));

export const spDecideRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: DecideRichiestaInput): DecideRichiestaInput => {
    if (!input?.richiestaId) throw new Error("richiestaId mancante");
    if (!input?.approvatoreId) throw new Error("approvatoreId mancante");
    if (input.decisione !== "Approvata" && input.decisione !== "Respinta")
      throw new Error("decisione non valida");
    return input;
  })
  .handler(async ({ data }): Promise<SpRichiesta> => decideRichiesta(data));

export const spCancelRichiesta = createServerFn({ method: "POST" })
  .inputValidator((input: { richiestaId: string; richiedenteId: string }) => {
    if (!input?.richiestaId) throw new Error("richiestaId mancante");
    if (!input?.richiedenteId) throw new Error("richiedenteId mancante");
    return {
      richiestaId: String(input.richiestaId),
      richiedenteId: String(input.richiedenteId),
    };
  })
  .handler(async ({ data }): Promise<SpRichiesta> => cancelRichiesta(data));

// ---------------------------------------------------------------------------
// Operatore (Sprint 3): elenco dipendenti + timbrature manuali
// ---------------------------------------------------------------------------

export const spGetDipendenti = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpDipendente[]> => fetchDipendenti(),
);

export const spCreateTimbraturaManuale = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTimbraturaManualeInput): CreateTimbraturaManualeInput => {
    if (!input?.operatoreId) throw new Error("operatoreId mancante");
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    if (!input?.evento) throw new Error("evento mancante");
    if (!input?.dataOra) throw new Error("dataOra mancante");
    return input;
  })
  .handler(async ({ data }): Promise<SpTimbratura> => createTimbraturaManuale(data));

export const spCreateTurnoManuale = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTurnoManualeInput): CreateTurnoManualeInput => {
    if (!input?.operatoreId) throw new Error("operatoreId mancante");
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    if (!input?.entrata || !input?.uscita) throw new Error("entrata/uscita mancanti");
    return input;
  })
  .handler(async ({ data }): Promise<SpTimbratura[]> => createTurnoManuale(data));
