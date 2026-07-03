// DR Portal — Server functions per SharePoint (via Lovable Connector Gateway).
// Sono l'unico entry point che il codice client usa per parlare con
// Microsoft Graph. Le implementazioni vere vivono in `sharepoint.server.ts`
// (bloccato dal bundle client dal suffisso .server.ts).

import { createServerFn } from "@tanstack/react-start";
import {
  clearSpDiscoveryCache,
  computeHealth,
  createTimbratura,
  discoverSharePoint,
  fetchDipendenti,
  fetchTimbratureOggi,
  getLastSyncAt,
  getSpLog,
  markSync,
  runSelfTest,
  type CreateTimbraturaInput,
  type EventoTimbratura,
  type SpHealth,
  type SpLogEvent,
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
    const [dipendenti, timbrature] = await Promise.all([
      fetchDipendenti(),
      fetchTimbratureOggi(),
    ]);
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