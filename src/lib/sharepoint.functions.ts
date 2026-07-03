// DR Portal — Server functions per SharePoint (via Lovable Connector Gateway).
// Sono l'unico entry point che il codice client usa per parlare con
// Microsoft Graph. Le implementazioni vere vivono in `sharepoint.server.ts`
// (bloccato dal bundle client dal suffisso .server.ts).

import { createServerFn } from "@tanstack/react-start";
import {
  createTimbratura,
  fetchDipendenti,
  fetchTimbratureOggi,
  isSpReady,
  loadSpConfig,
  type CreateTimbraturaInput,
  type EventoTimbratura,
  type SpDipendente,
  type SpTimbratura,
} from "./sharepoint.server";

export interface SpDiagnostics {
  configured: boolean;
  siteId: string;
  listDipendenti: string;
  listTimbrature: string;
  hasLovableKey: boolean;
  hasConnectionKey: boolean;
}

export const spGetDiagnostics = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpDiagnostics> => {
    const cfg = loadSpConfig();
    return {
      configured: isSpReady(cfg),
      siteId: cfg.siteId,
      listDipendenti: cfg.listDipendenti,
      listTimbrature: cfg.listTimbrature,
      hasLovableKey: Boolean(process.env.LOVABLE_API_KEY),
      hasConnectionKey: Boolean(process.env.MICROSOFT_SHAREPOINT_API_KEY),
    };
  },
);

export interface SpSnapshot {
  dipendenti: SpDipendente[];
  timbrature: SpTimbratura[];
}

export const spGetSnapshot = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpSnapshot> => {
    const [dipendenti, timbrature] = await Promise.all([
      fetchDipendenti(),
      fetchTimbratureOggi(),
    ]);
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