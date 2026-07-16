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
  computeRendicontoPeriodo,
  computeSaldoFerie,
  type SaldoFerieRiga,
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
  protectAllPins,
  uploadFileToLibrary,
  fetchDocumentiAll,
  fetchDocumentiForUser,
  createDocumento,
  fetchComunicazioniAll,
  fetchComunicazioniForUser,
  createComunicazione,
  markPresaVisione,
  fetchPreseVisione,
  fetchPreseVisioneForUser,
  getVapidPublicKey,
  savePushSubscription,
  sendPushToSede,
  sendPushToDipendente,
  enqueueEmail,
  parseEmails,
  getEmailDipendente,
  fetchVoci,
  fetchAcquisti,
  createAcquisto,
  decideAcquisto,
  type SpVoce,
  type SpAcquisto,
  type CreateAcquistoInput,
  type SpDocumento,
  type CreateDocumentoInput,
  type SpComunicazione,
  type CreateComunicazioneInput,
  type SpPresaVisione,
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

// Protezione massiva dei PIN (S3): converte in hash tutti i PIN in chiaro.
// SOLO amministratore.
export const spProtectPins = createServerFn({ method: "POST" }).handler(
  async (): Promise<{ protetti: number; giaProtetti: number }> => {
    assertCap(isAdmin(await currentUser()));
    return protectAllPins();
  },
);

// ---------------------------------------------------------------------------
// Documenti + Comunicazioni interne (Sprint 4)
// ---------------------------------------------------------------------------
// Capability di pubblicazione: responsabile, amministratore o operatore (DR000).
const canPubblicare = (me: ServerSessionUser) =>
  me.ruolo === "responsabile" || isAdmin(me) || me.operatore;

// Upload generico su libreria (solo pubblicatori). subfolder ristretto.
export const spUploadFile = createServerFn({ method: "POST" })
  .inputValidator((input: { subfolder: string; filename: string; contentBase64: string }) => {
    if (!input?.contentBase64) throw new Error("Contenuto file mancante");
    const allowed = new Set(["Documenti", "Comunicazioni"]);
    return {
      subfolder: allowed.has(input.subfolder) ? input.subfolder : "Documenti",
      filename: String(input.filename ?? "documento"),
      contentBase64: String(input.contentBase64),
    };
  })
  .handler(async ({ data }): Promise<UploadGiustificativoResult> => {
    const me = await currentUser();
    assertCap(canPubblicare(me));
    return uploadFileToLibrary(data.subfolder, data.filename, data.contentBase64);
  });

export const spGetDocumenti = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpDocumento[]> => {
    const me = await currentUser();
    // Pubblicatori vedono tutti i documenti; il dipendente solo i propri
    // (personali) + i generali per la sua sede / per tutti.
    if (canPubblicare(me)) return fetchDocumentiAll();
    return fetchDocumentiForUser(me.id, String(me.sede));
  },
);

export const spCreateDocumento = createServerFn({ method: "POST" })
  .inputValidator((input: Omit<CreateDocumentoInput, "caricatoDa">) => {
    if (!input?.categoria) throw new Error("Categoria mancante");
    if (!input?.titolo) throw new Error("Titolo mancante");
    if (!input?.file) throw new Error("File mancante");
    if (input.ambito !== "Personale" && input.ambito !== "Generale")
      throw new Error("Ambito non valido");
    if (input.ambito === "Personale" && !input.destinatarioId)
      throw new Error("Destinatario mancante per un documento personale");
    return input;
  })
  .handler(async ({ data }): Promise<SpDocumento> => {
    const me = await currentUser();
    assertCap(canPubblicare(me));
    const created = await createDocumento({
      ...data,
      caricatoDa: `${me.nome} ${me.cognome}`.trim(),
    });
    // Documento personale → notifica push al destinatario (best-effort).
    if (data.ambito === "Personale" && data.destinatarioId) {
      await sendPushToDipendente(data.destinatarioId, {
        title: "Nuovo documento",
        body: `${data.categoria}: ${data.titolo}`,
        url: "/documenti",
      }).catch(() => {});
    }
    return created;
  });

export const spGetComunicazioni = createServerFn({ method: "GET" }).handler(
  async (): Promise<SpComunicazione[]> => {
    const me = await currentUser();
    if (canPubblicare(me)) return fetchComunicazioniAll();
    return fetchComunicazioniForUser(String(me.sede));
  },
);

export const spCreateComunicazione = createServerFn({ method: "POST" })
  .inputValidator(
    (input: Omit<CreateComunicazioneInput, "autore"> & { destinatariEmail?: string }) => {
      if (!input?.titolo) throw new Error("Titolo mancante");
      if (!input?.testo) throw new Error("Testo mancante");
      if (input.tipo !== "Riunione" && input.tipo !== "Comunicazione")
        throw new Error("Tipo non valido");
      return {
        titolo: String(input.titolo),
        testo: String(input.testo),
        tipo: input.tipo,
        sede: String(input.sede ?? "Tutte"),
        allegato: input.allegato ? String(input.allegato) : undefined,
        richiedePresaVisione: Boolean(input.richiedePresaVisione),
        destinatariEmail: input.destinatariEmail ? String(input.destinatariEmail) : undefined,
      };
    },
  )
  .handler(async ({ data }): Promise<SpComunicazione & { pushEsito: string }> => {
    const me = await currentUser();
    assertCap(canPubblicare(me));
    const autore = `${me.nome} ${me.cognome}`.trim();
    const { destinatariEmail, ...comInput } = data;
    const created = await createComunicazione({ ...comInput, autore });
    // Notifica push ai dispositivi registrati della sede destinataria.
    // Best-effort: un errore qui non deve annullare la pubblicazione.
    // L'esito è restituito al pubblicatore per visibilità immediata.
    let pushEsito = "";
    try {
      const r = await sendPushToSede(data.sede, {
        title: data.tipo === "Riunione" ? "Nuova riunione" : "Nuova comunicazione",
        body: data.titolo,
        url: "/comunicazioni",
      });
      pushEsito =
        r.dispositivi === 0 && r.errori.length === 0
          ? "Nessun dispositivo registrato per le notifiche push."
          : `Notifiche push: ${r.sent} inviate su ${r.dispositivi} dispositivi${
              r.failed ? `, ${r.failed} fallite (${r.errori.join(" · ")})` : ""
            }${r.errori.length && !r.failed ? ` — ${r.errori.join(" · ")}` : ""}`;
    } catch (err) {
      pushEsito = `Notifiche push non inviate: ${err instanceof Error ? err.message : String(err)}`;
    }
    // Invio email (via coda + Power Automate) ai destinatari indicati.
    // Mittente = email di chi pubblica (fallback: casella Segreteria).
    const emails = parseEmails(destinatariEmail ?? "");
    if (emails.length) {
      const mittente = await getEmailDipendente(me.id).catch(() => "");
      const ok = await enqueueEmail({
        destinatari: emails,
        oggetto: `[DR Logistica] ${data.tipo === "Riunione" ? "Riunione" : "Comunicazione"}: ${data.titolo}`,
        corpo: `${data.testo}\n\n— ${autore}\nPortale DR Logistica: https://portal.drlogistica.it/comunicazioni`,
        allegato: data.allegato,
        mittente,
      }).catch(() => false);
      pushEsito += ok
        ? ` · Email in coda per ${emails.length} destinatari.`
        : " · Email NON accodate (lista CodaEmail assente).";
    }
    return { ...created, pushEsito };
  });

// Chi ha letto una comunicazione — solo pubblicatori.
export const spGetPreseVisione = createServerFn({ method: "GET" })
  .inputValidator((input: { comunicazioneId: string }) => {
    if (!input?.comunicazioneId) throw new Error("comunicazioneId mancante");
    return { comunicazioneId: String(input.comunicazioneId) };
  })
  .handler(async ({ data }): Promise<SpPresaVisione[]> => {
    const me = await currentUser();
    assertCap(canPubblicare(me));
    return fetchPreseVisione(data.comunicazioneId);
  });

// Comunicazioni già confermate dall'utente corrente.
export const spGetMiePreseVisione = createServerFn({ method: "GET" }).handler(
  async (): Promise<string[]> => {
    const me = await currentUser();
    return fetchPreseVisioneForUser(me.id);
  },
);

// ---------------------------------------------------------------------------
// Voci di spesa + Procurement (richieste di acquisto)
// ---------------------------------------------------------------------------
export const spGetVoci = createServerFn({ method: "GET" })
  .inputValidator((input: { ambito: string }) => {
    if (input?.ambito !== "Rimborso" && input?.ambito !== "Acquisto")
      throw new Error("Ambito non valido");
    return { ambito: input.ambito };
  })
  .handler(async ({ data }): Promise<SpVoce[]> => {
    await currentUser();
    return fetchVoci(data.ambito);
  });

export const spGetAcquisti = createServerFn({ method: "GET" })
  .inputValidator((input?: { mie?: boolean; stato?: string }) => ({
    mie: Boolean(input?.mie),
    stato: input?.stato ? String(input.stato) : undefined,
  }))
  .handler(async ({ data }): Promise<SpAcquisto[]> => {
    const me = await currentUser();
    if (data.mie) return fetchAcquisti({ richiedenteId: me.id, stato: data.stato });
    // Vista completa: approvatori (DR005) e admin.
    assertCap(me.autorizza || isAdmin(me));
    return fetchAcquisti({ stato: data.stato });
  });

export const spCreateAcquisto = createServerFn({ method: "POST" })
  .inputValidator((input: CreateAcquistoInput) => {
    if (!input?.macro) throw new Error("Voce di acquisto mancante");
    if (!input?.descrizione) throw new Error("Descrizione mancante");
    return {
      macro: String(input.macro),
      dettaglio: String(input.dettaglio ?? ""),
      descrizione: String(input.descrizione),
      importo:
        input.importo != null && Number.isFinite(Number(input.importo))
          ? Number(input.importo)
          : undefined,
    };
  })
  .handler(async ({ data }): Promise<SpAcquisto> => {
    const me = await currentUser();
    // La sede storica è ri-verificata dentro createAcquisto sul record SP.
    return createAcquisto(me.id, data);
  });

export const spDecideAcquisto = createServerFn({ method: "POST" })
  .inputValidator(
    (input: {
      acquistoId: string;
      decisione: "Approvata" | "Respinta";
      noteDecisione?: string;
    }) => {
      if (!input?.acquistoId) throw new Error("acquistoId mancante");
      if (input.decisione !== "Approvata" && input.decisione !== "Respinta")
        throw new Error("decisione non valida");
      return {
        acquistoId: String(input.acquistoId),
        decisione: input.decisione,
        noteDecisione: input.noteDecisione ? String(input.noteDecisione) : undefined,
      };
    },
  )
  .handler(async ({ data }): Promise<SpAcquisto> => {
    const me = await currentUser();
    // L'autorizzazione vera (DR005/admin) è verificata server-side su SP.
    return decideAcquisto({ ...data, approvatoreId: me.id });
  });

// --- Web Push: chiave pubblica + registrazione dispositivo -----------------
export const spGetVapidPublicKey = createServerFn({ method: "GET" }).handler(
  async (): Promise<{ publicKey: string }> => {
    await currentUser(); // basta una sessione valida
    return { publicKey: await getVapidPublicKey() };
  },
);

export const spSavePushSubscription = createServerFn({ method: "POST" })
  .inputValidator((input: { endpoint: string; p256dh: string; auth: string }) => {
    if (!input?.endpoint?.startsWith("https://")) throw new Error("Endpoint non valido");
    if (!input?.p256dh || !input?.auth) throw new Error("Chiavi subscription mancanti");
    return {
      endpoint: String(input.endpoint),
      p256dh: String(input.p256dh),
      auth: String(input.auth),
    };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    const me = await currentUser();
    await savePushSubscription(me.id, String(me.sede), data);
    return { ok: true };
  });

export const spMarkPresaVisione = createServerFn({ method: "POST" })
  .inputValidator((input: { comunicazioneId: string }) => {
    if (!input?.comunicazioneId) throw new Error("comunicazioneId mancante");
    return { comunicazioneId: String(input.comunicazioneId) };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    const me = await currentUser();
    await markPresaVisione(data.comunicazioneId, me.id, `${me.nome} ${me.cognome}`.trim());
    return { ok: true };
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

// Rendiconto su periodo arbitrario (settimana fiscale / settimana del mese).
export const spGetRendicontoPeriodo = createServerFn({ method: "GET" })
  .inputValidator((input: { from: string; to: string }) => {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    if (!re.test(input?.from ?? "") || !re.test(input?.to ?? ""))
      throw new Error("Periodo non valido");
    const days =
      (new Date(`${input.to}T00:00:00`).getTime() - new Date(`${input.from}T00:00:00`).getTime()) /
      86400000;
    if (days < 0 || days > 45) throw new Error("Periodo non valido (max 45 giorni)");
    return { from: input.from, to: input.to };
  })
  .handler(async ({ data }): Promise<RendicontoRiga[]> => {
    const me = await currentUser();
    assertCap(me.operatore || me.autorizza || me.ruolo === "responsabile" || isAdmin(me));
    return computeRendicontoPeriodo(data.from, data.to);
  });

export const spGetSaldoFerie = createServerFn({ method: "GET" })
  .inputValidator((input: { anno: number }) => {
    const anno = Number(input?.anno);
    if (!Number.isFinite(anno)) throw new Error("anno non valido");
    return { anno };
  })
  .handler(async ({ data }): Promise<SaldoFerieRiga[]> => {
    const me = await currentUser();
    assertCap(me.operatore || me.autorizza || me.ruolo === "responsabile" || isAdmin(me));
    return computeSaldoFerie(data.anno);
  });
