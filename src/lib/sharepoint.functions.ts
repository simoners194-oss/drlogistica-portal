// DR Portal — Server functions per SharePoint (via Lovable Connector Gateway).
// Sono l'unico entry point che il codice client usa per parlare con
// Microsoft Graph. Le implementazioni vere vivono in `sharepoint.server.ts`
// (bloccato dal bundle client dal suffisso .server.ts).

import { createServerFn } from "@tanstack/react-start";
import { normalizeRuolo } from "./session";
import { isSupervisoreGlobale } from "./richieste-logic";
import { arubaProvaConnessione, type ArubaProbeResult } from "./aruba.server";
import { lunediDellaSettimana, ymd } from "./rendiconto-logic";
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
  fetchVoci,
  fetchAcquisti,
  createAcquisto,
  decideAcquisto,
  fetchMovimenti,
  fetchMovimentiChiavi,
  importMovimenti,
  updateMovimento,
  fetchImportStorico,
  annullaImport,
  LEGACY_IMPORT_ID,
  fetchRegoleFinanza,
  createRegolaFinanza,
  deleteRegolaFinanza,
  applicaRegolaAiMovimenti,
  annullaRegolaAiMovimenti,
  type RegolaFinanza,
  fetchFatture,
  importFatture,
  fetchTerminiPagamento,
  setIncassiAruba,
  setRettificaNumero,
  setIncassoManuale,
  trovaFattureSenzaCliente,
  importTermini,
  deleteTermine,
  copiaTerminiSuFornitori,
  ultimoAggiornamentoFatture,
  fetchRegoleFatture,
  createRegolaFattura,
  deleteRegolaFattura,
  setClassificazione,
  eliminaFatture,
  fetchAbbinamenti,
  createAbbinamenti,
  deleteAbbinamento,
  type SpFattura,
  type ImportFattureResult,
  type FatturaRaw,
  type DirezioneFattura,
  type TerminePagamento,
  type AbbinamentoIncasso,
  getArubaStato,
  saveArubaCredenziali,
  type ArubaStato,
  getEbStato,
  saveEbApp,
  ebProvaApplicazione,
  ebSaldoAttuale,
  ebImpostaSaldoManuale,
  type EbSaldoInfo,
  ebAvviaCollegamento,
  ebCompletaCollegamento,
  ebScegliConto,
  ebTagliaUltimoGiorno,
  ebSincronizza,
  ebSincronizzaProgrammato,
  ebCronToken,
  cronToken,
  promemoriaUsciteAperte,
  type EbStato,
  type EbSyncResult,
  getCodiceDipendente,
  type SpMovimento,
  type MovimentiFilter,
  type ImportMovimentoRow,
  type ImportMovimentiResult,
  type UpdateMovimentoInput,
  type ImportStoricoRiga,
  type SpVoce,
  type SpAcquisto,
  type CreateAcquistoInput,
  type SpDocumento,
  type CreateDocumentoInput,
  type SpComunicazione,
  type CreateComunicazioneInput,
  type SpPresaVisione,
  fetchTimbratureRecenti,
  fetchTimbratureDaISO,
  fetchCorrezioni,
  createCorrezione,
  decideCorrezione,
  parseOrariProposti,
  type SpCorrezione,
  annullaUltimaTimbratura,
  fetchTimbratureGiorno,
  resocontoGiorno,
  type ResocontoGiornoRiga,
  deleteTimbratura,
  deleteTimbraturaOperatore,
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
    const me = await currentUser();
    // Garantisce discovery prima delle chiamate in parallelo (evita doppia
    // esecuzione della discovery quando la cache è fredda).
    await discoverSharePoint();
    // Finestra ~36h, non solo oggi: il turno notturno aperto ieri sera deve
    // risultare "presente" e chiudibile anche dopo mezzanotte.
    const [dipendenti, timbrature] = await Promise.all([
      fetchDipendenti(),
      fetchTimbratureRecenti(),
    ]);
    markSync();
    // Scope per ruolo: il dipendente vede SOLO il proprio record e le proprie
    // timbrature. Responsabile/operatore/autorizzatore/amministratore vedono
    // lo snapshot completo, coerente con le viste HR/operative.
    const puoVedereTutti =
      me.operatore || me.autorizza || me.ruolo === "responsabile" || isAdmin(me);
    if (!puoVedereTutti) {
      return {
        dipendenti: dipendenti.filter((d) => d.id === me.id),
        timbrature: timbrature.filter((t) => t.dipendenteId === me.id),
      };
    }
    return { dipendenti, timbrature };
  },
);

export const spCreateTimbratura = createServerFn({ method: "POST" })
  .inputValidator((input: CreateTimbraturaInput): CreateTimbraturaInput => {
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    const validi: EventoTimbratura[] = ["entrata", "inizio-pausa", "fine-pausa", "uscita"];
    if (!validi.includes(input.evento)) throw new Error("evento non valido");
    const dataOraClient = input.dataOraClient
      ? String(input.dataOraClient).slice(0, 40)
      : undefined;
    return {
      ...input,
      dataOraClient,
      note: input.note ? String(input.note).slice(0, 200) : undefined,
    };
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
        preposto: d.preposto,
        codice: d.codice,
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
    // Mittente SEMPRE la casella Segreteria (unica con "Invia come" concesso
    // alla connessione del flusso); chi pubblica è indicato nella firma.
    const emails = parseEmails(destinatariEmail ?? "");
    if (emails.length) {
      // Il corpo deve essere AUTOSUFFICIENTE: i destinatari possono essere
      // esterni senza accesso al portale. Niente rimandi al portale; l'eventuale
      // allegato è linkato esplicitamente (NB: richiede accesso SharePoint).
      const ok = await enqueueEmail({
        destinatari: emails,
        oggetto: `[DR Logistica] ${data.tipo === "Riunione" ? "Riunione" : "Comunicazione"}: ${data.titolo}`,
        corpo:
          `${data.testo}\n\n— ${autore}\nDR Logistica` +
          (data.allegato ? `\n\nAllegato: ${data.allegato}` : ""),
        allegato: data.allegato,
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

// ---------------------------------------------------------------------------
// Finanza (movimenti bancari) — SOLO direttore DR005 o amministratore.
// ---------------------------------------------------------------------------
// Il cookie firmato identifica l'utente; il codice DR005 è ri-verificato sul
// record SharePoint autorevole (come per l'approvazione Procurement).
async function assertDirettore(me: ServerSessionUser): Promise<void> {
  if (isAdmin(me)) return;
  const codice = await getCodiceDipendente(me.id);
  assertCap(isSupervisoreGlobale(codice));
}

export const spGetMovimenti = createServerFn({ method: "GET" })
  .inputValidator((input?: MovimentiFilter): MovimentiFilter => {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    return {
      from: input?.from && re.test(input.from) ? input.from : undefined,
      to: input?.to && re.test(input.to) ? input.to : undefined,
      soloDaVerificare: Boolean(input?.soloDaVerificare),
    };
  })
  .handler(async ({ data }): Promise<SpMovimento[]> => {
    await assertDirettore(await currentUser());
    return fetchMovimenti(data);
  });

export const spGetMovimentiChiavi = createServerFn({ method: "GET" }).handler(
  async (): Promise<string[]> => {
    await assertDirettore(await currentUser());
    return fetchMovimentiChiavi();
  },
);

export const spImportMovimenti = createServerFn({ method: "POST" })
  .inputValidator((input: { rows: ImportMovimentoRow[]; importId?: string }) => {
    if (!Array.isArray(input?.rows) || input.rows.length === 0)
      throw new Error("Nessun movimento da importare");
    if (input.rows.length > 150) throw new Error("Blocco troppo grande (max 150 movimenti)");
    const importId = String(input.importId ?? "").slice(0, 60);
    const re = /^\d{4}-\d{2}-\d{2}$/;
    const rows = input.rows.map((r): ImportMovimentoRow => {
      if (!re.test(r?.dataContabile ?? "") || !re.test(r?.dataValuta ?? ""))
        throw new Error("Data movimento non valida");
      const importo = Number(r.importo);
      if (!Number.isFinite(importo)) throw new Error("Importo movimento non valido");
      const occ = Math.floor(Number(r.occ));
      if (!Number.isFinite(occ) || occ < 1) throw new Error("Occorrenza non valida");
      return {
        dataContabile: r.dataContabile,
        dataValuta: r.dataValuta,
        importo,
        divisa: String(r.divisa ?? "EUR").slice(0, 8),
        causale: String(r.causale ?? "").slice(0, 20),
        descrizione: String(r.descrizione ?? "").slice(0, 1000),
        occ,
      };
    });
    return { rows, importId };
  })
  .handler(async ({ data }): Promise<ImportMovimentiResult> => {
    await assertDirettore(await currentUser());
    return importMovimenti(data.rows, data.importId);
  });

export const spGetImportStorico = createServerFn({ method: "GET" }).handler(
  async (): Promise<ImportStoricoRiga[]> => {
    await assertDirettore(await currentUser());
    return fetchImportStorico();
  },
);

// Annulla un import: cancella un blocco di movimenti per chiamata; il client
// ripete finché `rimanenti` è 0. LEGACY_IMPORT_ID annulla il gruppo senza id.
export const spAnnullaImport = createServerFn({ method: "POST" })
  .inputValidator((input: { importId: string }) => {
    const importId = String(input?.importId ?? "").trim();
    if (!importId) throw new Error("importId mancante");
    if (
      importId !== LEGACY_IMPORT_ID &&
      !importId.startsWith("IMP-") &&
      !importId.startsWith("SYNC-")
    )
      throw new Error("importId non valido");
    return { importId };
  })
  .handler(async ({ data }): Promise<{ eliminati: number; rimanenti: number }> => {
    await assertDirettore(await currentUser());
    return annullaImport(data.importId);
  });

export const spUpdateMovimento = createServerFn({ method: "POST" })
  .inputValidator((input: UpdateMovimentoInput): UpdateMovimentoInput => {
    if (!input?.movimentoId) throw new Error("movimentoId mancante");
    return {
      movimentoId: String(input.movimentoId),
      tipologia: input.tipologia !== undefined ? String(input.tipologia).slice(0, 60) : undefined,
      cliente: input.cliente !== undefined ? String(input.cliente).slice(0, 120) : undefined,
      nrFattura: input.nrFattura !== undefined ? String(input.nrFattura).slice(0, 160) : undefined,
      note: input.note !== undefined ? String(input.note).slice(0, 500) : undefined,
      daVerificare: input.daVerificare !== undefined ? Boolean(input.daVerificare) : undefined,
    };
  })
  .handler(async ({ data }): Promise<SpMovimento> => {
    await assertDirettore(await currentUser());
    return updateMovimento(data);
  });

// --- Regole apprese Finanza -------------------------------------------------
function validateRegola(input: Partial<RegolaFinanza>): RegolaFinanza {
  const pattern = String(input?.pattern ?? "")
    .trim()
    .slice(0, 120);
  if (!pattern) throw new Error("Pattern mancante");
  const tipologia = input.tipologia ? String(input.tipologia).trim().slice(0, 60) : undefined;
  const cliente = input.cliente ? String(input.cliente).trim().slice(0, 120) : undefined;
  if (!tipologia && !cliente)
    throw new Error("La regola deve impostare tipologia o nome controparte");
  return {
    pattern,
    campo: input.campo === "descrizione" ? "descrizione" : "cliente",
    modo: input.modo === "contiene" ? "contiene" : "esatto",
    tipologia,
    cliente,
  };
}

export const spGetRegoleFinanza = createServerFn({ method: "GET" }).handler(
  async (): Promise<RegolaFinanza[]> => {
    await assertDirettore(await currentUser());
    return fetchRegoleFinanza();
  },
);

export const spCreateRegolaFinanza = createServerFn({ method: "POST" })
  .inputValidator(validateRegola)
  .handler(async ({ data }): Promise<RegolaFinanza> => {
    await assertDirettore(await currentUser());
    return createRegolaFinanza(data);
  });

export const spDeleteRegolaFinanza = createServerFn({ method: "POST" })
  .inputValidator((input: { regolaId: string }) => {
    if (!input?.regolaId) throw new Error("regolaId mancante");
    return { regolaId: String(input.regolaId) };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await deleteRegolaFinanza(data.regolaId);
    return { ok: true };
  });

// Applica una regola all'archivio esistente, un blocco per chiamata (il
// client ripete finché rimanenti=0).
export const spApplicaRegolaFinanza = createServerFn({ method: "POST" })
  .inputValidator(validateRegola)
  .handler(async ({ data }): Promise<{ aggiornati: number; rimanenti: number }> => {
    await assertDirettore(await currentUser());
    return applicaRegolaAiMovimenti(data);
  });

// Ripristina i movimenti toccati da una regola GIÀ eliminata: si passa la
// definizione della regola (non l'id) e si ripete finché rimanenti=0.
export const spAnnullaRegolaFinanza = createServerFn({ method: "POST" })
  .inputValidator(validateRegola)
  .handler(async ({ data }): Promise<{ aggiornati: number; rimanenti: number }> => {
    await assertDirettore(await currentUser());
    return annullaRegolaAiMovimenti(data);
  });

// --- Finanza → Fatture emesse, termini, abbinamenti (solo direttore) --------

export const spGetAggiornamentoFatture = createServerFn({ method: "GET" })
  .inputValidator((input?: { direzione?: string }) => ({
    direzione: (input?.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
  }))
  .handler(async ({ data }): Promise<{ aggiornatoAl: string | null }> => {
    await assertDirettore(await currentUser());
    return { aggiornatoAl: await ultimoAggiornamentoFatture(data.direzione) };
  });

export const spGetFatture = createServerFn({ method: "GET" })
  .inputValidator((input?: { direzione?: string }) => ({
    direzione: (input?.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
  }))
  .handler(async ({ data }): Promise<SpFattura[]> => {
    await assertDirettore(await currentUser());
    return fetchFatture(data.direzione);
  });

export const spImportFatture = createServerFn({ method: "POST" })
  .inputValidator((input: { rows: FatturaRaw[]; direzione?: string }) => {
    if (!Array.isArray(input?.rows) || input.rows.length === 0)
      throw new Error("Nessuna fattura da importare");
    if (input.rows.length > 150) throw new Error("Blocco troppo grande (max 150 fatture)");
    const direzione = (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura;
    const re = /^\d{4}-\d{2}-\d{2}$/;
    const rows = input.rows.map((r): FatturaRaw => {
      if (!r?.nomeFile?.trim()) throw new Error("Nome file mancante");
      if (!re.test(r?.dataDocumento ?? "")) throw new Error("Data documento non valida");
      const num = (v: unknown) => (Number.isFinite(Number(v)) ? Number(v) : 0);
      return {
        nomeFile: String(r.nomeFile).trim().slice(0, 100),
        numero: String(r.numero ?? "").slice(0, 40),
        idSdi: String(r.idSdi ?? "").slice(0, 30),
        dataInvio: re.test(r.dataInvio ?? "") ? r.dataInvio : r.dataDocumento,
        dataDocumento: r.dataDocumento,
        tipoDocumento: String(r.tipoDocumento ?? "").slice(0, 60),
        cliente: String(r.cliente ?? "").slice(0, 120),
        piva: String(r.piva ?? "").slice(0, 20),
        metodoPagamento: String(r.metodoPagamento ?? "").slice(0, 60),
        imponibile: num(r.imponibile),
        iva: num(r.iva),
        totale: num(r.totale),
        netto: num(r.netto),
        statoSdI: String(r.statoSdI ?? "").slice(0, 40),
        direzione,
        scadenza: re.test(r.scadenza ?? "") ? r.scadenza : undefined,
        incassoAruba: r.incassoAruba ? String(r.incassoAruba).slice(0, 30) : undefined,
        dataIncasso: re.test(r.dataIncasso ?? "") ? r.dataIncasso : undefined,
        rettificaNumero: r.rettificaNumero ? String(r.rettificaNumero).slice(0, 60) : undefined,
        // Classificazione gestionale dal report compilato a mano.
        meseCompetenza: r.meseCompetenza ? String(r.meseCompetenza).slice(0, 40) : undefined,
        tipologiaCosto: r.tipologiaCosto ? String(r.tipologiaCosto).slice(0, 120) : undefined,
        clienteRif: r.clienteRif ? String(r.clienteRif).slice(0, 80) : undefined,
        oggetto: r.oggetto ? String(r.oggetto).slice(0, 500) : undefined,
      };
    });
    return { rows, direzione };
  })
  .handler(async ({ data }): Promise<ImportFattureResult> => {
    await assertDirettore(await currentUser());
    return importFatture(data.rows, data.direzione);
  });

// Incassi registrati su Aruba (report movimenti): importi per fattura.
export const spSetIncassiAruba = createServerFn({ method: "POST" })
  .inputValidator(
    (input: {
      righe: { nomeFile: string; incassato: number; ultimaData?: string; id?: string }[];
      direzione?: string;
    }) => {
      if (!Array.isArray(input?.righe) || input.righe.length === 0)
        throw new Error("Nessun incasso da registrare");
      if (input.righe.length > 200) throw new Error("Blocco troppo grande (max 200)");
      const re = /^\d{4}-\d{2}-\d{2}$/;
      return {
        direzione: (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
        righe: input.righe.map((r) => {
          const nomeFile = String(r?.nomeFile ?? "").trim();
          const incassato = Number(r?.incassato);
          if (!nomeFile || !Number.isFinite(incassato)) throw new Error("Riga incasso non valida");
          const ultimaData = r.ultimaData && re.test(r.ultimaData) ? r.ultimaData : undefined;
          const id = String(r.id ?? "").trim() || undefined;
          return { nomeFile, incassato: Math.round(incassato * 100) / 100, ultimaData, id };
        }),
      };
    },
  )
  .handler(async ({ data }): Promise<{ aggiornate: number; errori: string[] }> => {
    await assertDirettore(await currentUser());
    return setIncassiAruba(data.righe, data.direzione);
  });

// Collegamento manuale NOTA DI CREDITO → fattura rettificata: serve quando lo
// storno è stato fatto dentro Aruba e l'XML non porta il riferimento.
export const spSetRettificaNumero = createServerFn({ method: "POST" })
  .inputValidator((input: { nomeFile: string; numeroFattura: string; direzione?: string }) => {
    const nomeFile = String(input?.nomeFile ?? "").trim();
    if (!nomeFile) throw new Error("Nota di credito non indicata");
    // Vuoto = scollega. Il numero è testo libero (le numerazioni cambiano
    // formato fra anni e sezionali): si limita solo la lunghezza.
    const numeroFattura = String(input?.numeroFattura ?? "").trim();
    if (numeroFattura.length > 60) throw new Error("Numero fattura troppo lungo");
    return {
      nomeFile,
      numeroFattura,
      direzione: (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
    };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await setRettificaNumero(data.nomeFile, data.numeroFattura, data.direzione);
    return { ok: true };
  });

// Stato d'incasso corretto A MANO (senza aspettare il prossimo report).
export const spSetIncassoManuale = createServerFn({ method: "POST" })
  .inputValidator(
    (input: { nomeFile: string; stato: string; direzione?: string; dataIncasso?: string }) => {
      const nomeFile = String(input?.nomeFile ?? "").trim();
      if (!nomeFile) throw new Error("Fattura non indicata");
      if (input.stato !== "Incassata" && input.stato !== "Non incassata")
        throw new Error("Stato non valido");
      const dataIncasso =
        input.dataIncasso && /^\d{4}-\d{2}-\d{2}$/.test(input.dataIncasso)
          ? input.dataIncasso
          : undefined;
      return {
        nomeFile,
        stato: input.stato as "Incassata" | "Non incassata",
        direzione: (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
        dataIncasso,
      };
    },
  )
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await setIncassoManuale(data.nomeFile, data.stato, data.direzione, data.dataIncasso);
    return { ok: true };
  });

// Pulizia dei documenti senza controparte (file letto col tracciato sbagliato).
export const spTrovaFattureSenzaCliente = createServerFn({ method: "POST" })
  .inputValidator((input: { direzione?: string }) => ({
    direzione: (input?.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
  }))
  .handler(async ({ data }) => {
    await assertDirettore(await currentUser());
    return trovaFattureSenzaCliente(data.direzione);
  });

export const spEliminaFatture = createServerFn({ method: "POST" })
  .inputValidator((input: { ids: string[]; direzione?: string }) => {
    if (!Array.isArray(input?.ids) || input.ids.length === 0)
      throw new Error("Nessun documento da eliminare");
    if (input.ids.length > 80) throw new Error("Blocco troppo grande (max 80)");
    return {
      ids: input.ids.map((i) => String(i).trim()).filter(Boolean),
      direzione: (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
    };
  })
  .handler(async ({ data }) => {
    await assertDirettore(await currentUser());
    return eliminaFatture(data.ids, data.direzione);
  });

// Termini di pagamento dal foglio contratti (upsert per cliente).
export const spImportTermini = createServerFn({ method: "POST" })
  .inputValidator(
    (input: {
      rows: {
        cliente: string;
        giorni: number;
        direzione?: string;
        email?: string;
        oggetto?: string;
      }[];
    }) => {
      if (!Array.isArray(input?.rows) || input.rows.length === 0)
        throw new Error("Nessun termine da importare");
      if (input.rows.length > 200) throw new Error("Blocco troppo grande (max 200)");
      return {
        rows: input.rows
          .map((r) => ({
            cliente: String(r?.cliente ?? "").trim(),
            giorni: Number(r?.giorni),
            direzione: (r?.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
            email: r?.email === undefined ? undefined : String(r.email).trim().slice(0, 120),
            oggetto: r?.oggetto ? String(r.oggetto).trim().slice(0, 120) : undefined,
          }))
          // 0 giorni (a vista) e' ammesso solo per le regole con parola chiave.
          .filter(
            (r) =>
              r.cliente &&
              Number.isFinite(r.giorni) &&
              (r.giorni > 0 || (r.giorni === 0 && !!r.oggetto)),
          ),
      };
    },
  )
  .handler(async ({ data }) => {
    await assertDirettore(await currentUser());
    return importTermini(data.rows);
  });

export const spDeleteTermine = createServerFn({ method: "POST" })
  .inputValidator((input: { cliente: string; direzione?: string; oggetto?: string }) => {
    const cliente = String(input?.cliente ?? "").trim();
    if (!cliente) throw new Error("Cliente non indicato");
    return {
      cliente,
      direzione: (input?.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
      oggetto: input?.oggetto ? String(input.oggetto).trim().slice(0, 120) : undefined,
    };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await deleteTermine(data.cliente, data.direzione, data.oggetto);
    return { ok: true };
  });

// Tasto del direttore: copia i termini dei clienti sui fornitori omonimi.
export const spCopiaTerminiSuFornitori = createServerFn({ method: "POST" }).handler(async () => {
  await assertDirettore(await currentUser());
  return copiaTerminiSuFornitori();
});

// Regole di classificazione delle passive + classificazione manuale.
export const spGetRegoleFatture = createServerFn({ method: "GET" }).handler(async () => {
  await assertDirettore(await currentUser());
  return fetchRegoleFatture();
});

export const spCreateRegolaFattura = createServerFn({ method: "POST" })
  .inputValidator((input: { fornitore: string; tipologia?: string; clienteRif?: string }) => ({
    fornitore: String(input?.fornitore ?? "")
      .trim()
      .slice(0, 120),
    tipologia:
      String(input?.tipologia ?? "")
        .trim()
        .slice(0, 120) || undefined,
    clienteRif:
      String(input?.clienteRif ?? "")
        .trim()
        .slice(0, 80) || undefined,
  }))
  .handler(async ({ data }) => {
    await assertDirettore(await currentUser());
    return createRegolaFattura(data);
  });

export const spDeleteRegolaFattura = createServerFn({ method: "POST" })
  .inputValidator((input: { id: string }) => {
    const id = String(input?.id ?? "").trim();
    if (!id) throw new Error("Regola non indicata");
    return { id };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await deleteRegolaFattura(data.id);
    return { ok: true };
  });

export const spSetClassificazione = createServerFn({ method: "POST" })
  .inputValidator(
    (input: {
      nomeFile: string;
      direzione?: string;
      meseCompetenza?: string;
      tipologiaCosto?: string;
      clienteRif?: string;
    }) => {
      const nomeFile = String(input?.nomeFile ?? "").trim();
      if (!nomeFile) throw new Error("Fattura non indicata");
      const campo = (v: unknown, max: number) =>
        v === undefined ? undefined : String(v).trim().slice(0, max);
      return {
        nomeFile,
        direzione: (input.direzione === "Ricevuta" ? "Ricevuta" : "Emessa") as DirezioneFattura,
        meseCompetenza: campo(input.meseCompetenza, 40),
        tipologiaCosto: campo(input.tipologiaCosto, 120),
        clienteRif: campo(input.clienteRif, 80),
      };
    },
  )
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await setClassificazione(data.nomeFile, data.direzione, {
      meseCompetenza: data.meseCompetenza,
      tipologiaCosto: data.tipologiaCosto,
      clienteRif: data.clienteRif,
    });
    return { ok: true };
  });

export const spGetTerminiPagamento = createServerFn({ method: "GET" }).handler(
  async (): Promise<TerminePagamento[]> => {
    await assertDirettore(await currentUser());
    return fetchTerminiPagamento();
  },
);

export const spGetAbbinamenti = createServerFn({ method: "GET" }).handler(
  async (): Promise<AbbinamentoIncasso[]> => {
    await assertDirettore(await currentUser());
    return fetchAbbinamenti();
  },
);

export const spCreateAbbinamenti = createServerFn({ method: "POST" })
  .inputValidator((input: { rows: AbbinamentoIncasso[] }) => {
    if (!Array.isArray(input?.rows) || input.rows.length === 0)
      throw new Error("Nessun abbinamento da salvare");
    if (input.rows.length > 150) throw new Error("Blocco troppo grande (max 150 abbinamenti)");
    const rows = input.rows.map((r): AbbinamentoIncasso => {
      if (!r?.fatturaFile || !r?.movimentoChiave) throw new Error("Abbinamento incompleto");
      const importo = Number(r.importo);
      if (!Number.isFinite(importo) || importo <= 0) throw new Error("Importo non valido");
      return {
        fatturaFile: String(r.fatturaFile).slice(0, 100),
        movimentoChiave: String(r.movimentoChiave).slice(0, 200),
        importo: Math.round(importo * 100) / 100,
        origine: r.origine === "Manuale" ? "Manuale" : r.origine === "FIFO" ? "FIFO" : "Auto",
      };
    });
    return { rows };
  })
  .handler(async ({ data }): Promise<{ creati: number; scartati: number; errori: string[] }> => {
    await assertDirettore(await currentUser());
    return createAbbinamenti(data.rows);
  });

export const spDeleteAbbinamento = createServerFn({ method: "POST" })
  .inputValidator((input: { abbinamentoId: string }) => {
    if (!input?.abbinamentoId) throw new Error("abbinamentoId mancante");
    return { abbinamentoId: String(input.abbinamentoId) };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await deleteAbbinamento(data.abbinamentoId);
    return { ok: true };
  });

// --- Collegamento Aruba Fatturazione Elettronica (solo direttore) -----------
// Sola lettura in produzione: signin + ricerca. L'INVIO fatture non esiste.
// Il probe restituisce la forma reale della risposta per finalizzare il sync.

export const spGetArubaStato = createServerFn({ method: "GET" }).handler(
  async (): Promise<ArubaStato> => {
    await assertDirettore(await currentUser());
    return getArubaStato();
  },
);

export const spSetArubaCredenziali = createServerFn({ method: "POST" })
  .inputValidator((input: { username: string; password: string }) => {
    const username = String(input?.username ?? "").trim();
    const password = String(input?.password ?? "");
    if (!username || username.length > 120) throw new Error("Username non valido");
    if (!password || password.length > 200) throw new Error("Password non valida");
    return { username, password };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await saveArubaCredenziali(data.username, data.password);
    return { ok: true };
  });

export const spArubaProvaConnessione = createServerFn({ method: "POST" }).handler(
  async (): Promise<ArubaProbeResult> => {
    await assertDirettore(await currentUser());
    return arubaProvaConnessione();
  },
);

// --- Collegamento banca Enable Banking / PSD2 (solo direttore) ---------------
// SOLA LETTURA del conto aziendale: autorizzazione con SCA della banca,
// sincronizzazione dei movimenti nella lista MovimentiBancari.

export const spEbStato = createServerFn({ method: "GET" }).handler(async (): Promise<EbStato> => {
  await assertDirettore(await currentUser());
  return getEbStato();
});

export const spEbSalvaApp = createServerFn({ method: "POST" })
  .inputValidator((input: { appId: string; privateKeyPem: string }) => {
    const appId = String(input?.appId ?? "").trim();
    const privateKeyPem = String(input?.privateKeyPem ?? "").trim();
    if (!/^[0-9a-f-]{36}$/i.test(appId)) throw new Error("App id non valido (atteso un UUID).");
    if (!privateKeyPem || privateKeyPem.length > 10000)
      throw new Error("Chiave privata mancante o troppo lunga.");
    // Gli id Enable Banking sono UUID canonici minuscoli: il `kid` del JWT
    // deve coincidere esattamente, quindi si normalizza qui.
    return { appId: appId.toLowerCase(), privateKeyPem };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await saveEbApp(data.appId, data.privateKeyPem);
    return { ok: true };
  });

export const spEbProva = createServerFn({ method: "POST" }).handler(
  async (): Promise<{ nome: string; ambiente: string; redirect: string[] }> => {
    await assertDirettore(await currentUser());
    return ebProvaApplicazione();
  },
);

export const spEbSaldo = createServerFn({ method: "GET" }).handler(
  async (): Promise<EbSaldoInfo | null> => {
    await assertDirettore(await currentUser());
    return ebSaldoAttuale();
  },
);

export const spEbImpostaSaldo = createServerFn({ method: "POST" })
  .inputValidator((input: { saldo: number }) => {
    const saldo = Number(input?.saldo);
    if (!Number.isFinite(saldo) || Math.abs(saldo) > 1e12) throw new Error("Saldo non valido.");
    return { saldo };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await ebImpostaSaldoManuale(data.saldo);
    return { ok: true };
  });

// Innesco ESTERNO del promemoria "manca l'uscita" (fase F): stesso schema del
// sync bancario, token dedicato. Nessuna sessione: solo il token.
export const spCronTurni = createServerFn({ method: "POST" })
  .inputValidator((input: { token: string }) => {
    const token = String(input?.token ?? "").trim();
    if (!token || token.length > 100) throw new Error("Token mancante.");
    return { token };
  })
  .handler(async ({ data }) => promemoriaUsciteAperte(data.token));

// Indirizzo del promemoria turni, visibile a operatore/admin.
export const spCronTurniToken = createServerFn({ method: "GET" }).handler(
  async (): Promise<{ token: string }> => {
    const me = await currentUser();
    assertCap(me.operatore || isAdmin(me));
    return { token: await cronToken("turni") };
  },
);

export const spEbAvviaCollegamento = createServerFn({ method: "POST" }).handler(
  async (): Promise<{ url: string }> => {
    await assertDirettore(await currentUser());
    return ebAvviaCollegamento();
  },
);

export const spEbCompletaCollegamento = createServerFn({ method: "POST" })
  .inputValidator((input: { code: string }) => {
    const code = String(input?.code ?? "").trim();
    if (!code || code.length > 2000) throw new Error("Codice di autorizzazione mancante.");
    return { code };
  })
  .handler(async ({ data }) => {
    await assertDirettore(await currentUser());
    return ebCompletaCollegamento(data.code);
  });

export const spEbScegliConto = createServerFn({ method: "POST" })
  .inputValidator((input: { uid: string; iban: string }) => {
    const uid = String(input?.uid ?? "").trim();
    const iban = String(input?.iban ?? "").trim();
    if (!uid) throw new Error("Conto mancante.");
    return { uid, iban };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    await assertDirettore(await currentUser());
    await ebScegliConto(data.uid, data.iban);
    return { ok: true };
  });

// Passaggio Excel → API: elimina un blocco di movimenti dell'ultimo giorno
// importato per chiamata; il client ripete finché rimanenti > 0.
export const spEbTaglia = createServerFn({ method: "POST" }).handler(
  async (): Promise<{ dataTaglio: string; eliminati: number; rimanenti: number }> => {
    await assertDirettore(await currentUser());
    return ebTagliaUltimoGiorno();
  },
);

// Token della sincronizzazione programmata: visibile SOLO al direttore, da
// incollare nel flusso Power Automate.
export const spEbCronToken = createServerFn({ method: "GET" }).handler(
  async (): Promise<{ token: string }> => {
    await assertDirettore(await currentUser());
    return { token: await ebCronToken() };
  },
);

// Innesco ESTERNO (Power Automate a orari fissi): nessuna sessione, l'unica
// credenziale è il token. Non espone dati: risponde solo con l'esito.
export const spEbCron = createServerFn({ method: "POST" })
  .inputValidator((input: { token: string }) => {
    const token = String(input?.token ?? "").trim();
    if (!token || token.length > 100) throw new Error("Token mancante.");
    return { token };
  })
  .handler(async ({ data }) => {
    return ebSincronizzaProgrammato(data.token);
  });

export const spEbSincronizza = createServerFn({ method: "POST" })
  .inputValidator((input: { importId: string; continuation?: string }) => {
    const importId = String(input?.importId ?? "").slice(0, 60);
    if (!importId.startsWith("SYNC-")) throw new Error("importId non valido");
    const continuation = input?.continuation
      ? String(input.continuation).slice(0, 2000)
      : undefined;
    return { importId, continuation };
  })
  .handler(async ({ data }): Promise<EbSyncResult> => {
    await assertDirettore(await currentUser());
    return ebSincronizza(data.importId, data.continuation);
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

// Annulla l'ultima timbratura di OGGI dell'utente in sessione. La finestra
// dei 5 minuti è verificata dal server sul dato reale.
export const spAnnullaUltimaTimbratura = createServerFn({ method: "POST" }).handler(
  async (): Promise<SpTimbratura> => {
    const me = await currentUser();
    return annullaUltimaTimbratura(me.id);
  },
);

// Resoconto di un giorno per sede: tutti i dipendenti con i loro eventi,
// inclusi quelli senza timbrature (vista correzione operatore).
export const spGetResocontoGiorno = createServerFn({ method: "GET" })
  .inputValidator((input: { sede?: string; data: string }) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input?.data ?? "")) throw new Error("data non valida");
    return { sede: String(input.sede ?? "tutte"), data: input.data };
  })
  .handler(async ({ data }): Promise<import("./sharepoint.server").ResocontoGiornoRiga[]> => {
    const me = await currentUser();
    assertCap(me.operatore || isAdmin(me));
    return resocontoGiorno(data.sede, data.data);
  });

// Timbrature di un dipendente in un giorno (vista correzione operatore).
export const spGetTimbratureGiorno = createServerFn({ method: "GET" })
  .inputValidator((input: { dipendenteId: string; data: string }) => {
    if (!input?.dipendenteId) throw new Error("dipendenteId mancante");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input?.data ?? "")) throw new Error("data non valida");
    return { dipendenteId: String(input.dipendenteId), data: input.data };
  })
  .handler(async ({ data }): Promise<SpTimbratura[]> => {
    const me = await currentUser();
    assertCap(me.operatore || isAdmin(me));
    return fetchTimbratureGiorno(data.dipendenteId, data.data);
  });

// Eliminazione di una timbratura errata (operatore). Il flag Operatore è
// ri-verificato server-side sul record SharePoint.
export const spDeleteTimbratura = createServerFn({ method: "POST" })
  .inputValidator((input: { timbraturaId: string }) => {
    if (!input?.timbraturaId) throw new Error("timbraturaId mancante");
    return { timbraturaId: String(input.timbraturaId) };
  })
  .handler(async ({ data }): Promise<{ ok: true }> => {
    const me = await currentUser();
    if (isAdmin(me)) await deleteTimbratura(data.timbraturaId);
    else await deleteTimbraturaOperatore(me.id, data.timbraturaId);
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

// --- Le mie ore (self-service) ---------------------------------------------
// Ognuno vede SOLO le proprie timbrature: nessuna capability richiesta.
export const spGetMieTimbrature = createServerFn({ method: "GET" })
  .inputValidator((input: { from: string; to: string }) => {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    if (!re.test(input?.from ?? "") || !re.test(input?.to ?? ""))
      throw new Error("Periodo non valido");
    const giorni =
      (new Date(`${input.to}T00:00:00`).getTime() - new Date(`${input.from}T00:00:00`).getTime()) /
      86400000;
    if (giorni < 0 || giorni > 62) throw new Error("Periodo non valido (max 62 giorni)");
    return { from: input.from, to: input.to };
  })
  .handler(async ({ data }): Promise<SpTimbratura[]> => {
    const me = await currentUser();
    // +1 giorno in coda: il turno notturno chiude il giorno dopo.
    const fine = new Date(`${data.to}T00:00:00`);
    fine.setDate(fine.getDate() + 2);
    const tutte = await fetchTimbratureDaISO(new Date(`${data.from}T00:00:00`).toISOString());
    return tutte.filter((t) => t.dipendenteId === me.id && t.dataOra < fine.toISOString());
  });

// --- Correzione timbrature: richiesta dal dipendente, decisa dall'operatore -
export const spGetCorrezioni = createServerFn({ method: "GET" })
  .inputValidator((input?: { tutte?: boolean }) => ({ tutte: Boolean(input?.tutte) }))
  .handler(async ({ data }): Promise<SpCorrezione[]> => {
    const me = await currentUser();
    // La coda completa è per chi ha il flag Operatore (e per l'admin);
    // chiunque altro vede soltanto le proprie richieste.
    const puoVedereTutte = me.operatore || isAdmin(me);
    return fetchCorrezioni(data.tutte && puoVedereTutte ? undefined : me.id);
  });

export const spCreateCorrezione = createServerFn({ method: "POST" })
  .inputValidator((input: { giorno: string; orariProposti: string; motivo: string }) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input?.giorno ?? "")) throw new Error("Giorno non valido");
    const orariProposti = String(input?.orariProposti ?? "").trim();
    const motivo = String(input?.motivo ?? "").trim();
    if (!orariProposti) throw new Error("Indica gli orari corretti.");
    if (!motivo) throw new Error("Indica il motivo della correzione.");
    return {
      giorno: input.giorno,
      orariProposti: orariProposti.slice(0, 500),
      motivo: motivo.slice(0, 500),
    };
  })
  .handler(async ({ data }): Promise<SpCorrezione> => {
    const me = await currentUser();
    if (!parseOrariProposti(data.orariProposti))
      throw new Error('Formato orari non valido. Esempio: "entrata 08:00, uscita 17:00".');
    // Sempre per sé stessi: l'id arriva dalla sessione, mai dal client.
    return createCorrezione({ ...data, dipendenteId: me.id });
  });

export const spDecideCorrezione = createServerFn({ method: "POST" })
  .inputValidator((input: { correzioneId: string; approvata: boolean; note?: string }) => {
    if (!input?.correzioneId) throw new Error("correzioneId mancante");
    return {
      correzioneId: String(input.correzioneId),
      approvata: Boolean(input.approvata),
      note: input.note ? String(input.note).slice(0, 500) : undefined,
    };
  })
  .handler(async ({ data }): Promise<SpCorrezione> => {
    const me = await currentUser();
    assertCap(me.operatore || isAdmin(me));
    return decideCorrezione(
      data.correzioneId,
      data.approvata,
      `${me.nome} ${me.cognome}`.trim(),
      data.note,
    );
  });

// --- Vista di sede del PREPOSTO (sola lettura) ------------------------------
// Il preposto (es. capo appalto) vede turni e ore dei dipendenti della PROPRIA
// sede: nessuna modifica, nessuna approvazione. Il perimetro è imposto qui.
export const spGetSedeGiorno = createServerFn({ method: "GET" })
  .inputValidator((input: { data: string }) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input?.data ?? "")) throw new Error("Data non valida");
    return { data: input.data };
  })
  .handler(async ({ data }): Promise<{ sede: string; righe: ResocontoGiornoRiga[] }> => {
    const me = await currentUser();
    assertCap(Boolean(me.preposto) || me.operatore || isAdmin(me));
    // La sede è quella del record SharePoint, non del client.
    const mio = (await fetchDipendenti()).find((d) => d.id === me.id);
    const sede = isAdmin(me) || me.operatore ? "tutte" : (mio?.sede ?? "");
    if (!sede) throw new Error("Sede non impostata sul tuo record.");
    return { sede, righe: await resocontoGiorno(sede, data.data) };
  });

export const spGetSedeOre = createServerFn({ method: "GET" })
  .inputValidator((input: { from: string; to: string }) => {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    if (!re.test(input?.from ?? "") || !re.test(input?.to ?? ""))
      throw new Error("Periodo non valido");
    return { from: input.from, to: input.to };
  })
  .handler(async ({ data }): Promise<RendicontoRiga[]> => {
    const me = await currentUser();
    assertCap(Boolean(me.preposto) || me.operatore || isAdmin(me));
    const mio = (await fetchDipendenti()).find((d) => d.id === me.id);
    const righe = await computeRendicontoPeriodo(data.from, data.to);
    if (isAdmin(me) || me.operatore) return righe;
    const sede = (mio?.sede ?? "").trim().toLowerCase();
    return righe.filter((r) => r.sede.trim().toLowerCase() === sede);
  });

// Riepilogo della SETTIMANA CORRENTE del dipendente collegato: nessuna
// capability richiesta (ognuno vede solo i propri dati). Alimenta l'avviso
// "settimana sotto le ore previste" e, a seguire, la vista "Le mie ore".
export const spGetMiaSettimana = createServerFn({ method: "GET" })
  .inputValidator((input?: { oggi?: string }) => {
    const re = /^\d{4}-\d{2}-\d{2}$/;
    return { oggi: input?.oggi && re.test(input.oggi) ? input.oggi : undefined };
  })
  .handler(
    async ({
      data,
    }): Promise<{
      lunedi: string;
      domenica: string;
      oreLavorate: number;
      orePreviste: number | null;
      giorniNonChiusi: number;
    }> => {
      const me = await currentUser();
      const oggi = data.oggi ?? new Date().toISOString().slice(0, 10);
      const lunedi = lunediDellaSettimana(oggi);
      const dom = new Date(`${lunedi}T00:00:00`);
      dom.setDate(dom.getDate() + 6);
      const domenica = ymd(dom);
      const righe = await computeRendicontoPeriodo(lunedi, domenica);
      const mia = righe.find((r) => r.dipendenteId === me.id);
      return {
        lunedi,
        domenica,
        oreLavorate: mia?.oreLavorate ?? 0,
        orePreviste: mia?.oreSettimanali ?? null,
        giorniNonChiusi: mia?.giorniNonChiusi ?? 0,
      };
    },
  );

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
