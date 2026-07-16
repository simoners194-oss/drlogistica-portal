// DR Portal — Regole di business per il Modulo Richieste.
// Logica PURA e condivisa client/server (nessuna dipendenza runtime): il client
// la usa per validare i form e abilitare i pulsanti, il server la ri-applica in
// `sharepoint.server.ts` prima di scrivere su SharePoint (mai fidarsi del
// client). Tenendo un solo modulo si evita la divergenza di regole segnalata
// nell'audit.
//
// I valori delle union sono ESATTAMENTE i valori di scelta su SharePoint
// (`TipoRichiesta`, `Stato`, `Modalita`): così scrittura e confronto non
// richiedono mappature.

export type TipoRichiesta =
  | "Ferie"
  | "Permesso"
  | "Straordinario"
  | "Smart Working"
  | "Malattia"
  | "Reperibilità"
  | "Rimborso spese";
export type StatoRichiesta =
  "Bozza" | "Inviata" | "Comunicata" | "Approvata" | "Respinta" | "Annullata";
export type ModalitaStraordinario = "Preventivo" | "Consuntivo";
export type DecisioneRichiesta = "Approvata" | "Respinta";

export const TIPI_RICHIESTA: readonly TipoRichiesta[] = [
  "Ferie",
  "Permesso",
  "Straordinario",
  "Smart Working",
  "Malattia",
  "Reperibilità",
  "Rimborso spese",
];
export const STATI_RICHIESTA: readonly StatoRichiesta[] = [
  "Bozza",
  "Inviata",
  "Comunicata",
  "Approvata",
  "Respinta",
  "Annullata",
];
export const MODALITA: readonly ModalitaStraordinario[] = ["Preventivo", "Consuntivo"];

// Tipologia di acquisto per i rimborsi spese. Le voci sono gestite dalla
// lista SharePoint "Voci" (macro → dettaglio): il tipo è quindi una stringa
// libera. TIPI_ACQUISTO resta come fallback quando la lista Voci è vuota.
export type TipoAcquisto = string;
export const TIPI_ACQUISTO: readonly TipoAcquisto[] = ["Pasto", "Viaggio", "Alloggio", "Altro"];
export function parseTipoAcquisto(v: unknown): TipoAcquisto | null {
  const s = String(v ?? "").trim();
  return s.length > 0 ? s : null;
}
export function isRimborso(tipo: TipoRichiesta): boolean {
  return tipo === "Rimborso spese";
}

// Classificazione dei tipi -----------------------------------------------------
// Tipi misurati in GIORNI (intervallo da–a); gli altri in ORE (fascia oraria).
const TIPI_A_GIORNI: readonly TipoRichiesta[] = ["Ferie", "Smart Working", "Malattia"];
// Tipi che NON passano dall'approvazione: si comunicano e basta.
const TIPI_SENZA_APPROVAZIONE: readonly TipoRichiesta[] = ["Malattia"];
// Tipi con motivazione obbligatoria.
const TIPI_MOTIVAZIONE_OBBLIGATORIA: readonly TipoRichiesta[] = ["Permesso", "Straordinario"];

export function misuraInGiorni(tipo: TipoRichiesta): boolean {
  return TIPI_A_GIORNI.includes(tipo);
}
export function richiedeApprovazione(tipo: TipoRichiesta): boolean {
  return !TIPI_SENZA_APPROVAZIONE.includes(tipo);
}

// Finestra massima per l'invio a posteriori (ore): straordinario a Consuntivo
// e reperibilità (che è sempre a consuntivo, inserita il giorno dopo).
export const CONSUNTIVO_LIMITE_ORE = 72;

// Marcatore scritto in NoteDecisione quando richiedente = approvatore.
export const NOTA_AUTO_APPROVAZIONE = "Auto-approvazione (autorizzatore = richiedente)";

// ---------------------------------------------------------------------------
// Modello di input (ciò che la UI raccoglie e passa al server)
// ---------------------------------------------------------------------------
export interface RichiestaInput {
  tipo: TipoRichiesta;
  dataInizio: string; // "YYYY-MM-DD"
  dataFine: string; // "YYYY-MM-DD"
  oraInizio?: string; // "HH:MM" — tipi a ore
  oraFine?: string; // "HH:MM" — tipi a ore
  motivazione?: string;
  modalita?: ModalitaStraordinario; // solo Straordinario
  protocolloInps?: string; // solo Malattia (facoltativo)
  // Rimborso spese (dataInizio = data acquisto):
  importo?: number;
  tipoAcquisto?: TipoAcquisto;
  giustificativo?: string; // link/URL del documento (upload in fase B2)
}

// ---------------------------------------------------------------------------
// Parser tolleranti (per la lettura da SharePoint)
// ---------------------------------------------------------------------------
export function parseTipo(v: unknown): TipoRichiesta | null {
  const s = String(v ?? "").trim();
  return (TIPI_RICHIESTA as readonly string[]).includes(s) ? (s as TipoRichiesta) : null;
}
export function parseStato(v: unknown): StatoRichiesta | null {
  const s = String(v ?? "").trim();
  return (STATI_RICHIESTA as readonly string[]).includes(s) ? (s as StatoRichiesta) : null;
}
export function parseModalita(v: unknown): ModalitaStraordinario | null {
  const s = String(v ?? "").trim();
  return (MODALITA as readonly string[]).includes(s) ? (s as ModalitaStraordinario) : null;
}

// ---------------------------------------------------------------------------
// Helper data/ora
// ---------------------------------------------------------------------------
const RE_DATE = /^\d{4}-\d{2}-\d{2}$/;
const RE_TIME = /^([01]\d|2[0-3]):[0-5]\d$/;

export function isValidDate(s: string | undefined): boolean {
  if (!s || !RE_DATE.test(s)) return false;
  const d = new Date(`${s}T00:00:00`);
  return !Number.isNaN(d.getTime());
}
export function isValidTime(s: string | undefined): boolean {
  return Boolean(s && RE_TIME.test(s));
}
function timeToMinutes(s: string): number {
  const [h, m] = s.split(":").map(Number);
  return h * 60 + m;
}

// ---------------------------------------------------------------------------
// Calcoli (valore retributivo → coperti da unit test)
// ---------------------------------------------------------------------------

// Giorni: numero di giorni di calendario inclusivi (inizio e fine compresi).
// NB: v1 conta i giorni solari, non i lavorativi; il calcolo del saldo/giorni
// lavorativi è rimandato allo sprint dedicato ferie.
export function computeDurataGiorni(dataInizio: string, dataFine: string): number {
  if (!isValidDate(dataInizio) || !isValidDate(dataFine)) return 0;
  const a = new Date(`${dataInizio}T00:00:00`).getTime();
  const b = new Date(`${dataFine}T00:00:00`).getTime();
  if (b < a) return 0;
  return Math.floor((b - a) / 86400000) + 1;
}

// Ore: differenza oraFine-oraInizio in ore, 2 decimali.
export function computeDurataOre(oraInizio: string, oraFine: string): number {
  if (!isValidTime(oraInizio) || !isValidTime(oraFine)) return 0;
  const diff = timeToMinutes(oraFine) - timeToMinutes(oraInizio);
  if (diff <= 0) return 0;
  return Math.round((diff / 60) * 100) / 100;
}

// Anno di competenza = anno della data di inizio.
export function computeAnnoCompetenza(dataInizio: string): number {
  if (!isValidDate(dataInizio)) return new Date().getFullYear();
  return new Date(`${dataInizio}T00:00:00`).getFullYear();
}

// Momento di fine (data inizio + ora fine) per il calcolo delle 72h. Se manca
// l'ora fine si assume fine giornata.
export function fineEvento(input: RichiestaInput): Date {
  const ora = isValidTime(input.oraFine) ? input.oraFine! : "23:59";
  return new Date(`${input.dataInizio}T${ora}:00`);
}

// Serve la finestra 72h? Straordinario a Consuntivo e Reperibilità (sempre a
// consuntivo, inserita a posteriori).
export function richiedeFinestra72h(input: RichiestaInput): boolean {
  return (
    (input.tipo === "Straordinario" && input.modalita === "Consuntivo") ||
    input.tipo === "Reperibilità"
  );
}

// Inserimento a posteriori fuori finestra: trascorse più di 72h dalla fine.
export function isFuoriFinestra72h(input: RichiestaInput, now: Date = new Date()): boolean {
  if (!richiedeFinestra72h(input)) return false;
  if (!isValidDate(input.dataInizio)) return false;
  const diffOre = (now.getTime() - fineEvento(input).getTime()) / 3600000;
  return diffOre > CONSUNTIVO_LIMITE_ORE;
}

// ---------------------------------------------------------------------------
// Validazione condizionale per tipo (regole applicate dall'app, non da SP)
// ---------------------------------------------------------------------------
export interface ValidationResult {
  ok: boolean;
  errors: string[];
}

export function validateRichiesta(input: RichiestaInput, now: Date = new Date()): ValidationResult {
  const errors: string[] = [];
  const tipo = parseTipo(input.tipo);
  if (!tipo) {
    return { ok: false, errors: ["Tipo di richiesta non valido."] };
  }

  // Rimborso spese: data acquisto (=dataInizio) + importo + tipologia; niente
  // intervallo/ore/modalità.
  if (isRimborso(tipo)) {
    if (!isValidDate(input.dataInizio)) errors.push("Data di acquisto non valida.");
    if (input.importo == null || !(input.importo > 0))
      errors.push("Importo non valido (maggiore di 0).");
    if (!parseTipoAcquisto(input.tipoAcquisto)) errors.push("Tipologia di acquisto obbligatoria.");
    if (input.oraInizio || input.oraFine) errors.push("Il rimborso non prevede orari.");
    if (input.modalita) errors.push("Il rimborso non prevede una modalità.");
    return { ok: errors.length === 0, errors };
  }

  if (!isValidDate(input.dataInizio)) errors.push("Data di inizio non valida.");
  if (!isValidDate(input.dataFine)) errors.push("Data di fine non valida.");
  if (isValidDate(input.dataInizio) && isValidDate(input.dataFine)) {
    if (new Date(`${input.dataFine}T00:00:00`) < new Date(`${input.dataInizio}T00:00:00`)) {
      errors.push("La data di fine non può precedere la data di inizio.");
    }
  }

  if (misuraInGiorni(tipo)) {
    // Ferie / Smart Working / Malattia: intervallo di giorni, niente ore/modalità.
    if (input.oraInizio || input.oraFine)
      errors.push("Questo tipo di richiesta non prevede orari di inizio/fine.");
    if (input.modalita) errors.push("Questo tipo di richiesta non prevede una modalità.");
  } else {
    // Permesso / Straordinario / Reperibilità: giorno singolo, ore obbligatorie.
    if (
      isValidDate(input.dataInizio) &&
      isValidDate(input.dataFine) &&
      input.dataInizio !== input.dataFine
    ) {
      errors.push("Questo tipo di richiesta si riferisce a un solo giorno.");
    }
    if (!isValidTime(input.oraInizio)) errors.push("Ora di inizio non valida (formato HH:MM).");
    if (!isValidTime(input.oraFine)) errors.push("Ora di fine non valida (formato HH:MM).");
    if (
      isValidTime(input.oraInizio) &&
      isValidTime(input.oraFine) &&
      timeToMinutes(input.oraFine!) <= timeToMinutes(input.oraInizio!)
    ) {
      errors.push("L'ora di fine deve essere successiva all'ora di inizio.");
    }
  }

  if (
    TIPI_MOTIVAZIONE_OBBLIGATORIA.includes(tipo) &&
    (!input.motivazione || !input.motivazione.trim())
  )
    errors.push("La motivazione è obbligatoria per questo tipo di richiesta.");

  if (tipo === "Straordinario") {
    if (!parseModalita(input.modalita))
      errors.push("Modalità dello straordinario obbligatoria (Preventivo o Consuntivo).");
  } else if (input.modalita) {
    errors.push("La modalità si applica solo allo straordinario.");
  }

  if (isFuoriFinestra72h(input, now)) {
    errors.push(
      `Inserimento non consentito: sono trascorse più di ${CONSUNTIVO_LIMITE_ORE} ore dal giorno di riferimento.`,
    );
  }

  return { ok: errors.length === 0, errors };
}

// Validazione della decisione (approvazione/rifiuto).
export function validateDecisione(
  decisione: DecisioneRichiesta,
  noteDecisione: string | undefined,
): ValidationResult {
  const errors: string[] = [];
  if (decisione !== "Approvata" && decisione !== "Respinta") errors.push("Decisione non valida.");
  if (decisione === "Respinta" && (!noteDecisione || !noteDecisione.trim()))
    errors.push("La nota è obbligatoria per un rifiuto.");
  return { ok: errors.length === 0, errors };
}

// ---------------------------------------------------------------------------
// Macchina a stati del ciclo di vita
// ---------------------------------------------------------------------------

// Stato in cui entra la richiesta all'invio: Inviata (attende approvazione)
// oppure Comunicata (tipi senza approvazione, es. Malattia).
export function statoDopoInvio(tipo: TipoRichiesta): StatoRichiesta {
  return richiedeApprovazione(tipo) ? "Inviata" : "Comunicata";
}

// Il richiedente può inviare solo da Bozza (o nuova).
export function canSubmit(stato: StatoRichiesta | null): boolean {
  return stato === null || stato === "Bozza";
}
// Il richiedente può annullare finché non c'è una decisione (incluse le
// comunicazioni, es. una malattia inserita per errore).
export function canCancel(stato: StatoRichiesta | null): boolean {
  return stato === "Bozza" || stato === "Inviata" || stato === "Comunicata";
}
// L'approvatore può decidere solo su richieste Inviate.
export function canDecide(stato: StatoRichiesta | null): boolean {
  return stato === "Inviata";
}

// ---------------------------------------------------------------------------
// Routing approvazione
// ---------------------------------------------------------------------------

// Auto-approvazione: il richiedente è anche l'autorizzatore (Autorizza=true che
// invia una propria richiesta soggetta ad approvazione — oggi Francesco).
export function isAutoApprovazione(richiedenteId: string, richiedenteAutorizza: boolean): boolean {
  return richiedenteAutorizza === true && Boolean(richiedenteId);
}

// ---------------------------------------------------------------------------
// Routing supervisione per sede
// ---------------------------------------------------------------------------
// Le due sedi storiche (Fiano Romano e San Giuliano) fanno capo al supervisore
// DR005. Ogni ALTRA sede — comprese quelle caricate in futuro — fa capo a DR000.
// Elenco CONGELATO di proposito: NON aggiungere qui le nuove sedi, così ricadono
// automaticamente su DR000. Confronto tollerante a nome/ID e maiuscole.
const SEDI_DR005 = new Set(["fiano romano", "san giuliano", "roma", "san-giuliano"]);

export function codiceSupervisoreDiSede(sedeRichiedente: string): "DR005" | "DR000" {
  const s = (sedeRichiedente ?? "").trim().toLowerCase();
  return SEDI_DR005.has(s) ? "DR005" : "DR000";
}

// DR005 è il supervisore "globale/onnisciente": in VISUALIZZAZIONE vede tutte le
// sedi (come l'admin). Può però AUTORIZZARE solo le sedi storiche
// (vedi supervisionaSede). DR000 e gli altri vedono solo le proprie sedi.
export function isSupervisoreGlobale(codice: string): boolean {
  return (codice ?? "").trim().toUpperCase() === "DR005";
}

// Un autorizzatore (identificato dal suo Codice) è competente sulla richiesta di
// quella sede? Confronto case-insensitive sul codice.
export function supervisionaSede(codiceAutorizzatore: string, sedeRichiedente: string): boolean {
  return (
    (codiceAutorizzatore ?? "").trim().toUpperCase() === codiceSupervisoreDiSede(sedeRichiedente)
  );
}

// Sede "storica" (Fiano Romano / San Giuliano)? Usato dal Procurement: le
// richieste di acquisto sono attive solo per le risorse delle sedi storiche.
export function isSedeStorica(sede: string): boolean {
  return SEDI_DR005.has((sede ?? "").trim().toLowerCase());
}

// Costruisce il Title leggibile: REQ-<anno>-<idNativo>.
export function formatTitle(anno: number, itemId: string | number): string {
  return `REQ-${anno}-${itemId}`;
}
