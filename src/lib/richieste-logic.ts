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

export type TipoRichiesta = "Ferie" | "Permesso" | "Straordinario";
export type StatoRichiesta = "Bozza" | "Inviata" | "Approvata" | "Respinta" | "Annullata";
export type ModalitaStraordinario = "Preventivo" | "Consuntivo";
export type DecisioneRichiesta = "Approvata" | "Respinta";

export const TIPI_RICHIESTA: readonly TipoRichiesta[] = ["Ferie", "Permesso", "Straordinario"];
export const STATI_RICHIESTA: readonly StatoRichiesta[] = [
  "Bozza",
  "Inviata",
  "Approvata",
  "Respinta",
  "Annullata",
];
export const MODALITA: readonly ModalitaStraordinario[] = ["Preventivo", "Consuntivo"];

// Finestra massima per l'invio di uno straordinario a Consuntivo (ore).
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
  oraInizio?: string; // "HH:MM" — Permesso/Straordinario
  oraFine?: string; // "HH:MM" — Permesso/Straordinario
  motivazione?: string;
  modalita?: ModalitaStraordinario; // solo Straordinario
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

// Giorni di ferie: numero di giorni di calendario inclusivi (inizio e fine
// compresi). NB: v1 conta i giorni solari, non i giorni lavorativi; il calcolo
// del saldo/giorni lavorativi è rimandato allo sprint dedicato ferie.
export function computeDurataGiorni(dataInizio: string, dataFine: string): number {
  if (!isValidDate(dataInizio) || !isValidDate(dataFine)) return 0;
  const a = new Date(`${dataInizio}T00:00:00`).getTime();
  const b = new Date(`${dataFine}T00:00:00`).getTime();
  if (b < a) return 0;
  return Math.floor((b - a) / 86400000) + 1;
}

// Ore di permesso/straordinario: differenza oraFine-oraInizio in ore, 2 decimali.
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

// Momento di fine dello straordinario (data inizio + ora fine), per il calcolo
// delle 72h. Se manca l'ora fine si assume fine giornata.
export function fineStraordinario(input: RichiestaInput): Date {
  const ora = isValidTime(input.oraFine) ? input.oraFine! : "23:59";
  return new Date(`${input.dataInizio}T${ora}:00`);
}

// Straordinario a Consuntivo scaduto: sono trascorse più di 72h dalla fine.
export function isConsuntivoScaduto(input: RichiestaInput, now: Date = new Date()): boolean {
  if (input.tipo !== "Straordinario" || input.modalita !== "Consuntivo") return false;
  if (!isValidDate(input.dataInizio)) return false;
  const fine = fineStraordinario(input);
  const diffOre = (now.getTime() - fine.getTime()) / 3600000;
  return diffOre > CONSUNTIVO_LIMITE_ORE;
}

// ---------------------------------------------------------------------------
// Validazione condizionale per tipo (regole applicate dall'app, non da SP)
// ---------------------------------------------------------------------------
export interface ValidationResult {
  ok: boolean;
  errors: string[];
}

export function validateRichiesta(
  input: RichiestaInput,
  now: Date = new Date(),
): ValidationResult {
  const errors: string[] = [];
  const tipo = parseTipo(input.tipo);
  if (!tipo) {
    return { ok: false, errors: ["Tipo di richiesta non valido."] };
  }

  if (!isValidDate(input.dataInizio)) errors.push("Data di inizio non valida.");
  if (!isValidDate(input.dataFine)) errors.push("Data di fine non valida.");
  if (isValidDate(input.dataInizio) && isValidDate(input.dataFine)) {
    if (new Date(`${input.dataFine}T00:00:00`) < new Date(`${input.dataInizio}T00:00:00`)) {
      errors.push("La data di fine non può precedere la data di inizio.");
    }
  }

  if (tipo === "Ferie") {
    // Motivazione opzionale; nessuna ora; nessuna modalità.
    if (input.oraInizio || input.oraFine)
      errors.push("Le ferie non prevedono orari di inizio/fine.");
    if (input.modalita) errors.push("Le ferie non prevedono una modalità.");
  } else {
    // Permesso / Straordinario: giorno singolo, ore obbligatorie, motivazione.
    if (
      isValidDate(input.dataInizio) &&
      isValidDate(input.dataFine) &&
      input.dataInizio !== input.dataFine
    ) {
      errors.push("Per permesso e straordinario la data di fine coincide con quella di inizio.");
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
    if (!input.motivazione || !input.motivazione.trim())
      errors.push("La motivazione è obbligatoria per permessi e straordinari.");
  }

  if (tipo === "Straordinario") {
    if (!parseModalita(input.modalita))
      errors.push("Modalità dello straordinario obbligatoria (Preventivo o Consuntivo).");
    if (isConsuntivoScaduto(input, now)) {
      errors.push(
        `Consuntivo non inviabile: sono trascorse più di ${CONSUNTIVO_LIMITE_ORE} ore dallo straordinario.`,
      );
    }
  } else if (input.modalita) {
    errors.push("La modalità si applica solo allo straordinario.");
  }

  return { ok: errors.length === 0, errors };
}

// Validazione della decisione (approvazione/rifiuto).
export function validateDecisione(
  decisione: DecisioneRichiesta,
  noteDecisione: string | undefined,
): ValidationResult {
  const errors: string[] = [];
  if (decisione !== "Approvata" && decisione !== "Respinta")
    errors.push("Decisione non valida.");
  if (decisione === "Respinta" && (!noteDecisione || !noteDecisione.trim()))
    errors.push("La nota è obbligatoria per un rifiuto.");
  return { ok: errors.length === 0, errors };
}

// ---------------------------------------------------------------------------
// Macchina a stati del ciclo di vita
// ---------------------------------------------------------------------------

// Il richiedente può inviare una richiesta solo da Bozza (o nuova).
export function canSubmit(stato: StatoRichiesta | null): boolean {
  return stato === null || stato === "Bozza";
}
// Il richiedente può annullare finché non c'è una decisione.
export function canCancel(stato: StatoRichiesta | null): boolean {
  return stato === "Bozza" || stato === "Inviata";
}
// L'approvatore può decidere solo su richieste Inviate.
export function canDecide(stato: StatoRichiesta | null): boolean {
  return stato === "Inviata";
}

// ---------------------------------------------------------------------------
// Routing approvazione
// ---------------------------------------------------------------------------

// Auto-approvazione: il richiedente è anche l'autorizzatore (stesso id).
// Vale per chi ha Autorizza=true e invia una propria richiesta (oggi Francesco).
export function isAutoApprovazione(richiedenteId: string, richiedenteAutorizza: boolean): boolean {
  return richiedenteAutorizza === true && Boolean(richiedenteId);
}

// Costruisce il Title leggibile: REQ-<anno>-<idNativo>.
export function formatTitle(anno: number, itemId: string | number): string {
  return `REQ-${anno}-${itemId}`;
}
