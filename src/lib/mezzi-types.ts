// DR Portal — modulo Mezzi: tipi di dominio e logica pura (client-safe).
// -----------------------------------------------------------------------------
// Il parco mezzi vive in un unico snapshot JSON (`MezziData/mezzi-db.json`
// nella libreria del sito SharePoint) letto/scritto da mezzi.server.ts.
// Qui NIENTE import server: solo tipi e funzioni pure riusate da UI e server.
// -----------------------------------------------------------------------------

export type MezzoGruppo = "operativo" | "altro";
export type MezzoProprieta = "noleggio" | "proprieta";
export type MezzoStato =
  | "operativo"
  | "manutenzione"
  | "sostitutivo"
  | "fermo"
  | "reso"
  | "dirigenziale"
  | "dismesso"
  | "venduto"
  | "rottamato";

export const MEZZO_STATI: MezzoStato[] = [
  "operativo",
  "manutenzione",
  "sostitutivo",
  "fermo",
  "reso",
  "dirigenziale",
  "dismesso",
  "venduto",
  "rottamato",
];

// Stati che collocano il mezzo nel gruppo "operativo" di default.
export function gruppoDaStato(stato: MezzoStato): MezzoGruppo {
  return stato === "operativo" || stato === "manutenzione" || stato === "sostitutivo"
    ? "operativo"
    : "altro";
}

/** Dati tecnici dalla carta di circolazione (estratti dai PDF o inseriti a mano). */
export interface DatiLibretto {
  telaio?: string;
  immatricolazione?: string; // ISO YYYY-MM-DD (campo B)
  intestatario?: string; // C.2.1 (società di noleggio / proprietà)
  massaMaxKg?: number; // F.2
  massaVuotoKg?: number;
  portataKg?: number;
  cilindrata?: number; // P.1
  potenzaKw?: number; // P.2
  alimentazione?: string; // P.3
  classeEuro?: string;
  lunghezzaM?: number;
  larghezzaM?: number;
  carrozzeria?: string; // J.2 (es. "F3 furgone isotermico con gruppo frigorifero")
  allestimento?: string; // Lamberet, Lecapitaine, Icy Truck…
  posti?: number;
}

export interface Mezzo {
  id: string; // targa normalizzata (maiuscola, senza spazi)
  targa: string;
  gruppo: MezzoGruppo;
  stato: MezzoStato;
  proprieta: MezzoProprieta;
  societa?: string; // intestazione operativa (DR Logistica / DR Soluzione Logistica)
  tipoMezzo?: string; // FURGONE, FURGONE COIBENTATO FRIGO, CASSONATO CON SPONDA…
  marca?: string;
  modello?: string;
  frigo?: boolean;
  atp?: boolean;
  sponda?: boolean;
  appalto?: string; // appalto corrente (denormalizzato dall'affidamento attivo)
  sede?: string;
  kmAttuali?: number;
  kmAggiornatiAl?: string; // ISO date
  /** Numero dispositivo Telepass installato. */
  telepass?: string;
  /** Mezzo titolare di cui questo è il sostitutivo (targa). */
  sostitutivoDi?: string;
  libretto?: DatiLibretto;
  /** Link (SharePoint/OneDrive) ai documenti del mezzo. */
  docLibretto?: string;
  docAssicurazione?: string;
  docContratto?: string;
  note?: string;
}

export interface Affidamento {
  id: string;
  mezzoId: string;
  autistaNome: string;
  autistaCodice?: string; // codice dipendente DR se noto
  appalto?: string;
  dal: string; // ISO date
  al?: string; // ISO date — assente = affidamento in corso
  verbaleConsegnaUrl?: string;
  verbaleRiconsegnaUrl?: string;
  note?: string;
}

export type ScadenzaTipo =
  | "revisione"
  | "assicurazione"
  | "bollo"
  | "atp"
  | "tagliando"
  | "fine_contratto"
  | "disdetta"
  | "ztl"
  | "patente"
  | "altro";

export const SCADENZA_TIPI: ScadenzaTipo[] = [
  "revisione",
  "assicurazione",
  "bollo",
  "atp",
  "tagliando",
  "fine_contratto",
  "disdetta",
  "ztl",
  "patente",
  "altro",
];

export interface Scadenza {
  id: string;
  mezzoId?: string; // assente per scadenze non legate a un mezzo (es. patente)
  tipo: ScadenzaTipo;
  descrizione?: string;
  scadenza: string; // ISO date
  /** Alert email già inviati: soglia giorni → data invio ISO. Si azzera quando
   *  la data di scadenza viene aggiornata (rinnovo). */
  alertInviati?: Record<string, string>;
  chiusa?: boolean; // gestita/rinnovata: esce dal semaforo
  note?: string;
}

export interface Contratto {
  id: string;
  mezzoId: string;
  noleggiatore?: string;
  canoneMensileEur?: number;
  canonePiuIva?: boolean;
  dataInizio?: string;
  dataFine?: string;
  kmMeseInclusi?: number; // assente/0 = illimitato
  extraKmEur?: number;
  preavvisoDisdettaGiorni?: number;
  franchigiaRca?: string;
  franchigiaKasko?: string;
  franchigiaFurto?: string;
  serviziInclusi?: string[];
  penali?: string;
  deposito?: number;
  attivo?: boolean;
  note?: string;
}

export type MultaStato =
  | "ricevuta"
  | "contestata_dipendente"
  | "risposta_ricevuta"
  | "in_detrazione"
  | "detratta"
  | "pagata"
  | "ricorso"
  | "annullata";

export const MULTA_STATI: MultaStato[] = [
  "ricevuta",
  "contestata_dipendente",
  "risposta_ricevuta",
  "in_detrazione",
  "detratta",
  "pagata",
  "ricorso",
  "annullata",
];

export type MultaResponsabilita = "da_definire" | "autista" | "ufficio";

export interface MultaEvento {
  data: string; // ISO datetime
  stato: MultaStato;
  nota?: string;
  utente?: string;
}

export interface Multa {
  id: string;
  mezzoId: string;
  dataInfrazione: string; // ISO datetime (per il match autista)
  dataNotifica?: string; // ISO date
  comune?: string;
  tipo?: string; // ZTL, velocità, sosta, semaforo, altro
  importoEur?: number;
  importoRidottoEur?: number;
  scadenzaPagamento?: string;
  autistaNome?: string; // proposto dallo storico affidamenti, modificabile
  autistaCodice?: string;
  responsabilita: MultaResponsabilita;
  stato: MultaStato;
  storico?: MultaEvento[];
  verbaleUrl?: string;
  note?: string;
}

export type ZtlStato = "richiesto" | "attivo" | "scaduto" | "respinto";

export interface PermessoZtl {
  id: string;
  mezzoId?: string;
  comune: string;
  stato: ZtlStato;
  dataRichiesta?: string;
  protocollo?: string;
  dal?: string;
  al?: string;
  note?: string;
}

export interface InterventoOfficina {
  id: string;
  mezzoId: string;
  data: string; // ISO date
  tipo?: string; // tagliando, riparazione, gomme, carrozzeria, frigo, revisione…
  descrizione?: string;
  fornitore?: string;
  costoEur?: number;
  km?: number;
  note?: string;
}

export interface LetturaKm {
  id: string;
  mezzoId: string;
  data: string; // ISO date
  km: number;
  fonte?: string; // autista, foto, officina, telepass…
  note?: string;
}

export interface Rifornimento {
  id: string;
  mezzoId: string;
  data: string; // ISO date
  litri?: number;
  importoEur?: number;
  carta?: string;
  km?: number;
  note?: string;
}

export interface ParametriMezzi {
  /** Destinatari email degli alert scadenze. */
  emailAlert: string[];
  /** Soglie giorni per gli alert (decrescenti). */
  soglieAlertGiorni: number[];
  /** Costo extra-km di default quando il contratto non lo indica. */
  extraKmDefaultEur: number;
  /** Scostamento % dal canone mediano di categoria oltre cui un contratto è "fuori mercato". */
  sogliaOutlierPct: number;
  /** URL base della cartella "Mezzi DR" su SharePoint (per costruire i link ai documenti). */
  urlBaseDocumenti?: string;
  /** Causali di mancata consegna (Fase B). */
  causaliNonConsegna: string[];
  /** Parametri stima costi viaggio. */
  costoOrarioAutistaEur: number;
  costoCarburanteEurLitro: number;
  consumoMedioKmLitro: number;
}

export const PARAMETRI_DEFAULT: ParametriMezzi = {
  emailAlert: ["d.gabelli@drlogistica.it", "f.r@drlogistica.it"],
  soglieAlertGiorni: [90, 60, 30],
  extraKmDefaultEur: 0.15,
  sogliaOutlierPct: 20,
  causaliNonConsegna: [
    "Destinatario assente",
    "Indirizzo errato",
    "Rifiuto",
    "Problemi contrassegno",
    "Causa di forza maggiore",
    "Non ho fatto in tempo",
  ],
  costoOrarioAutistaEur: 18,
  costoCarburanteEurLitro: 1.75,
  consumoMedioKmLitro: 9,
};

export interface MezziDb {
  versione: number; // rev incrementale dello snapshot (concorrenza ottimistica)
  mezzi: Mezzo[];
  affidamenti: Affidamento[];
  scadenze: Scadenza[];
  contratti: Contratto[];
  multe: Multa[];
  ztl: PermessoZtl[];
  officina: InterventoOfficina[];
  km: LetturaKm[];
  carburante: Rifornimento[];
  parametri: ParametriMezzi;
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export function emptyMezziDb(): MezziDb {
  return {
    versione: 0,
    mezzi: [],
    affidamenti: [],
    scadenze: [],
    contratti: [],
    multe: [],
    ztl: [],
    officina: [],
    km: [],
    carburante: [],
    parametri: { ...PARAMETRI_DEFAULT },
  };
}

// ---------------------------------------------------------------------------
// Logica pura
// ---------------------------------------------------------------------------

export function normalizzaTarga(t: string): string {
  return (t ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

/** Giorni (interi) da oggi alla data ISO; negativi se già passata. */
export function giorniA(iso: string, oggi = new Date()): number {
  const d = new Date(iso.slice(0, 10) + "T12:00:00");
  const base = new Date(oggi.getFullYear(), oggi.getMonth(), oggi.getDate(), 12);
  return Math.round((d.getTime() - base.getTime()) / 86400000);
}

export type Semaforo = "scaduta" | "rosso" | "arancio" | "giallo" | "ok";

export function semaforoScadenza(giorni: number): Semaforo {
  if (giorni < 0) return "scaduta";
  if (giorni <= 30) return "rosso";
  if (giorni <= 60) return "arancio";
  if (giorni <= 90) return "giallo";
  return "ok";
}

/** Chi aveva il mezzo a una certa data/ora, dallo storico affidamenti. */
export function autistaAllaData(
  affidamenti: Affidamento[],
  mezzoId: string,
  dataOraIso: string,
): Affidamento | null {
  const t = dataOraIso.slice(0, 10);
  const candidati = affidamenti
    .filter(
      (a) => a.mezzoId === mezzoId && a.dal.slice(0, 10) <= t && (!a.al || a.al.slice(0, 10) >= t),
    )
    .sort((a, b) => (a.dal < b.dal ? 1 : -1));
  return candidati[0] ?? null;
}

/** Affidamento attualmente aperto per un mezzo. */
export function affidamentoCorrente(
  affidamenti: Affidamento[],
  mezzoId: string,
): Affidamento | null {
  const aperti = affidamenti
    .filter((a) => a.mezzoId === mezzoId && !a.al)
    .sort((a, b) => (a.dal < b.dal ? 1 : -1));
  return aperti[0] ?? null;
}

export interface OutlierCanone {
  contratto: Contratto;
  mezzo: Mezzo;
  categoria: string;
  medianaEur: number;
  scostamentoPct: number;
}

/** Categoria omogenea per il confronto canoni: tipo mezzo normalizzato. */
export function categoriaCanone(m: Mezzo): string {
  const t = (m.tipoMezzo ?? "").toUpperCase().trim();
  if (t.includes("FRIGO") || m.frigo) return "FURGONE FRIGO";
  if (t.includes("SPONDA")) return "FURGONE/CASSONATO CON SPONDA";
  if (t.includes("CASSONATO")) return "CASSONATO";
  if (t.includes("AUTO")) return "AUTO";
  return t || "ALTRO";
}

function mediana(nums: number[]): number {
  const s = [...nums].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** Contratti con canone oltre soglia% sopra la mediana della loro categoria. */
export function canoniFuoriMercato(
  mezzi: Mezzo[],
  contratti: Contratto[],
  sogliaPct: number,
): OutlierCanone[] {
  const byId = new Map(mezzi.map((m) => [m.id, m]));
  const gruppi = new Map<string, { c: Contratto; m: Mezzo }[]>();
  for (const c of contratti) {
    if (c.attivo === false || !c.canoneMensileEur) continue;
    const m = byId.get(c.mezzoId);
    if (!m) continue;
    const cat = categoriaCanone(m);
    if (!gruppi.has(cat)) gruppi.set(cat, []);
    gruppi.get(cat)!.push({ c, m });
  }
  const out: OutlierCanone[] = [];
  for (const [cat, righe] of gruppi) {
    if (righe.length < 3) continue; // confronto poco significativo
    const med = mediana(righe.map((r) => r.c.canoneMensileEur!));
    if (med <= 0) continue;
    for (const r of righe) {
      const pct = ((r.c.canoneMensileEur! - med) / med) * 100;
      if (pct >= sogliaPct) {
        out.push({
          contratto: r.c,
          mezzo: r.m,
          categoria: cat,
          medianaEur: med,
          scostamentoPct: pct,
        });
      }
    }
  }
  return out.sort((a, b) => b.scostamentoPct - a.scostamentoPct);
}

export interface SforamentoKm {
  mezzo: Mezzo;
  contratto: Contratto;
  kmMedioMese: number;
  kmInclusi: number;
  extraKmMese: number;
  costoExtraMeseEur: number;
}

/** Km medi/mese dalle letture (ultime 2+) e proiezione costo extra-km. */
export function sforamentiKm(
  mezzi: Mezzo[],
  contratti: Contratto[],
  letture: LetturaKm[],
  extraKmDefault: number,
): SforamentoKm[] {
  const byId = new Map(mezzi.map((m) => [m.id, m]));
  const out: SforamentoKm[] = [];
  for (const c of contratti) {
    if (c.attivo === false || !c.kmMeseInclusi) continue;
    const m = byId.get(c.mezzoId);
    if (!m) continue;
    const ls = letture
      .filter((l) => l.mezzoId === c.mezzoId)
      .sort((a, b) => (a.data < b.data ? -1 : 1));
    if (ls.length < 2) continue;
    const prima = ls[0];
    const ultima = ls[ls.length - 1];
    const giorni = Math.max(1, giorniA(ultima.data, new Date(prima.data)));
    const kmMedioMese = ((ultima.km - prima.km) / giorni) * 30.44;
    const extraKmMese = kmMedioMese - c.kmMeseInclusi;
    if (extraKmMese <= 0) continue;
    const tariffa = c.extraKmEur ?? extraKmDefault;
    out.push({
      mezzo: m,
      contratto: c,
      kmMedioMese: Math.round(kmMedioMese),
      kmInclusi: c.kmMeseInclusi,
      extraKmMese: Math.round(extraKmMese),
      costoExtraMeseEur: Math.round(extraKmMese * tariffa * 100) / 100,
    });
  }
  return out.sort((a, b) => b.costoExtraMeseEur - a.costoExtraMeseEur);
}

/** Id compatto per nuove entità (timestamp + suffisso casuale). */
export function nuovoId(prefix: string): string {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}
