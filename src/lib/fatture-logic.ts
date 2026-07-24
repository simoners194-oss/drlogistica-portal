// DR Portal — Fatture emesse (sezione Finanza, direttore DR005): logica pura.
// -----------------------------------------------------------------------------
// Sorgente dati v1: l'export "Check fatture inviate" del pannello Aruba
// (xlsx). In futuro la STESSA lista sarà alimentata dalle API Aruba v2 in
// sola lettura — chiave univoca in entrambi i casi = NOME FILE SdI
// (es. IT01879020517A2026_xxxxx.xml.p7m), mai il numero fattura.
// Qui vivono: parsing dell'export, calcolo scadenza/ritardo (termini di
// pagamento per cliente), stato incasso (Pagata/Parziale/Non incassata) e la
// proposta di abbinamento automatico fatture ↔ incassi bancari.

import {
  canonicalCliente,
  clienteGroupKey,
  cellToIsoDate,
  cellToImporto,
  normalizeTesto,
} from "./finanza-logic";

// --- Modello -----------------------------------------------------------------

export interface FatturaRaw {
  nomeFile: string; // chiave univoca (Title su SharePoint)
  numero: string; // es. "FPR 201/26" — NON univoco
  idSdi: string;
  dataInvio: string; // YYYY-MM-DD
  dataDocumento: string; // YYYY-MM-DD
  tipoDocumento: string; // "Fattura - TD01" | "Nota di credito - TD04" | ...
  cliente: string;
  piva: string;
  metodoPagamento: string;
  imponibile: number;
  iva: number;
  totale: number;
  netto: number;
  statoSdI: string; // Consegnata / Scartata / ...
}

export interface TerminePagamento {
  cliente: string;
  giorni: number;
  descrizione?: string;
}

export interface AbbinamentoIncasso {
  id?: string;
  fatturaFile: string; // nomeFile della fattura
  movimentoChiave: string; // chiave del movimento bancario (Title)
  importo: number;
  origine: "Auto" | "Manuale";
}

/** Giorni di pagamento di default quando il cliente non è nei termini. */
export const TERMINI_DEFAULT_GIORNI = 30;

/** Tolleranza (€) sotto la quale un residuo si considera saldato. */
export const TOLLERANZA_SALDO = 1;

export function isNotaCredito(tipoDocumento: string): boolean {
  return /td04|nota di credito/i.test(tipoDocumento);
}

/** Fattura che NON concorre al credito: scartata/rifiutata dallo SdI (i
 *  reinvii hanno un nuovo nome file, quindi restano in archivio come storia). */
export function isEsclusaDalCredito(f: Pick<FatturaRaw, "statoSdI">): boolean {
  return /scartat|rifiutat/i.test(f.statoSdI);
}

// --- Termini di pagamento ----------------------------------------------------

/** Giorni di pagamento per un cliente (match per chiave canonica; la riga
 *  senza descrizione è quella generica). */
export function giorniPerCliente(cliente: string, termini: readonly TerminePagamento[]): number {
  const key = clienteGroupKey(cliente);
  const match = termini.filter((t) => clienteGroupKey(t.cliente) === key && t.giorni > 0);
  if (!match.length) return TERMINI_DEFAULT_GIORNI;
  const generico = match.find((t) => !t.descrizione?.trim());
  return (generico ?? match[0]).giorni;
}

export function scadenzaFattura(dataDocumento: string, giorni: number): string {
  const d = new Date(`${dataDocumento}T00:00:00`);
  d.setDate(d.getDate() + giorni);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const g = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${g}`;
}

// --- Stato incasso -----------------------------------------------------------

export type StatoIncasso = "Pagata" | "Parziale" | "Non incassata" | "NC";

export interface FatturaStato {
  incassato: number;
  residuo: number;
  stato: StatoIncasso;
  scadenza: string;
  inRitardo: boolean;
  giorniRitardo: number;
}

export function computeStatoFattura(
  f: FatturaRaw,
  incassato: number,
  termini: readonly TerminePagamento[],
  oggiISO: string,
): FatturaStato {
  const giorni = giorniPerCliente(f.cliente, termini);
  const scadenza = scadenzaFattura(f.dataDocumento, giorni);
  // Le note di credito non si "incassano" e le scartate/rifiutate dallo SdI
  // non sono crediti: entrambe fuori dal computo di residui e ritardi.
  if (isNotaCredito(f.tipoDocumento) || isEsclusaDalCredito(f) || f.totale <= 0) {
    return { incassato, residuo: 0, stato: "NC", scadenza, inRitardo: false, giorniRitardo: 0 };
  }
  const residuo = Math.max(0, f.totale - incassato);
  const stato: StatoIncasso =
    residuo <= TOLLERANZA_SALDO
      ? "Pagata"
      : incassato > TOLLERANZA_SALDO
        ? "Parziale"
        : "Non incassata";
  const inRitardo = stato !== "Pagata" && oggiISO > scadenza;
  const giorniRitardo = inRitardo
    ? Math.floor(
        (new Date(`${oggiISO}T00:00:00`).getTime() - new Date(`${scadenza}T00:00:00`).getTime()) /
          86400000,
      )
    : 0;
  return { incassato, residuo, stato, scadenza, inRitardo, giorniRitardo };
}

// --- Riconciliazione automatica ---------------------------------------------
// Un bonifico può pagare PIÙ fatture ("saldo ft n. 170-171-172-173") e una
// fattura può essere pagata da più bonifici (acconti): il modello è n:n, con
// un importo allocato per coppia. L'automatismo è CONSERVATIVO: aggancia solo
// incassi dello stesso cliente che citano il numero della fattura, oppure che
// coincidono al centesimo con il residuo. Il resto si abbina a mano.

export interface MovimentoPerRiconciliazione {
  chiave: string;
  dataContabile: string;
  importo: number;
  tipologia: string;
  cliente: string;
  descrizione: string;
  nrFattura: string;
}

/** Estrae numero e anno dal numero documento ("FPR 201/26" → {n:"201", anno:"26"}). */
export function parseNumeroFattura(numero: string): { n: string; anno: string } | null {
  const m = normalizeTesto(numero).match(/(\d+)\s*\/\s*(\d{2,4})/);
  return m ? { n: m[1], anno: m[2].slice(-2) } : null;
}

/** true se il testo del movimento cita il numero della fattura (con confini
 *  numerici, per non confondere 1/26 con 171/26). */
export function movimentoCitaFattura(mov: MovimentoPerRiconciliazione, numero: string): boolean {
  const num = parseNumeroFattura(numero);
  if (!num) return false;
  const testo = `${mov.nrFattura} ${mov.descrizione}`.toLowerCase();
  const conAnno = new RegExp(`(^|[^0-9])${num.n}\\s*/\\s*${num.anno}([^0-9]|$)`);
  if (conAnno.test(testo)) return true;
  const solo = new RegExp(`(^|[^0-9])${num.n}([^0-9]|$)`);
  return solo.test(mov.nrFattura.toLowerCase());
}

export interface PropostaAbbinamento extends AbbinamentoIncasso {
  motivo: "numero" | "importo";
}

/** Propone gli abbinamenti automatici. Considera SOLO incassi (importo>0,
 *  tipologia Incasso) e fatture TD01 aperte; rispetta gli abbinamenti già
 *  registrati (residui per fattura E per movimento) ed è deterministica. */
export function proponiAbbinamenti(
  fatture: readonly FatturaRaw[],
  movimenti: readonly MovimentoPerRiconciliazione[],
  esistenti: readonly AbbinamentoIncasso[],
): PropostaAbbinamento[] {
  const round = (n: number) => Math.round(n * 100) / 100;
  const incassatoPerFattura = new Map<string, number>();
  const allocatoPerMovimento = new Map<string, number>();
  const coppie = new Set<string>();
  for (const a of esistenti) {
    incassatoPerFattura.set(
      a.fatturaFile,
      (incassatoPerFattura.get(a.fatturaFile) ?? 0) + a.importo,
    );
    allocatoPerMovimento.set(
      a.movimentoChiave,
      (allocatoPerMovimento.get(a.movimentoChiave) ?? 0) + a.importo,
    );
    coppie.add(`${a.fatturaFile}|${a.movimentoChiave}`);
  }

  const aperte = fatture
    .filter((f) => !isNotaCredito(f.tipoDocumento) && !isEsclusaDalCredito(f) && f.totale > 0)
    .map((f) => ({
      f,
      key: clienteGroupKey(f.cliente),
      residuo: round(f.totale - (incassatoPerFattura.get(f.nomeFile) ?? 0)),
    }))
    .filter((x) => x.residuo > TOLLERANZA_SALDO)
    // Più vecchie prima: un bonifico cumulativo salda in ordine cronologico.
    .sort((a, b) => a.f.dataDocumento.localeCompare(b.f.dataDocumento));

  const incassi = movimenti
    .filter((m) => m.importo > 0 && m.tipologia === "Incasso" && m.cliente)
    .map((m) => ({
      m,
      key: clienteGroupKey(m.cliente),
      residuo: round(m.importo - (allocatoPerMovimento.get(m.chiave) ?? 0)),
    }))
    .filter((x) => x.residuo > 0.01)
    .sort((a, b) => a.m.dataContabile.localeCompare(b.m.dataContabile));

  const proposte: PropostaAbbinamento[] = [];
  const alloca = (
    fat: (typeof aperte)[number],
    inc: (typeof incassi)[number],
    motivo: "numero" | "importo",
  ) => {
    const importo = round(Math.min(fat.residuo, inc.residuo));
    if (importo <= 0.01) return;
    proposte.push({
      fatturaFile: fat.f.nomeFile,
      movimentoChiave: inc.m.chiave,
      importo,
      origine: "Auto",
      motivo,
    });
    fat.residuo = round(fat.residuo - importo);
    inc.residuo = round(inc.residuo - importo);
    coppie.add(`${fat.f.nomeFile}|${inc.m.chiave}`);
  };

  // Passata 1 — il movimento cita il numero della fattura (stesso cliente).
  for (const inc of incassi) {
    if (inc.residuo <= 0.01) continue;
    for (const fat of aperte) {
      if (inc.residuo <= 0.01) break;
      if (fat.residuo <= TOLLERANZA_SALDO) continue;
      if (fat.key !== inc.key) continue;
      if (coppie.has(`${fat.f.nomeFile}|${inc.m.chiave}`)) continue;
      if (!movimentoCitaFattura(inc.m, fat.f.numero)) continue;
      if (inc.m.dataContabile < fat.f.dataDocumento) continue;
      alloca(fat, inc, "numero");
    }
  }
  // Passata 2 — importo del movimento identico al residuo di UNA sola fattura
  // dello stesso cliente (fallback prudente).
  for (const inc of incassi) {
    if (inc.residuo <= 0.01) continue;
    const candidate = aperte.filter(
      (fat) =>
        fat.key === inc.key &&
        fat.residuo > TOLLERANZA_SALDO &&
        Math.abs(fat.residuo - inc.residuo) <= 0.01 &&
        inc.m.dataContabile >= fat.f.dataDocumento &&
        !coppie.has(`${fat.f.nomeFile}|${inc.m.chiave}`),
    );
    if (candidate.length === 1) alloca(candidate[0], inc, "importo");
  }
  return proposte;
}

// --- Parsing dell'export Aruba (xlsx) ---------------------------------------
// L'export del pannello ("Check fatture inviate") ha intestazioni note; il
// direttore vi aggiunge colonne proprie, quindi si mappa PER NOME colonna,
// non per posizione. Righe senza Nome file o Totale → scartate.

const H = {
  numero: "numero",
  nomeFile: "nome file",
  idSdi: "id sdi",
  dataInvio: "data invio",
  dataDocumento: "data documento",
  tipoDocumento: "tipo documento",
  cliente: "cliente",
  piva: "p.iva",
  metodoPagamento: "metodo di pagamento",
  imponibile: "totale imponibile",
  iva: "totale iva",
  totale: "totale documento",
  netto: "netto a pagare",
  statoSdI: "stato",
} as const;

export interface ParseFattureResult {
  rows: FatturaRaw[];
  scartate: number;
}

export function parseFattureMatrice(matrix: unknown[][]): ParseFattureResult | null {
  const headerIdx = matrix.findIndex((r) => {
    const cells = r.map((c) => normalizeTesto(String(c ?? "")));
    return cells.includes(H.nomeFile) && cells.includes(H.numero) && cells.includes(H.totale);
  });
  if (headerIdx < 0) return null;
  const header = matrix[headerIdx].map((c) => normalizeTesto(String(c ?? "")));
  const col = (name: string) => header.indexOf(name);
  const idx = {
    numero: col(H.numero),
    nomeFile: col(H.nomeFile),
    idSdi: col(H.idSdi),
    dataInvio: col(H.dataInvio),
    dataDocumento: col(H.dataDocumento),
    tipoDocumento: col(H.tipoDocumento),
    cliente: col(H.cliente),
    piva: col(H.piva),
    metodoPagamento: col(H.metodoPagamento),
    imponibile: col(H.imponibile),
    iva: col(H.iva),
    totale: col(H.totale),
    netto: col(H.netto),
    statoSdI: col(H.statoSdI),
  };
  const cell = (r: unknown[], i: number) => (i >= 0 ? r[i] : undefined);
  const rows: FatturaRaw[] = [];
  let scartate = 0;
  for (const r of matrix.slice(headerIdx + 1)) {
    const nomeFile = String(cell(r, idx.nomeFile) ?? "").trim();
    const totale = cellToImporto(cell(r, idx.totale));
    const dataDocumento = cellToIsoDate(cell(r, idx.dataDocumento));
    if (!nomeFile || totale == null || !dataDocumento) {
      if (r.some((c) => c != null && String(c).trim() !== "")) scartate++;
      continue;
    }
    rows.push({
      nomeFile,
      numero: String(cell(r, idx.numero) ?? "").trim(),
      idSdi: String(cell(r, idx.idSdi) ?? "").trim(),
      dataInvio: cellToIsoDate(cell(r, idx.dataInvio)) ?? dataDocumento,
      dataDocumento,
      tipoDocumento: String(cell(r, idx.tipoDocumento) ?? "").trim(),
      cliente: canonicalCliente(String(cell(r, idx.cliente) ?? "")).toUpperCase(),
      piva: String(cell(r, idx.piva) ?? "").trim(),
      metodoPagamento: String(cell(r, idx.metodoPagamento) ?? "").trim(),
      imponibile: cellToImporto(cell(r, idx.imponibile)) ?? 0,
      iva: cellToImporto(cell(r, idx.iva)) ?? 0,
      totale,
      netto: cellToImporto(cell(r, idx.netto)) ?? totale,
      statoSdI: String(cell(r, idx.statoSdI) ?? "").trim(),
    });
  }
  return { rows, scartate };
}
