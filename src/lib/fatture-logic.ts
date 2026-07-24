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

export type DirezioneFattura = "Emessa" | "Ricevuta";

export interface FatturaRaw {
  nomeFile: string; // chiave univoca (Title su SharePoint) = nome file SdI
  numero: string; // es. "FPR 201/26" — NON univoco
  idSdi: string;
  dataInvio: string; // YYYY-MM-DD
  dataDocumento: string; // YYYY-MM-DD
  tipoDocumento: string; // "Fattura - TD01" | "Nota di credito - TD04" | ...
  cliente: string; // controparte: cliente (emesse) o fornitore (ricevute)
  piva: string;
  metodoPagamento: string;
  imponibile: number;
  iva: number;
  totale: number;
  netto: number;
  statoSdI: string; // Consegnata / Scartata / ... ("" per le ricevute)
  direzione: DirezioneFattura;
  /** Scadenza DICHIARATA in fattura (DatiPagamento) — quando c'è vince sui
   *  termini di pagamento. YYYY-MM-DD. */
  scadenza?: string;
}

/** P.IVA dell'azienda: decide la direzione di un XML FatturaPA (cedente = noi
 *  → emessa; cessionario = noi → ricevuta). */
export const PIVA_AZIENDA = "16935881009";

/** Chiave univoca NORMALIZZATA dal nome file SdI: la stessa fattura appare
 *  come "IT…_x.xml.p7m" nell'export xlsx e "IT…_x.xml" nello ZIP XML — le
 *  estensioni vanno rimosse o le due sorgenti si duplicherebbero a vicenda. */
export function normalizzaNomeFile(nome: string): string {
  return nome
    .trim()
    .replace(/\.p7m$/i, "")
    .replace(/\.xml$/i, "");
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

/** Individua i REINVII: quando la stessa fattura è stata scartata dallo SdI e
 *  rispedita, in archivio esistono più file con identici numero, data, totale
 *  e controparte. L'XML FatturaPA non porta lo stato SdI, quindi i tentativi
 *  scartati non sono riconoscibili dal singolo file: qui se ne tiene UNO solo
 *  (quello con stato SdI valido se noto, altrimenti il primo per nome file) e
 *  gli altri vengono esclusi dal credito. Ritorna i nomeFile da escludere. */
export function individuaReinvii(fatture: readonly FatturaRaw[]): Set<string> {
  const gruppi = new Map<string, FatturaRaw[]>();
  for (const f of fatture) {
    if (isNotaCredito(f.tipoDocumento) || f.totale <= 0) continue;
    const k = `${f.direzione}|${normalizeTesto(f.numero)}|${f.dataDocumento}|${f.totale.toFixed(2)}|${f.piva}`;
    const g = gruppi.get(k) ?? [];
    g.push(f);
    gruppi.set(k, g);
  }
  const esclusi = new Set<string>();
  for (const g of gruppi.values()) {
    if (g.length < 2) continue;
    const ordinati = [...g].sort((a, b) => a.nomeFile.localeCompare(b.nomeFile));
    const nonScartati = ordinati.filter((f) => !isEsclusaDalCredito(f));
    const vincitore =
      nonScartati.find((f) => /consegnat|accettat|inviat|presa/i.test(f.statoSdI)) ??
      nonScartati[0] ??
      ordinati[0];
    for (const f of ordinati) if (f.nomeFile !== vincitore.nomeFile) esclusi.add(f.nomeFile);
  }
  return esclusi;
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
  // La scadenza dichiarata in fattura (XML DatiPagamento) vince sui termini.
  const scadenza =
    f.scadenza && /^\d{4}-\d{2}-\d{2}$/.test(f.scadenza)
      ? f.scadenza
      : scadenzaFattura(f.dataDocumento, giorniPerCliente(f.cliente, termini));
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

// Uscite che NON sono pagamenti a fornitori: escluse dalla riconciliazione
// delle fatture ricevute.
const TIPOLOGIE_NON_FORNITORE = new Set([
  "Commissioni",
  "Imposte / F24",
  "Imposta di bollo",
  "Prelievo ATM",
  "PagoPA / Multe",
  "Pagamento Salario",
  "Storno",
  "Incasso",
]);

/** Propone gli abbinamenti automatici, per direzione: EMESSE ↔ incassi
 *  (importo>0, tipologia Incasso), RICEVUTE ↔ uscite verso fornitori
 *  (importo<0 con controparte, tipologie non-fornitore escluse). Considera
 *  solo fatture aperte, rispetta gli abbinamenti già registrati (residui per
 *  fattura E per movimento) ed è deterministica. */
export function proponiAbbinamenti(
  fatture: readonly FatturaRaw[],
  movimenti: readonly MovimentoPerRiconciliazione[],
  esistenti: readonly AbbinamentoIncasso[],
  direzione: DirezioneFattura = "Emessa",
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
    .filter(
      (f) =>
        f.direzione === direzione &&
        !isNotaCredito(f.tipoDocumento) &&
        !isEsclusaDalCredito(f) &&
        f.totale > 0,
    )
    .map((f) => ({
      f,
      key: clienteGroupKey(f.cliente),
      residuo: round(f.totale - (incassatoPerFattura.get(f.nomeFile) ?? 0)),
    }))
    .filter((x) => x.residuo > TOLLERANZA_SALDO)
    // Più vecchie prima: un bonifico cumulativo salda in ordine cronologico.
    .sort((a, b) => a.f.dataDocumento.localeCompare(b.f.dataDocumento));

  // Gli importi dei movimenti si trattano in VALORE ASSOLUTO (le uscite sono
  // negative): l'allocazione registrata è sempre positiva.
  const incassi = movimenti
    .filter((m) =>
      direzione === "Emessa"
        ? m.importo > 0 && m.tipologia === "Incasso" && m.cliente
        : m.importo < 0 && m.cliente && !TIPOLOGIE_NON_FORNITORE.has(m.tipologia),
    )
    .map((m) => ({
      m,
      key: clienteGroupKey(m.cliente),
      residuo: round(Math.abs(m.importo) - (allocatoPerMovimento.get(m.chiave) ?? 0)),
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

// --- Parsing XML FatturaPA (tracciato SdI, versione FPR12/FPA12) ------------
// Mini-parser XML senza dipendenze (funziona in browser, Workers e Node):
// ignora attributi e namespace (i prefissi vengono rimossi), gestisce CDATA,
// commenti e prolog. Sufficiente e robusto per il tracciato FatturaPA.

interface XmlNodo {
  tag: string;
  figli: XmlNodo[];
  testo: string;
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export function parseXmlSemplice(xml: string): XmlNodo | null {
  const radice: XmlNodo = { tag: "__root__", figli: [], testo: "" };
  const stack: XmlNodo[] = [radice];
  let i = 0;
  const n = xml.length;
  while (i < n) {
    const lt = xml.indexOf("<", i);
    if (lt < 0) break;
    const testo = xml.slice(i, lt);
    if (testo.trim()) stack[stack.length - 1].testo += decodeXmlEntities(testo.trim());
    if (xml.startsWith("<!--", lt)) {
      const end = xml.indexOf("-->", lt);
      if (end < 0) break;
      i = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const end = xml.indexOf("]]>", lt);
      if (end < 0) break;
      stack[stack.length - 1].testo += xml.slice(lt + 9, end);
      i = end + 3;
      continue;
    }
    if (xml.startsWith("<?", lt)) {
      const end = xml.indexOf("?>", lt);
      if (end < 0) break;
      i = end + 2;
      continue;
    }
    if (xml.startsWith("<!", lt)) {
      const end = xml.indexOf(">", lt);
      if (end < 0) break;
      i = end + 1;
      continue;
    }
    const gt = xml.indexOf(">", lt);
    if (gt < 0) break;
    const dentro = xml.slice(lt + 1, gt).trim();
    if (dentro.startsWith("/")) {
      if (stack.length > 1) stack.pop();
      i = gt + 1;
      continue;
    }
    const autoChiuso = dentro.endsWith("/");
    const nomeGrezzo = dentro.replace(/\/$/, "").split(/[\s]/)[0];
    const tag = nomeGrezzo.includes(":") ? nomeGrezzo.split(":").pop()! : nomeGrezzo;
    const nodo: XmlNodo = { tag, figli: [], testo: "" };
    stack[stack.length - 1].figli.push(nodo);
    if (!autoChiuso) stack.push(nodo);
    i = gt + 1;
  }
  return radice.figli[0] ?? null;
}

function figliDi(nodo: XmlNodo, percorso: string): XmlNodo[] {
  let correnti = [nodo];
  for (const parte of percorso.split("/")) {
    const prossimi: XmlNodo[] = [];
    for (const c of correnti) prossimi.push(...c.figli.filter((f) => f.tag === parte));
    correnti = prossimi;
  }
  return correnti;
}
function testoDi(nodo: XmlNodo, percorso: string): string {
  return figliDi(nodo, percorso)[0]?.testo.trim() ?? "";
}
function numeroDi(nodo: XmlNodo, percorso: string): number {
  const n = Number(testoDi(nodo, percorso).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

const TIPO_DOC_LABEL: Record<string, string> = {
  TD01: "Fattura",
  TD02: "Acconto/anticipo su fattura",
  TD03: "Acconto/anticipo su parcella",
  TD04: "Nota di credito",
  TD05: "Nota di debito",
  TD06: "Parcella",
  TD16: "Integrazione reverse charge",
  TD17: "Autofattura estero",
  TD18: "Integrazione acquisto UE",
  TD19: "Autofattura art.17",
  TD20: "Autofattura regolarizzazione",
  TD24: "Fattura differita",
  TD25: "Fattura differita (triangolare)",
  TD26: "Cessione beni ammortizzabili",
  TD27: "Autoconsumo/cessioni gratuite",
};

const MODALITA_PAG_LABEL: Record<string, string> = {
  MP01: "Contanti",
  MP02: "Assegno",
  MP05: "Bonifico",
  MP08: "Carta",
  MP12: "RiBa",
  MP15: "Giroconto",
  MP16: "Domiciliazione bancaria",
  MP17: "Domiciliazione postale",
  MP19: "SDD",
  MP20: "SDD CORE",
  MP21: "SDD B2B",
  MP23: "PagoPA",
};

function normalizzaPiva(v: string): string {
  return v.replace(/\D/g, "").replace(/^0+/, "");
}

export interface ParseXmlFattureResult {
  rows: FatturaRaw[];
  /** File XML non riconosciuti come FatturaPA o senza direzione certa. */
  scartati: string[];
}

/** Interpreta UN file XML FatturaPA. La direzione deriva dalla P.IVA
 *  aziendale: cedente = noi → emessa; cessionario = noi → ricevuta. Un file
 *  può contenere più FatturaElettronicaBody (rari: lotti): in tal caso la
 *  chiave dei body successivi è suffissata con #2, #3… */
export function parseFatturaPA(
  xmlText: string,
  nomeFileGrezzo: string,
  pivaAzienda: string = PIVA_AZIENDA,
): ParseXmlFattureResult {
  const nomeFile = normalizzaNomeFile(nomeFileGrezzo);
  const scartati: string[] = [];
  const root = parseXmlSemplice(xmlText);
  if (!root || root.tag !== "FatturaElettronica") {
    return { rows: [], scartati: [nomeFile] };
  }
  const header = figliDi(root, "FatturaElettronicaHeader")[0];
  if (!header) return { rows: [], scartati: [nomeFile] };

  const anagrafica = (lato: "CedentePrestatore" | "CessionarioCommittente") => {
    const den = testoDi(header, `${lato}/DatiAnagrafici/Anagrafica/Denominazione`);
    const nome = testoDi(header, `${lato}/DatiAnagrafici/Anagrafica/Nome`);
    const cognome = testoDi(header, `${lato}/DatiAnagrafici/Anagrafica/Cognome`);
    return {
      nome: den || `${nome} ${cognome}`.trim(),
      piva:
        testoDi(header, `${lato}/DatiAnagrafici/IdFiscaleIVA/IdCodice`) ||
        testoDi(header, `${lato}/DatiAnagrafici/CodiceFiscale`),
    };
  };
  const cedente = anagrafica("CedentePrestatore");
  const cessionario = anagrafica("CessionarioCommittente");
  const noi = normalizzaPiva(pivaAzienda);
  let direzione: DirezioneFattura;
  let controparte: { nome: string; piva: string };
  if (normalizzaPiva(cedente.piva) === noi) {
    direzione = "Emessa";
    controparte = cessionario;
  } else if (normalizzaPiva(cessionario.piva) === noi) {
    direzione = "Ricevuta";
    controparte = cedente;
  } else {
    return { rows: [], scartati: [nomeFile] };
  }

  const rows: FatturaRaw[] = [];
  const bodies = figliDi(root, "FatturaElettronicaBody");
  bodies.forEach((body, idx) => {
    const doc = figliDi(body, "DatiGenerali/DatiGeneraliDocumento")[0];
    if (!doc) {
      scartati.push(`${nomeFile}#${idx + 1}`);
      return;
    }
    const td = testoDi(doc, "TipoDocumento").toUpperCase();
    const dataDocumento = testoDi(doc, "Data").slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dataDocumento)) {
      scartati.push(`${nomeFile}#${idx + 1}`);
      return;
    }
    const riepiloghi = figliDi(body, "DatiBeniServizi/DatiRiepilogo");
    const imponibile = riepiloghi.reduce((s, r) => s + numeroDi(r, "ImponibileImporto"), 0);
    const iva = riepiloghi.reduce((s, r) => s + numeroDi(r, "Imposta"), 0);
    const totDich = numeroDi(doc, "ImportoTotaleDocumento");
    const totale = totDich || Math.round((imponibile + iva) * 100) / 100;
    const pagamenti = figliDi(body, "DatiPagamento/DettaglioPagamento");
    // Netto a pagare: somma degli ImportoPagamento (tiene conto di ritenute e
    // simili); se assente, il totale documento.
    const nettoPag = pagamenti.reduce((s, p) => s + numeroDi(p, "ImportoPagamento"), 0);
    const scadenze = pagamenti
      .map((p) => testoDi(p, "DataScadenzaPagamento").slice(0, 10))
      .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))
      .sort();
    const mp = pagamenti.map((p) => testoDi(p, "ModalitaPagamento")).find(Boolean) ?? "";
    rows.push({
      nomeFile: idx === 0 ? nomeFile : `${nomeFile}#${idx + 1}`,
      numero: testoDi(doc, "Numero"),
      idSdi: "",
      dataInvio: dataDocumento,
      dataDocumento,
      tipoDocumento: td ? `${TIPO_DOC_LABEL[td] ?? "Documento"} - ${td}` : "",
      cliente: canonicalCliente(controparte.nome).toUpperCase(),
      piva: controparte.piva,
      metodoPagamento: mp ? `${mp} - ${MODALITA_PAG_LABEL[mp] ?? mp}` : "",
      imponibile: Math.round(imponibile * 100) / 100,
      iva: Math.round(iva * 100) / 100,
      totale,
      netto: nettoPag ? Math.round(nettoPag * 100) / 100 : totale,
      statoSdI: "",
      direzione,
      scadenza: scadenze.length ? scadenze[scadenze.length - 1] : undefined,
    });
  });
  return { rows, scartati };
}

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
    const nomeFile = normalizzaNomeFile(String(cell(r, idx.nomeFile) ?? ""));
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
      direzione: "Emessa", // l'export "Check fatture inviate" è delle emesse
    });
  }
  return { rows, scartate };
}
