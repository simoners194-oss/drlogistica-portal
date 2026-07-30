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
  /** Stato d'incasso REGISTRATO SU ARUBA (colonna "Incassi" dell'export):
   *  "Incassata" | "Non incassata" | "Non gestita" | "". È la fonte primaria
   *  dello stato; la riconciliazione bancaria resta un'informazione in più. */
  incassoAruba?: string;
  /** Data incasso registrata su Aruba (YYYY-MM-DD). */
  dataIncasso?: string;
  /** Incassato REGISTRATO SU ARUBA (somma delle rate del report movimenti):
   *  è l'unico dato che quantifica gli incassi PARZIALI. undefined se il
   *  report movimenti non è stato caricato per quell'anno. */
  incassatoAruba?: number;
  /** Solo per le NOTE DI CREDITO: numero della fattura che rettificano,
   *  dichiarato nell'XML in DatiGenerali/DatiFattureCollegate/IdDocumento.
   *  Serve ad abbattere il credito della fattura collegata. */
  rettificaNumero?: string;
}

export type IncassoAruba = "Incassata" | "Non incassata" | "Non gestita" | "";

/** Normalizza il valore della colonna "Incassi" dell'export Aruba. */
export function parseIncassoAruba(v: unknown): IncassoAruba {
  const s = normalizeTesto(String(v ?? ""));
  if (s.startsWith("incassat")) return "Incassata";
  if (s.startsWith("non incassat")) return "Non incassata";
  if (s.startsWith("non gestit")) return "Non gestita";
  return "";
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
  // "Auto" = numero citato/importo esatto; "FIFO" = imputazione a scalare
  // (pulsante dedicato); "Manuale" = scelto dall'utente.
  origine: "Auto" | "Manuale" | "FIFO";
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

/** I documenti che NON esistono ai fini contabili: gli scarti dello SdI (una
 *  fattura scartata non è mai stata emessa, quindi non può essere incassata) e
 *  i tentativi doppi di un reinvio. Vanno tenuti fuori da elenchi, conteggi,
 *  abbinamenti e collegamenti delle note di credito: se restano dentro anche
 *  solo come riga, il fatturato di un cliente si moltiplica per il numero di
 *  tentativi (FPR 38/25: 4 copie, 3 scartate → fatturato contato 4 volte). */
export function fattureEscluse(fatture: readonly FatturaRaw[]): Set<string> {
  const esclusi = individuaReinvii(fatture);
  for (const f of fatture) if (isEsclusaDalCredito(f)) esclusi.add(f.nomeFile);
  return esclusi;
}

// --- Report MOVIMENTI di Aruba (rate incassate/pagate per fattura) ----------
// È il terzo export del pannello: per ogni fattura elenca le rate con data e
// importo. È l'unica fonte che quantifica gli incassi PARZIALI — lo stato
// testuale del report fatture dice solo incassata/non incassata.

export interface MovimentoAruba {
  data: string; // YYYY-MM-DD
  cliente: string;
  numeroFattura: string;
  /** Data della fattura dichiarata nel report: serve a capire quali ANNI
   *  copre il file, per distinguere "nessuna rata" da "report non caricato". */
  dataFattura: string;
  /** INCASSO (fatture emesse) o PAGAMENTO (ricevute). */
  flusso: "INCASSO" | "PAGAMENTO";
  importo: number; // sempre positivo
  modalita: string;
}

const H_MOV = {
  data: "data",
  cliente: "cliente/fornitore",
  numero: "numero fattura",
  dataFattura: "data fattura",
  flusso: "flusso",
  importo: "importo",
  modalita: "modalita di pagamento",
} as const;

/** Riconosce ed estrae il report movimenti; null se il foglio non è quello. */
export function parseMovimentiArubaMatrice(matrix: unknown[][]): MovimentoAruba[] | null {
  const headerIdx = matrix.findIndex((r) => {
    const c = r.map((x) => normalizeTesto(String(x ?? "")));
    return c.includes(H_MOV.numero) && c.includes(H_MOV.flusso) && c.includes(H_MOV.importo);
  });
  if (headerIdx < 0) return null;
  const header = matrix[headerIdx].map((c) => normalizeTesto(String(c ?? "")));
  const col = (n: string) => header.indexOf(n);
  const idx = {
    data: col(H_MOV.data),
    cliente: col(H_MOV.cliente),
    numero: col(H_MOV.numero),
    dataFattura: col(H_MOV.dataFattura),
    flusso: col(H_MOV.flusso),
    importo: col(H_MOV.importo),
    modalita: col(H_MOV.modalita),
  };
  const out: MovimentoAruba[] = [];
  for (const r of matrix.slice(headerIdx + 1)) {
    const numero = String(r[idx.numero] ?? "").trim();
    const importo = cellToImporto(r[idx.importo]);
    const flusso = String(r[idx.flusso] ?? "")
      .trim()
      .toUpperCase();
    if (!numero || importo == null || (flusso !== "INCASSO" && flusso !== "PAGAMENTO")) continue;
    out.push({
      data: cellToIsoDate(r[idx.data]) ?? "",
      cliente: String(r[idx.cliente] ?? "").trim(),
      numeroFattura: numero,
      dataFattura: idx.dataFattura >= 0 ? (cellToIsoDate(r[idx.dataFattura]) ?? "") : "",
      flusso: flusso as "INCASSO" | "PAGAMENTO",
      importo: Math.abs(importo),
      modalita: idx.modalita >= 0 ? String(r[idx.modalita] ?? "").trim() : "",
    });
  }
  return out.length ? out : null;
}

/** Somma le rate per fattura e le aggancia all'archivio (numero + controparte
 *  + direzione: due clienti possono avere fatture con lo stesso numero). */
export function aggregaIncassiAruba(
  movimenti: readonly MovimentoAruba[],
  fatture: readonly FatturaRaw[],
  /** Per gli ANNI coperti dal report, le fatture senza alcuna rata vengono
   *  segnate a ZERO: così "nessun incasso registrato" si distingue da "report
   *  di quell'anno non importato". */
  azzeraNelPeriodo = true,
  /** Scarti SdI e reinvii: non possono ricevere incassi (vedi fattureEscluse). */
  esclusi: ReadonlySet<string> = new Set(),
): Map<string, { incassato: number; ultimaData: string; rate: number }> {
  const valide = fatture.filter((f) => !esclusi.has(f.nomeFile));
  const perChiave = new Map<string, FatturaRaw>();
  // Indice per il fallback (stesso numero, controparte scritta diversamente):
  // PRECALCOLATO, mai una scansione lineare. La versione con `find` dentro il
  // ciclo rinormalizzava ogni numero a ogni rata — migliaia di rate per
  // migliaia di fatture = decine di milioni di regex, e il browser si
  // bloccava per ore sull'import del report movimenti.
  const perNumero = new Map<string, FatturaRaw>();
  for (const f of valide) {
    perChiave.set(`${f.direzione}|${clienteGroupKey(f.cliente)}|${normalizeTesto(f.numero)}`, f);
    const kNum = `${f.direzione}|${normalizeTesto(f.numero)}`;
    if (!perNumero.has(kNum)) perNumero.set(kNum, f);
  }
  const out = new Map<string, { incassato: number; ultimaData: string; rate: number }>();
  for (const m of movimenti) {
    const direzione: DirezioneFattura = m.flusso === "INCASSO" ? "Emessa" : "Ricevuta";
    const numero = normalizeTesto(m.numeroFattura);
    const f =
      perChiave.get(`${direzione}|${clienteGroupKey(m.cliente)}|${numero}`) ??
      perNumero.get(`${direzione}|${numero}`);
    if (!f) continue;
    const v = out.get(f.nomeFile) ?? { incassato: 0, ultimaData: "", rate: 0 };
    v.incassato = Math.round((v.incassato + m.importo) * 100) / 100;
    v.rate++;
    if (m.data > v.ultimaData) v.ultimaData = m.data;
    out.set(f.nomeFile, v);
  }
  if (azzeraNelPeriodo) {
    // Anni coperti dal file, per direzione: si guarda la DATA FATTURA delle
    // rate (il report elenca solo le fatture con movimenti, quindi l'anno si
    // deduce da lì).
    const anniPerDirezione = new Map<DirezioneFattura, Set<string>>();
    for (const m of movimenti) {
      const d: DirezioneFattura = m.flusso === "INCASSO" ? "Emessa" : "Ricevuta";
      const anno = (m.dataFattura || m.data).slice(0, 4);
      if (!anno) continue;
      const set = anniPerDirezione.get(d) ?? new Set<string>();
      set.add(anno);
      anniPerDirezione.set(d, set);
    }
    // Intervallo di DATE RATA coperto dal file: un file completo di un anno
    // contiene solo le rate DATATE in quell'anno.
    let minData = "";
    let maxData = "";
    for (const m of movimenti) {
      if (!m.data) continue;
      if (!minData || m.data < minData) minData = m.data;
      if (!maxData || m.data > maxData) maxData = m.data;
    }
    for (const f of valide) {
      if (out.has(f.nomeFile)) continue;
      if (isNotaCredito(f.tipoDocumento) || isEsclusaDalCredito(f)) continue;
      if (!anniPerDirezione.get(f.direzione)?.has(f.dataDocumento.slice(0, 4))) continue;
      // CONFINE D'ANNO: una fattura di dicembre pagata a gennaio non ha rate
      // nel file del suo anno, ed è NORMALE — il file di un anno non dice
      // nulla sulle fatture incassate in un altro anno. Un incasso già
      // registrato si azzera solo se la sua ULTIMA RATA cade nel periodo
      // coperto dal file: quel dato è stato smentito dal file stesso.
      // (Successo davvero: il file 2025 ha azzerato la 336/25, incassata a
      // gennaio 2026 dal file precedente.)
      const registrato = typeof f.incassatoAruba === "number" && f.incassatoAruba > 0;
      if (
        registrato &&
        !(f.dataIncasso && minData && f.dataIncasso >= minData && f.dataIncasso <= maxData)
      )
        continue;
      out.set(f.nomeFile, { incassato: 0, ultimaData: "", rate: 0 });
    }
  }
  return out;
}

/** Collega le NOTE DI CREDITO alle fatture che rettificano (numero dichiarato
 *  nell'XML). Ritorna, per ogni fattura, l'importo complessivo delle NC
 *  collegate e i loro numeri: il credito residuo va calcolato al netto,
 *  perché il cliente paga la differenza (compensazione).
 *  Il collegamento richiede stessa direzione e stessa controparte: due clienti
 *  possono avere fatture con lo stesso numero. */
export function collegaNoteCredito(
  fatture: readonly FatturaRaw[],
  /** Scarti SdI e reinvii: non sono bersagli validi (vedi fattureEscluse). */
  esclusi: ReadonlySet<string> = new Set(),
): Map<string, { importo: number; numeri: string[] }> {
  const perChiave = new Map<string, FatturaRaw>();
  for (const f of fatture) {
    if (isNotaCredito(f.tipoDocumento) || f.totale <= 0 || esclusi.has(f.nomeFile)) continue;
    perChiave.set(`${f.direzione}|${clienteGroupKey(f.cliente)}|${normalizeTesto(f.numero)}`, f);
  }
  const out = new Map<string, { importo: number; numeri: string[] }>();
  for (const nc of fatture) {
    if (!isNotaCredito(nc.tipoDocumento) || !nc.rettificaNumero || esclusi.has(nc.nomeFile))
      continue;
    for (const rif of nc.rettificaNumero.split(/[,;]+/)) {
      const target = perChiave.get(
        `${nc.direzione}|${clienteGroupKey(nc.cliente)}|${normalizeTesto(rif)}`,
      );
      if (!target) continue;
      const riga = out.get(target.nomeFile) ?? { importo: 0, numeri: [] };
      riga.importo = Math.round((riga.importo + Math.abs(nc.totale)) * 100) / 100;
      riga.numeri.push(nc.numero);
      out.set(target.nomeFile, riga);
    }
  }
  return out;
}

// --- Termini di pagamento ----------------------------------------------------

/** Giorni di pagamento per un cliente (match per chiave canonica; la riga
 *  senza descrizione è quella generica). */
export function giorniPerCliente(cliente: string, termini: readonly TerminePagamento[]): number {
  const key = clienteGroupKey(cliente);
  let match = termini.filter((t) => clienteGroupKey(t.cliente) === key && t.giorni > 0);
  if (!match.length) {
    // Il foglio contratti usa nomi BREVI ("IMILE" per "IMILE ITALY SRL"):
    // vale il termine il cui nome e' interamente contenuto nel nome del
    // cliente; a parita', vince il piu' specifico (piu' parole).
    const parole = new Set(key.split(" ").filter(Boolean));
    match = termini
      .filter((t) => {
        if (t.giorni <= 0) return false;
        const tk = clienteGroupKey(t.cliente).split(" ").filter(Boolean);
        return tk.length > 0 && tk.every((x) => parole.has(x));
      })
      .sort(
        (a, b) =>
          clienteGroupKey(b.cliente).split(" ").length -
          clienteGroupKey(a.cliente).split(" ").length,
      );
  }
  if (!match.length) return TERMINI_DEFAULT_GIORNI;
  const generico = match.find((t) => !t.descrizione?.trim());
  return (generico ?? match[0]).giorni;
}

// --- Import del foglio contratti del direttore -------------------------------
// "CLIENTI_FORNITORI check contratti.xlsx": colonna "GG pagamento" = giorni
// contrattuali VERIFICATI per cliente. Righe senza valore = default 30 a
// runtime (non si salvano). Riconosciuto dalle intestazioni, come gli altri.

export interface TermineImport {
  cliente: string;
  giorni: number;
}

export function parseTerminiMatrice(matrix: unknown[][]): TermineImport[] | null {
  const headerIdx = matrix.findIndex((r) => {
    const c = r.map((x) => normalizeTesto(String(x ?? "")));
    return c.includes("cliente") && c.includes("gg pagamento");
  });
  if (headerIdx < 0) return null;
  const header = matrix[headerIdx].map((c) => normalizeTesto(String(c ?? "")));
  const iCli = header.indexOf("cliente");
  const iGg = header.indexOf("gg pagamento");
  const out: TermineImport[] = [];
  for (const r of matrix.slice(headerIdx + 1)) {
    const cliente = String(r[iCli] ?? "").trim();
    const giorni = Number(r[iGg]);
    if (!cliente || !Number.isFinite(giorni) || giorni <= 0) continue;
    out.push({ cliente, giorni: Math.round(giorni) });
  }
  return out.length ? out : null;
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

// DUE LETTURE AFFIANCATE, nessuna prevale sull'altra:
//  - FATTURAZIONE: quanto risulta all'amministrazione dentro Aruba;
//  - BANCA: quanto risulta arrivato sul conto dagli abbinamenti.
// Possono dire cose diverse: la differenza è un'informazione, non un errore da
// nascondere. Il campo `stato` resta come lettura combinata (fatturazione se
// c'è, altrimenti banca) per filtri e ordinamenti.
export interface FatturaStato {
  /** Incassato secondo la lettura combinata. */
  incassato: number;
  /** Incassato secondo la FATTURAZIONE (Aruba); null se non gestita lì. */
  incassatoFatturazione: number | null;
  /** Incassato secondo i soli ABBINAMENTI bancari. */
  incassatoBanca: number;
  /** Residuo secondo la lettura combinata. */
  residuo: number;
  /** Residuo secondo la FATTURAZIONE; null se non gestita su Aruba. */
  residuoFatturazione: number | null;
  /** Residuo secondo i soli abbinamenti bancari. */
  residuoBanca: number;
  /** Lettura combinata (fatturazione se presente, altrimenti banca). */
  stato: StatoIncasso;
  /** Stato secondo la FATTURAZIONE; null se la fattura non è gestita su Aruba. */
  statoFatturazione: StatoIncasso | null;
  /** Stato secondo gli INCASSI registrati su Aruba (report movimenti):
   *  quantifica i parziali. null se il report non è stato caricato. */
  statoIncassi: StatoIncasso | null;
  /** Importo incassato secondo il report movimenti; null se non caricato. */
  incassatoIncassi: number | null;
  /** Residuo secondo il report movimenti; null se non caricato. */
  residuoIncassi: number | null;
  /** Stato ricavato dai soli abbinamenti bancari (riconciliazione). */
  statoBanca: StatoIncasso;
  /** Valore registrato su Aruba, "" se la fattura non è gestita lì. */
  aruba: IncassoAruba;
  /** Aruba dice incassata ma la banca non lo conferma (o viceversa). */
  discordante: boolean;
  /** Le note di credito collegate coprono l'intero importo: non c'è più
   *  nulla da incassare, ma non è un pagamento. */
  annullataDaNC: boolean;
  /** Importo delle note di credito collegate (0 se nessuna). */
  notaCredito: number;
  scadenza: string;
  inRitardo: boolean;
  giorniRitardo: number;
}

export function computeStatoFattura(
  f: FatturaRaw,
  incassato: number,
  termini: readonly TerminePagamento[],
  oggiISO: string,
  /** Importo delle note di credito collegate: il credito si calcola al netto. */
  notaCredito = 0,
): FatturaStato {
  // ATTIVE: la scadenza si calcola SEMPRE dai termini contrattuali per
  // cliente (foglio contratti del direttore; senza termine = 30 giorni): il
  // conteggio del ritardo parte da li'. Per le PASSIVE ("solo attive per
  // ora") vale ancora la scadenza dichiarata nell'XML, quando c'e'.
  const scadenza =
    f.direzione === "Emessa"
      ? scadenzaFattura(f.dataDocumento, giorniPerCliente(f.cliente, termini))
      : f.scadenza && /^\d{4}-\d{2}-\d{2}$/.test(f.scadenza)
        ? f.scadenza
        : scadenzaFattura(f.dataDocumento, giorniPerCliente(f.cliente, termini));
  // Le note di credito non si "incassano" e le scartate/rifiutate dallo SdI
  // non sono crediti: entrambe fuori dal computo di residui e ritardi.
  if (isNotaCredito(f.tipoDocumento) || isEsclusaDalCredito(f) || f.totale <= 0) {
    // Fuori dal credito, ma le informazioni delle fonti NON si buttano: per
    // una nota di credito "incassata" su Aruba significa che è stata
    // compensata/liquidata, e va potuto leggere (le viste la etichettano
    // come compensata, non come pagata).
    const arubaNC = parseIncassoAruba(f.incassoAruba);
    const impNC = Math.abs(f.totale);
    const incassatoNC = typeof f.incassatoAruba === "number" ? Math.abs(f.incassatoAruba) : null;
    const chiusa = (v: number | null) =>
      v == null ? null : Math.max(0, impNC - v) <= TOLLERANZA_SALDO ? "Pagata" : "Non incassata";
    return {
      annullataDaNC: false,
      notaCredito: 0,
      incassato,
      incassatoFatturazione:
        arubaNC === "Incassata" ? impNC : arubaNC === "Non incassata" ? 0 : null,
      incassatoIncassi: incassatoNC,
      incassatoBanca: incassato,
      residuo: 0,
      residuoFatturazione: null,
      residuoIncassi: incassatoNC == null ? null : Math.max(0, impNC - incassatoNC),
      residuoBanca: 0,
      stato: "NC",
      statoFatturazione:
        arubaNC === "Incassata" ? "Pagata" : arubaNC === "Non incassata" ? "Non incassata" : null,
      statoIncassi: chiusa(incassatoNC),
      statoBanca: "NC",
      aruba: arubaNC,
      discordante: false,
      scadenza,
      inRitardo: false,
      giorniRitardo: 0,
    };
  }

  // Base del credito: totale al netto delle note di credito collegate.
  const base = Math.max(0, Math.round((f.totale - notaCredito) * 100) / 100);
  const residuoBanca = Math.max(0, base - incassato);
  // Stato dagli ABBINAMENTI bancari (riconciliazione): informazione di
  // dettaglio, mostra quanto risulta effettivamente arrivato sul conto.
  const statoBanca: StatoIncasso =
    residuoBanca <= TOLLERANZA_SALDO
      ? "Pagata"
      : incassato > TOLLERANZA_SALDO
        ? "Parziale"
        : "Non incassata";
  const aruba = parseIncassoAruba(f.incassoAruba);
  // Discordanza: utile per capire cosa manca in banca (o cosa la banca ha
  // trovato e l'amministrazione non ha ancora registrato su Aruba).
  // Sulle PASSIVE la banca non è una controprova: molti costi non transitano
  // dal conto aziendale (carte di credito dei soci, contanti) pur avendo
  // fattura, quindi "pagata su Aruba ma nessun movimento" è la normalità e
  // non va segnalata. Resta invece informativo il caso opposto: dal conto è
  // uscito un pagamento che la fatturazione non ha registrato.
  const discordante =
    f.direzione === "Ricevuta"
      ? aruba === "Non incassata" && statoBanca === "Pagata"
      : (aruba === "Incassata" && statoBanca !== "Pagata") ||
        (aruba === "Non incassata" && statoBanca === "Pagata");
  // Lettura FATTURAZIONE: esiste solo se Aruba ha uno stato registrato.
  // "Non incassata" su Aruba non esclude un acconto già visto in banca: in
  // quel caso la fatturazione mostra comunque il parziale che risulta.
  // LETTURA 1 — FATTURAZIONE: lo stato testuale registrato su Aruba, invariato.
  const statoFatturazione: StatoIncasso | null =
    aruba === "Incassata" ? "Pagata" : aruba === "Non incassata" ? "Non incassata" : null;
  const incassatoFatturazione = aruba === "Incassata" ? base : aruba === "Non incassata" ? 0 : null;
  // LETTURA 2 — INCASSI registrati (report movimenti): stato E importo, con i
  // parziali quantificati al centesimo. Indipendente dalla lettura 1: le due
  // possono divergere ed è esattamente ciò che si vuole poter confrontare.
  const daMovimenti = typeof f.incassatoAruba === "number";
  const incassatoIncassi = daMovimenti ? f.incassatoAruba! : null;
  const residuoIncassi = incassatoIncassi == null ? null : Math.max(0, base - incassatoIncassi);
  const statoIncassi: StatoIncasso | null =
    incassatoIncassi == null
      ? null
      : residuoIncassi! <= TOLLERANZA_SALDO
        ? "Pagata"
        : incassatoIncassi > TOLLERANZA_SALDO
          ? "Parziale"
          : "Non incassata";
  // Lettura combinata (solo per filtri, ritardi e ordinamenti): la fonte più
  // precisa disponibile.
  const stato: StatoIncasso = statoIncassi ?? statoFatturazione ?? statoBanca;
  const residuoFatturazione =
    incassatoFatturazione == null ? null : Math.max(0, base - incassatoFatturazione);
  // Lettura combinata: serve a filtri, ritardi e ordinamenti. Si prende la
  // fonte più precisa disponibile (incassi registrati → fatturazione → banca);
  // nelle viste le TRE colonne restano comunque affiancate.
  const inRitardo = stato !== "Pagata" && oggiISO > scadenza;
  const giorniRitardo = inRitardo
    ? Math.floor(
        (new Date(`${oggiISO}T00:00:00`).getTime() - new Date(`${scadenza}T00:00:00`).getTime()) /
          86400000,
      )
    : 0;
  const residuo = residuoIncassi ?? residuoFatturazione ?? residuoBanca;
  const incassatoComb = incassatoIncassi ?? incassatoFatturazione ?? incassato;
  return {
    // Coperta per intero dalle note di credito: nulla da incassare, ma non è
    // un incasso — chi legge deve poter distinguere le due cose.
    annullataDaNC: notaCredito > 0 && base <= TOLLERANZA_SALDO,
    notaCredito,
    incassato: incassatoComb,
    incassatoFatturazione,
    incassatoIncassi,
    incassatoBanca: incassato,
    residuo,
    residuoFatturazione,
    residuoIncassi,
    residuoBanca,
    stato,
    statoFatturazione,
    statoIncassi,
    statoBanca,
    aruba,
    discordante,
    scadenza,
    inRitardo,
    giorniRitardo,
  };
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
  // Credito al netto delle note di credito collegate: il cliente paga la
  // differenza, quindi non va cercato in banca l'importo pieno.
  const noteCredito = collegaNoteCredito(fatture);
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
      residuo: round(
        f.totale -
          (noteCredito.get(f.nomeFile)?.importo ?? 0) -
          (incassatoPerFattura.get(f.nomeFile) ?? 0),
      ),
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

/** Riconciliazione A SCALARE (FIFO): i clienti pagano per acconti tondi,
 *  saldi mensili e compensazioni — importi che non coincidono mai con le
 *  singole fatture. Qui ogni movimento residuo viene imputato alle fatture
 *  APERTE più vecchie dello stesso cliente (parziali inclusi), come
 *  l'imputazione al debito più antico dell'art. 1193 c.c. Va lanciata
 *  ESPLICITAMENTE (pulsante dedicato): l'attribuzione per-fattura è
 *  contabile, non documentale, e gli abbinamenti "FIFO" restano
 *  riconoscibili ed eliminabili. Stesse regole di selezione movimenti di
 *  proponiAbbinamenti; il movimento non copre fatture posteriori alla sua
 *  data. */
export function proponiAbbinamentiFIFO(
  fatture: readonly FatturaRaw[],
  movimenti: readonly MovimentoPerRiconciliazione[],
  esistenti: readonly AbbinamentoIncasso[],
  direzione: DirezioneFattura = "Emessa",
): PropostaAbbinamento[] {
  const round = (n: number) => Math.round(n * 100) / 100;
  const incassatoPerFattura = new Map<string, number>();
  const allocatoPerMovimento = new Map<string, number>();
  for (const a of esistenti) {
    incassatoPerFattura.set(
      a.fatturaFile,
      (incassatoPerFattura.get(a.fatturaFile) ?? 0) + a.importo,
    );
    allocatoPerMovimento.set(
      a.movimentoChiave,
      (allocatoPerMovimento.get(a.movimentoChiave) ?? 0) + a.importo,
    );
  }
  const noteCredito = collegaNoteCredito(fatture);
  const apertePerCliente = new Map<string, { f: FatturaRaw; residuo: number }[]>();
  for (const f of fatture) {
    if (f.direzione !== direzione) continue;
    if (isNotaCredito(f.tipoDocumento) || isEsclusaDalCredito(f) || f.totale <= 0) continue;
    const residuo = round(
      f.totale -
        (noteCredito.get(f.nomeFile)?.importo ?? 0) -
        (incassatoPerFattura.get(f.nomeFile) ?? 0),
    );
    if (residuo <= TOLLERANZA_SALDO) continue;
    const key = clienteGroupKey(f.cliente);
    const l = apertePerCliente.get(key) ?? [];
    l.push({ f, residuo });
    apertePerCliente.set(key, l);
  }
  for (const l of apertePerCliente.values())
    l.sort((a, b) => a.f.dataDocumento.localeCompare(b.f.dataDocumento));

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
  for (const inc of incassi) {
    const aperte = apertePerCliente.get(inc.key);
    if (!aperte) continue;
    for (const fat of aperte) {
      if (inc.residuo <= 0.01) break;
      if (fat.residuo <= 0.01) continue;
      if (inc.m.dataContabile < fat.f.dataDocumento) continue;
      const importo = round(Math.min(fat.residuo, inc.residuo));
      proposte.push({
        fatturaFile: fat.f.nomeFile,
        movimentoChiave: inc.m.chiave,
        importo,
        origine: "FIFO",
        motivo: "importo",
      });
      fat.residuo = round(fat.residuo - importo);
      inc.residuo = round(inc.residuo - importo);
    }
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
    // Nota di credito: fattura rettificata (può essercene più d'una).
    const collegate = figliDi(body, "DatiGenerali/DatiFattureCollegate")
      .map((c) => testoDi(c, "IdDocumento").trim())
      .filter(Boolean);
    // Le note di credito valgono in DIMINUZIONE: l'XML le scrive con importi
    // positivi (il segno sta nel tipo documento), l'export xlsx negativi. Qui
    // si uniforma al segno negativo, così la lettura è coerente ovunque.
    const segno = td === "TD04" ? -1 : 1;
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
      imponibile: (segno * Math.round(Math.abs(imponibile) * 100)) / 100,
      iva: (segno * Math.round(Math.abs(iva) * 100)) / 100,
      totale: segno * Math.abs(totale),
      netto: segno * Math.abs(nettoPag ? Math.round(nettoPag * 100) / 100 : totale),
      statoSdI: "",
      direzione,
      scadenza: scadenze.length ? scadenze[scadenze.length - 1] : undefined,
      rettificaNumero: collegate.length ? collegate.join(", ") : undefined,
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
  // Stato d'incasso gestito dall'amministrazione dentro Aruba.
  incassoAruba: "incassi",
  dataIncasso: "data incasso",
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
  // L'export delle RICEVUTE ha lo stesso tracciato ma la controparte si
  // chiama "fornitore": riconoscerlo qui evita il disastro gia' successo
  // (1.103 fatture di fornitori archiviate come emesse, senza controparte).
  const colFornitore = col("fornitore");
  const ricevute = col(H.cliente) < 0 && colFornitore >= 0;
  const idx = {
    numero: col(H.numero),
    nomeFile: col(H.nomeFile),
    idSdi: col(H.idSdi),
    dataInvio: col(H.dataInvio),
    dataDocumento: col(H.dataDocumento),
    tipoDocumento: col(H.tipoDocumento),
    cliente: ricevute ? colFornitore : col(H.cliente),
    piva: col(H.piva),
    metodoPagamento: col(H.metodoPagamento),
    imponibile: col(H.imponibile),
    iva: col(H.iva),
    totale: col(H.totale),
    netto: col(H.netto),
    statoSdI: col(H.statoSdI),
    incassoAruba: col(H.incassoAruba),
    dataIncasso: col(H.dataIncasso),
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
    const segnoNc = isNotaCredito(String(cell(r, idx.tipoDocumento) ?? "")) ? -1 : 1;
    const val = (n: number) => segnoNc * Math.abs(n);
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
      imponibile: val(cellToImporto(cell(r, idx.imponibile)) ?? 0),
      iva: val(cellToImporto(cell(r, idx.iva)) ?? 0),
      totale: val(totale),
      netto: val(cellToImporto(cell(r, idx.netto)) ?? totale),
      // Lo "stato" dell'export ricevute non e' lo stato SdI delle emesse
      // (scartata/consegnata): non va confuso con gli scarti.
      statoSdI: ricevute ? "" : String(cell(r, idx.statoSdI) ?? "").trim(),
      direzione: ricevute ? "Ricevuta" : "Emessa",
      incassoAruba: parseIncassoAruba(cell(r, idx.incassoAruba)) || undefined,
      dataIncasso: cellToIsoDate(cell(r, idx.dataIncasso)) ?? undefined,
    });
  }
  return { rows, scartate };
}

// --- Spiegazione dei bonifici (compensazioni incluse) ------------------------
// Certi clienti (iMile su tutti) pagano con bonifici CUMULATIVI al netto di
// note di credito, trattenute e perfino delle fatture che LORO emettono a noi:
// l'importo del bonifico non coincide mai con una fattura e la riconciliazione
// classica non aggancia nulla. Qui si cerca, per ogni movimento non attribuito,
// la combinazione di documenti aperti che lo spiega al centesimo — fatture in
// positivo, storni in negativo. La ricerca e' conservativa: si accetta solo la
// combinazione esatta (tolleranza da arrotondamento), mai un'imputazione
// parziale, e la spiegazione resta una PROPOSTA finche' l'utente non la applica.

/** Un documento che puo' comporre un bonifico. */
export interface PezzoSpiegazione {
  /** nomeFile del documento (chiave per salvare l'abbinamento). */
  chiave: string;
  /** Numero leggibile ("FPR 63/26"). */
  etichetta: string;
  /** FT = fattura nostra (positivo), NC = nota di credito non collegata
   *  (negativo), CONTRO = fattura della controparte verso di noi (negativo,
   *  compensata dentro il bonifico). */
  tipo: "FT" | "NC" | "CONTRO";
  /** Importo con cui il documento entra nel bonifico (segno incluso). */
  valore: number;
  /** Totale documento, per mostrare l'eventuale trattenuta. */
  totale: number;
  data: string; // YYYY-MM-DD
}

export interface SpiegazioneBonifico {
  movimentoChiave: string;
  data: string;
  importo: number;
  /** Vuoto = nessuna combinazione trovata per questo movimento. */
  pezzi: PezzoSpiegazione[];
  /** importo - somma dei pezzi (residui di arrotondamento, entro tolleranza). */
  delta: number;
}

export const SPIEGA_TOLLERANZA = 0.15;
const SPIEGA_MAX_PEZZI = 7;

export function spiegaBonifici(
  movimenti: readonly { chiave: string; dataContabile: string; importo: number }[],
  pezzi: readonly PezzoSpiegazione[],
  toll = SPIEGA_TOLLERANZA,
): SpiegazioneBonifico[] {
  const giorni = (a: string, b: string) =>
    Math.abs(new Date(`${a}T00:00:00`).getTime() - new Date(`${b}T00:00:00`).getTime()) / 86400000;
  const liberi = pezzi.map((p) => ({ ...p, usato: false }));
  const out: SpiegazioneBonifico[] = [];
  // In ordine cronologico: ogni documento spiega UN solo bonifico, e i
  // bonifici vecchi devono potersi prendere i documenti vecchi.
  const ordinati = [...movimenti].sort((a, b) => a.dataContabile.localeCompare(b.dataContabile));
  for (const m of ordinati) {
    const target = Math.round(Math.abs(m.importo) * 100) / 100;
    // Candidati: documenti liberi emessi PRIMA dell'arrivo del bonifico, i
    // piu' vicini nel tempo per primi (un bonifico paga le fatture recenti).
    const cand = liberi
      .filter((p) => !p.usato && p.data <= m.dataContabile)
      .sort((a, b) => giorni(a.data, m.dataContabile) - giorni(b.data, m.dataContabile));
    const negTot = cand.filter((p) => p.valore < 0).reduce((s, p) => s + p.valore, 0);
    let trovato: typeof cand | null = null;
    // Iterative deepening: prima la spiegazione piu' semplice (meno pezzi).
    for (let maxN = 1; maxN <= SPIEGA_MAX_PEZZI && !trovato; maxN++) {
      const dfs = (idx: number, resto: number, presi: typeof cand, negResiduo: number) => {
        if (trovato) return;
        if (presi.length > 0 && Math.abs(resto) <= toll && presi.some((p) => p.valore > 0)) {
          trovato = [...presi];
          return;
        }
        if (idx >= cand.length || presi.length >= maxN) return;
        for (let j = idx; j < cand.length; j++) {
          const p = cand[j];
          const nuovoResto = Math.round((resto - p.valore) * 100) / 100;
          const negDopo = p.valore < 0 ? negResiduo - p.valore : negResiduo;
          // Uno sforamento in negativo e' lecito solo se gli storni ancora
          // disponibili possono riassorbirlo (e' cosi' che si spiegano i
          // bonifici "fattura meno nota di credito").
          if (nuovoResto < negDopo - toll) continue;
          dfs(j + 1, nuovoResto, [...presi, p], negDopo);
          if (trovato) return;
        }
      };
      dfs(0, target, [], negTot);
    }
    const scelti: PezzoSpiegazione[] = trovato ?? [];
    for (const p of scelti) {
      const l = liberi.find((x) => x.chiave === p.chiave && !x.usato);
      if (l) l.usato = true;
    }
    const somma = scelti.reduce((s, p) => s + p.valore, 0);
    out.push({
      movimentoChiave: m.chiave,
      data: m.dataContabile,
      importo: target,
      pezzi: scelti,
      delta: scelti.length ? Math.round((target - somma) * 100) / 100 : 0,
    });
  }
  return out;
}
