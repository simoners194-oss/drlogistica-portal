// DR Portal — Fatture ancora aperte su Aruba ma già passate in BANCA
// (Simone 12/11/2026: "Bin Consulenze è stato pagato, perché lo trovo ancora
// in cassa? … anche RF, controlla che non ce ne siano altre").
//
// Il portale chiude una fattura quando Aruba dice pagata/incassata. Tra il
// bonifico e la registrazione su Aruba passano ore o giorni, e in quel tempo
// la fattura conta due volte: il saldo in banca è già sceso (o salito) e la
// fattura è ancora tra le uscite (o le entrate). Qui si agganciano i
// movimenti di banca alle fatture aperte, con regole strette per non
// sbagliare:
//   - stessa controparte (parole del nome, tollerando i nomi troncati dalla
//     banca e la descrizione del bonifico);
//   - stesso importo al centesimo: residuo della fattura, oppure il netto a
//     pagare (ritenuta d'acconto);
//   - movimento NON prima della data fattura;
//   - un movimento vale per una cosa sola: prima si tolgono quelli già
//     spiegati dalle fatture che Aruba dà per pagate (stessa controparte,
//     stesso importo, data vicina a quella registrata).
// Oltre al caso semplice (un movimento, una fattura) si riconoscono il
// bonifico unico che paga più fatture e la fattura pagata in più bonifici.
//
// Logica pura: la usano la Cassa proprietà (per togliere dalle uscite quello
// che è già uscito) e la pagina "Da correggere su Aruba".
import {
  incassatoRegistrato,
  isNotaCredito,
  parseIncassoAruba,
  residuoAperto,
  type DirezioneFattura,
  type FatturaConStato,
} from "./fatture-logic";

export interface PagamentoBanca {
  id?: string;
  cliente?: string;
  descrizione?: string;
  importo: number;
  dataContabile: string;
}

export interface AggancioBanca {
  nomeFile: string;
  /** "singolo": un movimento per la fattura · "somma": un movimento che paga
   *  più fatture · "rate": la fattura pagata in più movimenti. */
  modo: "singolo" | "somma" | "rate";
  /** Data dell'ultimo movimento (YYYY-MM-DD): quella da registrare su Aruba. */
  data: string;
  movimenti: { data: string; importo: number }[];
  /** Numeri delle altre fatture pagate dallo stesso movimento ("somma"). */
  insieme?: string[];
}

const FORME = new Set([
  "srl",
  "srls",
  "spa",
  "snc",
  "sas",
  "scarl",
  "scpa",
  "soc",
  "societa",
  "coop",
  "cooperativa",
  "ditta",
  "benefit",
]);
const VUOTE = new Set([
  "del",
  "dei",
  "della",
  "delle",
  "degli",
  "con",
  "per",
  "the",
  "and",
  "italia",
  "italy",
  "italiana",
  "group",
  "gruppo",
]);

/** Le parole che contano nel nome di una controparte: minuscole, senza
 *  accenti, senza forma societaria né parole vuote, almeno tre lettere. */
export function paroleControparte(nome: string): string[] {
  return (nome ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .split(" ")
    .filter((w) => w.length > 2 && !FORME.has(w) && !VUOTE.has(w));
}

/** Vero se i due nomi indicano la stessa controparte. Il più corto dei due
 *  deve stare (quasi) tutto dentro l'altro; la banca tronca i nomi, quindi
 *  l'ultima parola può essere mozza. Un nome di una parola sola vale solo se
 *  è lunga almeno cinque lettere ("polidir", non "bin"). */
export function stessaControparte(a: string, b: string): boolean {
  return stesseParole(paroleControparte(a), paroleControparte(b));
}

/** Come `stessaControparte`, sulle parole già estratte (per i confronti in
 *  serie: le parole di un nome si calcolano una volta sola). */
export function stesseParole(A: readonly string[], B: readonly string[]): boolean {
  if (!A.length || !B.length) return false;
  const [corto, lungo] = A.length <= B.length ? [A, B] : [B, A];
  const c = [...new Set(corto)];
  const trovata = (w: string, ultima: boolean) =>
    lungo.some((x) => x === w || (ultima && w.length >= 4 && x.startsWith(w)));
  const comuni = c.filter((w, i) => trovata(w, i === c.length - 1)).length;
  if (c.length === 1) return comuni === 1 && c[0].length >= 5;
  return comuni >= 2 && comuni / c.length >= 0.6;
}

interface Mov {
  data: string;
  cent: number;
  nome: string;
  desc: string;
  usato: boolean;
}

const cent = (n: number) => Math.round(Math.abs(n) * 100);
const giorniTra = (a: string, b: string) =>
  Math.abs(Date.parse(`${a}T00:00:00Z`) - Date.parse(`${b}T00:00:00Z`)) / 86400000;
/** Tolleranza sull'importo: cinque centesimi (arrotondamenti). */
const TOLL = 5;

/** Aggancia i movimenti di banca alle fatture ancora aperte.
 *  @param fatture   tutte le fatture di quella direzione, pagate comprese
 *                   (servono a "consumare" i movimenti già spiegati)
 *  @param movimenti tutti i movimenti di banca
 *  @param direzione "Ricevuta" = pagamenti (uscite), "Emessa" = incassi */
export function agganciaPagamenti(
  fatture: readonly FatturaConStato[],
  movimenti: readonly PagamentoBanca[],
  direzione: DirezioneFattura,
): Map<string, AggancioBanca> {
  const out = new Map<string, AggancioBanca>();
  const uscite = direzione === "Ricevuta";
  const movs: Mov[] = movimenti
    .filter((m) => (uscite ? m.importo < -0.005 : m.importo > 0.005))
    .map((m) => ({
      data: (m.dataContabile ?? "").slice(0, 10),
      cent: cent(m.importo),
      nome: m.cliente ?? "",
      desc: m.descrizione ?? "",
      usato: false,
    }))
    .filter((m) => /^\d{4}-\d{2}-\d{2}$/.test(m.data))
    .sort((a, b) => a.data.localeCompare(b.data));
  if (!movs.length) return out;
  const perCent = new Map<number, Mov[]>();
  for (const m of movs) {
    const l = perCent.get(m.cent) ?? [];
    l.push(m);
    perCent.set(m.cent, l);
  }
  const conImporto = (c: number): Mov[] => {
    const l: Mov[] = [];
    for (let d = -TOLL; d <= TOLL; d++) for (const m of perCent.get(c + d) ?? []) l.push(m);
    return l;
  };
  // Stessa controparte: memoria per (movimento, nome) — i confronti sono tanti.
  const memo = new Map<string, boolean>();
  const combacia = (m: Mov, nome: string) => {
    const k = `${m.nome}\u0001${m.desc}\u0001${nome}`;
    let v = memo.get(k);
    if (v == null) {
      v = stessaControparte(m.nome, nome) || stessaControparte(m.desc, nome);
      memo.set(k, v);
    }
    return v;
  };

  const vere = fatture.filter((x) => !isNotaCredito(x.f.tipoDocumento) && x.f.totale > 0);
  // 1) I movimenti già spiegati dalle fatture che Aruba dà per pagate.
  for (const x of vere) {
    if (x.f.pagataCdc) continue; // pagata con carta: nessun bonifico suo
    if (residuoAperto(x) > 1) continue;
    const stato = parseIncassoAruba(x.f.incassoAruba);
    if (stato !== "Incassata" && x.s.statoFatturazione !== "Pagata") continue;
    const importi = [
      ...new Set([cent(x.f.incassatoAruba ?? x.f.totale), cent(x.f.totale), cent(x.f.netto)]),
    ];
    const di = x.f.dataIncasso ?? "";
    let scelto: Mov | undefined;
    for (const c of importi) {
      if (!c) continue;
      const candidati = conImporto(c).filter(
        (m) =>
          !m.usato &&
          combacia(m, x.f.cliente) &&
          (di ? giorniTra(m.data, di) <= 5 : m.data >= x.f.dataDocumento),
      );
      if (!candidati.length) continue;
      scelto = di
        ? candidati.sort((a, b) => giorniTra(a.data, di) - giorniTra(b.data, di))[0]
        : candidati[0];
      break;
    }
    if (scelto) scelto.usato = true;
  }

  // 2) Fatture aperte, dalla più vecchia.
  const aperte = vere
    .filter((x) => !x.f.pagataCdc && residuoAperto(x) > 1)
    .sort(
      (a, b) =>
        a.f.dataDocumento.localeCompare(b.f.dataDocumento) || a.f.numero.localeCompare(b.f.numero),
    );
  const residuo = new Map<string, number>(
    aperte.map((x) => [x.f.nomeFile, cent(residuoAperto(x))]),
  );
  // 2a) un movimento, una fattura
  for (const x of aperte) {
    const res = residuo.get(x.f.nomeFile) ?? 0;
    const importi = [...new Set([res, cent(x.f.netto)])].filter((c) => c > 100);
    for (const c of importi) {
      const m = conImporto(c).find(
        (v) => !v.usato && v.data >= x.f.dataDocumento && combacia(v, x.f.cliente),
      );
      if (!m) continue;
      m.usato = true;
      out.set(x.f.nomeFile, {
        nomeFile: x.f.nomeFile,
        modo: "singolo",
        data: m.data,
        movimenti: [{ data: m.data, importo: m.cent / 100 }],
      });
      break;
    }
  }
  // 2b) un movimento che paga più fatture della stessa controparte
  const perControparte = new Map<string, FatturaConStato[]>();
  for (const x of aperte) {
    if (out.has(x.f.nomeFile)) continue;
    const l = perControparte.get(x.f.cliente) ?? [];
    l.push(x);
    perControparte.set(x.f.cliente, l);
  }
  for (const [nome, lista] of perControparte) {
    if (lista.length < 2) continue;
    // Solo le più vecchie se sono tante: i sottoinsiemi raddoppiano a ogni fattura.
    let libere = lista.slice(0, 12);
    const prima = libere[0].f.dataDocumento;
    const candidati = movs.filter((m) => !m.usato && m.data >= prima && combacia(m, nome));
    for (const m of candidati) {
      const possibili = libere.filter((x) => x.f.dataDocumento <= m.data);
      const n = possibili.length;
      if (n < 2) continue;
      const importi = possibili.map((x) => residuo.get(x.f.nomeFile) ?? 0);
      if (importi.reduce((s, v) => s + v, 0) < m.cent - TOLL) continue;
      let trovato = 0;
      for (let mask = 3; mask < 1 << n && !trovato; mask++) {
        if ((mask & (mask - 1)) === 0) continue; // una fattura sola: caso 2a
        let s = 0;
        for (let i = 0; i < n; i++) if (mask & (1 << i)) s += importi[i];
        if (Math.abs(s - m.cent) <= TOLL) trovato = mask;
      }
      if (!trovato) continue;
      m.usato = true;
      const scelte = possibili.filter((_, i) => trovato & (1 << i));
      for (const x of scelte)
        out.set(x.f.nomeFile, {
          nomeFile: x.f.nomeFile,
          modo: "somma",
          data: m.data,
          movimenti: [{ data: m.data, importo: m.cent / 100 }],
          insieme: scelte.filter((y) => y !== x).map((y) => y.f.numero),
        });
      libere = libere.filter((x) => !out.has(x.f.nomeFile));
      if (libere.length < 2) break;
    }
  }
  // 2c) una fattura pagata in due o tre movimenti vicini
  for (const x of aperte) {
    if (out.has(x.f.nomeFile)) continue;
    const res = residuo.get(x.f.nomeFile) ?? 0;
    const candidati = movs
      .filter(
        (m) => !m.usato && m.cent < res && m.data >= x.f.dataDocumento && combacia(m, x.f.cliente),
      )
      .slice(-14);
    let scelti: Mov[] | null = null;
    for (let i = 0; i < candidati.length && !scelti; i++)
      for (let j = i + 1; j < candidati.length && !scelti; j++) {
        const a = candidati[i];
        const b = candidati[j];
        if (giorniTra(a.data, b.data) > 15) continue;
        if (Math.abs(a.cent + b.cent - res) <= TOLL) scelti = [a, b];
        for (let k = j + 1; k < candidati.length && !scelti; k++) {
          const c = candidati[k];
          if (giorniTra(a.data, c.data) > 15) continue;
          if (Math.abs(a.cent + b.cent + c.cent - res) <= TOLL) scelti = [a, b, c];
        }
      }
    if (!scelti) continue;
    for (const m of scelti) m.usato = true;
    out.set(x.f.nomeFile, {
      nomeFile: x.f.nomeFile,
      modo: "rate",
      data: scelti.map((m) => m.data).sort()[scelti.length - 1],
      movimenti: scelti.map((m) => ({ data: m.data, importo: m.cent / 100 })),
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// INCASSI ARRIVATI IN BANCA E NON ANCORA REGISTRATI SU ARUBA (1.121.0, Simone
// 18/11/2026: "prova a vedere se lo riconosce dai movimenti di banca").
//
// `agganciaPagamenti` riconosce una fattura solo quando il movimento ha il suo
// importo al centesimo. Con i clienti non basta: Number 1 paga due fatture
// meno una sua (19.316,07 per 19.354,63), Univex paga "la competenza di
// agosto" al netto delle compensazioni (69.192,30 per 74.853,10), Posta Doc
// manda acconti tondi. Qui si ragiona sull'accredito:
//   1. a chi appartiene: la controparte delle fatture che somiglia di più al
//      nome del movimento (o alla descrizione); se sono due alla pari, niente;
//   2. quanto di quell'accredito Aruba conosce già: gli incassi registrati
//      per quella controparte da cinque giorni prima a trenta dopo (chi
//      registra può scrivere la data del giorno in cui registra);
//   3. il resto è arrivato ma Aruba non lo sa: si scala dalle fatture aperte.
//      Prima quelle con lo stesso importo al centesimo, poi quelle citate
//      nella causale (numero, o "competenza agosto 2026"), poi dalla più
//      vecchia. Se la causale cita fatture che Aruba ha già chiuso,
//      l'accredito è spiegato e non si scala nulla.
// PRUDENZA SUL NOME: scalare "dalla più vecchia" è il criterio più debole
// (nessun importo uguale, nessuna fattura citata), quindi vale solo quando il
// nome del movimento combacia PER INTERO con la controparte ("postadoc srl",
// non un'etichetta corta come "number"). Lo stesso per gli accordi.
// Si guardano solo gli accrediti degli ultimi 45 giorni: più indietro è una
// storia da sistemare su Aruba, non una cosa da togliere dalla cassa.
// Logica pura; la Cassa proprietà la usa per le entrate attese e gli accordi.
// ---------------------------------------------------------------------------

export const GIORNI_INCASSI_BANCA = 45;
/** Aruba può portare una data diversa da quella della banca: questi giorni
 *  prima e questi dopo l'accredito. */
const REG_GIORNI_PRIMA = 5;
const REG_GIORNI_DOPO = 30;

export interface IncassoInBanca {
  /** Chiave stabile dell'accredito: id del movimento, oppure data|importo|nome. */
  chiave: string;
  /** La controparte, col nome che ha sulle fatture. */
  cliente: string;
  data: string;
  importo: number;
  /** Parte che Aruba ha già registrato negli stessi giorni. */
  registrato: number;
  /** Parte scalata dalle fatture ancora aperte. */
  scalato: number;
  /** Parte che non ha trovato fatture aperte da cui scalarsi. */
  avanzo: number;
  /** Come sono state scelte le fatture: stesso importo, numero citato in
   *  causale, mese di competenza citato, oppure dalla più vecchia. */
  criterio: "importo" | "numero" | "competenza" | "vecchie";
  fatture: { nomeFile: string; numero: string; importo: number; resta: number }[];
  descrizione: string;
}

export interface AccreditoCliente {
  chiave: string;
  cliente: string;
  data: string;
  importo: number;
  /** Il nome del movimento combacia per intero con la controparte: solo
   *  questi accrediti scalano un accordo. */
  nomeCerto: boolean;
}

export interface EsitoIncassiBanca {
  /** Gli accrediti con una parte non ancora registrata su Aruba. */
  incassi: IncassoInBanca[];
  /** Per fattura aperta (nomeFile): quanto è già arrivato in banca. */
  perFattura: Map<string, number>;
  /** Tutti gli accrediti del periodo con la loro controparte (registrati o
   *  no): servono a scalare gli accordi. */
  accrediti: AccreditoCliente[];
}

const MESI_CAUSALE = [
  "gennaio",
  "febbraio",
  "marzo",
  "aprile",
  "maggio",
  "giugno",
  "luglio",
  "agosto",
  "settembre",
  "ottobre",
  "novembre",
  "dicembre",
];

/** "Competenza Agosto 2026" in una causale → "2026-08". Senza anno vale
 *  l'ultimo mese con quel nome non dopo la data del movimento. */
export function competenzaInCausale(descrizione: string, data: string): string | null {
  const m = (descrizione ?? "")
    .toLowerCase()
    .match(
      /competenz[ae]\s+(?:di\s+|del\s+mese\s+di\s+|mese\s+di\s+|mese\s+)?(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\.?\s*(\d{4}|\d{2})?/,
    );
  if (!m) return null;
  const mese = MESI_CAUSALE.indexOf(m[1]) + 1;
  const annoMov = Number(data.slice(0, 4));
  const meseMov = Number(data.slice(5, 7));
  let anno = m[2] ? Number(m[2].length === 2 ? `20${m[2]}` : m[2]) : annoMov;
  if (!m[2] && mese > meseMov) anno = annoMov - 1;
  return `${anno}-${String(mese).padStart(2, "0")}`;
}

/** I numeri di fattura citati in una causale ("FPR 264/26", "ft 12/2026"):
 *  "264/26". Le date scritte coi punti ("1.10.2026") o con due barre
 *  ("25/09/2026") non contano. */
export function numeriInCausale(descrizione: string): Set<string> {
  const out = new Set<string>();
  const re = /(^|[^0-9/.])(\d{1,5})\s*\/\s*(\d{4}|\d{2})(?![0-9/])/g;
  for (const m of (descrizione ?? "").matchAll(re)) {
    const n = String(Number(m[2]));
    if (n !== "0") out.add(`${n}/${m[3].slice(-2)}`);
  }
  return out;
}

const numeroBreve = (numero: string): string | null => {
  const m = (numero ?? "").match(/(\d+)\s*\/\s*(\d{2,4})/);
  return m ? `${String(Number(m[1]))}/${m[2].slice(-2)}` : null;
};
const piuGiorniIso = (iso: string, g: number) =>
  new Date(Date.parse(`${iso}T00:00:00Z`) + g * 86400000).toISOString().slice(0, 10);
/** Giorni da `da` ad `a` (negativo se `a` viene prima). */
const giorniDa = (da: string, a: string) =>
  (Date.parse(`${a}T00:00:00Z`) - Date.parse(`${da}T00:00:00Z`)) / 86400000;

export function incassiInBanca(
  attive: readonly FatturaConStato[],
  movimenti: readonly PagamentoBanca[],
  oggi: string,
  opzioni: {
    /** Quanti giorni indietro guardare (45). */
    giorni?: number;
    /** Mese di competenza di una fattura ("YYYY-MM"); senza, il mese della data. */
    competenza?: (x: FatturaConStato) => string;
    /** Accrediti rifiutati a mano (chiavi): non si scalano e non contano. */
    rifiutati?: ReadonlySet<string>;
  } = {},
): EsitoIncassiBanca {
  const esito: EsitoIncassiBanca = { incassi: [], perFattura: new Map(), accrediti: [] };
  const dal = piuGiorniIso(oggi, -(opzioni.giorni ?? GIORNI_INCASSI_BANCA));
  const competenza = opzioni.competenza ?? ((x: FatturaConStato) => x.f.dataDocumento.slice(0, 7));
  const vere = attive.filter((x) => !isNotaCredito(x.f.tipoDocumento) && x.f.totale > 0);
  if (!vere.length) return esito;

  // Le controparti, col nome delle fatture, e le loro parole.
  const perCliente = new Map<string, FatturaConStato[]>();
  for (const x of vere) {
    const l = perCliente.get(x.f.cliente) ?? [];
    l.push(x);
    perCliente.set(x.f.cliente, l);
  }
  const clienti = [...perCliente.keys()].map((nome) => ({
    nome,
    parole: [...new Set(paroleControparte(nome))],
  }));

  // 1) A CHI APPARTIENE OGNI ACCREDITO.
  type Appartenenza = { nome: string; certo: boolean } | null;
  const memoDiChi = new Map<string, Appartenenza>();
  const diChi = (nomeMov: string, desc: string): Appartenenza => {
    const k = `${nomeMov}\u0001${desc}`;
    if (memoDiChi.has(k)) return memoDiChi.get(k) ?? null;
    let trovato: Appartenenza = null;
    for (const testo of [nomeMov, desc]) {
      const P = paroleControparte(testo);
      if (!P.length) continue;
      const candidati = clienti
        .filter((c) => stesseParole(P, c.parole))
        .map((c) => {
          const comuni = c.parole.filter((w) =>
            P.some((p) => p === w || (w.length >= 4 && p.startsWith(w))),
          ).length;
          return { nome: c.nome, comuni, fuori: c.parole.length - comuni };
        })
        .sort((a, b) => b.comuni - a.comuni || a.fuori - b.fuori);
      if (!candidati.length) continue;
      // Due controparti alla pari ("Univex" tra Univex Freight e Univex
      // Healthcare): non si sa di chi è, non si tocca nulla.
      const [primo, secondo] = candidati;
      trovato =
        secondo && secondo.comuni === primo.comuni && secondo.fuori === primo.fuori
          ? null
          : { nome: primo.nome, certo: primo.fuori === 0 };
      break;
    }
    memoDiChi.set(k, trovato);
    return trovato;
  };
  const accreditiDi = new Map<
    string,
    { chiave: string; data: string; cent: number; desc: string; certo: boolean }[]
  >();
  const visti = new Map<string, number>();
  for (const m of movimenti) {
    if (!(m.importo > 0.005)) continue;
    const data = (m.dataContabile ?? "").slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || data < dal || data > oggi) continue;
    const di = diChi(m.cliente ?? "", m.descrizione ?? "");
    if (!di) continue;
    const nome = di.nome;
    let chiave = m.id ? String(m.id) : `${data}|${cent(m.importo)}|${nome.toLowerCase()}`;
    // Due accrediti gemelli senza id (stessa data, importo, cliente): il
    // secondo prende un suffisso, così la × ne rifiuta uno solo.
    const n = (visti.get(chiave) ?? 0) + 1;
    visti.set(chiave, n);
    if (n > 1) chiave = `${chiave}#${n}`;
    if (opzioni.rifiutati?.has(chiave)) continue;
    const l = accreditiDi.get(nome) ?? [];
    l.push({ chiave, data, cent: cent(m.importo), desc: m.descrizione ?? "", certo: di.certo });
    accreditiDi.set(nome, l);
    esito.accrediti.push({
      chiave,
      cliente: nome,
      data,
      importo: cent(m.importo) / 100,
      nomeCerto: di.certo,
    });
  }

  for (const [nome, accrediti] of accreditiDi) {
    const fatture = perCliente.get(nome) ?? [];
    // 2) Quello che Aruba ha già registrato: un evento per fattura incassata.
    const eventi = fatture
      .map((x) => ({
        data: (x.f.dataIncasso ?? "").slice(0, 10),
        resto: cent(Math.max(0, incassatoRegistrato(x))),
      }))
      .filter((e) => /^\d{4}-\d{2}-\d{2}$/.test(e.data) && e.resto > 0);
    // Le fatture ancora aperte, dalla più vecchia, col loro residuo.
    const aperte = fatture
      .map((x) => ({ x, resta: cent(residuoAperto(x)) }))
      .filter((a) => a.resta > 100)
      .sort(
        (a, b) =>
          a.x.f.dataDocumento.localeCompare(b.x.f.dataDocumento) ||
          a.x.f.numero.localeCompare(b.x.f.numero, "it", { numeric: true }),
      );
    accrediti.sort((a, b) => a.data.localeCompare(b.data) || a.chiave.localeCompare(b.chiave));

    for (const c of accrediti) {
      let bisogno = c.cent;
      // 2a) lo stesso importo registrato in un giorno solo.
      const perGiorno = new Map<string, typeof eventi>();
      for (const e of eventi) {
        if (e.resto <= 0) continue;
        const g = giorniDa(c.data, e.data);
        if (g < -REG_GIORNI_PRIMA || g > REG_GIORNI_DOPO) continue;
        perGiorno.set(e.data, [...(perGiorno.get(e.data) ?? []), e]);
      }
      const giornoUguale = [...perGiorno.values()].find(
        (l) => Math.abs(l.reduce((s, e) => s + e.resto, 0) - c.cent) <= TOLL,
      );
      if (giornoUguale) {
        for (const e of giornoUguale) e.resto = 0;
        bisogno = 0;
      } else {
        // 2b) gli incassi registrati più vicini, fino a coprire l'accredito.
        const vicini = [...perGiorno.values()]
          .flat()
          .sort(
            (a, b) =>
              Math.abs(giorniDa(c.data, a.data)) - Math.abs(giorniDa(c.data, b.data)) ||
              a.data.localeCompare(b.data),
          );
        for (const e of vicini) {
          if (bisogno <= 0) break;
          const preso = Math.min(e.resto, bisogno);
          e.resto -= preso;
          bisogno -= preso;
        }
      }
      // Differenze piccole (spese, arrotondamenti): l'accredito è registrato.
      if (bisogno <= Math.max(100, Math.round(c.cent * 0.01))) continue;

      // 3) Il resto non è su Aruba: da quali fatture aperte si scala.
      const candidate = aperte.filter((a) => a.resta > 0 && a.x.f.dataDocumento <= c.data);
      const citati = numeriInCausale(c.desc);
      const meseCitato = competenzaInCausale(c.desc, c.data);
      const eCitata = (x: FatturaConStato) => {
        const n = numeroBreve(x.f.numero);
        return n != null && citati.has(n);
      };
      // Il mese citato si cerca in due modi: la competenza che il portale dà
      // alla fattura, oppure il mese della sua data ("competenza agosto" =
      // fatture del 31/08).
      const delMese = (x: FatturaConStato, perData: boolean) =>
        meseCitato != null &&
        (perData ? x.f.dataDocumento.slice(0, 7) : competenza(x)) === meseCitato;
      // La causale cita fatture che Aruba ha già chiuso tutte: l'accredito è
      // spiegato da quelle (registrate con una data lontana).
      const tutteChiuse = (l: FatturaConStato[]) =>
        l.length > 0 && l.every((x) => residuoAperto(x) <= 1);
      if (
        tutteChiuse(fatture.filter(eCitata)) ||
        tutteChiuse(fatture.filter((x) => delMese(x, false))) ||
        tutteChiuse(fatture.filter((x) => delMese(x, true)))
      )
        continue;

      const scelte: { a: (typeof aperte)[number]; cent: number }[] = [];
      let criterio: IncassoInBanca["criterio"] = "vecchie";
      let daScalare = bisogno;
      // 3a) stesso importo: una fattura, o più fatture insieme.
      const una = candidate.find(
        (a) => Math.abs(a.resta - bisogno) <= TOLL || Math.abs(cent(a.x.f.netto) - bisogno) <= TOLL,
      );
      if (una) {
        scelte.push({ a: una, cent: una.resta });
        criterio = "importo";
        daScalare = 0;
      } else {
        const prime = candidate.slice(0, 12);
        let trovato = 0;
        if (prime.reduce((s, a) => s + a.resta, 0) >= bisogno - TOLL)
          for (let mask = 3; mask < 1 << prime.length && !trovato; mask++) {
            if ((mask & (mask - 1)) === 0) continue;
            let s = 0;
            for (let i = 0; i < prime.length; i++) if (mask & (1 << i)) s += prime[i].resta;
            if (Math.abs(s - bisogno) <= TOLL) trovato = mask;
          }
        if (trovato) {
          prime.forEach((a, i) => {
            if (trovato & (1 << i)) scelte.push({ a, cent: a.resta });
          });
          criterio = "importo";
          daScalare = 0;
        }
      }
      // 3b) le fatture citate (numero, poi competenza), poi dalla più vecchia.
      if (daScalare > 0) {
        const perNumero = candidate.filter((a) => eCitata(a.x));
        // Tra le due letture del mese vale quella in cui l'accredito ci sta
        // (un bonifico al netto delle compensazioni non supera le fatture).
        const somma = (l: typeof candidate) => l.reduce((s, a) => s + a.resta, 0);
        const letture = [false, true]
          .map((perData) =>
            candidate.filter((a) => !perNumero.includes(a) && delMese(a.x, perData)),
          )
          .filter((l) => l.length > 0);
        const perMese =
          letture
            .filter((l) => somma(l) >= daScalare - TOLL)
            .sort((a, b) => somma(a) - somma(b))[0] ??
          letture[0] ??
          [];
        const altre = candidate.filter((a) => !perNumero.includes(a) && !perMese.includes(a));
        if (perNumero.length) criterio = "numero";
        else if (perMese.length) criterio = "competenza";
        // Né importo, né numero, né mese: si scala dalla più vecchia solo se
        // il nome del movimento è per intero quello della controparte.
        else if (!c.certo) continue;
        for (const a of [...perNumero, ...perMese, ...altre]) {
          if (daScalare <= 0) break;
          const preso = Math.min(a.resta, daScalare);
          scelte.push({ a, cent: preso });
          daScalare -= preso;
        }
      }
      const scalato = scelte.reduce((s, v) => s + v.cent, 0);
      if (scalato <= 0) continue;
      for (const v of scelte) {
        v.a.resta -= v.cent;
        esito.perFattura.set(
          v.a.x.f.nomeFile,
          Math.round(((esito.perFattura.get(v.a.x.f.nomeFile) ?? 0) + v.cent / 100) * 100) / 100,
        );
      }
      esito.incassi.push({
        chiave: c.chiave,
        cliente: nome,
        data: c.data,
        importo: c.cent / 100,
        registrato: (c.cent - bisogno) / 100,
        scalato: scalato / 100,
        avanzo: Math.max(0, daScalare) / 100,
        criterio,
        fatture: scelte.map((v) => ({
          nomeFile: v.a.x.f.nomeFile,
          numero: v.a.x.f.numero,
          importo: v.cent / 100,
          resta: v.a.resta / 100,
        })),
        descrizione: c.desc,
      });
    }
  }
  esito.incassi.sort(
    (a, b) => b.data.localeCompare(a.data) || a.cliente.localeCompare(b.cliente, "it"),
  );
  return esito;
}
