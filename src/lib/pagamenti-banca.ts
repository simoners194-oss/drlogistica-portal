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
  const A = paroleControparte(a);
  const B = paroleControparte(b);
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
