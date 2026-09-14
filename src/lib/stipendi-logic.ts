// DR Portal — Finanza → Stipendi: logica pura di parsing del file paghe.
// -----------------------------------------------------------------------------
// Fonte: OneDrive "Personale\MENSILITA'\<MESE ANNO>\COSTI <MESE> <ANNO>.xlsx",
// export del consulente paghe: una riga per dipendente per "codice
// indirizzamento" (1 = costi ordinari, 2 = straordinari, 910 = ferie/permessi,
// 920 = mensilità aggiuntive, 930 = TFR, 0 = totalizzazioni). La riga 0 porta
// "Totale costo", "Totale ore" e "Costo medio". La PRIMA colonna cambia nome
// da un mese all'altro (APPALTO, Nome Cognome…), quindi le colonne si
// riconoscono per INTESTAZIONE, mai per posizione.

export interface StipendioDipendente {
  codice: string;
  cognome: string;
  nome: string;
  /** Prima colonna del file (appalto o etichetta libera del consulente). */
  etichetta: string;
  /** "Descrizione ripartizione 1" (es. filiale), quando presente. */
  ripartizione?: string;
  oreOrdinarie: number;
  oreStraordinarie: number;
  costoOrdinario: number; // codice 1 (retribuzione + contributi + INAIL)
  costoStraordinario: number; // codice 2
  feriePermessi: number; // codice 910
  mensilitaAggiuntive: number; // codice 920
  tfr: number; // codice 930
  totaleCosto: number; // riga 0 — la colonna "Totale costo" chiesta da Simone
  totaleOre: number;
  costoMedio: number;
}

export interface StipendiMese {
  mese: string; // YYYY-MM (competenza)
  fonteFile: string;
  caricatoIl?: string;
  caricatoDa?: string;
  dipendenti: StipendioDipendente[];
}

export interface StipendiDb {
  versione: number;
  mesi: StipendiMese[];
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export function emptyStipendiDb(): StipendiDb {
  return { versione: 0, mesi: [] };
}

export function totaleMese(m: StipendiMese): number {
  return Math.round(m.dipendenti.reduce((a, d) => a + d.totaleCosto, 0) * 100) / 100;
}

export interface ParseCostiResult {
  mese: string;
  dipendenti: StipendioDipendente[];
  totale: number;
  /** Dipendenti la cui somma componenti non torna col totale (oltre 1 €). */
  squadrature: number;
  scartate: number;
}

const norm = (s: unknown) =>
  String(s ?? "")
    .trim()
    .toLowerCase();

function numCell(v: unknown): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  let s = String(v ?? "").trim();
  if (!s) return 0;
  // "1.234,56" (formato italiano) → il punto è separatore di migliaia SOLO
  // se c'è anche la virgola; "633.43" resta un decimale.
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

/** Riconosce e interpreta un foglio "COSTI <mese>"; null se il foglio non è
 *  nel tracciato (si prova il successivo). */
export function parseCostiMese(matrix: unknown[][]): ParseCostiResult | null {
  const headerIdx = matrix.findIndex((r) => {
    const cells = (r ?? []).map(norm);
    return cells.includes("codice indirizzamento") && cells.includes("totale costo");
  });
  if (headerIdx < 0) return null;
  const header = (matrix[headerIdx] ?? []).map(norm);
  const col = (nome: string) => header.indexOf(nome);
  const C = {
    codice: col("codice dipendente"),
    cognome: col("cognome"),
    nome: col("nome"),
    indirizzamento: col("codice indirizzamento"),
    retrib: col("costo retribuzione"),
    contrib: col("costo contributivo"),
    inail: col("costo inail"),
    oreOrd: col("ore ordinarie"),
    oreStr: col("ore straordinarie"),
    totCosto: col("totale costo"),
    totOre: col("totale ore"),
    medio: col("costo medio"),
    meseDa: col("mese da"),
    annoDa: col("anno da"),
    ripart: col("descrizione ripartizione 1"),
  };
  if (C.codice < 0 || C.indirizzamento < 0 || C.totCosto < 0) return null;

  const perDip = new Map<string, StipendioDipendente>();
  const mesi = new Map<string, number>();
  let scartate = 0;
  for (const r of matrix.slice(headerIdx + 1)) {
    const codice = String(r?.[C.codice] ?? "").trim();
    if (!codice) {
      if ((r ?? []).some((c) => c != null && String(c).trim() !== "")) scartate++;
      continue;
    }
    let d = perDip.get(codice);
    if (!d) {
      d = {
        codice,
        cognome: String(r[C.cognome] ?? "").trim(),
        nome: String(r[C.nome] ?? "").trim(),
        etichetta: String(r[0] ?? "").trim(),
        ripartizione: C.ripart >= 0 ? String(r[C.ripart] ?? "").trim() || undefined : undefined,
        oreOrdinarie: 0,
        oreStraordinarie: 0,
        costoOrdinario: 0,
        costoStraordinario: 0,
        feriePermessi: 0,
        mensilitaAggiuntive: 0,
        tfr: 0,
        totaleCosto: 0,
        totaleOre: 0,
        costoMedio: 0,
      };
      perDip.set(codice, d);
    }
    const componente =
      numCell(r[C.retrib]) +
      (C.contrib >= 0 ? numCell(r[C.contrib]) : 0) +
      (C.inail >= 0 ? numCell(r[C.inail]) : 0);
    const ind = String(r[C.indirizzamento] ?? "").trim();
    if (ind === "1") {
      d.costoOrdinario += componente;
      if (C.oreOrd >= 0) d.oreOrdinarie += numCell(r[C.oreOrd]);
    } else if (ind === "2") {
      d.costoStraordinario += componente;
      if (C.oreStr >= 0) d.oreStraordinarie += numCell(r[C.oreStr]);
    } else if (ind === "910") d.feriePermessi += componente;
    else if (ind === "920") d.mensilitaAggiuntive += componente;
    else if (ind === "930") d.tfr += componente;
    else if (ind === "0") {
      d.totaleCosto += numCell(r[C.totCosto]);
      if (C.totOre >= 0) d.totaleOre += numCell(r[C.totOre]);
      if (C.medio >= 0) d.costoMedio = numCell(r[C.medio]);
    }
    if (C.meseDa >= 0 && C.annoDa >= 0) {
      const mm = String(r[C.meseDa] ?? "").trim();
      const aa = String(r[C.annoDa] ?? "").trim();
      if (/^\d{1,2}$/.test(mm) && /^\d{4}$/.test(aa)) {
        const k = `${aa}-${mm.padStart(2, "0")}`;
        mesi.set(k, (mesi.get(k) ?? 0) + 1);
      }
    }
  }
  if (perDip.size === 0) return null;

  const mese =
    [...mesi.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? new Date().toISOString().slice(0, 7);
  const dipendenti = [...perDip.values()].sort((a, b) =>
    `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`),
  );
  // Quadratura per dipendente: componenti vs riga Totalizzazioni.
  let squadrature = 0;
  for (const d of dipendenti) {
    const somma =
      d.costoOrdinario + d.costoStraordinario + d.feriePermessi + d.mensilitaAggiuntive + d.tfr;
    if (Math.abs(somma - d.totaleCosto) > 1) squadrature++;
    // Arrotondamenti a 2 decimali per non salvare code binarie.
    d.costoOrdinario = Math.round(d.costoOrdinario * 100) / 100;
    d.costoStraordinario = Math.round(d.costoStraordinario * 100) / 100;
    d.feriePermessi = Math.round(d.feriePermessi * 100) / 100;
    d.mensilitaAggiuntive = Math.round(d.mensilitaAggiuntive * 100) / 100;
    d.tfr = Math.round(d.tfr * 100) / 100;
    d.totaleCosto = Math.round(d.totaleCosto * 100) / 100;
    d.totaleOre = Math.round(d.totaleOre * 100) / 100;
  }
  const totale = Math.round(dipendenti.reduce((a, d) => a + d.totaleCosto, 0) * 100) / 100;
  return { mese, dipendenti, totale, squadrature, scartate };
}

/** Mese successivo in formato YYYY-MM (default per il pagamento stipendi). */
export function meseSuccessivo(yyyymm: string): string {
  const [y, m] = yyyymm.split("-").map(Number);
  const d = new Date(y, m, 1); // m è già il mese successivo (0-based)
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
