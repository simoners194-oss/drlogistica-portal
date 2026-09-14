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
  /** "" quando il file non dichiara il mese (tracciato per-appalto). */
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

/** Descrizione indirizzamento → codice, per i fogli SENZA la colonna codice
 *  (tracciato per-appalto di gennaio–maggio 2026: un foglio per appalto). */
function codiceDaDescrizione(desc: string): string {
  const d = norm(desc);
  if (d.startsWith("costi ord")) return "1";
  if (d.startsWith("costi str")) return "2";
  if (d.startsWith("ferie")) return "910";
  if (d.startsWith("mens")) return "920";
  if (d.startsWith("t.f.r") || d === "tfr") return "930";
  if (d.startsWith("totalizzazioni")) return "0";
  return "";
}

/** Riconosce e interpreta un foglio costi; null se il foglio non è nel
 *  tracciato (si prova il successivo). Due varianti: export paghe completo
 *  (con "Codice indirizzamento", "Codice dipendente", "Mese da") e foglio
 *  per-appalto (solo "Descrizione indirizzamento", niente codici né mese —
 *  l'appalto è `etichettaDefault`, cioè il nome del foglio). */
export function parseCostiMese(
  matrix: unknown[][],
  etichettaDefault = "",
): ParseCostiResult | null {
  const headerIdx = matrix.findIndex((r) => {
    const cells = (r ?? []).map(norm);
    return (
      cells.includes("totale costo") &&
      (cells.includes("codice indirizzamento") ||
        cells.some((c) => c.startsWith("descrizione indirizzament")))
    );
  });
  if (headerIdx < 0) return null;
  const header = (matrix[headerIdx] ?? []).map(norm);
  const col = (nome: string) => header.indexOf(nome);
  const C = {
    codice: col("codice dipendente"),
    cognome: col("cognome"),
    nome: col("nome"),
    indirizzamento: col("codice indirizzamento"),
    indirizzamentoDesc: header.findIndex((c) => c.startsWith("descrizione indirizzament")),
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
  if (C.cognome < 0 || C.totCosto < 0) return null;
  if (C.indirizzamento < 0 && C.indirizzamentoDesc < 0) return null;
  // La colonna etichetta (appalto) esiste solo nel tracciato completo, dove
  // la prima colonna NON è il cognome.
  const haEtichetta = C.cognome > 0;

  const perDip = new Map<string, StipendioDipendente>();
  const mesi = new Map<string, number>();
  let scartate = 0;
  for (const r of matrix.slice(headerIdx + 1)) {
    const cognome = String(r?.[C.cognome] ?? "").trim();
    const nome = C.nome >= 0 ? String(r?.[C.nome] ?? "").trim() : "";
    const codice = C.codice >= 0 ? String(r?.[C.codice] ?? "").trim() : "";
    // Nel tracciato completo la chiave è il codice dipendente; in quello
    // per-appalto (senza codici) vale cognome+nome.
    const chiave = codice || (cognome ? `${cognome}|${nome}` : "");
    if (!chiave) {
      if ((r ?? []).some((c) => c != null && String(c).trim() !== "")) scartate++;
      continue;
    }
    let d = perDip.get(chiave);
    if (!d) {
      d = {
        codice,
        cognome,
        nome,
        etichetta: haEtichetta ? String(r[0] ?? "").trim() : etichettaDefault,
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
      perDip.set(chiave, d);
    }
    const componente =
      numCell(r[C.retrib]) +
      (C.contrib >= 0 ? numCell(r[C.contrib]) : 0) +
      (C.inail >= 0 ? numCell(r[C.inail]) : 0);
    const ind =
      C.indirizzamento >= 0
        ? String(r[C.indirizzamento] ?? "").trim()
        : codiceDaDescrizione(String(r[C.indirizzamentoDesc] ?? ""));
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

  const mese = [...mesi.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "";
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

const MESI_IT = [
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

/** "01 - Costi Personale Gennaio 2026.xlsx" → "2026-01" ("" se non riconosciuto). */
export function meseDaNomeFile(nomeFile: string): string {
  const s = norm(nomeFile);
  for (let i = 0; i < 12; i++) {
    const m = s.match(new RegExp(`${MESI_IT[i]}[^0-9]*([0-9]{4})`));
    if (m) return `${m[1]}-${String(i + 1).padStart(2, "0")}`;
  }
  return "";
}

export interface ParseCostiFileResult extends ParseCostiResult {
  /** Fogli del file effettivamente letti (per l'anteprima). */
  fogli: string[];
}

/** Interpreta un intero file costi: prova OGNI foglio e unisce quelli validi.
 *  Serve per il tracciato per-appalto (gennaio–maggio 2026: un foglio per
 *  appalto, l'appalto è il nome del foglio); per l'export paghe completo
 *  passa di solito un solo foglio. Il mese viene dal contenuto quando c'è,
 *  altrimenti dal nome del file ("Costi Personale Gennaio 2026"). */
export function parseCostiFile(
  fogli: { nome: string; matrix: unknown[][] }[],
  nomeFile: string,
): ParseCostiFileResult | null {
  const parziali: { nome: string; res: ParseCostiResult }[] = [];
  for (const f of fogli) {
    const res = parseCostiMese(f.matrix, f.nome);
    if (res) parziali.push({ nome: f.nome, res });
  }
  if (parziali.length === 0) return null;
  const dipendenti = parziali.flatMap((p) => p.res.dipendenti);
  const mese = parziali.map((p) => p.res.mese).find((m) => m) || meseDaNomeFile(nomeFile) || "";
  const totale = Math.round(dipendenti.reduce((a, d) => a + d.totaleCosto, 0) * 100) / 100;
  return {
    mese,
    dipendenti,
    totale,
    squadrature: parziali.reduce((a, p) => a + p.res.squadrature, 0),
    scartate: parziali.reduce((a, p) => a + p.res.scartate, 0),
    fogli: parziali.map((p) => p.nome),
  };
}
