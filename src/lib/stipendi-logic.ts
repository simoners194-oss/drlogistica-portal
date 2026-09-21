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

/** Riga del file "Stipendi Dr.xlsx" (Personale): il NETTO da bonificare.
 *  saldo = resto da pagare (stipendio − anticipi − addebiti); il file lo
 *  riporta esplicito nei mesi recenti, altrimenti si ricalcola. */
export interface NettoDipendente {
  nome: string; // "COGNOME NOME" come scritto nel file
  appalto?: string;
  stipendio: number;
  anticipo: number;
  addebito: number;
  saldo: number;
}

export interface NettiMese {
  mese: string; // YYYY-MM (competenza, dal nome del foglio)
  fonteFile: string;
  caricatoIl?: string;
  caricatoDa?: string;
  dipendenti: NettoDipendente[];
  totaleStipendio: number;
  totaleAnticipi: number;
  /** Somma dei saldi: quello che esce dal conto al pagamento (mese dopo). */
  totaleSaldo: number;
}

/** Riga della "Mappatura Dipendenti" (Personale\DIPENDENTI): anagrafica
 *  contrattuale per nome. */
export interface AnagraficaDipendente {
  nome: string; // come scritto nel file
  mansione?: string; // es. "AUTISTA G1", "IMPIEGATO 5"
  livello?: string; // estratto dalla mansione (G1, D2, 5…)
  /** "Indeterminato" | "Determinato" */
  contratto?: string;
  /** Per i determinati: la fine più avanzata tra data fine e proroghe. */
  fineContratto?: string; // ISO
  orario?: string; // FULL TIME / P.TIME 20H…
  stato?: string; // vuoto = in forza, "NON IN FORZA"…
  /** Mensilità dichiarate a mano (vincono sulla stima dai ratei paghe). */
  mensilita?: number;
}

/** Chiave di confronto nomi tra fonti diverse ("GABELLI SERVENTI DIEGO" vs
 *  cognome+nome dei file paghe): minuscole, solo lettere, parole ordinate. */
export function chiaveNome(s: string): string {
  return s
    .toLowerCase()
    .replace(/[^a-zà-ù ]/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .sort()
    .join(" ");
}

export interface StipendiDb {
  versione: number;
  mesi: StipendiMese[];
  /** Netti da "Stipendi Dr.xlsx" (un foglio per mese). */
  netti?: NettiMese[];
  /** Anagrafica contrattuale dalla "Mappatura Dipendenti". */
  anagrafica?: AnagraficaDipendente[];
  anagraficaFonte?: string;
  /** PAGATO SI/NO manuale (richiesta Simone 15/09): per mese di COMPETENZA,
   *  le chiavi nome (chiaveNome) dei dipendenti segnati come pagati. Guida
   *  la riga Stipendi dei Flussi: tabella reale = solo i sì, tabella "solo
   *  fatturazioni" = tutti. */
  pagatiPerMese?: Record<string, string[]>;
  /** Correzioni A MANO dei valori (richiesta Simone 21/09): vivono FUORI dai
   *  mesi importati, così un re-import del file non le cancella; la tabella
   *  mostra la "M" con chi/quando/prima. Vedi applicaModifiche. */
  modifiche?: ModificaManuale[];
  /** Perché un dipendente NON ha il netto nel mese (aspettativa, infortunio,
   *  cessato…) o compare nei netti senza costo (liquidazione, contanti):
   *  risposte HR del 21/09. Vive a parte come le modifiche: il re-import non
   *  la tocca. Vedi MOTIVI_NETTO. */
  motiviNetto?: MotivoNetto[];
  /** Esito dell'ultima lettura automatica dei file da SharePoint (1.79.0). */
  ultimaSync?: {
    il: string;
    da: string;
    nettiMesi: string[];
    costiMesi: string[];
    errori: string[];
  };
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export function emptyStipendiDb(): StipendiDb {
  return { versione: 0, mesi: [] };
}

// --- Modifiche manuali ---------------------------------------------------------
export const CAMPI_COSTI = [
  "costoOrdinario",
  "costoStraordinario",
  "feriePermessi",
  "mensilitaAggiuntive",
  "tfr",
  "totaleCosto",
] as const;
export const CAMPI_NETTI = ["stipendio", "anticipo", "saldo"] as const;
export const CAMPI_MODIFICABILI = [...CAMPI_COSTI, ...CAMPI_NETTI] as const;
export type CampoModificabile = (typeof CAMPI_MODIFICABILI)[number];

export interface ModificaManuale {
  mese: string; // YYYY-MM (competenza)
  /** chiaveNome("Cognome Nome"): aggancia sia la riga COSTI sia quella dei netti. */
  chiave: string;
  /** "Cognome Nome" leggibile: serve alle righe sintetiche dei netti. */
  nome: string;
  campo: CampoModificabile;
  valore: number;
  /** Valore che aveva il file al momento della prima modifica (null = non c'era). */
  prima: number | null;
  da: string;
  il: string; // ISO
}

export function chiaveModifica(mese: string, chiave: string, campo: string): string {
  return `${mese}|${chiave}|${campo}`;
}

// --- Motivo del netto assente / del costo assente -------------------------------
export const MOTIVI_NETTO = [
  "aspettativa",
  "infortunio",
  "malattia",
  "cessato",
  "fuori-file",
  "contanti",
  "altro",
] as const;
export type MotivoNettoTipo = (typeof MOTIVI_NETTO)[number];

export interface MotivoNetto {
  mese: string; // YYYY-MM (competenza)
  chiave: string; // chiaveNome("Cognome Nome")
  nome: string;
  motivo: MotivoNettoTipo;
  nota?: string;
  da: string;
  il: string; // ISO
}

export function chiaveMotivo(mese: string, chiave: string): string {
  return `${mese}|${chiave}`;
}

const r2 = (n: number) => Math.round(n * 100) / 100;

/** Applica le modifiche manuali a una copia del db: i mesi COSTI e i netti
 *  restano quelli del file, ma i valori corretti a mano vincono. Se un
 *  dipendente ha netti modificati ma NON compare nel foglio Stipendi Dr del
 *  mese (o il mese non è mai stato caricato), nasce una riga/un mese
 *  sintetici: così si riempie a mano anche un buco, non solo un errore. */
export function applicaModifiche(db: StipendiDb): StipendiDb {
  const mods = db.modifiche ?? [];
  if (mods.length === 0) return db;
  const perMese = new Map<string, ModificaManuale[]>();
  for (const m of mods) {
    if (!perMese.has(m.mese)) perMese.set(m.mese, []);
    perMese.get(m.mese)!.push(m);
  }
  const isCosto = (c: string) => (CAMPI_COSTI as readonly string[]).includes(c);
  const isNetto = (c: string) => (CAMPI_NETTI as readonly string[]).includes(c);

  const mesi = db.mesi.map((m) => {
    const ms = (perMese.get(m.mese) ?? []).filter((x) => isCosto(x.campo));
    if (ms.length === 0) return m;
    return {
      ...m,
      dipendenti: m.dipendenti.map((d) => {
        const k = chiaveNome(`${d.cognome} ${d.nome}`);
        const mie = ms.filter((x) => x.chiave === k);
        if (mie.length === 0) return d;
        const nd: StipendioDipendente = { ...d };
        for (const x of mie) (nd as unknown as Record<string, number>)[x.campo] = x.valore;
        return nd;
      }),
    };
  });

  const nettiMap = new Map<string, NettiMese>((db.netti ?? []).map((n) => [n.mese, n]));
  for (const [mese, lista] of perMese) {
    const ms = lista.filter((x) => isNetto(x.campo));
    if (ms.length === 0) continue;
    const base = nettiMap.get(mese) ?? {
      mese,
      fonteFile: "manuale",
      dipendenti: [],
      totaleStipendio: 0,
      totaleAnticipi: 0,
      totaleSaldo: 0,
    };
    const dipendenti = base.dipendenti.map((d) => ({ ...d }));
    const coperti = new Set<string>();
    for (const d of dipendenti) {
      const k = chiaveNome(d.nome);
      const mie = ms.filter((x) => x.chiave === k);
      if (mie.length === 0) continue;
      coperti.add(k);
      for (const x of mie) (d as unknown as Record<string, number>)[x.campo] = x.valore;
    }
    // Righe sintetiche per chi ha correzioni ma non sta nel foglio.
    const nuove = new Map<string, NettoDipendente>();
    for (const x of ms) {
      if (coperti.has(x.chiave)) continue;
      const d = nuove.get(x.chiave) ?? {
        nome: x.nome,
        stipendio: 0,
        anticipo: 0,
        addebito: 0,
        saldo: 0,
      };
      (d as unknown as Record<string, number>)[x.campo] = x.valore;
      nuove.set(x.chiave, d);
    }
    dipendenti.push(...nuove.values());
    nettiMap.set(mese, {
      ...base,
      dipendenti,
      totaleStipendio: r2(dipendenti.reduce((a, d) => a + d.stipendio, 0)),
      totaleAnticipi: r2(dipendenti.reduce((a, d) => a + d.anticipo, 0)),
      totaleSaldo: r2(dipendenti.reduce((a, d) => a + d.saldo, 0)),
    });
  }
  const netti = [...nettiMap.values()].sort((a, b) => (a.mese < b.mese ? -1 : 1));
  return { ...db, mesi, netti };
}

/** Valore del FILE (senza modifiche) per la traccia "prima: …". */
export function valoreDaFile(
  db: StipendiDb,
  mese: string,
  chiave: string,
  campo: CampoModificabile,
): number | null {
  if ((CAMPI_COSTI as readonly string[]).includes(campo)) {
    const d = db.mesi
      .find((m) => m.mese === mese)
      ?.dipendenti.find((x) => chiaveNome(`${x.cognome} ${x.nome}`) === chiave);
    return d ? ((d as unknown as Record<string, number>)[campo] ?? null) : null;
  }
  const n = (db.netti ?? [])
    .find((m) => m.mese === mese)
    ?.dipendenti.find((x) => chiaveNome(x.nome) === chiave);
  return n ? ((n as unknown as Record<string, number>)[campo] ?? null) : null;
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
  // Chiavi con almeno una riga di indirizzamento riconosciuta: una "persona"
  // senza nemmeno una voce è una nota scritta nella colonna cognome
  // ("COSTO EXTRA 1", "Nel costo sono inclusi…"), non un dipendente.
  const conVoci = new Set<string>();
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
    if (ind) conVoci.add(chiave);
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
  for (const chiave of perDip.keys()) {
    if (conVoci.has(chiave)) continue;
    perDip.delete(chiave);
    scartate++;
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

/** Interpreta "Stipendi Dr.xlsx": un foglio per mese ("Febbraio 2026"…) con
 *  APPALTO, NOME, STIPENDIO, ANTICIPO, ADDEBITO, SALDO (+ IBAN). Le colonne
 *  si trovano per intestazione (cambiano posizione tra i fogli); il mese
 *  viene dal NOME DEL FOGLIO. null se nessun foglio è riconosciuto. */
export function parseStipendiDr(
  fogli: { nome: string; matrix: unknown[][] }[],
  fonteFile: string,
): NettiMese[] | null {
  const out: NettiMese[] = [];
  for (const f of fogli) {
    // Mese dal nome foglio ("Maggio 2026").
    const mese = meseDaNomeFile(f.nome);
    if (!mese) continue;
    const headerIdx = f.matrix.findIndex((r) => {
      const cells = (r ?? []).map(norm);
      return cells.includes("nome") && cells.includes("stipendio");
    });
    if (headerIdx < 0) continue;
    const header = (f.matrix[headerIdx] ?? []).map(norm);
    const C = {
      appalto: header.indexOf("appalto"),
      nome: header.indexOf("nome"),
      stipendio: header.indexOf("stipendio"),
      anticipo: header.indexOf("anticipo"),
      addebito: header.indexOf("addebito"),
      saldo: header.indexOf("saldo"),
    };
    if (C.nome < 0 || C.stipendio < 0) continue;
    const dipendenti: NettoDipendente[] = [];
    for (const r of f.matrix.slice(headerIdx + 1)) {
      const nome = String(r?.[C.nome] ?? "").trim();
      if (!nome || /^totale/i.test(nome)) continue;
      const stipendio = numCell(r[C.stipendio]);
      const anticipo = C.anticipo >= 0 ? numCell(r[C.anticipo]) : 0;
      const addebito = C.addebito >= 0 ? numCell(r[C.addebito]) : 0;
      const saldoFile = C.saldo >= 0 ? numCell(r[C.saldo]) : 0;
      const saldo =
        Math.abs(saldoFile) > 0.004
          ? saldoFile
          : Math.round((stipendio - anticipo - addebito) * 100) / 100;
      if (Math.abs(stipendio) < 0.004 && Math.abs(saldo) < 0.004) continue;
      dipendenti.push({
        nome,
        appalto: C.appalto >= 0 ? String(r[C.appalto] ?? "").trim() || undefined : undefined,
        stipendio: Math.round(stipendio * 100) / 100,
        anticipo: Math.round(anticipo * 100) / 100,
        addebito: Math.round(addebito * 100) / 100,
        saldo,
      });
    }
    if (dipendenti.length === 0) continue;
    const r2 = (n: number) => Math.round(n * 100) / 100;
    out.push({
      mese,
      fonteFile,
      dipendenti,
      totaleStipendio: r2(dipendenti.reduce((a, d) => a + d.stipendio, 0)),
      totaleAnticipi: r2(dipendenti.reduce((a, d) => a + d.anticipo, 0)),
      totaleSaldo: r2(dipendenti.reduce((a, d) => a + d.saldo, 0)),
    });
  }
  if (out.length === 0) return null;
  out.sort((a, b) => (a.mese < b.mese ? -1 : 1));
  return out;
}

/** Estrae il livello contrattuale dalla mansione ("AUTISTA G1" → "G1",
 *  "IMPIEGATO 5" → "5", "MAGAZZINIERE 6 L" → "6 L"). */
export function livelloDaMansione(mansione: string): string {
  // "IMPIEGATO 1 LIVELLO" → "1"
  const liv = mansione.trim().match(/(?:^|\s)([A-Z]?\d[A-Z]?)\s+LIVELLO\s*$/i);
  if (liv) return liv[1].toUpperCase();
  // "AUTISTA G1" → "G1", "MAGAZZINIERE 6 L" → "6 L"
  const m = mansione.trim().match(/\s([A-Z]?\d[A-Z]?)(\s?[SL])?\s*$/i);
  return m ? `${m[1]}${m[2] ? ` ${m[2].trim()}` : ""}`.toUpperCase() : "";
}

/** Interpreta la "Mappatura Dipendenti Dr Logistica.xlsx" (foglio 1):
 *  colonne per intestazione NOME, MANSIONE, DATA FINE (data oppure
 *  "INDETERMINATO"), ORARIO, STATO, PROROGA 1..4. null se non riconosciuta. */
export function parseMappatura(
  fogli: { nome: string; matrix: unknown[][] }[],
): AnagraficaDipendente[] | null {
  for (const f of fogli) {
    const headerIdx = f.matrix.findIndex((r) => {
      const cells = (r ?? []).map(norm);
      return cells.includes("nome") && cells.includes("mansione");
    });
    if (headerIdx < 0) continue;
    const header = (f.matrix[headerIdx] ?? []).map(norm);
    const col = (n: string) => header.findIndex((c) => c === n || c.startsWith(n));
    const C = {
      nome: col("nome"),
      mansione: col("mansione"),
      fine: col("data fine"),
      orario: col("orario"),
      stato: col("stato"),
      proroghe: [1, 2, 3, 4].map((i) => col(`proroga ${i}`)),
    };
    if (C.nome < 0 || C.mansione < 0) continue;
    const toIso = (v: unknown): string => {
      // Con `raw:true` le date Excel arrivano come numero seriale.
      if (typeof v === "number" && Number.isFinite(v) && v > 25569 && v < 80000)
        return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000)
          .toISOString()
          .slice(0, 10);
      const s = String(v ?? "").trim();
      const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return `${m[1]}-${m[2]}-${m[3]}`;
      const it = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (it) return `${it[3]}-${it[2].padStart(2, "0")}-${it[1].padStart(2, "0")}`;
      return "";
    };
    const out: AnagraficaDipendente[] = [];
    for (const r of f.matrix.slice(headerIdx + 1)) {
      const nome = String(r?.[C.nome] ?? "").trim();
      if (!nome) continue;
      const mansione = String(r[C.mansione] ?? "").trim() || undefined;
      // La cella va passata GREZZA a toIso: stringificarla prima trasforma
      // il seriale Excel in "46112", che nessun ramo riconosce più.
      const fineCell = C.fine >= 0 ? r[C.fine] : null;
      const fineRaw = String(fineCell ?? "").trim();
      const indet = /^indet/i.test(fineRaw);
      // Fine effettiva: la più avanzata tra data fine e proroghe.
      const date = [toIso(fineCell), ...C.proroghe.map((i) => (i >= 0 ? toIso(r[i]) : ""))]
        .filter(Boolean)
        .sort();
      out.push({
        nome,
        mansione,
        livello: mansione ? livelloDaMansione(mansione) : undefined,
        contratto: indet ? "Indeterminato" : fineRaw || date.length ? "Determinato" : undefined,
        fineContratto: indet ? undefined : date[date.length - 1] || undefined,
        orario: C.orario >= 0 ? String(r[C.orario] ?? "").trim() || undefined : undefined,
        stato: C.stato >= 0 ? String(r[C.stato] ?? "").trim() || undefined : undefined,
      });
    }
    if (out.length) return out;
  }
  return null;
}

/** Mensilità stimate (12/13/14) dai ratei di "Mens.Agg." dei mesi caricati:
 *  mediana di (rateo mensilità aggiuntive / costo ordinario) → 0≈12,
 *  1/12≈13, 2/12≈14. I mesi con ordinario basso (assunzioni/cessazioni) si
 *  scartano; null se non c'è abbastanza storia. */
export function mensilitaStimate(ratios: number[]): number | null {
  const validi = ratios.filter((r) => Number.isFinite(r));
  if (validi.length === 0) return null;
  const s = [...validi].sort((a, b) => a - b);
  const med =
    s.length % 2 ? s[Math.floor(s.length / 2)] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
  const extra = Math.round(med * 12);
  return 12 + Math.max(0, Math.min(2, extra));
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
    // "Dettagli1", "Dettagli2"…: fogli che Excel genera da solo col doppio
    // clic su una cella della pivot (caso COSTI GIUGNO 2026, 21/09): sono
    // un ESTRATTO del foglio vero e sommarli raddoppia le persone.
    if (/^dettagli\s*\d*$/i.test(f.nome.trim())) continue;
    const res = parseCostiMese(f.matrix, f.nome);
    if (res) parziali.push({ nome: f.nome, res });
  }
  if (parziali.length === 0) return null;
  // Tracciato COMPLETO (righe col codice dipendente): la stessa persona può
  // comparire in più fogli solo per copie/estratti → vince il foglio più
  // ricco e ogni codice conta una volta. Nel tracciato per-appalto (senza
  // codici, un foglio per appalto) la stessa persona in due fogli sono due
  // costi veri e si sommano come prima.
  parziali.sort((a, b) => b.res.dipendenti.length - a.res.dipendenti.length);
  const visti = new Set<string>();
  const dipendenti = parziali.flatMap((p) =>
    p.res.dipendenti.filter((d) => {
      if (!d.codice) return true;
      if (visti.has(d.codice)) return false;
      visti.add(d.codice);
      return true;
    }),
  );
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
