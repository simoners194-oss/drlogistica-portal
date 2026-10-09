// DR Portal — ESTRATTO CONTO DI GRUPPO: logica pura (testabile senza UI).
//
// Nato dal caso Univex (ottobre 2026): il cliente è un GRUPPO di società
// (Healthcare e Freight ci pagano, Nolvex e Univex Srl ci fatturano) che
// regola per COMPENSAZIONE — nessun bonifico nostro verso di loro, e i loro
// bonifici coprono le nostre fatture al netto di quello che dobbiamo. Lo
// "scaduto per società" dipende da come ciascuno attribuisce le
// compensazioni: il cliente le scarica tutte su una società (Freight), su
// Aruba erano finite altrove, e per settimane i numeri non tornavano.
//
// Qui si ricostruisce l'estratto conto COME LO VEDE IL CLIENTE:
//   - per società e mese di competenza: fatturato (al netto delle note di
//     credito), bonifici ricevuti, compensazioni, aperto;
//   - regola fissa per gruppo: le fatture passive di chi non ci fattura
//     (Nolvex…) si compensano sulla società "che assorbe" (Freight); quelle
//     di una società che ci fattura si compensano sulle sue stesse fatture;
//   - accanto, la vista "Aruba" (quello che il portale registra oggi) e la
//     differenza per società: zero quando le attribuzioni coincidono.
// Più il SEMAFORO mensile: incassi registrati su Aruba − bonifici in banca
// deve essere uguale alle fatture passive pagate per compensazione; se non
// torna, qualcuno ha registrato un incasso senza soldi e senza
// compensazione (il 17/09 avrebbe acceso il rosso lo stesso giorno). E la
// PREVISIONE: cosa scade nei prossimi mesi, al netto di quanto verrà
// trattenuto.
import { groupKeyAlgoritmica } from "./finanza-logic";
import {
  collegaNoteCredito,
  computeStatoFattura,
  fattureEscluse,
  isNotaCredito,
  parseIncassoAruba,
  type FatturaRaw,
  type TerminePagamento,
} from "./fatture-logic";

export interface GruppoDef {
  nome: string;
  /** Membri come sulla lista GruppiControparti: nomi (o pezzi) separati da
   *  virgola, punto e virgola o a capo; match per parole contenute. */
  membri: string;
}

export interface MovimentoEC {
  dataContabile: string; // YYYY-MM-DD
  importo: number; // >0 entrata, <0 uscita
  cliente: string;
}

export interface SocietaEC {
  chiave: string;
  nome: string;
  /** true = ci fattura (ha fatture attive nel gruppo). */
  ciFattura: boolean;
  /** true = ci manda fatture (passive). */
  ciFatturaPassive: boolean;
}

export interface FatturaEC {
  nomeFile: string;
  numero: string;
  data: string;
  societa: string; // chiave società
  competenza: string; // YYYY-MM
  scadenza: string; // YYYY-MM-DD
  /** Totale al netto delle note di credito collegate. */
  netto: number;
  /** Incassato registrato su Aruba (importo delle rate). */
  incassatoAruba: number;
  residuoAruba: number;
  /** Residuo nella vista cliente (dopo bonifici e compensazioni). */
  residuoCliente: number;
  bonifici: number;
  compensazioni: number;
  oggetto?: string;
}

export interface PassivaEC {
  nomeFile: string;
  numero: string;
  data: string;
  societa: string;
  competenza: string;
  netto: number;
  /** Dove è finita nella vista cliente: chiave della società su cui si
   *  compensa, oppure "bonifico" se l'abbiamo pagata noi. */
  compensataSu?: string;
  /** Parte non allocata (nessuna fattura attiva aperta su cui scaricarla). */
  nonAllocata: number;
}

export interface RigaMeseEC {
  societa: string;
  competenza: string;
  fatturato: number;
  bonifici: number;
  compensazioni: number;
  apertoCliente: number;
  apertoAruba: number;
  /** Scaduto (vista cliente) alla data di riferimento. */
  scadutoCliente: number;
  fatture: FatturaEC[];
}

export interface TotaleSocietaEC {
  societa: string;
  fatturato: number;
  bonifici: number;
  compensazioni: number;
  apertoCliente: number;
  scadutoCliente: number;
  apertoAruba: number;
  /** apertoAruba − apertoCliente: attribuzioni diverse delle compensazioni. */
  delta: number;
}

export interface SemaforoMeseEC {
  mese: string; // YYYY-MM
  arubaIncassato: number;
  banca: number;
  compensazioni: number;
  /** arubaIncassato − banca − compensazioni: ≈ 0 quando tutto torna. */
  delta: number;
  stato: "ok" | "attenzione" | "rosso";
}

export interface PrevisioneMeseEC {
  mese: string; // YYYY-MM, oppure "scaduto"
  daIncassare: number;
  daTrattenere: number;
  netto: number;
}

export interface EstrattoConto {
  gruppo: string;
  assorbe: string | null;
  societa: SocietaEC[];
  fatture: FatturaEC[];
  passive: PassivaEC[];
  righe: RigaMeseEC[];
  totali: TotaleSocietaEC[];
  totaleGruppo: TotaleSocietaEC;
  /** Fatture passive (vista cliente) ancora da compensare: quanto dobbiamo. */
  passivoAperto: number;
  bonificiNonAllocati: number;
  semaforo: SemaforoMeseEC[];
  previsione: PrevisioneMeseEC[];
}

/** Soglie del semaforo: sotto la prima è arrotondamento o timing, sopra la
 *  seconda è una registrazione senza soldi e senza compensazione. */
export const SOGLIA_SEMAFORO_ATTENZIONE = 1_000;
export const SOGLIA_SEMAFORO_ROSSO = 10_000;

const MESI = [
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

const r2 = (n: number) => Math.round(n * 100) / 100;

/** Parole "nude" di un nome (senza forma societaria, ordinate). */
export function paroleNome(nome: string): string[] {
  return groupKeyAlgoritmica(nome).split(" ").filter(Boolean);
}

/** I membri del gruppo come liste di parole. */
export function tokensMembri(membri: string): string[][] {
  return membri
    .split(/[,;\n\r]+/)
    .map((m) => paroleNome(m))
    .filter((w) => w.length > 0);
}

/** Una controparte appartiene al gruppo se TUTTE le parole di almeno un
 *  membro stanno nel suo nome ("univex" prende Univex Healthcare, Univex
 *  Freight e Univex Srl; "nolvex" prende Nolvex). */
export function appartieneAlGruppo(nome: string, tokens: readonly string[][]): boolean {
  const parole = paroleNome(nome);
  if (!parole.length) return false;
  return tokens.some((tk) => tk.every((w) => parole.includes(w)));
}

/** Chiave della SOCIETÀ: parole nude, senza alias di gruppo (le società del
 *  gruppo devono restare distinte). */
export function chiaveSocieta(nome: string): string {
  return paroleNome(nome).join(" ");
}

/** Mese di competenza di una fattura: il mese scritto nell'oggetto
 *  ("…NEL MESE DI GIUGNO 2026…"); altrimenti il mese prima della data
 *  documento se il termine del cliente dice "competenza precedente",
 *  altrimenti il mese della data documento. */
export function competenzaDi(
  f: Pick<FatturaRaw, "oggetto" | "causaleDoc" | "dataDocumento" | "meseCompetenza">,
  competenzaPrecedente = false,
): string {
  const testo = `${f.oggetto ?? ""} ${f.causaleDoc ?? ""} ${f.meseCompetenza ?? ""}`.toLowerCase();
  const m = testo.match(
    /\b(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\b(?:\s+(?:di\s+)?(\d{4}))?/,
  );
  const docAnno = Number(f.dataDocumento.slice(0, 4));
  const docMese = Number(f.dataDocumento.slice(5, 7));
  if (m) {
    const mese = MESI.indexOf(m[1]) + 1;
    let anno = m[2] ? Number(m[2]) : docAnno;
    // Senza anno nell'oggetto: "dicembre" su una fattura di gennaio è l'anno prima.
    if (!m[2] && mese > docMese) anno = docAnno - 1;
    return `${anno}-${String(mese).padStart(2, "0")}`;
  }
  if (competenzaPrecedente) {
    return new Date(Date.UTC(docAnno, docMese - 2, 1)).toISOString().slice(0, 7);
  }
  return f.dataDocumento.slice(0, 7);
}

function termineDi(
  termini: readonly TerminePagamento[],
  nome: string,
): TerminePagamento | undefined {
  const k = chiaveSocieta(nome);
  return termini.find(
    (t) =>
      (!t.direzione || t.direzione === "Emessa") && !t.oggetto && chiaveSocieta(t.cliente) === k,
  );
}

function meseRelativo(yyyymm: string, delta: number): string {
  const [y, m] = yyyymm.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1 + delta, 1)).toISOString().slice(0, 7);
}

export interface InputEstratto {
  gruppo: GruppoDef;
  attive: readonly FatturaRaw[];
  passive: readonly FatturaRaw[];
  movimenti: readonly MovimentoEC[];
  termini: readonly TerminePagamento[];
  /** Data di riferimento (YYYY-MM-DD): scaduto, previsione, semaforo. */
  oggi: string;
  /** Chiave della società che assorbe le compensazioni di chi non ci
   *  fattura; vuoto = quella col fatturato maggiore. */
  assorbe?: string;
  /** Mesi di semaforo da mostrare (a ritroso da oggi). */
  mesiSemaforo?: number;
  /** Mesi di previsione in avanti. */
  mesiPrevisione?: number;
}

export function calcolaEstratto(input: InputEstratto): EstrattoConto {
  const tokens = tokensMembri(input.gruppo.membri);
  const oggi = input.oggi;
  const inGruppo = (nome: string) => appartieneAlGruppo(nome, tokens);

  // --- Società del gruppo (nome più frequente per chiave) ----------------
  const nomi = new Map<string, Map<string, number>>();
  const conta = (nome: string) => {
    const k = chiaveSocieta(nome);
    if (!k) return;
    const m = nomi.get(k) ?? new Map<string, number>();
    m.set(nome.trim(), (m.get(nome.trim()) ?? 0) + 1);
    nomi.set(k, m);
  };
  const attiveG = input.attive.filter((f) => inGruppo(f.cliente));
  const passiveG = input.passive.filter((f) => inGruppo(f.cliente));
  const moviG = input.movimenti.filter((m) => inGruppo(m.cliente));
  for (const f of attiveG) conta(f.cliente);
  for (const f of passiveG) conta(f.cliente);
  for (const m of moviG) conta(m.cliente);
  const societa: SocietaEC[] = [...nomi.entries()].map(([chiave, m]) => ({
    chiave,
    nome: [...m.entries()].sort((a, b) => b[1] - a[1])[0][0],
    ciFattura: attiveG.some(
      (f) => chiaveSocieta(f.cliente) === chiave && !isNotaCredito(f.tipoDocumento),
    ),
    ciFatturaPassive: passiveG.some(
      (f) => chiaveSocieta(f.cliente) === chiave && !isNotaCredito(f.tipoDocumento),
    ),
  }));
  societa.sort((a, b) => a.nome.localeCompare(b.nome, "it"));

  // --- Fatture attive: netto NC, competenza, scadenza, Aruba ---------------
  const escluseA = fattureEscluse(input.attive);
  const ncA = collegaNoteCredito(input.attive, escluseA);
  const fatture: FatturaEC[] = [];
  for (const f of attiveG) {
    if (isNotaCredito(f.tipoDocumento) || escluseA.has(f.nomeFile)) continue;
    if (parseIncassoAruba(f.incassoAruba) === "Stornata") continue;
    const nc = ncA.get(f.nomeFile)?.importo ?? 0;
    const netto = r2(f.totale - nc);
    if (netto <= 0.005) continue;
    const termine = termineDi(input.termini, f.cliente);
    const stato = computeStatoFattura(f, 0, input.termini, oggi, nc);
    const incAruba =
      f.incassatoAruba != null
        ? Math.min(f.incassatoAruba, netto)
        : parseIncassoAruba(f.incassoAruba) === "Incassata"
          ? netto
          : 0;
    fatture.push({
      nomeFile: f.nomeFile,
      numero: f.numero,
      data: f.dataDocumento,
      societa: chiaveSocieta(f.cliente),
      competenza: competenzaDi(f, termine?.competenzaPrecedente ?? false),
      scadenza: stato.scadenza || f.dataDocumento,
      netto,
      incassatoAruba: r2(incAruba),
      residuoAruba: r2(netto - incAruba),
      residuoCliente: netto,
      bonifici: 0,
      compensazioni: 0,
      oggetto: f.oggetto,
    });
  }
  fatture.sort((a, b) => a.competenza.localeCompare(b.competenza) || a.data.localeCompare(b.data));

  // --- Passive: netto NC, competenza (mese prima della data: ci fatturano
  // il mese dopo il servizio) ----------------------------------------------
  const escluseP = fattureEscluse(input.passive);
  const ncP = collegaNoteCredito(input.passive, escluseP);
  const passive: PassivaEC[] = [];
  for (const f of passiveG) {
    if (isNotaCredito(f.tipoDocumento) || escluseP.has(f.nomeFile)) continue;
    if (parseIncassoAruba(f.incassoAruba) === "Stornata") continue;
    const netto = r2(Math.abs(f.netto || f.totale) - (ncP.get(f.nomeFile)?.importo ?? 0));
    if (netto <= 0.005) continue;
    passive.push({
      nomeFile: f.nomeFile,
      numero: f.numero,
      data: f.dataDocumento,
      societa: chiaveSocieta(f.cliente),
      competenza: competenzaDi(f, true),
      netto,
      nonAllocata: netto,
    });
  }
  passive.sort((a, b) => a.data.localeCompare(b.data));

  // --- Bonifici: entrate per società, FIFO con abbinamento esatto prima ---
  const entrate = moviG
    .filter((m) => m.importo > 0.005)
    .map((m) => ({ ...m, societa: chiaveSocieta(m.cliente), resto: m.importo }))
    .sort((a, b) => a.dataContabile.localeCompare(b.dataContabile));
  const perSocieta = (chiave: string) => fatture.filter((f) => f.societa === chiave);
  for (const p of entrate) {
    const mie = perSocieta(p.societa);
    // Abbinamento esatto solo per importi "non tondi": 78.080,00 è la fattura
    // 156/26 pagata a sé, 50.000,00 è un acconto e va sul più vecchio (FIFO)
    // anche se per caso una fattura vale proprio 50.000.
    const tondo = Math.abs(p.resto % 500) < 0.005;
    const esatta = tondo
      ? undefined
      : mie.find((f) => f.data <= p.dataContabile && Math.abs(f.residuoCliente - p.resto) < 0.01);
    if (esatta) {
      esatta.bonifici = r2(esatta.bonifici + p.resto);
      esatta.residuoCliente = 0;
      p.resto = 0;
      continue;
    }
    for (const f of mie) {
      if (p.resto <= 0.005) break;
      if (f.residuoCliente <= 0.005) continue;
      const u = Math.min(p.resto, f.residuoCliente);
      f.bonifici = r2(f.bonifici + u);
      f.residuoCliente = r2(f.residuoCliente - u);
      p.resto = r2(p.resto - u);
    }
  }
  const bonificiNonAllocati = r2(entrate.reduce((s, p) => s + p.resto, 0));

  // --- Uscite nostre verso il gruppo (rare): pagano le passive più vecchie,
  // il resto delle passive è compensazione --------------------------------
  let uscite = r2(
    moviG.filter((m) => m.importo < -0.005).reduce((s, m) => s + Math.abs(m.importo), 0),
  );
  for (const p of passive) {
    if (uscite <= 0.005) break;
    const u = Math.min(uscite, p.nonAllocata);
    p.nonAllocata = r2(p.nonAllocata - u);
    uscite = r2(uscite - u);
    if (p.nonAllocata <= 0.005) p.compensataSu = "bonifico";
  }

  // --- Compensazioni: sulla stessa società se ci fattura, altrimenti su chi
  // assorbe; prima la fattura della STESSA competenza, poi la più vecchia --
  const fatturatoPer = new Map<string, number>();
  for (const f of fatture)
    fatturatoPer.set(f.societa, (fatturatoPer.get(f.societa) ?? 0) + f.netto);
  const assorbe =
    input.assorbe && fatturatoPer.has(input.assorbe)
      ? input.assorbe
      : ([...fatturatoPer.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null);
  const scarica = (p: PassivaEC, chiave: string) => {
    const mie = perSocieta(chiave);
    const candidate = [
      ...mie.filter((f) => f.competenza === p.competenza && f.residuoCliente > 0.005),
      ...mie.filter((f) => f.competenza !== p.competenza && f.residuoCliente > 0.005),
    ];
    for (const f of candidate) {
      if (p.nonAllocata <= 0.005) break;
      const u = Math.min(p.nonAllocata, f.residuoCliente);
      f.compensazioni = r2(f.compensazioni + u);
      f.residuoCliente = r2(f.residuoCliente - u);
      p.nonAllocata = r2(p.nonAllocata - u);
      p.compensataSu = chiave;
    }
  };
  for (const p of passive) {
    if (p.nonAllocata <= 0.005) continue;
    const propria = societa.find((s) => s.chiave === p.societa)?.ciFattura;
    if (propria) scarica(p, p.societa);
    if (p.nonAllocata > 0.005 && assorbe) scarica(p, assorbe);
  }
  const passivoAperto = r2(passive.reduce((s, p) => s + p.nonAllocata, 0));

  // --- Righe per società × mese e totali -----------------------------------
  const righeMap = new Map<string, RigaMeseEC>();
  for (const f of fatture) {
    const k = `${f.societa}|${f.competenza}`;
    const r =
      righeMap.get(k) ??
      ({
        societa: f.societa,
        competenza: f.competenza,
        fatturato: 0,
        bonifici: 0,
        compensazioni: 0,
        apertoCliente: 0,
        apertoAruba: 0,
        scadutoCliente: 0,
        fatture: [],
      } as RigaMeseEC);
    r.fatturato = r2(r.fatturato + f.netto);
    r.bonifici = r2(r.bonifici + f.bonifici);
    r.compensazioni = r2(r.compensazioni + f.compensazioni);
    r.apertoCliente = r2(r.apertoCliente + f.residuoCliente);
    r.apertoAruba = r2(r.apertoAruba + f.residuoAruba);
    if (f.scadenza <= oggi) r.scadutoCliente = r2(r.scadutoCliente + f.residuoCliente);
    r.fatture.push(f);
    righeMap.set(k, r);
  }
  const righe = [...righeMap.values()].sort(
    (a, b) => a.societa.localeCompare(b.societa) || a.competenza.localeCompare(b.competenza),
  );
  const somma = (xs: readonly RigaMeseEC[], k: keyof RigaMeseEC) =>
    r2(xs.reduce((acc, r) => acc + (r[k] as number), 0));
  const totali: TotaleSocietaEC[] = societa
    .filter((s) => s.ciFattura)
    .map((s) => {
      const mie = righe.filter((r) => r.societa === s.chiave);
      const t: TotaleSocietaEC = {
        societa: s.chiave,
        fatturato: somma(mie, "fatturato"),
        bonifici: somma(mie, "bonifici"),
        compensazioni: somma(mie, "compensazioni"),
        apertoCliente: somma(mie, "apertoCliente"),
        scadutoCliente: somma(mie, "scadutoCliente"),
        apertoAruba: somma(mie, "apertoAruba"),
        delta: 0,
      };
      t.delta = r2(t.apertoAruba - t.apertoCliente);
      return t;
    });
  const sommaT = (k: keyof TotaleSocietaEC) => r2(totali.reduce((s, t) => s + (t[k] as number), 0));
  const totaleGruppo: TotaleSocietaEC = {
    societa: "*",
    fatturato: sommaT("fatturato"),
    bonifici: sommaT("bonifici"),
    compensazioni: sommaT("compensazioni"),
    apertoCliente: sommaT("apertoCliente"),
    scadutoCliente: sommaT("scadutoCliente"),
    apertoAruba: sommaT("apertoAruba"),
    delta: 0,
  };
  totaleGruppo.delta = r2(totaleGruppo.apertoAruba - totaleGruppo.apertoCliente);

  // --- Semaforo mensile: Aruba − banca − compensazioni registrate ---------
  const nMesi = input.mesiSemaforo ?? 12;
  const meseOggi = oggi.slice(0, 7);
  const mesi: string[] = [];
  for (let i = nMesi - 1; i >= 0; i--) mesi.push(meseRelativo(meseOggi, -i));
  const arubaPerMese = new Map<string, number>();
  for (const f of attiveG) {
    if (isNotaCredito(f.tipoDocumento) || !f.dataIncasso) continue;
    const inc =
      f.incassatoAruba != null
        ? f.incassatoAruba
        : parseIncassoAruba(f.incassoAruba) === "Incassata"
          ? Math.abs(f.totale) - (ncA.get(f.nomeFile)?.importo ?? 0)
          : 0;
    if (inc > 0) {
      const k = f.dataIncasso.slice(0, 7);
      arubaPerMese.set(k, r2((arubaPerMese.get(k) ?? 0) + inc));
    }
  }
  const compPerMese = new Map<string, number>();
  for (const f of passiveG) {
    if (isNotaCredito(f.tipoDocumento) || !f.dataIncasso) continue;
    const pag =
      f.incassatoAruba != null
        ? f.incassatoAruba
        : parseIncassoAruba(f.incassoAruba) === "Incassata"
          ? Math.abs(f.netto || f.totale) - (ncP.get(f.nomeFile)?.importo ?? 0)
          : 0;
    if (pag > 0) {
      const k = f.dataIncasso.slice(0, 7);
      compPerMese.set(k, r2((compPerMese.get(k) ?? 0) + pag));
    }
  }
  const bancaPerMese = new Map<string, number>();
  const uscitePerMese = new Map<string, number>();
  for (const m of moviG) {
    const k = m.dataContabile.slice(0, 7);
    if (m.importo > 0) bancaPerMese.set(k, r2((bancaPerMese.get(k) ?? 0) + m.importo));
    else uscitePerMese.set(k, r2((uscitePerMese.get(k) ?? 0) + Math.abs(m.importo)));
  }
  const semaforo: SemaforoMeseEC[] = mesi.map((mese) => {
    const aruba = arubaPerMese.get(mese) ?? 0;
    const banca = bancaPerMese.get(mese) ?? 0;
    const comp = r2((compPerMese.get(mese) ?? 0) - (uscitePerMese.get(mese) ?? 0));
    const delta = r2(aruba - banca - comp);
    const abs = Math.abs(delta);
    return {
      mese,
      arubaIncassato: aruba,
      banca,
      compensazioni: comp,
      delta,
      stato:
        abs < SOGLIA_SEMAFORO_ATTENZIONE
          ? "ok"
          : abs < SOGLIA_SEMAFORO_ROSSO
            ? "attenzione"
            : "rosso",
    };
  });

  // --- Previsione: aperto per mese di scadenza, meno le passive da trattenere
  const nPrev = input.mesiPrevisione ?? 6;
  const previsione: PrevisioneMeseEC[] = [];
  const scaduto = r2(
    fatture.filter((f) => f.scadenza < oggi).reduce((s, f) => s + f.residuoCliente, 0),
  );
  previsione.push({ mese: "scaduto", daIncassare: scaduto, daTrattenere: 0, netto: scaduto });
  for (let i = 0; i < nPrev; i++) {
    const mese = meseRelativo(meseOggi, i);
    const daInc = r2(
      fatture
        .filter((f) => f.scadenza >= oggi && f.scadenza.slice(0, 7) === mese)
        .reduce((s, f) => s + f.residuoCliente, 0),
    );
    previsione.push({ mese, daIncassare: daInc, daTrattenere: 0, netto: daInc });
  }
  // Le passive ancora da compensare si trattengono sul primo incasso utile.
  let daTrattenere = passivoAperto;
  for (const p of previsione) {
    if (daTrattenere <= 0.005) break;
    const u = Math.min(daTrattenere, p.daIncassare);
    p.daTrattenere = r2(u);
    p.netto = r2(p.daIncassare - u);
    daTrattenere = r2(daTrattenere - u);
  }

  return {
    gruppo: input.gruppo.nome,
    assorbe,
    societa,
    fatture,
    passive,
    righe,
    totali,
    totaleGruppo,
    passivoAperto,
    bonificiNonAllocati,
    semaforo,
    previsione,
  };
}

/** Etichetta mese "giu 2026". */
export function etichettaMese(yyyymm: string): string {
  const [y, m] = yyyymm.split("-");
  const nomi = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"];
  return `${nomi[Number(m) - 1] ?? m} ${y}`;
}

/** Configurazione del gruppo salvata sui Flussi (genere "gruppo", Note JSON
 *  compatto {"a": chiave della società che assorbe}). */
export interface ConfigGruppoEC {
  assorbe?: string;
}

export function parseConfigGruppo(note?: string | null): ConfigGruppoEC {
  if (!note) return {};
  try {
    const j = JSON.parse(note) as Record<string, unknown>;
    return { assorbe: typeof j.a === "string" && j.a ? j.a : undefined };
  } catch {
    return {};
  }
}

export function serializeConfigGruppo(cfg: ConfigGruppoEC): string {
  return JSON.stringify({ a: cfg.assorbe ?? "" });
}
