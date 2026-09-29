// DR Portal — Flussi di cassa: logica pura su MODALITÀ DI PAGAMENTO delle
// fatture passive (RID / RiBa / nessuna) e sulla scheda Fornitori.
// Richiesta Simone 29/09/2026: "il flusso di cassa deve essere diviso tra
// fatture in entrata con RIBA, con RID e senza una di queste", impostate
// dai file di Sabrina e modificabili a mano; più una scheda con una riga
// per fornitore presente nel flusso (chi se ne occupa, macrovoce, appalto,
// scaduto, aperto, conteggi, date, pagato).
//
// FONTE AUTOMATICA: il campo ModalitaPagamento dell'XML FatturaPA di OGNI
// fattura (MP12 = RiBa; MP09/10/11 RID, MP16/17 domiciliazioni, MP19/20/21
// SEPA Direct Debit = addebiti automatici, per il direttore "RID"). Il foglio
// "RID" di Sabrina (DR SITUAZIONE 2026.xlsx) combacia: Califano Carrelli e
// CB Pneumatici "RI.BA OK", Kuwait/Consat/Arval/Kion/BNP in SDD. L'unica
// differenza (DKV: bonifico in XML, RID per Sabrina) si sana con l'override
// manuale per fornitore, salvato sulla lista FlussiCassa (genere "fornitore",
// Note = JSON compatto) insieme a referente, macrovoce e appalto.
import { clienteGroupKey } from "./finanza-logic";
import {
  incassatoRegistrato,
  isNotaCredito,
  residuoAperto,
  type FatturaConStato,
} from "./fatture-logic";

export type ModalitaPagamento = "rid" | "riba" | "altro";

/** Ordine di presentazione nei Flussi: prima gli addebiti automatici. */
export const MODALITA_ORDINE: readonly ModalitaPagamento[] = ["riba", "rid", "altro"];

/** Modalità dal campo MetodoPagamento della fattura ("MP19 - SDD", oppure
 *  il testo libero dell'export Aruba). Tutto ciò che non è un addebito
 *  automatico né una RiBa è "altro" (bonifico, carta, contanti…). */
export function modalitaDaMetodo(metodo: string | undefined | null): ModalitaPagamento {
  const s = String(metodo ?? "").trim();
  const code = s.match(/\bMP(\d{2})\b/i)?.[1];
  if (code) {
    if (code === "12") return "riba";
    if (["09", "10", "11", "16", "17", "19", "20", "21"].includes(code)) return "rid";
    return "altro";
  }
  const low = s.toLowerCase();
  if (/ri\.?\s?ba/.test(low)) return "riba";
  if (/\brid\b|sdd|sepa|domicil|addebito\s+dirett/.test(low)) return "rid";
  return "altro";
}

export interface FornitoreInfo {
  /** Override manuale: assente = automatico dalle fatture. */
  modalita?: ModalitaPagamento;
  referente?: string;
  macrovoce?: string;
  appalto?: string;
}

/** Chiave di raggruppamento del fornitore (stessa dei Flussi/Resoconto). */
export function chiaveFornitore(nome: string): string {
  return clienteGroupKey(nome) || nome.trim().toLowerCase();
}

const MAX_CAMPO = 60;
const pulisci = (v: unknown) =>
  String(v ?? "")
    .trim()
    .slice(0, MAX_CAMPO);

/** Note della riga FlussiCassa genere "fornitore": JSON compatto
 *  {"m":"rid","r":"Sabrina","c":"Carburante","a":"iMile Cerro"} — la colonna
 *  Note è testo a riga singola (255 caratteri), quindi chiavi corte e campi
 *  tagliati a 60. Tollerante: note vuote o non JSON = nessuna informazione. */
export function parseFornitoreInfo(note?: string | null): FornitoreInfo {
  const s = String(note ?? "").trim();
  if (!s.startsWith("{")) return {};
  try {
    const j = JSON.parse(s) as Record<string, unknown>;
    const m = String(j.m ?? "").toLowerCase();
    const out: FornitoreInfo = {};
    if (m === "rid" || m === "riba" || m === "altro") out.modalita = m;
    if (pulisci(j.r)) out.referente = pulisci(j.r);
    if (pulisci(j.c)) out.macrovoce = pulisci(j.c);
    if (pulisci(j.a)) out.appalto = pulisci(j.a);
    return out;
  } catch {
    return {};
  }
}

export function serializeFornitoreInfo(info: FornitoreInfo): string {
  const j: Record<string, string> = {};
  if (info.modalita) j.m = info.modalita;
  if (pulisci(info.referente)) j.r = pulisci(info.referente);
  if (pulisci(info.macrovoce)) j.c = pulisci(info.macrovoce);
  if (pulisci(info.appalto)) j.a = pulisci(info.appalto);
  return JSON.stringify(j);
}

export interface FornitoreSalvato extends FornitoreInfo {
  /** Title della riga su SharePoint: il nome con cui il fornitore è stato
   *  salvato (si riusa per aggiornare la stessa riga). */
  nomeRiga: string;
}

/** Mappa chiave fornitore → informazioni salvate (righe genere "fornitore"). */
export function mappaFornitori(
  righe: readonly { genere: string; nome: string; note?: string }[],
): Map<string, FornitoreSalvato> {
  const out = new Map<string, FornitoreSalvato>();
  for (const r of righe) {
    if (r.genere !== "fornitore" || !r.nome.trim()) continue;
    out.set(chiaveFornitore(r.nome), { nomeRiga: r.nome, ...parseFornitoreInfo(r.note) });
  }
  return out;
}

/** Modalità EFFETTIVA di una fattura: override del fornitore se c'è,
 *  altrimenti quella dichiarata nell'XML della fattura stessa. */
export function modalitaFattura(
  f: { cliente: string; metodoPagamento?: string },
  info: ReadonlyMap<string, FornitoreInfo>,
): ModalitaPagamento {
  return info.get(chiaveFornitore(f.cliente))?.modalita ?? modalitaDaMetodo(f.metodoPagamento);
}

/** Esclusioni dei Flussi (righe genere "esclusione", con finestra di mesi
 *  facoltativa): vero se la controparte è fuori dalla vista in quel mese. */
export function esclusaDaFlussi(
  esclusioni: readonly { nome: string; mese?: string; meseFine?: string }[],
  nomeControparte: string,
  meseScadenza: string,
): boolean {
  const chiave = clienteGroupKey(nomeControparte) || nomeControparte.toLowerCase();
  return esclusioni.some((e) => {
    const token = clienteGroupKey(e.nome) || e.nome.trim().toLowerCase();
    if (!token || !chiave.includes(token)) return false;
    if (e.mese && meseScadenza < e.mese) return false;
    if (e.meseFine && meseScadenza > e.meseFine) return false;
    return true;
  });
}

export interface RigaFornitore {
  chiave: string;
  /** Nome da mostrare e da usare per salvare (quello della riga salvata,
   *  altrimenti la dicitura più frequente sulle fatture). */
  nome: string;
  /** Modalità effettiva prevalente sulle fatture aperte (override compreso). */
  modalita: ModalitaPagamento;
  /** Modalità prevalente dichiarata nelle fatture (senza override). */
  modalitaAuto: ModalitaPagamento;
  override?: ModalitaPagamento;
  /** Fatture aperte con modalità dichiarate diverse (es. TIM: bollettino e SDD). */
  mista: boolean;
  referente: string;
  macrovoce: string;
  appalto: string;
  /** Proposte dalle fatture (tipologia di costo / cliente di riferimento). */
  macrovoceAuto: string;
  appaltoAuto: string;
  scaduto: number;
  /** Residuo aperto totale nel flusso: scaduto + a scadere. */
  aperto: number;
  nScadute: number;
  nAperte: number;
  /** Documenti totali del fornitore in archivio (note di credito escluse). */
  nTotale: number;
  nPagate: number;
  pagato: number;
  primaFattura: string;
  ultimaFattura: string;
  /** Prima scadenza tra le fatture aperte nel flusso ("" se nessuna). */
  prossimaScadenza: string;
}

function piuFrequente(valori: readonly string[]): string {
  const cnt = new Map<string, number>();
  for (const v of valori) {
    const s = v.trim();
    if (!s) continue;
    cnt.set(s, (cnt.get(s) ?? 0) + 1);
  }
  let best = "";
  let n = 0;
  for (const [v, c] of cnt) if (c > n) [best, n] = [v, c];
  return best;
}

/** Una riga per fornitore PRESENTE nel flusso (almeno una fattura aperta
 *  che passa `inFlusso`): conteggi e importi sulle aperte, storia (prima/
 *  ultima fattura, pagate, pagato) su tutto l'archivio del fornitore. */
export function riepilogoFornitori(
  passive: readonly FatturaConStato[],
  info: ReadonlyMap<string, FornitoreSalvato>,
  inFlusso: (x: FatturaConStato) => boolean,
): RigaFornitore[] {
  type Acc = {
    nomi: string[];
    modAperte: ModalitaPagamento[];
    modTutte: ModalitaPagamento[];
    macro: string[];
    appalti: string[];
    scaduto: number;
    aperto: number;
    nScadute: number;
    nAperte: number;
    nTotale: number;
    nPagate: number;
    pagato: number;
    prima: string;
    ultima: string;
    prossima: string;
  };
  const per = new Map<string, Acc>();
  for (const x of passive) {
    const k = chiaveFornitore(x.f.cliente);
    let a = per.get(k);
    if (!a) {
      a = {
        nomi: [],
        modAperte: [],
        modTutte: [],
        macro: [],
        appalti: [],
        scaduto: 0,
        aperto: 0,
        nScadute: 0,
        nAperte: 0,
        nTotale: 0,
        nPagate: 0,
        pagato: 0,
        prima: "",
        ultima: "",
        prossima: "",
      };
      per.set(k, a);
    }
    a.nomi.push(x.f.cliente);
    const nc = isNotaCredito(x.f.tipoDocumento);
    const modDich = modalitaDaMetodo(x.f.metodoPagamento);
    if (!nc) a.modTutte.push(modDich);
    a.macro.push(x.f.tipologiaCosto || x.f.sottocategoria || "");
    a.appalti.push(x.f.clienteRif || "");
    const data = x.f.dataDocumento.slice(0, 10);
    if (data) {
      if (!a.prima || data < a.prima) a.prima = data;
      if (data > a.ultima) a.ultima = data;
    }
    a.pagato += incassatoRegistrato(x);
    if (!nc) {
      a.nTotale++;
      const residuo = residuoAperto(x);
      if (residuo <= 1 && incassatoRegistrato(x) > 0.005) a.nPagate++;
    }
    if (inFlusso(x)) {
      const residuo = residuoAperto(x);
      a.nAperte++;
      a.aperto += residuo;
      a.modAperte.push(modDich);
      if (x.s.inRitardo) {
        a.nScadute++;
        a.scaduto += residuo;
      }
      const scad = x.s.scadenza.slice(0, 10);
      if (scad && (!a.prossima || scad < a.prossima)) a.prossima = scad;
    }
  }
  const out: RigaFornitore[] = [];
  for (const [k, a] of per) {
    if (a.nAperte === 0) continue;
    const salvato = info.get(k);
    const modalitaAuto = (piuFrequente(a.modAperte.length ? a.modAperte : a.modTutte) ||
      "altro") as ModalitaPagamento;
    const r2 = (n: number) => Math.round(n * 100) / 100;
    out.push({
      chiave: k,
      nome: salvato?.nomeRiga || piuFrequente(a.nomi) || a.nomi[0],
      modalita: salvato?.modalita ?? modalitaAuto,
      modalitaAuto,
      override: salvato?.modalita,
      mista: !salvato?.modalita && new Set(a.modAperte).size > 1,
      referente: salvato?.referente ?? "",
      macrovoce: salvato?.macrovoce ?? "",
      appalto: salvato?.appalto ?? "",
      macrovoceAuto: piuFrequente(a.macro),
      appaltoAuto: piuFrequente(a.appalti),
      scaduto: r2(a.scaduto),
      aperto: r2(a.aperto),
      nScadute: a.nScadute,
      nAperte: a.nAperte,
      nTotale: a.nTotale,
      nPagate: a.nPagate,
      pagato: r2(a.pagato),
      primaFattura: a.prima,
      ultimaFattura: a.ultima,
      prossimaScadenza: a.prossima,
    });
  }
  return out.sort((x, y) => y.aperto - x.aperto || x.nome.localeCompare(y.nome, "it"));
}

// --- Costi fissi nelle "Altre spese" (Simone 29/09, v1.85.0) --------------------
// "Le 4 righe di affitti vanno su costi generali però con costo effettivo,
// non calcolato su media": una riga FlussiCassa genere "fisso" per ogni
// uscita fissa (Title = nome, Importo = importo mensile, Note = parola con
// cui riconoscerla nei movimenti bancari, Mese/MeseFine = validità). I
// movimenti che la riguardano escono dalla media dei costi variabili e la
// voce li conta al loro importo: quello reale se nel mese è già uscito,
// altrimenti quello impostato.

export interface CostoFisso {
  id: string;
  nome: string;
  /** Importo mensile atteso (positivo). */
  importo: number;
  /** Parola cercata (senza maiuscole) in controparte + descrizione del movimento. */
  token: string;
  mese?: string;
  meseFine?: string;
}

export function costiFissiDa(
  righe: readonly {
    id: string;
    genere: string;
    nome: string;
    importo: number;
    note?: string;
    mese?: string;
    meseFine?: string;
  }[],
): CostoFisso[] {
  return righe
    .filter((r) => r.genere === "fisso" && r.nome.trim() && r.importo > 0)
    .map((r) => ({
      id: r.id,
      nome: r.nome.trim(),
      importo: Math.round(r.importo * 100) / 100,
      token: (r.note ?? "").trim().toLowerCase(),
      mese: r.mese,
      meseFine: r.meseFine,
    }))
    .sort((a, b) => a.nome.localeCompare(b.nome, "it"));
}

/** Il costo fisso vale nel mese (finestra da/a facoltativa). */
export function fissoAttivo(f: Pick<CostoFisso, "mese" | "meseFine">, mese: string): boolean {
  if (f.mese && mese < f.mese) return false;
  if (f.meseFine && mese > f.meseFine) return false;
  return true;
}

/** Il movimento è il pagamento di questo costo fisso: uscita, parola chiave
 *  presente e importo vicino a quello atteso (±35%: distingue l'affitto
 *  Zekaj da 2.000 dalle altre disposizioni verso lo stesso nome, 1.000 o
 *  25.000). Tra più candidati nello stesso mese vince il più vicino
 *  all'importo atteso (vedi `piuVicinoAlFisso`). */
export const TOLLERANZA_FISSO = 0.35;

export function matchFisso(
  f: Pick<CostoFisso, "token" | "importo">,
  m: { cliente?: string; descrizione?: string; importo: number },
): boolean {
  if (m.importo >= 0 || !f.token) return false;
  const testo = `${m.cliente ?? ""} ${m.descrizione ?? ""}`.toLowerCase();
  if (!testo.includes(f.token)) return false;
  const a = Math.abs(m.importo);
  return a >= f.importo * (1 - TOLLERANZA_FISSO) && a <= f.importo * (1 + TOLLERANZA_FISSO);
}

/** true se `candidato` è un pagamento più vicino all'importo atteso di
 *  `attuale` (o non c'è ancora un attuale). */
export function piuVicinoAlFisso(
  f: Pick<CostoFisso, "importo">,
  attuale: { importo: number } | undefined,
  candidato: { importo: number },
): boolean {
  if (!attuale) return true;
  return (
    Math.abs(Math.abs(candidato.importo) - f.importo) <
    Math.abs(Math.abs(attuale.importo) - f.importo)
  );
}

// --- Medie senza anomalie (Simone 29/09, v1.86.0) ------------------------------
// "Le anomalie come quelle di questo mese non possono entrare nella media":
// un movimento è ANOMALO quando da solo supera del 25% quanto quella
// tipologia spende in un mese tipico (mediana dei totali mensili dei 6 mesi
// pieni precedenti alla finestra, con un minimo di 1.000 €): un pagamento
// pari al mese tipico è la normale rata mensile, non un'anomalia
// (Consulenze: mese tipico 12.260 → soglia 15.325 → le disposizioni da
// 16.705 e 25.000 sono anomale, la parcella da 9.600 no). Le anomalie contano
// nel REALE del mese corrente (sono soldi usciti) ma non nella media che
// alimenta i mesi futuri. Ogni movimento si può forzare a mano (righe
// FlussiCassa genere "anomalia": Title = chiave movimento, Importo 1 =
// anomalia, 0 = normale).

export const ANOMALIA_MINIMO = 1000;
// 1,25 (1.86.2): con 1,5 la disposizione Zekaj da 16.705 (mese tipico
// Consulenze 12.260 → soglia 18.390) restava "normale" ed entrava nella
// media di ottobre; col 25% la soglia è 15.325 e resta fuori. Una parcella
// pari al mese tipico resta normale.
export const ANOMALIA_MOLTIPLICATORE = 1.25;

export interface MovimentoMedia {
  chiave: string;
  /** Data contabile YYYY-MM-DD. */
  data: string;
  /** Uscite negative (si usa il valore assoluto). */
  importo: number;
  tipologia: string;
  controparte: string;
}

export interface RigaTipologiaMedia {
  tip: string;
  /** Totali dei due mesi pieni SENZA anomalie. */
  m1: number;
  m2: number;
  /** Mese corrente finora, anomalie COMPRESE (è il reale). */
  corrente: number;
  /** Parte anomala del mese corrente. */
  correnteAnomalie: number;
}

export interface MovimentoAnalizzato {
  m: MovimentoMedia;
  mese: string;
  soglia: number;
  anomalia: boolean;
  /** true se decisa dalla regola, false se forzata a mano. */
  auto: boolean;
}

export interface AnalisiMedia {
  /** I due mesi pieni della media. */
  mesi: [string, string];
  meseCorrente: string;
  righe: RigaTipologiaMedia[];
  /** Movimenti dei tre mesi (due pieni + corrente), dal più grande. */
  movimenti: MovimentoAnalizzato[];
  soglie: Map<string, number>;
}

/** Chiave YYYY-MM del mese `delta` mesi prima (delta > 0) di oggi. */
export function meseRelativo(oggiISO: string, delta: number): string {
  const y = Number(oggiISO.slice(0, 4));
  const m = Number(oggiISO.slice(5, 7));
  const d = new Date(Date.UTC(y, m - 1 - delta, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

function mediana(v: number[]): number {
  if (!v.length) return 0;
  const s = [...v].sort((a, b) => a - b);
  const k = Math.floor(s.length / 2);
  return s.length % 2 ? s[k] : (s[k - 1] + s[k]) / 2;
}

export function analizzaMedia(
  movs: readonly MovimentoMedia[],
  oggiISO: string,
  override: ReadonlyMap<string, boolean> = new Map(),
  mesiStorico = 6,
): AnalisiMedia {
  const meseCorrente = oggiISO.slice(0, 7);
  const mesi: [string, string] = [meseRelativo(oggiISO, 2), meseRelativo(oggiISO, 1)];
  const storico = new Set<string>();
  for (let i = 3; i < 3 + mesiStorico; i++) storico.add(meseRelativo(oggiISO, i));
  const r2 = (n: number) => Math.round(n * 100) / 100;
  // Totali mensili storici per tipologia → soglia.
  const totStorico = new Map<string, Map<string, number>>();
  for (const m of movs) {
    if (m.importo >= 0) continue;
    const mese = m.data.slice(0, 7);
    if (!storico.has(mese)) continue;
    const per = totStorico.get(m.tipologia) ?? new Map<string, number>();
    per.set(mese, (per.get(mese) ?? 0) + Math.abs(m.importo));
    totStorico.set(m.tipologia, per);
  }
  const soglie = new Map<string, number>();
  const sogliaDi = (tip: string) => {
    let s = soglie.get(tip);
    if (s == null) {
      const mensili = [...(totStorico.get(tip)?.values() ?? [])].filter((x) => x > 0);
      s = Math.max(
        ANOMALIA_MINIMO,
        Math.round(mediana(mensili) * ANOMALIA_MOLTIPLICATORE * 100) / 100,
      );
      soglie.set(tip, s);
    }
    return s;
  };
  const righe = new Map<string, RigaTipologiaMedia>();
  const movimenti: MovimentoAnalizzato[] = [];
  for (const m of movs) {
    if (m.importo >= 0) continue;
    const mese = m.data.slice(0, 7);
    const idx = mese === meseCorrente ? 2 : mesi.indexOf(mese);
    if (idx < 0) continue;
    const soglia = sogliaDi(m.tipologia);
    const forzata = override.get(m.chiave);
    const anomalia = forzata ?? Math.abs(m.importo) > soglia;
    movimenti.push({ m, mese, soglia, anomalia, auto: forzata == null });
    const r = righe.get(m.tipologia) ?? {
      tip: m.tipologia,
      m1: 0,
      m2: 0,
      corrente: 0,
      correnteAnomalie: 0,
    };
    const a = Math.abs(m.importo);
    if (idx === 2) {
      r.corrente += a;
      if (anomalia) r.correnteAnomalie += a;
    } else if (!anomalia) {
      if (idx === 0) r.m1 += a;
      else r.m2 += a;
    }
    righe.set(m.tipologia, r);
  }
  return {
    mesi,
    meseCorrente,
    righe: [...righe.values()]
      .map((r) => ({
        ...r,
        m1: r2(r.m1),
        m2: r2(r.m2),
        corrente: r2(r.corrente),
        correnteAnomalie: r2(r.correnteAnomalie),
      }))
      .sort((a, b) => b.m1 + b.m2 - (a.m1 + a.m2)),
    movimenti: movimenti.sort((a, b) => Math.abs(b.m.importo) - Math.abs(a.m.importo)),
    soglie,
  };
}

export function giorniNelMese(mese: string): number {
  const [y, m] = mese.split("-").map(Number);
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** Mese corrente delle "Altre spese" variabili: quanto è già uscito più la
 *  media per i giorni che mancano — all'ultimo giorno del mese coincide col
 *  reale, il primo giorno con la media. */
export function proiezioneMeseCorrente(
  realeFinora: number,
  media: number,
  oggiISO: string,
): { proiezione: number; giornoOggi: number; giorniMese: number } {
  const giorniMese = giorniNelMese(oggiISO.slice(0, 7));
  const giornoOggi = Math.min(giorniMese, Math.max(1, Number(oggiISO.slice(8, 10)) || 1));
  const restanti = giorniMese - giornoOggi;
  const proiezione = Math.round((realeFinora + (media * restanti) / giorniMese) * 100) / 100;
  return { proiezione, giornoOggi, giorniMese };
}
