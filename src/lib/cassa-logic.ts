// DR Portal — "CASSA PROPRIETÀ": seconda vista dei Flussi di cassa, logica pura.
//
// Richiesta Simone 19/10/2026, sul modello del file del commercialista
// ("nuovo schema flusso di cassa_071026.xlsx"): non una previsione ma una
// FOTOGRAFIA della cassa a una data, nello STESSO foglio che la proprietà
// legge — colonne Entrate | Note | Saldo | Uscite | 1-5 del mese dopo |
// Standby | Note | Note2 e i totali in fondo con le sue etichette.
//
// Semantica del foglio (ricavata dai fogli "1210" e "1510"):
// - ENTRATE = gli incassi arrivati in banca dal primo del mese (già dentro
//   il saldo) più quelli attesi dai clienti; "standby incasso" (B55) = solo
//   gli attesi, non ancora in banca. Le entrate NON si sommano al saldo.
// - USCITE in tre colonne: da pagare ORA (scadenze del mese, scadute
//   comprese), PROSSIMI giorni (fino alla data limite, di solito il 5 del
//   mese dopo), STANDBY (parcheggiate apposta).
// - saldo dopo incassi e pagamenti = saldo − da pagare; la cella accanto =
//   quello − prossimi; "totale da pagare" = da pagare + standby;
//   "cassa+standby incassi−standby da pagare" = (saldo − ora − prossimi)
//   + standby incasso − standby da pagare.
// Una riga per FORNITORE/CLIENTE (come fa lui), non per fattura: i gruppi
// riproducono i suoi numeri al centesimo (MG Autoparts 6.527,01,
// l'autorimessa 30.599, N.A.B 4.625,20, Califano 2.975,11…).
// Niente medie e niente stime: entra solo quello che ha un documento o un
// importo noto (fatture aperte, stipendi netti dal file paghe, scadenze
// fiscali non "previsionali", costi fissi) più le righe manuali della
// proprietà (divisione utili, compensazioni) e le spese passate da altri
// conti (carta Amex = "CDC" nel suo foglio).
import { residuoAperto, type FatturaConStato } from "./fatture-logic";
import {
  esclusaDaFlussi,
  fissoAttivo,
  matchFisso,
  modalitaFattura,
  type CostoFisso,
  type FornitoreInfo,
} from "./flussi-logic";
import type { ScadenzaFiscale } from "./fiscale-logic";
import { chiaveNome, type StipendiDb } from "./stipendi-logic";

export type SecchioCassa = "ora" | "prossimi" | "standby";
/** Priorità delle cose da fare (agenda): bassa / media / alta. */
export type PrioritaCassa = "l" | "m" | "h";
export type TipoRigaCassa =
  "incasso" | "attiva" | "passiva" | "stipendi" | "fiscale" | "fisso" | "manuale" | "amex";

/** Stato salvato per riga (chiave stabile): colonna scelta, nota, flag. */
export interface StatoRigaCassa {
  /** Colonna forzata a mano (ora / prossimi / standby). */
  b?: SecchioCassa;
  /** Nota libera ("contestare", "compensazione", "sentire Marelli"). */
  n?: string;
  /** Contestata (non più usato dalla vista, conservato). */
  c?: boolean;
  /** Riservato (non più usato: gli attesi sono sempre "standby incasso"). */
  s?: boolean;
  /** Importo corretto a mano (stipendi, fiscale, costi fissi, attesi). */
  i?: number;
  /** Tolta dal foglio. */
  x?: boolean;
  /** Data scelta a mano (entrate attese: quando il cliente ha detto che paga). */
  d?: string;
  /** Priorità della nota (agenda cose da fare). */
  p?: PrioritaCassa;
}

export interface RigaManualeCassa {
  id: string;
  tipo: "manuale" | "amex";
  descrizione: string;
  importo: number;
  data: string; // YYYY-MM-DD (scadenza / addebito)
  secchio: SecchioCassa;
  nota?: string;
  /** true = entrata attesa (es. "Univex 110.000"); default uscita. */
  entrata?: boolean;
  /** Priorità della nota (agenda). */
  p?: PrioritaCassa;
  /** Si ripresenta ogni mese, stesso giorno (stipendi iMile, buoni pasto,
   *  commercialista…): le righe che non hanno una fattura nel portale. */
  ricorrente?: boolean;
}

/** Voce libera dell'agenda (cose da fare non legate a una riga). */
export interface VoceAgenda {
  id: string;
  testo: string;
  p: PrioritaCassa;
  /** Scadenza (YYYY-MM-DD), facoltativa. */
  scad?: string;
  fatta?: boolean;
  creata: string;
}

export interface RigaFotografia {
  k: string;
  t: TipoRigaCassa;
  c: string;
  d: string;
  i: number;
  s: string;
  b: SecchioCassa;
  e: boolean;
  n?: string;
  si?: boolean;
  /** Documenti raggruppati nella riga. */
  nd?: number;
  /** Priorità della nota. */
  p?: PrioritaCassa;
}

export interface TotaliCassa {
  saldo: number;
  /** B2: somma di tutte le entrate elencate (arrivate + attese). */
  entrate: number;
  /** Incassi già in banca dal primo del mese (informativo). */
  entrateRicevute: number;
  /** B55 "standby incasso": entrate attese, non ancora in banca. */
  entrateStandby: number;
  /** E55 "da pagare". */
  usciteOra: number;
  /** F55 "da pagare 1-5/11". */
  usciteProssimi: number;
  /** G55 "standby da pagare". */
  usciteStandby: number;
  /** H55 "totale da pagare" = da pagare + standby. */
  totaleDaPagare: number;
  /** E58 "saldo dopo incassi e pagamenti" = saldo − da pagare. */
  saldoDopoOra: number;
  /** F58 = saldo dopo − prossimi. */
  saldoDopoProssimi: number;
  /** I55 "cassa+standby incassi-standby da pagare". */
  cassaConStandby: number;
  /** Righe fuori orizzonte (scadenza oltre il limite), solo informative. */
  nOltre: number;
  usciteOltre: number;
}

export interface FotografiaCassa {
  id: string;
  data: string;
  limite: string;
  creataIl: string;
  da: string;
  saldo: number;
  totali: TotaliCassa;
  righe: RigaFotografia[];
  nota?: string;
}

export interface CassaDb {
  versione: number;
  /** Giorno del mese dopo fino a cui guardare ("1-5 novembre" → 5). */
  giornoLimite: number;
  stati: Record<string, StatoRigaCassa>;
  manuali: RigaManualeCassa[];
  fotografie: FotografiaCassa[];
  /** Saldo scritto a mano quando la banca non risponde (con la sua data). */
  saldoManuale?: { importo: number; data: string };
  /** Agenda: cose da fare libere, con priorità. */
  agenda?: VoceAgenda[];
  /** Appalti chiusi: ultimo mese di competenza con stipendi (YYYY-MM), per
   *  prefisso del nome appalto. Sovrascrive/aggiunge a CHIUSURE_APPALTI. */
  chiusure?: Record<string, string>;
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

export function emptyCassaDb(): CassaDb {
  return { versione: 0, giornoLimite: 5, stati: {}, manuali: [], fotografie: [] };
}

export interface RigaCassa {
  chiave: string;
  tipo: TipoRigaCassa;
  controparte: string;
  descrizione: string;
  importo: number;
  /** Importo prima della correzione a mano (uguale se non corretto). */
  importoOrigine: number;
  scadenza: string; // YYYY-MM-DD (per gli incassi: data in banca)
  /** Addebito automatico (RID/RiBa), stipendi, fisco: esce senza decidere. */
  certa: boolean;
  entrata: boolean;
  secchio: SecchioCassa;
  /** Colonna che avrebbe senza interventi a mano. */
  secchioAuto: SecchioCassa | "oltre";
  /** Entrata attesa (non in banca): conta in "standby incasso". */
  standbyIncasso: boolean;
  nota: string;
  contestata: boolean;
  appalto?: string;
  /** Documenti raggruppati nella riga (fatture, scadenze). */
  documenti: number;
  /** L'importo si può correggere a mano. */
  editabile: boolean;
  /** Priorità della nota (agenda). */
  priorita?: PrioritaCassa;
  /** La data è stata scelta a mano. */
  dataScelta: boolean;
  /** Modalità di pagamento delle passive quando è un addebito automatico
   *  ("RID", "RIBA"): esce da sola, non si può parcheggiare. */
  modalita?: string;
  /** Stipendi STIMATI (mesi senza ancora i netti di Lucrezia). */
  stima?: boolean;
}

const r2 = (n: number) => Math.round(n * 100) / 100;
const ddmm = (iso: string) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}`;
const chiaveCliente = (s: string) => s.trim().toLowerCase().replace(/\s+/g, " ");

export function fineMese(iso: string): string {
  const [y, m] = iso.slice(0, 7).split("-").map(Number);
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return `${iso.slice(0, 7)}-${String(last).padStart(2, "0")}`;
}

export function meseDopo(yyyymm: string): string {
  const [y, m] = yyyymm.split("-").map(Number);
  return new Date(Date.UTC(y, m, 1)).toISOString().slice(0, 7);
}
export function mesePrima(yyyymm: string): string {
  const [y, m] = yyyymm.split("-").map(Number);
  return new Date(Date.UTC(y, m - 2, 1)).toISOString().slice(0, 7);
}

/** Limite di default: il giorno `giorno` del mese dopo. */
export function limiteDefault(oggi: string, giorno: number): string {
  const g = Math.min(Math.max(1, Math.round(giorno || 5)), 28);
  return `${meseDopo(oggi.slice(0, 7))}-${String(g).padStart(2, "0")}`;
}

export interface StipendiCassaRiga {
  appalto: string;
  meseCompetenza: string;
  mesePagamento: string;
  importo: number;
  dipendenti: number;
  /** Stima sugli ultimi netti disponibili (mese non ancora caricato). */
  stima?: boolean;
}

/** Appalti chiusi (Simone 20/10/2026): ultimo mese di competenza con
 *  stipendi. Postadoc chiuso il 30/09/2026, Ceva il 30/09/2026, Univex
 *  Savona il 31/10/2026. Prefisso del nome appalto, maiuscolo. */
export const CHIUSURE_APPALTI: Record<string, string> = {
  "POSTA DOC": "2026-09",
  POSTADOC: "2026-09",
  CEVA: "2026-09",
  "UNIVEX - SAVONA": "2026-10",
  "UNIVEX SAVONA": "2026-10",
};
export function appaltoChiuso(
  appalto: string,
  meseCompetenza: string,
  chiusure: Record<string, string> = CHIUSURE_APPALTI,
): boolean {
  const a = appalto.trim().toUpperCase();
  for (const [prefisso, ultimo] of Object.entries({ ...CHIUSURE_APPALTI, ...chiusure }))
    if (a.startsWith(prefisso.toUpperCase()) && meseCompetenza > ultimo) return true;
  return false;
}

/** Stipendi ancora da pagare, per appalto e mese: dal file "Stipendi Dr"
 *  (netti), tolte le spunte Pagato della tab Stipendi. Mese di pagamento =
 *  mese dopo la competenza; entrano i mesi con pagamento fino a `fino`
 *  (YYYY-MM) compreso, così si vedono anche gli arretrati (settembre non
 *  pagato a ottobre). Gli arretrati si guardano solo fino al mese di
 *  pagamento PRECEDENTE a quello corrente, e solo se quel mese ha almeno
 *  una spunta Pagato: più indietro le spunte non si usavano o sono parziali
 *  (cessati, liquidazioni) e gennaio-agosto 2026 risorgevano come arretrati
 *  nel foglio della proprietà — visto live il 19/10. */
export function stipendiDaPagare(
  db: StipendiDb,
  fino: string,
  meseCorrente: string = fino,
  chiusure: Record<string, string> = {},
): StipendiCassaRiga[] {
  const out: StipendiCassaRiga[] = [];
  for (const n of db.netti ?? []) {
    const mesePag = meseDopo(n.mese);
    if (mesePag > fino) continue;
    const flag = db.pagatiPerMese?.[n.mese] ?? [];
    if (mesePag < mesePrima(meseCorrente)) continue;
    if (mesePag < meseCorrente && flag.length === 0) continue;
    const pagati = new Set(flag);
    const perAppalto = new Map<string, { importo: number; dipendenti: number }>();
    for (const d of n.dipendenti) {
      if (pagati.has(chiaveNome(d.nome))) continue;
      if (!(d.saldo > 0.005)) continue;
      const k = (d.appalto ?? "").trim() || "Altri";
      const a = perAppalto.get(k) ?? { importo: 0, dipendenti: 0 };
      a.importo = r2(a.importo + d.saldo);
      a.dipendenti++;
      perAppalto.set(k, a);
    }
    for (const [appalto, a] of perAppalto)
      out.push({
        appalto,
        meseCompetenza: n.mese,
        mesePagamento: mesePag,
        importo: a.importo,
        dipendenti: a.dipendenti,
      });
  }
  // STIMA per i mesi di competenza dopo l'ultimo caricato (Simone 20/10:
  // "per i mesi futuri devi fare una stima"): media del netto in busta
  // (stipendio) per appalto sugli ultimi due mesi disponibili, senza gli
  // appalti chiusi. Si corregge a mano come le altre righe.
  const mesi = [...new Set((db.netti ?? []).map((n) => n.mese))].sort();
  const ultimo = mesi[mesi.length - 1];
  if (ultimo) {
    const base = (db.netti ?? []).filter((n) => n.mese === ultimo || n.mese === mesePrima(ultimo));
    for (let mese = meseDopo(ultimo); meseDopo(mese) <= fino; mese = meseDopo(mese)) {
      const perAppalto = new Map<string, { tot: number; mesi: Set<string>; dip: number }>();
      for (const n of base)
        for (const d of n.dipendenti) {
          const k = (d.appalto ?? "").trim() || "Altri";
          if (appaltoChiuso(k, mese, chiusure)) continue;
          const imp = Number(d.stipendio) || 0;
          if (imp <= 0.005) continue;
          const a = perAppalto.get(k) ?? { tot: 0, mesi: new Set<string>(), dip: 0 };
          a.tot += imp;
          a.mesi.add(n.mese);
          a.dip = Math.max(
            a.dip,
            n.dipendenti.filter((x) => ((x.appalto ?? "").trim() || "Altri") === k).length,
          );
          perAppalto.set(k, a);
        }
      for (const [appalto, a] of perAppalto)
        out.push({
          appalto,
          meseCompetenza: mese,
          mesePagamento: meseDopo(mese),
          importo: r2(a.tot / Math.max(1, a.mesi.size)),
          dipendenti: a.dip,
          stima: true,
        });
    }
  }
  return out.sort(
    (a, b) =>
      a.mesePagamento.localeCompare(b.mesePagamento) || a.appalto.localeCompare(b.appalto, "it"),
  );
}

export interface InputCassa {
  /** Data della fotografia (YYYY-MM-DD). */
  oggi: string;
  /** Fino a quando guardare (YYYY-MM-DD), oltre la fine del mese. */
  limite: string;
  saldo: number;
  attive: readonly FatturaConStato[];
  passive: readonly FatturaConStato[];
  esclusioni: readonly { nome: string; mese?: string; meseFine?: string }[];
  infoFornitori: ReadonlyMap<string, FornitoreInfo>;
  fiscale: readonly ScadenzaFiscale[];
  fissi: readonly CostoFisso[];
  /** Movimenti bancari: gli accrediti dal primo del mese sono le ENTRATE;
   *  un costo fisso già uscito nel mese non si ripropone. */
  movimenti: readonly {
    id?: string;
    cliente?: string;
    descrizione?: string;
    importo: number;
    dataContabile: string;
  }[];
  stipendi: readonly StipendiCassaRiga[];
  db: CassaDb;
}

export interface EsitoCassa {
  righe: RigaCassa[];
  /** Fuori orizzonte: scadenza oltre il limite e nessuna scelta a mano. */
  oltre: RigaCassa[];
  totali: TotaliCassa;
}

const MESI_IT = [
  "gen",
  "feb",
  "mar",
  "apr",
  "mag",
  "giu",
  "lug",
  "ago",
  "set",
  "ott",
  "nov",
  "dic",
];
export function etichettaMeseCassa(yyyymm: string): string {
  const [y, m] = yyyymm.split("-");
  return `${MESI_IT[Number(m) - 1] ?? m} ${y}`;
}

const MESI_LUNGHI = [
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
/** Intestazione della colonna dei prossimi giorni come nel foglio: "1-5 novembre". */
export function etichettaProssimi(limite: string): string {
  const g = Number(limite.slice(8, 10));
  const m = MESI_LUNGHI[Number(limite.slice(5, 7)) - 1] ?? limite.slice(5, 7);
  return `1-${g} ${m}`;
}
/** Etichetta corta del totale: "da pagare 1-5/11". */
export function etichettaProssimiBreve(limite: string): string {
  return `1-${Number(limite.slice(8, 10))}/${limite.slice(5, 7)}`;
}

/** Scadenze fiscali "previsionali" (stime del consulente): non entrano. */
export function scadenzaPrevisionale(sc: ScadenzaFiscale): boolean {
  return /previsional/i.test([sc.voce, sc.periodo, sc.modalita, sc.note].join(" "));
}

type Grezza = Omit<
  RigaCassa,
  | "secchio"
  | "secchioAuto"
  | "nota"
  | "contestata"
  | "standbyIncasso"
  | "importoOrigine"
  | "priorita"
  | "dataScelta"
>;

function elencoNumeri(numeri: readonly string[]): string {
  const n = numeri.length;
  const testa = numeri.slice(0, 3).join(", ");
  return n > 3 ? `${testa} +${n - 3}` : testa;
}

export function costruisciCassa(input: InputCassa): EsitoCassa {
  const { oggi, limite, db } = input;
  const fm = fineMese(oggi);
  const meseCorr = oggi.slice(0, 7);
  const inizioMese = `${meseCorr}-01`;
  const stato = (k: string): StatoRigaCassa => db.stati[k] ?? {};
  const secchioAuto = (scadenza: string): SecchioCassa | "oltre" =>
    scadenza <= fm ? "ora" : scadenza <= limite ? "prossimi" : "oltre";

  const grezze: Grezza[] = [];

  // ENTRATE arrivate: accrediti in banca dal primo del mese alla data.
  const incassi = input.movimenti
    .filter((m) => m.importo > 0.005)
    .map((m) => ({ ...m, d: m.dataContabile.slice(0, 10) }))
    .filter((m) => m.d >= inizioMese && m.d <= oggi)
    .sort((a, b) => a.d.localeCompare(b.d));
  for (const m of incassi) {
    const nome = (m.cliente || m.descrizione || "").trim();
    grezze.push({
      chiave: `inc|${m.id ?? `${m.d}|${m.importo}|${chiaveCliente(nome)}`}`,
      tipo: "incasso",
      controparte: nome,
      descrizione: ddmm(m.d),
      importo: r2(m.importo),
      scadenza: m.d,
      certa: true,
      entrata: true,
      documenti: 1,
      editabile: false,
    });
  }
  // ENTRATE attese: fatture attive aperte con scadenza entro il limite,
  // una riga per cliente (l'importo si corregge a mano con quello che il
  // cliente annuncia). Esclusioni dei Flussi rispettate.
  const gruppiAtt = new Map<
    string,
    { nome: string; importo: number; n: number; scad: string; numeri: string[] }
  >();
  for (const x of input.attive) {
    const res = residuoAperto(x);
    if (res <= 1 || !x.s.scadenza) continue;
    const scad = x.s.scadenza.slice(0, 10);
    if (scad > limite) continue;
    if (esclusaDaFlussi(input.esclusioni, x.f.cliente, scad.slice(0, 7))) continue;
    const k = chiaveCliente(x.f.cliente);
    const g = gruppiAtt.get(k) ?? { nome: x.f.cliente, importo: 0, n: 0, scad, numeri: [] };
    g.importo = r2(g.importo + res);
    g.n++;
    if (scad < g.scad) g.scad = scad;
    g.numeri.push(x.f.numero);
    gruppiAtt.set(k, g);
  }
  for (const [k, g] of gruppiAtt)
    grezze.push({
      chiave: `attc|${meseCorr}|${k}`,
      tipo: "attiva",
      controparte: g.nome,
      descrizione: `${g.n} fatt. ${elencoNumeri(g.numeri)} · scad. ${ddmm(g.scad)}`,
      importo: g.importo,
      scadenza: g.scad,
      certa: false,
      entrata: true,
      documenti: g.n,
      editabile: true,
    });
  // USCITE: fatture passive aperte, una riga per fornitore e colonna
  // automatica (RID/RiBa = certe). Le decisioni (standby, nota) restano
  // sul fornitore anche quando arrivano fatture nuove.
  const gruppiPas = new Map<
    string,
    {
      nome: string;
      auto: SecchioCassa | "oltre";
      importo: number;
      n: number;
      scad: string;
      mods: Set<string>;
      numeri: string[];
    }
  >();
  for (const x of input.passive) {
    const res = residuoAperto(x);
    if (res <= 1 || !x.s.scadenza) continue;
    const scad = x.s.scadenza.slice(0, 10);
    if (esclusaDaFlussi(input.esclusioni, x.f.cliente, scad.slice(0, 7))) continue;
    const mod = modalitaFattura(x.f, input.infoFornitori);
    const auto = secchioAuto(scad);
    const k = `${chiaveCliente(x.f.cliente)}|${auto}`;
    const g = gruppiPas.get(k) ?? {
      nome: x.f.cliente,
      auto,
      importo: 0,
      n: 0,
      scad,
      mods: new Set<string>(),
      numeri: [],
    };
    g.importo = r2(g.importo + res);
    g.n++;
    if (scad < g.scad) g.scad = scad;
    g.mods.add(mod);
    g.numeri.push(x.f.numero);
    gruppiPas.set(k, g);
  }
  for (const [k, g] of gruppiPas) {
    const certa = g.mods.size > 0 && ![...g.mods].includes("altro");
    const mod = certa ? [...g.mods][0].toUpperCase() : "";
    grezze.push({
      chiave: `pasg|${k}`,
      tipo: "passiva",
      controparte: g.nome,
      descrizione: `${g.n} fatt. ${elencoNumeri(g.numeri)} · scad. ${ddmm(g.scad)}${mod ? ` · ${mod}` : ""}`,
      importo: g.importo,
      scadenza: g.scad,
      certa,
      entrata: false,
      documenti: g.n,
      editabile: false,
      modalita: mod || undefined,
    });
  }
  // Stipendi: netti non pagati per appalto, il 10 del mese di pagamento.
  for (const s of input.stipendi) {
    grezze.push({
      chiave: `stip|${s.meseCompetenza}|${s.appalto}`,
      tipo: "stipendi",
      controparte: `Stipendi ${s.appalto}`,
      descrizione: `${s.stima ? "stima · " : ""}${etichettaMeseCassa(s.meseCompetenza)} · ${s.dipendenti} dip.`,
      importo: s.importo,
      scadenza: `${s.mesePagamento}-10`,
      certa: true,
      entrata: false,
      appalto: s.appalto,
      documenti: s.dipendenti,
      editabile: true,
      stima: s.stima,
    });
  }
  // Fiscale: scadenze non pagate (anche scadute), senza le "previsionali",
  // una riga per colonna come il suo "Fiscale + BPM".
  const gruppiFis = new Map<
    SecchioCassa | "oltre",
    { importo: number; n: number; scad: string; voci: string[]; fin: boolean }
  >();
  for (const sc of input.fiscale) {
    if (sc.pagato || !sc.dataPagamento || scadenzaPrevisionale(sc)) continue;
    const scad = sc.dataPagamento.slice(0, 10);
    const auto = secchioAuto(scad);
    const g = gruppiFis.get(auto) ?? { importo: 0, n: 0, scad, voci: [], fin: false };
    g.importo = r2(g.importo + Math.abs(sc.importo));
    g.n++;
    if (scad < g.scad) g.scad = scad;
    g.voci.push([sc.voce, sc.periodo].filter(Boolean).join(" "));
    if (sc.categoria === "finanziamento") g.fin = true;
    gruppiFis.set(auto, g);
  }
  for (const [auto, g] of gruppiFis)
    grezze.push({
      chiave: `fiscg|${meseCorr}|${auto}`,
      tipo: "fiscale",
      controparte: g.fin ? "Fiscale + BPM" : "Fiscale",
      descrizione: elencoNumeri(g.voci),
      importo: g.importo,
      scadenza: g.scad,
      certa: true,
      entrata: false,
      documenti: g.n,
      editabile: true,
    });
  // Costi fissi del mese: se è già uscito un movimento che li somiglia, no.
  const moviMese = input.movimenti.filter((m) => m.dataContabile.slice(0, 7) === meseCorr);
  for (const f of input.fissi) {
    if (!fissoAttivo(f, meseCorr)) continue;
    if (moviMese.some((m) => matchFisso(f, m))) continue;
    grezze.push({
      chiave: `fisso|${f.id}|${meseCorr}`,
      tipo: "fisso",
      controparte: f.nome,
      descrizione: `costo fisso ${etichettaMeseCassa(meseCorr)}`,
      importo: f.importo,
      scadenza: fm,
      certa: false,
      entrata: false,
      documenti: 1,
      editabile: true,
    });
  }
  // Righe manuali (proprietà) e carta Amex ("CDC"). Una riga RICORRENTE
  // si ripresenta ogni mese allo stesso giorno (dal mese della sua data in
  // poi), finché non la si elimina.
  for (const m of db.manuali) {
    let data = m.data || oggi;
    if (m.ricorrente && data.slice(0, 7) < meseCorr) {
      const giorno = Math.min(Number(data.slice(8, 10)) || 1, Number(fm.slice(8, 10)));
      data = `${meseCorr}-${String(giorno).padStart(2, "0")}`;
    }
    grezze.push({
      chiave: `man|${m.id}`,
      tipo: m.tipo,
      controparte: m.descrizione,
      descrizione: m.tipo === "amex" ? "CDC" : m.ricorrente ? "ogni mese" : "",
      importo: r2(Math.abs(m.importo)),
      scadenza: data,
      certa: m.tipo === "amex",
      entrata: Boolean(m.entrata),
      documenti: 1,
      editabile: true,
    });
  }

  const righe: RigaCassa[] = [];
  const oltre: RigaCassa[] = [];
  for (const g of grezze) {
    const st = stato(g.chiave);
    if (st.x) continue;
    const manuale = db.manuali.find((m) => `man|${m.id}` === g.chiave);
    const auto: SecchioCassa | "oltre" =
      g.tipo === "incasso" || g.tipo === "attiva" ? "ora" : secchioAuto(g.scadenza);
    const secchio: SecchioCassa =
      st.b ?? manuale?.secchio ?? (auto === "oltre" ? "prossimi" : auto);
    const importo = g.editabile && st.i != null && st.i >= 0 ? r2(st.i) : g.importo;
    // Data scelta a mano: solo per le entrate attese (quando il cliente ha
    // detto che paga) — alimenta il calendario entrate.
    const dataScelta = Boolean(g.entrata && g.tipo !== "incasso" && st.d);
    const riga: RigaCassa = {
      ...g,
      importoOrigine: g.importo,
      importo,
      scadenza: dataScelta ? st.d! : g.scadenza,
      secchio,
      secchioAuto: auto,
      standbyIncasso: g.entrata && g.tipo !== "incasso",
      nota: st.n ?? manuale?.nota ?? "",
      contestata: Boolean(st.c),
      priorita: st.p ?? manuale?.p,
      dataScelta,
    };
    if (auto === "oltre" && !st.b && !manuale) oltre.push(riga);
    else righe.push(riga);
  }
  // Ordine del foglio: a sinistra gli incassi per data poi gli attesi; a
  // destra per scadenza e nome, a prescindere dalla colonna (così una riga
  // spostata di colonna non cambia posto).
  const tipoOrd = (r: RigaCassa) => (r.tipo === "incasso" ? 0 : 1);
  righe.sort(
    (a, b) =>
      Number(b.entrata) - Number(a.entrata) ||
      tipoOrd(a) - tipoOrd(b) ||
      a.scadenza.localeCompare(b.scadenza) ||
      a.controparte.localeCompare(b.controparte, "it"),
  );
  oltre.sort(
    (a, b) =>
      a.scadenza.localeCompare(b.scadenza) || a.controparte.localeCompare(b.controparte, "it"),
  );

  const somma = (xs: readonly RigaCassa[]) => r2(xs.reduce((s, r) => s + r.importo, 0));
  const entrate = righe.filter((r) => r.entrata);
  const uscite = righe.filter((r) => !r.entrata);
  const entrateRicevute = somma(entrate.filter((r) => !r.standbyIncasso));
  const entrateStandby = somma(entrate.filter((r) => r.standbyIncasso));
  const usciteOra = somma(uscite.filter((r) => r.secchio === "ora"));
  const usciteProssimi = somma(uscite.filter((r) => r.secchio === "prossimi"));
  const usciteStandby = somma(uscite.filter((r) => r.secchio === "standby"));
  const saldoDopoOra = r2(input.saldo - usciteOra);
  const saldoDopoProssimi = r2(saldoDopoOra - usciteProssimi);
  const totali: TotaliCassa = {
    saldo: r2(input.saldo),
    entrate: r2(entrateRicevute + entrateStandby),
    entrateRicevute,
    entrateStandby,
    usciteOra,
    usciteProssimi,
    usciteStandby,
    totaleDaPagare: r2(usciteOra + usciteStandby),
    saldoDopoOra,
    saldoDopoProssimi,
    cassaConStandby: r2(saldoDopoProssimi + entrateStandby - usciteStandby),
    nOltre: oltre.length,
    usciteOltre: somma(oltre.filter((r) => !r.entrata)),
  };
  return { righe, oltre, totali };
}

export function creaFotografia(
  esito: EsitoCassa,
  p: { data: string; limite: string; da: string; nota?: string; id?: string },
): FotografiaCassa {
  return {
    id: p.id ?? `F${Date.now().toString(36)}`,
    data: p.data,
    limite: p.limite,
    creataIl: new Date().toISOString(),
    da: p.da,
    saldo: esito.totali.saldo,
    totali: esito.totali,
    nota: p.nota,
    righe: esito.righe.map((r) => ({
      k: r.chiave,
      t: r.tipo,
      c: r.controparte,
      d: r.descrizione,
      i: r.importo,
      s: r.scadenza,
      b: r.secchio,
      e: r.entrata,
      n: r.nota || undefined,
      si: r.standbyIncasso || undefined,
      nd: r.documenti > 1 ? r.documenti : undefined,
      p: r.priorita,
    })),
  };
}

/** Confronto tra due fotografie: righe nuove, sparite, cambiate di importo o colonna. */
export function confrontaFotografie(
  prima: FotografiaCassa,
  dopo: FotografiaCassa,
): {
  nuove: RigaFotografia[];
  sparite: RigaFotografia[];
  cambiate: { prima: RigaFotografia; dopo: RigaFotografia }[];
} {
  const a = new Map(prima.righe.map((r) => [r.k, r]));
  const b = new Map(dopo.righe.map((r) => [r.k, r]));
  const nuove = dopo.righe.filter((r) => !a.has(r.k));
  const sparite = prima.righe.filter((r) => !b.has(r.k));
  const cambiate: { prima: RigaFotografia; dopo: RigaFotografia }[] = [];
  for (const [k, r] of b) {
    const p = a.get(k);
    if (p && (Math.abs(p.i - r.i) > 0.005 || p.b !== r.b || Boolean(p.si) !== Boolean(r.si)))
      cambiate.push({ prima: p, dopo: r });
  }
  return { nuove, sparite, cambiate };
}

/** Calendario delle entrate attese: per data (scelta a mano o scadenza),
 *  con il progressivo — "calendario entrate futuro, mettere date" (Simone
 *  20/10/2026). */
export interface VoceCalendario {
  chiave: string;
  data: string;
  nome: string;
  importo: number;
  progressivo: number;
  dataScelta: boolean;
}
export function calendarioEntrate(righe: readonly RigaCassa[]): VoceCalendario[] {
  const attese = righe
    .filter((r) => r.entrata && r.standbyIncasso && r.importo > 0.005)
    .sort(
      (a, b) =>
        a.scadenza.localeCompare(b.scadenza) || a.controparte.localeCompare(b.controparte, "it"),
    );
  let prog = 0;
  return attese.map((r) => {
    prog = r2(prog + r.importo);
    return {
      chiave: r.chiave,
      data: r.scadenza,
      nome: r.controparte,
      importo: r.importo,
      progressivo: prog,
      dataScelta: r.dataScelta,
    };
  });
}

/** Agenda: le note con priorità scritte sulle righe più le voci libere,
 *  ordinate per priorità (alta prima) e data. */
export interface VoceAgendaVista {
  id: string;
  testo: string;
  p: PrioritaCassa;
  data?: string;
  /** Riga del foglio a cui è legata (vuoto = voce libera). */
  riga?: string;
  rigaChiave?: string;
  fatta: boolean;
}
const ORD_P: Record<PrioritaCassa, number> = { h: 0, m: 1, l: 2 };
export function agendaCassa(
  righe: readonly RigaCassa[],
  libere: readonly VoceAgenda[],
): VoceAgendaVista[] {
  const out: VoceAgendaVista[] = [];
  for (const r of righe)
    if (r.priorita)
      out.push({
        id: `r|${r.chiave}`,
        testo: r.nota || r.controparte,
        p: r.priorita,
        data: r.scadenza,
        riga: r.controparte,
        rigaChiave: r.chiave,
        fatta: false,
      });
  for (const v of libere)
    out.push({ id: v.id, testo: v.testo, p: v.p, data: v.scad, fatta: Boolean(v.fatta) });
  return out.sort(
    (a, b) =>
      Number(a.fatta) - Number(b.fatta) ||
      ORD_P[a.p] - ORD_P[b.p] ||
      (a.data ?? "9999").localeCompare(b.data ?? "9999") ||
      a.testo.localeCompare(b.testo, "it"),
  );
}
