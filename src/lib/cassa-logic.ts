// DR Portal — "CASSA PROPRIETÀ": seconda vista dei Flussi di cassa, logica pura.
//
// Richiesta Simone 19/10/2026, sul modello del file del commercialista
// ("nuovo schema flusso di cassa_071026.xlsx"): non una previsione ma una
// FOTOGRAFIA della cassa a una data — saldo in banca, incassi attesi
// (certi o in standby), uscite in tre secchi: da pagare ORA (scadenze del
// mese, scadute comprese), PROSSIMI giorni (fino a una data a scelta, di
// solito il 5 del mese dopo), STANDBY (parcheggiate apposta, col motivo).
// Niente medie e niente stime: entra solo quello che ha un documento o un
// importo noto (fatture aperte, stipendi netti dal file paghe, scadenze
// fiscali, costi fissi) più le righe manuali della proprietà (divisione
// utili, compensazioni) e le spese passate da altri conti (carta Amex).
// Ogni fotografia si può salvare e rileggere: come i fogli "1210"/"1510".
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
export type TipoRigaCassa =
  "attiva" | "passiva" | "stipendi" | "fiscale" | "fisso" | "manuale" | "amex";

/** Stato salvato per riga (chiave stabile): secchio scelto, nota, flag. */
export interface StatoRigaCassa {
  /** Secchio forzato a mano. */
  b?: SecchioCassa;
  /** Nota libera ("contestare", "compensazione", "sentire Marelli"). */
  n?: string;
  /** Contestata. */
  c?: boolean;
  /** Standby incasso (solo entrate). */
  s?: boolean;
  /** Importo corretto a mano (stipendi, costi fissi). */
  i?: number;
  /** Tolta dalla vista. */
  x?: boolean;
}

export interface RigaManualeCassa {
  id: string;
  tipo: "manuale" | "amex";
  descrizione: string;
  importo: number;
  data: string; // YYYY-MM-DD (scadenza / addebito)
  secchio: SecchioCassa;
  nota?: string;
  /** true = entrata (es. rimborso); default uscita. */
  entrata?: boolean;
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
}

export interface TotaliCassa {
  saldo: number;
  entrateCerte: number;
  entrateStandby: number;
  usciteOra: number;
  usciteProssimi: number;
  usciteStandby: number;
  /** saldo − uscite ora ("saldo dopo pagamenti" del commercialista). */
  saldoDopoOra: number;
  /** saldo − ora − prossimi. */
  saldoDopoProssimi: number;
  /** saldo − ora − prossimi + entrate certe. */
  cassaFinale: number;
  /** cassaFinale + entrate in standby − uscite in standby. */
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
  scadenza: string; // YYYY-MM-DD
  /** Addebito automatico (RID/RiBa), stipendi, fisco: esce senza decidere. */
  certa: boolean;
  entrata: boolean;
  secchio: SecchioCassa;
  /** Secchio che avrebbe senza interventi a mano. */
  secchioAuto: SecchioCassa | "oltre";
  standbyIncasso: boolean;
  nota: string;
  contestata: boolean;
  appalto?: string;
}

const r2 = (n: number) => Math.round(n * 100) / 100;

export function fineMese(iso: string): string {
  const [y, m] = iso.slice(0, 7).split("-").map(Number);
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return `${iso.slice(0, 7)}-${String(last).padStart(2, "0")}`;
}

export function meseDopo(yyyymm: string): string {
  const [y, m] = yyyymm.split("-").map(Number);
  return new Date(Date.UTC(y, m, 1)).toISOString().slice(0, 7);
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
}

/** Stipendi ancora da pagare, per appalto e mese: dal file "Stipendi Dr"
 *  (netti), tolte le spunte Pagato della tab Stipendi. Mese di pagamento =
 *  mese dopo la competenza; entrano i mesi con pagamento fino a `fino`
 *  (YYYY-MM) compreso, così si vedono anche gli arretrati (settembre non
 *  pagato a ottobre). */
export function stipendiDaPagare(db: StipendiDb, fino: string): StipendiCassaRiga[] {
  const out: StipendiCassaRiga[] = [];
  for (const n of db.netti ?? []) {
    const mesePag = meseDopo(n.mese);
    if (mesePag > fino) continue;
    const pagati = new Set(db.pagatiPerMese?.[n.mese] ?? []);
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
  /** Movimenti del mese corrente: un costo fisso già uscito non si ripropone. */
  movimenti: readonly {
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

export function costruisciCassa(input: InputCassa): EsitoCassa {
  const { oggi, limite, db } = input;
  const fm = fineMese(oggi);
  const meseCorr = oggi.slice(0, 7);
  const stato = (k: string): StatoRigaCassa => db.stati[k] ?? {};
  const secchioAuto = (scadenza: string): SecchioCassa | "oltre" =>
    scadenza <= fm ? "ora" : scadenza <= limite ? "prossimi" : "oltre";

  const grezze: Omit<
    RigaCassa,
    "secchio" | "secchioAuto" | "nota" | "contestata" | "standbyIncasso" | "importoOrigine"
  >[] = [];

  // Entrate: fatture attive aperte con scadenza (esclusioni dei Flussi).
  for (const x of input.attive) {
    const res = residuoAperto(x);
    if (res <= 1 || !x.s.scadenza) continue;
    const scad = x.s.scadenza.slice(0, 10);
    if (esclusaDaFlussi(input.esclusioni, x.f.cliente, scad.slice(0, 7))) continue;
    grezze.push({
      chiave: `att|${x.f.nomeFile}`,
      tipo: "attiva",
      controparte: x.f.cliente,
      descrizione: `${x.f.numero} del ${x.f.dataDocumento.slice(0, 10).split("-").reverse().join("/")}`,
      importo: r2(res),
      scadenza: scad,
      certa: false,
      entrata: true,
    });
  }
  // Uscite: fatture passive aperte (RID/RiBa = certe).
  for (const x of input.passive) {
    const res = residuoAperto(x);
    if (res <= 1 || !x.s.scadenza) continue;
    const scad = x.s.scadenza.slice(0, 10);
    if (esclusaDaFlussi(input.esclusioni, x.f.cliente, scad.slice(0, 7))) continue;
    const mod = modalitaFattura(x.f, input.infoFornitori);
    grezze.push({
      chiave: `pas|${x.f.nomeFile}`,
      tipo: "passiva",
      controparte: x.f.cliente,
      descrizione: `${x.f.numero} del ${x.f.dataDocumento.slice(0, 10).split("-").reverse().join("/")}${mod !== "altro" ? ` · ${mod.toUpperCase()}` : ""}`,
      importo: r2(res),
      scadenza: scad,
      certa: mod !== "altro",
      entrata: false,
    });
  }
  // Stipendi: netti non pagati per appalto, il 10 del mese di pagamento.
  for (const s of input.stipendi) {
    grezze.push({
      chiave: `stip|${s.meseCompetenza}|${s.appalto}`,
      tipo: "stipendi",
      controparte: `Stipendi ${s.appalto}`,
      descrizione: `${etichettaMeseCassa(s.meseCompetenza)} · ${s.dipendenti} dip.`,
      importo: s.importo,
      scadenza: `${s.mesePagamento}-10`,
      certa: true,
      entrata: false,
      appalto: s.appalto,
    });
  }
  // Fiscale: scadenze non pagate (anche scadute).
  for (const sc of input.fiscale) {
    if (sc.pagato || !sc.dataPagamento) continue;
    grezze.push({
      chiave: `fisc|${sc.id ?? `${sc.voce}|${sc.dataPagamento}|${sc.importo}`}`,
      tipo: "fiscale",
      controparte: "Fiscale",
      descrizione: [sc.voce, sc.periodo, sc.anno, sc.modalita].filter(Boolean).join(" "),
      importo: r2(Math.abs(sc.importo)),
      scadenza: sc.dataPagamento.slice(0, 10),
      certa: true,
      entrata: false,
    });
  }
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
    });
  }
  // Righe manuali (proprietà) e Amex.
  for (const m of db.manuali) {
    grezze.push({
      chiave: `man|${m.id}`,
      tipo: m.tipo,
      controparte: m.tipo === "amex" ? "Carta Amex" : "Manuale",
      descrizione: m.descrizione,
      importo: r2(Math.abs(m.importo)),
      scadenza: m.data || oggi,
      certa: m.tipo === "amex",
      entrata: Boolean(m.entrata),
    });
  }

  const righe: RigaCassa[] = [];
  const oltre: RigaCassa[] = [];
  for (const g of grezze) {
    const st = stato(g.chiave);
    if (st.x) continue;
    const manuale = db.manuali.find((m) => `man|${m.id}` === g.chiave);
    const auto = secchioAuto(g.scadenza);
    const secchio: SecchioCassa =
      st.b ?? manuale?.secchio ?? (auto === "oltre" ? "prossimi" : auto);
    const importo = st.i != null && st.i >= 0 ? r2(st.i) : g.importo;
    const riga: RigaCassa = {
      ...g,
      importoOrigine: g.importo,
      importo,
      secchio,
      secchioAuto: auto,
      standbyIncasso: Boolean(st.s),
      nota: st.n ?? manuale?.nota ?? "",
      contestata: Boolean(st.c),
    };
    if (auto === "oltre" && !st.b && !manuale) oltre.push(riga);
    else righe.push(riga);
  }
  const ordine: Record<SecchioCassa, number> = { ora: 0, prossimi: 1, standby: 2 };
  righe.sort(
    (a, b) =>
      Number(b.entrata) - Number(a.entrata) ||
      ordine[a.secchio] - ordine[b.secchio] ||
      a.scadenza.localeCompare(b.scadenza) ||
      a.controparte.localeCompare(b.controparte, "it"),
  );
  oltre.sort((a, b) => a.scadenza.localeCompare(b.scadenza));

  const somma = (xs: readonly RigaCassa[]) => r2(xs.reduce((s, r) => s + r.importo, 0));
  const entrate = righe.filter((r) => r.entrata);
  const uscite = righe.filter((r) => !r.entrata);
  const entrateCerte = somma(entrate.filter((r) => !r.standbyIncasso && r.secchio !== "standby"));
  const entrateStandby = somma(entrate.filter((r) => r.standbyIncasso || r.secchio === "standby"));
  const usciteOra = somma(uscite.filter((r) => r.secchio === "ora"));
  const usciteProssimi = somma(uscite.filter((r) => r.secchio === "prossimi"));
  const usciteStandby = somma(uscite.filter((r) => r.secchio === "standby"));
  const saldoDopoOra = r2(input.saldo - usciteOra);
  const saldoDopoProssimi = r2(saldoDopoOra - usciteProssimi);
  const cassaFinale = r2(saldoDopoProssimi + entrateCerte);
  const totali: TotaliCassa = {
    saldo: r2(input.saldo),
    entrateCerte,
    entrateStandby,
    usciteOra,
    usciteProssimi,
    usciteStandby,
    saldoDopoOra,
    saldoDopoProssimi,
    cassaFinale,
    cassaConStandby: r2(cassaFinale + entrateStandby - usciteStandby),
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
    })),
  };
}

/** Confronto tra due fotografie: righe nuove, sparite, cambiate di importo o secchio. */
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
