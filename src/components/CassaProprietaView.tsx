// DR Portal — Finanza → Flussi → vista "Cassa proprietà".
// IL FOGLIO del commercialista, tale e quale: colonne Entrate | Note |
// Saldo | Uscite | 1-5 del mese dopo | Standby | Note | Note2, i totali in
// fondo con le sue etichette, le fotografie come le linguette dei fogli
// Excel ("1210"/"1510"). Niente cruscotti: la proprietà legge solo questo.
// Si lavora dentro le celle: clic su una cella vuota delle tre colonne per
// spostare un'uscita, clic sull'importo per correggerlo (stipendi, fiscale,
// fissi, attesi), Note2 è una cella scrivibile, la × a bordo riga toglie la
// riga. Logica in lib/cassa-logic; dati in FinanzaData/cassa.json.
import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { toast } from "sonner";
import {
  CalendarDays,
  Camera,
  Filter,
  ListTodo,
  Loader2,
  Lock,
  Pencil,
  Plus,
  Search,
  StickyNote,
  X,
} from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import {
  SchedaClienteCassa,
  SchedaFiscaleCassa,
  SchedaGruppoCassa,
} from "@/components/CassaSchede";
import type { FatturaConStato } from "@/lib/fatture-logic";
import type {
  CostoFisso,
  FornitoreInfo,
  ParcoPerRipartizione,
  QuotaAppalto,
} from "@/lib/flussi-logic";
import type { ScadenzaFiscale } from "@/lib/fiscale-logic";
import type { SpMovimento } from "@/lib/sharepoint.server";
import { spEbSaldoCache } from "@/lib/sharepoint.functions";
import { spStipendiGet } from "@/lib/stipendi.functions";
import { spCassaGet, spCassaSalva } from "@/lib/cassa.functions";
import {
  COLONNE_USCITE,
  FILTRI_VUOTI,
  SENZA,
  bloccoEntrata,
  clientiDi,
  distinti,
  filtraEntrate,
  filtraUscite,
  filtroEntrateAttivo,
  filtroUsciteAttivo,
  sommaCelle,
  tipologiaDi,
  type FiltriFoglio,
} from "@/lib/cassa-filtri";
import type { StipendiDb } from "@/lib/stipendi-logic";
import {
  agendaCassa,
  calendarioEntrate,
  costruisciCassa,
  creaFotografia,
  emptyCassaDb,
  etichettaProssimi,
  etichettaProssimiBreve,
  limiteDefault,
  stipendiDaPagare,
  tokenGruppi,
  type CassaDb,
  type FotografiaCassa,
  type PrioritaCassa,
  type RigaCassa,
  type RigaManualeCassa,
  type SecchioCassa,
  type StatoRigaCassa,
  type TipoRigaCassa,
  type TotaliCassa,
  type VoceAgenda,
  type VoceAgendaVista,
  type VoceDettaglioCassa,
} from "@/lib/cassa-logic";

type TFn = ReturnType<typeof useLang>["t"];

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtData = (iso: string) => (iso ? iso.slice(0, 10).split("-").reverse().join("/") : "—");
const ddmm = (iso: string) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : "");
/** Con l'anno a due cifre: per le attese ("entranti metti anche anno", Simone 02/11/2026). */
const ddmmyy = (iso: string) =>
  iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(2, 4)}` : "";
const oggiISO = () => new Date().toISOString().slice(0, 10);
const parseNum = (s: string): number | null => {
  const n = Number(s.trim().replace(/\./g, "").replace(",", "."));
  return s.trim() && Number.isFinite(n) ? n : null;
};
const inputCls =
  "rounded-lg border border-border bg-background px-2 py-1 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";
/** Ultimo saldo banca visto in questa sessione: si dipinge subito al rientro
 *  nella vista, poi si riaggiorna dalla cache del server. */
let saldoMemo: { saldo: number; al?: string } | null = null;

/** Riga del foglio, uguale per la vista viva e per una fotografia. */
interface Cella {
  chiave: string;
  tipo: TipoRigaCassa;
  nome: string;
  dettaglio: string;
  importo: number;
  importoOrigine: number;
  secchio: SecchioCassa;
  entrata: boolean;
  attesa: boolean;
  /** Entrata attesa messa in standby (blocco "Standby" delle entrate). */
  parcheggiata: boolean;
  /** Entrata attesa da sollecitare (blocco "Da sollecitare"). */
  daSollecitare: boolean;
  nota: string;
  data: string;
  certa: boolean;
  documenti: number;
  editabile: boolean;
  priorita?: PrioritaCassa;
  dataScelta: boolean;
  modalita?: string;
  stima?: boolean;
  /** Per raggruppare le uscite per tipologia o per cliente. */
  tipologia?: string;
  cliente?: string;
  /** Righe di gruppo: i fornitori dentro, per aprirli con un clic. */
  membri?: { nome: string; importo: number; tipo: TipoRigaCassa }[];
  /** Promemoria della scheda fornitore (icona col testo al passaggio). */
  promemoria?: string;
  /** Scadenze dentro la riga (Fiscale/IVA/Banca): per il popup. */
  voci?: VoceDettaglioCassa[];
  /** Quote per appalto del fornitore: "Uscite per cliente" divide la riga. */
  quote?: QuotaAppalto[];
  /** Entrate unite per gruppo: le società dentro la riga. */
  membriClienti?: string[];
}

/** Vista raggruppata delle uscite (Simone 02/11/2026: "raggruppare le spese
 *  per tipologia o per cliente"): una riga per gruppo e colonna, in sola
 *  lettura (chiave "grp|…"), con l'elenco dei fornitori nel dettaglio. */
type Raggruppa = "fornitore" | "tipologia" | "cliente";
function raggruppaUscite(uscite: readonly Cella[], per: Raggruppa, t: TFn): Cella[] {
  if (per === "fornitore") return [...uscite];
  const gruppi = new Map<string, Cella & { nomi: string[] }>();
  for (const u of uscite) {
    // Per cliente, un fornitore ripartito su più appalti (quote della scheda
    // o dai mezzi) si divide: ogni appalto prende la sua percentuale —
    // Simone 09/11/2026 ("dividere meglio le spese").
    const parti =
      per === "cliente" && u.quote && u.quote.length > 1
        ? u.quote.map((q) => ({
            g: q.appalto.trim() || t("cpr.senzaCliente"),
            importo: Math.round(u.importo * q.quota) / 100,
            nome: `${u.nome} ${q.quota}%`,
          }))
        : [
            {
              g:
                (per === "tipologia" ? u.tipologia : u.cliente)?.trim() ||
                t(per === "tipologia" ? "cpr.senzaTipologia" : "cpr.senzaCliente"),
              importo: u.importo,
              nome: u.nome,
            },
          ];
    for (const p of parti) {
      const g = p.g;
      const k = `grp|${per}|${g.toLowerCase()}|${u.secchio}`;
      const r = gruppi.get(k) ?? {
        ...u,
        chiave: k,
        tipo: "passiva",
        nome: g,
        dettaglio: "",
        importo: 0,
        importoOrigine: 0,
        nota: "",
        certa: true,
        documenti: 0,
        editabile: false,
        priorita: undefined,
        dataScelta: false,
        modalita: undefined,
        stima: false,
        membri: [],
        nomi: [],
      };
      r.importo = Math.round((r.importo + p.importo) * 100) / 100;
      r.importoOrigine = r.importo;
      r.documenti += 1;
      if (u.data < r.data) r.data = u.data;
      r.nomi.push(`${p.nome} ${fmt(p.importo)}`);
      r.membri?.push({ nome: u.nome, importo: p.importo, tipo: u.tipo });
      gruppi.set(k, r);
    }
  }
  return [...gruppi.values()]
    .map((r) => ({ ...r, dettaglio: r.nomi.join(" · ") }))
    .sort((a, b) => b.importo - a.importo || a.nome.localeCompare(b.nome, "it"));
}
const eGruppo = (c: Cella) => c.chiave.startsWith("grp|");
/** I nomi del foglio sono link (Simone 04/11/2026, "come su vista classica"). */
const LINK_CLS =
  "text-left underline decoration-dotted decoration-border underline-offset-4 hover:text-primary hover:decoration-primary";

const daRiga = (r: RigaCassa): Cella => ({
  chiave: r.chiave,
  tipo: r.tipo,
  nome: r.controparte,
  dettaglio: r.descrizione,
  importo: r.importo,
  importoOrigine: r.importoOrigine,
  secchio: r.secchio,
  entrata: r.entrata,
  attesa: r.standbyIncasso || r.parcheggiata,
  parcheggiata: r.parcheggiata,
  daSollecitare: r.daSollecitare,
  nota: r.nota,
  data: r.scadenza,
  certa: r.certa,
  documenti: r.documenti,
  editabile: r.editabile,
  priorita: r.priorita,
  dataScelta: r.dataScelta,
  modalita: r.modalita,
  stima: r.stima,
  tipologia: r.tipologia,
  cliente: r.cliente,
  promemoria: r.promemoria,
  voci: r.voci,
  quote: r.quote,
  membriClienti: r.membriClienti,
});

/** La colonna Entrate in tre blocchi col loro totale: "Già entrate" (in
 *  banca, dentro il saldo), "Entranti" (attese = standby incasso, al netto
 *  degli acconti) e "Standby" (attese parcheggiate, non contano) — Simone
 *  19/10 e 02/11/2026. */
type SlotE =
  | { kind: "head"; label: string; value: number }
  | { kind: "riga"; c: Cella; figlio?: boolean }
  /** Cliente con più fatture attese: una riga sola col totale e il "+" che
   *  apre le singole fatture (Simone 10/11/2026, "un + per cliente"). */
  | { kind: "padre"; c: Cella; figli: Cella[]; aperto: boolean };
function slotsEntrate(
  entrate: readonly Cella[],
  totali: TotaliCassa,
  t: TFn,
  espansi: ReadonlySet<string> = new Set(),
  /** Con un filtro acceso ogni blocco mostra la somma delle voci visibili
   *  (non il totale del foglio) e i blocchi rimasti vuoti spariscono. */
  filtrate = false,
): SlotE[] {
  const entranti =
    totali.acconti > 0 && !filtrate
      ? `${t("cpr.entranti")} − ${t("cpr.acconti")} ${fmt(totali.acconti)}`
      : t("cpr.entranti");
  // Per cliente: le fatture dello stesso cliente (o gruppo) stanno sotto una
  // riga padre col totale; il "+" le apre una per scadenza. Vale per gli
  // entranti, per le "da sollecitare" e per lo standby (Simone 10/11/2026).
  const conPadri = (lista: readonly Cella[], prefisso: string): SlotE[] => {
    const perCliente = new Map<string, Cella[]>();
    for (const c of lista) {
      const l = perCliente.get(c.nome) ?? [];
      l.push(c);
      perCliente.set(c.nome, l);
    }
    const out: SlotE[] = [];
    for (const [nome, figli] of perCliente) {
      if (figli.length === 1) {
        out.push({ kind: "riga", c: figli[0] });
        continue;
      }
      const primo = figli[0];
      const importo = Math.round(figli.reduce((s, c) => s + c.importo, 0) * 100) / 100;
      const padre: Cella = {
        ...primo,
        chiave: `${prefisso}|${nome}`,
        importo,
        importoOrigine: importo,
        data: figli.reduce((m, c) => (c.data < m ? c.data : m), primo.data),
        documenti: figli.reduce((s, c) => s + c.documenti, 0),
        dettaglio: figli.map((c) => `${ddmm(c.data)} ${fmt(c.importo)}`).join(" · "),
        nota: "",
        editabile: false,
        dataScelta: false,
        priorita: undefined,
      };
      const aperto = espansi.has("*") || espansi.has(`${prefisso}|${nome}`);
      out.push({ kind: "padre", c: padre, figli, aperto });
      if (aperto) for (const c of figli) out.push({ kind: "riga", c, figlio: true });
    }
    return out;
  };
  const gia = entrate.filter((c) => bloccoEntrata(c) === "gia");
  const ent = entrate.filter((c) => bloccoEntrata(c) === "entranti");
  const sol = entrate.filter((c) => bloccoEntrata(c) === "sollecitare");
  const stb = entrate.filter((c) => bloccoEntrata(c) === "standby");
  const blocchi: { label: string; value: number; celle: Cella[]; slots: SlotE[] }[] = [
    {
      label: t("cpr.giaEntrate"),
      value: totali.entrateRicevute,
      celle: gia,
      slots: gia.map((c): SlotE => ({ kind: "riga", c })),
    },
    { label: entranti, value: totali.entrantiNette, celle: ent, slots: conPadri(ent, "cli") },
    {
      label: t("cpr.daSollecitare"),
      value: totali.entrateDaSollecitare,
      celle: sol,
      slots: conPadri(sol, "sol"),
    },
    {
      label: t("cpr.colStandby"),
      value: totali.entrateParcheggiate,
      celle: stb,
      slots: conPadri(stb, "stb"),
    },
  ];
  return blocchi
    .filter((b) => !filtrate || b.celle.length > 0)
    .flatMap((b): SlotE[] => [
      { kind: "head", label: b.label, value: filtrate ? sommaCelle(b.celle) : b.value },
      ...b.slots,
    ]);
}

function daFoto(r: FotografiaCassa["righe"][number]): Cella {
  return {
    chiave: r.k,
    tipo: r.t,
    nome: r.c,
    dettaglio: r.d,
    importo: r.i,
    importoOrigine: r.i,
    secchio: r.b,
    entrata: r.e,
    attesa: Boolean(r.si) || Boolean(r.sp),
    parcheggiata: Boolean(r.sp),
    daSollecitare: Boolean(r.sr),
    nota: r.n ?? "",
    data: r.s,
    certa: false,
    documenti: r.nd ?? 1,
    editabile: false,
    priorita: r.p,
    dataScelta: false,
  };
}

export function CassaProprietaView({
  attive,
  passive,
  esclusioni,
  infoFornitori,
  fiscale,
  fissi,
  movimenti,
  onApriFornitore,
  mezzi,
  gruppi,
}: {
  attive: readonly FatturaConStato[];
  passive: readonly FatturaConStato[];
  esclusioni: readonly { nome: string; mese?: string; meseFine?: string }[];
  infoFornitori: ReadonlyMap<string, FornitoreInfo>;
  fiscale: readonly ScadenzaFiscale[];
  fissi: readonly CostoFisso[];
  movimenti: readonly SpMovimento[];
  /** Clic sul nome di un fornitore: apre la scheda come nei Flussi. */
  onApriFornitore?: (nome: string) => void;
  /** Parco mezzi per la ripartizione automatica dei fornitori (facoltativo). */
  mezzi?: ParcoPerRipartizione;
  /** Gruppi di controparti (Resoconto → Gruppi): per unire le entrate per gruppo. */
  gruppi?: readonly { nome: string; membri: string }[];
}) {
  const { t } = useLang();
  const [db, setDb] = useState<CassaDb | null>(null);
  const [stipendiDb, setStipendiDb] = useState<StipendiDb | null>(null);
  const [saldoBanca, setSaldoBanca] = useState<{ saldo: number; al?: string } | null>(saldoMemo);
  const [saldoErr, setSaldoErr] = useState(false);
  const [oggi, setOggi] = useState(oggiISO());
  const [limite, setLimite] = useState("");
  const [saldoIn, setSaldoIn] = useState<string | null>(null);
  const [salvando, setSalvando] = useState(false);
  const [mostraOltre, setMostraOltre] = useState(false);
  const [fotoId, setFotoId] = useState<string | null>(null);
  // Uscite per fornitore (il foglio) oppure raggruppate per tipologia o cliente.
  const [raggruppa, setRaggruppa] = useState<Raggruppa>("fornitore");
  // Filtri del foglio: solo vista, non si salvano.
  const [filtri, setFiltri] = useState<FiltriFoglio>(FILTRI_VUOTI);
  // Schede che si aprono dai nomi del foglio (Simone 04/11/2026): cliente
  // (le sue fatture aperte) e gruppo (i fornitori dentro); i fornitori
  // aprono la scheda dei Flussi tramite onApriFornitore.
  const [clienteAperto, setClienteAperto] = useState<string | null>(null);
  const [gruppoAperto, setGruppoAperto] = useState<Cella | null>(null);
  const [fiscaleAperto, setFiscaleAperto] = useState<Cella | null>(null);

  useEffect(() => {
    spCassaGet()
      .then((d) => {
        setDb(d);
        setLimite(limiteDefault(oggiISO(), d.giornoLimite));
      })
      .catch((e) => {
        setDb(emptyCassaDb());
        setLimite(limiteDefault(oggiISO(), 5));
        toast.error(t("common.error"), { description: e instanceof Error ? e.message : String(e) });
      });
    spStipendiGet()
      .then((d) => setStipendiDb(d))
      .catch(() => setStipendiDb(null));
    // Saldo dalla cache del server (una riga SharePoint, niente banca e
    // niente rilettura dei movimenti): arriva subito.
    spEbSaldoCache()
      .then((s) => {
        if (s) {
          saldoMemo = { saldo: s.saldo, al: s.aggiornatoAl };
          setSaldoBanca(saldoMemo);
        } else setSaldoErr(true);
      })
      .catch(() => setSaldoErr(true));
  }, [t]);

  // Saldo: scritto a mano (vince), altrimenti dalla banca, altrimenti
  // l'ultimo scritto a mano e salvato.
  const saldo = useMemo(() => {
    if (saldoIn != null) {
      const n = parseNum(saldoIn);
      if (n != null) return n;
    }
    if (saldoBanca) return saldoBanca.saldo;
    return db?.saldoManuale?.importo ?? 0;
  }, [saldoIn, saldoBanca, db]);

  const stipendi = useMemo(
    () =>
      stipendiDb
        ? stipendiDaPagare(
            stipendiDb,
            (limite || oggi).slice(0, 7),
            oggi.slice(0, 7),
            db?.chiusure ?? {},
          )
        : [],
    [stipendiDb, limite, oggi, db?.chiusure],
  );

  const esito = useMemo(() => {
    if (!db || !limite) return null;
    return costruisciCassa({
      oggi,
      limite,
      saldo,
      attive,
      passive,
      esclusioni,
      infoFornitori,
      fiscale,
      fissi,
      movimenti: movimenti.map((m) => ({
        id: m.id,
        cliente: m.cliente,
        descrizione: m.descrizione,
        importo: m.importo,
        dataContabile: m.dataContabile,
      })),
      stipendi,
      mezzi,
      gruppiClienti: tokenGruppi(gruppi ?? []),
      unisciClienti: Boolean(db.unisciClienti),
      db,
    });
  }, [
    db,
    limite,
    oggi,
    saldo,
    attive,
    passive,
    esclusioni,
    infoFornitori,
    fiscale,
    fissi,
    movimenti,
    stipendi,
    mezzi,
    gruppi,
  ]);

  // --- salvataggio: ogni modifica scrive subito il db (piccolo) -------------
  const salva = async (prossimo: CassaDb) => {
    setDb(prossimo);
    setSalvando(true);
    try {
      const saved = await spCassaSalva({ data: { db: prossimo } });
      setDb(saved);
    } catch (e) {
      toast.error(t("cpr.errSalva"), { description: e instanceof Error ? e.message : String(e) });
      try {
        setDb(await spCassaGet());
      } catch {
        /* resta il db locale */
      }
    } finally {
      setSalvando(false);
    }
  };
  const setStato = (chiave: string, patch: Partial<StatoRigaCassa>) => {
    if (!db) return;
    const prev = db.stati[chiave] ?? {};
    const next: StatoRigaCassa = { ...prev, ...patch };
    for (const k of Object.keys(next) as (keyof StatoRigaCassa)[])
      if (next[k] === undefined || next[k] === false || next[k] === "") delete next[k];
    const stati = { ...db.stati };
    if (Object.keys(next).length) stati[chiave] = next;
    else delete stati[chiave];
    void salva({ ...db, stati });
  };
  const aggiungiManuale = (riga: Omit<RigaManualeCassa, "id">) => {
    if (!db) return;
    void salva({
      ...db,
      manuali: [...db.manuali, { ...riga, id: `M${Date.now().toString(36)}` }],
    });
  };
  const aggiornaManuale = (id: string, patch: Partial<RigaManualeCassa>) => {
    if (!db) return;
    void salva({
      ...db,
      manuali: db.manuali.map((m) => (m.id === id ? { ...m, ...patch } : m)),
    });
  };
  const eliminaManuale = (id: string) => {
    if (!db) return;
    const stati = { ...db.stati };
    delete stati[`man|${id}`];
    void salva({ ...db, manuali: db.manuali.filter((m) => m.id !== id), stati });
  };
  const salvaSaldoManuale = () => {
    if (!db || saldoIn == null) return;
    const n = parseNum(saldoIn);
    if (n == null) {
      setSaldoIn(null);
      return;
    }
    void salva({ ...db, saldoManuale: { importo: Math.round(n * 100) / 100, data: oggi } });
  };
  const salvaFotografia = () => {
    if (!db || !esito) return;
    const nota = window.prompt(t("cpr.fotoNotaPrompt"), "") ?? "";
    const foto = creaFotografia(esito, { data: oggi, limite, da: "", nota: nota || undefined });
    void salva({
      ...db,
      fotografie: [foto, ...db.fotografie].slice(0, 120),
      saldoManuale: { importo: saldo, data: oggi },
    });
    toast.success(t("cpr.fotoSalvata"));
  };
  const eliminaFotografia = (id: string) => {
    if (!db) return;
    if (!window.confirm(t("cpr.fotoEliminaConferma"))) return;
    void salva({ ...db, fotografie: db.fotografie.filter((f) => f.id !== id) });
    if (fotoId === id) setFotoId(null);
  };
  const cambiaLimite = (iso: string) => {
    if (!iso) return;
    setLimite(iso);
    const g = Number(iso.slice(8, 10));
    if (db && iso.slice(0, 7) === limiteDefault(oggi, 1).slice(0, 7) && g >= 1 && g <= 28)
      void salva({ ...db, giornoLimite: g });
  };

  // Entrate per società o per gruppo di controparti (Simone 10/11/2026):
  // scelta salvata nel db, così resta.
  const unisci = (on: boolean) => {
    if (!db) return;
    void salva({ ...db, unisciClienti: on || undefined });
  };
  // Spostare un'uscita in una colonna: se è la colonna automatica, si toglie
  // la forzatura; per le righe manuali si aggiorna la riga.
  const sposta = (c: Cella, col: SecchioCassa) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { secchio: col });
      return;
    }
    const auto = esito?.righe.find((r) => r.chiave === c.chiave)?.secchioAuto;
    setStato(c.chiave, { b: col === auto ? undefined : col });
  };
  const correggi = (c: Cella, importo: number) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { importo });
      return;
    }
    setStato(c.chiave, { i: Math.abs(importo - c.importoOrigine) < 0.005 ? undefined : importo });
  };
  const annota = (c: Cella, nota: string) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { nota: nota || undefined });
      return;
    }
    setStato(c.chiave, { n: nota });
  };
  const togli = (c: Cella) => {
    if (c.tipo === "manuale" || c.tipo === "amex") eliminaManuale(c.chiave.slice(4));
    else setStato(c.chiave, { x: true });
  };
  // Priorità della nota (agenda) e data attesa (calendario entrate).
  const setPriorita = (c: Cella, p: PrioritaCassa | undefined) => {
    if (c.tipo === "manuale" || c.tipo === "amex") aggiornaManuale(c.chiave.slice(4), { p });
    else setStato(c.chiave, { p });
  };
  const setData = (c: Cella, d: string) => {
    if (c.tipo === "manuale" || c.tipo === "amex")
      aggiornaManuale(c.chiave.slice(4), { data: d || c.data });
    else setStato(c.chiave, { d: d || undefined });
  };
  // Entrata attesa in standby / di nuovo tra le entranti (flag `s`, vale
  // anche per le righe manuali: lo stato è per chiave).
  const parcheggia = (c: Cella, on: boolean) => setStato(c.chiave, { s: on || undefined });
  // Da sollecitare (Simone 10/11/2026): resta tra le entranti, blocco a parte.
  const sollecita = (c: Cella, on: boolean) => setStato(c.chiave, { r: on || undefined });
  // Acconti già incassati sugli entranti del mese (colonna Reale).
  const setAcconti = (importo: number) => {
    if (!db) return;
    const acconti = { ...(db.acconti ?? {}) };
    const mese = oggi.slice(0, 7);
    if (importo > 0.005) acconti[mese] = Math.round(importo * 100) / 100;
    else delete acconti[mese];
    void salva({ ...db, acconti });
  };
  // Rinomina una riga scritta a mano (clic sul nome): "stipendi iMile" →
  // "Acconto per DR Logistics (Aldo)", Simone 02/11/2026.
  const rinomina = (c: Cella, nome: string) => {
    if ((c.tipo === "manuale" || c.tipo === "amex") && nome.trim())
      aggiornaManuale(c.chiave.slice(4), { descrizione: nome.trim() });
  };
  const aggiungiVoce = (testo: string, p: PrioritaCassa, scad?: string) => {
    if (!db) return;
    const v: VoceAgenda = {
      id: `A${Date.now().toString(36)}`,
      testo,
      p,
      scad: scad || undefined,
      creata: new Date().toISOString(),
    };
    void salva({ ...db, agenda: [...(db.agenda ?? []), v] });
  };
  const toggleVoce = (id: string) => {
    if (!db) return;
    void salva({
      ...db,
      agenda: (db.agenda ?? []).map((v) => (v.id === id ? { ...v, fatta: !v.fatta } : v)),
    });
  };
  const eliminaVoce = (id: string) => {
    if (!db) return;
    void salva({ ...db, agenda: (db.agenda ?? []).filter((v) => v.id !== id) });
  };

  if (!db || !esito) {
    return (
      <div className="flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );
  }

  const foto = fotoId ? (db.fotografie.find((f) => f.id === fotoId) ?? null) : null;
  // Tutte le voci del foglio; i filtri (solo vista) restringono quelle
  // mostrate, le uscite prima del raggruppamento per tipologia/cliente.
  const entrateTutte = foto
    ? foto.righe.filter((r) => r.e).map(daFoto)
    : esito.righe.filter((r) => r.entrata).map(daRiga);
  const usciteTutte = foto
    ? foto.righe.filter((r) => !r.e).map(daFoto)
    : esito.righe.filter((r) => !r.entrata).map(daRiga);
  const entrate = filtraEntrate(entrateTutte, filtri);
  const uscite = raggruppaUscite(filtraUscite(usciteTutte, filtri), raggruppa, t);
  const totali = foto ? foto.totali : esito.totali;
  const limiteVista = foto ? foto.limite : limite;
  const saldoVista = foto ? foto.saldo : saldo;
  const dataVista = foto ? foto.data : oggi;
  const calendario = foto ? [] : calendarioEntrate(esito.righe);
  const agenda = foto ? [] : agendaCassa(esito.righe, db.agenda ?? []);

  const righeCsv = () => {
    const out: (string | number)[][] = [];
    // L'export è sempre il foglio INTERO: i filtri sono solo vista.
    const sl = slotsEntrate(entrateTutte, totali, t);
    const usc = raggruppaUscite(usciteTutte, raggruppa, t);
    const n = Math.max(sl.length, usc.length);
    out.push([totali.entrate, "", saldoVista, totali.reale, "", "", "", "", "", ""]);
    for (let i = 0; i < n; i++) {
      const s = sl[i];
      const e = s?.kind === "riga" || s?.kind === "padre" ? s.c : undefined;
      const u = usc[i];
      out.push([
        s?.kind === "head" ? s.value : e ? e.importo : "",
        s?.kind === "head"
          ? s.label
          : e
            ? `${e.nome}${e.attesa ? ` (${t("cpr.attesa")} ${ddmmyy(e.data)})` : ""}`
            : "",
        "",
        "",
        u && u.secchio === "ora" ? u.importo : "",
        u && u.secchio === "prossimi" ? u.importo : "",
        u && u.secchio === "standby" ? u.importo : "",
        u && u.secchio === "standby2" ? u.importo : "",
        u ? u.nome : "",
        u
          ? [
              u.dettaglio === "CDC" ? "CDC" : "",
              u.priorita ? `[${u.priorita.toUpperCase()}]` : "",
              u.nota,
            ]
              .filter(Boolean)
              .join(" ")
          : "",
      ]);
    }
    out.push([
      totali.entrateStandby,
      totali.entrateParcheggiate,
      "",
      totali.acconti,
      totali.usciteOra,
      totali.usciteProssimi,
      totali.usciteStandby,
      totali.usciteStandby2,
      totali.totaleDaPagare,
      totali.cassaConStandby,
    ]);
    out.push([
      t("cpr.lblStandbyIncasso"),
      t("cpr.lblEntrateStandby"),
      "",
      t("cpr.acconti"),
      t("cpr.lblDaPagare"),
      `${t("cpr.lblDaPagare")} ${etichettaProssimiBreve(limiteVista)}`,
      t("cpr.lblStandbyDaPagare"),
      t("cpr.lblStandby2"),
      t("cpr.lblTotaleDaPagare"),
      t("cpr.lblCassa"),
    ]);
    out.push(["", "", "", "", totali.saldoDopoOra, totali.saldoDopoProssimi, "", "", "", ""]);
    out.push(["", "", "", "", t("cpr.lblSaldoDopo"), "", "", "", "", ""]);
    return out;
  };

  const saldoTesto =
    saldoIn ??
    (saldoBanca ? fmt(saldoBanca.saldo) : db.saldoManuale ? fmt(db.saldoManuale.importo) : "");
  const saldoNota = saldoBanca
    ? `${t("cpr.saldoDaBanca")} ${saldoBanca.al ? fmtData(saldoBanca.al) : ""}`
    : saldoErr && db.saldoManuale
      ? `${t("cpr.saldoScritto")} ${fmtData(db.saldoManuale.data)}`
      : saldoErr
        ? t("cpr.saldoNonDisponibile")
        : "";

  return (
    <div className="space-y-4">
      {/* Banda in alto: saldo grande, i tre numeri del foglio, data e limite */}
      <div className="rounded-2xl border border-border bg-card px-5 py-4 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-x-8 gap-y-4">
          <div>
            <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
              {t("cpr.saldoBanca")}
              {foto && <span className="ml-1 normal-case">· {fmtData(foto.data)}</span>}
            </div>
            {foto ? (
              <div className="mt-0.5 text-3xl font-bold tabular-nums text-foreground">
                {fmt(foto.saldo)} €
              </div>
            ) : (
              <div className="mt-0.5 flex items-baseline gap-1">
                <input
                  value={saldoTesto}
                  onChange={(e) => setSaldoIn(e.target.value)}
                  onBlur={salvaSaldoManuale}
                  placeholder="0,00"
                  title={t("cpr.saldoBanca")}
                  className="w-52 border-b-2 border-dotted border-border bg-transparent text-right text-3xl font-bold tabular-nums text-foreground focus:border-primary focus:outline-none"
                />
                <span className="text-xl font-semibold text-muted-foreground">€</span>
              </div>
            )}
            <div className="mt-1 text-[11px] text-muted-foreground">{saldoNota}</div>
          </div>
          <Stat label={t("cpr.lblDaPagare")} value={totali.usciteOra} tono="rosso" />
          <Stat label={t("cpr.lblSaldoDopo")} value={totali.saldoDopoOra} tono="blu" />
          <Stat label={t("cpr.lblCassa")} value={totali.cassaConStandby} tono="segno" />
          <div className="ml-auto flex flex-wrap items-center gap-x-4 gap-y-2 text-[13px]">
            <label className="flex items-center gap-2 text-muted-foreground">
              {t("cpr.situazioneAl")}
              <input
                type="date"
                value={oggi}
                onChange={(e) => {
                  const v = e.target.value || oggiISO();
                  setOggi(v);
                  setLimite(limiteDefault(v, db.giornoLimite));
                }}
                className={inputCls}
              />
            </label>
            <label className="flex items-center gap-2 text-muted-foreground">
              {t("cpr.prossimiFinoAl")}
              <input
                type="date"
                value={limite}
                onChange={(e) => cambiaLimite(e.target.value)}
                className={inputCls}
              />
            </label>
            {/* Uscite per fornitore / tipologia / cliente (Simone 02/11/2026) */}
            <span className="flex items-center gap-2 text-muted-foreground">
              {t("cpr.raggruppa")}
              <span
                className="flex overflow-hidden rounded-lg border border-border text-[12px]"
                title={t("cpr.raggruppateTip")}
              >
                {(["fornitore", "tipologia", "cliente"] as const).map((v) => (
                  <button
                    key={v}
                    type="button"
                    onClick={() => setRaggruppa(v)}
                    className={`px-2.5 py-1 ${raggruppa === v ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
                  >
                    {t(
                      v === "fornitore"
                        ? "cpr.perFornitore"
                        : v === "tipologia"
                          ? "cpr.perTipologia"
                          : "cpr.perCliente",
                    )}
                  </button>
                ))}
              </span>
            </span>
            {/* Entrate per società / per gruppo di controparti (Simone 10/11/2026) */}
            {gruppi && gruppi.length > 0 && (
              <span className="flex items-center gap-2 text-muted-foreground">
                {t("cpr.entratePer")}
                <span
                  className="flex overflow-hidden rounded-lg border border-border text-[12px]"
                  title={t("cpr.perGruppoTip")}
                >
                  {([false, true] as const).map((v) => (
                    <button
                      key={String(v)}
                      type="button"
                      onClick={() => unisci(v)}
                      className={`px-2.5 py-1 ${Boolean(db?.unisciClienti) === v ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
                    >
                      {t(v ? "cpr.perGruppo" : "cpr.perSocieta")}
                    </button>
                  ))}
                </span>
              </span>
            )}
            {salvando && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
            <button
              type="button"
              onClick={salvaFotografia}
              className="inline-flex items-center gap-1 rounded-lg border border-primary bg-primary px-3 py-1.5 text-primary-foreground shadow-sm hover:opacity-90"
            >
              <Camera className="h-4 w-4" /> {t("cpr.salvaFoto")}
            </button>
            <EsportaStampa
              nomeFile={`cassa-proprieta-${dataVista}`}
              testata={[
                t("cpr.colEntrate"),
                t("cpr.colNote"),
                t("cpr.colSaldo"),
                t("cpr.colReale"),
                t("cpr.colUscite"),
                etichettaProssimi(limiteVista),
                t("cpr.colStandby"),
                t("cpr.colStandby2"),
                t("cpr.colNote"),
                t("cpr.colNote2"),
              ]}
              righe={righeCsv}
            />
          </div>
        </div>
      </div>

      {/* Il foglio */}
      <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]">
        <Foglio
          t={t}
          entrate={entrate}
          uscite={uscite}
          entrateTutte={entrateTutte}
          usciteTutte={usciteTutte}
          filtri={filtri}
          onFiltri={setFiltri}
          totali={totali}
          saldo={saldoVista}
          limite={limiteVista}
          readOnly={Boolean(foto)}
          onSposta={sposta}
          onImporto={correggi}
          onNota={annota}
          onTogli={togli}
          onAggiungi={aggiungiManuale}
          onPriorita={setPriorita}
          onData={setData}
          onParcheggia={parcheggia}
          onSollecita={sollecita}
          onAcconti={setAcconti}
          onRinomina={rinomina}
          onApriFornitore={onApriFornitore}
          onApriCliente={setClienteAperto}
          onApriGruppo={setGruppoAperto}
          onApriFiscale={setFiscaleAperto}
          oggi={oggi}
        />
        {clienteAperto && (
          <SchedaClienteCassa
            nome={clienteAperto}
            nomi={esito?.righe.find((r) => r.controparte === clienteAperto)?.membriClienti}
            attive={attive}
            oggi={oggi}
            onChiudi={() => setClienteAperto(null)}
          />
        )}
        {gruppoAperto && (
          <SchedaGruppoCassa
            gruppo={gruppoAperto}
            onApriFornitore={onApriFornitore}
            onChiudi={() => setGruppoAperto(null)}
          />
        )}
        {fiscaleAperto && (
          <SchedaFiscaleCassa
            nome={fiscaleAperto.nome}
            voci={fiscaleAperto.voci ?? []}
            importo={fiscaleAperto.importo}
            onChiudi={() => setFiscaleAperto(null)}
          />
        )}
        {/* RID/RiBa già scaduti: la banca li ha già presi, fuori dal foglio */}
        {!foto && esito.addebitate.length > 0 && (
          <div
            className="border-t border-border px-3 py-1.5 text-[11px] text-muted-foreground"
            title={esito.addebitate
              .map((r) => `${r.controparte} ${fmt(r.importo)} (${ddmm(r.scadenza)})`)
              .join(" · ")}
          >
            <Lock className="mr-1 inline h-3 w-3" />
            {t("cpr.addebitati")}{" "}
            <span className="font-semibold text-foreground">
              {esito.addebitate.length} · {fmt(totali.usciteAddebitate)} €
            </span>{" "}
            (
            {esito.addebitate.slice(0, 6).map((r, i) => (
              <span key={r.chiave}>
                {i > 0 && ", "}
                {r.tipo === "passiva" && onApriFornitore ? (
                  <button
                    type="button"
                    onClick={() => onApriFornitore(r.controparte)}
                    title={t("cpr.apriFornitore")}
                    className={LINK_CLS}
                  >
                    {r.controparte}
                  </button>
                ) : (
                  r.controparte
                )}
              </span>
            ))}
            {esito.addebitate.length > 6 ? ", …" : ""})
          </div>
        )}
        {/* Legenda: cosa è scritto a mano, cosa è stimato, cosa esce da solo */}
        <div className="flex flex-wrap items-center gap-x-4 gap-y-1 border-t border-border px-3 py-1.5 text-[11px] text-muted-foreground">
          <span className="inline-flex items-center gap-1">
            <span className={`rounded px-1 text-[9px] font-bold uppercase ${PILL_MAN}`}>
              {t("cpr.manuale")}
            </span>
            {t("cpr.legendaManuale")} · {t("cpr.rinomina")}
          </span>
          <span className="inline-flex items-center gap-1">
            <span className={`rounded px-1 text-[9px] font-bold uppercase ${PILL_COMP}`}>
              {t("cpr.compensazione")}
            </span>
            {t("cpr.legendaCompensazione")}
          </span>
          <span>{t("cpr.regolaFineMese")}</span>
          <span className="inline-flex items-center gap-1">
            <span className={`rounded px-1 text-[9px] font-bold uppercase ${PILL_STIMA}`}>
              {t("cpr.stima")}
            </span>
            {t("cpr.legendaStima")}
          </span>
          <span className="inline-flex items-center gap-1">
            <span
              className={`inline-flex items-center gap-0.5 rounded px-1 text-[9px] font-bold uppercase ${PILL_AUTO}`}
            >
              <Lock className="h-2.5 w-2.5" /> RID
            </span>
            {t("cpr.legendaAuto")}
          </span>
        </div>
        {/* Linguette come i fogli di Excel */}
        <div className="flex flex-wrap items-center gap-1 border-t border-border bg-muted/50 px-3 pt-1.5 pb-2 text-[12px]">
          <button
            type="button"
            onClick={() => setFotoId(null)}
            className={`rounded-b-lg border border-t-0 px-3 py-1 ${!foto ? "border-primary/40 bg-card font-semibold text-primary shadow-sm" : "border-transparent text-muted-foreground hover:bg-card/70"}`}
          >
            {t("cpr.oggi")} · {ddmm(oggi)}
          </button>
          {db.fotografie.map((f) => (
            <span
              key={f.id}
              className={`group inline-flex items-center gap-1 rounded-b-lg border border-t-0 px-3 py-1 ${foto?.id === f.id ? "border-primary/40 bg-card font-semibold text-primary shadow-sm" : "border-transparent text-muted-foreground hover:bg-card/70"}`}
            >
              <button type="button" onClick={() => setFotoId(f.id)} title={f.nota ?? ""}>
                {ddmm(f.data)}
                {f.nota ? ` · ${f.nota.slice(0, 18)}` : ""}
              </button>
              <button
                type="button"
                onClick={() => eliminaFotografia(f.id)}
                className="opacity-0 group-hover:opacity-100 hover:text-status-absent"
                title={t("common.delete")}
              >
                <X className="h-3 w-3" />
              </button>
            </span>
          ))}
        </div>
      </div>

      {/* Calendario entrate e agenda (Simone 20/10/2026) */}
      {!foto && (
        <div className="grid gap-4 lg:grid-cols-2">
          <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
            <div className="mb-2 flex items-center gap-2 text-[13px] font-semibold text-foreground">
              <CalendarDays className="h-4 w-4 text-emerald-600" /> {t("cpr.calendario")}
              <span className="ml-auto text-[12px] font-normal tabular-nums text-muted-foreground">
                {fmt(totali.entrateStandby)} €
              </span>
            </div>
            {calendario.length === 0 ? (
              <p className="text-[12px] text-muted-foreground">{t("cpr.calendarioVuoto")}</p>
            ) : (
              <table className="w-full border-collapse text-[13px]">
                <thead>
                  <tr className="text-[10.5px] uppercase tracking-wider text-muted-foreground">
                    <th className="py-1 pr-2 text-left font-semibold">{t("cpr.data")}</th>
                    <th className="py-1 pr-2 text-left font-semibold">{t("cpr.colNote")}</th>
                    <th className="py-1 pr-2 text-right font-semibold">{t("cpr.importo")}</th>
                    <th className="py-1 text-right font-semibold">{t("cpr.progressivo")}</th>
                  </tr>
                </thead>
                <tbody>
                  {calendario.map((v) => {
                    const c = entrate.find((x) => x.chiave === v.chiave);
                    return (
                      <tr
                        key={v.chiave}
                        className={`border-t border-border/50 ${v.data < oggi ? "text-rose-600 dark:text-rose-400" : "text-foreground"}`}
                      >
                        <td className="py-1 pr-2 whitespace-nowrap">
                          {c ? (
                            <DataAttesa c={c} readOnly={false} onData={setData} t={t} />
                          ) : (
                            ddmm(v.data)
                          )}
                        </td>
                        <td className="max-w-[220px] truncate py-1 pr-2">{v.nome}</td>
                        <td className="py-1 pr-2 text-right tabular-nums">{fmt(v.importo)}</td>
                        <td className="py-1 text-right tabular-nums text-muted-foreground">
                          {fmt(v.progressivo)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
          <AgendaBox
            t={t}
            voci={agenda}
            onAggiungi={aggiungiVoce}
            onToggle={toggleVoce}
            onElimina={eliminaVoce}
            onTogliRiga={(chiave) => {
              const c = [...entrate, ...uscite].find((x) => x.chiave === chiave);
              if (c) setPriorita(c, undefined);
            }}
          />
        </div>
      )}

      {!foto && esito.oltre.length > 0 && (
        <div className="text-[12px] text-muted-foreground">
          <button
            type="button"
            onClick={() => setMostraOltre((v) => !v)}
            className="underline-offset-2 hover:underline"
          >
            {mostraOltre ? "▾" : "▸"} {t("cpr.oltreIl")} {ddmm(limite)}: {esito.oltre.length}{" "}
            {t("cpr.voci")} · {fmt(esito.totali.usciteOltre)}
          </button>
          {mostraOltre && (
            <table className="mt-1 border-collapse text-[12px]">
              <tbody>
                {esito.oltre.map((r) => (
                  <tr key={r.chiave} className="border-b border-border/40">
                    <td className="py-0.5 pr-3 whitespace-nowrap">{ddmm(r.scadenza)}</td>
                    <td className="py-0.5 pr-3 text-foreground">
                      {r.tipo === "passiva" && onApriFornitore ? (
                        <button
                          type="button"
                          onClick={() => onApriFornitore(r.controparte)}
                          title={t("cpr.apriFornitore")}
                          className={LINK_CLS}
                        >
                          {r.controparte}
                        </button>
                      ) : (
                        r.controparte
                      )}
                      {r.documenti > 1 && r.tipo !== "stipendi" && (
                        <span className="ml-1 text-muted-foreground">×{r.documenti}</span>
                      )}
                    </td>
                    <td className="py-0.5 pr-3 text-right tabular-nums text-foreground">
                      {fmt(r.importo)}
                    </td>
                    <td className="py-0.5 pr-2">
                      <button
                        type="button"
                        className="rounded border border-border px-1.5 hover:bg-muted"
                        onClick={() => setStato(r.chiave, { b: "prossimi" })}
                      >
                        {t("cpr.portaNeiProssimi")}
                      </button>
                    </td>
                    <td className="py-0.5">
                      <button
                        type="button"
                        className="rounded border border-border px-1.5 hover:bg-muted"
                        onClick={() => setStato(r.chiave, { b: "standby" })}
                      >
                        {t("cpr.parcheggia")}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}

function Stat({
  label,
  value,
  tono,
}: {
  label: string;
  value: number;
  tono: "rosso" | "blu" | "segno";
}) {
  const colore =
    tono === "rosso"
      ? "text-rose-600 dark:text-rose-400"
      : tono === "blu"
        ? "text-primary"
        : value < 0
          ? "text-rose-600 dark:text-rose-400"
          : "text-emerald-700 dark:text-emerald-400";
  return (
    <div>
      <div className="max-w-[220px] text-[11px] uppercase tracking-wider text-muted-foreground">
        {label}
      </div>
      <div className={`mt-0.5 text-xl font-semibold tabular-nums ${colore}`}>{fmt(value)} €</div>
    </div>
  );
}

// --- il foglio ---------------------------------------------------------------
// Tinte delle colonne (chiare in body, piene in testa): entrate verde,
// uscite rosso, prossimi ambra, standby ardesia, note neutre.
const TINTA = {
  entrate: {
    testa: "bg-emerald-600 text-white",
    col: "bg-emerald-50/60 dark:bg-emerald-950/25",
    num: "text-emerald-800 dark:text-emerald-300",
  },
  ora: {
    testa: "bg-rose-600 text-white",
    col: "bg-rose-50/60 dark:bg-rose-950/25",
    num: "text-rose-700 dark:text-rose-300",
  },
  prossimi: {
    testa: "bg-amber-500 text-white",
    col: "bg-amber-50/70 dark:bg-amber-950/25",
    num: "text-amber-700 dark:text-amber-300",
  },
  standby: {
    testa: "bg-slate-500 text-white",
    col: "bg-slate-100/70 dark:bg-slate-800/40",
    num: "text-slate-600 dark:text-slate-300",
  },
  // Standby prolungato ("Standby 2"): più scuro, fuori dai conti.
  standby2: {
    testa: "bg-slate-700 text-white",
    col: "bg-slate-200/70 dark:bg-slate-800/70",
    num: "text-slate-700 dark:text-slate-200",
  },
  note: { testa: "bg-muted text-foreground", col: "", num: "" },
} as const;
/** Colonna "Reale" accanto al saldo: saldo + entranti − da pagare − prossimi. */
const TINTA_REALE = "bg-sky-50/70 dark:bg-sky-950/30";
/** Righe scritte a mano (non calcolate): tinta viola e pillola "manuale". */
const TINTA_MAN = "bg-violet-50/80 dark:bg-violet-950/30";
const PILL_MAN = "bg-violet-600 text-white";
const PILL_COMP = "bg-rose-600 text-white";
const PILL_STIMA = "bg-amber-500 text-white";
const PILL_AUTO =
  "border border-rose-300 bg-rose-50 text-rose-700 dark:border-rose-700 dark:bg-rose-950/40 dark:text-rose-300";
const eManuale = (c: Cella) => c.tipo === "manuale" || c.tipo === "amex";
const cel = "border-b border-r border-border/60 px-3 py-1.5 align-middle last:border-r-0";
const celN = `${cel} text-right tabular-nums whitespace-nowrap text-[13.5px]`;
const celH = "px-3 py-2 text-left text-[12px] font-semibold uppercase tracking-wide";
const celT = `${celN} py-2.5 text-[15px] font-bold`;
const celL = `${cel} py-1 text-[10.5px] uppercase tracking-wider text-muted-foreground`;

function Foglio({
  t,
  entrate,
  uscite,
  entrateTutte,
  usciteTutte,
  filtri,
  onFiltri,
  totali,
  saldo,
  limite,
  readOnly,
  oggi,
  onSposta,
  onImporto,
  onNota,
  onTogli,
  onAggiungi,
  onPriorita,
  onData,
  onParcheggia,
  onSollecita,
  onAcconti,
  onRinomina,
  onApriFornitore,
  onApriCliente,
  onApriGruppo,
  onApriFiscale,
}: {
  t: TFn;
  /** Voci mostrate (già filtrate; le uscite anche raggruppate). */
  entrate: readonly Cella[];
  uscite: readonly Cella[];
  /** Tutte le voci, senza filtri né raggruppamento: per la barra dei filtri. */
  entrateTutte: readonly Cella[];
  usciteTutte: readonly Cella[];
  filtri: FiltriFoglio;
  onFiltri: (f: FiltriFoglio) => void;
  totali: TotaliCassa;
  saldo: number;
  limite: string;
  readOnly: boolean;
  oggi: string;
  onSposta: (c: Cella, col: SecchioCassa) => void;
  onImporto: (c: Cella, importo: number) => void;
  onNota: (c: Cella, nota: string) => void;
  onTogli: (c: Cella) => void;
  onAggiungi: (r: Omit<RigaManualeCassa, "id">) => void;
  onPriorita: (c: Cella, p: PrioritaCassa | undefined) => void;
  onData: (c: Cella, d: string) => void;
  /** Entrata attesa in standby (true) o di nuovo tra le entranti (false). */
  onParcheggia: (c: Cella, on: boolean) => void;
  /** Entrata attesa da sollecitare (true) o di nuovo tra le entranti (false). */
  onSollecita: (c: Cella, on: boolean) => void;
  /** Acconti già incassati sugli entranti del mese (colonna Reale). */
  onAcconti: (importo: number) => void;
  /** Nuovo nome per una riga scritta a mano. */
  onRinomina: (c: Cella, nome: string) => void;
  onApriFornitore?: (nome: string) => void;
  /** Clic sul nome di un cliente: le sue fatture aperte. */
  onApriCliente: (nome: string) => void;
  /** Clic sul nome di un gruppo (uscite per tipologia/cliente): i fornitori dentro. */
  onApriGruppo: (c: Cella) => void;
  /** Clic su Fiscale/IVA/Banca: le scadenze dentro la riga. */
  onApriFiscale: (c: Cella) => void;
}) {
  const [nuova, setNuova] = useState<null | {
    entrata: boolean;
    descrizione: string;
    importo: string;
    data: string;
    secchio: SecchioCassa;
    amex: boolean;
    ricorrente: boolean;
  }>(null);
  // Clienti aperti col "+" (singole fatture visibili); non si salva.
  const [espansi, setEspansi] = useState<Set<string>>(new Set());
  const toggleCliente = (nome: string) =>
    setEspansi((prev) => {
      const next = new Set(prev);
      if (next.has(nome)) next.delete(nome);
      else next.add(nome);
      return next;
    });
  const slotsE = slotsEntrate(entrate, totali, t, espansi, filtroEntrateAttivo(filtri));
  const n = Math.max(slotsE.length, uscite.length);
  const righe = Array.from({ length: n }, (_, i) => i);
  const breve = etichettaProssimiBreve(limite);
  // Barra di scorrimento orizzontale anche SOPRA il foglio (Simone
  // 11/11/2026, "barra sopra per spostarsi a destra"): segue quella in basso
  // e viceversa. Compare solo se il foglio è più largo della cornice.
  const corniceRef = useRef<HTMLDivElement>(null);
  const barraRef = useRef<HTMLDivElement>(null);
  const [largh, setLargh] = useState({ tot: 0, vis: 0, sb: 0 });
  useEffect(() => {
    const el = corniceRef.current;
    const misura = () => {
      if (!el) return;
      const nuova = {
        tot: el.scrollWidth,
        vis: el.clientWidth,
        sb: el.offsetWidth - el.clientWidth,
      };
      setLargh((p) =>
        p.tot === nuova.tot && p.vis === nuova.vis && p.sb === nuova.sb ? p : nuova,
      );
    };
    misura();
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(misura) : null;
    if (el && ro) {
      ro.observe(el);
      if (el.firstElementChild) ro.observe(el.firstElementChild);
    }
    return () => ro?.disconnect();
  }, []);
  const allinea = (da: HTMLDivElement | null, a: HTMLDivElement | null) => {
    if (da && a && Math.abs(a.scrollLeft - da.scrollLeft) > 1) a.scrollLeft = da.scrollLeft;
  };

  const conferma = () => {
    if (!nuova) return;
    const imp = parseNum(nuova.importo);
    if (!nuova.descrizione.trim() || imp == null || imp === 0) {
      toast.error(t("cpr.manualeIncompleta"));
      return;
    }
    onAggiungi({
      tipo: nuova.amex ? "amex" : "manuale",
      descrizione: nuova.descrizione.trim(),
      importo: Math.abs(imp),
      data: nuova.data || oggi,
      secchio: nuova.entrata ? "ora" : nuova.secchio,
      entrata: nuova.entrata || undefined,
      ricorrente: nuova.ricorrente || undefined,
    });
    setNuova(null);
  };
  const apri = (entrata: boolean) =>
    setNuova({
      entrata,
      descrizione: "",
      importo: "",
      data: oggi,
      secchio: "ora",
      amex: false,
      ricorrente: false,
    });

  const cellaUscita = (u: Cella | undefined, col: SecchioCassa) => {
    const tinta = TINTA[col];
    if (!u) return <td className={`${celN} ${tinta.col}`} />;
    const man = eManuale(u) ? TINTA_MAN : "";
    if (u.secchio === col)
      return (
        <td
          className={`${celN} ${tinta.col} ${man} ${tinta.num} font-medium ${u.stima ? "italic" : ""}`}
        >
          <Importo c={u} readOnly={readOnly} onImporto={onImporto} t={t} />
        </td>
      );
    if (readOnly) return <td className={`${celN} ${tinta.col} ${man}`} />;
    // Vista raggruppata (tipologia/cliente): sola lettura.
    if (eGruppo(u))
      return <td className={`${celN} ${tinta.col}`} title={t("cpr.raggruppateTip")} />;
    // RID/RiBa: l'addebito arriva da solo alla scadenza, non si parcheggia.
    if (u.tipo === "passiva" && u.certa)
      return <td className={`${celN} ${tinta.col}`} title={t("cpr.nonParcheggiabile")} />;
    return (
      <td className={`${celN} ${tinta.col} ${man} p-0`}>
        <button
          type="button"
          onClick={() => onSposta(u, col)}
          title={t("cpr.spostaQui")}
          className="h-full min-h-[34px] w-full px-3 text-right text-transparent hover:bg-primary/10 hover:text-muted-foreground"
        >
          →
        </button>
      </td>
    );
  };

  return (
    // Il foglio scorre dentro la sua cornice (barra di scorrimento sua, come
    // Excel), con le testate fisse in alto — Simone 02/11/2026.
    <div>
      <BarraFiltri
        t={t}
        filtri={filtri}
        onFiltri={onFiltri}
        entrateTutte={entrateTutte}
        usciteTutte={usciteTutte}
        limite={limite}
      />
      {largh.tot > largh.vis + 1 && (
        <div
          ref={barraRef}
          onScroll={() => allinea(barraRef.current, corniceRef.current)}
          title={t("cpr.barraSopra")}
          className="overflow-x-scroll overflow-y-hidden border-b border-border/60 bg-muted/40 [&::-webkit-scrollbar]:h-3 [&::-webkit-scrollbar-thumb]:rounded-full [&::-webkit-scrollbar-thumb]:bg-primary/60 [&::-webkit-scrollbar-thumb:hover]:bg-primary [&::-webkit-scrollbar-track]:bg-transparent"
        >
          <div style={{ width: largh.tot + largh.sb, height: 1 }} />
        </div>
      )}
      <div
        ref={corniceRef}
        onScroll={() => allinea(corniceRef.current, barraRef.current)}
        className="max-h-[78vh] overflow-auto"
      >
        <table className="w-full min-w-[1400px] border-collapse text-[13px]">
          <colgroup>
            <col className="w-[120px]" />
            <col className="w-[215px]" />
            <col className="w-[120px]" />
            <col className="w-[135px]" />
            <col className="w-[120px]" />
            <col className="w-[120px]" />
            <col className="w-[120px]" />
            <col className="w-[120px]" />
            <col className="w-[280px]" />
            <col />
          </colgroup>
          <thead className="sticky top-0 z-10 bg-card">
            <tr>
              <th
                colSpan={4}
                className={`${celH} ${TINTA.entrate.testa} border-r border-white/30 text-center text-[11px]`}
              >
                {t("cpr.colEntrate")}
              </th>
              <th
                colSpan={6}
                className={`${celH} bg-primary text-center text-[11px] text-primary-foreground`}
              >
                {t("cpr.colUscite")}
              </th>
            </tr>
            <tr className="border-b border-border">
              <th
                className={`${celH} ${TINTA.entrate.col} ${TINTA.entrate.num} border-r border-border/60 text-right`}
              >
                {t("cpr.colEntrate")}
              </th>
              <th
                className={`${celH} ${TINTA.entrate.col} border-r border-border/60 text-foreground`}
              >
                {t("cpr.colNote")}
              </th>
              <th
                className={`${celH} border-r border-border/60 bg-muted/60 text-right text-foreground`}
              >
                {t("cpr.colSaldo")}
              </th>
              <th
                className={`${celH} ${TINTA_REALE} border-r border-border/60 text-right text-sky-800 dark:text-sky-300`}
                title={t("cpr.lblReale")}
              >
                {t("cpr.colReale")}
              </th>
              <th
                className={`${celH} ${TINTA.ora.col} ${TINTA.ora.num} border-r border-border/60 text-right`}
              >
                {t("cpr.colUscite")}
              </th>
              <th
                className={`${celH} ${TINTA.prossimi.col} ${TINTA.prossimi.num} border-r border-border/60 text-right normal-case`}
              >
                {etichettaProssimi(limite)}
              </th>
              <th
                className={`${celH} ${TINTA.standby.col} ${TINTA.standby.num} border-r border-border/60 text-right`}
              >
                {t("cpr.colStandby")}
              </th>
              <th
                className={`${celH} ${TINTA.standby2.col} ${TINTA.standby2.num} border-r border-border/60 text-right`}
                title={t("cpr.lblStandby2")}
              >
                {t("cpr.colStandby2")}
              </th>
              <th className={`${celH} border-r border-border/60 bg-muted/60 text-foreground`}>
                {t("cpr.colNote")}
              </th>
              <th className={`${celH} bg-muted/60 text-foreground`}>{t("cpr.colNote2")}</th>
            </tr>
          </thead>
          <tbody>
            {/* riga 2: totale entrate, saldo e il "reale" (saldo + entranti −
              da pagare − prossimi), con gli acconti già incassati da togliere */}
            <tr className="bg-muted/30">
              <td className={`${celT} ${TINTA.entrate.num}`}>{fmt(totali.entrate)}</td>
              <td className={`${cel} text-[11px] uppercase tracking-wider text-muted-foreground`}>
                {t("cpr.lblTotaleEntrate")}
              </td>
              <td className={`${celT} text-foreground`}>{fmt(saldo)}</td>
              <td className={`${celN} ${TINTA_REALE} align-top`}>
                <div
                  className={`py-1 text-[15px] font-bold ${totali.reale < 0 ? "text-rose-600 dark:text-rose-400" : "text-sky-800 dark:text-sky-300"}`}
                  title={t("cpr.lblReale")}
                >
                  {fmt(totali.reale)}
                </div>
                <Acconti importo={totali.acconti} readOnly={readOnly} onAcconti={onAcconti} t={t} />
              </td>
              <td className={`${cel} ${TINTA.ora.col}`} />
              <td className={`${cel} ${TINTA.prossimi.col}`} />
              <td className={`${cel} ${TINTA.standby.col}`} />
              <td className={`${cel} ${TINTA.standby2.col}`} />
              <td className={cel} />
              <td className={cel} />
            </tr>
            {righe.map((i) => {
              const slot = slotsE[i];
              const e = slot?.kind === "riga" || slot?.kind === "padre" ? slot.c : undefined;
              const padre = slot?.kind === "padre" ? slot : undefined;
              const figlio = slot?.kind === "riga" && slot.figlio === true;
              const u = uscite[i];
              return (
                <tr key={i} className="group/r odd:bg-background/40 hover:bg-primary/5">
                  {slot?.kind === "head" ? (
                    <td
                      colSpan={2}
                      className={`${cel} bg-emerald-100/70 dark:bg-emerald-900/30 ${TINTA.entrate.num}`}
                    >
                      <span className="flex items-center justify-between gap-2">
                        <span className="text-[11px] font-semibold uppercase tracking-wider">
                          {slot.label}
                        </span>
                        <span className="text-[13.5px] font-bold tabular-nums">
                          {fmt(slot.value)}
                        </span>
                      </span>
                    </td>
                  ) : (
                    <>
                      <td
                        className={`${celN} ${e?.parcheggiata ? TINTA.standby.col : TINTA.entrate.col} ${e && eManuale(e) ? TINTA_MAN : ""} ${e?.parcheggiata ? "italic text-slate-500 line-through decoration-slate-400" : e?.attesa ? "italic text-muted-foreground" : `${TINTA.entrate.num} font-medium`}`}
                      >
                        {e && (
                          <span
                            className={
                              e.importo < 0
                                ? "text-rose-700 dark:text-rose-400"
                                : e.daSollecitare
                                  ? "text-amber-700 dark:text-amber-400"
                                  : ""
                            }
                          >
                            <Importo c={e} readOnly={readOnly} onImporto={onImporto} t={t} />
                          </span>
                        )}
                      </td>
                      <td
                        className={`${cel} ${e?.parcheggiata ? TINTA.standby.col : TINTA.entrate.col} ${e && eManuale(e) ? TINTA_MAN : ""} ${e?.attesa ? "text-muted-foreground" : "text-foreground"}`}
                      >
                        {e && (
                          <span className={`flex items-center gap-1.5 ${figlio ? "pl-5" : ""}`}>
                            {padre && (
                              <button
                                type="button"
                                onClick={() => toggleCliente(e.chiave)}
                                title={t(padre.aperto ? "cpr.chiudiCliente" : "cpr.espandiCliente")}
                                className="inline-flex h-4 w-4 shrink-0 items-center justify-center rounded border border-border text-[11px] leading-none text-muted-foreground hover:bg-muted hover:text-foreground"
                              >
                                {padre.aperto ? "−" : "+"}
                              </button>
                            )}
                            {e.tipo === "attiva" || e.tipo === "incasso" ? (
                              <button
                                type="button"
                                onClick={() => onApriCliente(e.nome)}
                                title={t("cpr.apriCliente")}
                                className={`max-w-[200px] truncate ${LINK_CLS}`}
                              >
                                {e.nome}
                              </button>
                            ) : e.tipo === "compensazione" && onApriFornitore ? (
                              <button
                                type="button"
                                onClick={() => onApriFornitore(e.nome)}
                                title={t("cpr.apriFornitore")}
                                className={`max-w-[200px] truncate ${LINK_CLS}`}
                              >
                                {e.nome}
                              </button>
                            ) : (
                              <NomeRiga
                                c={e}
                                readOnly={readOnly}
                                onRinomina={onRinomina}
                                t={t}
                                className={`max-w-[200px] truncate ${eManuale(e) ? "italic" : ""}`}
                              />
                            )}
                            {eManuale(e) && (
                              <span
                                className={`inline-flex shrink-0 items-center gap-0.5 rounded px-1 text-[9px] font-bold uppercase ${PILL_MAN}`}
                                title={t("cpr.legendaManuale")}
                              >
                                <Pencil className="h-2.5 w-2.5" /> {t("cpr.manuale")}
                              </span>
                            )}
                            {e.tipo === "compensazione" && (
                              <span
                                className={`inline-flex shrink-0 items-center rounded px-1 text-[9px] font-bold uppercase ${PILL_COMP}`}
                                title={e.dettaglio}
                              >
                                {t("cpr.compensazione")}
                              </span>
                            )}
                            {padre ? (
                              <span className="shrink-0 text-[10.5px] text-muted-foreground">
                                {t("cpr.dal")} {ddmmyy(e.data)} · {padre.figli.length}{" "}
                                {t("cpr.fattureBreve")}
                              </span>
                            ) : e.attesa ? (
                              <DataAttesa c={e} readOnly={readOnly} onData={onData} t={t} />
                            ) : (
                              <span className="shrink-0 text-[10.5px] text-muted-foreground">
                                {ddmm(e.data)}
                              </span>
                            )}
                            {!padre && e.attesa && !e.parcheggiata && !readOnly && (
                              <ToggleSollecito c={e} onSollecita={onSollecita} t={t} />
                            )}
                            {!padre && e.attesa && !readOnly && (
                              <ToggleStandby c={e} onParcheggia={onParcheggia} t={t} />
                            )}
                            {!padre && !readOnly && <Togli onClick={() => onTogli(e)} t={t} />}
                          </span>
                        )}
                      </td>
                    </>
                  )}
                  <td className={`${cel} bg-muted/20`} />
                  <td className={`${cel} ${TINTA_REALE}`} />
                  {cellaUscita(u, "ora")}
                  {cellaUscita(u, "prossimi")}
                  {cellaUscita(u, "standby")}
                  {cellaUscita(u, "standby2")}
                  <td className={`${cel} text-foreground ${u && eManuale(u) ? TINTA_MAN : ""}`}>
                    {u && (
                      <span className="flex items-center gap-1.5">
                        {u.tipo === "fiscale" && u.voci?.length ? (
                          <button
                            type="button"
                            onClick={() => onApriFiscale(u)}
                            title={t("cpr.apriFiscale")}
                            className={`truncate font-medium ${LINK_CLS}`}
                          >
                            {u.nome}
                          </button>
                        ) : u.tipo === "passiva" && onApriFornitore && !eGruppo(u) ? (
                          <button
                            type="button"
                            onClick={() => onApriFornitore(u.nome)}
                            title={t("cpr.apriFornitore")}
                            className="truncate text-left font-medium underline decoration-dotted decoration-border underline-offset-4 hover:text-primary hover:decoration-primary"
                          >
                            {u.nome}
                          </button>
                        ) : eGruppo(u) ? (
                          <button
                            type="button"
                            onClick={() => onApriGruppo(u)}
                            title={t("cpr.apriGruppo")}
                            className={`truncate font-medium ${LINK_CLS}`}
                          >
                            {u.nome}
                          </button>
                        ) : (
                          <NomeRiga
                            c={u}
                            readOnly={readOnly}
                            onRinomina={onRinomina}
                            t={t}
                            className={`truncate font-medium ${eManuale(u) ? "italic" : ""}`}
                          />
                        )}
                        {u.promemoria && (
                          <span
                            className="shrink-0 cursor-help text-amber-600"
                            title={u.promemoria}
                          >
                            <StickyNote className="h-3 w-3" />
                          </span>
                        )}
                        {u.documenti > 1 && u.tipo !== "stipendi" && (
                          <span className="shrink-0 text-[10.5px] text-muted-foreground">
                            ×{u.documenti}
                          </span>
                        )}
                        {u.tipo === "passiva" && u.certa && (
                          <span
                            className={`inline-flex shrink-0 items-center gap-0.5 rounded px-1 text-[9px] font-bold uppercase ${PILL_AUTO}`}
                            title={t("cpr.nonParcheggiabile")}
                          >
                            <Lock className="h-2.5 w-2.5" /> {u.modalita || "AUTO"}
                          </span>
                        )}
                        {eManuale(u) && (
                          <span
                            className={`inline-flex shrink-0 items-center gap-0.5 rounded px-1 text-[9px] font-bold uppercase ${PILL_MAN}`}
                            title={t("cpr.legendaManuale")}
                          >
                            <Pencil className="h-2.5 w-2.5" /> {t("cpr.manuale")}
                          </span>
                        )}
                        {u.dettaglio === "ogni mese" && (
                          <span className="shrink-0 rounded border border-violet-400 px-1 text-[9px] uppercase text-violet-700 dark:text-violet-300">
                            {t("cpr.ricorrente")}
                          </span>
                        )}
                        {u.stima && (
                          <span
                            className={`shrink-0 rounded px-1 text-[9px] font-bold uppercase ${PILL_STIMA}`}
                            title={t("cpr.legendaStima")}
                          >
                            {t("cpr.stima")}
                          </span>
                        )}
                        {(u.tipo === "stipendi" || u.tipo === "fisso" || u.tipo === "fiscale") && (
                          <span className="shrink-0 truncate text-[10.5px] text-muted-foreground">
                            {u.dettaglio.replace(/^stima · /, "")}
                          </span>
                        )}
                      </span>
                    )}
                  </td>
                  <td className={`${cel} p-0 ${u && eManuale(u) ? TINTA_MAN : ""}`}>
                    {u && !eGruppo(u) && (
                      <span className="flex items-center">
                        {u.dettaglio === "CDC" && (
                          <span className="pl-3 text-[11px] font-semibold text-muted-foreground">
                            CDC
                          </span>
                        )}
                        <Nota c={u} readOnly={readOnly} onNota={onNota} />
                        <Priorita c={u} readOnly={readOnly} onPriorita={onPriorita} t={t} />
                        {!readOnly && (
                          <span className="pr-2">
                            <Togli onClick={() => onTogli(u)} t={t} />
                          </span>
                        )}
                      </span>
                    )}
                  </td>
                </tr>
              );
            })}
            {/* riga per aggiungere */}
            {!readOnly && (
              <tr className="bg-muted/10">
                {nuova?.entrata ? (
                  <>
                    <td className={`${celN} ${TINTA.entrate.col} p-1`}>
                      <input
                        autoFocus
                        value={nuova.importo}
                        onChange={(e) => setNuova({ ...nuova, importo: e.target.value })}
                        onKeyDown={(e) => e.key === "Enter" && conferma()}
                        placeholder={t("cpr.importo")}
                        className={`${inputCls} w-full text-right`}
                      />
                    </td>
                    <td className={`${cel} ${TINTA.entrate.col} p-1`}>
                      <span className="flex items-center gap-1">
                        <input
                          value={nuova.descrizione}
                          onChange={(e) => setNuova({ ...nuova, descrizione: e.target.value })}
                          onKeyDown={(e) => e.key === "Enter" && conferma()}
                          placeholder={t("cpr.descrizione")}
                          className={`${inputCls} w-full`}
                        />
                        <input
                          type="date"
                          value={nuova.data}
                          onChange={(e) => setNuova({ ...nuova, data: e.target.value })}
                          title={t("cpr.dataAttesa")}
                          className={inputCls}
                        />
                        <label className="flex shrink-0 items-center gap-1 text-[12px] text-muted-foreground">
                          <input
                            type="checkbox"
                            checked={nuova.ricorrente}
                            onChange={(e) => setNuova({ ...nuova, ricorrente: e.target.checked })}
                          />
                          {t("cpr.ricorrente")}
                        </label>
                        <Ok onOk={conferma} onNo={() => setNuova(null)} t={t} />
                      </span>
                    </td>
                  </>
                ) : (
                  <>
                    <td className={`${cel} ${TINTA.entrate.col}`} />
                    <td className={`${cel} ${TINTA.entrate.col} py-1`}>
                      <Aggiungi onClick={() => apri(true)} t={t} />
                    </td>
                  </>
                )}
                <td className={`${cel} bg-muted/20`} />
                <td className={`${cel} ${TINTA_REALE}`} />
                {nuova && !nuova.entrata ? (
                  <>
                    <td className={`${celN} p-1`} colSpan={4}>
                      <span className="flex items-center gap-1">
                        <input
                          autoFocus
                          value={nuova.importo}
                          onChange={(e) => setNuova({ ...nuova, importo: e.target.value })}
                          onKeyDown={(e) => e.key === "Enter" && conferma()}
                          placeholder={t("cpr.importo")}
                          className={`${inputCls} w-28 text-right`}
                        />
                        <select
                          value={nuova.secchio}
                          onChange={(e) =>
                            setNuova({ ...nuova, secchio: e.target.value as SecchioCassa })
                          }
                          className={inputCls}
                        >
                          <option value="ora">{t("cpr.colUscite")}</option>
                          <option value="prossimi">{breve}</option>
                          <option value="standby">{t("cpr.colStandby")}</option>
                          <option value="standby2">{t("cpr.colStandby2")}</option>
                        </select>
                        <input
                          type="date"
                          value={nuova.data}
                          onChange={(e) => setNuova({ ...nuova, data: e.target.value })}
                          className={inputCls}
                        />
                      </span>
                    </td>
                    <td className={`${cel} p-1`}>
                      <input
                        value={nuova.descrizione}
                        onChange={(e) => setNuova({ ...nuova, descrizione: e.target.value })}
                        onKeyDown={(e) => e.key === "Enter" && conferma()}
                        placeholder={t("cpr.descrizione")}
                        className={`${inputCls} w-full`}
                      />
                    </td>
                    <td className={`${cel} p-1`}>
                      <span className="flex items-center gap-2">
                        <label className="flex items-center gap-1 text-[12px] text-muted-foreground">
                          <input
                            type="checkbox"
                            checked={nuova.amex}
                            onChange={(e) => setNuova({ ...nuova, amex: e.target.checked })}
                          />
                          CDC
                        </label>
                        <label className="flex items-center gap-1 text-[12px] text-muted-foreground">
                          <input
                            type="checkbox"
                            checked={nuova.ricorrente}
                            onChange={(e) => setNuova({ ...nuova, ricorrente: e.target.checked })}
                          />
                          {t("cpr.ricorrente")}
                        </label>
                        <Ok onOk={conferma} onNo={() => setNuova(null)} t={t} />
                      </span>
                    </td>
                  </>
                ) : (
                  <>
                    <td className={`${cel} ${TINTA.ora.col}`} />
                    <td className={`${cel} ${TINTA.prossimi.col}`} />
                    <td className={`${cel} ${TINTA.standby.col}`} />
                    <td className={`${cel} ${TINTA.standby2.col}`} />
                    <td className={`${cel} py-1`}>
                      <Aggiungi onClick={() => apri(false)} t={t} />
                    </td>
                    <td className={cel} />
                  </>
                )}
              </tr>
            )}
          </tbody>
          <tfoot>
            {/* totali come nel foglio: valori e sotto le sue etichette */}
            <tr className="border-t-2 border-border bg-muted/50">
              <td className={`${celT} ${TINTA.entrate.num}`}>
                {fmt(totali.entrateStandby)}
                {(totali.entrateDaSollecitare ?? 0) > 0 && (
                  <div className="text-[10px] font-normal text-amber-700 dark:text-amber-400">
                    {t("cpr.diCuiSollecitare")} {fmt(totali.entrateDaSollecitare)}
                  </div>
                )}
                {(totali.compensazioni ?? 0) !== 0 && (
                  <div className="text-[10px] font-normal text-rose-700 dark:text-rose-400">
                    {t("cpr.diCuiComp")} {fmt(totali.compensazioni)}
                  </div>
                )}
              </td>
              <td className={`${celT} ${TINTA.standby.num} text-left`}>
                {fmt(totali.entrateParcheggiate)}
              </td>
              <td className={cel} />
              <td className={`${celT} text-sky-800 dark:text-sky-300`}>
                {totali.acconti > 0 ? `−${fmt(totali.acconti)}` : ""}
              </td>
              <td className={`${celT} ${TINTA.ora.num}`}>{fmt(totali.usciteOra)}</td>
              <td className={`${celT} ${TINTA.prossimi.num}`}>{fmt(totali.usciteProssimi)}</td>
              <td className={`${celT} ${TINTA.standby.num}`}>{fmt(totali.usciteStandby)}</td>
              <td className={`${celT} ${TINTA.standby2.num}`}>{fmt(totali.usciteStandby2)}</td>
              <td className={`${celT} text-foreground`}>{fmt(totali.totaleDaPagare)}</td>
              <td
                className={`${celT} ${totali.cassaConStandby < 0 ? "text-rose-600 dark:text-rose-400" : "text-emerald-700 dark:text-emerald-400"}`}
              >
                {fmt(totali.cassaConStandby)}
              </td>
            </tr>
            <tr className="bg-muted/50">
              <td className={`${celL} text-right`}>{t("cpr.lblStandbyIncasso")}</td>
              <td className={celL}>{t("cpr.lblEntrateStandby")}</td>
              <td className={cel} />
              <td className={`${celL} text-right`}>{totali.acconti > 0 ? t("cpr.acconti") : ""}</td>
              <td className={`${celL} text-right`}>{t("cpr.lblDaPagare")}</td>
              <td className={`${celL} text-right`}>
                {t("cpr.lblDaPagare")} {breve}
              </td>
              <td className={`${celL} text-right`}>{t("cpr.lblStandbyDaPagare")}</td>
              <td className={`${celL} text-right`}>{t("cpr.lblStandby2")}</td>
              <td className={`${celL} text-right`}>{t("cpr.lblTotaleDaPagare")}</td>
              <td className={`${celL} text-right`}>{t("cpr.lblCassa")}</td>
            </tr>
            <tr>
              <td className={`${cel} py-1.5`} colSpan={10} />
            </tr>
            <tr className="bg-primary/5">
              <td className={cel} />
              <td className={cel} />
              <td className={cel} />
              <td className={cel} />
              <td
                className={`${celT} ${totali.saldoDopoOra < 0 ? "text-rose-600 dark:text-rose-400" : "text-primary"}`}
              >
                {fmt(totali.saldoDopoOra)}
              </td>
              <td
                className={`${celT} ${totali.saldoDopoProssimi < 0 ? "text-rose-600 dark:text-rose-400" : "text-primary"}`}
              >
                {fmt(totali.saldoDopoProssimi)}
              </td>
              <td className={cel} />
              <td className={cel} />
              <td className={cel} />
              <td className={cel} />
            </tr>
            <tr className="bg-primary/5">
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
              <td className={`${celL} border-b-0 text-right`} colSpan={2}>
                {t("cpr.lblSaldoDopo")} · {t("cpr.lblDopoProssimi")} {breve}
              </td>
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
              <td className={`${cel} border-b-0`} />
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

const filtroCls =
  "h-7 rounded-md border bg-background px-2 text-[12px] text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";
const filtroOn = "border-primary ring-1 ring-primary/40";
const filtroOff = "border-border";

/** Campo di ricerca della barra dei filtri, con la × per svuotarlo. */
function CercaVoce({
  value,
  onChange,
  placeholder,
  t,
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder: string;
  t: TFn;
}) {
  return (
    <span className="relative inline-flex items-center">
      <Search className="pointer-events-none absolute left-2 h-3.5 w-3.5 text-muted-foreground" />
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className={`${filtroCls} w-56 pl-7 pr-6 ${value ? filtroOn : filtroOff}`}
      />
      {value && (
        <button
          type="button"
          onClick={() => onChange("")}
          title={t("cpr.fSvuota")}
          className="absolute right-1.5 text-muted-foreground hover:text-foreground"
        >
          <X className="h-3 w-3" />
        </button>
      )}
    </span>
  );
}

/** Barra dei filtri sopra il foglio (Simone 11/11/2026): entrate per testo e
 *  blocco; uscite per testo, colonna, tipologia, cliente/appalto, priorità e
 *  nota. Con un filtro acceso, la riga sotto dice quante voci restano e
 *  quanto sommano; i totali in fondo al foglio restano quelli di tutto. */
function BarraFiltri({
  t,
  filtri,
  onFiltri,
  entrateTutte,
  usciteTutte,
  limite,
}: {
  t: TFn;
  filtri: FiltriFoglio;
  onFiltri: (f: FiltriFoglio) => void;
  entrateTutte: readonly Cella[];
  usciteTutte: readonly Cella[];
  limite: string;
}) {
  const set = (p: Partial<FiltriFoglio>) => onFiltri({ ...filtri, ...p });
  const fE = filtroEntrateAttivo(filtri);
  const fU = filtroUsciteAttivo(filtri);
  const tipologie = distinti(usciteTutte.map(tipologiaDi));
  const clienti = distinti(usciteTutte.flatMap(clientiDi));
  const eFil = fE ? filtraEntrate(entrateTutte, filtri) : [];
  const uFil = fU ? filtraUscite(usciteTutte, filtri) : [];
  const nomeColonna: Record<SecchioCassa, string> = {
    ora: t("cpr.colUscite"),
    prossimi: etichettaProssimiBreve(limite),
    standby: t("cpr.colStandby"),
    standby2: t("cpr.colStandby2"),
  };
  const etichetta =
    "inline-flex h-7 items-center gap-1 rounded-md px-2 text-[11px] font-semibold uppercase tracking-wide";
  return (
    <div className="border-b border-border bg-muted/20 px-3 py-2">
      <div className="flex flex-wrap items-center gap-x-5 gap-y-2">
        {/* Entrate */}
        <div className="flex flex-wrap items-center gap-1.5">
          <span className={`${etichetta} ${TINTA.entrate.testa}`}>
            <Filter className="h-3 w-3" /> {t("cpr.colEntrate")}
          </span>
          <CercaVoce
            value={filtri.eTesto}
            onChange={(v) => set({ eTesto: v })}
            placeholder={t("cpr.fCercaEntrate")}
            t={t}
          />
          <select
            value={filtri.eBlocco}
            onChange={(e) => set({ eBlocco: e.target.value as FiltriFoglio["eBlocco"] })}
            className={`${filtroCls} ${filtri.eBlocco ? filtroOn : filtroOff}`}
          >
            <option value="">{t("cpr.fBlocchiTutti")}</option>
            <option value="gia">{t("cpr.giaEntrate")}</option>
            <option value="entranti">{t("cpr.entranti")}</option>
            <option value="sollecitare">{t("cpr.daSollecitare")}</option>
            <option value="standby">{t("cpr.colStandby")}</option>
          </select>
        </div>
        {/* Uscite */}
        <div className="flex flex-wrap items-center gap-1.5">
          <span className={`${etichetta} bg-primary text-primary-foreground`}>
            <Filter className="h-3 w-3" /> {t("cpr.colUscite")}
          </span>
          <CercaVoce
            value={filtri.uTesto}
            onChange={(v) => set({ uTesto: v })}
            placeholder={t("cpr.fCercaUscite")}
            t={t}
          />
          {COLONNE_USCITE.map((col) => {
            const on = filtri.uColonne.includes(col);
            return (
              <button
                key={col}
                type="button"
                aria-pressed={on}
                title={t("cpr.fColonnaTip")}
                onClick={() =>
                  set({
                    uColonne: on
                      ? filtri.uColonne.filter((x) => x !== col)
                      : [...filtri.uColonne, col],
                  })
                }
                className={`h-7 rounded-md border px-2 text-[11px] font-semibold ${on ? `${TINTA[col].testa} border-transparent` : "border-border bg-background text-muted-foreground hover:bg-muted"}`}
              >
                {nomeColonna[col]}
              </button>
            );
          })}
          {tipologie.length > 1 && (
            <select
              value={filtri.uTipologia}
              onChange={(e) => set({ uTipologia: e.target.value })}
              className={`${filtroCls} max-w-44 ${filtri.uTipologia ? filtroOn : filtroOff}`}
            >
              <option value="">{t("cpr.fTipologieTutte")}</option>
              {tipologie.map((x) => (
                <option key={x} value={x}>
                  {x === SENZA ? t("cpr.senzaTipologia") : x}
                </option>
              ))}
            </select>
          )}
          {clienti.length > 1 && (
            <select
              value={filtri.uCliente}
              onChange={(e) => set({ uCliente: e.target.value })}
              className={`${filtroCls} max-w-44 ${filtri.uCliente ? filtroOn : filtroOff}`}
            >
              <option value="">{t("cpr.fClientiTutti")}</option>
              {clienti.map((x) => (
                <option key={x} value={x}>
                  {x === SENZA ? t("cpr.senzaCliente") : x}
                </option>
              ))}
            </select>
          )}
          <select
            value={filtri.uPriorita}
            onChange={(e) => set({ uPriorita: e.target.value as FiltriFoglio["uPriorita"] })}
            className={`${filtroCls} ${filtri.uPriorita ? filtroOn : filtroOff}`}
          >
            <option value="">{t("cpr.fPrioritaTutte")}</option>
            <option value="con">{t("cpr.fPrioritaCon")}</option>
            <option value="h">
              {t("cpr.fPriorita")} {t("cpr.prioH")}
            </option>
            <option value="m">
              {t("cpr.fPriorita")} {t("cpr.prioM")}
            </option>
            <option value="l">
              {t("cpr.fPriorita")} {t("cpr.prioL")}
            </option>
          </select>
          <label
            className={`inline-flex h-7 cursor-pointer items-center gap-1.5 rounded-md border bg-background px-2 text-[12px] ${filtri.uNota ? `${filtroOn} text-foreground` : `${filtroOff} text-muted-foreground`}`}
          >
            <input
              type="checkbox"
              checked={filtri.uNota}
              onChange={(e) => set({ uNota: e.target.checked })}
            />
            {t("cpr.fConNota")}
          </label>
        </div>
        {(fE || fU) && (
          <button
            type="button"
            onClick={() => onFiltri(FILTRI_VUOTI)}
            className="inline-flex h-7 items-center gap-1 rounded-md border border-amber-500 bg-amber-500/10 px-2 text-[12px] font-semibold text-amber-800 hover:bg-amber-500/20 dark:text-amber-300"
          >
            <X className="h-3 w-3" /> {t("cpr.fTogli")}
          </button>
        )}
      </div>
      {(fE || fU) && (
        <div className="mt-2 flex flex-wrap items-center gap-x-5 gap-y-1 rounded-lg border border-amber-300/70 bg-amber-50 px-3 py-1.5 text-[12px] text-foreground dark:border-amber-700/50 dark:bg-amber-950/30">
          {fE && (
            <span>
              <span className={`font-semibold ${TINTA.entrate.num}`}>
                {t("cpr.fEntrateFiltrate")}
              </span>{" "}
              <b className="tabular-nums">{eFil.length}</b> {t("cpr.fVociDi")} {entrateTutte.length}{" "}
              · <b className="tabular-nums">{fmt(sommaCelle(eFil))}</b>
            </span>
          )}
          {fU && (
            <span className="flex flex-wrap items-center gap-x-3">
              <span>
                <span className="font-semibold text-primary">{t("cpr.fUsciteFiltrate")}</span>{" "}
                <b className="tabular-nums">{uFil.length}</b> {t("cpr.fVociDi")}{" "}
                {usciteTutte.length} · <b className="tabular-nums">{fmt(sommaCelle(uFil))}</b>
              </span>
              {COLONNE_USCITE.map((col) => {
                const s = sommaCelle(uFil.filter((u) => u.secchio === col));
                return s !== 0 ? (
                  <span key={col} className={`tabular-nums ${TINTA[col].num}`}>
                    {nomeColonna[col]} {fmt(s)}
                  </span>
                ) : null;
              })}
            </span>
          )}
          <span className="text-[11px] text-muted-foreground">{t("cpr.fNota")}</span>
        </div>
      )}
    </div>
  );
}

function Importo({
  c,
  readOnly,
  onImporto,
  t,
}: {
  c: Cella;
  readOnly: boolean;
  onImporto: (c: Cella, importo: number) => void;
  t: TFn;
}) {
  const [v, setV] = useState<string | null>(null);
  const corretto = Math.abs(c.importo - c.importoOrigine) > 0.005;
  const titolo = corretto ? `${t("cpr.importoOrigine")} ${fmt(c.importoOrigine)}` : c.dettaglio;
  if (readOnly || !c.editabile)
    return (
      <span title={titolo}>
        {fmt(c.importo)}
        {corretto && <span className="ml-0.5 text-[10px] text-muted-foreground">*</span>}
      </span>
    );
  if (v == null)
    return (
      <button
        type="button"
        onClick={() => setV(fmt(c.importo))}
        title={titolo}
        className="w-full text-right underline decoration-dotted decoration-border underline-offset-4 hover:decoration-foreground"
      >
        {fmt(c.importo)}
        {corretto && <span className="ml-0.5 text-[10px] text-muted-foreground">*</span>}
      </button>
    );
  const commit = () => {
    const n = parseNum(v);
    setV(null);
    if (n != null && n >= 0 && Math.abs(n - c.importo) > 0.005)
      onImporto(c, Math.round(n * 100) / 100);
  };
  return (
    <input
      autoFocus
      value={v}
      onChange={(e) => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") commit();
        if (e.key === "Escape") setV(null);
      }}
      className={`${inputCls} w-full text-right`}
    />
  );
}

function Nota({
  c,
  readOnly,
  onNota,
}: {
  c: Cella;
  readOnly: boolean;
  onNota: (c: Cella, nota: string) => void;
}) {
  const [v, setV] = useState<string | null>(null);
  if (readOnly) return <span className="px-3 py-1.5 text-foreground">{c.nota}</span>;
  return (
    <input
      value={v ?? c.nota}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => {
        if (v != null && v.trim() !== c.nota) onNota(c, v.trim());
        setV(null);
      }}
      onKeyDown={(e) => e.key === "Enter" && (e.target as HTMLInputElement).blur()}
      className="w-full bg-transparent px-3 py-1.5 text-[13px] text-foreground focus:bg-background focus:outline-none"
    />
  );
}

function Togli({ onClick, t }: { onClick: () => void; t: TFn }) {
  return (
    <button
      type="button"
      onClick={onClick}
      title={t("cpr.togliTip")}
      className="ml-auto shrink-0 text-muted-foreground opacity-0 group-hover/r:opacity-100 hover:text-status-absent"
    >
      <X className="h-3.5 w-3.5" />
    </button>
  );
}

function Aggiungi({ onClick, t }: { onClick: () => void; t: TFn }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="inline-flex items-center gap-1 text-[12px] text-muted-foreground hover:text-primary"
    >
      <Plus className="h-3.5 w-3.5" /> {t("cpr.aggiungi")}
    </button>
  );
}

function Ok({ onOk, onNo, t }: { onOk: () => void; onNo: () => void; t: TFn }): ReactNode {
  return (
    <span className="flex shrink-0 items-center gap-1">
      <button
        type="button"
        onClick={onOk}
        className="rounded border border-primary bg-primary px-2 py-0.5 text-[12px] text-primary-foreground"
      >
        {t("common.save")}
      </button>
      <button
        type="button"
        onClick={onNo}
        className="rounded border border-border px-2 py-0.5 text-[12px] hover:bg-muted"
      >
        {t("common.cancel")}
      </button>
    </span>
  );
}

// --- data attesa, priorità, agenda ------------------------------------------
const PRIO_CLS: Record<PrioritaCassa, string> = {
  h: "bg-rose-500 text-white",
  m: "bg-amber-500 text-white",
  l: "bg-emerald-500 text-white",
};
const PRIO_TXT: Record<PrioritaCassa, string> = {
  h: "text-rose-600 dark:text-rose-400",
  m: "text-amber-600 dark:text-amber-400",
  l: "text-emerald-600 dark:text-emerald-400",
};
function etichettaPriorita(p: PrioritaCassa | undefined, t: TFn): string {
  return p === "h"
    ? t("cpr.prioH")
    : p === "m"
      ? t("cpr.prioM")
      : p === "l"
        ? t("cpr.prioL")
        : t("cpr.senzaPriorita");
}

/** "attesa · 31/10": la data in cui il cliente ha detto che paga, cliccabile. */
function DataAttesa({
  c,
  readOnly,
  onData,
  t,
}: {
  c: Cella;
  readOnly: boolean;
  onData: (c: Cella, d: string) => void;
  t: TFn;
}) {
  const [edit, setEdit] = useState(false);
  const base =
    "shrink-0 whitespace-nowrap rounded-full border px-1.5 text-[10px] " +
    (c.dataScelta
      ? "border-emerald-500 bg-emerald-500/10 font-semibold text-emerald-800 dark:text-emerald-200"
      : "border-emerald-300 text-emerald-700 dark:border-emerald-700 dark:text-emerald-300");
  if (readOnly)
    return (
      <span className={base}>
        {t("cpr.attesa")} · {ddmmyy(c.data)}
      </span>
    );
  if (edit)
    return (
      <input
        type="date"
        autoFocus
        defaultValue={c.data}
        onBlur={(e) => {
          setEdit(false);
          if (e.target.value && e.target.value !== c.data) onData(c, e.target.value);
        }}
        onKeyDown={(e) => {
          if (e.key === "Escape") setEdit(false);
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        }}
        className={`${inputCls} py-0 text-[11px]`}
      />
    );
  return (
    <button
      type="button"
      onClick={() => setEdit(true)}
      title={t("cpr.dataAttesa")}
      className={base}
    >
      {t("cpr.attesa")} · {ddmmyy(c.data)}
    </button>
  );
}

/** Acconti già incassati sugli entranti del mese: campo sotto il "reale",
 *  nella colonna Reale (Simone 02/11/2026: "possibilità di inserire un
 *  acconto di entrata, che me lo va a scalare dal totale entrante"). */
function Acconti({
  importo,
  readOnly,
  onAcconti,
  t,
}: {
  importo: number;
  readOnly: boolean;
  onAcconti: (importo: number) => void;
  t: TFn;
}) {
  const [testo, setTesto] = useState<string | null>(null);
  if (readOnly)
    return importo > 0 ? (
      <div className="text-[10.5px] text-muted-foreground">
        {t("cpr.acconti")}: −{fmt(importo)}
      </div>
    ) : null;
  return (
    <label
      className="flex items-center justify-end gap-1 text-[10px] uppercase tracking-wider text-muted-foreground"
      title={t("cpr.accontiTip")}
    >
      {t("cpr.acconti")} −
      <input
        value={testo ?? (importo > 0 ? fmt(importo) : "")}
        onChange={(e) => setTesto(e.target.value)}
        onBlur={() => {
          if (testo == null) return;
          const n = testo.trim() ? parseNum(testo) : 0;
          setTesto(null);
          if (n != null && Math.abs(n - importo) > 0.005) onAcconti(Math.abs(n));
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
          if (e.key === "Escape") setTesto(null);
        }}
        placeholder="0,00"
        className={`${inputCls} w-24 py-0 text-right text-[12px] normal-case tracking-normal`}
      />
    </label>
  );
}

/** Nome della riga: per le righe scritte a mano si rinomina con un clic. */
function NomeRiga({
  c,
  readOnly,
  onRinomina,
  t,
  className,
}: {
  c: Cella;
  readOnly: boolean;
  onRinomina: (c: Cella, nome: string) => void;
  t: TFn;
  className: string;
}) {
  const [edit, setEdit] = useState(false);
  if (!eManuale(c) || readOnly || c.chiave.startsWith("grp|"))
    return (
      <span className={className} title={c.dettaglio}>
        {c.nome}
      </span>
    );
  if (edit)
    return (
      <input
        autoFocus
        defaultValue={c.nome}
        onBlur={(e) => {
          setEdit(false);
          const v = e.target.value.trim();
          if (v && v !== c.nome) onRinomina(c, v);
        }}
        onKeyDown={(e) => {
          if (e.key === "Escape") setEdit(false);
          if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        }}
        className={`${inputCls} w-full py-0 text-[12px]`}
      />
    );
  return (
    <button
      type="button"
      onClick={() => setEdit(true)}
      title={t("cpr.rinomina")}
      className={`${className} text-left underline decoration-dotted decoration-border underline-offset-4 hover:text-primary hover:decoration-primary`}
    >
      {c.nome}
    </button>
  );
}

/** Entrata attesa: da sollecitare (conta, ma sta nel suo blocco) o di nuovo
 *  tra le entranti normali. */
function ToggleSollecito({
  c,
  onSollecita,
  t,
}: {
  c: Cella;
  onSollecita: (c: Cella, on: boolean) => void;
  t: TFn;
}) {
  return (
    <button
      type="button"
      onClick={() => onSollecita(c, !c.daSollecitare)}
      title={c.daSollecitare ? t("cpr.tornaEntranti") : t("cpr.sollecitaTip")}
      className={`shrink-0 rounded border px-1.5 text-[9px] font-bold uppercase ${c.daSollecitare ? "border-emerald-500 bg-emerald-500/10 text-emerald-700 hover:bg-emerald-500/20 dark:text-emerald-300" : "border-amber-500 bg-amber-500/10 text-amber-700 hover:bg-amber-500/20 dark:text-amber-300"}`}
    >
      {c.daSollecitare ? `↩ ${t("cpr.entranti")}` : `📣 ${t("cpr.sollecita")}`}
    </button>
  );
}

/** Entrata attesa: in standby (non conta) o di nuovo tra le entranti. */
function ToggleStandby({
  c,
  onParcheggia,
  t,
}: {
  c: Cella;
  onParcheggia: (c: Cella, on: boolean) => void;
  t: TFn;
}) {
  return (
    <button
      type="button"
      onClick={() => onParcheggia(c, !c.parcheggiata)}
      title={c.parcheggiata ? t("cpr.riattivaEntrata") : t("cpr.parcheggiaEntrata")}
      className={`shrink-0 rounded border px-1.5 text-[9px] font-bold uppercase ${c.parcheggiata ? "border-emerald-500 bg-emerald-500/10 text-emerald-700 hover:bg-emerald-500/20 dark:text-emerald-300" : "border-slate-400 bg-slate-500/10 text-slate-700 hover:bg-slate-500/20 dark:text-slate-200"}`}
    >
      {c.parcheggiata ? `↩ ${t("cpr.entranti")}` : `⏸ ${t("cpr.colStandby")}`}
    </button>
  );
}

/** Pallino di priorità in Note2: clic per passare bassa → media → alta → nessuna. */
function Priorita({
  c,
  readOnly,
  onPriorita,
  t,
}: {
  c: Cella;
  readOnly: boolean;
  onPriorita: (c: Cella, p: PrioritaCassa | undefined) => void;
  t: TFn;
}) {
  const p = c.priorita;
  const next: PrioritaCassa | undefined =
    p === undefined ? "l" : p === "l" ? "m" : p === "m" ? "h" : undefined;
  const cls = p
    ? PRIO_CLS[p]
    : "border border-dashed border-border text-muted-foreground opacity-0 group-hover/r:opacity-100";
  if (readOnly)
    return p ? (
      <span className={`mr-1 rounded px-1 text-[10px] font-bold uppercase ${PRIO_CLS[p]}`}>
        {p}
      </span>
    ) : null;
  return (
    <button
      type="button"
      onClick={() => onPriorita(c, next)}
      title={`${t("cpr.priorita")}: ${etichettaPriorita(p, t)}`}
      className={`mr-1 h-5 w-5 shrink-0 rounded text-[10px] font-bold uppercase ${cls}`}
    >
      {p ?? "·"}
    </button>
  );
}

function AgendaBox({
  t,
  voci,
  onAggiungi,
  onToggle,
  onElimina,
  onTogliRiga,
}: {
  t: TFn;
  voci: readonly VoceAgendaVista[];
  onAggiungi: (testo: string, p: PrioritaCassa, scad?: string) => void;
  onToggle: (id: string) => void;
  onElimina: (id: string) => void;
  onTogliRiga: (chiave: string) => void;
}) {
  const [testo, setTesto] = useState("");
  const [p, setP] = useState<PrioritaCassa>("m");
  const [scad, setScad] = useState("");
  const invia = () => {
    if (!testo.trim()) return;
    onAggiungi(testo.trim(), p, scad || undefined);
    setTesto("");
    setScad("");
  };
  const aperte = voci.filter((v) => !v.fatta);
  const conta = (q: PrioritaCassa) => aperte.filter((v) => v.p === q).length;
  return (
    <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
      <div className="mb-2 flex items-center gap-2 text-[13px] font-semibold text-foreground">
        <ListTodo className="h-4 w-4 text-primary" /> {t("cpr.agenda")}
        <span className="ml-auto flex items-center gap-1 text-[11px] font-normal">
          {(["h", "m", "l"] as const).map((q) => (
            <span key={q} className={`rounded px-1.5 ${PRIO_CLS[q]}`}>
              {q.toUpperCase()} {conta(q)}
            </span>
          ))}
        </span>
      </div>
      <div className="mb-3 flex flex-wrap items-center gap-2">
        <input
          value={testo}
          onChange={(e) => setTesto(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && invia()}
          placeholder={t("cpr.nuovaVoce")}
          className={`${inputCls} min-w-[180px] flex-1`}
        />
        <span className="flex overflow-hidden rounded-lg border border-border text-[11px] font-bold uppercase">
          {(["l", "m", "h"] as const).map((q) => (
            <button
              key={q}
              type="button"
              onClick={() => setP(q)}
              title={etichettaPriorita(q, t)}
              className={`px-2 py-1 ${p === q ? PRIO_CLS[q] : "text-muted-foreground hover:bg-muted"}`}
            >
              {q}
            </button>
          ))}
        </span>
        <input
          type="date"
          value={scad}
          onChange={(e) => setScad(e.target.value)}
          className={inputCls}
        />
        <button
          type="button"
          onClick={invia}
          className="inline-flex items-center gap-1 rounded-lg border border-primary bg-primary px-2.5 py-1 text-[12px] text-primary-foreground"
        >
          <Plus className="h-3.5 w-3.5" /> {t("cpr.aggiungi")}
        </button>
      </div>
      {voci.length === 0 ? (
        <p className="text-[12px] text-muted-foreground">{t("cpr.agendaVuota")}</p>
      ) : (
        <ul className="space-y-1">
          {voci.map((v) => (
            <li
              key={v.id}
              className={`group/a flex items-start gap-2 rounded-lg border-l-4 px-2 py-1 text-[13px] ${
                v.p === "h"
                  ? "border-rose-500 bg-rose-500/5"
                  : v.p === "m"
                    ? "border-amber-500 bg-amber-500/5"
                    : "border-emerald-500 bg-emerald-500/5"
              } ${v.fatta ? "opacity-50" : ""}`}
            >
              {v.rigaChiave ? (
                <span
                  className={`mt-0.5 rounded px-1 text-[10px] font-bold uppercase ${PRIO_CLS[v.p]}`}
                >
                  {v.p}
                </span>
              ) : (
                <input
                  type="checkbox"
                  checked={v.fatta}
                  onChange={() => onToggle(v.id)}
                  title={t("cpr.fatta")}
                  className="mt-1"
                />
              )}
              <span className={`flex-1 ${v.fatta ? "line-through" : "text-foreground"}`}>
                {v.testo}
                {v.riga && (
                  <span className="ml-1 text-[11px] text-muted-foreground">
                    · {v.riga} · {t("cpr.rigaAgenda")}
                  </span>
                )}
                {v.data && (
                  <span
                    className={`ml-1 text-[11px] ${!v.fatta && v.data < oggiISO() ? PRIO_TXT.h : "text-muted-foreground"}`}
                  >
                    {ddmm(v.data)}
                  </span>
                )}
              </span>
              <button
                type="button"
                onClick={() => (v.rigaChiave ? onTogliRiga(v.rigaChiave) : onElimina(v.id))}
                className="text-muted-foreground opacity-0 hover:text-status-absent group-hover/a:opacity-100"
                title={t("common.delete")}
              >
                <X className="h-3.5 w-3.5" />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
