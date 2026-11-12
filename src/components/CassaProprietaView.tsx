// DR Portal — Finanza → Flussi → vista "Cassa proprietà".
// IL FOGLIO del commercialista, tale e quale: colonne Entrate | Note |
// Saldo | Uscite | 1-5 del mese dopo | Standby | Note | Note2, i totali in
// fondo con le sue etichette, le fotografie come le linguette dei fogli
// Excel ("1210"/"1510"). Niente cruscotti: la proprietà legge solo questo.
// Si lavora dentro le celle: clic su una cella vuota delle tre colonne per
// spostare un'uscita, clic sull'importo per correggerlo (stipendi, fiscale,
// fissi, attesi), Note2 è una cella scrivibile, la × a bordo riga toglie la
// riga. Logica in lib/cassa-logic; dati in FinanzaData/cassa.json.
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { toast } from "sonner";
import {
  CalendarDays,
  Camera,
  Download,
  Filter,
  ListTodo,
  Loader2,
  Lock,
  Minimize2,
  Pencil,
  Plus,
  Printer,
  Search,
  StickyNote,
  Tv,
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
  type BloccoEntrate,
  type FiltriFoglio,
} from "@/lib/cassa-filtri";
import {
  apriStampa,
  csvSezioni,
  htmlStampaSezioni,
  sezioniCassa,
  type EtichetteStampa,
  type SezioneCassa,
  type SezioneCassaId,
} from "@/lib/cassa-stampa";
import { esportaCsvFile } from "@/lib/csv";
import type { StipendiDb } from "@/lib/stipendi-logic";
import {
  agendaCassa,
  calendarioEntrate,
  costruisciCassa,
  creaFotografia,
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
  type AltreSpeseCassa,
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
  /** Entrata attesa messa in standby: a mano = blocco "Standby 2" (non la
   *  incasseremo); dall'accordo (`daAccordo`) = blocco "Standby". */
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
  /** Entrate attese: chiave della controparte, per l'accordo sugli entranti. */
  chiaveControparte?: string;
  /** Riga dell'accordo (quanto entra); totale = quanto era atteso. */
  accordo?: { totale: number };
  /** In standby per via dell'accordo, col residuo non coperto. */
  daAccordo?: boolean;
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
  chiaveControparte: r.chiaveControparte,
  accordo: r.accordo,
  daAccordo: r.daAccordo,
});

/** La colonna Entrate in tre blocchi col loro totale: "Già entrate" (in
 *  banca, dentro il saldo), "Entranti" (attese = standby incasso, al netto
 *  degli acconti) e "Standby" (attese parcheggiate, non contano) — Simone
 *  19/10 e 02/11/2026. */
type SlotE =
  /** Testata di un blocco: somma delle voci che ha sotto, quante sono, e se
   *  è richiuso (Simone 11/11/2026: "chiudere anche entrate, entranti e
   *  standby con un +"). */
  | {
      kind: "head";
      id: BloccoEntrate;
      label: string;
      value: number;
      n: number;
      chiuso: boolean;
    }
  | { kind: "riga"; c: Cella; figlio?: boolean }
  /** Cliente con più fatture attese: una riga sola col totale e il "+" che
   *  apre le singole fatture (Simone 10/11/2026, "un + per cliente"). */
  | { kind: "padre"; c: Cella; figli: Cella[]; aperto: boolean };
function slotsEntrate(
  entrate: readonly Cella[],
  t: TFn,
  espansi: ReadonlySet<string> = new Set(),
  /** Blocchi richiusi ("e:gia", "e:entranti"…): resta solo la testata. */
  chiusi: ReadonlySet<string> = new Set(),
): SlotE[] {
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
      // La riga del cliente prende l'aspetto di una sua fattura, non della
      // compensazione: così resta la casella "Entra ora" anche nello Standby.
      const primo = figli.find((c) => c.tipo !== "compensazione") ?? figli[0];
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
  const blocco = (id: BloccoEntrate) => entrate.filter((c) => bloccoEntrata(c) === id);
  const gia = blocco("gia");
  const ent = blocco("entranti");
  const sol = blocco("sollecitare");
  const stb = blocco("standby");
  const stb2 = blocco("standby2");
  const blocchi: { id: BloccoEntrate; label: string; celle: Cella[]; slots: () => SlotE[] }[] = [
    {
      id: "gia",
      label: t("cpr.giaEntrate"),
      celle: gia,
      slots: () => gia.map((c): SlotE => ({ kind: "riga", c })),
    },
    { id: "entranti", label: t("cpr.entranti"), celle: ent, slots: () => conPadri(ent, "cli") },
    {
      id: "sollecitare",
      label: t("cpr.daSollecitare"),
      celle: sol,
      slots: () => conPadri(sol, "sol"),
    },
    // Due standby, come per le uscite (Simone 12/11/2026): "Standby" è il
    // resto degli accordi, quello che incasseremo più avanti; "Standby 2"
    // sono le attese messe da parte a mano, quelle che non incasseremo.
    { id: "standby", label: t("cpr.eStandby"), celle: stb, slots: () => conPadri(stb, "stb") },
    {
      id: "standby2",
      label: t("cpr.eStandby2"),
      celle: stb2,
      slots: () => conPadri(stb2, "stb2"),
    },
  ];
  // Ogni testata porta la somma delle SUE voci (così i numeri tornano a
  // colpo d'occhio); i blocchi senza voci non compaiono.
  return blocchi
    .filter((b) => b.celle.length > 0)
    .flatMap((b): SlotE[] => {
      const chiuso = chiusi.has(`e:${b.id}`);
      return [
        {
          kind: "head",
          id: b.id,
          label: b.label,
          value: sommaCelle(b.celle),
          n: b.celle.length,
          chiuso,
        },
        ...(chiuso ? [] : b.slots()),
      ];
    });
}

/** Le uscite in blocchi come le entrate (Simone 11/11/2026: "le uscite
 *  imminenti salgono in alto, poi le standby e poi le standby 2, tutto
 *  chiudibile con il +"): imminenti, 1-5 del mese dopo, standby, standby 2.
 *  Dentro ogni blocco resta l'ordine del foglio (scadenza e nome). */
type SlotU =
  | { kind: "head"; col: SecchioCassa; label: string; value: number; n: number; chiuso: boolean }
  | { kind: "riga"; c: Cella; aperta?: boolean }
  /** Una voce dentro la riga aperta col "+" (le tipologie delle Altre spese). */
  | { kind: "voce"; padre: Cella; v: VoceDettaglioCassa };
/** Le righe che si aprono col "+" per vedere da cosa sono composte. */
const haVoci = (c: Cella) => c.tipo === "altre" && (c.voci?.length ?? 0) > 0;
function slotsUscite(
  uscite: readonly Cella[],
  t: TFn,
  limite: string,
  chiusi: ReadonlySet<string> = new Set(),
  espansi: ReadonlySet<string> = new Set(),
): SlotU[] {
  const etichetta: Record<SecchioCassa, string> = {
    ora: t("cpr.vUsciteOra"),
    prossimi: `${t("cpr.colUscite")} ${etichettaProssimi(limite)}`,
    standby: t("cpr.vStandbyUscite"),
    standby2: t("cpr.vStandby2Uscite"),
  };
  return COLONNE_USCITE.flatMap((col): SlotU[] => {
    const celle = uscite.filter((u) => u.secchio === col);
    if (!celle.length) return [];
    const chiuso = chiusi.has(`u:${col}`);
    return [
      {
        kind: "head",
        col,
        label: etichetta[col],
        value: sommaCelle(celle),
        n: celle.length,
        chiuso,
      },
      ...(chiuso
        ? []
        : celle.flatMap((c): SlotU[] => {
            const aperta = haVoci(c) && espansi.has(c.chiave);
            return [
              { kind: "riga", c, aperta },
              ...(aperta ? (c.voci ?? []).map((v): SlotU => ({ kind: "voce", padre: c, v })) : []),
            ];
          })),
    ];
  });
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
    daAccordo: r.sa ? true : undefined,
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
  altreSpese,
  movimenti,
  onApriFornitore,
  mezzi,
  gruppi,
  avviso,
}: {
  /** Avviso da mostrare in testa (dati non caricati): sta dentro la vista
   *  perché si veda anche nella modalità TV. */
  avviso?: ReactNode;
  attive: readonly FatturaConStato[];
  passive: readonly FatturaConStato[];
  esclusioni: readonly { nome: string; mese?: string; meseFine?: string }[];
  infoFornitori: ReadonlyMap<string, FornitoreInfo>;
  fiscale: readonly ScadenzaFiscale[];
  fissi: readonly CostoFisso[];
  /** "Altre spese" stimate della vista classica (facoltative). */
  altreSpese?: AltreSpeseCassa;
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
  // SCHERMO TV (Simone 12/11/2026: "ottimizza la vista per la tv, me la
  // taglia sempre a destra"): la cassa a tutto schermo, senza menù né
  // testata, in sola lettura (senza i pulsanti di lavoro, che allargano il
  // foglio), disegnata sempre alla stessa larghezza e poi ingrandita o
  // rimpicciolita per riempire lo schermo: non resta mai tagliata.
  const [tv, setTv] = useState(false);
  // Dalla 1.113.1 la vista TV parte CON i pulsanti di lavoro: in riunione si
  // spostano le voci e si correggono gli importi dalla TV (Simone 12/11/2026:
  // "non mi fa spostare le voci, non mi fa cambiare le cifre"). Togliendo la
  // spunta "modifica" si ha la sola lettura.
  const [tvModifica, setTvModifica] = useState(true);
  const [finestra, setFinestra] = useState({ w: 0, h: 0 });
  useEffect(() => {
    const misura = () => setFinestra({ w: window.innerWidth, h: window.innerHeight });
    const tasto = (e: KeyboardEvent) => {
      const dentroUnCampo = (e.target as HTMLElement | null)?.closest?.("input,textarea,select");
      if (e.key === "Escape" && !dentroUnCampo) setTv(false);
    };
    // Uscendo dallo schermo intero del browser si esce anche dalla vista TV.
    const schermo = () => {
      if (!document.fullscreenElement) setTv(false);
    };
    const scorrimento = document.body.style.overflow;
    if (tv) {
      misura();
      document.body.style.overflow = "hidden";
      window.addEventListener("resize", misura);
      window.addEventListener("keydown", tasto);
      document.addEventListener("fullscreenchange", schermo);
    }
    return () => {
      document.body.style.overflow = scorrimento;
      window.removeEventListener("resize", misura);
      window.removeEventListener("keydown", tasto);
      document.removeEventListener("fullscreenchange", schermo);
    };
  }, [tv]);
  const entraTv = () => {
    setTvModifica(true);
    setTv(true);
    // Schermo intero del browser, se lo concede (sulla TV toglie le barre).
    try {
      void document.documentElement.requestFullscreen?.().catch(() => undefined);
    } catch {
      // senza schermo intero resta comunque la vista TV
    }
  };
  const esciTv = () => {
    setTv(false);
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
  };
  // Schede che si aprono dai nomi del foglio (Simone 04/11/2026): cliente
  // (le sue fatture aperte) e gruppo (i fornitori dentro); i fornitori
  // aprono la scheda dei Flussi tramite onApriFornitore.
  const [clienteAperto, setClienteAperto] = useState<string | null>(null);
  const [gruppoAperto, setGruppoAperto] = useState<Cella | null>(null);
  const [fiscaleAperto, setFiscaleAperto] = useState<Cella | null>(null);

  // LETTURA FALLITA (1.115.0, incidente del 12/11/2026: Microsoft 365 ha
  // rifiutato le letture per qualche minuto e il foglio è uscito vuoto, col
  // saldo a zero, senza dire niente). Se il foglio salvato non si legge NON
  // si parte da un foglio vuoto (sembrava "sparito tutto"): resta l'avviso
  // col pulsante per riprovare. Il saldo non letto si scrive in rosso.
  const [dbErr, setDbErr] = useState<string | null>(null);
  const caricaDb = () => {
    setDbErr(null);
    spCassaGet()
      .then((d) => {
        setDb(d);
        setLimite(limiteDefault(oggiISO(), d.giornoLimite));
      })
      .catch((e) => setDbErr(e instanceof Error ? e.message : String(e)));
  };
  // Saldo dalla cache del server (una riga SharePoint, niente banca e
  // niente rilettura dei movimenti): arriva subito.
  const caricaSaldo = () => {
    spEbSaldoCache()
      .then((s) => {
        if (s) {
          saldoMemo = { saldo: s.saldo, al: s.aggiornatoAl };
          setSaldoBanca(saldoMemo);
          setSaldoErr(false);
        } else setSaldoErr(true);
      })
      .catch(() => setSaldoErr(true));
  };
  useEffect(() => {
    caricaDb();
    spStipendiGet()
      .then((d) => setStipendiDb(d))
      .catch(() => setStipendiDb(null));
    caricaSaldo();
    // Una volta sola, all'apertura.
  }, []);

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
      altreSpese,
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
    altreSpese,
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
  // Accordo sugli entranti (Simone 12/11/2026): quanto entra di quel cliente;
  // null = nessun accordo (entra tutto), 0 = non entra nulla (tutto in standby).
  const setAccordo = (c: Cella, importo: number | null) => {
    if (!db || !c.chiaveControparte) return;
    const accordi = { ...(db.accordi ?? {}) };
    if (importo == null) delete accordi[c.chiaveControparte];
    else accordi[c.chiaveControparte] = Math.round(Math.abs(importo) * 100) / 100;
    void salva({ ...db, accordi: Object.keys(accordi).length ? accordi : undefined });
  };
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
      <div className="space-y-4">
        {avviso}
        {dbErr ? (
          <div
            role="alert"
            className="rounded-xl border-2 border-rose-400 bg-rose-50 px-4 py-3 text-rose-900 shadow-sm dark:border-rose-700 dark:bg-rose-950/50 dark:text-rose-100"
          >
            <div className="text-sm font-bold uppercase tracking-wide">
              {t("cpr.errLetturaTitolo")}
            </div>
            <p className="mt-1.5 text-[13px]">{t("cpr.errLetturaTesto")}</p>
            <p className="mt-1 text-[11px] opacity-70">{dbErr}</p>
            <button
              type="button"
              onClick={caricaDb}
              className="mt-2 rounded-lg border border-rose-500 bg-white px-3 py-1 text-[13px] font-semibold text-rose-800 hover:bg-rose-100 dark:bg-rose-900 dark:text-rose-50"
            >
              {t("fc.riprova")}
            </button>
          </div>
        ) : (
          <div className="flex items-center gap-2 text-sm text-muted-foreground">
            <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
          </div>
        )}
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
  const usciteFiltrate = filtraUscite(usciteTutte, filtri);
  const uscite = raggruppaUscite(usciteFiltrate, raggruppa, t);
  const totali = foto ? foto.totali : esito.totali;
  const limiteVista = foto ? foto.limite : limite;
  const saldoVista = foto ? foto.saldo : saldo;
  const dataVista = foto ? foto.data : oggi;
  const calendario = foto ? [] : calendarioEntrate(esito.righe);
  const agenda = foto ? [] : agendaCassa(esito.righe, db.agenda ?? []);

  const righeCsv = () => {
    const out: (string | number)[][] = [];
    const num = (v: number | undefined) => (v ?? 0).toFixed(2).replace(".", ",");
    // L'export è sempre il foglio INTERO, con tutti i blocchi aperti: filtri
    // e blocchi richiusi sono solo vista.
    const sl = slotsEntrate(entrateTutte, t, new Set(["*"]));
    const su = slotsUscite(raggruppaUscite(usciteTutte, raggruppa, t), t, limiteVista);
    const n = Math.max(sl.length, su.length);
    for (let i = 0; i < n; i++) {
      const s = sl[i];
      const e = s?.kind === "riga" ? s.c : undefined;
      const x = su[i];
      const u = x?.kind === "riga" ? x.c : undefined;
      const colonna = (col: SecchioCassa) =>
        x?.kind === "head"
          ? x.col === col
            ? num(x.value)
            : ""
          : u && u.secchio === col
            ? num(u.importo)
            : "";
      // I "padri" (cliente con più fatture) non si esportano: escono le
      // singole fatture, così i conti non raddoppiano.
      if (s?.kind === "padre" && !x) continue;
      out.push([
        s?.kind === "head" ? num(s.value) : e ? num(e.importo) : "",
        s?.kind === "head"
          ? `${s.label.toUpperCase()} (${s.n})`
          : e
            ? `${e.nome}${e.attesa ? ` (${t("cpr.attesa")} ${ddmmyy(e.data)})` : ""}`
            : "",
        colonna("ora"),
        colonna("prossimi"),
        colonna("standby"),
        colonna("standby2"),
        x?.kind === "head" ? `${x.label.toUpperCase()} (${x.n})` : u ? u.nome : "",
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
    // Il conto in fila, come nella banda in alto.
    const breve = etichettaProssimiBreve(limiteVista);
    const riga = (etichetta: string, v: number | undefined) =>
      out.push(["", "", "", "", "", "", etichetta, num(v)]);
    out.push(["", "", "", "", "", "", "", ""]);
    riga(t("cpr.saldoBanca"), saldoVista);
    riga(`- ${t("cpr.vUsciteOra")}`, totali.usciteOra);
    riga(`= ${t("cpr.rDopoOra")}`, totali.saldoDopoOra);
    riga(`- ${t("cpr.colUscite")} ${breve}`, totali.usciteProssimi);
    riga(`= ${t("cpr.rDopoProssimi")} ${breve}`, totali.saldoDopoProssimi);
    riga(
      `+ ${t("cpr.entranti")} (${t("cpr.acconti")} ${num(totali.acconti)})`,
      totali.entrantiNette,
    );
    riga(`= ${t("cpr.colReale")}`, totali.reale);
    riga(`- ${t("cpr.vStandbyUscite")}`, totali.usciteStandby);
    riga(`= ${t("cpr.rCassa")}`, totali.cassaConStandby);
    riga(`${t("cpr.rFuori")}: ${t("cpr.vStandbyEntrate")}`, totali.entrateFuture);
    riga(`${t("cpr.rFuori")}: ${t("cpr.vStandby2Entrate")}`, totali.entrateParcheggiate);
    riga(`${t("cpr.rFuori")}: ${t("cpr.vStandby2Uscite")}`, totali.usciteStandby2);
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

  // Vista TV: tutto disegnato a LARGHEZZA_TV e poi portato alla larghezza
  // dello schermo (più grande su una TV grande, più piccolo se serve).
  const zoomTv = tv && finestra.w > 0 ? Math.round((finestra.w / LARGHEZZA_TV) * 1000) / 1000 : 1;
  const soloLettura = Boolean(foto) || (tv && !tvModifica);

  return (
    <div
      className={tv ? "fixed inset-0 z-40 space-y-4 overflow-auto bg-background p-4" : "space-y-4"}
      style={tv && zoomTv !== 1 ? { zoom: zoomTv } : undefined}
    >
      {/* Dati non caricati: l'avviso sta QUI dentro, così si vede anche
          nella vista TV, che copre il resto della pagina. */}
      {avviso}
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
            <div
              className={`mt-1 text-[11px] ${saldoErr && !saldoBanca && !foto ? "font-semibold text-rose-600 dark:text-rose-400" : "text-muted-foreground"}`}
            >
              {saldoNota}
              {saldoErr && !saldoBanca && !foto && (
                <button
                  type="button"
                  onClick={caricaSaldo}
                  className="ml-2 underline underline-offset-2 hover:opacity-80"
                >
                  {t("fc.riprova")}
                </button>
              )}
            </div>
          </div>
          <RiepilogoCassa
            t={t}
            totali={totali}
            limite={limiteVista}
            readOnly={soloLettura}
            onAcconti={setAcconti}
          />
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
            {tv ? (
              <>
                <label
                  className="flex cursor-pointer items-center gap-1.5 text-muted-foreground"
                  title={t("cpr.tvModificaTip")}
                >
                  <input
                    type="checkbox"
                    checked={tvModifica}
                    onChange={(e) => setTvModifica(e.target.checked)}
                  />
                  {t("cpr.tvModifica")}
                </label>
                <button
                  type="button"
                  onClick={esciTv}
                  className="inline-flex items-center gap-2 rounded-lg border border-primary bg-primary px-3 py-1.5 text-sm text-primary-foreground shadow-sm hover:opacity-90"
                >
                  <Minimize2 className="h-4 w-4" /> {t("cpr.tvEsci")}
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={entraTv}
                title={t("cpr.tvEntraTip")}
                className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground transition-colors hover:bg-secondary"
              >
                <Tv className="h-4 w-4" /> {t("cpr.tvEntra")}
              </button>
            )}
            <MenuVoci
              t={t}
              entrate={entrate}
              uscite={usciteFiltrate}
              limite={limiteVista}
              data={dataVista}
              filtrate={filtroEntrateAttivo(filtri) || filtroUsciteAttivo(filtri)}
            />
            <EsportaStampa
              nomeFile={`cassa-proprieta-${dataVista}`}
              testata={[
                t("cpr.colEntrate"),
                t("cpr.colNote"),
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
          <FuoriConti t={t} totali={totali} />
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
          altezzaMax={
            tv && finestra.h > 0
              ? Math.max(320, Math.round((finestra.h / zoomTv) * 0.88))
              : undefined
          }
          limite={limiteVista}
          readOnly={soloLettura}
          onSposta={sposta}
          onImporto={correggi}
          onNota={annota}
          onTogli={togli}
          onAggiungi={aggiungiManuale}
          onPriorita={setPriorita}
          onData={setData}
          onParcheggia={parcheggia}
          onSollecita={sollecita}
          onAccordo={setAccordo}
          accordi={foto ? undefined : db.accordi}
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
        {/* Già pagate in banca ma ancora aperte su Aruba: fuori dalle uscite
            (il saldo è già sceso), in attesa che Aruba le chiuda. La × dice
            "non era quel pagamento" e rimette la fattura tra le uscite. */}
        {!foto && esito.pagateInBanca.length > 0 && (
          <div className="border-t border-border px-3 py-1.5 text-[11px] text-muted-foreground">
            <span className="mr-1 font-bold text-emerald-700 dark:text-emerald-400">✓</span>
            {t("cpr.pagateBanca")}{" "}
            <span className="font-semibold text-foreground">
              {esito.pagateInBanca.length} · {fmt(totali.usciteGiaPagate ?? 0)} €
            </span>{" "}
            (
            {esito.pagateInBanca.slice(0, 10).map((r, i) => (
              <span key={r.nomeFile} className="whitespace-nowrap">
                {i > 0 && ", "}
                {r.fornitore} {r.numero}{" "}
                <span className="tabular-nums">
                  {fmt(r.importo)} · {ddmm(r.data)}
                </span>
                {!soloLettura && (
                  <button
                    type="button"
                    onClick={() => setStato(`pagb|${r.nomeFile}`, { x: true })}
                    title={t("cpr.pagateBancaNoTip")}
                    className="ml-0.5 rounded px-1 text-muted-foreground hover:bg-muted hover:text-rose-600"
                  >
                    ×
                  </button>
                )}
              </span>
            ))}
            {esito.pagateInBanca.length > 10 ? ", …" : ""}) · {t("cpr.pagateBancaNota")}
          </div>
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

/** Un passo del conto: cosa si toglie o si aggiunge, e dove si arriva. */
function Passo({
  segno,
  voce,
  importo,
  tinta,
  titolo,
  valore,
  forte,
  children,
}: {
  segno: "−" | "+";
  voce: string;
  importo: number;
  tinta: string;
  titolo: string;
  valore: number;
  /** Risultato principale: verde se positivo. */
  forte?: boolean;
  /** Accanto alla prima riga (gli acconti, per gli entranti). */
  children?: ReactNode;
}) {
  const colore =
    valore < 0
      ? "text-rose-600 dark:text-rose-400"
      : forte
        ? "text-emerald-700 dark:text-emerald-400"
        : "text-primary";
  return (
    <div className="border-l-2 border-border pl-4">
      <div className="flex items-center gap-3 whitespace-nowrap">
        <span className={`text-[12px] tabular-nums ${tinta}`}>
          <b>{segno}</b> {voce} <b>{fmt(importo)}</b>
        </span>
        {children}
      </div>
      <div className="mt-1 max-w-[230px] text-[11px] uppercase tracking-wider text-muted-foreground">
        = {titolo}
      </div>
      <div className={`mt-0.5 whitespace-nowrap text-xl font-semibold tabular-nums ${colore}`}>
        {fmt(valore)} €
      </div>
    </div>
  );
}

/** Il conto della cassa letto in fila (Simone 11/11/2026: "rivalutiamo tutte
 *  le somme, la lettura deve essere più pulita"). Dal saldo in banca si
 *  tolgono le uscite imminenti, poi quelle dell'1-5 del mese dopo, si
 *  aggiungono gli entranti (meno gli acconti) e si arriva al Reale; tolto
 *  anche lo standby, alla Cassa. Sono gli stessi numeri e le stesse formule
 *  di prima (colonne Saldo e Reale, riga dei totali in fondo): ora stanno
 *  una volta sola, qui. */
export function RiepilogoCassa({
  t,
  totali,
  limite,
  readOnly,
  onAcconti,
}: {
  t: TFn;
  totali: TotaliCassa;
  limite: string;
  readOnly: boolean;
  onAcconti: (importo: number) => void;
}) {
  const breve = etichettaProssimiBreve(limite);
  // Le fotografie più vecchie non hanno tutti i totali: mancanti = zero.
  const n = (v: number | undefined) => v ?? 0;
  return (
    <>
      <Passo
        segno="−"
        voce={t("cpr.vUsciteOra")}
        importo={n(totali.usciteOra)}
        tinta={TINTA.ora.num}
        titolo={t("cpr.rDopoOra")}
        valore={n(totali.saldoDopoOra)}
      />
      <Passo
        segno="−"
        voce={`${t("cpr.colUscite")} ${breve}`}
        importo={n(totali.usciteProssimi)}
        tinta={TINTA.prossimi.num}
        titolo={`${t("cpr.rDopoProssimi")} ${breve}`}
        valore={n(totali.saldoDopoProssimi)}
      />
      <Passo
        segno="+"
        voce={t("cpr.entranti")}
        importo={n(totali.entrantiNette)}
        tinta={TINTA.entrate.num}
        titolo={t("cpr.colReale")}
        valore={n(totali.reale)}
        forte
      >
        <Acconti importo={n(totali.acconti)} readOnly={readOnly} onAcconti={onAcconti} t={t} />
      </Passo>
      <Passo
        segno="−"
        voce={t("cpr.vStandbyUscite")}
        importo={n(totali.usciteStandby)}
        tinta={TINTA.standby.num}
        titolo={t("cpr.rCassa")}
        valore={n(totali.cassaConStandby)}
        forte
      />
    </>
  );
}

/** Quello che NON entra nel conto, detto una volta sola sotto la banda. */
export function FuoriConti({ t, totali }: { t: TFn; totali: TotaliCassa }) {
  const voci: [string, number | undefined][] = [
    [t("cpr.vStandbyEntrate"), totali.entrateFuture],
    [t("cpr.vStandby2Entrate"), totali.entrateParcheggiate],
    [t("cpr.vStandby2Uscite"), totali.usciteStandby2],
    [t("cpr.rGiaEntrate"), totali.entrateRicevute],
  ];
  const visibili = voci.filter(([, v]) => Math.abs(v ?? 0) > 0.005);
  if (!visibili.length) return null;
  return (
    <div className="basis-full text-[11.5px] text-muted-foreground">
      <span className="font-semibold uppercase tracking-wider">{t("cpr.rFuori")}</span>
      {visibili.map(([nome, v]) => (
        <span key={nome} className="ml-3 whitespace-nowrap">
          {nome} <b className="tabular-nums text-foreground">{fmt(v ?? 0)}</b>
        </span>
      ))}
    </div>
  );
}

/** Il "+" e il "−" che aprono e richiudono un blocco del foglio. */
function PiuMeno({ aperto, onClick, t }: { aperto: boolean; onClick: () => void; t: TFn }) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-expanded={aperto}
      title={t(aperto ? "cpr.bChiudi" : "cpr.bApri")}
      className="inline-flex h-[18px] w-[18px] shrink-0 items-center justify-center rounded border border-current bg-background/80 text-[13px] font-bold leading-none hover:bg-background"
    >
      {aperto ? "−" : "+"}
    </button>
  );
}

/** Tinta della testata di ogni blocco: entrate e uscite. */
const TESTA_ENTRATE: Record<BloccoEntrate, string> = {
  gia: "bg-emerald-200/70 text-emerald-900 dark:bg-emerald-900/50 dark:text-emerald-200",
  entranti: "bg-emerald-100/80 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300",
  sollecitare: "bg-amber-100/90 text-amber-800 dark:bg-amber-900/30 dark:text-amber-300",
  standby: "bg-slate-200/80 text-slate-700 dark:bg-slate-700/60 dark:text-slate-200",
  standby2: "bg-slate-300/80 text-slate-800 dark:bg-slate-600/70 dark:text-slate-100",
};
const TESTA_USCITE: Record<SecchioCassa, string> = {
  ora: "bg-rose-100/90 text-rose-800 dark:bg-rose-900/40 dark:text-rose-200",
  prossimi: "bg-amber-100/90 text-amber-800 dark:bg-amber-900/30 dark:text-amber-300",
  standby: "bg-slate-200/80 text-slate-700 dark:bg-slate-700/60 dark:text-slate-200",
  standby2: "bg-slate-300/80 text-slate-800 dark:bg-slate-600/70 dark:text-slate-100",
};
/** Tutti i blocchi del foglio, per "chiudi tutti". */
const TUTTI_I_BLOCCHI = [
  "e:gia",
  "e:entranti",
  "e:sollecitare",
  "e:standby",
  "e:standby2",
  "u:ora",
  "u:prossimi",
  "u:standby",
  "u:standby2",
] as const;
/** Vista TV: larghezza a cui si disegna la cassa prima di portarla a
 *  quella dello schermo. */
const LARGHEZZA_TV = 1400;
/** Sotto questa scala il foglio non si rimpicciolisce più (non si
 *  leggerebbe): torna la barra di scorrimento. */
const ZOOM_MINIMO = 0.75;
/** Dove il browser ricorda i blocchi richiusi (preferenza di chi guarda). */
const CHIAVE_BLOCCHI = "dr.cassa.blocchiChiusi";
const CHIAVE_BLOCCHI_V = "dr.cassa.blocchiChiusi.v";

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
/** Righe scritte a mano (non calcolate): tinta viola e pillola "manuale". */
const TINTA_MAN = "bg-violet-50/80 dark:bg-violet-950/30";
const PILL_MAN = "bg-violet-600 text-white";
const PILL_COMP = "bg-rose-600 text-white";
const PILL_STIMA = "bg-amber-500 text-white";
const PILL_AUTO =
  "border border-rose-300 bg-rose-50 text-rose-700 dark:border-rose-700 dark:bg-rose-950/40 dark:text-rose-300";
const eManuale = (c: Cella) => c.tipo === "manuale" || c.tipo === "amex";
/** Tinta di una riga delle entrate: verde; ardesia per lo Standby (il resto
 *  degli accordi, lo incasseremo più avanti); ardesia più scura per lo
 *  Standby 2 (messe da parte a mano, non le incasseremo). */
const tintaEntrata = (e?: Cella) =>
  !e?.parcheggiata ? TINTA.entrate.col : e.daAccordo ? TINTA.standby.col : TINTA.standby2.col;
const cel = "border-b border-r border-border/60 px-3 py-1.5 align-middle last:border-r-0";
const celN = `${cel} text-right tabular-nums whitespace-nowrap text-[13.5px]`;
const celH = "px-3 py-2 text-left text-[12px] font-semibold uppercase tracking-wide";

export function Foglio({
  t,
  entrate,
  uscite,
  entrateTutte,
  usciteTutte,
  filtri,
  onFiltri,
  limite,
  readOnly,
  altezzaMax,
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
  onAccordo,
  accordi,
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
  limite: string;
  readOnly: boolean;
  /** Altezza massima della cornice in px (vista TV); senza, il 78% dello schermo. */
  altezzaMax?: number;
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
  /** Accordo sugli entranti: quanto entra di quel cliente (null = lo toglie). */
  onAccordo: (c: Cella, importo: number | null) => void;
  /** Gli accordi salvati, per controparte. */
  accordi?: Readonly<Record<string, number>>;
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
  // Blocchi richiusi col "−" (Simone 11/11/2026): resta la testata col suo
  // totale. È una preferenza di chi guarda, la ricorda il browser.
  const [chiusi, setChiusi] = useState<Set<string>>(new Set());
  useEffect(() => {
    try {
      const v: unknown = JSON.parse(window.localStorage.getItem(CHIAVE_BLOCCHI) ?? "[]");
      if (Array.isArray(v)) {
        let lista = v.filter((x): x is string => typeof x === "string");
        // 1.115.0: lo standby delle entrate di prima ora si chiama Standby 2.
        // Chi lo teneva chiuso lo ritrova chiuso (una volta sola).
        if (window.localStorage.getItem(CHIAVE_BLOCCHI_V) !== "2") {
          lista = lista.map((x) => (x === "e:standby" ? "e:standby2" : x));
          window.localStorage.setItem(CHIAVE_BLOCCHI, JSON.stringify(lista));
          window.localStorage.setItem(CHIAVE_BLOCCHI_V, "2");
        }
        setChiusi(new Set(lista));
      }
    } catch {
      // preferenza illeggibile: tutto aperto
    }
  }, []);
  const salvaChiusi = (next: Set<string>) => {
    setChiusi(next);
    try {
      window.localStorage.setItem(CHIAVE_BLOCCHI, JSON.stringify([...next]));
    } catch {
      // senza memoria del browser vale solo per questa visita
    }
  };
  const toggleBlocco = (id: string) => {
    const next = new Set(chiusi);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    salvaChiusi(next);
  };
  // Con un filtro acceso i blocchi di quel lato si aprono: quello che si
  // cerca non deve restare nascosto dentro un blocco richiuso.
  const VUOTO: ReadonlySet<string> = new Set();
  const slotsE = slotsEntrate(entrate, t, espansi, filtroEntrateAttivo(filtri) ? VUOTO : chiusi);
  const slotsU = slotsUscite(
    uscite,
    t,
    limite,
    filtroUsciteAttivo(filtri) ? VUOTO : chiusi,
    espansi,
  );
  const n = Math.max(slotsE.length, slotsU.length);
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
  // ADATTA ALLA LARGHEZZA (Simone 12/11/2026: "me la taglia sempre a
  // destra"). Se il foglio è più largo della sua cornice lo si rimpicciolisce
  // quanto basta per farcelo stare tutto; sotto il 75% non si leggerebbe più
  // e allora resta la barra di scorrimento. Si misura a grandezza naturale a
  // ogni disegno e quando cambia la larghezza della cornice.
  const tabellaRef = useRef<HTMLTableElement>(null);
  const adatta = () => {
    const el = corniceRef.current;
    const tb = tabellaRef.current;
    if (!el || !tb) return;
    tb.style.removeProperty("zoom");
    const spazio =
      el.getBoundingClientRect().width * (el.offsetWidth ? el.clientWidth / el.offsetWidth : 1);
    const naturale = tb.getBoundingClientRect().width;
    if (naturale > spazio + 1)
      tb.style.setProperty(
        "zoom",
        String(Math.max(ZOOM_MINIMO, Math.floor((spazio / naturale) * 1000) / 1000)),
      );
  };
  useLayoutEffect(adatta);
  useEffect(() => {
    const el = corniceRef.current;
    const ro = el && typeof ResizeObserver !== "undefined" ? new ResizeObserver(adatta) : null;
    if (el && ro) ro.observe(el);
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
        onApriTutti={() => salvaChiusi(new Set())}
        onChiudiTutti={() => salvaChiusi(new Set(TUTTI_I_BLOCCHI))}
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
        className={altezzaMax ? "overflow-auto" : "max-h-[78vh] overflow-auto"}
        style={altezzaMax ? { maxHeight: altezzaMax } : undefined}
      >
        <table ref={tabellaRef} className="w-full min-w-[1332px] border-collapse text-[13px]">
          <colgroup>
            <col className="w-[120px]" />
            <col className="w-[150px]" />
            <col className="w-[250px]" />
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
                colSpan={3}
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
                className={`${celH} ${TINTA.entrate.col} border-r border-border/60 text-right text-foreground`}
                title={t("cpr.accordoTip")}
              >
                {t("cpr.colAccordo")}
              </th>
              <th
                className={`${celH} ${TINTA.entrate.col} border-r border-border/60 text-foreground`}
              >
                {t("cpr.colNote")}
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
            {righe.map((i) => {
              const slot = slotsE[i];
              const e = slot?.kind === "riga" || slot?.kind === "padre" ? slot.c : undefined;
              const padre = slot?.kind === "padre" ? slot : undefined;
              const figlio = slot?.kind === "riga" && slot.figlio === true;
              const slotU = slotsU[i];
              const testaU = slotU?.kind === "head" ? slotU : undefined;
              const u = slotU?.kind === "riga" ? slotU.c : undefined;
              const voceU = slotU?.kind === "voce" ? slotU : undefined;
              const uAperta = slotU?.kind === "riga" && slotU.aperta === true;
              return (
                <tr key={i} className="group/r odd:bg-background/40 hover:bg-primary/5">
                  {slot?.kind === "head" ? (
                    <td colSpan={3} className={`${cel} ${TESTA_ENTRATE[slot.id]}`}>
                      <span className="flex items-center gap-2">
                        <PiuMeno
                          aperto={!slot.chiuso}
                          onClick={() => toggleBlocco(`e:${slot.id}`)}
                          t={t}
                        />
                        <span className="text-[11px] font-semibold uppercase tracking-wider">
                          {slot.label}
                        </span>
                        <span className="text-[10.5px] font-normal opacity-70">
                          {slot.n} {t("cpr.fVoci")}
                        </span>
                        <span className="ml-auto text-[13.5px] font-bold tabular-nums">
                          {fmt(slot.value)}
                        </span>
                      </span>
                    </td>
                  ) : (
                    <>
                      <td
                        className={`${celN} ${tintaEntrata(e)} ${e && eManuale(e) ? TINTA_MAN : ""} ${e?.parcheggiata && !e.daAccordo ? "italic text-slate-500 line-through decoration-slate-400" : e?.parcheggiata ? "italic text-slate-600 dark:text-slate-300" : e?.attesa ? "italic text-muted-foreground" : `${TINTA.entrate.num} font-medium`}`}
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
                      {/* "Entra ora": l'accordo sugli entranti, per cliente */}
                      <td className={`${celN} ${tintaEntrata(e)} px-1.5 py-1`}>
                        {e &&
                          !figlio &&
                          e.tipo === "attiva" &&
                          e.chiaveControparte &&
                          (e.accordo || e.daAccordo || (e.attesa && !e.parcheggiata)) && (
                            <CampoAccordo
                              c={e}
                              valore={accordi?.[e.chiaveControparte]}
                              readOnly={readOnly}
                              onAccordo={onAccordo}
                              t={t}
                            />
                          )}
                      </td>
                      <td
                        className={`${cel} ${tintaEntrata(e)} ${e && eManuale(e) ? TINTA_MAN : ""} ${e?.attesa ? "text-muted-foreground" : "text-foreground"}`}
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
                            {(e.accordo || e.daAccordo) && (
                              <span
                                className={`inline-flex shrink-0 items-center rounded px-1 text-[9px] font-bold uppercase ${e.accordo ? "bg-sky-600 text-white" : "border border-sky-500 text-sky-700 dark:text-sky-300"}`}
                                title={
                                  e.accordo
                                    ? e.dettaglio
                                    : e.tipo === "compensazione"
                                      ? t("cpr.compAccordoTip")
                                      : t("cpr.daAccordoTip")
                                }
                              >
                                {t("cpr.accordo")}
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
                            {!padre && e.attesa && !e.accordo && !readOnly && (
                              <ToggleStandby c={e} onParcheggia={onParcheggia} t={t} />
                            )}
                            {!padre && !e.accordo && !readOnly && (
                              <Togli onClick={() => onTogli(e)} t={t} />
                            )}
                          </span>
                        )}
                      </td>
                    </>
                  )}
                  {testaU ? (
                    <>
                      {COLONNE_USCITE.map((col) => (
                        <td
                          key={col}
                          className={`${celN} ${TESTA_USCITE[testaU.col]} text-[13.5px] font-bold`}
                        >
                          {col === testaU.col ? fmt(testaU.value) : ""}
                        </td>
                      ))}
                      <td colSpan={2} className={`${cel} ${TESTA_USCITE[testaU.col]}`}>
                        <span className="flex items-center gap-2">
                          <PiuMeno
                            aperto={!testaU.chiuso}
                            onClick={() => toggleBlocco(`u:${testaU.col}`)}
                            t={t}
                          />
                          <span className="text-[11px] font-semibold uppercase tracking-wider">
                            {testaU.label}
                          </span>
                          <span className="text-[10.5px] font-normal opacity-70">
                            {testaU.n} {t("cpr.fVoci")}
                          </span>
                        </span>
                      </td>
                    </>
                  ) : voceU ? (
                    <>
                      {COLONNE_USCITE.map((col) => (
                        <td
                          key={col}
                          className={`${celN} ${TINTA[col].col} text-[12.5px] italic text-muted-foreground`}
                        >
                          {col === voceU.padre.secchio ? fmt(voceU.v.importo) : ""}
                        </td>
                      ))}
                      <td className={`${cel} py-1 pl-9 text-[12.5px] text-muted-foreground`}>
                        {voceU.v.nome}
                      </td>
                      <td className={`${cel} py-1 text-[11px] text-muted-foreground`}>
                        {voceU.v.extra ?? ""}
                      </td>
                    </>
                  ) : (
                    <>
                      {cellaUscita(u, "ora")}
                      {cellaUscita(u, "prossimi")}
                      {cellaUscita(u, "standby")}
                      {cellaUscita(u, "standby2")}
                      <td className={`${cel} text-foreground ${u && eManuale(u) ? TINTA_MAN : ""}`}>
                        {u && (
                          <span className="flex items-center gap-1.5">
                            {haVoci(u) && (
                              <button
                                type="button"
                                onClick={() => toggleCliente(u.chiave)}
                                title={t(uAperta ? "cpr.chiudiCliente" : "cpr.apriVoci")}
                                aria-expanded={uAperta}
                                className="inline-flex h-4 w-4 shrink-0 items-center justify-center rounded border border-border text-[11px] leading-none text-muted-foreground hover:bg-muted hover:text-foreground"
                              >
                                {uAperta ? "−" : "+"}
                              </button>
                            )}
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
                            {(u.tipo === "stipendi" ||
                              u.tipo === "fisso" ||
                              u.tipo === "altre" ||
                              u.tipo === "fiscale") && (
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
                    </>
                  )}
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
                    <td className={`${cel} ${TINTA.entrate.col}`} />
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
                    <td className={`${cel} ${TINTA.entrate.col}`} />
                    <td className={`${cel} ${TINTA.entrate.col} py-1`}>
                      <Aggiungi onClick={() => apri(true)} t={t} />
                    </td>
                  </>
                )}
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
  onApriTutti,
  onChiudiTutti,
}: {
  t: TFn;
  filtri: FiltriFoglio;
  onFiltri: (f: FiltriFoglio) => void;
  entrateTutte: readonly Cella[];
  usciteTutte: readonly Cella[];
  limite: string;
  /** Apre / richiude tutti i blocchi del foglio in un colpo. */
  onApriTutti: () => void;
  onChiudiTutti: () => void;
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
            <option value="standby2">{t("cpr.colStandby2")}</option>
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
        <span className="ml-auto flex items-center gap-1.5 text-[12px] text-muted-foreground">
          {t("cpr.bBlocchi")}
          <button
            type="button"
            onClick={onApriTutti}
            title={t("cpr.bApriTuttiTip")}
            className="inline-flex h-7 items-center gap-1 rounded-md border border-border bg-background px-2 font-semibold text-foreground hover:bg-muted"
          >
            + {t("cpr.bApriTutti")}
          </button>
          <button
            type="button"
            onClick={onChiudiTutti}
            title={t("cpr.bChiudiTuttiTip")}
            className="inline-flex h-7 items-center gap-1 rounded-md border border-border bg-background px-2 font-semibold text-foreground hover:bg-muted"
          >
            − {t("cpr.bChiudiTutti")}
          </button>
        </span>
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

/** Pallino colorato di ogni sezione, come le colonne del foglio. */
const PUNTO_SEZIONE: Record<SezioneCassaId, string> = {
  uscite_ora: "bg-rose-600",
  uscite_prossimi: "bg-amber-500",
  entrati: "bg-emerald-700",
  entranti: "bg-emerald-500",
  standby_entrate: "bg-slate-400",
  standby2_entrate: "bg-slate-600",
  standby_uscite: "bg-slate-500",
  standby2_uscite: "bg-slate-700",
};

/** "Stampa / esporta voci" (Simone 11/11/2026): le sette sezioni del foglio,
 *  ognuna col numero di voci e il totale, da esportare in CSV o da stampare
 *  (dal riquadro di stampa si salva in PDF); in fondo, tutte insieme. Con i
 *  filtri accesi escono solo le voci filtrate, e la stampa lo scrive. Le
 *  uscite escono sempre fornitore per fornitore, anche con la vista
 *  raggruppata per tipologia o cliente. */
function MenuVoci({
  t,
  entrate,
  uscite,
  limite,
  data,
  filtrate,
}: {
  t: TFn;
  entrate: readonly Cella[];
  uscite: readonly Cella[];
  limite: string;
  data: string;
  filtrate: boolean;
}) {
  const [aperto, setAperto] = useState(false);
  const titoli: Record<SezioneCassaId, string> = {
    uscite_ora: t("cpr.vUsciteOra"),
    uscite_prossimi: `${t("cpr.colUscite")} ${etichettaProssimi(limite)}`,
    entrati: t("cpr.vEntrati"),
    entranti: t("cpr.entranti"),
    standby_entrate: t("cpr.vStandbyEntrate"),
    standby2_entrate: t("cpr.vStandby2Entrate"),
    standby_uscite: t("cpr.vStandbyUscite"),
    standby2_uscite: t("cpr.vStandby2Uscite"),
  };
  const etichette = (): EtichetteStampa => ({
    titolo: t("cpr.vTitolo"),
    sottotitolo: `DR Logistica · ${t("cpr.vSituazione")} ${fmtData(data)} · ${t("cpr.vStampato")} ${new Date().toLocaleString(
      "it-IT",
      { dateStyle: "short", timeStyle: "short" },
    )}${filtrate ? ` · ${t("cpr.vFiltrate")}` : ""}`,
    sezioni: titoli,
    colSezione: t("cpr.vColSezione"),
    colVoce: t("cpr.vColVoce"),
    colDettaglio: t("cpr.vColDettaglio"),
    colData: t("cpr.vColData"),
    colImporto: t("cpr.vColImporto"),
    colTipologia: t("cpr.vColTipologia"),
    colCliente: t("cpr.vColCliente"),
    colPriorita: t("cpr.fPriorita"),
    colSollecitare: t("cpr.daSollecitare"),
    colNota: t("cpr.vColNota"),
    totale: t("cpr.vTotale"),
    voci: t("cpr.fVoci"),
    si: t("cpr.vSi"),
    nessuna: t("cpr.vNessuna"),
    priorita: { h: t("cpr.prioH"), m: t("cpr.prioM"), l: t("cpr.prioL") },
    logo: `${window.location.origin}/favicon.png`,
  });
  const sezioni: SezioneCassa<Cella>[] = aperto ? sezioniCassa(entrate, uscite) : [];
  const stampa = (lista: readonly SezioneCassa[]) => {
    if (!apriStampa(htmlStampaSezioni(lista, etichette()))) toast.error(t("cpr.vPopup"));
    setAperto(false);
  };
  const csv = (lista: readonly SezioneCassa[], nome: string) => {
    const { testata, righe } = csvSezioni(lista, etichette());
    esportaCsvFile(`cassa-proprieta-${nome}-${data}`, testata, righe);
    setAperto(false);
  };
  const icona =
    "inline-flex h-7 w-7 items-center justify-center rounded-md border border-border text-muted-foreground hover:border-primary hover:bg-primary hover:text-primary-foreground disabled:pointer-events-none disabled:opacity-30";
  const conVoci = sezioni.filter((s) => s.voci.length > 0);
  return (
    <span className="no-print relative">
      <button
        type="button"
        onClick={() => setAperto((v) => !v)}
        aria-expanded={aperto}
        className={`inline-flex items-center gap-2 rounded-lg border px-3 py-1.5 text-sm transition-colors ${aperto ? "border-primary bg-primary text-primary-foreground" : "border-border bg-background text-foreground hover:bg-secondary"}`}
      >
        <Printer className="h-4 w-4" /> {t("cpr.vBottone")}
      </button>
      {aperto && (
        <>
          {/* un clic fuori dal riquadro lo chiude */}
          <div className="fixed inset-0 z-30" onClick={() => setAperto(false)} />
          <div className="absolute right-0 top-full z-40 mt-1.5 w-[460px] max-w-[92vw] rounded-xl border border-border bg-card p-2 text-left shadow-xl">
            <div className="px-2 pb-1.5 pt-1 text-[11px] text-muted-foreground">
              {t("cpr.vIntro")}
              {filtrate && (
                <span className="ml-1 rounded bg-amber-500/15 px-1.5 py-0.5 font-semibold text-amber-800 dark:text-amber-300">
                  {t("cpr.vFiltrate")}
                </span>
              )}
            </div>
            {sezioni.map((s) => (
              <div
                key={s.id}
                className="flex items-center gap-2 rounded-lg px-2 py-1.5 hover:bg-muted/60"
              >
                <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${PUNTO_SEZIONE[s.id]}`} />
                <span className="min-w-0 flex-1 truncate text-[13px] font-medium text-foreground">
                  {titoli[s.id]}
                </span>
                <span className="shrink-0 text-[11.5px] tabular-nums text-muted-foreground">
                  {s.voci.length} {t("cpr.fVoci")} ·{" "}
                  <b className="text-foreground">{fmt(s.totale)}</b>
                </span>
                <button
                  type="button"
                  disabled={!s.voci.length}
                  onClick={() => csv([s], s.id.replace(/_/g, "-"))}
                  title={t("cpr.vCsvTip")}
                  className={icona}
                >
                  <Download className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  disabled={!s.voci.length}
                  onClick={() => stampa([s])}
                  title={t("cpr.vStampaTip")}
                  className={icona}
                >
                  <Printer className="h-3.5 w-3.5" />
                </button>
              </div>
            ))}
            <div className="mt-1 flex items-center gap-2 border-t border-border px-2 pb-0.5 pt-2">
              <span className="flex-1 text-[13px] font-semibold text-foreground">
                {t("cpr.vTutte")}
              </span>
              <button
                type="button"
                disabled={!conVoci.length}
                onClick={() => csv(conVoci, "voci")}
                title={t("cpr.vCsvTip")}
                className={icona}
              >
                <Download className="h-3.5 w-3.5" />
              </button>
              <button
                type="button"
                disabled={!conVoci.length}
                onClick={() => stampa(conVoci)}
                title={t("cpr.vStampaTip")}
                className={icona}
              >
                <Printer className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>
        </>
      )}
    </span>
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

/** "Entra ora": l'accordo sugli entranti di un cliente (Simone 12/11/2026).
 *  Vuoto = entra tutto; un importo = entra quello e il resto delle sue
 *  fatture attese va in standby; 0 = non entra nulla, tutto in standby. */
function CampoAccordo({
  c,
  valore,
  readOnly,
  onAccordo,
  t,
}: {
  c: Cella;
  valore: number | undefined;
  readOnly: boolean;
  onAccordo: (c: Cella, importo: number | null) => void;
  t: TFn;
}) {
  const [testo, setTesto] = useState<string | null>(null);
  if (readOnly)
    return valore != null ? (
      <span className="text-[12px] tabular-nums text-sky-700 dark:text-sky-300">{fmt(valore)}</span>
    ) : null;
  return (
    <input
      value={testo ?? (valore != null ? fmt(valore) : "")}
      onChange={(e) => setTesto(e.target.value)}
      onBlur={() => {
        if (testo == null) return;
        const pulito = testo.trim();
        setTesto(null);
        if (!pulito) {
          if (valore != null) onAccordo(c, null);
          return;
        }
        const n = parseNum(pulito);
        if (n != null && n >= 0 && (valore == null || Math.abs(n - valore) > 0.005))
          onAccordo(c, n);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter") (e.target as HTMLInputElement).blur();
        if (e.key === "Escape") setTesto(null);
      }}
      placeholder={t("cpr.accordoTutto")}
      title={t("cpr.accordoTip")}
      className={`h-8 w-full min-w-[124px] rounded-md border bg-background px-2 text-right text-[14px] tabular-nums focus:outline-none focus:ring-2 focus:ring-primary/40 ${valore != null ? "border-sky-500 font-semibold text-sky-700 dark:text-sky-300" : "border-border/70 text-foreground placeholder:text-muted-foreground/50"}`}
    />
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

/** Entrata attesa: in "Standby 2" (messa da parte a mano: non la
 *  incasseremo, non conta) o di nuovo tra le entranti. Sulle righe dello
 *  Standby (il resto di un accordo) il pulsante manda nello Standby 2. */
function ToggleStandby({
  c,
  onParcheggia,
  t,
}: {
  c: Cella;
  onParcheggia: (c: Cella, on: boolean) => void;
  t: TFn;
}) {
  const inStandby2 = c.parcheggiata && !c.daAccordo;
  return (
    <button
      type="button"
      onClick={() => onParcheggia(c, !inStandby2)}
      title={inStandby2 ? t("cpr.riattivaEntrata") : t("cpr.parcheggiaEntrata")}
      className={`shrink-0 rounded border px-1.5 text-[9px] font-bold uppercase ${inStandby2 ? "border-emerald-500 bg-emerald-500/10 text-emerald-700 hover:bg-emerald-500/20 dark:text-emerald-300" : "border-slate-400 bg-slate-500/10 text-slate-700 hover:bg-slate-500/20 dark:text-slate-200"}`}
    >
      {inStandby2 ? `↩ ${t("cpr.entranti")}` : `⏸ ${t("cpr.colStandby2")}`}
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
