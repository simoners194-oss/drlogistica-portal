// DR Portal — Finanza (sezione riservata al direttore DR005 + admin).
// Estratto conto bancario: import da xlsx (con scelta del foglio), archivio
// movimenti classificati, overview incassi/spese per cliente, anomalie da
// sanare a mano, storico degli import con annullamento.
import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import {
  Lock,
  Upload,
  Table2,
  TrendingUp,
  AlertTriangle,
  Loader2,
  Download,
  CheckCircle2,
  History,
  Trash2,
  Pencil,
  Users,
  Filter,
  ArrowDown,
  ArrowUp,
  GraduationCap,
  Wand2,
  Receipt,
  ReceiptText,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  RefreshCw,
} from "lucide-react";
import { FattureTab } from "@/components/FattureTab";
import { csvData, csvPeriodo, esportaCsvFile } from "@/lib/csv";
import { MultiSelect } from "@/components/MultiSelect";
import { CampoVocabolario } from "@/components/CampoVocabolario";
import { ResocontoTab } from "@/components/ResocontoTab";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import { isSupervisoreGlobale } from "@/lib/richieste-logic";
import {
  parseEstratto,
  normalizeTesto,
  parseMatrice,
  clienteGroupKey,
  LEGACY_IMPORT_ID,
  TIPOLOGIE_MOVIMENTO,
  matchRegola,
  applicaRegole,
  matchDipendenteNome,
  type DipendenteRoster,
  type MovimentoParsed,
  type ParseFileResult,
  type RegolaFinanza,
} from "@/lib/finanza-logic";
import {
  spGetMovimenti,
  spGetMovimentiChiavi,
  spImportMovimenti,
  spUpdateMovimento,
  spGetImportStorico,
  spGetDettagliDistinte,
  spSetDistintaAppalto,
  spSetDistintaMovimento,
  spGetRosterDipendenti,
  spImportDistinta,
  spAnnullaImport,
  spEliminaMovimento,
  spCorreggiMovimento,
  spGetRegoleFinanza,
  spGetTerminiPagamento,
  spImportTermini,
  spDeleteTermine,
  spCopiaTerminiSuFornitori,
  spGetRegoleFatture,
  spCreateRegolaFattura,
  spDeleteRegolaFattura,
  spCreateRegolaFinanza,
  spDeleteRegolaFinanza,
  spApplicaRegolaFinanza,
  spUpdateRegolaFinanza,
  spApplicaRegolaDipendenti,
  spAssegnaContoLotto,
  spAnnullaRegolaFinanza,
  spEbStato,
  spEbSaldo,
  spEbSincronizza,
} from "@/lib/sharepoint.functions";
import type {
  SpMovimento,
  ImportStoricoRiga,
  DettaglioDistinta,
  EbStato,
  EbSyncResult,
  EbSaldoInfo,
} from "@/lib/sharepoint.server";

export const Route = createFileRoute("/finanza")({
  head: () => ({ meta: [{ title: "Finanze — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: FinanzaPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

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
const MESI_EN = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}
function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
// Numero per il CSV: virgola decimale (Excel italiano), niente separatore
// migliaia (che Excel leggerebbe come testo).
function csvNum(n: number): string {
  return (Math.round(n * 100) / 100).toString().replace(".", ",");
}
// "IMP-2026-07-22T15:30:12" → "22/07/2026 15:30" (gruppo legacy → etichetta;
// i lotti "SYNC-…" arrivano dal collegamento banca e sono marcati "API").
// Stralcio della descrizione attorno al testo cercato: fa capire PERCHÉ una
// riga senza cliente corrisponde alla ricerca (la descrizione non è in
// tabella — "romano" pescava le utenze di Fiano Romano e sembrava un bug).
function stralcioDescr(descrizione: string, cerca: string): string {
  const q = cerca.trim().toLowerCase();
  const i = descrizione.toLowerCase().indexOf(q);
  if (i < 0) return descrizione.slice(0, 40);
  const da = Math.max(0, i - 15);
  const a = Math.min(descrizione.length, i + q.length + 25);
  return `${da > 0 ? "…" : ""}${descrizione.slice(da, a)}${a < descrizione.length ? "…" : ""}`;
}

// Il campo Pattern di una regola regge 240 caratteri (margine sotto il
// limite di 255 delle colonne testo SharePoint): un elenco piu' lungo si
// SPEZZA in piu' regole gemelle — il match multi-termine e' un OR, quindi
// il comportamento e' identico. I termini doppi spariscono.
// Tasto laterale: porta in FONDO alla pagina con un colpo; arrivati in
// fondo si capovolge e riporta in cima.
function ScorriFondo() {
  const [inFondo, setInFondo] = useState(false);
  useEffect(() => {
    const controlla = () =>
      setInFondo(window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 80);
    controlla();
    window.addEventListener("scroll", controlla, { passive: true });
    window.addEventListener("resize", controlla);
    return () => {
      window.removeEventListener("scroll", controlla);
      window.removeEventListener("resize", controlla);
    };
  }, []);
  return (
    <button
      type="button"
      onClick={() =>
        window.scrollTo({
          top: inFondo ? 0 : document.documentElement.scrollHeight,
          behavior: "smooth",
        })
      }
      className="fixed bottom-6 right-4 z-40 rounded-full border border-border bg-card p-2.5 text-muted-foreground shadow-[var(--shadow-elegant)] hover:text-foreground"
    >
      {inFondo ? <ArrowUp className="h-4 w-4" /> : <ArrowDown className="h-4 w-4" />}
    </button>
  );
}

const MAX_PATTERN = 240;
function spezzaPattern(testo: string): string[] {
  const visti = new Set<string>();
  const termini: string[] = [];
  for (const parte of testo.split(/[,;\n]/)) {
    const termine = parte.trim();
    if (!termine) continue;
    const k = termine.toLowerCase();
    if (!visti.has(k)) {
      visti.add(k);
      termini.push(termine);
    }
  }
  const blocchi: string[] = [];
  let corrente = "";
  for (const termine of termini) {
    const candidato = corrente ? `${corrente}, ${termine}` : termine;
    if (candidato.length > MAX_PATTERN && corrente) {
      blocchi.push(corrente);
      corrente = termine;
    } else {
      corrente = candidato;
    }
  }
  if (corrente) blocchi.push(corrente);
  return blocchi.length ? blocchi : [testo.trim().slice(0, MAX_PATTERN)];
}

function fmtImportId(id: string, legacyLabel: string): string {
  if (!id) return legacyLabel;
  const m = id.match(/^(IMP|SYNC)-(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2})/);
  if (!m) return id;
  const base = `${m[4]}/${m[3]}/${m[2]} ${m[5]}`;
  return m[1] === "SYNC" ? `${base} · API` : base;
}

// Blocchi di upload verso il server (sotto il limite server di 150).
const CHUNK = 100;
// Righe per pagina nella tabella movimenti.
const RIGHE_PAGINA = 50;

type Tab =
  "movimenti" | "overview" | "resoconto" | "attive" | "passive" | "anomalie" | "storico" | "regole";

interface SheetInfo {
  name: string;
  res: ParseFileResult | null; // null = foglio non riconosciuto
}
interface SheetChoice {
  fileName: string;
  sheets: SheetInfo[];
  selected: string;
}

interface CorrezioneX100 {
  id: string;
  dataContabile: string;
  da: number;
  a: number;
  chiave: string;
  descrizione: string;
}

interface PreviewImport {
  fileName: string;
  righe: MovimentoParsed[];
  nuove: MovimentoParsed[];
  doppioni: number;
  scartate: number;
  anomalie: number;
  dal: string;
  al: string;
  /** Righe d'archivio corrotte (x100) da correggere IN POSTO col valore
   *  vero del file: importo e chiave cambiano, il resto resta. */
  correzioni: CorrezioneX100[];
}

function FinanzaPage() {
  const { t, lang } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [tab, setTab] = useState<Tab>("movimenti");
  // Anni selezionati (vuoto = tutti). Con UN solo anno l'overview mostra le
  // colonne-mese, altrimenti le colonne-anno.
  const [anni, setAnni] = useState<number[]>([new Date().getFullYear()]);

  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [anomalie, setAnomalie] = useState<SpMovimento[] | null>(null);
  const [storico, setStorico] = useState<ImportStoricoRiga[] | null>(null);
  // Distinte / esiti pagamenti: dettaglio dei pagamenti cumulativi.
  const [distinte, setDistinte] = useState<DettaglioDistinta[] | null>(null);
  const [trancheBusy, setTrancheBusy] = useState(false);
  const [rosterDip, setRosterDip] = useState<DipendenteRoster[] | null>(null);
  const [distPreview, setDistPreview] = useState<
    | {
        idPagamento: string;
        dataEsecuzione: string;
        beneficiario: string;
        importo: number;
        tipoPagamento: string;
        descrizione: string;
      }[]
    | null
  >(null);
  const [distBusy, setDistBusy] = useState(false);
  const [distModal, setDistModal] = useState<{
    data: string;
    tipo: string;
    somma: number;
    righe: DettaglioDistinta[];
  } | null>(null);

  // Filtri archivio movimenti
  const [tipiF, setTipiF] = useState<string[]>([]);
  const [cercaF, setCercaF] = useState("");
  const [mesiF, setMesiF] = useState<number[]>([]); // vuoto = tutti
  // Clienti selezionati nel menu a tendina (vuoto = tutti). Le voci proposte
  // sono le 15 controparti con piu' incassi, in ordine decrescente.
  const [clientiF, setClientiF] = useState<string[]>([]);
  const [paginaMov, setPaginaMov] = useState(1); // pagine da righePagina
  const [righePagina, setRighePagina] = useState(RIGHE_PAGINA);

  // Overview: incassi o spese (+ filtro tipologia, utile solo per le spese)
  const [ovMode, setOvMode] = useState<"incassi" | "spese" | "regole" | "appalti">("incassi");
  const [ovTipF, setOvTipF] = useState("tutte");
  const [sottF, setSottF] = useState<string[]>([]);
  const [allocPriF, setAllocPriF] = useState<string[]>([]);
  const [allocSecF, setAllocSecF] = useState<string[]>([]);

  // Import estratto conto (pannello a scomparsa dentro la tab Movimenti:
  // "Importa" da solo era ambiguo dopo l'arrivo delle Fatture).
  const [showImportEC, setShowImportEC] = useState(false);
  const [sheetChoice, setSheetChoice] = useState<SheetChoice | null>(null);
  const [preview, setPreview] = useState<PreviewImport | null>(null);
  const [parsing, setParsing] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState("");

  // Collegamento banca (PSD2): la CONFIGURAZIONE vive in Amministrazione
  // (BancaPsd2Panel). Qui restano il saldo e il tasto Aggiorna, che con il
  // collegamento attivo scarica anche i nuovi movimenti dalla banca.
  const [ebStato, setEbStato] = useState<EbStato | null>(null);
  const [ebSyncBusy, setEbSyncBusy] = useState(false);
  const [ebProgress, setEbProgress] = useState("");
  // Saldo attuale dalla banca (null = collegamento non attivo o non caricato).
  const [ebSaldoInfo, setEbSaldoInfo] = useState<EbSaldoInfo | null>(null);

  // Storico: annullamento in corso
  const [annullaBusy, setAnnullaBusy] = useState<string | null>(null);
  const [annullaProgress, setAnnullaProgress] = useState(0);

  // Regole apprese
  const [regole, setRegole] = useState<RegolaFinanza[] | null>(null);
  const [rPattern, setRPattern] = useState("");
  const [rCampo, setRCampo] = useState<"cliente" | "descrizione" | "entrambi">("cliente");
  // Regola in modifica: si precompila il form e al salvataggio la vecchia
  // viene sostituita (elimina + ricrea).
  const [rEditId, setREditId] = useState<string | null>(null);
  // Elenco regole raggruppato per tipologia: si apre un gruppo al tocco.
  const [catAperta, setCatAperta] = useState<string | null>(null);
  const [uniBusy, setUniBusy] = useState(false);
  const [bonificaBusy, setBonificaBusy] = useState(false);
  const [rCerca, setRCerca] = useState("");
  // Multi-selezione: spunta le righe e agisci in blocco.
  const [soloDistinte, setSoloDistinte] = useState(false);
  const [selMov, setSelMov] = useState<Set<string>>(new Set());
  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkTip, setBulkTip] = useState("");
  const [bulkSott, setBulkSott] = useState("");
  const [bulkPri, setBulkPri] = useState("");
  const [bulkSec, setBulkSec] = useState("");
  const [bulkBusy, setBulkBusy] = useState(false);
  const [selReg, setSelReg] = useState<Set<string>>(new Set());
  const [dataDaF, setDataDaF] = useState("");
  const [dataAF, setDataAF] = useState("");
  const [impMinF, setImpMinF] = useState("");
  const [impMaxF, setImpMaxF] = useState("");
  // UNIFICA DOPPIE: regole con gli stessi esiti (tipologia, sottocategoria,
  // allocazioni, controparte) e stesso campo/modo diventano UNA regola con
  // l'unione dei termini; i termini ripetuti spariscono anche dalle singole.
  // La semantica non cambia: il match multi-termine e' un OR.
  const unificaDoppie = async () => {
    const rs = regole ?? [];
    if (!rs.length) return;
    const fondi = (patterns: string[]) => {
      const visti = new Set<string>();
      const out: string[] = [];
      for (const pat of patterns)
        for (const parte of pat.split(/[,;\n]/)) {
          const termine = parte.trim();
          if (!termine) continue;
          const k2 = termine.toLowerCase();
          if (!visti.has(k2)) {
            visti.add(k2);
            out.push(termine);
          }
        }
      return out.join(", ");
    };
    const gruppi = new Map<string, RegolaFinanza[]>();
    for (const r of rs) {
      const k2 = [
        r.campo,
        r.modo,
        r.tipologia ?? "",
        r.sottocategoria ?? "",
        r.allocPrimaria ?? "",
        r.allocSecondaria ?? "",
        r.cliente ?? "",
        r.segno ?? "",
      ].join("|");
      gruppi.set(k2, [...(gruppi.get(k2) ?? []), r]);
    }
    const daUnire = [...gruppi.values()].filter((g) => g.length > 1);
    const inGruppo = new Set(daUnire.flat().map((r) => r.id));
    const daRipulire = rs.filter((r) => !inGruppo.has(r.id) && fondi([r.pattern]) !== r.pattern);
    const eliminate = daUnire.reduce((s2, g) => s2 + g.length - 1, 0);
    if (!daUnire.length && !daRipulire.length) {
      toast.success(t("fin.uniNiente"));
      return;
    }
    const msg = `${t("fin.uniConfirm1")} ${daUnire.length} ${t("fin.uniConfirm2")} ${eliminate} ${t("fin.uniConfirm3")} ${daRipulire.length} ${t("fin.uniConfirm4")}`;
    if (!window.confirm(msg)) return;
    setUniBusy(true);
    try {
      for (const g of daUnire) {
        const keep = g[0];
        const blocchiUni = spezzaPattern(g.map((r) => r.pattern).join(", "));
        const payload = {
          pattern: blocchiUni[0],
          campo: keep.campo,
          modo: keep.modo,
          tipologia: keep.tipologia,
          sottocategoria: keep.sottocategoria,
          allocPrimaria: keep.allocPrimaria,
          allocSecondaria: keep.allocSecondaria,
          cliente: keep.cliente,
          note: keep.note,
          segno: keep.segno,
        };
        await spUpdateRegolaFinanza({ data: { regolaId: keep.id ?? "", ...payload } });
        for (const blocco of blocchiUni.slice(1))
          await spCreateRegolaFinanza({ data: { ...payload, pattern: blocco } });
        for (const extra of g.slice(1))
          await spDeleteRegolaFinanza({ data: { regolaId: extra.id ?? "" } });
      }
      for (const r of daRipulire)
        await spUpdateRegolaFinanza({
          data: {
            regolaId: r.id ?? "",
            pattern: fondi([r.pattern]),
            campo: r.campo,
            modo: r.modo,
            tipologia: r.tipologia,
            sottocategoria: r.sottocategoria,
            allocPrimaria: r.allocPrimaria,
            allocSecondaria: r.allocSecondaria,
            cliente: r.cliente,
            note: r.note,
            segno: r.segno,
          },
        });
      const agg = (await spGetRegoleFinanza()) as RegolaFinanza[];
      setRegole(agg);
      toast.success(t("fin.uniFatto"), {
        description: `${daUnire.length + daRipulire.length} regole sistemate, ${eliminate} eliminate`,
      });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setUniBusy(false);
    }
  };
  const [dipBusy, setDipBusy] = useState(false);
  // Filtro conto sui movimenti + assegnazione conto per lotto (storico).
  const [contoF, setContoF] = useState("");
  const [contoLotto, setContoLotto] = useState<Record<string, string>>({});
  const [contoBusy, setContoBusy] = useState<string | null>(null);
  const [rModo, setRModo] = useState<"esatto" | "contiene">("esatto");
  const [rTipologia, setRTipologia] = useState("");
  const [rCliente, setRCliente] = useState("");
  const [rSottocat, setRSottocat] = useState("");
  const [rAllocPri, setRAllocPri] = useState("");
  const [rAllocSec, setRAllocSec] = useState("");
  // Vocabolario dalle regole apprese, A CASCATA: scelta la tipologia, le
  // sottocategorie si restringono a quelle usate con lei (idem le
  // allocazioni); se il filtro svuota l'elenco, si mostra tutto.
  const vocab = useMemo(() => {
    const rs = regole ?? [];
    const uniq = (xs: (string | undefined)[]) =>
      [...new Set(xs.filter(Boolean) as string[])].sort((a, b) => a.localeCompare(b));
    const perTip = rTipologia.trim()
      ? rs.filter((r) => (r.tipologia ?? "") === rTipologia.trim())
      : rs;
    const perPri = rAllocPri.trim()
      ? rs.filter((r) => (r.allocPrimaria ?? "") === rAllocPri.trim())
      : rs;
    const sotto = uniq(perTip.map((r) => r.sottocategoria));
    const sec = uniq(perPri.map((r) => r.allocSecondaria));
    return {
      tipologie: uniq(rs.map((r) => r.tipologia)),
      sottocat: sotto.length ? sotto : uniq(rs.map((r) => r.sottocategoria)),
      allocPri: uniq(rs.map((r) => r.allocPrimaria)),
      allocSec: sec.length ? sec : uniq(rs.map((r) => r.allocSecondaria)),
    };
  }, [regole, rTipologia, rAllocPri]);
  const [rNote, setRNote] = useState("");
  const [rSegno, setRSegno] = useState<"" | "entrate" | "uscite">("");
  // Movimento da cui e' partita la bacchetta: consente di applicare la
  // classificazione SOLO a lui, senza creare la regola.
  const [rSorgente, setRSorgente] = useState<string | null>(null);
  const [rApplica, setRApplica] = useState(true);
  const [rBusy, setRBusy] = useState(false);
  const [rProgress, setRProgress] = useState(0);

  // Sanatura anomalie
  const [editId, setEditId] = useState<string | null>(null);
  const [editTip, setEditTip] = useState("");
  const [editCliente, setEditCliente] = useState("");
  const [editSott, setEditSott] = useState("");
  const [editAllocPri, setEditAllocPri] = useState("");
  const [editAllocSec, setEditAllocSec] = useState("");
  const [editNrFatt, setEditNrFatt] = useState("");
  const [editNote, setEditNote] = useState("");
  const [saving, setSaving] = useState(false);
  // Vocabolario A CASCATA per il modal di correzione: stesse voci delle
  // regole apprese (scelta la tipologia restano le sottocategorie coerenti,
  // idem per le allocazioni).
  const vocabEdit = useMemo(() => {
    const rs = regole ?? [];
    const uniq = (xs: (string | undefined)[]) =>
      [...new Set(xs.filter(Boolean) as string[])].sort((a, b) => a.localeCompare(b));
    const perTip = editTip.trim() ? rs.filter((r) => (r.tipologia ?? "") === editTip.trim()) : rs;
    const perPri = editAllocPri.trim()
      ? rs.filter((r) => (r.allocPrimaria ?? "") === editAllocPri.trim())
      : rs;
    const sotto = uniq(perTip.map((r) => r.sottocategoria));
    const sec = uniq(perPri.map((r) => r.allocSecondaria));
    return {
      tipologie: uniq([...TIPOLOGIE_MOVIMENTO, ...rs.map((r) => r.tipologia)]),
      sottocat: sotto.length ? sotto : uniq(rs.map((r) => r.sottocategoria)),
      allocPri: uniq(rs.map((r) => r.allocPrimaria)),
      allocSec: sec.length ? sec : uniq(rs.map((r) => r.allocSecondaria)),
    };
  }, [regole, editTip, editAllocPri]);
  // Cascata che RICOMINCIA: cambiata la tipologia, la sottocategoria
  // scritta prima si svuota se non e' coerente con la nuova scelta
  // (idem allocazione primaria -> secondaria).
  const coerenteSott = (tip: string, sott: string) =>
    !tip.trim() ||
    !sott.trim() ||
    (regole ?? []).some(
      (r) => (r.tipologia ?? "") === tip.trim() && (r.sottocategoria ?? "") === sott.trim(),
    );
  const coerenteSec = (pri: string, sec: string) =>
    !pri.trim() ||
    !sec.trim() ||
    (regole ?? []).some(
      (r) => (r.allocPrimaria ?? "") === pri.trim() && (r.allocSecondaria ?? "") === sec.trim(),
    );
  const cambiaEditTip = (v: string) => {
    setEditTip(v);
    if (!coerenteSott(v, editSott)) setEditSott("");
  };
  const cambiaEditAllocPri = (v: string) => {
    setEditAllocPri(v);
    if (!coerenteSec(v, editAllocSec)) setEditAllocSec("");
  };
  const vocabBulk = useMemo(() => {
    const rs = regole ?? [];
    const uniq = (xs: (string | undefined)[]) =>
      [...new Set(xs.filter(Boolean) as string[])].sort((a, b) => a.localeCompare(b));
    const perTip = bulkTip.trim() ? rs.filter((r) => (r.tipologia ?? "") === bulkTip.trim()) : rs;
    const perPri = bulkPri.trim()
      ? rs.filter((r) => (r.allocPrimaria ?? "") === bulkPri.trim())
      : rs;
    const sotto = uniq(perTip.map((r) => r.sottocategoria));
    const sec = uniq(perPri.map((r) => r.allocSecondaria));
    return {
      tipologie: uniq([...TIPOLOGIE_MOVIMENTO, ...rs.map((r) => r.tipologia)]),
      sottocat: sotto.length ? sotto : uniq(rs.map((r) => r.sottocategoria)),
      allocPri: uniq(rs.map((r) => r.allocPrimaria)),
      allocSec: sec.length ? sec : uniq(rs.map((r) => r.allocSecondaria)),
    };
  }, [regole, bulkTip, bulkPri]);
  const cambiaBulkTip = (v: string) => {
    setBulkTip(v);
    if (!coerenteSott(v, bulkSott)) setBulkSott("");
  };
  const cambiaBulkPri = (v: string) => {
    setBulkPri(v);
    if (!coerenteSec(v, bulkSec)) setBulkSec("");
  };

  const salvaBulk = async () => {
    if (!selMov.size) return;
    setBulkBusy(true);
    try {
      let fatti = 0;
      for (const id of selMov) {
        await spUpdateMovimento({
          data: {
            movimentoId: id,
            // Vuoto = NON toccare (semantica del blocco, diversa dalla
            // matita singola dove vuoto = cancella).
            ...(bulkTip.trim() ? { tipologia: bulkTip.trim(), daVerificare: false } : {}),
            ...(bulkSott.trim() ? { sottocategoria: bulkSott.trim() } : {}),
            ...(bulkPri.trim() ? { allocPrimaria: bulkPri.trim() } : {}),
            ...(bulkSec.trim() ? { allocSecondaria: bulkSec.trim() } : {}),
          },
        });
        fatti++;
      }
      toast.success(t("fin.selFatto"), { description: `${fatti} ${t("fin.rows")}` });
      setBulkOpen(false);
      setSelMov(new Set());
      setBulkTip("");
      setBulkSott("");
      setBulkPri("");
      setBulkSec("");
      loadMovimenti(anni);
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setBulkBusy(false);
    }
  };

  const isDirettore =
    session != null &&
    (session.ruolo === "amministratore_sistema" || isSupervisoreGlobale(session.codice ?? ""));

  const loadMovimenti = (a: number[]) => {
    setMovimenti(null);
    // Dal server si chiede l'INTERVALLO min-max; le selezioni non contigue
    // (es. 2024+2026) vengono rifinite client-side nel filtro.
    const range = a.length
      ? { from: `${Math.min(...a)}-01-01`, to: `${Math.max(...a)}-12-31` }
      : {};
    spGetMovimenti({ data: range })
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch((err) => {
        setMovimenti([]);
        toast.error(t("fin.errLoad"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  };
  // Le anomalie non hanno filtro anno: sono poche e vanno sanate tutte.
  const loadAnomalie = () => {
    spGetMovimenti({ data: { soloDaVerificare: true } })
      .then((l) => setAnomalie(l as SpMovimento[]))
      .catch(() => setAnomalie([]));
  };
  useEffect(() => {
    spGetDettagliDistinte()
      .then((l) => setDistinte(l as DettaglioDistinta[]))
      .catch(() => setDistinte([]));
    spGetRosterDipendenti()
      .then((l) => setRosterDip(l as DipendenteRoster[]))
      .catch(() => setRosterDip([]));
  }, []);

  // Parser del report "Esiti pagamenti" BPM (xlsx o csv, 16 colonne):
  // riconosce l'intestazione ovunque sia e legge per NOME colonna.
  const parseEsiti = async (file: File) => {
    let tab2d: string[][] = [];
    const isoLocale = (d: Date) =>
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
        d.getDate(),
      ).padStart(2, "0")}`;
    if (/\.csv$/i.test(file.name)) {
      const testo = await file.text();
      tab2d = testo
        .split(/\r?\n/)
        .filter((l) => l.trim())
        .map((l) => l.split(";").map((c) => c.replace(/^"+|"+$/g, "").trim()));
    } else {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      tab2d = (XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as unknown[][]).map((r) =>
        (r ?? []).map((c) => (c instanceof Date ? isoLocale(c) : String(c ?? "").trim())),
      );
    }
    const hIdx = tab2d.findIndex((r) => r.some((c) => c.toLowerCase().startsWith("beneficiario")));
    if (hIdx < 0) return [];
    const header = tab2d[hIdx].map((c) => c.toLowerCase());
    const col = (pfx: string) => header.findIndex((h) => h.startsWith(pfx));
    const iData = col("esecuzione");
    const iBen = col("beneficiario");
    const iTipo = col("tipo pagamento");
    const iImp = col("importo");
    const iDesc = col("descrizione causale");
    const iId = col("identificativo");
    if (iBen < 0 || iImp < 0 || iId < 0) return [];
    const out: NonNullable<typeof distPreview> = [];
    for (const r of tab2d.slice(hIdx + 1)) {
      const ben = (r[iBen] ?? "").trim();
      const idRaw = (r[iId] ?? "").trim();
      const rawImp = (r[iImp] ?? "").trim();
      // Formato italiano "1314,2900" o numero gia' decimale dall'xlsx.
      const importo = rawImp.includes(",")
        ? Number(rawImp.replace(/\./g, "").replace(",", "."))
        : Number(rawImp);
      let dataEs = (iData >= 0 ? (r[iData] ?? "") : "").trim().slice(0, 10);
      const mIt = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(dataEs);
      if (mIt) dataEs = `${mIt[3]}-${mIt[2]}-${mIt[1]}`;
      if (!ben || !idRaw || !importo || !Number.isFinite(importo)) continue;
      out.push({
        // L'identificativo da solo NON e' univoco (distinte stipendi):
        // la chiave anti-doppioni e' identificativo|beneficiario.
        idPagamento: `${idRaw}|${ben}`.slice(0, 240),
        dataEsecuzione: dataEs,
        beneficiario: ben,
        importo: Math.round(importo * 100) / 100,
        tipoPagamento: iTipo >= 0 ? (r[iTipo] ?? "").trim() : "",
        descrizione: iDesc >= 0 ? (r[iDesc] ?? "").trim() : "",
      });
    }
    return out;
  };

  const importaDistinte = async () => {
    if (!distPreview?.length) return;
    setDistBusy(true);
    try {
      let create = 0;
      let gia = 0;
      for (let i = 0; i < distPreview.length; i += 60) {
        const esito = (await spImportDistinta({
          data: { rows: distPreview.slice(i, i + 60) },
        })) as { create: number; giaPresenti: number };
        create += esito.create;
        gia += esito.giaPresenti;
      }
      const agg = (await spGetDettagliDistinte()) as DettaglioDistinta[];
      setDistinte(agg);
      setDistPreview(null);
      toast.success(t("fin.distEsitoOk"), {
        description: `+${create} · ${t("fin.distGia")}: ${gia}`,
      });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setDistBusy(false);
    }
  };

  // Le disposizioni raggruppate per giorno+tipo: il gruppo la cui somma
  // coincide con l'addebito cumulativo (al centesimo, data entro 6 giorni)
  // e' il dettaglio di quel movimento.
  const distGruppi = useMemo(() => {
    const map = new Map<
      string,
      {
        data: string;
        tipo: string;
        somma: number;
        righe: DettaglioDistinta[];
        movChiavi: string[];
        primaRigaId: string;
      }
    >();
    for (const d of distinte ?? []) {
      const k = `${d.dataEsecuzione}|${d.tipoPagamento}`;
      const g = map.get(k) ?? {
        data: d.dataEsecuzione,
        tipo: d.tipoPagamento,
        somma: 0,
        righe: [] as DettaglioDistinta[],
        movChiavi: [] as string[],
        primaRigaId: d.id,
      };
      g.somma += d.importo;
      g.righe.push(d);
      if (d.movimentoChiave && !g.movChiavi.includes(d.movimentoChiave))
        g.movChiavi.push(d.movimentoChiave);
      map.set(k, g);
    }
    return [...map.values()].map((g) => ({ ...g, somma: Math.round(g.somma * 100) / 100 }));
  }, [distinte]);
  const distintaDi = (m: SpMovimento) => {
    if (!distGruppi.length) return null;
    // Aggancio MANUALE: vince su tutto (badge anche quando la somma non torna),
    // e una distinta puo' essere agganciata a PIU' movimenti (tranche).
    const manuale = distGruppi.find((g) => g.movChiavi.includes(m.chiave));
    if (manuale) return manuale;
    if (m.importo >= 0) return null;
    const target = Math.round(-m.importo * 100) / 100;
    let best: (typeof distGruppi)[number] | null = null;
    let bestDiff = 7;
    for (const g of distGruppi) {
      if (Math.abs(g.somma - target) > 1) continue;
      const diff = Math.abs(
        (new Date(`${m.dataContabile}T00:00:00`).getTime() -
          new Date(`${g.data}T00:00:00`).getTime()) /
          86400000,
      );
      if (diff <= 6 && diff < bestDiff) {
        best = g;
        bestDiff = diff;
      }
    }
    return best;
  };

  // TRANCHE: la banca puo' addebitare una distinta in piu' movimenti.
  // Per ogni distinta NON agganciata si cerca il sottoinsieme di uscite
  // libere (entro 10 giorni, fino a 30 candidate) che somma piu' vicino
  // alla distinta: quadratura al centesimo = tranche certe.
  const trovaTranche = (
    g: (typeof distGruppi)[number],
  ): { scelti: SpMovimento[]; diffCent: number } | null => {
    // Priorita' A DUE LIVELLI: "beneficiari vari/distinta/stipendi" e'
    // un segnale FORTE, "vostra disposizione" e' la dicitura di qualsiasi
    // bonifico (segnale debole). In un giorno di paghe con decine di
    // bonifici singoli, senza questa gerarchia il tetto dei 30 candidati
    // si riempiva dei bonifici sbagliati e le tranche vere restavano fuori.
    const kwForte = /beneficiari|distint|stipend|emolument|salari/i;
    const kwDebole = /disposizione/i;
    const peso = (m: SpMovimento) => {
      const txt = `${m.cliente} ${m.descrizione} ${m.causale ?? ""}`;
      return kwForte.test(txt) ? 0 : kwDebole.test(txt) ? 1 : 2;
    };
    const giorniDa = (m: SpMovimento) =>
      Math.abs(
        (new Date(`${m.dataContabile}T00:00:00`).getTime() -
          new Date(`${g.data}T00:00:00`).getTime()) /
          86400000,
      );
    const cand = (movimenti ?? [])
      .filter((m) => m.importo < 0 && distintaDi(m) == null && giorniDa(m) <= 10)
      .sort((a2, b2) => {
        const ka = peso(a2);
        const kb = peso(b2);
        return ka !== kb ? ka - kb : giorniDa(a2) - giorniDa(b2);
      })
      .slice(0, 30);
    if (!cand.length) return null;
    // Meet in the middle SUL PIU' VICINO: le somme di meta' candidati in
    // un array ordinato, l'altra meta' cerca il complemento migliore.
    const arr = cand.map((m) => Math.round(-m.importo * 100));
    const target = Math.round(g.somma * 100);
    const metaN = Math.ceil(cand.length / 2);
    const nB = cand.length - metaN;
    const sommeA: [number, number][] = [];
    for (let mask = 0; mask < 1 << metaN; mask++) {
      let s2 = 0;
      for (let i2 = 0; i2 < metaN; i2++) if (mask & (1 << i2)) s2 += arr[i2];
      sommeA.push([s2, mask]);
    }
    sommeA.sort((x, y) => x[0] - y[0]);
    let best: { diff: number; maskA: number; maskB: number } | null = null;
    for (let mask = 0; mask < 1 << nB; mask++) {
      let s2 = 0;
      for (let i2 = 0; i2 < nB; i2++) if (mask & (1 << i2)) s2 += arr[metaN + i2];
      const want = target - s2;
      let lo = 0;
      let hi = sommeA.length - 1;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (sommeA[mid][0] < want) lo = mid + 1;
        else hi = mid;
      }
      for (const idx of [lo - 1, lo]) {
        if (idx < 0 || idx >= sommeA.length) continue;
        const coppia = sommeA[idx];
        if (coppia[1] === 0 && mask === 0) continue;
        const diff = Math.abs(coppia[0] + s2 - target);
        if (!best || diff < best.diff) best = { diff, maskA: coppia[1], maskB: mask };
      }
      if (best && best.diff === 0) break;
    }
    if (!best) return null;
    const scelto = best;
    const scelti = cand.filter((_, i2) =>
      i2 < metaN ? scelto.maskA & (1 << i2) : scelto.maskB & (1 << (i2 - metaN)),
    );
    // Anche UN solo movimento e' una risposta utile (addebito unico
    // con commissioni): la soglia minima di 2 scartava proprio quel caso.
    return scelti.length >= 1 ? { scelti, diffCent: scelto.diff } : null;
  };
  const trancheMap = useMemo(() => {
    const out = new Map<string, { scelti: SpMovimento[]; diffCent: number }>();
    if (!movimenti) return out;
    for (const g of distGruppi) {
      const k = `${g.data}|${g.tipo}`;
      if (movimenti.some((m) => g.movChiavi.includes(m.chiave))) continue;
      const auto = movimenti.some(
        (m) =>
          m.importo < 0 &&
          Math.abs(Math.round(-m.importo * 100) / 100 - g.somma) <= 1 &&
          Math.abs(
            (new Date(`${m.dataContabile}T00:00:00`).getTime() -
              new Date(`${g.data}T00:00:00`).getTime()) /
              86400000,
          ) <= 6,
      );
      if (auto) continue;
      const r = trovaTranche(g);
      if (r) out.set(k, r);
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [movimenti, distGruppi]);
  // AGGANCIO AUTOMATICO: quadratura al centesimo = si collega da solo,
  // senza click. Il ref evita partenze doppie mentre le scritture corrono.
  const trancheAutoBusy = useRef(false);
  useEffect(() => {
    if (trancheAutoBusy.current) return;
    const esatte = [...trancheMap.entries()].filter(([, r]) => r.diffCent === 0);
    if (!esatte.length) return;
    trancheAutoBusy.current = true;
    void (async () => {
      try {
        for (const [k, r] of esatte) {
          const g = distGruppi.find((x) => `${x.data}|${x.tipo}` === k);
          if (!g) continue;
          const libere = g.righe.filter((x) => !x.movimentoChiave);
          if (libere.length < r.scelti.length) continue;
          const fatti: { id: string; chiave: string }[] = [];
          for (let i = 0; i < r.scelti.length; i++) {
            await spSetDistintaMovimento({
              data: { id: libere[i].id, chiave: r.scelti[i].chiave },
            });
            fatti.push({ id: libere[i].id, chiave: r.scelti[i].chiave });
          }
          setDistinte((prev) =>
            (prev ?? []).map((x) => {
              const f = fatti.find((y) => y.id === x.id);
              return f ? { ...x, movimentoChiave: f.chiave } : x;
            }),
          );
          toast.success(t("fin.distAutoTrancheOk"), {
            description: `${fmtData(g.data)} ${g.tipo || ""} · ${r.scelti.length} × = ${fmtImporto(g.somma)} €`,
          });
        }
      } catch (err) {
        toast.error(t("common.error"), {
          description: err instanceof Error ? err.message : String(err),
        });
      } finally {
        trancheAutoBusy.current = false;
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [trancheMap]);

  // Spaccato del modal per APPALTO: ogni beneficiario passa nel
  // riconoscitore fuzzy dei dipendenti e prende l'appalto dall'anagrafica.
  // Memoria dell'appalto MANUALE per beneficiario: assegnato una volta,
  // vale per tutte le distinte (anche future) con lo stesso nome.
  const appaltoPerNome = useMemo(() => {
    const out = new Map<string, string>();
    for (const d of distinte ?? [])
      if (d.appalto?.trim()) out.set(d.beneficiario.toLowerCase(), d.appalto.trim());
    return out;
  }, [distinte]);
  // Risoluzione dell'appalto di una disposizione: anagrafica -> assegnazione
  // manuale della riga -> memoria per nome -> niente.
  const risolviAppaltoDist = (d: DettaglioDistinta) => {
    const nomi = (rosterDip ?? []).map((r) => r.nome);
    const nome = nomi.length ? matchDipendenteNome(d.beneficiario, nomi) : null;
    if (nome) {
      return {
        nome,
        appalto: (rosterDip ?? []).find((r) => r.nome === nome)?.appalto || "__senza__",
        manuale: false,
      };
    }
    const man = d.appalto?.trim() || appaltoPerNome.get(d.beneficiario.toLowerCase()) || "";
    return { nome: null, appalto: man, manuale: Boolean(man) };
  };

  const distSpaccato = useMemo(() => {
    if (!distModal || !rosterDip?.length) return null;
    const per = new Map<string, { n: number; somma: number }>();
    const nonRic: string[] = [];
    for (const d of distModal.righe) {
      const ris = risolviAppaltoDist(d);
      if (!ris.appalto) {
        nonRic.push(d.beneficiario);
        continue;
      }
      const g = per.get(ris.appalto) ?? { n: 0, somma: 0 };
      g.n++;
      g.somma += d.importo;
      per.set(ris.appalto, g);
    }
    if (!per.size) return null;
    return {
      righe: [...per.entries()]
        .map(([app, g]) => ({ app, n: g.n, somma: Math.round(g.somma * 100) / 100 }))
        .sort((a, b) => b.somma - a.somma),
      nonRic,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [distModal, rosterDip, distinte]);

  // CSV dello spaccato distinta ("diviso -> riconducibile, esportabile"):
  // una riga per disposizione, con dipendente riconosciuto e appalto.
  const esportaDistintaCsv = () => {
    if (!distModal) return;
    const nomi = (rosterDip ?? []).map((r) => r.nome);
    const righe = [...distModal.righe]
      .sort((a, b) => b.importo - a.importo)
      .map((d) => {
        const ris = risolviAppaltoDist(d);
        const nome = ris.nome;
        const app = ris.appalto === "__senza__" ? "" : ris.appalto;
        return [
          d.dataEsecuzione,
          d.tipoPagamento,
          d.beneficiario,
          d.importo.toFixed(2).replace(".", ","),
          nome ?? "",
          app,
          d.descrizione,
        ];
      });
    esportaCsvFile(
      `distinta-${distModal.data}`,
      [
        "Data esecuzione",
        "Tipo pagamento",
        "Beneficiario",
        "Importo",
        "Dipendente riconosciuto",
        "Appalto",
        "Causale",
      ],
      righe,
    );
  };

  // SPLIT AUTOMATICO: il movimento cumulativo agganciato a una distinta
  // si presenta come le sue DISPOSIZIONI (una riga per beneficiario, gia'
  // classificata), piu' una riga di RESTO se le somme non coincidono —
  // cosi' i totali quadrano al centesimo con l'estratto conto. L'archivio
  // banca non viene toccato: e' una lente, non una riscrittura.
  type MovVista = SpMovimento & { distVirtuale?: boolean; distChiave?: string };
  const movimentiVista = useMemo((): MovVista[] | null => {
    if (!movimenti) return null;
    if (!distGruppi.length) return movimenti;
    // TRANCHE MULTIPLE: la banca puo' addebitare la stessa distinta in
    // piu' movimenti (es. stipendi 69.345,22 pagati in 7 addebiti). Il
    // gruppo si espande UNA volta sola — al primo movimento incontrato —
    // e il resto si calcola sul TOTALE di tutte le tranche agganciate,
    // cosi' beneficiari mai duplicati e totali che quadrano al centesimo.
    const gruppoDi = new Map<string, (typeof distGruppi)[number]>();
    const totaleTranche = new Map<string, number>();
    for (const m of movimenti) {
      const g = m.importo < 0 ? distintaDi(m) : null;
      if (!g) continue;
      const k = `${g.data}|${g.tipo}`;
      gruppoDi.set(m.id, g);
      totaleTranche.set(k, Math.round(((totaleTranche.get(k) ?? 0) + m.importo) * 100) / 100);
    }
    const espansi = new Set<string>();
    const out: MovVista[] = [];
    for (const m of movimenti) {
      const g = gruppoDi.get(m.id);
      if (!g) {
        out.push(m);
        continue;
      }
      const chiaveG = `${g.data}|${g.tipo}`;
      if (espansi.has(chiaveG)) continue;
      espansi.add(chiaveG);
      let somma = 0;
      // "Pagamento Salario" vale SOLO per le distinte stipendi (o per i
      // beneficiari riconosciuti in anagrafica): le RiBa e le altre distinte
      // sono pagamenti fornitori e passano nelle REGOLE APPRESE, come
      // qualsiasi movimento (Califano -> la sua regola, non "salario").
      const eStipendi = (g.tipo || "").toLowerCase().includes("stipend");
      for (const d of g.righe) {
        somma += d.importo;
        const ris = risolviAppaltoDist(d);
        const appalto = ris.appalto && ris.appalto !== "__senza__" ? ris.appalto : "";
        let riga: MovVista = {
          ...m,
          id: `dist:${d.id}`,
          importo: Math.round(-d.importo * 100) / 100,
          cliente: d.beneficiario,
          descrizione: `${g.tipo || "distinta"} ${fmtData(g.data)}${d.descrizione ? ` · ${d.descrizione}` : ""}`,
          causale: "",
          tipologia: "",
          sottocategoria: "",
          allocPrimaria: "",
          allocSecondaria: "",
          nrFattura: "",
          daVerificare: false,
          distVirtuale: true,
          distChiave: chiaveG,
        };
        // PRIMA le regole apprese: le regole per nominativo (che portano
        // anche l'appalto) vincono sempre. Poi, per le distinte stipendi o
        // i beneficiari in anagrafica, il ripiego "Pagamento Salario" +
        // appalto della persona riempie SOLO i buchi rimasti.
        riga = applicaRegole(riga, regole ?? []);
        if (eStipendi || ris.nome) {
          if (!riga.tipologia) riga.tipologia = "Pagamento Salario";
          if (!riga.sottocategoria) riga.sottocategoria = "Pagamento Salario";
          if (!riga.allocSecondaria && appalto) {
            riga.allocSecondaria = appalto;
            if (!riga.allocPrimaria)
              riga.allocPrimaria = appalto.toLowerCase().startsWith("ufficio")
                ? "Costi generali"
                : "Appalto";
          }
        }
        out.push(riga);
      }
      const resto = Math.round(((totaleTranche.get(chiaveG) ?? m.importo) + somma) * 100) / 100;
      if (Math.abs(resto) > 0.005)
        out.push({
          ...m,
          id: `distresto:${m.id}`,
          importo: resto,
          descrizione: `${t("fin.distResto")} · ${m.descrizione}`,
          distVirtuale: true,
          distChiave: chiaveG,
        });
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [movimenti, distGruppi, rosterDip, distinte, regole]);

  const distinteCard = (
    <div className="mb-4 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      <div className="text-sm font-semibold text-foreground mb-1">{t("fin.distTitolo")}</div>
      <p className="text-xs text-muted-foreground mb-3">{t("fin.distDesc")}</p>
      <input
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={(e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (!file) return;
          void parseEsiti(file).then((righe) => {
            if (!righe.length) toast.error(t("fin.distNoRighe"));
            setDistPreview(righe.length ? righe : null);
          });
        }}
        className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-1.5 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
      />
      {distPreview && (
        <div className="mt-3 rounded-xl border border-border/60 bg-muted/20 p-3 text-sm">
          <p className="mb-2 text-foreground">
            {distPreview.length} {t("fin.distRighe")}
          </p>
          <ul className="mb-3 space-y-0.5 text-xs text-muted-foreground">
            {[
              ...distPreview
                .reduce((map, r) => {
                  const k = `${r.dataEsecuzione}|${r.tipoPagamento}`;
                  const g = map.get(k) ?? { n: 0, somma: 0 };
                  g.n++;
                  g.somma += r.importo;
                  map.set(k, g);
                  return map;
                }, new Map<string, { n: number; somma: number }>())
                .entries(),
            ].map(([k, g]) => (
              <li key={k}>
                {fmtData(k.split("|")[0])} · {k.split("|")[1] || "—"} — {g.n} ×{" "}
                {fmtImporto(Math.round(g.somma * 100) / 100)} €
              </li>
            ))}
          </ul>
          <div className="flex gap-2">
            <button
              type="button"
              disabled={distBusy}
              onClick={() => void importaDistinte()}
              className="rounded-lg bg-primary px-4 py-1.5 text-sm font-medium text-primary-foreground disabled:opacity-50"
            >
              {distBusy ? t("common.loading") : t("fin.distImporta")}
            </button>
            <button
              type="button"
              onClick={() => setDistPreview(null)}
              className="rounded-lg border border-border px-4 py-1.5 text-sm hover:bg-muted"
            >
              {t("common.cancel")}
            </button>
          </div>
        </div>
      )}
      {distinte != null && distinte.length > 0 && !distPreview && (
        <div className="mt-3">
          <p className="mb-1 text-[11px] text-muted-foreground">
            {distinte.length} {t("fin.distArchivio")}
          </p>
          {/* DIAGNOSTICA AGGANCIO: per ogni distinta si dice subito se ha
              trovato il suo movimento cumulativo (stessa somma ±1€, data
              entro 6 giorni) — e se no, quale importo si sta cercando. */}
          <table className="w-full text-[12px]">
            <tbody>
              {distGruppi
                .sort((a, b) => (a.data < b.data ? 1 : -1))
                .map((g) => {
                  // Copertura: somma dei movimenti agganciati (manuali) o
                  // il singolo movimento trovato in automatico.
                  const collegati = (movimenti ?? []).filter((m) => g.movChiavi.includes(m.chiave));
                  const autoMov =
                    collegati.length === 0
                      ? (movimenti ?? []).find(
                          (m) =>
                            m.importo < 0 &&
                            Math.abs(Math.round(-m.importo * 100) / 100 - g.somma) <= 1 &&
                            Math.abs(
                              (new Date(`${m.dataContabile}T00:00:00`).getTime() -
                                new Date(`${g.data}T00:00:00`).getTime()) /
                                86400000,
                            ) <= 6,
                        )
                      : undefined;
                  const coperto =
                    Math.round(collegati.reduce((s2, m) => s2 - m.importo, 0) * 100) / 100;
                  const completo = collegati.length
                    ? Math.abs(coperto - g.somma) <= 1
                    : Boolean(autoMov);
                  // La chiave nuova si salva su una riga della distinta ancora
                  // senza chiave: 70 disposizioni = spazio per 70 tranche.
                  const rigaLibera = g.righe.find((r) => !r.movimentoChiave);
                  // Tranche dal motore condiviso: quadratura esatta =
                  // aggancio automatico (il bottone resta come riserva);
                  // quasi-quadratura = indizio per l'analisi manuale.
                  const tr =
                    collegati.length > 0 || autoMov
                      ? undefined
                      : trancheMap.get(`${g.data}|${g.tipo}`);
                  const trancheCand = tr && tr.diffCent === 0 ? tr.scelti : null;
                  const trancheQuasi = tr && tr.diffCent > 0 ? tr : null;
                  return (
                    <tr key={`${g.data}|${g.tipo}`} className="border-t border-border/40">
                      <td className="py-1 pr-3 whitespace-nowrap">{fmtData(g.data)}</td>
                      <td className="py-1 pr-3">{g.tipo || "—"}</td>
                      <td className="py-1 pr-3 text-right tabular-nums whitespace-nowrap">
                        {g.righe.length} × {fmtImporto(g.somma)} €
                      </td>
                      <td className="py-1 pr-3">
                        {collegati.length > 0 ? (
                          <button
                            type="button"
                            onClick={() => setDistModal(g)}
                            className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-medium ${completo ? "bg-status-present/10 text-status-present" : "bg-primary/10 text-primary"}`}
                          >
                            <Users className="h-3 w-3" /> {collegati.length} mov ·{" "}
                            {fmtImporto(coperto)}
                            {completo ? " €" : ` di ${fmtImporto(g.somma)} €`}
                          </button>
                        ) : autoMov ? (
                          <button
                            type="button"
                            onClick={() => setDistModal(g)}
                            className="inline-flex items-center gap-1 rounded-full bg-status-present/10 px-2 py-0.5 text-[11px] font-medium text-status-present"
                          >
                            <Users className="h-3 w-3" /> {t("fin.distAgganciata")}{" "}
                            {fmtData(autoMov.dataContabile)}
                          </button>
                        ) : (
                          <span className="rounded-full bg-status-absent/10 px-2 py-0.5 text-[11px] font-medium text-status-absent">
                            {t("fin.distNonAgganciata")}
                          </span>
                        )}
                        {!completo && (
                          <span className="ml-1.5 inline-flex flex-wrap items-center gap-1.5">
                            {/* Candidati vicini per data (uscite, ±15 gg),
                                ordinati per somiglianza d'importo: scelto
                                uno, l'aggancio si salva su SharePoint e il
                                badge compare anche con somma diversa. */}
                            <select
                              defaultValue=""
                              onChange={(e) => {
                                const chiave = e.target.value;
                                if (!chiave) return;
                                void spSetDistintaMovimento({
                                  data: { id: rigaLibera?.id ?? g.primaRigaId, chiave },
                                })
                                  .then(() => {
                                    const idScelto = rigaLibera?.id ?? g.primaRigaId;
                                    setDistinte((prev) =>
                                      (prev ?? []).map((x) =>
                                        x.id === idScelto ? { ...x, movimentoChiave: chiave } : x,
                                      ),
                                    );
                                    toast.success(t("fin.distAggOk"));
                                  })
                                  .catch((err) =>
                                    toast.error(t("common.error"), {
                                      description: err instanceof Error ? err.message : String(err),
                                    }),
                                  );
                              }}
                              className="max-w-72 rounded border border-border bg-background px-1.5 py-0.5 text-[11px]"
                            >
                              <option value="">{t("fin.distAggScegli")}</option>
                              {(movimenti ?? []).filter(
                                (m) =>
                                  m.importo < 0 &&
                                  Math.abs(
                                    (new Date(`${m.dataContabile}T00:00:00`).getTime() -
                                      new Date(`${g.data}T00:00:00`).getTime()) /
                                      86400000,
                                  ) <= 15,
                              ).length === 0 && (
                                <option value="" disabled>
                                  {g.data > new Date().toISOString().slice(0, 10)
                                    ? t("fin.distAggFuturo")
                                    : t("fin.distAggNessunMov")}
                                </option>
                              )}
                              {(movimenti ?? [])
                                .filter(
                                  (m) =>
                                    m.importo < 0 &&
                                    !g.movChiavi.includes(m.chiave) &&
                                    Math.abs(
                                      (new Date(`${m.dataContabile}T00:00:00`).getTime() -
                                        new Date(`${g.data}T00:00:00`).getTime()) /
                                        86400000,
                                    ) <= 15,
                                )
                                .sort(
                                  (a, b) =>
                                    Math.abs(-a.importo - (g.somma - coperto)) -
                                    Math.abs(-b.importo - (g.somma - coperto)),
                                )
                                .slice(0, 8)
                                .map((m) => (
                                  <option key={m.chiave} value={m.chiave}>
                                    {fmtData(m.dataContabile)} · {fmtImporto(m.importo)} ·{" "}
                                    {m.descrizione.slice(0, 40)}
                                  </option>
                                ))}
                            </select>
                            {trancheCand && (
                              <button
                                type="button"
                                disabled={trancheBusy}
                                onClick={() => {
                                  void (async () => {
                                    const libere = g.righe.filter((r) => !r.movimentoChiave);
                                    if (libere.length < trancheCand.length) return;
                                    setTrancheBusy(true);
                                    try {
                                      const fatti: { id: string; chiave: string }[] = [];
                                      for (let i = 0; i < trancheCand.length; i++) {
                                        await spSetDistintaMovimento({
                                          data: { id: libere[i].id, chiave: trancheCand[i].chiave },
                                        });
                                        fatti.push({
                                          id: libere[i].id,
                                          chiave: trancheCand[i].chiave,
                                        });
                                      }
                                      setDistinte((prev) =>
                                        (prev ?? []).map((x) => {
                                          const f = fatti.find((y) => y.id === x.id);
                                          return f ? { ...x, movimentoChiave: f.chiave } : x;
                                        }),
                                      );
                                      toast.success(t("fin.distTrancheOk"), {
                                        description: `${trancheCand.length} × = ${fmtImporto(g.somma)} €`,
                                      });
                                    } catch (err) {
                                      toast.error(t("common.error"), {
                                        description:
                                          err instanceof Error ? err.message : String(err),
                                      });
                                    } finally {
                                      setTrancheBusy(false);
                                    }
                                  })();
                                }}
                                className="rounded-full bg-primary px-2.5 py-0.5 text-[11px] font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                              >
                                {trancheBusy
                                  ? t("common.loading")
                                  : `${t("fin.distTrancheBtn")} (${trancheCand.length} × = ${fmtImporto(g.somma)} €)`}
                              </button>
                            )}
                            {trancheQuasi && (
                              <span className="text-[11px] text-muted-foreground">
                                {t("fin.distQuasi")}: {trancheQuasi.scelti.length} × ={" "}
                                {fmtImporto(
                                  Math.round(
                                    trancheQuasi.scelti.reduce((s2, m) => s2 - m.importo, 0) * 100,
                                  ) / 100,
                                )}{" "}
                                € (Δ {fmtImporto(trancheQuasi.diffCent / 100)} €)
                              </span>
                            )}
                            {trancheQuasi && (
                              <button
                                type="button"
                                disabled={trancheBusy}
                                onClick={() => {
                                  void (async () => {
                                    const scelti = trancheQuasi.scelti;
                                    const libere = g.righe.filter((r) => !r.movimentoChiave);
                                    if (libere.length < scelti.length) return;
                                    const elenco = scelti
                                      .map(
                                        (m) =>
                                          `${fmtData(m.dataContabile)} · ${fmtImporto(m.importo)} € · ${m.descrizione.slice(0, 40)}`,
                                      )
                                      .join("\n");
                                    if (
                                      !window.confirm(
                                        `${t("fin.distQuasiConfirm")}\n\n${elenco}\n\nΔ ${fmtImporto(trancheQuasi.diffCent / 100)} €`,
                                      )
                                    )
                                      return;
                                    setTrancheBusy(true);
                                    try {
                                      const fatti: { id: string; chiave: string }[] = [];
                                      for (let i = 0; i < scelti.length; i++) {
                                        await spSetDistintaMovimento({
                                          data: { id: libere[i].id, chiave: scelti[i].chiave },
                                        });
                                        fatti.push({ id: libere[i].id, chiave: scelti[i].chiave });
                                      }
                                      setDistinte((prev) =>
                                        (prev ?? []).map((x) => {
                                          const f = fatti.find((y) => y.id === x.id);
                                          return f ? { ...x, movimentoChiave: f.chiave } : x;
                                        }),
                                      );
                                      toast.success(t("fin.distAggOk"), {
                                        description: `${scelti.length} × · Δ ${fmtImporto(trancheQuasi.diffCent / 100)} €`,
                                      });
                                    } catch (err) {
                                      toast.error(t("common.error"), {
                                        description:
                                          err instanceof Error ? err.message : String(err),
                                      });
                                    } finally {
                                      setTrancheBusy(false);
                                    }
                                  })();
                                }}
                                className="rounded-full border border-primary/40 px-2.5 py-0.5 text-[11px] font-medium text-primary hover:bg-primary/10 disabled:opacity-50"
                              >
                                {trancheBusy ? t("common.loading") : t("fin.distQuasiBtn")}
                              </button>
                            )}
                            {collegati.length === 0 && !autoMov && !tr && (
                              <span className="text-[11px] text-muted-foreground">
                                {t("fin.distNessunCand")}
                              </span>
                            )}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
            </tbody>
          </table>
          <p className="mt-1 text-[10px] text-muted-foreground">{t("fin.distAggancioNota")}</p>
          {(() => {
            // Vista speculare: le uscite che dalla descrizione sembrano
            // pagamenti cumulativi (beneficiari vari, distinte, stipendi)
            // ma NON hanno trovato nessuna distinta. Per ognuna si mostra
            // la distinta piu' vicina e di quanto manca l'aggancio.
            const orfani = (movimenti ?? [])
              .filter(
                (m) =>
                  m.importo < 0 &&
                  /beneficiari|distint|stipend|emolument|salari|disposizione/i.test(
                    `${m.descrizione} ${m.causale ?? ""}`,
                  ) &&
                  distintaDi(m) == null,
              )
              .sort((a, b) => (a.dataContabile < b.dataContabile ? 1 : -1))
              .slice(0, 10);
            if (!orfani.length) return null;
            return (
              <div className="mt-3 rounded-lg border border-status-absent/30 bg-status-absent/5 p-2">
                <p className="mb-1 text-[11px] font-medium text-status-absent">
                  {t("fin.distOrfaniTitolo")} ({orfani.length})
                </p>
                <table className="w-full text-[11px]">
                  <tbody>
                    {orfani.map((m) => {
                      const target = Math.round(-m.importo * 100) / 100;
                      let best: (typeof distGruppi)[number] | null = null;
                      let bestScore = Infinity;
                      for (const g of distGruppi) {
                        const dEuro = Math.abs(g.somma - target);
                        const dGg = Math.abs(
                          (new Date(`${m.dataContabile}T00:00:00`).getTime() -
                            new Date(`${g.data}T00:00:00`).getTime()) /
                            86400000,
                        );
                        const score = dEuro + dGg * 10;
                        if (score < bestScore) {
                          bestScore = score;
                          best = g;
                        }
                      }
                      const dEuro = best
                        ? Math.round(Math.abs(best.somma - target) * 100) / 100
                        : 0;
                      const dGg = best
                        ? Math.round(
                            Math.abs(
                              (new Date(`${m.dataContabile}T00:00:00`).getTime() -
                                new Date(`${best.data}T00:00:00`).getTime()) /
                                86400000,
                            ),
                          )
                        : 0;
                      return (
                        <tr key={m.id} className="border-t border-border/30">
                          <td className="py-1 pr-3 whitespace-nowrap">
                            {fmtData(m.dataContabile)}
                          </td>
                          <td className="py-1 pr-3 text-right tabular-nums whitespace-nowrap">
                            {fmtImporto(m.importo)} €
                          </td>
                          <td className="max-w-64 truncate py-1 pr-3" title={m.descrizione}>
                            {m.descrizione}
                          </td>
                          <td className="py-1 text-muted-foreground">
                            {best
                              ? `${t("fin.distOrfanoVicina")}: ${fmtData(best.data)} · ${best.righe.length} × ${fmtImporto(best.somma)} € (Δ ${fmtImporto(dEuro)} €, ${dGg} gg)`
                              : t("fin.distOrfanoNiente")}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <p className="mt-1 text-[10px] text-muted-foreground">{t("fin.distOrfaniNota")}</p>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );

  const loadStorico = () => {
    spGetImportStorico()
      .then((l) => setStorico(l as ImportStoricoRiga[]))
      .catch(() => setStorico([]));
  };
  const loadRegole = () => {
    spGetRegoleFinanza()
      .then((l) => setRegole(l as RegolaFinanza[]))
      .catch(() => setRegole([]));
    // Termini d'incasso per cliente: vivono nella stessa tab.
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TermineRiga[]))
      .catch(() => setTermini([]));
    // Regole di classificazione delle fatture passive.
    spGetRegoleFatture()
      .then((l) => setRegoleFat(l as RegolaFatturaRiga[]))
      .catch(() => setRegoleFat([]));
  };

  // --- Regole di classificazione passive -------------------------------------
  type RegolaFatturaRiga = {
    id?: string;
    fornitore: string;
    tipologia?: string;
    clienteRif?: string;
  };
  const [regoleFat, setRegoleFat] = useState<RegolaFatturaRiga[] | null>(null);
  const [rfFornitore, setRfFornitore] = useState("");
  const [rfTipologia, setRfTipologia] = useState("");
  const [rfCliente, setRfCliente] = useState("");
  const [rfBusy, setRfBusy] = useState(false);

  const salvaRegolaFat = async () => {
    if (!rfFornitore.trim() || (!rfTipologia.trim() && !rfCliente.trim())) {
      toast.error(t("fin.rfInvalida"));
      return;
    }
    setRfBusy(true);
    try {
      await spCreateRegolaFattura({
        data: {
          fornitore: rfFornitore.trim(),
          tipologia: rfTipologia.trim() || undefined,
          clienteRif: rfCliente.trim() || undefined,
        },
      });
      setRfFornitore("");
      setRfTipologia("");
      setRfCliente("");
      loadRegole();
      toast.success(t("fin.rfSalvata"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setRfBusy(false);
    }
  };

  const eliminaRegolaFat = async (r: RegolaFatturaRiga) => {
    if (!r.id) return;
    if (!window.confirm(t("fin.rfDeleteConfirm"))) return;
    setRfBusy(true);
    try {
      await spDeleteRegolaFattura({ data: { id: r.id } });
      setRegoleFat((prev) => (prev ? prev.filter((x) => x.id !== r.id) : prev));
      toast.success(t("fin.rfEliminata"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setRfBusy(false);
    }
  };

  // --- Termini d'incasso (giorni di pagamento per cliente) ------------------
  type TermineRiga = {
    cliente: string;
    giorni: number;
    direzione?: "Emessa" | "Ricevuta";
    email?: string;
    oggetto?: string;
  };
  const [termini, setTermini] = useState<TermineRiga[] | null>(null);
  const [tCliente, setTCliente] = useState("");
  const [tGiorni, setTGiorni] = useState("");
  const [tEmail, setTEmail] = useState("");
  // Parole chiave sull'oggetto fattura: rendono il termine una REGOLA
  // (es. IMILE + "locazione, affitto" -> 0 giorni, a vista).
  const [tOggetto, setTOggetto] = useState("");
  const [tBusy, setTBusy] = useState(false);
  // Termini DIREZIONALI: scheda Clienti (quanto ci pagano) e Fornitori
  // (quando paghiamo noi, ritardo sui NOSTRI pagamenti — default 30gg).
  const [tDirezione, setTDirezione] = useState<"Emessa" | "Ricevuta">("Emessa");

  const salvaTermine = async () => {
    const giorni = Number(tGiorni);
    const oggetto = tOggetto.trim();
    // 0 giorni (pagamento a vista) e' valido solo insieme a una parola
    // chiave sull'oggetto: da solo sarebbe un termine senza senso.
    if (!tCliente.trim() || !Number.isFinite(giorni) || giorni < 0 || (giorni === 0 && !oggetto)) {
      toast.error(t("fin.termInvalido"));
      return;
    }
    setTBusy(true);
    try {
      await spImportTermini({
        data: {
          rows: [
            {
              cliente: tCliente.trim(),
              giorni,
              direzione: tDirezione,
              email: tEmail.trim(),
              oggetto: oggetto || undefined,
            },
          ],
        },
      });
      setTCliente("");
      setTGiorni("");
      setTEmail("");
      setTOggetto("");
      loadRegole();
      toast.success(t("fin.termSalvato"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setTBusy(false);
    }
  };

  const copiaTerminiFornitori = async () => {
    if (!window.confirm(t("fin.termCopiaConfirm"))) return;
    setTBusy(true);
    try {
      const r = (await spCopiaTerminiSuFornitori()) as { copiati: number; esistenti: number };
      loadRegole();
      toast.success(`${r.copiati} ${t("fin.termCopiati")}`);
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setTBusy(false);
    }
  };

  const eliminaTermine = async (cliente: string, oggetto?: string) => {
    if (!window.confirm(`${t("fin.termDeleteConfirm")} ${cliente}?`)) return;
    setTBusy(true);
    try {
      await spDeleteTermine({ data: { cliente, direzione: tDirezione, oggetto } });
      setTermini((prev) =>
        prev
          ? prev.filter((x) => !(x.cliente === cliente && (x.direzione ?? "Emessa") === tDirezione))
          : prev,
      );
      toast.success(t("fin.termEliminato"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setTBusy(false);
    }
  };
  // Saldo attuale dalla banca. Silenzioso: senza collegamento attivo
  // semplicemente non si mostra (né la colonna Saldo); la diagnosi vive nel
  // pannello Banca in Amministrazione.
  const loadEbSaldo = () => {
    spEbSaldo()
      .then((s) => setEbSaldoInfo(s as EbSaldoInfo | null))
      .catch(() => setEbSaldoInfo(null));
  };
  const refreshAll = (a: number[]) => {
    loadMovimenti(a);
    loadAnomalie();
    loadStorico();
    loadRegole();
    // Il saldo è ancorato al progressivo dell'archivio: si ricarica insieme
    // ai movimenti per restare coerente dopo import/annullamenti/sync.
    loadEbSaldo();
  };

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    const dir = s.ruolo === "amministratore_sistema" || isSupervisoreGlobale(s.codice ?? "");
    if (!dir) return;
    refreshAll(anni);
    loadEbStato(); // serve al tasto Aggiorna per sapere se il sync è attivo
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const cambiaAnni = (a: number[]) => {
    setAnni(a);
    loadMovimenti(a);
  };

  // --- Collegamento banca (PSD2): solo uso quotidiano -----------------------
  const errMsg = (err: unknown) => (err instanceof Error ? err.message : String(err));

  const loadEbStato = () => {
    spEbStato()
      .then((s) => setEbStato(s as EbStato))
      .catch(() => setEbStato(null));
  };
  // Saldo dopo la riga, ancorato al saldo attuale della banca: esatto ovunque
  // l'archivio sia continuo tra la riga e oggi.
  const saldoDopo = (m: SpMovimento): number =>
    Math.round((ebSaldoInfo!.saldo - (ebSaldoInfo!.progressivoFinale - m.progressivo)) * 100) / 100;

  // Collegamento pronto all'uso: Aggiorna scarica anche dalla banca.
  const ebAttivo = Boolean(ebStato?.configurato && ebStato?.contoIban && ebStato?.dataTaglio);

  const ebSync = async () => {
    setEbSyncBusy(true);
    const importId = `SYNC-${new Date().toISOString().slice(0, 19)}`;
    let scritti = 0;
    let doppioni = 0;
    let pendenti = 0;
    let saldoErrore: string | undefined;
    try {
      let continuation: string | undefined;
      let guard = 0;
      for (;;) {
        const r = (await spEbSincronizza({ data: { importId, continuation } })) as EbSyncResult;
        scritti += r.scritti;
        doppioni += r.doppioni;
        pendenti += r.pendenti;
        saldoErrore = r.saldoErrore;
        if (r.errori.length) throw new Error(r.errori[0]);
        setEbProgress(String(scritti));
        if (!r.continuation || ++guard > 100) break;
        continuation = r.continuation;
      }
      toast.success(t("fin.ebSyncDone"), {
        description: `${scritti} ${t("fin.ebNuovi")} · ${doppioni} ${t("fin.ebGiaPresenti")} · ${pendenti} ${t("fin.ebPendenti")}`,
      });
      // Sync ok ma saldo non aggiornato: si dice PERCHÉ (limite banca,
      // colonna mancante…), altrimenti l'assenza del saldo resta un mistero.
      if (saldoErrore) toast.warning(t("fin.ebSaldoAttuale"), { description: saldoErrore });
      loadEbStato();
      refreshAll(anni);
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setEbSyncBusy(false);
      setEbProgress("");
    }
  };

  // Tasto Aggiorna: con il collegamento banca attivo prima scarica i nuovi
  // movimenti (il sync ricarica poi tutto), altrimenti ricarica soltanto.
  const aggiorna = () => {
    if (ebAttivo && !ebSyncBusy) void ebSync();
    else refreshAll(anni);
  };

  // --- Import xlsx ----------------------------------------------------------
  const costruisciPreview = async (fileName: string, res: ParseFileResult) => {
    setParsing(true);
    try {
      // Le regole apprese valgono anche per l'anteprima (stesse del server).
      const regoleAttive = (await spGetRegoleFinanza().catch(() => [])) as RegolaFinanza[];
      const righe = parseEstratto(res.rows, regoleAttive);
      const chiavi = new Set((await spGetMovimentiChiavi()) as string[]);
      const nuove = righe.filter((r) => !chiavi.has(r.chiave));
      const date = righe.map((r) => r.dataContabile).sort();
      // CORREZIONI x100: una riga del file "nuova" che in archivio esiste
      // con lo STESSO giorno, stessa descrizione e importo pari a 100 volte
      // e' il gemello corrotto di un vecchio import col punto decimale.
      // Si corregge la riga d'archivio IN POSTO (importo + chiave): le
      // classificazioni e il lavoro manuale fatto sopra NON si toccano.
      const perTupla = new Map<string, MovimentoParsed[]>();
      for (const r of nuove) {
        if (Number.isInteger(r.importo)) continue; // solo importi coi centesimi
        const k = `${r.dataContabile}|${Math.round(Math.abs(r.importo) * 100)}|${r.importo < 0 ? "-" : "+"}`;
        const arr = perTupla.get(k) ?? [];
        arr.push(r);
        perTupla.set(k, arr);
      }
      const usate = new Set<string>();
      const correzioni: CorrezioneX100[] = [];
      for (const m of movimenti ?? []) {
        if (!Number.isInteger(m.importo) || m.importo === 0) continue;
        const k = `${m.dataContabile}|${Math.abs(m.importo)}|${m.importo < 0 ? "-" : "+"}`;
        const f = (perTupla.get(k) ?? []).find(
          (c) =>
            !usate.has(c.chiave) &&
            normalizeTesto(c.descrizione).slice(0, 60) ===
              normalizeTesto(m.descrizione).slice(0, 60),
        );
        if (!f) continue;
        usate.add(f.chiave);
        correzioni.push({
          id: m.id,
          dataContabile: m.dataContabile,
          da: m.importo,
          a: f.importo,
          chiave: f.chiave,
          descrizione: m.descrizione,
        });
      }
      setPreview({
        fileName,
        righe,
        nuove: nuove.filter((r) => !usate.has(r.chiave)),
        doppioni: righe.length - nuove.length,
        scartate: res.scartate,
        anomalie: nuove.filter((r) => r.daVerificare).length,
        dal: date[0] ?? "",
        al: date[date.length - 1] ?? "",
        correzioni,
      });
    } catch (err) {
      toast.error(t("fin.errFile"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setParsing(false);
    }
  };

  const onFile = async (file: File) => {
    setPreview(null);
    setSheetChoice(null);
    setParsing(true);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
      const sheets: SheetInfo[] = wb.SheetNames.map((name) => {
        const matrix = XLSX.utils.sheet_to_json(wb.Sheets[name], {
          header: 1,
          raw: true,
        }) as unknown[][];
        const res = parseMatrice(matrix);
        return { name, res: res && res.rows.length ? res : null };
      });
      const riconosciuti = sheets.filter((s) => s.res);
      if (riconosciuti.length === 0) {
        toast.error(t("fin.errFile"), { description: t("fin.errFileDesc") });
        return;
      }
      // Un solo foglio nel file → si importa quello. Più fogli → si chiede
      // SEMPRE quale usare (preselezionando il primo riconosciuto).
      if (sheets.length === 1) {
        await costruisciPreview(file.name, riconosciuti[0].res!);
      } else {
        setSheetChoice({ fileName: file.name, sheets, selected: riconosciuti[0].name });
      }
    } catch (err) {
      toast.error(t("fin.errFile"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setParsing(false);
    }
  };

  const confermaFoglio = async () => {
    if (!sheetChoice) return;
    const scelto = sheetChoice.sheets.find((s) => s.name === sheetChoice.selected);
    if (!scelto?.res) {
      toast.error(t("fin.errFile"), { description: t("fin.sheetNotRecognized") });
      return;
    }
    await costruisciPreview(sheetChoice.fileName, scelto.res);
    setSheetChoice(null);
  };

  const eseguiImport = async () => {
    if (!preview || preview.nuove.length === 0) return;
    setImporting(true);
    // Lotto di import: identifica queste righe nello storico (annullabile).
    const importId = `IMP-${new Date().toISOString().slice(0, 19)}`;
    let importati = 0;
    let doppioni = 0;
    const errori: string[] = [];
    try {
      for (let i = 0; i < preview.nuove.length; i += CHUNK) {
        const rows = preview.nuove.slice(i, i + CHUNK).map((r) => ({
          dataContabile: r.dataContabile,
          dataValuta: r.dataValuta,
          importo: r.importo,
          divisa: r.divisa,
          causale: r.causale,
          descrizione: r.descrizione,
          occ: r.occ,
        }));
        setImportProgress(`${Math.min(i + CHUNK, preview.nuove.length)} / ${preview.nuove.length}`);
        const res = await spImportMovimenti({ data: { rows, importId } });
        importati += res.importati;
        doppioni += res.doppioni;
        errori.push(...res.errori);
      }
      toast.success(t("fin.importDone"), {
        description: `${importati} ${t("fin.importedRows")}${doppioni ? ` · ${doppioni} ${t("fin.skippedDup")}` : ""}${errori.length ? ` · ${errori.length} ${t("common.error").toLowerCase()}` : ""}`,
      });
      if (errori.length) console.warn("Import movimenti — errori:", errori);
      setPreview(null);
      setShowImportEC(false);
      refreshAll(anni);
    } catch (err) {
      toast.error(t("fin.errImport"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImporting(false);
      setImportProgress("");
    }
  };

  // --- Annullamento import --------------------------------------------------
  const annullaImportGruppo = async (riga: ImportStoricoRiga) => {
    const etichetta = fmtImportId(riga.importId, t("fin.legacyImport"));
    if (
      !window.confirm(
        `${t("fin.annullaConfirm")}\n${etichetta} — ${riga.movimenti} ${t("fin.rows")}`,
      )
    )
      return;
    const importId = riga.importId || LEGACY_IMPORT_ID;
    setAnnullaBusy(riga.importId);
    setAnnullaProgress(0);
    try {
      let tot = 0;
      for (;;) {
        const r = (await spAnnullaImport({ data: { importId } })) as {
          eliminati: number;
          rimanenti: number;
        };
        tot += r.eliminati;
        setAnnullaProgress(tot);
        if (r.rimanenti <= 0) break;
        if (r.eliminati === 0) throw new Error("Annullamento interrotto: nessun progresso.");
      }
      toast.success(t("fin.annullaDone"), { description: `${tot} ${t("fin.rows")}` });
      refreshAll(anni);
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
      refreshAll(anni);
    } finally {
      setAnnullaBusy(null);
      setAnnullaProgress(0);
    }
  };

  // --- Regole apprese -------------------------------------------------------
  // Prefill del form regola a partire da un movimento ("insegna al sistema").
  const creaRegolaDa = (m: SpMovimento) => {
    setRSorgente(m.id);
    setRPattern(m.cliente || "");
    setRCampo("cliente");
    setRModo("esatto");
    setRTipologia(m.tipologia || "");
    setRSottocat(m.sottocategoria || "");
    setRAllocPri(m.allocPrimaria || "");
    setRAllocSec(m.allocSecondaria || "");
    setRCliente("");
    setRNote("");
    setRApplica(true);
    setTab("regole");
  };

  const submitRegola = async () => {
    setRBusy(true);
    setRProgress(0);
    try {
      // Doppioni via e lista spezzata se oltre il limite del campo.
      const blocchi = spezzaPattern(rPattern);
      const payload = {
        pattern: blocchi[0],
        campo: rCampo,
        modo: rModo,
        tipologia: rTipologia.trim() || undefined,
        sottocategoria: rSottocat.trim() || undefined,
        allocPrimaria: rAllocPri.trim() || undefined,
        allocSecondaria: rAllocSec.trim() || undefined,
        cliente: rCliente.trim() || undefined,
        note: rNote.trim() || undefined,
        segno: rSegno || undefined,
      };
      // PARACADUTE: prima di salvare si mostra QUANTI movimenti verrebbero
      // toccati — una regola troppo larga si riconosce dal numero.
      // Con l'archivio ancora in caricamento il conteggio mentirebbe ("0
      // movimenti" su una lista vuota): in quel caso lo si dice chiaramente.
      const conferma =
        movimenti == null
          ? window.confirm(t("fin.regolaImpattoND"))
          : window.confirm(
              `${t("fin.regolaImpatto1")} ${
                movimenti.filter((m) =>
                  blocchi.some((blocco) => matchRegola(m, { ...payload, pattern: blocco })),
                ).length
              } ${t("fin.regolaImpatto2")}`,
            );
      if (!conferma) {
        setRBusy(false);
        setRProgress(0);
        return;
      }
      // Modifica = aggiornamento SUL POSTO: mai piu' cancella-e-ricrea (una
      // create fallita dopo la delete ha bruciato due regole del direttore).
      // Regola GEMELLA gia' esistente (stessi esiti, stesso criterio)?
      // I termini nuovi si uniscono a lei invece di creare un doppione.
      const gemella =
        !rEditId &&
        (regole ?? []).find(
          (r) =>
            r.campo === payload.campo &&
            r.modo === payload.modo &&
            (r.tipologia ?? "") === (payload.tipologia ?? "") &&
            (r.sottocategoria ?? "") === (payload.sottocategoria ?? "") &&
            (r.allocPrimaria ?? "") === (payload.allocPrimaria ?? "") &&
            (r.allocSecondaria ?? "") === (payload.allocSecondaria ?? "") &&
            (r.cliente ?? "") === (payload.cliente ?? "") &&
            `${r.pattern}, ${payload.pattern}`.length <= MAX_PATTERN + 100,
        );
      if (gemella) {
        const fusi = spezzaPattern(`${gemella.pattern}, ${blocchi.join(", ")}`);
        await spUpdateRegolaFinanza({
          data: {
            regolaId: gemella.id ?? "",
            ...payload,
            pattern: fusi[0],
            note: payload.note ?? gemella.note,
          },
        });
        for (const blocco of fusi.slice(1))
          await spCreateRegolaFinanza({ data: { ...payload, pattern: blocco } });
        toast.info(t("fin.regolaUnitaEsistente"), {
          description: gemella.pattern.slice(0, 80),
        });
      } else if (rEditId) {
        await spUpdateRegolaFinanza({ data: { regolaId: rEditId, ...payload } });
      } else {
        await spCreateRegolaFinanza({ data: payload });
      }
      // Blocchi oltre il primo: regole gemelle (stessi esiti, altri termini).
      if (!gemella)
        for (const blocco of blocchi.slice(1))
          await spCreateRegolaFinanza({ data: { ...payload, pattern: blocco } });
      if (blocchi.length > 1)
        toast.info(`${t("fin.regolaSpezzata1")} ${blocchi.length} ${t("fin.regolaSpezzata2")}`);
      let applicati = 0;
      if (rApplica) {
        // Applicazione retroattiva a blocchi finché il server non ha finito.
        // Se i RIMANENTI non calano tra un giro e l'altro, qualcosa non si
        // riesce a scrivere: ci si ferma invece di girare a vuoto.
        for (const blocco of blocchi) {
          let ultimoRimanenti = Number.POSITIVE_INFINITY;
          for (;;) {
            const r = (await spApplicaRegolaFinanza({
              data: { ...payload, pattern: blocco },
            })) as {
              aggiornati: number;
              rimanenti: number;
            };
            applicati += r.aggiornati;
            setRProgress(applicati);
            if (r.rimanenti <= 0) break;
            if (r.aggiornati === 0) break; // safety: niente progresso
            if (r.rimanenti >= ultimoRimanenti) {
              toast.warning(t("fin.regolaLoopStop"));
              break;
            }
            ultimoRimanenti = r.rimanenti;
          }
        }
      }
      toast.success(t("fin.regolaCreata"), {
        description: rApplica ? `${applicati} ${t("fin.regolaApplicati")}` : undefined,
      });
      setRPattern("");
      setRTipologia("");
      setRSottocat("");
      setRAllocPri("");
      setRAllocSec("");
      setRCliente("");
      setRNote("");
      setRSegno("");
      setRSorgente(null);
      setREditId(null);
      loadRegole();
      if (rApplica) {
        loadMovimenti(anni);
        loadAnomalie();
      }
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setRBusy(false);
      setRProgress(0);
    }
  };

  const eliminaRegola = async (r: RegolaFinanza) => {
    if (!r.id) return;
    if (!window.confirm(t("fin.regolaDeleteConfirm"))) return;
    // Secondo bivio: OK = anche i movimenti modificati dalla regola tornano
    // alla classificazione automatica; ANNULLA = si elimina solo la regola.
    const ripristina = window.confirm(t("fin.regolaRipristinoConfirm"));
    try {
      await spDeleteRegolaFinanza({ data: { regolaId: r.id } });
      setRegole((prev) => (prev ?? []).filter((x) => x.id !== r.id));
      let ripristinati = 0;
      if (ripristina) {
        const payload = {
          pattern: r.pattern,
          campo: r.campo,
          modo: r.modo,
          tipologia: r.tipologia,
          cliente: r.cliente,
        };
        for (;;) {
          const esito = (await spAnnullaRegolaFinanza({ data: payload })) as {
            aggiornati: number;
            rimanenti: number;
          };
          ripristinati += esito.aggiornati;
          if (esito.rimanenti <= 0 || esito.aggiornati === 0) break;
        }
        loadMovimenti(anni);
        loadAnomalie();
      }
      toast.success(t("fin.regolaDeleted"), {
        description: ripristina ? `${ripristinati} ${t("fin.regolaRipristinati")}` : undefined,
      });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    }
  };

  // --- Sanatura -------------------------------------------------------------
  const apriEdit = (m: SpMovimento) => {
    setEditId(m.id);
    setEditTip(m.tipologia || "Altro");
    setEditSott(m.sottocategoria || "");
    setEditAllocPri(m.allocPrimaria || "");
    setEditAllocSec(m.allocSecondaria || "");
    setEditCliente(m.cliente);
    setEditNrFatt(m.nrFattura);
    setEditNote(m.note);
  };
  // "PERCHE' questa classificazione?": per il movimento in modifica, le
  // regole che fanno match e il TERMINE esatto che ha colpito. Si riusa
  // matchRegola su cloni a singolo termine: stessa semantica garantita.
  const spiegaRegole = (m: SpMovimento) =>
    (regole ?? [])
      .map((r) => {
        if (!matchRegola(m, r)) return null;
        const termine =
          r.pattern
            .split(/[,;\n]/)
            .map((x) => x.trim())
            .filter(Boolean)
            .find((termine2) => matchRegola(m, { ...r, pattern: termine2 })) ?? r.pattern;
        return { r, termine };
      })
      .filter(Boolean) as { r: RegolaFinanza; termine: string }[];

  const salvaEdit = async () => {
    if (!editId) return;
    setSaving(true);
    try {
      const updated = (await spUpdateMovimento({
        data: {
          movimentoId: editId,
          tipologia: editTip,
          sottocategoria: editSott.trim(),
          allocPrimaria: editAllocPri.trim(),
          allocSecondaria: editAllocSec.trim(),
          cliente: editCliente.trim(),
          nrFattura: editNrFatt.trim(),
          note: editNote.trim(),
          daVerificare: false,
        },
      })) as SpMovimento;
      setAnomalie((prev) => (prev ?? []).filter((m) => m.id !== editId));
      setMovimenti((prev) => (prev ? prev.map((m) => (m.id === updated.id ? updated : m)) : prev));
      setEditId(null);
      toast.success(t("fin.fixSaved"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSaving(false);
    }
  };

  // --- Derivati -------------------------------------------------------------
  // TUTTE le controparti con incassi negli anni selezionati, in ordine
  // DECRESCENTE di incassato: sono le voci del menu clienti. Calcolate prima
  // degli altri filtri, cosi' l'elenco non cambia mentre si affina la ricerca.
  const clientiTop = useMemo(() => {
    const somme = new Map<string, { nome: string; tot: number }>();
    for (const m of movimenti ?? []) {
      if (anni.length && !anni.includes(Number(m.dataContabile.slice(0, 4)))) continue;
      if (m.importo <= 0 || !m.cliente) continue;
      const k = clienteGroupKey(m.cliente) || m.cliente;
      const v = somme.get(k) ?? { nome: m.cliente, tot: 0 };
      v.tot += m.importo;
      somme.set(k, v);
    }
    return [...somme.entries()]
      .sort((a, b) => b[1].tot - a[1].tot)
      .map(([k, v]) => ({ v: k, label: v.nome }));
  }, [movimenti, anni]);

  // Imbuti di colonna sulla tabella movimenti (stile Excel, come su
  // Fatture): spunte sui valori distinti, a cascata con gli altri filtri.
  const [movFiltriTh, setMovFiltriTh] = useState<Record<string, Set<string>>>({});
  const [movThAperto, setMovThAperto] = useState<string | null>(null);
  const [movThCerca, setMovThCerca] = useState("");
  useEffect(() => {
    setMovFiltriTh({});
    setMovThAperto(null);
  }, [anni]);
  type MovColTh = {
    key: string;
    label: string;
    get: (m: SpMovimento) => string;
    /** Chiave di ORDINAMENTO della tendina: date in cronologico, importi
     *  in numerico (senza, l'elenco andrebbe in alfabetico: 01/04, 01/06...). */
    ord?: (v: string) => string | number;
  };
  const ordData = (v: string) => v.split("/").reverse().join("-");
  const ordImporto = (v: string) => Number(v.replace(/\./g, "").replace(",", ".")) || 0;
  const movColTh = useMemo<MovColTh[]>(
    () => [
      {
        key: "data",
        label: t("fin.dataContabile"),
        get: (m) => fmtData(m.dataContabile),
        ord: ordData,
      },
      {
        key: "valuta",
        label: t("fin.dataValuta"),
        get: (m) => fmtData(m.dataValuta),
        ord: ordData,
      },
      {
        key: "importo",
        label: t("common.amount"),
        get: (m) => fmtImporto(m.importo),
        ord: ordImporto,
      },
      { key: "causale", label: t("fin.causaleCol"), get: (m) => m.descrizione },
      { key: "tipo", label: t("common.type"), get: (m) => m.tipologia },
      { key: "cliforn", label: t("fin.cliForn"), get: (m) => m.cliente },
      { key: "nrfatt", label: t("fin.nrFattura"), get: (m) => m.nrFattura },
      { key: "note", label: t("fin.note"), get: (m) => m.note },
    ],
    [t],
  );
  const passaMovTh = (m: SpMovimento, escludi?: string) => {
    for (const c of movColTh) {
      if (c.key === escludi) continue;
      const sel = movFiltriTh[c.key];
      if (sel && sel.size > 0 && !sel.has(c.get(m))) return false;
    }
    return true;
  };

  const filtratiBase = useMemo(() => {
    let out: MovVista[] = movimentiVista ?? [];
    // Anni: il server ha fornito l'intervallo min-max; qui si rifiniscono le
    // selezioni non contigue. Tipologie e mesi: vuoto = tutti, altrimenti OR.
    if (anni.length) out = out.filter((m) => anni.includes(Number(m.dataContabile.slice(0, 4))));
    if (clientiF.length)
      out = out.filter(
        (m) => m.cliente && clientiF.includes(clienteGroupKey(m.cliente) || m.cliente),
      );
    if (tipiF.length) out = out.filter((m) => tipiF.includes(m.tipologia));
    // Sottocategoria e allocazioni: "__vuoto__" = senza valore, per stanare
    // il non classificato.
    if (sottF.length) out = out.filter((m) => sottF.includes(m.sottocategoria || "__vuoto__"));
    if (allocPriF.length)
      out = out.filter((m) => allocPriF.includes(m.allocPrimaria || "__vuoto__"));
    if (allocSecF.length)
      out = out.filter((m) => allocSecF.includes(m.allocSecondaria || "__vuoto__"));
    if (mesiF.length) out = out.filter((m) => mesiF.includes(Number(m.dataContabile.slice(5, 7))));
    if (contoF) out = out.filter((m) => (m.conto || "") === (contoF === "__vuoto__" ? "" : contoF));
    // Range di DATE (contabile) e di IMPORTI: un solo estremo = "da" o
    // "fino a" (per gli importi: maggiore/minore di).
    if (dataDaF) out = out.filter((m) => m.dataContabile >= dataDaF);
    if (dataAF) out = out.filter((m) => m.dataContabile <= dataAF);
    const impMin = impMinF.trim() ? Number(impMinF.replace(",", ".")) : null;
    const impMax = impMaxF.trim() ? Number(impMaxF.replace(",", ".")) : null;
    if (impMin != null && Number.isFinite(impMin)) out = out.filter((m) => m.importo >= impMin);
    if (impMax != null && Number.isFinite(impMax)) out = out.filter((m) => m.importo <= impMax);
    if (cercaF.trim()) {
      // Più termini separati da , o ; = basta che UNO corrisponda (per
      // trovare in un colpo "aereo, treno, dirigibile" e regolarli insieme).
      const termini = cercaF
        .toLowerCase()
        .split(/[,;]/)
        .map((x) => x.trim())
        .filter(Boolean);
      out = out.filter((m) =>
        termini.some(
          (q) =>
            m.cliente.toLowerCase().includes(q) ||
            m.descrizione.toLowerCase().includes(q) ||
            m.nrFattura.toLowerCase().includes(q) ||
            m.note.toLowerCase().includes(q),
        ),
      );
    }
    return out;
  }, [
    movimentiVista,
    anni,
    tipiF,
    mesiF,
    cercaF,
    clientiF,
    contoF,
    sottF,
    allocPriF,
    allocSecF,
    dataDaF,
    dataAF,
    impMinF,
    impMaxF,
  ]);
  const filtrati = useMemo(
    () =>
      filtratiBase.filter(
        (m) => passaMovTh(m) && (!soloDistinte || m.distChiave != null || distintaDi(m) != null),
      ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [filtratiBase, movFiltriTh, soloDistinte, distGruppi],
  );
  // Totale (e spezzato entrate/uscite) di QUELLO CHE SI VEDE coi filtri.
  const totaleFiltrato = useMemo(() => {
    let entrate = 0;
    let uscite = 0;
    for (const m of filtrati) {
      if (m.importo >= 0) entrate += m.importo;
      else uscite += m.importo;
    }
    return {
      totale: Math.round((entrate + uscite) * 100) / 100,
      entrate: Math.round(entrate * 100) / 100,
      uscite: Math.round(uscite * 100) / 100,
    };
  }, [filtrati]);

  // Pagine da RIGHE_PAGINA sulla tabella movimenti; il cambio di filtri o
  // anno riparte dalla prima (il clamp copre i ricaricamenti che accorciano
  // la lista, es. dopo una sincronizzazione).
  useEffect(() => {
    setPaginaMov(1);
  }, [
    tipiF,
    mesiF,
    cercaF,
    anni,
    clientiF,
    sottF,
    allocPriF,
    allocSecF,
    movFiltriTh,
    dataDaF,
    dataAF,
    impMinF,
    impMaxF,
  ]);
  const pagineMovTot = Math.max(1, Math.ceil(filtrati.length / righePagina));
  const pagMov = Math.min(paginaMov, pagineMovTot);
  const inizioMov = (pagMov - 1) * righePagina;

  const thFiltroMov = (c: MovColTh, extra = "") => {
    const sel = movFiltriTh[c.key];
    const attivo = (sel?.size ?? 0) > 0;
    const aperto = movThAperto === c.key;
    return (
      <th key={c.key} className={`py-2 pr-3 whitespace-nowrap relative ${extra}`}>
        <span className="inline-flex items-center gap-1">
          {c.label}
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              setMovThAperto(aperto ? null : c.key);
              setMovThCerca("");
            }}
            className={attivo ? "text-primary" : "text-muted-foreground/50 hover:text-foreground"}
            title={t("ft.thFiltra")}
          >
            <Filter className="h-3 w-3" fill={attivo ? "currentColor" : "none"} />
          </button>
        </span>
        {aperto &&
          (() => {
            // Valori distinti a CASCATA: contano tutti i filtri attivi
            // tranne quello di questa colonna.
            const base = filtratiBase.filter((m) => passaMovTh(m, c.key));
            const conteggi = new Map<string, number>();
            for (const m of base) {
              const v2 = c.get(m);
              conteggi.set(v2, (conteggi.get(v2) ?? 0) + 1);
            }
            const q = movThCerca.trim().toLowerCase();
            const valori = [...conteggi.entries()]
              .sort((a, b) => {
                const ka = c.ord ? c.ord(a[0]) : a[0];
                const kb = c.ord ? c.ord(b[0]) : b[0];
                return typeof ka === "number" && typeof kb === "number"
                  ? ka - kb
                  : String(ka).localeCompare(String(kb));
              })
              .filter(([v2]) => !q || v2.toLowerCase().includes(q));
            const scelte = sel ?? new Set<string>();
            const setSel = (ns: Set<string>) => setMovFiltriTh({ ...movFiltriTh, [c.key]: ns });
            return (
              <>
                <div className="fixed inset-0 z-30" onClick={() => setMovThAperto(null)} />
                <div
                  className="absolute left-0 top-full z-40 mt-1 w-64 rounded-lg border border-border bg-card p-2 shadow-[var(--shadow-elegant)] font-normal"
                  onClick={(e) => e.stopPropagation()}
                >
                  <input
                    autoFocus
                    value={movThCerca}
                    onChange={(e) => setMovThCerca(e.target.value)}
                    placeholder={t("ft.cerca")}
                    className="mb-1.5 w-full rounded-md border border-border bg-background px-2 py-1 text-xs text-foreground"
                  />
                  <div className="mb-1.5 flex gap-3 text-[11px]">
                    <button
                      type="button"
                      className="underline underline-offset-2"
                      onClick={() => setSel(new Set())}
                    >
                      {t("ft.thTutti")}
                    </button>
                    <button
                      type="button"
                      className="underline underline-offset-2"
                      onClick={() => setSel(new Set(valori.map(([v2]) => v2)))}
                    >
                      {t("ft.thSoloVisibili")}
                    </button>
                  </div>
                  <div className="max-h-60 overflow-auto">
                    {valori.slice(0, 500).map(([v2, cnt]) => (
                      <label
                        key={v2 || "__vuoto__"}
                        className="flex cursor-pointer items-center gap-2 rounded px-1 py-0.5 text-xs text-foreground hover:bg-muted"
                      >
                        <input
                          type="checkbox"
                          className="accent-primary"
                          checked={scelte.size === 0 || scelte.has(v2)}
                          onChange={() => {
                            const ns = new Set(
                              scelte.size === 0 ? [...conteggi.keys()] : [...scelte],
                            );
                            if (ns.has(v2)) ns.delete(v2);
                            else ns.add(v2);
                            setSel(ns.size === conteggi.size ? new Set() : ns);
                          }}
                        />
                        <span className="flex-1 truncate">{v2 || t("ft.thVuoto")}</span>
                        <span className="text-muted-foreground tabular-nums">{cnt}</span>
                      </label>
                    ))}
                    {valori.length > 500 && (
                      <p className="px-1 py-0.5 text-[10px] text-muted-foreground">
                        +{valori.length - 500}…
                      </p>
                    )}
                  </div>
                </div>
              </>
            );
          })()}
      </th>
    );
  };

  const tipologiePresenti = useMemo(() => {
    const set = new Set((movimenti ?? []).map((m) => m.tipologia).filter(Boolean));
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [movimenti]);
  // Valori presenti per i filtri di classificazione (con la voce "(vuota)").
  const opzClassifica = useMemo(() => {
    const raccogli = (get: (m: SpMovimento) => string) => {
      const set = new Set((movimenti ?? []).map(get).filter(Boolean));
      return [
        { v: "__vuoto__", label: t("fin.vuota") },
        ...[...set].sort((a, b) => a.localeCompare(b)).map((x) => ({ v: x, label: x })),
      ];
    };
    return {
      sottocategorie: raccogli((m) => m.sottocategoria),
      allocPrimarie: raccogli((m) => m.allocPrimaria),
      allocSecondarie: raccogli((m) => m.allocSecondaria),
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [movimenti]);

  const mesi = lang === "it" ? MESI_IT : MESI_EN;

  // Overview: pivot per cliente. Con un anno selezionato le colonne sono i
  // mesi; con "tutti gli anni" le colonne diventano gli anni. Incassi = soli
  // accrediti di tipo Incasso; Spese = tutti gli addebiti (importo < 0,
  // raggruppati per controparte o, in mancanza, per tipologia), in valore
  // assoluto.
  const overview = useMemo(() => {
    // Con UN solo anno selezionato le colonne sono i mesi; con più anni (o
    // nessun filtro) le colonne sono gli anni della selezione.
    const annoSingolo = anni.length === 1 ? anni[0] : 0;
    const all = (movimentiVista ?? []).filter(
      (m) => anni.length === 0 || anni.includes(Number(m.dataContabile.slice(0, 4))),
    );
    const anniColonna = [...new Set(all.map((m) => m.dataContabile.slice(0, 4)))]
      .filter(Boolean)
      .sort();
    const colonne = annoSingolo > 0 ? mesi : anniColonna;
    const colIdx = (m: SpMovimento): number =>
      annoSingolo > 0
        ? Number(m.dataContabile.slice(5, 7)) - 1
        : anniColonna.indexOf(m.dataContabile.slice(0, 4));
    let selezione =
      ovMode === "incassi"
        ? all.filter((m) => m.tipologia === "Incasso" && m.importo > 0)
        : ovMode === "spese"
          ? all.filter((m) => m.importo < 0)
          : all;
    if (ovMode === "spese" && ovTipF !== "tutte")
      selezione = selezione.filter((m) => m.tipologia === ovTipF);
    // Raggruppamento per chiave canonica (accorpa varianti dello stesso nome,
    // anche nei dati importati con regole più vecchie); come etichetta si
    // mostra la variante più frequente del gruppo.
    const byRiga = new Map<
      string,
      { valori: number[]; tot: number; labels: Map<string, number> }
    >();
    const totCol = colonne.map(() => 0);
    let tot = 0;
    for (const m of selezione) {
      const i = colIdx(m);
      if (i < 0 || i >= colonne.length) continue;
      const valore = ovMode === "incassi" ? m.importo : ovMode === "spese" ? -m.importo : m.importo;
      // "Per regole": righe tipologia - sottocategoria; "Per appalto":
      // righe dall'allocazione (secondaria, o primaria in mancanza).
      const label =
        ovMode === "regole"
          ? m.tipologia
            ? m.tipologia + (m.sottocategoria ? ` \u00b7 ${m.sottocategoria}` : "")
            : `(${t("fin.nonClassificato")})`
          : ovMode === "appalti"
            ? m.allocSecondaria || m.allocPrimaria || `(${t("fin.nonAllocato")})`
            : m.cliente ||
              (ovMode === "spese" && m.tipologia ? m.tipologia : `(${t("fin.unknownClient")})`);
      const key =
        ovMode === "regole" || ovMode === "appalti"
          ? label
          : m.cliente
            ? clienteGroupKey(m.cliente) || label
            : label;
      const row = byRiga.get(key) ?? {
        valori: colonne.map(() => 0),
        tot: 0,
        labels: new Map<string, number>(),
      };
      row.valori[i] += valore;
      row.tot += valore;
      row.labels.set(label, (row.labels.get(label) ?? 0) + 1);
      byRiga.set(key, row);
      totCol[i] += valore;
      tot += valore;
    }
    const righe = [...byRiga.values()]
      .map((r) => {
        const label = [...r.labels.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "";
        return [label, { valori: r.valori, tot: r.tot }] as const;
      })
      .sort((a, b) => b[1].tot - a[1].tot);
    const tipologieSpese = [
      ...new Set(
        all
          .filter((m) => m.importo < 0)
          .map((m) => m.tipologia)
          .filter(Boolean),
      ),
    ].sort((a, b) => a.localeCompare(b));
    return { righe, colonne, totCol, tot, count: selezione.length, tipologieSpese };
  }, [movimentiVista, ovMode, ovTipF, anni, mesi, t]);

  const esportaMovimenti = () => {
    esportaCsvFile(
      `movimenti-${anni.length ? [...anni].sort().join("-") : "tutti"}`,
      [
        "Data contabile",
        // Anno/trimestre/mese già pronti: la pivot raggruppa senza dover
        // convertire la colonna data a mano.
        "Anno",
        "Trimestre",
        "Mese",
        "Data valuta",
        "Importo",
        "Divisa",
        "Causale",
        "Tipologia",
        "Sottocategoria",
        "Allocazione primaria",
        "Allocazione secondaria",
        "Conto",
        "Cliente",
        "Nr fattura",
        "Note",
        "Descrizione",
        "Da verificare",
      ],
      filtrati.map((m) => [
        csvData(m.dataContabile),
        ...csvPeriodo(m.dataContabile),
        csvData(m.dataValuta),
        csvNum(m.importo),
        m.divisa,
        m.causale,
        m.tipologia,
        m.sottocategoria,
        m.allocPrimaria,
        m.allocSecondaria,
        m.conto,
        m.cliente,
        m.nrFattura,
        m.note,
        m.descrizione,
        m.daVerificare ? "Sì" : "No",
      ]),
    );
  };
  const esportaOverview = () => {
    esportaCsvFile(
      `overview-${ovMode}-${anni.length ? [...anni].sort().join("-") : "tutti"}`,
      [ovMode === "incassi" ? "Cliente" : "Controparte / Tipologia", ...overview.colonne, "Totale"],
      [
        ...overview.righe.map(([riga, r]) => [
          riga,
          ...r.valori.map((v) => (v ? csvNum(v) : "")),
          csvNum(r.tot),
        ]),
        [
          t("common.total"),
          ...overview.totCol.map((v) => (v ? csvNum(v) : "")),
          csvNum(overview.tot),
        ],
      ],
    );
  };

  // --- Render ---------------------------------------------------------------
  if (session && !isDirettore) {
    return (
      <AppShell title={t("fin.title")}>
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">{t("common.restricted")}</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">{t("fin.restrictedMsg")}</p>
          </div>
        </div>
      </AppShell>
    );
  }

  const tabBtn = (id: Tab, icon: ReactNode, label: string, badge?: number) => (
    <button
      type="button"
      onClick={() => setTab(id)}
      className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === id ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
    >
      {icon} {label}
      {badge != null && badge > 0 && (
        <span
          className={`rounded-full px-1.5 text-[10px] font-semibold ${tab === id ? "bg-primary-foreground/20" : "bg-status-absent/15 text-status-absent"}`}
        >
          {badge}
        </span>
      )}
    </button>
  );

  return (
    <AppShell title={t("fin.title")} subtitle={t("fin.subtitle")} wide>
      <div className="mb-4 flex flex-wrap items-center gap-3">
        <div className="inline-flex flex-wrap rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
          {tabBtn("movimenti", <Table2 className="h-4 w-4" />, t("fin.tabMovimenti"))}
          {tabBtn("overview", <TrendingUp className="h-4 w-4" />, t("fin.tabOverview"))}
          {tabBtn("resoconto", <Users className="h-4 w-4" />, t("fin.tabResoconto"))}
          {tabBtn("attive", <Receipt className="h-4 w-4" />, t("fin.tabAttive"))}
          {tabBtn("passive", <ReceiptText className="h-4 w-4" />, t("fin.tabPassive"))}
          {tabBtn(
            "anomalie",
            <AlertTriangle className="h-4 w-4" />,
            t("fin.tabAnomalie"),
            anomalie?.length ?? 0,
          )}
          {tabBtn("storico", <History className="h-4 w-4" />, t("fin.tabStorico"))}
          {tabBtn("regole", <GraduationCap className="h-4 w-4" />, t("fin.tabRegole"))}
        </div>
        {(tab === "movimenti" || tab === "overview") && (
          <MultiSelect
            label={t("ft.anno")}
            tuttiLabel={t("fin.allYears")}
            selLabel={t("fin.msSel")}
            opzioni={Array.from({ length: 6 }, (_, i) => new Date().getFullYear() - i).map((a) => ({
              v: a,
              label: String(a),
            }))}
            valori={anni}
            onChange={cambiaAnni}
            className="w-44"
          />
        )}
      </div>

      {/* ------------------------------- Resoconto ------------------------- */}
      {tab === "resoconto" && <ResocontoTab />}

      {/* ------------------------------- Movimenti ------------------------- */}
      {tab === "movimenti" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-end gap-3 mb-4">
            <MultiSelect
              label={t("common.type")}
              tuttiLabel={t("common.allF")}
              selLabel={t("fin.msSel")}
              opzioni={tipologiePresenti.map((tp) => ({ v: tp, label: tp }))}
              valori={tipiF}
              onChange={setTipiF}
              className="w-48"
            />
            <MultiSelect
              label={t("fin.month")}
              tuttiLabel={t("common.all")}
              selLabel={t("fin.msSel")}
              opzioni={mesi.map((m, i) => ({ v: i + 1, label: m }))}
              valori={mesiF}
              onChange={setMesiF}
              className="w-40"
            />
            <MultiSelect
              label={t("fin.cliForn")}
              tuttiLabel={t("common.allF")}
              selLabel={t("fin.msSel")}
              opzioni={clientiTop}
              valori={clientiF}
              onChange={setClientiF}
              className="w-56"
            />
            <MultiSelect
              label={t("fin.sottocat")}
              tuttiLabel={t("common.allF")}
              selLabel={t("fin.msSel")}
              opzioni={opzClassifica.sottocategorie}
              valori={sottF}
              onChange={setSottF}
              className="w-48"
            />
            <MultiSelect
              label={t("fin.allocPri")}
              tuttiLabel={t("common.allF")}
              selLabel={t("fin.msSel")}
              opzioni={opzClassifica.allocPrimarie}
              valori={allocPriF}
              onChange={setAllocPriF}
              className="w-44"
            />
            <MultiSelect
              label={t("fin.allocSec")}
              tuttiLabel={t("common.allF")}
              selLabel={t("fin.msSel")}
              opzioni={opzClassifica.allocSecondarie}
              valori={allocSecF}
              onChange={setAllocSecF}
              className="w-44"
            />
            <div className="flex-1 min-w-48">
              <label className="text-xs text-muted-foreground">{t("fin.conto")}</label>
              <select
                value={contoF}
                onChange={(e) => setContoF(e.target.value)}
                className={inputCls}
              >
                <option value="">{t("common.allF")}</option>
                <option value="__vuoto__">{t("fin.contoNonAssegnato")}</option>
                {[...new Set((movimenti ?? []).map((m) => m.conto).filter(Boolean))]
                  .sort()
                  .map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
              </select>
            </div>
            <div>
              <label className="text-xs text-muted-foreground">{t("fin.search")}</label>
              <input
                value={cercaF}
                onChange={(e) => setCercaF(e.target.value)}
                placeholder={t("fin.searchPh")}
                className={inputCls}
              />
            </div>
            <div className="self-end">
              <button
                type="button"
                onClick={() => setSoloDistinte((x) => !x)}
                className={`inline-flex items-center gap-1.5 rounded-full border px-3 py-2 text-sm font-medium transition-colors ${
                  soloDistinte
                    ? "border-primary bg-primary text-primary-foreground"
                    : "border-border bg-background text-foreground hover:bg-muted"
                }`}
              >
                <Users className="h-4 w-4" /> {t("fin.soloDistinte")}
              </button>
            </div>
            <div>
              <label className="text-xs text-muted-foreground">{t("fin.rangeDate")}</label>
              <div className="flex items-center gap-1">
                <input
                  type="date"
                  value={dataDaF}
                  onChange={(e) => setDataDaF(e.target.value)}
                  className={inputCls}
                />
                <span className="text-xs text-muted-foreground">→</span>
                <input
                  type="date"
                  value={dataAF}
                  onChange={(e) => setDataAF(e.target.value)}
                  className={inputCls}
                />
              </div>
            </div>
            <div>
              <label className="text-xs text-muted-foreground">{t("fin.rangeImporti")}</label>
              <div className="flex items-center gap-1">
                <input
                  value={impMinF}
                  onChange={(e) => setImpMinF(e.target.value)}
                  placeholder={t("fin.rangeMin")}
                  className="w-24 rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
                <span className="text-xs text-muted-foreground">→</span>
                <input
                  value={impMaxF}
                  onChange={(e) => setImpMaxF(e.target.value)}
                  placeholder={t("fin.rangeMax")}
                  className="w-24 rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
              </div>
            </div>
            <button
              type="button"
              onClick={aggiorna}
              disabled={movimenti == null || ebSyncBusy}
              className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
            >
              <RefreshCw
                className={`h-4 w-4 ${movimenti == null || ebSyncBusy ? "animate-spin" : ""}`}
              />{" "}
              {t("fin.aggiorna")}
              {ebSyncBusy && ebProgress && ` (${ebProgress})`}
            </button>
            <button
              type="button"
              onClick={() => setShowImportEC((v) => !v)}
              className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
            >
              <Upload className="h-4 w-4" /> {t("fin.tabImport")}
            </button>
            <button
              type="button"
              onClick={esportaMovimenti}
              disabled={filtrati.length === 0}
              className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
            >
              <Download className="h-4 w-4" /> {t("common.exportCsv")}
            </button>
            {/* Il totale del filtrato, sempre sott'occhio (richiesta del
                direttore): entrate e uscite separate, saldo in grande. */}
            <div className="rounded-xl border border-border bg-secondary/40 px-4 py-2">
              <div className="text-[11px] text-muted-foreground">
                {t("fin.totFiltrato")} · {filtrati.length}
              </div>
              <div
                className={`text-lg font-semibold tabular-nums ${totaleFiltrato.totale >= 0 ? "text-status-present" : "text-status-absent"}`}
              >
                {fmtImporto(totaleFiltrato.totale)}
              </div>
              <div className="text-[11px] tabular-nums text-muted-foreground">
                +{fmtImporto(totaleFiltrato.entrate)} / {fmtImporto(totaleFiltrato.uscite)}
              </div>
            </div>
            {selMov.size > 0 && (
              <div className="mb-2 flex flex-wrap items-center gap-3 rounded-lg bg-primary/10 px-3 py-2 text-sm">
                <b>{selMov.size}</b> {t("fin.selN")}
                <button
                  type="button"
                  onClick={() => setBulkOpen(true)}
                  className="rounded-lg bg-primary px-3 py-1 text-sm font-medium text-primary-foreground"
                >
                  {t("fin.selCorreggi")}
                </button>
                <button
                  type="button"
                  onClick={() => setSelMov(new Set())}
                  className="rounded-lg border border-border px-3 py-1 text-sm hover:bg-muted"
                >
                  {t("fin.selDeselez")}
                </button>
              </div>
            )}
            {bulkOpen && (
              <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
                <div className="w-full max-w-lg rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-elegant)]">
                  <div className="mb-1 text-[15px] font-semibold text-foreground">
                    {t("fin.selCorreggi")} ({selMov.size})
                  </div>
                  <p className="mb-3 text-xs text-muted-foreground">{t("fin.selVuotoNonCambia")}</p>
                  <div className="grid gap-3 sm:grid-cols-2">
                    <div>
                      <CampoVocabolario
                        label={t("common.type")}
                        valore={bulkTip}
                        onChange={cambiaBulkTip}
                        opzioni={vocabBulk.tipologie}
                        testoNessuno={t("fin.selNonCambiare")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.sottocat")}
                        valore={bulkSott}
                        onChange={setBulkSott}
                        opzioni={vocabBulk.sottocat}
                        testoNessuno={t("fin.selNonCambiare")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.allocPri")}
                        valore={bulkPri}
                        onChange={cambiaBulkPri}
                        opzioni={vocabBulk.allocPri}
                        testoNessuno={t("fin.selNonCambiare")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.allocSec")}
                        valore={bulkSec}
                        onChange={setBulkSec}
                        opzioni={vocabBulk.allocSec}
                        testoNessuno={t("fin.selNonCambiare")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                  </div>
                  <div className="mt-4 flex justify-end gap-2">
                    <button
                      type="button"
                      onClick={() => setBulkOpen(false)}
                      className="rounded-lg border border-border px-4 py-2 text-sm hover:bg-muted"
                    >
                      {t("common.cancel")}
                    </button>
                    <button
                      type="button"
                      disabled={bulkBusy}
                      onClick={() => void salvaBulk()}
                      className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
                    >
                      {bulkBusy ? t("common.loading") : t("common.save")}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {/* Dentro il pagamento cumulativo: l'elenco delle disposizioni
                della distinta agganciata (somma uguale, data vicina). */}
            {distModal && (
              <div
                className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
                onClick={() => setDistModal(null)}
              >
                <div
                  className="max-h-[80vh] w-full max-w-lg overflow-auto rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-elegant)]"
                  onClick={(e) => e.stopPropagation()}
                >
                  <div className="mb-1 text-[15px] font-semibold text-foreground">
                    {t("fin.distModalTitolo")}
                  </div>
                  <p className="mb-3 text-xs text-muted-foreground">
                    {fmtData(distModal.data)} · {distModal.tipo || "—"} · {distModal.righe.length}{" "}
                    {t("fin.distRighe")} · {t("fin.distTotale")} {fmtImporto(distModal.somma)} €
                  </p>
                  <datalist id="appalti-distinta">
                    {[
                      ...new Set([
                        ...(rosterDip ?? []).map((r) => r.appalto ?? ""),
                        ...(distinte ?? []).map((d3) => d3.appalto ?? ""),
                      ]),
                    ]
                      .filter(Boolean)
                      .sort()
                      .map((a) => (
                        <option key={a} value={a} />
                      ))}
                  </datalist>
                  <table className="w-full text-sm">
                    <tbody>
                      {[...distModal.righe]
                        .sort((a, b) => b.importo - a.importo)
                        .map((d2) => (
                          <tr key={d2.id} className="border-b border-border/40">
                            <td className="py-1 pr-2">{d2.beneficiario}</td>
                            <td className="py-1 pr-2 text-right tabular-nums whitespace-nowrap">
                              {fmtImporto(d2.importo)} €
                            </td>
                            <td className="py-1 pr-2 whitespace-nowrap">
                              {(() => {
                                const ris = risolviAppaltoDist(d2);
                                if (ris.nome)
                                  return (
                                    <span className="text-[11px] text-muted-foreground">
                                      {ris.appalto === "__senza__"
                                        ? t("fin.distSenzaApp")
                                        : ris.appalto}
                                    </span>
                                  );
                                // NON riconosciuto (es. ex dipendente fuori
                                // anagrafica): appalto assegnabile A MANO,
                                // ricordato per nome sulle distinte future.
                                return (
                                  <input
                                    list="appalti-distinta"
                                    defaultValue={ris.appalto}
                                    placeholder={t("fin.distAppaltoPh")}
                                    onBlur={(e) => {
                                      const val = e.target.value.trim();
                                      if (val === (d2.appalto ?? "")) return;
                                      void spSetDistintaAppalto({
                                        data: { id: d2.id, appalto: val },
                                      })
                                        .then(() => {
                                          setDistinte((prev) =>
                                            (prev ?? []).map((x) =>
                                              x.id === d2.id
                                                ? { ...x, appalto: val || undefined }
                                                : x,
                                            ),
                                          );
                                          toast.success(t("fin.distAppaltoOk"));
                                        })
                                        .catch((err) =>
                                          toast.error(t("common.error"), {
                                            description:
                                              err instanceof Error ? err.message : String(err),
                                          }),
                                        );
                                    }}
                                    className="w-36 rounded border border-border bg-background px-1.5 py-0.5 text-[11px]"
                                  />
                                );
                              })()}
                            </td>
                            <td
                              className="max-w-40 truncate py-1 text-[11px] text-muted-foreground"
                              title={d2.descrizione}
                            >
                              {d2.descrizione}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                  {distSpaccato && (
                    <div className="mt-3 border-t border-border/60 pt-2">
                      <div className="mb-1 text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                        {t("fin.distPerAppalto")}
                      </div>
                      <table className="w-full text-sm">
                        <tbody>
                          {distSpaccato.righe.map((r) => (
                            <tr key={r.app} className="border-b border-border/40">
                              <td className="py-1 pr-2">
                                {r.app === "__senza__" ? t("fin.distSenzaApp") : r.app}
                              </td>
                              <td className="py-1 pr-2 text-right text-[11px] text-muted-foreground">
                                {r.n}
                              </td>
                              <td className="py-1 text-right tabular-nums whitespace-nowrap font-medium">
                                {fmtImporto(r.somma)} €
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {distSpaccato.nonRic.length > 0 && (
                        <p
                          className="mt-1.5 text-[11px] text-muted-foreground"
                          title={distSpaccato.nonRic.join(", ")}
                        >
                          {t("fin.distNonRic")}: {distSpaccato.nonRic.length} —{" "}
                          {distSpaccato.nonRic.slice(0, 4).join(", ")}
                          {distSpaccato.nonRic.length > 4 ? "…" : ""}
                        </p>
                      )}
                    </div>
                  )}
                  <div className="mt-3 flex justify-end gap-2">
                    <button
                      type="button"
                      onClick={esportaDistintaCsv}
                      className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground"
                    >
                      {t("fin.distEsporta")}
                    </button>
                    <button
                      type="button"
                      onClick={() => setDistModal(null)}
                      className="rounded-lg border border-border px-4 py-2 text-sm hover:bg-muted"
                    >
                      {t("common.close")}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {/* CORREZIONE del singolo movimento (matita): stessi campi della
                sanatura — svuotare un campo e salvare = cancellare il valore. */}
            {editId != null && (
              <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
                <div className="w-full max-w-lg rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-elegant)]">
                  <div className="mb-3 text-[15px] font-semibold text-foreground">
                    {t("fin.editMovTitolo")}
                  </div>
                  <div className="grid gap-3 sm:grid-cols-2">
                    <div>
                      <CampoVocabolario
                        label={t("common.type")}
                        valore={editTip}
                        onChange={cambiaEditTip}
                        opzioni={vocabEdit.tipologie}
                        testoNessuno={t("fin.vuota")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.sottocat")}
                        valore={editSott}
                        onChange={setEditSott}
                        opzioni={vocabEdit.sottocat}
                        testoNessuno={t("fin.vuota")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.allocPri")}
                        valore={editAllocPri}
                        onChange={cambiaEditAllocPri}
                        opzioni={vocabEdit.allocPri}
                        testoNessuno={t("fin.vuota")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <CampoVocabolario
                        label={t("fin.allocSec")}
                        valore={editAllocSec}
                        onChange={setEditAllocSec}
                        opzioni={vocabEdit.allocSec}
                        testoNessuno={t("fin.vuota")}
                        testoNuova={t("fin.vocNuova")}
                      />
                    </div>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("fin.cliForn")}</label>
                      <input
                        value={editCliente}
                        onChange={(e) => setEditCliente(e.target.value)}
                        className={inputCls}
                      />
                    </div>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("fin.nrFattura")}</label>
                      <input
                        value={editNrFatt}
                        onChange={(e) => setEditNrFatt(e.target.value)}
                        className={inputCls}
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="text-xs text-muted-foreground">{t("fin.note")}</label>
                      <input
                        value={editNote}
                        onChange={(e) => setEditNote(e.target.value)}
                        className={inputCls}
                      />
                    </div>
                  </div>
                  <p className="mt-2 text-[11px] text-muted-foreground">{t("fin.editMovSvuota")}</p>
                  {(() => {
                    const m2 = (movimenti ?? []).find((x) => x.id === editId);
                    const colpite = m2 ? spiegaRegole(m2) : [];
                    if (!m2) return null;
                    return (
                      <div className="mt-2 rounded-lg bg-muted/40 p-2 text-[11px]">
                        <div className="mb-1 font-semibold text-foreground">
                          {t("fin.spiegaTitolo")}
                        </div>
                        {colpite.length === 0 ? (
                          <p className="text-muted-foreground">{t("fin.spiegaNessuna")}</p>
                        ) : (
                          colpite.map(({ r, termine }) => (
                            <p key={r.id} className="text-muted-foreground">
                              «<b className="text-foreground">{termine}</b>» ({r.campo} · {r.modo})
                              → {r.tipologia ?? ""}
                              {r.sottocategoria ? ` · ${r.sottocategoria}` : ""}
                              {r.allocSecondaria ? ` · ${r.allocSecondaria}` : ""}
                            </p>
                          ))
                        )}
                      </div>
                    );
                  })()}
                  <div className="mt-4 flex items-center justify-end gap-2">
                    {/* Eliminazione CHIRURGICA: per righe corrotte (importi
                        x100 da import sbagliati). Doppia conferma, con
                        l'importo nel testo per non sbagliare riga. */}
                    <button
                      type="button"
                      disabled={saving}
                      onClick={() => {
                        const m2 = (movimenti ?? []).find((x) => x.id === editId);
                        if (!m2) return;
                        if (
                          !window.confirm(
                            `${t("fin.eliminaMovConfirm")}

${fmtData(m2.dataContabile)} · ${fmtImporto(m2.importo)} € · ${m2.descrizione.slice(0, 60)}`,
                          )
                        )
                          return;
                        setSaving(true);
                        void spEliminaMovimento({ data: { id: m2.id } })
                          .then(() => {
                            setMovimenti((prev) => (prev ?? []).filter((x) => x.id !== m2.id));
                            setEditId(null);
                            toast.success(t("fin.eliminaMovOk"));
                          })
                          .catch((err) =>
                            toast.error(t("common.error"), {
                              description: err instanceof Error ? err.message : String(err),
                            }),
                          )
                          .finally(() => setSaving(false));
                      }}
                      className="mr-auto rounded-lg border border-status-absent/40 px-4 py-2 text-sm text-status-absent hover:bg-status-absent/10 disabled:opacity-50"
                    >
                      {t("fin.eliminaMovBtn")}
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditId(null)}
                      className="rounded-lg border border-border px-4 py-2 text-sm hover:bg-muted"
                    >
                      {t("common.cancel")}
                    </button>
                    <button
                      type="button"
                      disabled={saving}
                      onClick={() => void salvaEdit()}
                      className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
                    >
                      {saving ? t("common.loading") : t("common.save")}
                    </button>
                  </div>
                </div>
              </div>
            )}
            {ebSaldoInfo && (
              <div className="ml-auto rounded-xl border border-border bg-secondary/40 px-4 py-2 text-right">
                <div className="text-[11px] text-muted-foreground">
                  {/* BPM via PSD2 espone SOLO il saldo contabile (verificato
                      sul campo, 01/08): il titolo lo dice senza ambiguità.
                      Se un giorno la banca esponesse anche il disponibile,
                      la riga sotto comparirebbe da sola. */}
                  {t("fin.saldoContabileTitolo")}
                  {ebSaldoInfo.riferimento && ` · ${fmtData(ebSaldoInfo.riferimento)}`}
                </div>
                <div
                  className={`text-lg font-semibold tabular-nums ${ebSaldoInfo.saldo >= 0 ? "text-status-present" : "text-status-absent"}`}
                >
                  {fmtImporto(ebSaldoInfo.saldo)} {ebSaldoInfo.divisa}
                </div>
                {ebSaldoInfo.disponibile != null && (
                  <div className="text-[11px] text-muted-foreground">
                    {t("fin.saldoDisponibile")}:{" "}
                    <b className="tabular-nums">{fmtImporto(ebSaldoInfo.disponibile)}</b>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Import estratto conto (a scomparsa) */}
          {showImportEC && (
            <div className="mb-4 rounded-xl border border-border p-4">
              <div className="text-sm font-semibold text-foreground mb-1">
                {t("fin.importTitle")}
              </div>
              <p className="text-xs text-muted-foreground mb-4">{t("fin.importDesc")}</p>
              <input
                type="file"
                accept=".xlsx,.xls"
                disabled={parsing || importing}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) void onFile(f);
                  e.target.value = "";
                }}
                className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
              />
              {parsing && (
                <p className="mt-3 text-sm text-muted-foreground inline-flex items-center gap-2">
                  <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
                </p>
              )}
              {sheetChoice && (
                <div className="mt-4 rounded-xl border border-border p-4">
                  <div className="text-sm font-medium text-foreground">{sheetChoice.fileName}</div>
                  <p className="mt-1 text-[13px] text-muted-foreground">{t("fin.sheetChoose")}</p>
                  <div className="mt-3 flex flex-wrap items-end gap-3">
                    <div className="min-w-64">
                      <label className="text-xs text-muted-foreground">{t("fin.sheet")}</label>
                      <select
                        value={sheetChoice.selected}
                        onChange={(e) =>
                          setSheetChoice({ ...sheetChoice, selected: e.target.value })
                        }
                        className={inputCls}
                      >
                        {sheetChoice.sheets.map((s) => (
                          <option key={s.name} value={s.name}>
                            {s.name}
                            {s.res
                              ? ` (${s.res.rows.length} ${t("fin.rows")})`
                              : ` — ${t("fin.sheetNotRecognized")}`}
                          </option>
                        ))}
                      </select>
                    </div>
                    <button
                      type="button"
                      onClick={() => void confermaFoglio()}
                      className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90"
                    >
                      {t("fin.sheetUse")}
                    </button>
                    <button
                      type="button"
                      onClick={() => setSheetChoice(null)}
                      className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                    >
                      {t("common.cancel")}
                    </button>
                  </div>
                </div>
              )}
              {preview && (
                <div className="mt-4 rounded-xl border border-border p-4">
                  <div className="text-sm font-medium text-foreground">{preview.fileName}</div>
                  <ul className="mt-2 text-[13px] text-muted-foreground space-y-1">
                    <li>
                      {t("fin.previewPeriod")}: {fmtData(preview.dal)} → {fmtData(preview.al)}
                    </li>
                    <li>
                      {t("fin.previewTotal")}: <b>{preview.righe.length}</b>
                      {preview.scartate > 0 && ` (${preview.scartate} ${t("fin.previewSkipped")})`}
                    </li>
                    <li>
                      {t("fin.previewNew")}:{" "}
                      <b className="text-status-present">{preview.nuove.length}</b>
                    </li>
                    <li>
                      {t("fin.previewDup")}: <b>{preview.doppioni}</b>
                    </li>
                    <li>
                      {t("fin.previewAnomalie")}:{" "}
                      <b className={preview.anomalie ? "text-status-absent" : ""}>
                        {preview.anomalie}
                      </b>
                    </li>
                    {preview.correzioni.length > 0 && (
                      <li className="text-status-absent">
                        {t("fin.previewCorrezioni")}: <b>{preview.correzioni.length}</b> —{" "}
                        {t("fin.previewCorrezioniDesc")}
                      </li>
                    )}
                  </ul>
                  {preview.correzioni.length > 0 && (
                    <div className="mt-2 max-h-44 overflow-y-auto rounded-lg border border-status-absent/30 bg-status-absent/5 p-2">
                      <table className="w-full text-[11px]">
                        <tbody>
                          {preview.correzioni.map((c) => (
                            <tr key={c.id} className="border-t border-border/30">
                              <td className="py-0.5 pr-2 whitespace-nowrap">
                                {fmtData(c.dataContabile)}
                              </td>
                              <td className="py-0.5 pr-2 text-right tabular-nums whitespace-nowrap text-status-absent">
                                {fmtImporto(c.da)} €
                              </td>
                              <td className="py-0.5 pr-2 whitespace-nowrap">→</td>
                              <td className="py-0.5 pr-2 text-right font-medium tabular-nums whitespace-nowrap text-status-present">
                                {fmtImporto(c.a)} €
                              </td>
                              <td className="max-w-72 truncate py-0.5" title={c.descrizione}>
                                {c.descrizione}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  <div className="mt-3 flex items-center gap-3">
                    {preview.correzioni.length > 0 && (
                      <button
                        type="button"
                        disabled={importing}
                        onClick={() => {
                          void (async () => {
                            if (
                              !window.confirm(
                                `${t("fin.correggiConfirm")} (${preview.correzioni.length})`,
                              )
                            )
                              return;
                            setImporting(true);
                            try {
                              let fatte = 0;
                              for (const c of preview.correzioni) {
                                await spCorreggiMovimento({
                                  data: { id: c.id, importo: c.a, chiave: c.chiave },
                                });
                                fatte++;
                                setImportProgress(`${fatte} / ${preview.correzioni.length}`);
                              }
                              setMovimenti((prev) =>
                                (prev ?? []).map((x) => {
                                  const c = preview.correzioni.find((y) => y.id === x.id);
                                  return c ? { ...x, importo: c.a, chiave: c.chiave } : x;
                                }),
                              );
                              setPreview((prev) => (prev ? { ...prev, correzioni: [] } : prev));
                              toast.success(t("fin.correggiOk"), { description: `${fatte}` });
                            } catch (err) {
                              toast.error(t("common.error"), {
                                description: err instanceof Error ? err.message : String(err),
                              });
                            } finally {
                              setImporting(false);
                              setImportProgress("");
                            }
                          })();
                        }}
                        className="inline-flex items-center gap-2 rounded-lg bg-status-absent px-4 py-2 text-sm font-medium text-white hover:opacity-90 disabled:opacity-50"
                      >
                        {t("fin.correggiBtn")} ({preview.correzioni.length})
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={eseguiImport}
                      disabled={importing || preview.nuove.length === 0}
                      className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                    >
                      {importing ? (
                        <Loader2 className="h-4 w-4 animate-spin" />
                      ) : (
                        <Upload className="h-4 w-4" />
                      )}
                      {importing
                        ? `${t("fin.importing")} ${importProgress}`
                        : `${t("fin.importBtn")} (${preview.nuove.length})`}
                    </button>
                    {!importing && (
                      <button
                        type="button"
                        onClick={() => setPreview(null)}
                        className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                      >
                        {t("common.cancel")}
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}

          {movimenti == null ? (
            <div className="py-10 text-center text-sm text-muted-foreground">
              <Loader2 className="h-5 w-5 animate-spin inline-block" />
            </div>
          ) : filtrati.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground">{t("fin.empty")}</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  {/* Ogni colonna ha il suo imbuto (l'ordine dei th deve
                      combaciare con le celle del corpo). Il Saldo e' un
                      progressivo calcolato: niente filtro. */}
                  <tr className="text-left text-xs text-muted-foreground border-b border-border">
                    <th className="py-2 pr-2">
                      <input
                        type="checkbox"
                        className="accent-primary"
                        checked={
                          filtrati
                            .slice(inizioMov, inizioMov + righePagina)
                            .every((m) => selMov.has(m.id)) && filtrati.length > 0
                        }
                        onChange={(e) => {
                          const ns = new Set(selMov);
                          for (const m of filtrati.slice(inizioMov, inizioMov + righePagina))
                            if (e.target.checked) ns.add(m.id);
                            else ns.delete(m.id);
                          setSelMov(ns);
                        }}
                      />
                    </th>
                    {thFiltroMov(movColTh[0])}
                    {thFiltroMov(movColTh[1])}
                    {thFiltroMov(movColTh[2], "text-right")}
                    {ebSaldoInfo && <th className="py-2 pr-3 text-right">{t("fin.saldo")}</th>}
                    {thFiltroMov(movColTh[3])}
                    {thFiltroMov(movColTh[4])}
                    {thFiltroMov(movColTh[5])}
                    {thFiltroMov(movColTh[6])}
                    {thFiltroMov(movColTh[7])}
                    <th className="py-2" />
                  </tr>
                </thead>
                <tbody>
                  {filtrati.slice(inizioMov, inizioMov + righePagina).map((m) => (
                    <tr
                      key={m.id}
                      className="border-b border-border/50 hover:bg-muted/40"
                      title={m.descrizione}
                    >
                      <td className="py-1.5 pr-2">
                        {!m.distVirtuale && (
                          <input
                            type="checkbox"
                            className="accent-primary"
                            checked={selMov.has(m.id)}
                            onChange={() => {
                              const ns = new Set(selMov);
                              if (ns.has(m.id)) ns.delete(m.id);
                              else ns.add(m.id);
                              setSelMov(ns);
                            }}
                            onClick={(e) => e.stopPropagation()}
                          />
                        )}
                      </td>
                      <td className="py-1.5 pr-3 whitespace-nowrap">{fmtData(m.dataContabile)}</td>
                      <td className="py-1.5 pr-3 whitespace-nowrap text-muted-foreground">
                        {fmtData(m.dataValuta)}
                      </td>
                      <td
                        className={`py-1.5 pr-3 text-right font-medium whitespace-nowrap ${m.importo > 0 ? "text-status-present" : "text-foreground"}`}
                      >
                        {fmtImporto(m.importo)}
                      </td>
                      {ebSaldoInfo && (
                        <td className="py-1.5 pr-3 text-right whitespace-nowrap text-muted-foreground tabular-nums">
                          {fmtImporto(saldoDopo(m))}
                        </td>
                      )}
                      <td className="py-1.5 pr-3 max-w-64">
                        {/* La causale VERA e' il testo del movimento; il
                            codice ABI (usato dal classificatore) resta
                            sotto, in piccolo. */}
                        <div className="truncate text-muted-foreground" title={m.descrizione}>
                          {m.descrizione || "—"}
                        </div>
                        {m.causale && (
                          <div className="text-[10px] tabular-nums text-muted-foreground/60">
                            ABI {m.causale}
                          </div>
                        )}
                        {(() => {
                          const g = m.distChiave
                            ? distGruppi.find((x) => `${x.data}|${x.tipo}` === m.distChiave)
                            : distintaDi(m);
                          return g ? (
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setDistModal(g);
                              }}
                              title={t("fin.distBadgeTip")}
                              className="mt-0.5 inline-flex items-center gap-1 rounded-full bg-primary/10 px-2 py-0.5 text-[10px] font-medium text-primary hover:bg-primary/20"
                            >
                              <Users className="h-3 w-3" /> {g.righe.length}
                            </button>
                          ) : null;
                        })()}
                      </td>
                      <td className="py-1.5 pr-3">
                        {m.tipologia}
                        {m.sottocategoria && (
                          <span className="text-[11px] text-muted-foreground">
                            {" "}
                            · {m.sottocategoria}
                          </span>
                        )}
                        {(m.allocPrimaria || m.allocSecondaria) && (
                          <div className="text-[11px] text-muted-foreground">
                            {[m.allocPrimaria, m.allocSecondaria].filter(Boolean).join(" / ")}
                          </div>
                        )}
                        {m.daVerificare && (
                          <AlertTriangle className="h-3.5 w-3.5 inline-block ml-1 text-status-absent" />
                        )}
                      </td>
                      <td className="py-1.5 pr-3">
                        {m.cliente || "—"}
                        {/* La ricerca guarda anche la DESCRIZIONE (che in
                            tabella non c'è): quando è l'unico punto in cui
                            il testo cercato compare, se ne mostra uno
                            stralcio — sennò la riga sembra un falso
                            positivo (successo con "romano" → utenze di
                            Fiano Romano). */}
                        {cercaF.trim() &&
                          !m.cliente.toLowerCase().includes(cercaF.trim().toLowerCase()) &&
                          m.descrizione.toLowerCase().includes(cercaF.trim().toLowerCase()) && (
                            <span
                              className="ml-1 text-[11px] italic text-muted-foreground"
                              title={m.descrizione}
                            >
                              «{stralcioDescr(m.descrizione, cercaF)}»
                            </span>
                          )}
                      </td>
                      <td className="py-1.5 pr-3 text-muted-foreground">{m.nrFattura || "—"}</td>
                      <td className="py-1.5 pr-3 text-muted-foreground">{m.note || "—"}</td>
                      <td className="py-1.5 text-right whitespace-nowrap">
                        {!m.distVirtuale && (
                          <button
                            type="button"
                            onClick={() => apriEdit(m)}
                            title={t("fin.editMovTip")}
                            className="rounded-md p-1 text-muted-foreground hover:text-foreground hover:bg-muted"
                          >
                            <Pencil className="h-3.5 w-3.5" />
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => creaRegolaDa(m)}
                          title={t("fin.creaRegolaTip")}
                          className="rounded-md p-1 text-muted-foreground hover:text-primary hover:bg-primary/10"
                        >
                          <Wand2 className="h-3.5 w-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className="mt-2 flex flex-wrap items-center justify-between gap-2 text-xs text-muted-foreground">
                <span>
                  {filtrati.length <= RIGHE_PAGINA
                    ? `${filtrati.length} ${t("fin.rows")}`
                    : `${inizioMov + 1}–${Math.min(inizioMov + righePagina, filtrati.length)} ${t("fin.pageOf")} ${filtrati.length} ${t("fin.rows")}`}
                </span>
                <select
                  value={righePagina}
                  onChange={(e) => {
                    setRighePagina(Number(e.target.value));
                    setPaginaMov(1);
                  }}
                  title={t("fin.perPagina")}
                  className="rounded-lg border border-border bg-background px-2 py-1 text-xs"
                >
                  {[50, 100, 200, 500, 1000].map((nr) => (
                    <option key={nr} value={nr}>
                      {nr} / {t("fin.page").toLowerCase()}
                    </option>
                  ))}
                </select>
                {pagineMovTot > 1 && (
                  <span className="inline-flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setPaginaMov(pagMov - 1)}
                      disabled={pagMov <= 1}
                      className="rounded-lg border border-border p-1.5 text-foreground hover:bg-muted disabled:opacity-40"
                      aria-label="‹"
                    >
                      <ChevronLeft className="h-3.5 w-3.5" />
                    </button>
                    <span>
                      {t("fin.page")} {pagMov} {t("fin.pageOf")} {pagineMovTot}
                    </span>
                    <button
                      type="button"
                      onClick={() => setPaginaMov(pagMov + 1)}
                      disabled={pagMov >= pagineMovTot}
                      className="rounded-lg border border-border p-1.5 text-foreground hover:bg-muted disabled:opacity-40"
                      aria-label="›"
                    >
                      <ChevronRight className="h-3.5 w-3.5" />
                    </button>
                  </span>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ------------------------------- Overview -------------------------- */}
      {tab === "overview" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div>
              <div className="flex flex-wrap items-center gap-2 mb-2">
                <div className="inline-flex rounded-lg border border-border p-0.5 text-sm">
                  <button
                    type="button"
                    onClick={() => setOvMode("incassi")}
                    className={`rounded-md px-3 py-1 font-medium ${ovMode === "incassi" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
                  >
                    {t("fin.ovIncassi")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setOvMode("spese")}
                    className={`rounded-md px-3 py-1 font-medium ${ovMode === "spese" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
                  >
                    {t("fin.ovSpese")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setOvMode("regole")}
                    className={`rounded-md px-3 py-1 font-medium ${ovMode === "regole" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
                  >
                    {t("fin.ovRegole")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setOvMode("appalti")}
                    className={`rounded-md px-3 py-1 font-medium ${ovMode === "appalti" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
                  >
                    {t("fin.ovAppalti")}
                  </button>
                </div>
                {ovMode === "spese" && (
                  <select
                    value={ovTipF}
                    onChange={(e) => setOvTipF(e.target.value)}
                    className={`${inputCls} w-auto`}
                  >
                    <option value="tutte">{t("common.allF")}</option>
                    {overview.tipologieSpese.map((tp) => (
                      <option key={tp} value={tp}>
                        {tp}
                      </option>
                    ))}
                  </select>
                )}
              </div>
              <p className="text-xs text-muted-foreground">
                {overview.count}{" "}
                {ovMode === "incassi" ? t("fin.overviewCount") : t("fin.overviewCountSpese")} ·{" "}
                {t("common.total")}{" "}
                <span
                  className={`font-semibold ${ovMode === "incassi" ? "text-status-present" : "text-foreground"}`}
                >
                  {fmtImporto(overview.tot)} €
                </span>
              </p>
            </div>
            <button
              type="button"
              onClick={esportaOverview}
              disabled={overview.righe.length === 0}
              className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
            >
              <Download className="h-4 w-4" /> {t("common.exportCsv")}
            </button>
          </div>
          {movimenti == null ? (
            <div className="py-10 text-center text-sm text-muted-foreground">
              <Loader2 className="h-5 w-5 animate-spin inline-block" />
            </div>
          ) : overview.righe.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground">{t("fin.empty")}</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">
                      {ovMode === "incassi" ? t("fin.cliente") : t("fin.controparte")}
                    </th>
                    {overview.colonne.map((c) => (
                      <th key={c} className="py-2 px-2 text-right">
                        {c}
                      </th>
                    ))}
                    <th className="py-2 pl-2 text-right">{t("common.total")}</th>
                  </tr>
                </thead>
                <tbody>
                  {overview.righe.map(([riga, r]) => (
                    <tr key={riga} className="border-b border-border/50 hover:bg-muted/40">
                      <td className="py-1.5 pr-3 max-w-56 truncate" title={riga}>
                        {riga}
                      </td>
                      {r.valori.map((v, i) => (
                        <td
                          key={i}
                          className="py-1.5 px-2 text-right whitespace-nowrap text-muted-foreground"
                        >
                          {v ? fmtImporto(v) : ""}
                        </td>
                      ))}
                      <td className="py-1.5 pl-2 text-right font-semibold whitespace-nowrap">
                        {fmtImporto(r.tot)}
                      </td>
                    </tr>
                  ))}
                  <tr className="border-t border-border font-semibold">
                    <td className="py-2 pr-3">{t("common.total")}</td>
                    {overview.totCol.map((v, i) => (
                      <td key={i} className="py-2 px-2 text-right whitespace-nowrap">
                        {v ? fmtImporto(v) : ""}
                      </td>
                    ))}
                    <td
                      className={`py-2 pl-2 text-right whitespace-nowrap ${ovMode === "incassi" ? "text-status-present" : ""}`}
                    >
                      {fmtImporto(overview.tot)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ------------------------------- Fatture --------------------------- */}
      {tab === "attive" && <FattureTab direzione="Emessa" />}
      {tab === "passive" && <FattureTab direzione="Ricevuta" />}

      {/* ------------------------------- Anomalie -------------------------- */}
      {tab === "anomalie" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="text-sm font-semibold text-foreground mb-1">{t("fin.anomalieTitle")}</div>
          <p className="text-xs text-muted-foreground mb-4">{t("fin.anomalieDesc")}</p>
          {anomalie == null ? (
            <div className="py-10 text-center text-sm text-muted-foreground">
              <Loader2 className="h-5 w-5 animate-spin inline-block" />
            </div>
          ) : anomalie.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground inline-flex items-center gap-2 w-full justify-center">
              <CheckCircle2 className="h-4 w-4 text-status-present" /> {t("fin.noAnomalie")}
            </p>
          ) : (
            <ul className="space-y-3">
              {anomalie.map((m) => (
                <li key={m.id} className="rounded-xl border border-border p-4">
                  <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-sm">
                    <span className="whitespace-nowrap">{fmtData(m.dataContabile)}</span>
                    <span
                      className={`font-semibold whitespace-nowrap ${m.importo > 0 ? "text-status-present" : ""}`}
                    >
                      {fmtImporto(m.importo)} {m.divisa}
                    </span>
                    <span className="text-xs rounded-full bg-muted px-2 py-0.5">
                      {m.tipologia || "—"}
                    </span>
                    <span className="text-xs text-muted-foreground">
                      {t("fin.causale")} {m.causale || "—"}
                    </span>
                  </div>
                  <p className="mt-1 text-[13px] text-muted-foreground break-all">
                    {m.descrizione}
                  </p>
                  {editId === m.id ? (
                    <div className="mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                      <div>
                        <label className="text-xs text-muted-foreground">{t("common.type")}</label>
                        <input
                          list="tipologie-sanatura"
                          value={editTip}
                          onChange={(e) => setEditTip(e.target.value)}
                          className={inputCls}
                        />
                        <datalist id="tipologie-sanatura">
                          {[
                            ...new Set([
                              ...TIPOLOGIE_MOVIMENTO,
                              ...(movimenti ?? []).map((x) => x.tipologia).filter(Boolean),
                            ]),
                          ]
                            .sort((a, b) => a.localeCompare(b))
                            .map((tp) => (
                              <option key={tp} value={tp} />
                            ))}
                        </datalist>
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.sottocat")}</label>
                        <input
                          list="sottocategorie-note"
                          value={editSott}
                          onChange={(e) => setEditSott(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.allocPri")}</label>
                        <input
                          list="alloc-primarie"
                          value={editAllocPri}
                          onChange={(e) => setEditAllocPri(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.allocSec")}</label>
                        <input
                          list="alloc-secondarie"
                          value={editAllocSec}
                          onChange={(e) => setEditAllocSec(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.cliForn")}</label>
                        <input
                          value={editCliente}
                          onChange={(e) => setEditCliente(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">
                          {t("fin.nrFattura")}
                        </label>
                        <input
                          value={editNrFatt}
                          onChange={(e) => setEditNrFatt(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.note")}</label>
                        <input
                          value={editNote}
                          onChange={(e) => setEditNote(e.target.value)}
                          className={inputCls}
                        />
                      </div>
                      <div className="sm:col-span-2 lg:col-span-4 flex gap-2">
                        <button
                          type="button"
                          onClick={salvaEdit}
                          disabled={saving}
                          className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                        >
                          {saving ? (
                            <Loader2 className="h-4 w-4 animate-spin" />
                          ) : (
                            <CheckCircle2 className="h-4 w-4" />
                          )}
                          {t("fin.fixConfirm")}
                        </button>
                        <button
                          type="button"
                          onClick={() => setEditId(null)}
                          className="rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                        >
                          {t("common.cancel")}
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div className="mt-3 flex flex-wrap gap-2">
                      <button
                        type="button"
                        onClick={() => apriEdit(m)}
                        className="rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                      >
                        {t("fin.fix")}
                      </button>
                      <button
                        type="button"
                        onClick={() => creaRegolaDa(m)}
                        className="inline-flex items-center gap-1.5 rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                      >
                        <Wand2 className="h-3.5 w-3.5" /> {t("fin.creaRegolaTip")}
                      </button>
                    </div>
                  )}
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

      {/* ------------------------------- Storico import -------------------- */}
      <ScorriFondo />
      {tab === "storico" &&
        (() => {
          // GEMELLI x100: un vecchio import ha letto la virgola dei decimali
          // come separatore delle migliaia (5.315,55 -> 531555). La riga
          // corrotta e' sempre INTERA, vale 100 volte una riga vera con i
          // centesimi, stessa data e stesso riferimento/descrizione.
          const rifDi = (m: SpMovimento) => {
            const r = /rif\.?\s*([a-z0-9/]+)/i.exec(m.descrizione);
            return r
              ? r[1].toLowerCase()
              : m.descrizione.toLowerCase().replace(/\s+/g, " ").slice(0, 60);
          };
          const gruppi = new Map<string, SpMovimento[]>();
          for (const m of movimenti ?? []) {
            const k = `${m.dataContabile}|${rifDi(m)}`;
            const arr = gruppi.get(k) ?? [];
            arr.push(m);
            gruppi.set(k, arr);
          }
          const corrotti: SpMovimento[] = [];
          for (const arr of gruppi.values()) {
            if (arr.length < 2) continue;
            for (const x of arr) {
              if (!Number.isInteger(x.importo)) continue;
              const gemello = arr.find(
                (y) =>
                  y.id !== x.id &&
                  x.importo * y.importo > 0 &&
                  Math.abs(y.importo % 1) > 0.001 &&
                  Math.round(Math.abs(y.importo) * 100) === Math.abs(x.importo),
              );
              if (gemello) corrotti.push(x);
            }
          }
          if (!corrotti.length) return null;
          const totale = Math.round(corrotti.reduce((s2, m) => s2 + m.importo, 0) * 100) / 100;
          return (
            <div className="mb-4 rounded-2xl border border-status-absent/40 bg-status-absent/5 p-5">
              <div className="mb-1 text-sm font-semibold text-status-absent">
                {t("fin.bonificaTitolo")} ({corrotti.length})
              </div>
              <p className="mb-2 text-xs text-muted-foreground">{t("fin.bonificaDesc")}</p>
              <table className="w-full text-[12px]">
                <tbody>
                  {corrotti.slice(0, 20).map((m) => (
                    <tr key={m.id} className="border-t border-border/40">
                      <td className="py-1 pr-3 whitespace-nowrap">{fmtData(m.dataContabile)}</td>
                      <td className="py-1 pr-3 text-right font-medium tabular-nums whitespace-nowrap text-status-absent">
                        {fmtImporto(m.importo)} €
                      </td>
                      <td
                        className="max-w-96 truncate py-1 pr-3 text-muted-foreground"
                        title={m.descrizione}
                      >
                        {m.descrizione}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button
                type="button"
                disabled={bonificaBusy}
                onClick={() => {
                  void (async () => {
                    if (
                      !window.confirm(
                        `${t("fin.bonificaConfirm")} (${corrotti.length} · ${fmtImporto(totale)} €)`,
                      )
                    )
                      return;
                    setBonificaBusy(true);
                    try {
                      let fatti = 0;
                      for (const m of corrotti) {
                        await spEliminaMovimento({ data: { id: m.id } });
                        fatti++;
                      }
                      setMovimenti((prev) =>
                        (prev ?? []).filter((x) => !corrotti.some((c) => c.id === x.id)),
                      );
                      toast.success(t("fin.bonificaOk"), { description: `${fatti}` });
                    } catch (err) {
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      });
                    } finally {
                      setBonificaBusy(false);
                    }
                  })();
                }}
                className="mt-3 rounded-lg bg-status-absent px-4 py-2 text-sm font-medium text-white hover:opacity-90 disabled:opacity-50"
              >
                {bonificaBusy
                  ? t("common.loading")
                  : `${t("fin.bonificaBtn")} (${corrotti.length})`}
              </button>
            </div>
          );
        })()}
      {tab === "storico" && distinteCard}
      {tab === "storico" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="text-sm font-semibold text-foreground mb-1">{t("fin.storicoTitle")}</div>
          <p className="text-xs text-muted-foreground mb-4">{t("fin.storicoDesc")}</p>
          {storico == null ? (
            <div className="py-10 text-center text-sm text-muted-foreground">
              <Loader2 className="h-5 w-5 animate-spin inline-block" />
            </div>
          ) : storico.length === 0 ? (
            <p className="py-8 text-center text-sm text-muted-foreground">
              {t("fin.storicoEmpty")}
            </p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">{t("fin.colImport")}</th>
                    <th className="py-2 pr-3">{t("fin.previewPeriod")}</th>
                    <th className="py-2 pr-3 text-right">{t("fin.colMovimenti")}</th>
                    <th className="py-2 pr-3 text-right">{t("fin.tabAnomalie")}</th>
                    <th className="py-2 pr-3 text-right">{t("fin.colSaldo")}</th>
                    <th className="py-2 pr-3" />
                  </tr>
                </thead>
                <tbody>
                  {storico.map((r) => (
                    <tr key={r.importId || "legacy"} className="border-b border-border/50">
                      <td className="py-2 pr-3 whitespace-nowrap">
                        {fmtImportId(r.importId, t("fin.legacyImport"))}
                        {/* CONTO del lotto: badge se assegnato, campo+bottone
                            per assegnarlo (a blocchi, tutto il lotto). */}
                        <div className="mt-1 flex items-center gap-1.5">
                          {r.conto && r.conto !== "misto" ? (
                            <span className="rounded-full bg-primary/10 px-2 py-0.5 text-[11px] font-medium text-primary">
                              {r.conto}
                            </span>
                          ) : (
                            <>
                              <input
                                list="conti-noti"
                                value={contoLotto[r.importId] ?? ""}
                                onChange={(e) =>
                                  setContoLotto((prev) => ({
                                    ...prev,
                                    [r.importId]: e.target.value,
                                  }))
                                }
                                placeholder={t("fin.contoPh")}
                                className="w-32 rounded-md border border-border bg-background px-2 py-1 text-[11px]"
                              />
                              <button
                                type="button"
                                disabled={
                                  contoBusy != null || !(contoLotto[r.importId] ?? "").trim()
                                }
                                onClick={() => {
                                  const conto = (contoLotto[r.importId] ?? "").trim();
                                  setContoBusy(r.importId);
                                  void (async () => {
                                    try {
                                      for (;;) {
                                        const esito = (await spAssegnaContoLotto({
                                          data: { importId: r.importId, conto },
                                        })) as { aggiornati: number; rimanenti: number };
                                        if (esito.rimanenti <= 0 || esito.aggiornati === 0) break;
                                      }
                                      toast.success(t("fin.contoAssegnato"));
                                      loadStorico();
                                      loadMovimenti(anni);
                                    } catch (err) {
                                      toast.error(t("common.error"), {
                                        description:
                                          err instanceof Error ? err.message : String(err),
                                      });
                                    } finally {
                                      setContoBusy(null);
                                    }
                                  })();
                                }}
                                className="rounded-md border border-border px-2 py-1 text-[11px] hover:bg-muted disabled:opacity-50"
                              >
                                {contoBusy === r.importId ? "…" : t("fin.contoAssegna")}
                              </button>
                            </>
                          )}
                          {r.conto === "misto" && (
                            <span className="text-[11px] text-status-absent">misto</span>
                          )}
                        </div>
                        <datalist id="conti-noti">
                          {[
                            ...new Set([
                              "BPM 3681",
                              "Qonto",
                              ...(movimenti ?? []).map((m) => m.conto).filter(Boolean),
                            ]),
                          ].map((c) => (
                            <option key={c} value={c} />
                          ))}
                        </datalist>
                      </td>
                      <td className="py-2 pr-3 whitespace-nowrap text-muted-foreground">
                        {fmtData(r.dal)} → {fmtData(r.al)}
                      </td>
                      <td className="py-2 pr-3 text-right">{r.movimenti}</td>
                      <td className="py-2 pr-3 text-right">
                        {r.anomalie > 0 ? (
                          <span className="text-status-absent font-medium">{r.anomalie}</span>
                        ) : (
                          "0"
                        )}
                      </td>
                      <td
                        className={`py-2 pr-3 text-right whitespace-nowrap font-medium ${r.totale > 0 ? "text-status-present" : ""}`}
                      >
                        {fmtImporto(r.totale)}
                      </td>
                      <td className="py-2 text-right">
                        <button
                          type="button"
                          onClick={() => void annullaImportGruppo(r)}
                          disabled={annullaBusy != null}
                          className="inline-flex items-center gap-1.5 rounded-lg border border-status-absent/40 px-3 py-1.5 text-xs font-medium text-status-absent hover:bg-status-absent/10 disabled:opacity-50"
                        >
                          {annullaBusy === r.importId ? (
                            <>
                              <Loader2 className="h-3.5 w-3.5 animate-spin" />
                              {annullaProgress} / {r.movimenti}
                            </>
                          ) : (
                            <>
                              <Trash2 className="h-3.5 w-3.5" /> {t("fin.annulla")}
                            </>
                          )}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ------------------------------- Regole apprese -------------------- */}
      {tab === "regole" && (
        <div className="space-y-4">
          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="text-sm font-semibold text-foreground mb-1">{t("fin.regoleTitle")}</div>
            <p className="text-xs text-muted-foreground mb-4">{t("fin.regoleDesc")}</p>
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              <div className="sm:col-span-2">
                <label className="text-xs text-muted-foreground">{t("fin.regolaPattern")}</label>
                {/* Area di testo: il direttore incolla ELENCHI di nominativi
                    (virgola, punto e virgola o a capo = termini alternativi). */}
                <textarea
                  value={rPattern}
                  onChange={(e) => setRPattern(e.target.value)}
                  placeholder={t("fin.regolaPatternPh")}
                  rows={2}
                  className={`${inputCls} min-h-[42px] resize-y`}
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaCampo")}</label>
                <select
                  value={rCampo}
                  onChange={(e) =>
                    setRCampo(e.target.value as "cliente" | "descrizione" | "entrambi")
                  }
                  className={inputCls}
                >
                  <option value="cliente">{t("fin.campoCliente")}</option>
                  <option value="descrizione">{t("fin.campoDescrizione")}</option>
                  <option value="entrambi">{t("fin.campoEntrambi")}</option>
                </select>
              </div>
              {rCampo !== "descrizione" && (
                <div>
                  <label className="text-xs text-muted-foreground">{t("fin.regolaModo")}</label>
                  <select
                    value={rModo}
                    onChange={(e) => setRModo(e.target.value as "esatto" | "contiene")}
                    className={inputCls}
                  >
                    <option value="esatto">{t("fin.modoEsatto")}</option>
                    <option value="contiene">{t("fin.modoContiene")}</option>
                  </select>
                </div>
              )}
              <div>
                <CampoVocabolario
                  label={t("fin.regolaTipologia")}
                  valore={rTipologia}
                  onChange={setRTipologia}
                  opzioni={vocab.tipologie}
                  testoNessuno={t("fin.regolaTipNoChange")}
                  testoNuova={t("fin.vocNuova")}
                />
              </div>
              <div>
                <CampoVocabolario
                  label={t("fin.regolaSottocat")}
                  valore={rSottocat}
                  onChange={setRSottocat}
                  opzioni={vocab.sottocat}
                  testoNessuno={t("fin.regolaTipNoChange")}
                  testoNuova={t("fin.vocNuova")}
                />
              </div>
              <div>
                <CampoVocabolario
                  label={t("fin.regolaAllocPri")}
                  valore={rAllocPri}
                  onChange={setRAllocPri}
                  opzioni={vocab.allocPri}
                  testoNessuno={t("fin.regolaTipNoChange")}
                  testoNuova={t("fin.vocNuova")}
                />
              </div>
              <div>
                <CampoVocabolario
                  label={t("fin.regolaAllocSec")}
                  valore={rAllocSec}
                  onChange={setRAllocSec}
                  opzioni={vocab.allocSec}
                  testoNessuno={t("fin.regolaTipNoChange")}
                  testoNuova={t("fin.vocNuova")}
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaSegno")}</label>
                <select
                  value={rSegno}
                  onChange={(e) => setRSegno(e.target.value as "" | "entrate" | "uscite")}
                  className={inputCls}
                >
                  <option value="">{t("fin.regolaSegnoTutti")}</option>
                  <option value="entrate">{t("fin.regolaSegnoEntrate")}</option>
                  <option value="uscite">{t("fin.regolaSegnoUscite")}</option>
                </select>
              </div>
              <div className="sm:col-span-2">
                <label className="text-xs text-muted-foreground">{t("fin.regolaNote")}</label>
                <input
                  value={rNote}
                  onChange={(e) => setRNote(e.target.value)}
                  placeholder={t("fin.regolaNotePh")}
                  className={inputCls}
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">
                  {t("fin.regolaClienteNuovo")}
                </label>
                <input
                  value={rCliente}
                  onChange={(e) => setRCliente(e.target.value)}
                  placeholder={t("fin.regolaClienteNuovoPh")}
                  className={inputCls}
                />
              </div>
              <label className="flex items-end gap-2 pb-2 text-sm text-foreground">
                <input
                  type="checkbox"
                  checked={rApplica}
                  onChange={(e) => setRApplica(e.target.checked)}
                  className="h-4 w-4 accent-primary"
                />
                {t("fin.regolaApplicaEsistenti")}
              </label>
            </div>
            {rSorgente && !rEditId && (
              <button
                type="button"
                disabled={
                  rBusy ||
                  (!rTipologia.trim() &&
                    !rSottocat.trim() &&
                    !rAllocPri.trim() &&
                    !rAllocSec.trim() &&
                    !rCliente.trim())
                }
                onClick={() => {
                  void (async () => {
                    setRBusy(true);
                    try {
                      // La classificazione RESTA nelle categorie del movimento
                      // anche senza salvare la regola (richiesta direzione).
                      await spUpdateMovimento({
                        data: {
                          movimentoId: rSorgente,
                          ...(rTipologia.trim()
                            ? { tipologia: rTipologia.trim(), daVerificare: false }
                            : {}),
                          ...(rSottocat.trim() ? { sottocategoria: rSottocat.trim() } : {}),
                          ...(rAllocPri.trim() ? { allocPrimaria: rAllocPri.trim() } : {}),
                          ...(rAllocSec.trim() ? { allocSecondaria: rAllocSec.trim() } : {}),
                          ...(rCliente.trim() ? { cliente: rCliente.trim() } : {}),
                        },
                      });
                      toast.success(t("fin.soloMovFatto"));
                      setRSorgente(null);
                      loadMovimenti(anni);
                    } catch (err) {
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      });
                    } finally {
                      setRBusy(false);
                    }
                  })();
                }}
                className="mt-4 mr-2 inline-flex items-center gap-2 rounded-lg border border-border px-4 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
              >
                {t("fin.soloMovBtn")}
              </button>
            )}
            <button
              type="button"
              onClick={() => void submitRegola()}
              disabled={rBusy || !rPattern.trim() || (!rTipologia.trim() && !rCliente.trim())}
              className="mt-4 inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
            >
              {rBusy ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  {rApplica ? `${t("fin.regolaApplying")} ${rProgress}` : t("common.loading")}
                </>
              ) : (
                <>
                  <GraduationCap className="h-4 w-4" />{" "}
                  {rEditId ? t("fin.regolaAggiorna") : t("fin.regolaCrea")}
                </>
              )}
            </button>
          </div>

          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="mb-3 flex items-center gap-3">
              <span className="text-sm font-semibold text-foreground">
                {t("fin.regoleElencoTitle")}
              </span>
              {/* REGOLA UNICA DIPENDENTI: bonifici in uscita verso nomi che
                  combaciano col roster (troncati/invertiti compresi) →
                  Pagamento Salario col nome pulito. Retroattiva a blocchi. */}
              <button
                type="button"
                disabled={dipBusy}
                onClick={() => {
                  if (!window.confirm(t("fin.regDipConfirm"))) return;
                  setDipBusy(true);
                  void (async () => {
                    let tot = 0;
                    try {
                      for (;;) {
                        const r = (await spApplicaRegolaDipendenti()) as {
                          aggiornati: number;
                          rimanenti: number;
                        };
                        tot += r.aggiornati;
                        if (r.rimanenti <= 0 || r.aggiornati === 0) break;
                      }
                      toast.success(`${tot} ${t("fin.regolaApplicati")}`);
                      loadMovimenti(anni);
                      loadAnomalie();
                    } catch (err) {
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      });
                    } finally {
                      setDipBusy(false);
                    }
                  })();
                }}
                className="ml-auto rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {dipBusy ? t("fin.regolaApplying") : t("fin.regDipBtn")}
              </button>
              {/* RIAPPLICA TUTTE: dopo aver creato colonne mancanti (o dopo
                  un pasticcio) un click ripassa ogni regola sull'archivio. */}
              <button
                type="button"
                disabled={dipBusy || (regole ?? []).length === 0}
                onClick={() => {
                  if (!window.confirm(t("fin.riapplicaConfirm"))) return;
                  setDipBusy(true);
                  void (async () => {
                    let tot = 0;
                    let errori = 0;
                    try {
                      for (const r of regole ?? []) {
                        const payload = {
                          pattern: r.pattern,
                          campo: r.campo,
                          modo: r.modo,
                          tipologia: r.tipologia,
                          sottocategoria: r.sottocategoria,
                          allocPrimaria: r.allocPrimaria,
                          allocSecondaria: r.allocSecondaria,
                          cliente: r.cliente,
                        };
                        let ultimoRimanenti = Number.POSITIVE_INFINITY;
                        try {
                          for (;;) {
                            const esito = (await spApplicaRegolaFinanza({ data: payload })) as {
                              aggiornati: number;
                              rimanenti: number;
                            };
                            tot += esito.aggiornati;
                            if (esito.rimanenti <= 0 || esito.aggiornati === 0) break;
                            if (esito.rimanenti >= ultimoRimanenti) break;
                            ultimoRimanenti = esito.rimanenti;
                          }
                        } catch {
                          errori++;
                        }
                      }
                      toast.success(
                        `${tot} ${t("fin.regolaApplicati")}${errori ? ` · ${errori} regole con errori` : ""}`,
                      );
                      loadMovimenti(anni);
                      loadAnomalie();
                    } finally {
                      setDipBusy(false);
                    }
                  })();
                }}
                className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {t("fin.riapplicaBtn")}
              </button>
              <button
                type="button"
                disabled={uniBusy || regole == null}
                onClick={() => void unificaDoppie()}
                className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {uniBusy ? t("common.loading") : t("fin.uniBtn")}
              </button>
              <button
                type="button"
                disabled={uniBusy || regole == null}
                onClick={() => {
                  void (async () => {
                    // MIGRAZIONE UNA TANTUM: le regole storiche (spese,
                    // salari) valgono solo per le USCITE — cosi' un bonifico
                    // in ENTRATA non puo' piu' finire "Pagamento Salario".
                    const senza = (regole ?? []).filter(
                      (r) => !r.segno && r.pattern.trim() !== "*",
                    );
                    if (!senza.length) {
                      toast.success(t("fin.segnoNiente"));
                      return;
                    }
                    if (!window.confirm(`${t("fin.segnoConfirm")} (${senza.length})`)) return;
                    setUniBusy(true);
                    try {
                      for (const r of senza)
                        await spUpdateRegolaFinanza({
                          data: {
                            regolaId: r.id ?? "",
                            pattern: r.pattern,
                            campo: r.campo,
                            modo: r.modo,
                            tipologia: r.tipologia,
                            sottocategoria: r.sottocategoria,
                            allocPrimaria: r.allocPrimaria,
                            allocSecondaria: r.allocSecondaria,
                            cliente: r.cliente,
                            note: r.note,
                            segno: "uscite",
                          },
                        });
                      const agg = (await spGetRegoleFinanza()) as RegolaFinanza[];
                      setRegole(agg);
                      toast.success(t("fin.segnoFatto"), { description: `${senza.length}` });
                    } catch (err) {
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      });
                    } finally {
                      setUniBusy(false);
                    }
                  })();
                }}
                className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {t("fin.segnoBtn")}
              </button>
              <button
                type="button"
                disabled={regole == null || regole.length === 0}
                onClick={() => {
                  // Excel-friendly: CSV con BOM e punto e virgola.
                  esportaCsvFile(
                    "regole-movimenti",
                    [
                      "Testo da riconoscere",
                      "Campo",
                      "Modo",
                      "Tipologia",
                      "Sottocategoria",
                      "Allocazione primaria",
                      "Allocazione secondaria",
                      "Nome controparte",
                      "Nota",
                    ],
                    (regole ?? []).map((r) => [
                      r.pattern,
                      r.campo,
                      r.modo,
                      r.tipologia ?? "",
                      r.sottocategoria ?? "",
                      r.allocPrimaria ?? "",
                      r.allocSecondaria ?? "",
                      r.cliente ?? "",
                      r.note ?? "",
                    ]),
                  );
                }}
                className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {t("fin.regoleEsporta")}
              </button>
              {selReg.size > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    void (async () => {
                      if (!window.confirm(`${t("fin.selEliminaConfirm")} (${selReg.size})`)) return;
                      try {
                        for (const rid of selReg)
                          await spDeleteRegolaFinanza({ data: { regolaId: rid } });
                        toast.success(t("fin.selFatto"), { description: `${selReg.size}` });
                        setSelReg(new Set());
                        loadRegole();
                      } catch (err) {
                        toast.error(t("common.error"), {
                          description: err instanceof Error ? err.message : String(err),
                        });
                      }
                    })();
                  }}
                  className="rounded-lg border border-status-absent/40 px-3 py-1.5 text-xs text-status-absent hover:bg-status-absent/10"
                >
                  {t("fin.selElimina")} ({selReg.size})
                </button>
              )}
            </div>
            {/* Ricerca ISTANTANEA: mentre si scrive restano solo le regole
                col termine (nei pattern o negli esiti) e le categorie si
                aprono da sole sul match. */}
            <input
              value={rCerca}
              onChange={(e) => setRCerca(e.target.value)}
              placeholder={t("fin.regoleCercaPh")}
              className="mb-3 w-full max-w-md rounded-lg border border-border bg-background px-3 py-2 text-sm"
            />
            {regole == null ? (
              <div className="py-6 text-center text-sm text-muted-foreground">
                <Loader2 className="h-5 w-5 animate-spin inline-block" />
              </div>
            ) : regole.length === 0 ? (
              <p className="py-6 text-center text-sm text-muted-foreground">
                {t("fin.regoleEmpty")}
              </p>
            ) : (
              (() => {
                const q = rCerca.trim().toLowerCase();
                const visibili = q
                  ? regole.filter((r) =>
                      [
                        r.pattern,
                        r.tipologia ?? "",
                        r.sottocategoria ?? "",
                        r.allocPrimaria ?? "",
                        r.allocSecondaria ?? "",
                        r.cliente ?? "",
                      ]
                        .join(" ")
                        .toLowerCase()
                        .includes(q),
                    )
                  : regole;
                const perCat = new Map<string, RegolaFinanza[]>();
                for (const r of visibili) {
                  const k = r.tipologia?.trim() || t("fin.regoleAltro");
                  perCat.set(k, [...(perCat.get(k) ?? []), r]);
                }
                return [...perCat.entries()]
                  .sort((a, b) => a[0].localeCompare(b[0]))
                  .map(([cat, lista]) => (
                    <div key={cat} className="border-b border-border/60 last:border-0">
                      <button
                        type="button"
                        onClick={() => setCatAperta(catAperta === cat ? null : cat)}
                        className="flex w-full items-center gap-2 py-2 text-sm font-semibold text-foreground hover:text-primary"
                      >
                        <span
                          className={`transition-transform ${catAperta === cat ? "rotate-90" : ""}`}
                        >
                          ▸
                        </span>
                        {cat}
                        <span className="text-xs font-normal text-muted-foreground">
                          ({lista.length})
                        </span>
                      </button>
                      {(catAperta === cat || rCerca.trim() !== "") && (
                        <ul className="divide-y divide-border/40 pl-5">
                          {lista.map((r) => (
                            <li key={r.id} className="py-2.5 flex items-center gap-3 text-sm">
                              {/* Frase per esteso, come la leggerebbe una persona:
                        "Se il nome contiene «amazon» → tipologia … · nome …" */}
                              <span className="flex-1 min-w-0">
                                {t("fin.regolaFraseSe")}{" "}
                                <span className="text-muted-foreground">
                                  {r.campo === "descrizione"
                                    ? t("fin.regolaFraseDescr")
                                    : r.campo === "entrambi"
                                      ? t("fin.regolaFraseEntrambi")
                                      : t("fin.regolaFraseNome")}{" "}
                                  {r.campo !== "descrizione" && r.modo === "esatto"
                                    ? t("fin.regolaFraseUguale")
                                    : t("fin.regolaFraseContiene")}
                                </span>{" "}
                                <span className="font-medium text-foreground">«{r.pattern}»</span>
                                <span className="text-muted-foreground">
                                  {" "}
                                  {t("fin.regolaFraseAllora")}{" "}
                                </span>
                                {r.tipologia && (
                                  <span className="text-xs rounded-full bg-primary/10 text-primary px-2 py-0.5 mr-1">
                                    {t("fin.regolaFraseTip")} {r.tipologia}
                                  </span>
                                )}
                                {r.sottocategoria && (
                                  <span className="text-xs rounded-full bg-primary/10 text-primary px-2 py-0.5 mr-1">
                                    {t("fin.regolaFraseSottocat")} {r.sottocategoria}
                                  </span>
                                )}
                                {(r.allocPrimaria || r.allocSecondaria) && (
                                  <span className="text-xs rounded-full bg-muted px-2 py-0.5 mr-1">
                                    {[r.allocPrimaria, r.allocSecondaria]
                                      .filter(Boolean)
                                      .join(" / ")}
                                  </span>
                                )}
                                {r.cliente && (
                                  <span className="text-xs rounded-full bg-muted px-2 py-0.5">
                                    {t("fin.regolaFraseNomeNuovo")} {r.cliente}
                                  </span>
                                )}
                                {r.note && (
                                  <span className="text-xs italic text-muted-foreground">
                                    {" "}
                                    — {r.note}
                                  </span>
                                )}
                              </span>
                              <input
                                type="checkbox"
                                className="accent-primary mr-1 shrink-0"
                                checked={selReg.has(r.id ?? "")}
                                onChange={() => {
                                  const ns = new Set(selReg);
                                  const rid = r.id ?? "";
                                  if (ns.has(rid)) ns.delete(rid);
                                  else ns.add(rid);
                                  setSelReg(ns);
                                }}
                                onClick={(e) => e.stopPropagation()}
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  setRPattern(r.pattern);
                                  setRCampo(r.campo);
                                  setRModo(r.modo);
                                  setRTipologia(r.tipologia ?? "");
                                  setRSottocat(r.sottocategoria ?? "");
                                  setRAllocPri(r.allocPrimaria ?? "");
                                  setRAllocSec(r.allocSecondaria ?? "");
                                  setRCliente(r.cliente ?? "");
                                  setRNote(r.note ?? "");
                                  setRSegno(r.segno ?? "");
                                  setRSorgente(null);
                                  setRApplica(true);
                                  setREditId(r.id ?? null);
                                  window.scrollTo({ top: 0, behavior: "smooth" });
                                }}
                                title={t("common.edit")}
                                className="shrink-0 rounded-md p-1.5 text-muted-foreground hover:text-foreground hover:bg-muted"
                              >
                                <Pencil className="h-4 w-4" />
                              </button>
                              <button
                                type="button"
                                onClick={() => void eliminaRegola(r)}
                                title={t("fin.regolaDelete")}
                                className="shrink-0 rounded-md p-1.5 text-muted-foreground hover:text-status-absent hover:bg-status-absent/10"
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  ));
              })()
            )}
          </div>

          {/* Regole di classificazione delle fatture passive: per fornitore
              fissano tipologia e cliente di riferimento (immagine del
              direttore: Nolvex → sanzioni/franchigie, cliente NOLVEX). */}
          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="text-sm font-semibold text-foreground mb-1">
              {t("fin.rfTitle")}
              {regoleFat != null && (
                <span className="ml-1.5 text-xs font-normal text-muted-foreground">
                  ({regoleFat.length})
                </span>
              )}
            </div>
            <p className="text-xs text-muted-foreground mb-4">{t("fin.rfDesc")}</p>
            <div className="flex flex-wrap items-end gap-3 mb-4">
              <div className="flex-1 min-w-48">
                <label className="text-xs text-muted-foreground">{t("fin.rfFornitore")}</label>
                <input
                  value={rfFornitore}
                  onChange={(e) => setRfFornitore(e.target.value)}
                  className={inputCls}
                />
              </div>
              <div className="flex-1 min-w-56">
                <label className="text-xs text-muted-foreground">{t("fin.rfTipologia")}</label>
                <input
                  value={rfTipologia}
                  onChange={(e) => setRfTipologia(e.target.value)}
                  className={inputCls}
                />
              </div>
              <div className="w-44">
                <label className="text-xs text-muted-foreground">{t("fin.rfCliente")}</label>
                <input
                  value={rfCliente}
                  onChange={(e) => setRfCliente(e.target.value)}
                  className={inputCls}
                />
              </div>
              <button
                type="button"
                disabled={rfBusy}
                onClick={() => void salvaRegolaFat()}
                className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
              >
                {t("common.save")}
              </button>
            </div>
            {regoleFat == null ? (
              <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
            ) : regoleFat.length === 0 ? (
              <p className="text-sm text-muted-foreground">{t("fin.rfEmpty")}</p>
            ) : (
              <ul className="divide-y divide-border/60">
                {[...regoleFat]
                  .sort((a, b) => a.fornitore.localeCompare(b.fornitore))
                  .map((r) => (
                    <li
                      key={r.id ?? r.fornitore}
                      className="flex items-center gap-3 py-1.5 text-sm"
                    >
                      <span className="w-56 truncate font-medium">{r.fornitore}</span>
                      <span className="flex-1 truncate text-muted-foreground">
                        {r.tipologia ?? "—"}
                      </span>
                      <span className="w-40 truncate">{r.clienteRif ?? ""}</span>
                      <button
                        type="button"
                        disabled={rfBusy}
                        onClick={() => void eliminaRegolaFat(r)}
                        className="rounded-md p-1 text-muted-foreground hover:text-status-absent"
                        title={t("common.delete")}
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </li>
                  ))}
              </ul>
            )}
          </div>

          {/* Termini d'incasso per cliente: i giorni contrattuali da cui
              partono scadenze e ritardi delle fatture attive. Si impostano
              qui a mano, oppure caricando il foglio contratti nell'import
              della tab Fatture. Chi non è in elenco = 30 giorni. */}
          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="text-sm font-semibold text-foreground mb-1">
              {t("fin.termTitle")}
              {termini != null && (
                <span className="ml-1.5 text-xs font-normal text-muted-foreground">
                  ({termini.length})
                </span>
              )}
            </div>
            <p className="text-xs text-muted-foreground mb-4">{t("fin.termDesc")}</p>
            {/* Direzione: clienti (ci pagano) o fornitori (paghiamo noi). */}
            <div className="mb-3 flex flex-wrap items-center gap-1.5">
              {(
                [
                  ["Emessa", t("fin.termClientiTab")],
                  ["Ricevuta", t("fin.termFornitoriTab")],
                ] as const
              ).map(([dir2, label]) => (
                <button
                  key={dir2}
                  type="button"
                  onClick={() => setTDirezione(dir2)}
                  className={`rounded-full border px-3 py-1.5 text-xs font-medium transition-colors ${
                    tDirezione === dir2
                      ? "border-primary bg-primary text-primary-foreground"
                      : "border-border bg-background text-foreground hover:bg-muted"
                  }`}
                >
                  {label}
                </button>
              ))}
              <button
                type="button"
                disabled={tBusy}
                onClick={() => void copiaTerminiFornitori()}
                className="ml-auto rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
              >
                {t("fin.termCopiaBtn")}
              </button>
            </div>
            <div className="flex flex-wrap items-end gap-3 mb-4">
              <div className="flex-1 min-w-56">
                <label className="text-xs text-muted-foreground">{t("fin.cliente")}</label>
                <input
                  value={tCliente}
                  onChange={(e) => setTCliente(e.target.value)}
                  placeholder={t("fin.termClientePh")}
                  className={inputCls}
                />
              </div>
              <div className="w-32">
                <label className="text-xs text-muted-foreground">{t("fin.termGiorni")}</label>
                <input
                  type="number"
                  min={1}
                  value={tGiorni}
                  onChange={(e) => setTGiorni(e.target.value)}
                  placeholder="30"
                  className={inputCls}
                />
              </div>
              <div className="flex-1 min-w-56">
                <label className="text-xs text-muted-foreground">{t("fin.termOggetto")}</label>
                <input
                  value={tOggetto}
                  onChange={(e) => setTOggetto(e.target.value)}
                  placeholder={t("fin.termOggettoPh")}
                  className={inputCls}
                />
              </div>
              <div className="flex-1 min-w-56">
                <label className="text-xs text-muted-foreground">{t("fin.termEmail")}</label>
                <input
                  type="email"
                  value={tEmail}
                  onChange={(e) => setTEmail(e.target.value)}
                  placeholder={t("fin.termEmailPh")}
                  className={inputCls}
                />
              </div>
              <button
                type="button"
                disabled={tBusy}
                onClick={() => void salvaTermine()}
                className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
              >
                {t("common.save")}
              </button>
            </div>
            {termini == null ? (
              <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />
            ) : termini.length === 0 ? (
              <p className="text-sm text-muted-foreground">{t("fin.termEmpty")}</p>
            ) : (
              <ul className="divide-y divide-border/60">
                {[...termini]
                  .filter((x) => (x.direzione ?? "Emessa") === tDirezione)
                  .sort((a, b) => a.cliente.localeCompare(b.cliente))
                  .map((x) => (
                    <li
                      key={`${x.cliente}|${x.oggetto ?? ""}`}
                      className="flex items-center gap-3 py-1.5 text-sm"
                    >
                      <span className="flex-1 truncate">
                        {x.cliente}
                        {x.oggetto && (
                          <span className="ml-2 rounded-full bg-primary/10 px-2 py-0.5 text-[11px] text-primary">
                            {x.oggetto}
                          </span>
                        )}
                        {x.email && (
                          <span className="ml-2 text-xs text-muted-foreground">{x.email}</span>
                        )}
                      </span>
                      <b className="tabular-nums">
                        {x.giorni} {t("fin.termGg")}
                      </b>
                      <button
                        type="button"
                        onClick={() => {
                          setTCliente(x.cliente);
                          setTGiorni(String(x.giorni));
                          setTEmail(x.email ?? "");
                          setTOggetto(x.oggetto ?? "");
                        }}
                        className="rounded-md p-1 text-muted-foreground hover:text-foreground"
                        title={t("common.edit")}
                      >
                        <Pencil className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        disabled={tBusy}
                        onClick={() => void eliminaTermine(x.cliente, x.oggetto)}
                        className="rounded-md p-1 text-muted-foreground hover:text-status-absent"
                        title={t("common.delete")}
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </li>
                  ))}
              </ul>
            )}
          </div>
        </div>
      )}
    </AppShell>
  );
}
