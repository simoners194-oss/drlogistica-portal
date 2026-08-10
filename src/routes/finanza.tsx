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
import { ResocontoTab } from "@/components/ResocontoTab";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import { isSupervisoreGlobale } from "@/lib/richieste-logic";
import {
  parseEstratto,
  parseMatrice,
  clienteGroupKey,
  LEGACY_IMPORT_ID,
  TIPOLOGIE_MOVIMENTO,
  matchRegola,
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
  spAnnullaImport,
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
  EbStato,
  EbSyncResult,
  EbSaldoInfo,
} from "@/lib/sharepoint.server";

export const Route = createFileRoute("/finanza")({
  head: () => ({ meta: [{ title: "Finanza — DR Portal" }] }),
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
const RIGHE_PAGINA = 500;

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

interface PreviewImport {
  fileName: string;
  righe: MovimentoParsed[];
  nuove: MovimentoParsed[];
  doppioni: number;
  scartate: number;
  anomalie: number;
  dal: string;
  al: string;
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

  // Filtri archivio movimenti
  const [tipiF, setTipiF] = useState<string[]>([]);
  const [cercaF, setCercaF] = useState("");
  const [mesiF, setMesiF] = useState<number[]>([]); // vuoto = tutti
  // Clienti selezionati nel menu a tendina (vuoto = tutti). Le voci proposte
  // sono le 15 controparti con piu' incassi, in ordine decrescente.
  const [clientiF, setClientiF] = useState<string[]>([]);
  const [paginaMov, setPaginaMov] = useState(1); // pagine da RIGHE_PAGINA

  // Overview: incassi o spese (+ filtro tipologia, utile solo per le spese)
  const [ovMode, setOvMode] = useState<"incassi" | "spese">("incassi");
  const [ovTipF, setOvTipF] = useState("tutte");

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
      setPreview({
        fileName,
        righe,
        nuove,
        doppioni: righe.length - nuove.length,
        scartate: res.scartate,
        anomalie: nuove.filter((r) => r.daVerificare).length,
        dal: date[0] ?? "",
        al: date[date.length - 1] ?? "",
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
    setRPattern(m.cliente || "");
    setRCampo("cliente");
    setRModo("esatto");
    setRTipologia(m.tipologia || "");
    setRSottocat(m.sottocategoria || "");
    setRAllocPri(m.allocPrimaria || "");
    setRAllocSec(m.allocSecondaria || "");
    setRCliente("");
    setRApplica(true);
    setTab("regole");
  };

  const submitRegola = async () => {
    setRBusy(true);
    setRProgress(0);
    try {
      const payload = {
        pattern: rPattern.trim(),
        campo: rCampo,
        modo: rModo,
        tipologia: rTipologia.trim() || undefined,
        sottocategoria: rSottocat.trim() || undefined,
        allocPrimaria: rAllocPri.trim() || undefined,
        allocSecondaria: rAllocSec.trim() || undefined,
        cliente: rCliente.trim() || undefined,
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
                movimenti.filter((m) => matchRegola(m, payload)).length
              } ${t("fin.regolaImpatto2")}`,
            );
      if (!conferma) {
        setRBusy(false);
        setRProgress(0);
        return;
      }
      // Modifica = aggiornamento SUL POSTO: mai piu' cancella-e-ricrea (una
      // create fallita dopo la delete ha bruciato due regole del direttore).
      if (rEditId) await spUpdateRegolaFinanza({ data: { regolaId: rEditId, ...payload } });
      else await spCreateRegolaFinanza({ data: payload });
      let applicati = 0;
      if (rApplica) {
        // Applicazione retroattiva a blocchi finché il server non ha finito.
        // Se i RIMANENTI non calano tra un giro e l'altro, qualcosa non si
        // riesce a scrivere: ci si ferma invece di girare a vuoto.
        let ultimoRimanenti = Number.POSITIVE_INFINITY;
        for (;;) {
          const r = (await spApplicaRegolaFinanza({ data: payload })) as {
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
      toast.success(t("fin.regolaCreata"), {
        description: rApplica ? `${applicati} ${t("fin.regolaApplicati")}` : undefined,
      });
      setRPattern("");
      setRTipologia("");
      setRSottocat("");
      setRAllocPri("");
      setRAllocSec("");
      setRCliente("");
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

  const filtrati = useMemo(() => {
    let out = movimenti ?? [];
    // Anni: il server ha fornito l'intervallo min-max; qui si rifiniscono le
    // selezioni non contigue. Tipologie e mesi: vuoto = tutti, altrimenti OR.
    if (anni.length) out = out.filter((m) => anni.includes(Number(m.dataContabile.slice(0, 4))));
    if (clientiF.length)
      out = out.filter(
        (m) => m.cliente && clientiF.includes(clienteGroupKey(m.cliente) || m.cliente),
      );
    if (tipiF.length) out = out.filter((m) => tipiF.includes(m.tipologia));
    if (mesiF.length) out = out.filter((m) => mesiF.includes(Number(m.dataContabile.slice(5, 7))));
    if (contoF) out = out.filter((m) => (m.conto || "") === (contoF === "__vuoto__" ? "" : contoF));
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
  }, [movimenti, anni, tipiF, mesiF, cercaF, clientiF, contoF]);
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
  }, [tipiF, mesiF, cercaF, anni, clientiF]);
  const pagineMovTot = Math.max(1, Math.ceil(filtrati.length / RIGHE_PAGINA));
  const pagMov = Math.min(paginaMov, pagineMovTot);
  const inizioMov = (pagMov - 1) * RIGHE_PAGINA;

  const tipologiePresenti = useMemo(() => {
    const set = new Set((movimenti ?? []).map((m) => m.tipologia).filter(Boolean));
    return [...set].sort((a, b) => a.localeCompare(b));
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
    const all = (movimenti ?? []).filter(
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
        : all.filter((m) => m.importo < 0);
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
      const valore = ovMode === "incassi" ? m.importo : -m.importo;
      const label =
        m.cliente ||
        (ovMode === "spese" && m.tipologia ? m.tipologia : `(${t("fin.unknownClient")})`);
      const key = m.cliente ? clienteGroupKey(m.cliente) || label : label;
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
  }, [movimenti, ovMode, ovTipF, anni, mesi, t]);

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
                      <label className="text-xs text-muted-foreground">{t("common.type")}</label>
                      <input
                        list="tipologie-mov-edit"
                        value={editTip}
                        onChange={(e) => setEditTip(e.target.value)}
                        className={inputCls}
                      />
                      <datalist id="tipologie-mov-edit">
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
                        list="sottocat-mov-edit"
                        value={editSott}
                        onChange={(e) => setEditSott(e.target.value)}
                        className={inputCls}
                      />
                      <datalist id="sottocat-mov-edit">
                        {[
                          ...new Set(
                            (movimenti ?? []).map((x) => x.sottocategoria).filter(Boolean),
                          ),
                        ]
                          .sort((a, b) => a.localeCompare(b))
                          .map((sc) => (
                            <option key={sc} value={sc} />
                          ))}
                      </datalist>
                    </div>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("fin.allocPri")}</label>
                      <input
                        list="allocpri-mov-edit"
                        value={editAllocPri}
                        onChange={(e) => setEditAllocPri(e.target.value)}
                        className={inputCls}
                      />
                      <datalist id="allocpri-mov-edit">
                        {[
                          ...new Set([
                            "Costi generali",
                            "Appalto",
                            ...(movimenti ?? []).map((x) => x.allocPrimaria).filter(Boolean),
                          ]),
                        ]
                          .sort((a, b) => a.localeCompare(b))
                          .map((a) => (
                            <option key={a} value={a} />
                          ))}
                      </datalist>
                    </div>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("fin.allocSec")}</label>
                      <input
                        list="allocsec-mov-edit"
                        value={editAllocSec}
                        onChange={(e) => setEditAllocSec(e.target.value)}
                        className={inputCls}
                      />
                      <datalist id="allocsec-mov-edit">
                        {[
                          ...new Set(
                            (movimenti ?? []).map((x) => x.allocSecondaria).filter(Boolean),
                          ),
                        ]
                          .sort((a, b) => a.localeCompare(b))
                          .map((a) => (
                            <option key={a} value={a} />
                          ))}
                      </datalist>
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
                  <div className="mt-4 flex justify-end gap-2">
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
                  </ul>
                  <div className="mt-3 flex items-center gap-3">
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
                  <tr className="text-left text-xs text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">{t("fin.dataContabile")}</th>
                    <th className="py-2 pr-3">{t("fin.dataValuta")}</th>
                    <th className="py-2 pr-3 text-right">{t("common.amount")}</th>
                    {ebSaldoInfo && <th className="py-2 pr-3 text-right">{t("fin.saldo")}</th>}
                    <th className="py-2 pr-3">{t("fin.causaleCol")}</th>
                    <th className="py-2 pr-3">{t("common.type")}</th>
                    <th className="py-2 pr-3">{t("fin.cliForn")}</th>
                    <th className="py-2 pr-3">{t("fin.nrFattura")}</th>
                    <th className="py-2 pr-3">{t("fin.note")}</th>
                    <th className="py-2" />
                  </tr>
                </thead>
                <tbody>
                  {filtrati.slice(inizioMov, inizioMov + RIGHE_PAGINA).map((m) => (
                    <tr
                      key={m.id}
                      className="border-b border-border/50 hover:bg-muted/40"
                      title={m.descrizione}
                    >
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
                        <button
                          type="button"
                          onClick={() => apriEdit(m)}
                          title={t("fin.editMovTip")}
                          className="rounded-md p-1 text-muted-foreground hover:text-foreground hover:bg-muted"
                        >
                          <Pencil className="h-3.5 w-3.5" />
                        </button>
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
                    : `${inizioMov + 1}–${Math.min(inizioMov + RIGHE_PAGINA, filtrati.length)} ${t("fin.pageOf")} ${filtrati.length} ${t("fin.rows")}`}
                </span>
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
                <label className="text-xs text-muted-foreground">{t("fin.regolaTipologia")}</label>
                {/* Campo LIBERO con suggerimenti: si può digitare una
                    tipologia nuova (es. "Consulenza notarile") e nasce con
                    la regola — vuoto = non cambiare. */}
                <input
                  list="tipologie-regola"
                  value={rTipologia}
                  onChange={(e) => setRTipologia(e.target.value)}
                  placeholder={t("fin.regolaTipNoChange")}
                  className={inputCls}
                />
                <datalist id="tipologie-regola">
                  {[...new Set((regole ?? []).map((r) => r.tipologia ?? "").filter(Boolean))]
                    .sort((a, b) => a.localeCompare(b))
                    .map((tp) => (
                      <option key={tp} value={tp} />
                    ))}
                </datalist>
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaSottocat")}</label>
                {/* Libera con suggerimenti: le sottocategorie "nascono"
                    scrivendole (auto-aggiunta chiesta dal direttore). */}
                <input
                  list="sottocategorie-note"
                  value={rSottocat}
                  onChange={(e) => setRSottocat(e.target.value)}
                  placeholder={t("fin.regolaTipNoChange")}
                  className={inputCls}
                />
                <datalist id="sottocategorie-note">
                  {[...new Set((regole ?? []).map((r) => r.sottocategoria ?? "").filter(Boolean))]
                    .sort((a, b) => a.localeCompare(b))
                    .map((sc) => (
                      <option key={sc} value={sc} />
                    ))}
                </datalist>
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaAllocPri")}</label>
                <input
                  list="alloc-primarie"
                  value={rAllocPri}
                  onChange={(e) => setRAllocPri(e.target.value)}
                  placeholder={t("fin.regolaTipNoChange")}
                  className={inputCls}
                />
                <datalist id="alloc-primarie">
                  {[...new Set((regole ?? []).map((r) => r.allocPrimaria ?? "").filter(Boolean))]
                    .sort((a, b) => a.localeCompare(b))
                    .map((a) => (
                      <option key={a} value={a} />
                    ))}
                </datalist>
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaAllocSec")}</label>
                <input
                  list="alloc-secondarie"
                  value={rAllocSec}
                  onChange={(e) => setRAllocSec(e.target.value)}
                  placeholder={t("fin.regolaTipNoChange")}
                  className={inputCls}
                />
                <datalist id="alloc-secondarie">
                  {[...new Set((regole ?? []).map((r) => r.allocSecondaria ?? "").filter(Boolean))]
                    .sort((a, b) => a.localeCompare(b))
                    .map((a) => (
                      <option key={a} value={a} />
                    ))}
                </datalist>
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
            </div>
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
                const perCat = new Map<string, RegolaFinanza[]>();
                for (const r of regole) {
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
                      {catAperta === cat && (
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
                              </span>
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
