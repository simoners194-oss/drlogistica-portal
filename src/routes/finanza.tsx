// DR Portal — Finanza (sezione riservata al direttore DR005 + admin).
// Estratto conto bancario: import da xlsx (con scelta del foglio), archivio
// movimenti classificati, overview incassi/spese per cliente, anomalie da
// sanare a mano, storico degli import con annullamento.
import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState, type ReactNode } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import {
  Landmark,
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
  GraduationCap,
  Wand2,
} from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import { isSupervisoreGlobale } from "@/lib/richieste-logic";
import {
  parseEstratto,
  parseMatrice,
  clienteGroupKey,
  LEGACY_IMPORT_ID,
  TIPOLOGIE_MOVIMENTO,
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
  spCreateRegolaFinanza,
  spDeleteRegolaFinanza,
  spApplicaRegolaFinanza,
} from "@/lib/sharepoint.functions";
import type { SpMovimento, ImportStoricoRiga } from "@/lib/sharepoint.server";

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
// "IMP-2026-07-22T15:30:12" → "22/07/2026 15:30" (gruppo legacy → etichetta).
function fmtImportId(id: string, legacyLabel: string): string {
  if (!id) return legacyLabel;
  const m = id.match(/^IMP-(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2})/);
  return m ? `${m[3]}/${m[2]}/${m[1]} ${m[4]}` : id;
}

// Blocchi di upload verso il server (sotto il limite server di 150).
const CHUNK = 100;

type Tab = "movimenti" | "overview" | "anomalie" | "import" | "storico" | "regole";

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
  // 0 = tutti gli anni (l'overview passa da colonne-mese a colonne-anno).
  const [anno, setAnno] = useState(new Date().getFullYear());

  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [anomalie, setAnomalie] = useState<SpMovimento[] | null>(null);
  const [storico, setStorico] = useState<ImportStoricoRiga[] | null>(null);

  // Filtri archivio movimenti
  const [tipF, setTipF] = useState("tutte");
  const [cercaF, setCercaF] = useState("");
  const [meseF, setMeseF] = useState(0); // 0 = tutti

  // Overview: incassi o spese (+ filtro tipologia, utile solo per le spese)
  const [ovMode, setOvMode] = useState<"incassi" | "spese">("incassi");
  const [ovTipF, setOvTipF] = useState("tutte");

  // Import
  const [sheetChoice, setSheetChoice] = useState<SheetChoice | null>(null);
  const [preview, setPreview] = useState<PreviewImport | null>(null);
  const [parsing, setParsing] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState("");

  // Storico: annullamento in corso
  const [annullaBusy, setAnnullaBusy] = useState<string | null>(null);
  const [annullaProgress, setAnnullaProgress] = useState(0);

  // Regole apprese
  const [regole, setRegole] = useState<RegolaFinanza[] | null>(null);
  const [rPattern, setRPattern] = useState("");
  const [rCampo, setRCampo] = useState<"cliente" | "descrizione">("cliente");
  const [rModo, setRModo] = useState<"esatto" | "contiene">("esatto");
  const [rTipologia, setRTipologia] = useState("");
  const [rCliente, setRCliente] = useState("");
  const [rApplica, setRApplica] = useState(true);
  const [rBusy, setRBusy] = useState(false);
  const [rProgress, setRProgress] = useState(0);

  // Sanatura anomalie
  const [editId, setEditId] = useState<string | null>(null);
  const [editTip, setEditTip] = useState("");
  const [editCliente, setEditCliente] = useState("");
  const [editNrFatt, setEditNrFatt] = useState("");
  const [editNote, setEditNote] = useState("");
  const [saving, setSaving] = useState(false);

  const isDirettore =
    session != null &&
    (session.ruolo === "amministratore_sistema" || isSupervisoreGlobale(session.codice ?? ""));

  const loadMovimenti = (a: number) => {
    setMovimenti(null);
    const range = a > 0 ? { from: `${a}-01-01`, to: `${a}-12-31` } : {};
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
  };
  const refreshAll = (a: number) => {
    loadMovimenti(a);
    loadAnomalie();
    loadStorico();
    loadRegole();
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
    refreshAll(anno);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const cambiaAnno = (a: number) => {
    setAnno(a);
    loadMovimenti(a);
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
      refreshAll(anno);
      setTab("movimenti");
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
      refreshAll(anno);
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
      refreshAll(anno);
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
        cliente: rCliente.trim() || undefined,
      };
      await spCreateRegolaFinanza({ data: payload });
      let applicati = 0;
      if (rApplica) {
        // Applicazione retroattiva a blocchi finché il server non ha finito.
        for (;;) {
          const r = (await spApplicaRegolaFinanza({ data: payload })) as {
            aggiornati: number;
            rimanenti: number;
          };
          applicati += r.aggiornati;
          setRProgress(applicati);
          if (r.rimanenti <= 0) break;
          if (r.aggiornati === 0) break; // safety: niente progresso
        }
      }
      toast.success(t("fin.regolaCreata"), {
        description: rApplica ? `${applicati} ${t("fin.regolaApplicati")}` : undefined,
      });
      setRPattern("");
      setRTipologia("");
      setRCliente("");
      loadRegole();
      if (rApplica) {
        loadMovimenti(anno);
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
    try {
      await spDeleteRegolaFinanza({ data: { regolaId: r.id } });
      setRegole((prev) => (prev ?? []).filter((x) => x.id !== r.id));
      toast.success(t("fin.regolaDeleted"));
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
  const filtrati = useMemo(() => {
    let out = movimenti ?? [];
    if (tipF !== "tutte") out = out.filter((m) => m.tipologia === tipF);
    if (meseF > 0) out = out.filter((m) => Number(m.dataContabile.slice(5, 7)) === meseF);
    if (cercaF.trim()) {
      const q = cercaF.trim().toLowerCase();
      out = out.filter(
        (m) =>
          m.cliente.toLowerCase().includes(q) ||
          m.descrizione.toLowerCase().includes(q) ||
          m.nrFattura.toLowerCase().includes(q) ||
          m.note.toLowerCase().includes(q),
      );
    }
    return out;
  }, [movimenti, tipF, meseF, cercaF]);

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
    const all = movimenti ?? [];
    const anni = [...new Set(all.map((m) => m.dataContabile.slice(0, 4)))].filter(Boolean).sort();
    const colonne = anno > 0 ? mesi : anni;
    const colIdx = (m: SpMovimento): number =>
      anno > 0
        ? Number(m.dataContabile.slice(5, 7)) - 1
        : anni.indexOf(m.dataContabile.slice(0, 4));
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
  }, [movimenti, ovMode, ovTipF, anno, mesi, t]);

  const esportaMovimenti = () => {
    esportaCsvFile(
      `movimenti-${anno > 0 ? anno : "tutti"}`,
      [
        "Data contabile",
        "Data valuta",
        "Importo",
        "Divisa",
        "Causale",
        "Tipologia",
        "Cliente",
        "Nr fattura",
        "Note",
        "Descrizione",
        "Da verificare",
      ],
      filtrati.map((m) => [
        fmtData(m.dataContabile),
        fmtData(m.dataValuta),
        csvNum(m.importo),
        m.divisa,
        m.causale,
        m.tipologia,
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
      `overview-${ovMode}-${anno > 0 ? anno : "tutti"}`,
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
    <AppShell title={t("fin.title")} subtitle={t("fin.subtitle")}>
      <div className="mb-4 flex flex-wrap items-center gap-3">
        <div className="inline-flex flex-wrap rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
          {tabBtn("movimenti", <Table2 className="h-4 w-4" />, t("fin.tabMovimenti"))}
          {tabBtn("overview", <TrendingUp className="h-4 w-4" />, t("fin.tabOverview"))}
          {tabBtn(
            "anomalie",
            <AlertTriangle className="h-4 w-4" />,
            t("fin.tabAnomalie"),
            anomalie?.length ?? 0,
          )}
          {tabBtn("import", <Upload className="h-4 w-4" />, t("fin.tabImport"))}
          {tabBtn("storico", <History className="h-4 w-4" />, t("fin.tabStorico"))}
          {tabBtn("regole", <GraduationCap className="h-4 w-4" />, t("fin.tabRegole"))}
        </div>
        {(tab === "movimenti" || tab === "overview") && (
          <select
            value={anno}
            onChange={(e) => cambiaAnno(Number(e.target.value))}
            className={`${inputCls} w-auto`}
          >
            <option value={0}>{t("fin.allYears")}</option>
            {Array.from({ length: 6 }, (_, i) => new Date().getFullYear() - i).map((a) => (
              <option key={a} value={a}>
                {a}
              </option>
            ))}
          </select>
        )}
      </div>

      {/* ------------------------------- Movimenti ------------------------- */}
      {tab === "movimenti" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-end gap-3 mb-4">
            <div className="w-44">
              <label className="text-xs text-muted-foreground">{t("common.type")}</label>
              <select value={tipF} onChange={(e) => setTipF(e.target.value)} className={inputCls}>
                <option value="tutte">{t("common.allF")}</option>
                {tipologiePresenti.map((tp) => (
                  <option key={tp} value={tp}>
                    {tp}
                  </option>
                ))}
              </select>
            </div>
            <div className="w-36">
              <label className="text-xs text-muted-foreground">{t("fin.month")}</label>
              <select
                value={meseF}
                onChange={(e) => setMeseF(Number(e.target.value))}
                className={inputCls}
              >
                <option value={0}>{t("common.all")}</option>
                {mesi.map((m, i) => (
                  <option key={m} value={i + 1}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex-1 min-w-48">
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
              onClick={esportaMovimenti}
              disabled={filtrati.length === 0}
              className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
            >
              <Download className="h-4 w-4" /> {t("common.exportCsv")}
            </button>
          </div>
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
                    <th className="py-2 pr-3">{t("common.type")}</th>
                    <th className="py-2 pr-3">{t("fin.cliente")}</th>
                    <th className="py-2 pr-3">{t("fin.nrFattura")}</th>
                    <th className="py-2 pr-3">{t("fin.note")}</th>
                    <th className="py-2" />
                  </tr>
                </thead>
                <tbody>
                  {filtrati.slice(0, 500).map((m) => (
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
                      <td className="py-1.5 pr-3">
                        {m.tipologia}
                        {m.daVerificare && (
                          <AlertTriangle className="h-3.5 w-3.5 inline-block ml-1 text-status-absent" />
                        )}
                      </td>
                      <td className="py-1.5 pr-3">{m.cliente || "—"}</td>
                      <td className="py-1.5 pr-3 text-muted-foreground">{m.nrFattura || "—"}</td>
                      <td className="py-1.5 pr-3 text-muted-foreground">{m.note || "—"}</td>
                      <td className="py-1.5 text-right">
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
              <div className="mt-2 text-xs text-muted-foreground">
                {filtrati.length > 500
                  ? `${t("fin.first500")} ${filtrati.length}`
                  : `${filtrati.length} ${t("fin.rows")}`}
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
                        <select
                          value={editTip}
                          onChange={(e) => setEditTip(e.target.value)}
                          className={inputCls}
                        >
                          {TIPOLOGIE_MOVIMENTO.map((tp) => (
                            <option key={tp} value={tp}>
                              {tp}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">{t("fin.cliente")}</label>
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

      {/* ------------------------------- Import ---------------------------- */}
      {tab === "import" && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="text-sm font-semibold text-foreground mb-1">{t("fin.importTitle")}</div>
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
                    onChange={(e) => setSheetChoice({ ...sheetChoice, selected: e.target.value })}
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
                  <b className={preview.anomalie ? "text-status-absent" : ""}>{preview.anomalie}</b>
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
          <div className="mt-6 flex items-start gap-2 text-xs text-muted-foreground">
            <Landmark className="h-4 w-4 shrink-0 mt-0.5" />
            <p>{t("fin.apiNote")}</p>
          </div>
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
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaPattern")}</label>
                <input
                  value={rPattern}
                  onChange={(e) => setRPattern(e.target.value)}
                  placeholder={t("fin.regolaPatternPh")}
                  className={inputCls}
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.regolaCampo")}</label>
                <select
                  value={rCampo}
                  onChange={(e) => setRCampo(e.target.value as "cliente" | "descrizione")}
                  className={inputCls}
                >
                  <option value="cliente">{t("fin.campoCliente")}</option>
                  <option value="descrizione">{t("fin.campoDescrizione")}</option>
                </select>
              </div>
              {rCampo === "cliente" && (
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
                <select
                  value={rTipologia}
                  onChange={(e) => setRTipologia(e.target.value)}
                  className={inputCls}
                >
                  <option value="">{t("fin.regolaTipNoChange")}</option>
                  {TIPOLOGIE_MOVIMENTO.map((tp) => (
                    <option key={tp} value={tp}>
                      {tp}
                    </option>
                  ))}
                </select>
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
                  <GraduationCap className="h-4 w-4" /> {t("fin.regolaCrea")}
                </>
              )}
            </button>
          </div>

          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="text-sm font-semibold text-foreground mb-3">
              {t("fin.regoleElencoTitle")}
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
              <ul className="divide-y divide-border/60">
                {regole.map((r) => (
                  <li key={r.id} className="py-2.5 flex items-center gap-3 text-sm">
                    <span className="flex-1 min-w-0">
                      <span className="font-medium text-foreground">{r.pattern}</span>{" "}
                      <span className="text-xs text-muted-foreground">
                        (
                        {r.campo === "descrizione"
                          ? t("fin.campoDescrizione")
                          : `${t("fin.campoCliente")} · ${r.modo === "contiene" ? t("fin.modoContiene") : t("fin.modoEsatto")}`}
                        )
                      </span>
                      <span className="text-muted-foreground"> → </span>
                      {r.tipologia && (
                        <span className="text-xs rounded-full bg-primary/10 text-primary px-2 py-0.5 mr-1">
                          {r.tipologia}
                        </span>
                      )}
                      {r.cliente && (
                        <span className="text-xs rounded-full bg-muted px-2 py-0.5">
                          {r.cliente}
                        </span>
                      )}
                    </span>
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
        </div>
      )}
    </AppShell>
  );
}
