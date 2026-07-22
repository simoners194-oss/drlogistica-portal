// DR Portal — Finanza (sezione riservata al direttore DR005 + admin).
// Estratto conto bancario: import da xlsx, archivio movimenti classificati,
// overview incassi per cliente/mese, anomalie da sanare a mano.
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
} from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import { isSupervisoreGlobale } from "@/lib/richieste-logic";
import {
  parseEstratto,
  parseMatrice,
  TIPOLOGIE_MOVIMENTO,
  type MovimentoParsed,
} from "@/lib/finanza-logic";
import {
  spGetMovimenti,
  spGetMovimentiChiavi,
  spImportMovimenti,
  spUpdateMovimento,
} from "@/lib/sharepoint.functions";
import type { SpMovimento } from "@/lib/sharepoint.server";

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

// Blocchi di upload verso il server (sotto il limite server di 150).
const CHUNK = 100;

type Tab = "movimenti" | "overview" | "anomalie" | "import";

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
  const [anno, setAnno] = useState(new Date().getFullYear());

  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [anomalie, setAnomalie] = useState<SpMovimento[] | null>(null);

  // Filtri archivio movimenti
  const [tipF, setTipF] = useState("tutte");
  const [cercaF, setCercaF] = useState("");
  const [meseF, setMeseF] = useState(0); // 0 = tutti

  // Import
  const [preview, setPreview] = useState<PreviewImport | null>(null);
  const [parsing, setParsing] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importProgress, setImportProgress] = useState("");

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
    spGetMovimenti({ data: { from: `${a}-01-01`, to: `${a}-12-31` } })
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
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    const dir = s.ruolo === "amministratore_sistema" || isSupervisoreGlobale(s.codice ?? "");
    if (!dir) return;
    loadMovimenti(anno);
    loadAnomalie();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const cambiaAnno = (a: number) => {
    setAnno(a);
    loadMovimenti(a);
  };

  // --- Import xlsx ----------------------------------------------------------
  const onFile = async (file: File) => {
    setPreview(null);
    setParsing(true);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
      let parsedRows: ReturnType<typeof parseMatrice> = null;
      for (const name of wb.SheetNames) {
        const matrix = XLSX.utils.sheet_to_json(wb.Sheets[name], {
          header: 1,
          raw: true,
        }) as unknown[][];
        const res = parseMatrice(matrix);
        if (res && res.rows.length) {
          parsedRows = res;
          break;
        }
      }
      if (!parsedRows) {
        toast.error(t("fin.errFile"), { description: t("fin.errFileDesc") });
        return;
      }
      const righe = parseEstratto(parsedRows.rows);
      const chiavi = new Set((await spGetMovimentiChiavi()) as string[]);
      const nuove = righe.filter((r) => !chiavi.has(r.chiave));
      const date = righe.map((r) => r.dataContabile).sort();
      setPreview({
        fileName: file.name,
        righe,
        nuove,
        doppioni: righe.length - nuove.length,
        scartate: parsedRows.scartate,
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

  const eseguiImport = async () => {
    if (!preview || preview.nuove.length === 0) return;
    setImporting(true);
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
        const res = await spImportMovimenti({ data: { rows } });
        importati += res.importati;
        doppioni += res.doppioni;
        errori.push(...res.errori);
      }
      toast.success(t("fin.importDone"), {
        description: `${importati} ${t("fin.importedRows")}${doppioni ? ` · ${doppioni} ${t("fin.skippedDup")}` : ""}${errori.length ? ` · ${errori.length} ${t("common.error").toLowerCase()}` : ""}`,
      });
      if (errori.length) console.warn("Import movimenti — errori:", errori);
      setPreview(null);
      loadMovimenti(anno);
      loadAnomalie();
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

  // Overview incassi: pivot cliente × mese sui soli importi positivi di tipo
  // Incasso (come il foglio "Overview incassi", ma raggruppato per cliente).
  const overview = useMemo(() => {
    const incassi = (movimenti ?? []).filter((m) => m.tipologia === "Incasso" && m.importo > 0);
    const byCliente = new Map<string, { mesi: number[]; tot: number }>();
    const totMesi = Array.from({ length: 12 }, () => 0);
    let tot = 0;
    for (const m of incassi) {
      const mese = Number(m.dataContabile.slice(5, 7)) - 1;
      if (mese < 0 || mese > 11) continue;
      const key = m.cliente || `(${t("fin.unknownClient")})`;
      const row = byCliente.get(key) ?? { mesi: Array.from({ length: 12 }, () => 0), tot: 0 };
      row.mesi[mese] += m.importo;
      row.tot += m.importo;
      byCliente.set(key, row);
      totMesi[mese] += m.importo;
      tot += m.importo;
    }
    const righe = [...byCliente.entries()].sort((a, b) => b[1].tot - a[1].tot);
    return { righe, totMesi, tot, count: incassi.length };
  }, [movimenti, t]);

  const mesi = lang === "it" ? MESI_IT : MESI_EN;

  const esportaMovimenti = () => {
    esportaCsvFile(
      `movimenti-${anno}`,
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
        m.importo,
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
      `overview-incassi-${anno}`,
      ["Cliente", ...mesi, "Totale"],
      [
        ...overview.righe.map(([cliente, r]) => [
          cliente,
          ...r.mesi.map((v) => (v ? v : "")),
          r.tot,
        ]),
        [t("common.total"), ...overview.totMesi.map((v) => (v ? v : "")), overview.tot],
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
        <div className="inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
          {tabBtn("movimenti", <Table2 className="h-4 w-4" />, t("fin.tabMovimenti"))}
          {tabBtn("overview", <TrendingUp className="h-4 w-4" />, t("fin.tabOverview"))}
          {tabBtn(
            "anomalie",
            <AlertTriangle className="h-4 w-4" />,
            t("fin.tabAnomalie"),
            anomalie?.length ?? 0,
          )}
          {tabBtn("import", <Upload className="h-4 w-4" />, t("fin.tabImport"))}
        </div>
        {tab !== "anomalie" && tab !== "import" && (
          <select
            value={anno}
            onChange={(e) => cambiaAnno(Number(e.target.value))}
            className={`${inputCls} w-auto`}
          >
            {Array.from({ length: 4 }, (_, i) => new Date().getFullYear() - i).map((a) => (
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
          <div className="flex items-center justify-between mb-4">
            <div>
              <div className="text-sm font-semibold text-foreground">{t("fin.overviewTitle")}</div>
              <p className="text-xs text-muted-foreground mt-0.5">
                {overview.count} {t("fin.overviewCount")} · {t("common.total")}{" "}
                <span className="font-semibold text-status-present">
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
                    <th className="py-2 pr-3">{t("fin.cliente")}</th>
                    {mesi.map((m) => (
                      <th key={m} className="py-2 px-2 text-right">
                        {m}
                      </th>
                    ))}
                    <th className="py-2 pl-2 text-right">{t("common.total")}</th>
                  </tr>
                </thead>
                <tbody>
                  {overview.righe.map(([cliente, r]) => (
                    <tr key={cliente} className="border-b border-border/50 hover:bg-muted/40">
                      <td className="py-1.5 pr-3 max-w-56 truncate" title={cliente}>
                        {cliente}
                      </td>
                      {r.mesi.map((v, i) => (
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
                    {overview.totMesi.map((v, i) => (
                      <td key={i} className="py-2 px-2 text-right whitespace-nowrap">
                        {v ? fmtImporto(v) : ""}
                      </td>
                    ))}
                    <td className="py-2 pl-2 text-right text-status-present whitespace-nowrap">
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
                    <button
                      type="button"
                      onClick={() => apriEdit(m)}
                      className="mt-3 rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                    >
                      {t("fin.fix")}
                    </button>
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
    </AppShell>
  );
}
