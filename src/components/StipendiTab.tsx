// DR Portal — Finanza → tab Stipendi (direzione).
// Costo del personale per mese dal file paghe "COSTI <MESE> <ANNO>.xlsx"
// (OneDrive Personale\MENSILITA'): import del foglio costi, dettaglio per
// dipendente, totale mese e aggancio della riga "Stipendi" ai Flussi di
// cassa (voce del mese di PAGAMENTO, di default il mese successivo alla
// competenza).

import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Banknote, Loader2, Trash2, Upload } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { esportaCsvFile } from "@/lib/csv";
import {
  meseSuccessivo,
  parseCostiMese,
  totaleMese,
  type ParseCostiResult,
  type StipendiDb,
} from "@/lib/stipendi-logic";
import {
  spStipendiEliminaMese,
  spStipendiGet,
  spStipendiSalvaMese,
} from "@/lib/stipendi.functions";
import { spUpsertFlussoCassa } from "@/lib/sharepoint.functions";

const inputCls =
  "rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function eur(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtMese(yyyymm: string): string {
  const [y, m] = yyyymm.split("-");
  const nomi = [
    "Gennaio",
    "Febbraio",
    "Marzo",
    "Aprile",
    "Maggio",
    "Giugno",
    "Luglio",
    "Agosto",
    "Settembre",
    "Ottobre",
    "Novembre",
    "Dicembre",
  ];
  return `${nomi[Number(m) - 1] ?? m} ${y}`;
}

interface PreviewStip {
  fileName: string;
  res: ParseCostiResult;
  sostituisce: boolean;
}

export function StipendiTab() {
  const { t } = useLang();
  const [db, setDb] = useState<StipendiDb | null>(null);
  const [errore, setErrore] = useState<string | null>(null);
  const [meseSel, setMeseSel] = useState("");
  const [q, setQ] = useState("");
  const [etichettaF, setEtichettaF] = useState("");
  const [showImport, setShowImport] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [preview, setPreview] = useState<PreviewStip | null>(null);
  const [aggiornaFlussi, setAggiornaFlussi] = useState(true);
  const [meseFlussi, setMeseFlussi] = useState("");
  const [saving, setSaving] = useState(false);
  const [confermaElimina, setConfermaElimina] = useState(false);

  const refresh = () =>
    spStipendiGet()
      .then((res) => {
        setDb(res);
        setMeseSel((prev) =>
          prev && res.mesi.some((m) => m.mese === prev)
            ? prev
            : (res.mesi[res.mesi.length - 1]?.mese ?? ""),
        );
      })
      .catch((err) => {
        const msg = err instanceof Error ? err.message : String(err);
        setErrore(msg);
        toast.error(msg);
      });
  useEffect(() => {
    void refresh();
  }, []);

  const mese = useMemo(() => db?.mesi.find((m) => m.mese === meseSel) ?? null, [db, meseSel]);
  const precedente = useMemo(() => {
    if (!db || !mese) return null;
    const i = db.mesi.findIndex((m) => m.mese === mese.mese);
    return i > 0 ? db.mesi[i - 1] : null;
  }, [db, mese]);

  const etichette = useMemo(
    () => [...new Set((mese?.dipendenti ?? []).map((d) => d.etichetta).filter(Boolean))].sort(),
    [mese],
  );
  const righe = useMemo(() => {
    const ql = q.trim().toLowerCase();
    return (mese?.dipendenti ?? [])
      .filter((d) => !etichettaF || d.etichetta === etichettaF)
      .filter((d) => !ql || `${d.cognome} ${d.nome} ${d.etichetta}`.toLowerCase().includes(ql));
  }, [mese, q, etichettaF]);
  const totRighe = useMemo(
    () => ({
      totale: righe.reduce((a, d) => a + d.totaleCosto, 0),
      ore: righe.reduce((a, d) => a + d.totaleOre, 0),
      ord: righe.reduce((a, d) => a + d.costoOrdinario, 0),
      str: righe.reduce((a, d) => a + d.costoStraordinario, 0),
      ferie: righe.reduce((a, d) => a + d.feriePermessi, 0),
      mens: righe.reduce((a, d) => a + d.mensilitaAggiuntive, 0),
      tfr: righe.reduce((a, d) => a + d.tfr, 0),
    }),
    [righe],
  );

  const onFile = async (f: File) => {
    setParsing(true);
    setPreview(null);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await f.arrayBuffer(), { cellDates: false });
      let res: ParseCostiResult | null = null;
      for (const nome of wb.SheetNames) {
        const matrix = XLSX.utils.sheet_to_json(wb.Sheets[nome], {
          header: 1,
          raw: true,
          defval: null,
        }) as unknown[][];
        res = parseCostiMese(matrix);
        if (res) break;
      }
      if (!res) {
        toast.error(t("stip.errTracciato"));
        return;
      }
      setPreview({
        fileName: f.name,
        res,
        sostituisce: !!db?.mesi.some((m) => m.mese === res!.mese),
      });
      setMeseFlussi(meseSuccessivo(res.mese));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsing(false);
    }
  };

  const conferma = async () => {
    if (!preview) return;
    setSaving(true);
    try {
      const res = await spStipendiSalvaMese({
        data: {
          mese: {
            mese: preview.res.mese,
            fonteFile: preview.fileName,
            dipendenti: preview.res.dipendenti,
          },
        },
      });
      if (aggiornaFlussi && /^\d{4}-\d{2}$/.test(meseFlussi)) {
        await spUpsertFlussoCassa({
          data: {
            nome: "Stipendi",
            genere: "voce",
            mese: meseFlussi,
            importo: -preview.res.totale,
            note: `${t("stip.flussiNota")} ${fmtMese(preview.res.mese)} (${preview.fileName})`,
          },
        });
      }
      setDb(res);
      setMeseSel(preview.res.mese);
      setPreview(null);
      setShowImport(false);
      toast.success(t("stip.importOk"), {
        description: `${fmtMese(preview.res.mese)} · ${preview.res.dipendenti.length} ${t("stip.dipendenti").toLowerCase()} · ${eur(preview.res.totale)} €${aggiornaFlussi ? ` · ${t("stip.flussiOk")} ${fmtMese(meseFlussi)}` : ""}`,
      });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const elimina = async () => {
    if (!mese) return;
    if (!confermaElimina) {
      setConfermaElimina(true);
      setTimeout(() => setConfermaElimina(false), 4000);
      return;
    }
    setConfermaElimina(false);
    try {
      const res = await spStipendiEliminaMese({ data: { mese: mese.mese } });
      setDb(res);
      setMeseSel(res.mesi[res.mesi.length - 1]?.mese ?? "");
      toast.success(t("stip.eliminato"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    }
  };

  const esporta = () => {
    if (!mese) return;
    esportaCsvFile(
      `stipendi-${mese.mese}`,
      [
        "Mese",
        "Etichetta",
        "Codice",
        "Cognome",
        "Nome",
        "Ore ordinarie",
        "Ore straordinarie",
        "Costo ordinario",
        "Straordinari",
        "Ferie/Permessi",
        "Mensilita aggiuntive",
        "TFR",
        "Totale costo",
        "Costo medio",
      ],
      mese.dipendenti.map((d) => [
        mese.mese,
        d.etichetta,
        d.codice,
        d.cognome,
        d.nome,
        d.oreOrdinarie,
        d.oreStraordinarie,
        d.costoOrdinario,
        d.costoStraordinario,
        d.feriePermessi,
        d.mensilitaAggiuntive,
        d.tfr,
        d.totaleCosto,
        d.costoMedio,
      ]),
    );
  };

  if (!db && !errore) {
    return (
      <div className="flex items-center gap-2 py-16 text-muted-foreground">
        <Loader2 className="h-5 w-5 animate-spin" /> {t("common.loading")}
      </div>
    );
  }
  if (errore && !db) return <p className="text-sm text-destructive">{errore}</p>;

  const totale = mese ? totaleMese(mese) : 0;
  const totalePrec = precedente ? totaleMese(precedente) : null;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end gap-3">
        <div>
          <label className="text-xs text-muted-foreground">{t("stip.mese")}</label>
          <select
            className={`${inputCls} mt-1 block`}
            value={meseSel}
            onChange={(e) => setMeseSel(e.target.value)}
          >
            {(db?.mesi ?? []).map((m) => (
              <option key={m.mese} value={m.mese}>
                {fmtMese(m.mese)}
              </option>
            ))}
            {(db?.mesi ?? []).length === 0 && <option value="">—</option>}
          </select>
        </div>
        <input
          className={`${inputCls} w-56`}
          placeholder={t("stip.cerca")}
          value={q}
          onChange={(e) => setQ(e.target.value)}
        />
        <select
          className={`${inputCls} max-w-56`}
          value={etichettaF}
          onChange={(e) => setEtichettaF(e.target.value)}
        >
          <option value="">
            {t("stip.etichetta")}: {t("common.all").toLowerCase()}
          </option>
          {etichette.map((a) => (
            <option key={a} value={a}>
              {a}
            </option>
          ))}
        </select>
        <div className="ml-auto flex gap-2">
          <button
            type="button"
            onClick={() => setShowImport((x) => !x)}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("stip.importBtn")}
          </button>
          <button
            type="button"
            onClick={esporta}
            disabled={!mese}
            className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
          >
            {t("common.exportCsv")}
          </button>
          {mese && (
            <button
              type="button"
              onClick={elimina}
              className={`inline-flex items-center gap-1 rounded-lg border px-3 py-2 text-sm ${confermaElimina ? "border-destructive bg-destructive text-destructive-foreground" : "border-border text-muted-foreground hover:bg-muted"}`}
            >
              <Trash2 className="h-4 w-4" />
              {confermaElimina ? t("stip.confermaElimina") : t("common.delete")}
            </button>
          )}
        </div>
      </div>

      {showImport && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-1 text-sm font-semibold text-foreground">{t("stip.importTitle")}</div>
          <p className="mb-4 text-xs text-muted-foreground">{t("stip.importDesc")}</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            disabled={parsing || saving}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void onFile(f);
              e.target.value = "";
            }}
            className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
          />
          {parsing && (
            <p className="mt-3 inline-flex items-center gap-2 text-sm text-muted-foreground">
              <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
            </p>
          )}
          {preview && (
            <div className="mt-4 rounded-xl border border-border p-4">
              <div className="text-sm font-medium text-foreground">{preview.fileName}</div>
              <ul className="mt-2 space-y-1 text-[13px] text-muted-foreground">
                <li>
                  {t("stip.mese")}: <b>{fmtMese(preview.res.mese)}</b>
                  {preview.sostituisce && (
                    <span className="ml-2 rounded bg-status-break/20 px-1.5 py-0.5 text-xs text-status-break">
                      {t("stip.sovrascrive")}
                    </span>
                  )}
                </li>
                <li>
                  {t("stip.dipendenti")}: <b>{preview.res.dipendenti.length}</b>
                  {preview.res.scartate > 0 &&
                    ` (${preview.res.scartate} ${t("fin.previewSkipped")})`}
                </li>
                <li>
                  {t("stip.totaleMese")}: <b>{eur(preview.res.totale)} €</b>
                </li>
                {preview.res.squadrature > 0 && (
                  <li className="text-status-break">
                    {preview.res.squadrature} {t("stip.squadrature")}
                  </li>
                )}
              </ul>
              <div className="mt-3 flex flex-wrap items-center gap-3 text-sm">
                <label className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={aggiornaFlussi}
                    onChange={(e) => setAggiornaFlussi(e.target.checked)}
                  />
                  {t("stip.flussiCheck")}
                </label>
                <input
                  type="month"
                  className={inputCls}
                  value={meseFlussi}
                  onChange={(e) => setMeseFlussi(e.target.value)}
                  disabled={!aggiornaFlussi}
                />
              </div>
              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  onClick={() => void conferma()}
                  disabled={saving}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {saving && <Loader2 className="h-4 w-4 animate-spin" />}
                  {t("stip.conferma")}
                </button>
                <button
                  type="button"
                  onClick={() => setPreview(null)}
                  disabled={saving}
                  className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                >
                  {t("common.cancel")}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {mese ? (
        <>
          <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.totaleMese")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {eur(totale)} €
              </p>
              {totalePrec != null && totalePrec > 0 && (
                <p
                  className={`text-xs ${totale > totalePrec ? "text-status-out" : "text-status-present"}`}
                >
                  {totale >= totalePrec ? "+" : ""}
                  {eur(totale - totalePrec)} € {t("stip.vsPrec")} {fmtMese(precedente!.mese)}
                </p>
              )}
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.dipendenti")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {mese.dipendenti.length}
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.oreTot")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {mese.dipendenti.reduce((a, d) => a + d.totaleOre, 0).toLocaleString("it-IT")}
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.mediaDip")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {eur(mese.dipendenti.length ? totale / mese.dipendenti.length : 0)} €
              </p>
            </div>
          </div>

          <div className="overflow-x-auto rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-border text-left text-xs uppercase tracking-wider text-muted-foreground">
                  <th className="px-3 py-2">{t("stip.etichetta")}</th>
                  <th className="px-3 py-2">{t("common.employee")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.oreOrd")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.oreStr")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.ordinario")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.straord")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.ferie")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.mensAgg")}</th>
                  <th className="px-3 py-2 text-right">TFR</th>
                  <th className="px-3 py-2 text-right">{t("stip.totale")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.medio")}</th>
                </tr>
              </thead>
              <tbody>
                {righe.map((d) => (
                  <tr key={d.codice} className="border-b border-border/60 hover:bg-muted/40">
                    <td className="max-w-44 truncate px-3 py-1.5 text-xs text-muted-foreground">
                      {d.etichetta || "—"}
                    </td>
                    <td className="px-3 py-1.5 font-medium">
                      {d.cognome} {d.nome}
                      <span className="ml-1 font-mono text-[10px] text-muted-foreground">
                        {d.codice}
                      </span>
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{d.oreOrdinarie || "—"}</td>
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {d.oreStraordinarie || "—"}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{eur(d.costoOrdinario)}</td>
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {d.costoStraordinario ? eur(d.costoStraordinario) : "—"}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{eur(d.feriePermessi)}</td>
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {eur(d.mensilitaAggiuntive)}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{eur(d.tfr)}</td>
                    <td className="px-3 py-1.5 text-right font-semibold tabular-nums">
                      {eur(d.totaleCosto)}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums text-muted-foreground">
                      {d.costoMedio ? eur(d.costoMedio) : "—"}
                    </td>
                  </tr>
                ))}
                <tr className="bg-muted/40 font-semibold">
                  <td className="px-3 py-2" colSpan={2}>
                    {t("common.total")} ({righe.length})
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {totRighe.ore ? totRighe.ore.toLocaleString("it-IT") : ""}
                  </td>
                  <td className="px-3 py-2" />
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.ord)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.str)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.ferie)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.mens)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.tfr)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.totale)}</td>
                  <td className="px-3 py-2" />
                </tr>
              </tbody>
            </table>
          </div>

          {(db?.mesi.length ?? 0) > 1 && (
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="mb-2 flex items-center gap-2 text-sm font-semibold text-foreground">
                <Banknote className="h-4 w-4 text-primary" /> {t("stip.trend")}
              </p>
              <div className="flex flex-wrap gap-2">
                {db!.mesi.map((m) => (
                  <button
                    key={m.mese}
                    type="button"
                    onClick={() => setMeseSel(m.mese)}
                    className={`rounded-lg border px-3 py-1.5 text-xs ${m.mese === meseSel ? "border-primary bg-primary/10 text-primary" : "border-border text-foreground hover:bg-muted"}`}
                  >
                    {fmtMese(m.mese)}
                    <span className="ml-1.5 font-semibold tabular-nums">
                      {eur(totaleMese(m))} €
                    </span>
                  </button>
                ))}
              </div>
            </div>
          )}
        </>
      ) : (
        <p className="py-8 text-center text-sm text-muted-foreground">{t("stip.vuoto")}</p>
      )}
    </div>
  );
}
