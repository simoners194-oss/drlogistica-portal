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
  parseCostiFile,
  totaleMese,
  type ParseCostiFileResult,
  type StipendiDb,
  type StipendioDipendente,
} from "@/lib/stipendi-logic";
import {
  spStipendiEliminaMese,
  spStipendiGet,
  spStipendiSalvaMese,
} from "@/lib/stipendi.functions";
import { spGetDettagliDistinte, spUpsertFlussoCassa } from "@/lib/sharepoint.functions";
import type { DettaglioDistinta } from "@/lib/sharepoint.server";

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
  res: ParseCostiFileResult;
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
  // Mese di competenza del file in anteprima: proposto dal contenuto (o dal
  // nome file per il tracciato per-appalto) e sempre correggibile a mano.
  const [meseComp, setMeseComp] = useState("");
  const [saving, setSaving] = useState(false);
  const [confermaElimina, setConfermaElimina] = useState(false);
  // Distinte stipendi (report BPM "Esiti pagamenti"): per l'EFFETTIVO versato.
  const [distinte, setDistinte] = useState<DettaglioDistinta[] | null>(null);

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
    spGetDettagliDistinte()
      .then((l) => setDistinte(l as DettaglioDistinta[]))
      .catch(() => setDistinte([]));
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

  // EFFETTIVO VERSATO (richiesta direzione 14/09): il netto uscito dal conto
  // verso i dipendenti, dalle distinte stipendi BPM (report "Esiti pagamenti"
  // importato in Finanze → Storico estratti). Gli stipendi del mese M si
  // pagano nel mese M+1: per il mese selezionato si guardano le disposizioni
  // salari ESEGUITE nel mese successivo, abbinate al dipendente per nome.
  const nameKey = (s: string) =>
    s
      .toLowerCase()
      .replace(/[^a-zà-ù ]/g, " ")
      .split(/\s+/)
      .filter(Boolean)
      .sort()
      .join(" ");
  const salariPag = useMemo(() => {
    const mesePag = /^\d{4}-\d{2}$/.test(meseSel) ? meseSuccessivo(meseSel) : "";
    // I valori reali del campo sono "Stipendi SEPA" (e "Pagamento Riba" da
    // escludere): si accettano entrambe le diciture stipendi/salari.
    const righeS = (distinte ?? []).filter(
      (d) =>
        /salar|stipend/i.test(d.tipoPagamento) && (d.dataEsecuzione || "").slice(0, 7) === mesePag,
    );
    const perNome = new Map<string, number>();
    for (const d of righeS) {
      const k = nameKey(d.beneficiario);
      perNome.set(k, (perNome.get(k) ?? 0) + d.importo);
    }
    return {
      mesePag,
      righe: righeS,
      perNome,
      totale: righeS.reduce((a, d) => a + d.importo, 0),
    };
  }, [distinte, meseSel]);
  const versatoDi = (d: StipendioDipendente): number | null =>
    salariPag.perNome.get(nameKey(`${d.cognome} ${d.nome}`)) ?? null;
  const nonAbbinate = useMemo(() => {
    const chiavi = new Set((mese?.dipendenti ?? []).map((x) => nameKey(`${x.cognome} ${x.nome}`)));
    return salariPag.righe.filter((r) => !chiavi.has(nameKey(r.beneficiario)));
  }, [salariPag, mese]);

  const onFile = async (f: File) => {
    setParsing(true);
    setPreview(null);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await f.arrayBuffer(), { cellDates: false });
      const fogli = wb.SheetNames.map((nome) => ({
        nome,
        matrix: XLSX.utils.sheet_to_json(wb.Sheets[nome], {
          header: 1,
          raw: true,
          defval: null,
        }) as unknown[][],
      }));
      const res = parseCostiFile(fogli, f.name);
      if (!res) {
        toast.error(t("stip.errTracciato"));
        return;
      }
      setPreview({ fileName: f.name, res });
      setMeseComp(res.mese);
      setMeseFlussi(res.mese ? meseSuccessivo(res.mese) : "");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsing(false);
    }
  };

  const conferma = async () => {
    if (!preview || !/^\d{4}-\d{2}$/.test(meseComp)) return;
    setSaving(true);
    try {
      const res = await spStipendiSalvaMese({
        data: {
          mese: {
            mese: meseComp,
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
            note: `${t("stip.flussiNota")} ${fmtMese(meseComp)} (${preview.fileName})`,
          },
        });
      }
      setDb(res);
      setMeseSel(meseComp);
      setPreview(null);
      setShowImport(false);
      toast.success(t("stip.importOk"), {
        description: `${fmtMese(meseComp)} · ${preview.res.dipendenti.length} ${t("stip.dipendenti").toLowerCase()} · ${eur(preview.res.totale)} €${aggiornaFlussi && meseFlussi ? ` · ${t("stip.flussiOk")} ${fmtMese(meseFlussi)}` : ""}`,
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
        "Versato in banca",
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
        versatoDi(d) ?? "",
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
                <li className="flex items-center gap-2">
                  {t("stip.mese")}:
                  <input
                    type="month"
                    className={inputCls}
                    value={meseComp}
                    onChange={(e) => {
                      setMeseComp(e.target.value);
                      if (/^\d{4}-\d{2}$/.test(e.target.value))
                        setMeseFlussi(meseSuccessivo(e.target.value));
                    }}
                  />
                  {!preview.res.mese && (
                    <span className="text-xs text-status-break">{t("stip.meseNonNelFile")}</span>
                  )}
                  {/^\d{4}-\d{2}$/.test(meseComp) && db?.mesi.some((m) => m.mese === meseComp) && (
                    <span className="rounded bg-status-break/20 px-1.5 py-0.5 text-xs text-status-break">
                      {t("stip.sovrascrive")}
                    </span>
                  )}
                </li>
                <li>
                  {t("stip.fogli")}: {preview.res.fogli.join(", ")}
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
                  disabled={saving || !/^\d{4}-\d{2}$/.test(meseComp)}
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
          <div className="grid grid-cols-2 gap-3 lg:grid-cols-5">
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
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.versato")} · {salariPag.mesePag ? fmtMese(salariPag.mesePag) : "—"}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {salariPag.righe.length ? `${eur(salariPag.totale)} €` : "—"}
              </p>
              <p className="text-xs text-muted-foreground">
                {salariPag.righe.length
                  ? `${salariPag.righe.length} ${t("stip.disposizioni")}${nonAbbinate.length ? ` · ${nonAbbinate.length} ${t("stip.nonAbbinate")}` : ""}`
                  : t("stip.versatoVuoto")}
              </p>
            </div>
          </div>
          <p className="text-[11px] text-muted-foreground">{t("stip.versatoNota")}</p>

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
                  <th className="px-3 py-2 text-right">{t("stip.versatoCol")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.medio")}</th>
                </tr>
              </thead>
              <tbody>
                {righe.map((d) => (
                  <tr
                    key={`${d.codice}|${d.cognome} ${d.nome}|${d.etichetta}`}
                    className="border-b border-border/60 hover:bg-muted/40"
                  >
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
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {versatoDi(d) != null ? eur(versatoDi(d)!) : "—"}
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
                  <td className="px-3 py-2 text-right tabular-nums">
                    {eur(righe.reduce((a, d) => a + (versatoDi(d) ?? 0), 0))}
                  </td>
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
