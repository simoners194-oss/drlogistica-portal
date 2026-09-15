// DR Portal — Finanza → tab Fiscale (direzione).
// Sostituisce l'Excel "SCADENZARIO FISCALE" di Sabrina: le sue viste
// (elenco scadenze, pivot Da pagare, pivot Pagato) diventano compilabili
// direttamente sul portale, con export CSV di ogni tabella. Le voci
// "Costo fiscale rate/corrente" dei Flussi leggono questi stessi dati.
// L'import del file (pannello nei Flussi) resta per la transizione: un
// re-import sostituisce le righe DA FILE e conserva quelle del portale.

import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Loader2, Pencil, Plus, Trash2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { esportaCsvFile } from "@/lib/csv";
import {
  spFiscaleDaRateizzare,
  spFiscaleEliminaScadenza,
  spFiscaleGet,
  spFiscaleUpsertScadenza,
} from "@/lib/fiscale.functions";
import type { DaRateizzareFiscale, FiscaleDb, ScadenzaFiscale } from "@/lib/fiscale-logic";

const inputCls =
  "rounded-lg border border-border bg-background px-2 py-1 text-[13px] text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function eur(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function dataIt(iso: string): string {
  const [y, m, g] = (iso ?? "").slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : (iso ?? "");
}
function numDaTesto(s: string): number {
  const t = s.trim();
  if (!t) return NaN;
  if (t.includes(",")) return Number(t.replace(/\./g, "").replace(",", "."));
  // "66.285" senza virgola: punto = separatore di MIGLIAIA (formato
  // italiano), non un decimale — altrimenti 66.285,00 € diventavano 66,29.
  if (/^-?\d{1,3}(\.\d{3})+$/.test(t)) return Number(t.replace(/\./g, ""));
  return Number(t);
}

/** Numero per il CSV in convenzione Excel italiano (virgola decimale). */
function csvNum(n: number): string {
  return n.toFixed(2).replace(".", ",");
}

const MESI_LBL = [
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

/** Colonne pivot: i mesi dell'anno di riferimento uno a uno, gli altri anni
 *  aggregati (stessa forma delle pivot di Sabrina). */
function colonnePivot(mesi: string[], annoRif: number): { chiave: string; label: string }[] {
  const set = new Set<string>();
  for (const m of mesi) set.add(m.slice(0, 4) === String(annoRif) ? m : m.slice(0, 4));
  return [...set].sort().map((c) => ({
    chiave: c,
    label: c.length === 7 ? `${MESI_LBL[Number(c.slice(5)) - 1]} ${c.slice(0, 4)}` : c,
  }));
}

interface FormScadenza {
  id?: string;
  voce: string;
  anno: string;
  periodo: string;
  importo: string;
  dataPagamento: string;
  modalita: string;
  note: string;
  pagato: boolean;
}

const FORM_VUOTO: FormScadenza = {
  voce: "",
  anno: "",
  periodo: "",
  importo: "",
  dataPagamento: "",
  modalita: "F24",
  note: "",
  pagato: false,
};

export function FiscaleTab() {
  const { t } = useLang();
  const [db, setDb] = useState<FiscaleDb | null>(null);
  const [errore, setErrore] = useState<string | null>(null);
  const [vista, setVista] = useState<"scadenze" | "dapagare" | "pagato">("scadenze");
  const [q, setQ] = useState("");
  const [statoF, setStatoF] = useState<"aperte" | "pagate" | "tutte">("aperte");
  const [voceF, setVoceF] = useState("");
  // Filtro periodo dal drill-down delle pivot: "2026-09" (mese) o "2027".
  const [meseF, setMeseF] = useState("");
  // Form "da registrare in futuro" (voce/anno/periodo/importo).
  const [drVoce, setDrVoce] = useState("");
  const [drAnno, setDrAnno] = useState("");
  const [drPeriodo, setDrPeriodo] = useState("");
  const [drImporto, setDrImporto] = useState("");
  const [busy, setBusy] = useState(false);
  // Form nuova scadenza / modifica (stesso form, id pieno = modifica).
  const [form, setForm] = useState<FormScadenza | null>(null);
  const [confermaElimina, setConfermaElimina] = useState<string | null>(null);

  useEffect(() => {
    spFiscaleGet()
      .then((res) => setDb(res))
      .catch((err) => setErrore(err instanceof Error ? err.message : String(err)));
  }, []);

  const oggiISO = new Date().toISOString().slice(0, 10);
  const annoRif = Number(oggiISO.slice(0, 4));

  const scadenze = useMemo(() => db?.scadenze ?? [], [db]);
  const voci = useMemo(() => [...new Set(scadenze.map((s) => s.voce))].sort(), [scadenze]);
  const modalitaNote = useMemo(
    () => [...new Set(scadenze.map((s) => s.modalita).filter(Boolean))] as string[],
    [scadenze],
  );

  const filtrate = useMemo(() => {
    const ql = q.trim().toLowerCase();
    return scadenze
      .filter((s) => (statoF === "tutte" ? true : statoF === "pagate" ? s.pagato : !s.pagato))
      .filter((s) => !voceF || s.voce === voceF)
      .filter((s) =>
        !meseF ? true : s.dataPagamento.slice(0, meseF.length === 7 ? 7 : 4) === meseF,
      )
      .filter(
        (s) =>
          !ql ||
          `${s.voce} ${s.voceOld ?? ""} ${s.periodo ?? ""} ${s.note ?? ""} ${s.modalita ?? ""}`
            .toLowerCase()
            .includes(ql),
      )
      .sort((a, b) => (a.dataPagamento < b.dataPagamento ? -1 : 1));
  }, [scadenze, q, statoF, voceF, meseF]);

  const totFiltrate = useMemo(() => filtrate.reduce((a, s) => a + s.importo, 0), [filtrate]);

  // Riepilogo (solo NON pagate).
  const aperte = useMemo(() => scadenze.filter((s) => !s.pagato), [scadenze]);
  const riepilogo = useMemo(() => {
    const somma = (c: ScadenzaFiscale["categoria"]) =>
      aperte.filter((s) => s.categoria === c).reduce((a, s) => a + s.importo, 0);
    return {
      totale: aperte.reduce((a, s) => a + s.importo, 0),
      rate: somma("rate"),
      correnti: somma("corrente"),
      finanziamenti: somma("finanziamento"),
      ultima: aperte.reduce((m, s) => (s.dataPagamento > m ? s.dataPagamento : m), ""),
    };
  }, [aperte]);

  /** Pivot voce × periodo su un sottoinsieme di scadenze. */
  const pivot = (righe: ScadenzaFiscale[]) => {
    const mesi = righe.map((s) => s.dataPagamento.slice(0, 7));
    const cols = colonnePivot(mesi, annoRif);
    const chiaveCol = (s: ScadenzaFiscale) => {
      const m = s.dataPagamento.slice(0, 7);
      return m.slice(0, 4) === String(annoRif) ? m : m.slice(0, 4);
    };
    const perVoce = new Map<string, Map<string, number>>();
    for (const s of righe) {
      if (!perVoce.has(s.voce)) perVoce.set(s.voce, new Map());
      const r = perVoce.get(s.voce)!;
      const k = chiaveCol(s);
      r.set(k, (r.get(k) ?? 0) + s.importo);
    }
    const vociOrd = [...perVoce.keys()].sort();
    return { cols, perVoce, vociOrd };
  };
  const pivotDaPagare = useMemo(() => pivot(aperte), [aperte, annoRif]); // eslint-disable-line react-hooks/exhaustive-deps
  const pivotPagato = useMemo(
    () => pivot(scadenze.filter((s) => s.pagato)),
    [scadenze, annoRif], // eslint-disable-line react-hooks/exhaustive-deps
  );

  const salva = async (f: FormScadenza) => {
    const importo = numDaTesto(f.importo);
    if (!Number.isFinite(importo) || importo <= 0) {
      toast.error(t("fis.errImporto"));
      return;
    }
    if (!f.voce.trim() || !/^\d{4}-\d{2}-\d{2}$/.test(f.dataPagamento)) {
      toast.error(t("fis.errCampi"));
      return;
    }
    setBusy(true);
    try {
      const res = await spFiscaleUpsertScadenza({
        data: {
          scadenza: {
            id: f.id,
            voce: f.voce.trim(),
            anno: f.anno.trim() ? Number(f.anno.trim()) : undefined,
            periodo: f.periodo.trim() || undefined,
            importo: Math.round(importo * 100) / 100,
            dataPagamento: f.dataPagamento,
            pagato: f.pagato,
            modalita: f.modalita.trim() || undefined,
            note: f.note.trim() || undefined,
            categoria: "corrente", // ricalcolata dal server
          },
        },
      });
      setDb(res);
      setForm(null);
      toast.success(t(f.id ? "fis.okModifica" : "fis.okNuova"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const togglePagato = async (s: ScadenzaFiscale) => {
    setBusy(true);
    try {
      const res = await spFiscaleUpsertScadenza({
        data: { scadenza: { ...s, pagato: !s.pagato } },
      });
      setDb(res);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const salvaDaRateizzare = async (lista: DaRateizzareFiscale[]): Promise<boolean> => {
    setBusy(true);
    try {
      const res = await spFiscaleDaRateizzare({ data: { lista } });
      setDb(res);
      return true;
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
      return false;
    } finally {
      setBusy(false);
    }
  };

  const elimina = async (id: string) => {
    setBusy(true);
    try {
      const res = await spFiscaleEliminaScadenza({ data: { id } });
      setDb(res);
      setConfermaElimina(null);
      toast.success(t("fis.okElimina"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  };

  const esportaScadenze = () =>
    esportaCsvFile(
      "scadenzario-fiscale",
      [
        t("fis.colData"),
        t("fis.colVoce"),
        t("fis.colDettaglio"),
        t("fis.colAnno"),
        t("fis.colPeriodo"),
        t("fis.colImporto"),
        t("fis.colModalita"),
        t("fis.colCategoria"),
        t("fis.colPagato"),
        t("fis.colNote"),
      ],
      filtrate.map((s) => [
        s.dataPagamento,
        s.voce,
        s.voceOld ?? "",
        s.anno ?? "",
        s.periodo ?? "",
        csvNum(s.importo),
        s.modalita ?? "",
        s.categoria,
        s.pagato ? "SI" : "NO",
        s.note ?? "",
      ]),
    );

  const esportaPivot = (p: ReturnType<typeof pivot>, nome: string) =>
    esportaCsvFile(
      nome,
      [t("fis.colVoce"), ...p.cols.map((c) => c.label), t("fis.colTotale")],
      [
        ...p.vociOrd.map((v) => {
          const r = p.perVoce.get(v)!;
          const tot = [...r.values()].reduce((a, x) => a + x, 0);
          return [
            v,
            ...p.cols.map((c) => (r.get(c.chiave) != null ? csvNum(r.get(c.chiave)!) : "")),
            csvNum(tot),
          ];
        }),
        [
          t("fis.colTotale"),
          ...p.cols.map((c) =>
            csvNum(p.vociOrd.reduce((a, v) => a + (p.perVoce.get(v)!.get(c.chiave) ?? 0), 0)),
          ),
          csvNum(
            p.vociOrd.reduce(
              (a, v) => a + [...p.perVoce.get(v)!.values()].reduce((x, y) => x + y, 0),
              0,
            ),
          ),
        ],
      ],
    );

  const badgeCat = (c: ScadenzaFiscale["categoria"]) => (
    <span
      className={`rounded-full px-2 py-0.5 text-[10px] font-medium ${
        c === "rate"
          ? "bg-status-absent/15 text-status-absent"
          : c === "corrente"
            ? "bg-primary/10 text-primary"
            : "bg-muted text-muted-foreground"
      }`}
    >
      {t(c === "rate" ? "fis.catRate" : c === "corrente" ? "fis.catCorrente" : "fis.catFin")}
    </span>
  );

  const formUI = form && (
    <div className="mb-4 grid grid-cols-2 gap-2 rounded-xl border border-primary/40 p-3 text-[13px] md:grid-cols-4 lg:grid-cols-8">
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colVoce")} *</span>
        <input
          list="fis-voci"
          value={form.voce}
          onChange={(e) => setForm({ ...form, voce: e.target.value })}
          className={inputCls}
        />
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colAnno")}</span>
        <input
          value={form.anno}
          onChange={(e) => setForm({ ...form, anno: e.target.value })}
          placeholder={String(annoRif)}
          className={inputCls}
        />
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colPeriodo")}</span>
        <input
          value={form.periodo}
          onChange={(e) => setForm({ ...form, periodo: e.target.value })}
          placeholder={t("fis.periodoPh")}
          className={inputCls}
        />
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colImporto")} *</span>
        <input
          value={form.importo}
          onChange={(e) => setForm({ ...form, importo: e.target.value })}
          placeholder="0,00"
          className={`${inputCls} text-right`}
        />
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colData")} *</span>
        <input
          type="date"
          value={form.dataPagamento}
          onChange={(e) => setForm({ ...form, dataPagamento: e.target.value })}
          className={inputCls}
        />
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colModalita")}</span>
        <input
          list="fis-modalita"
          value={form.modalita}
          onChange={(e) => setForm({ ...form, modalita: e.target.value })}
          className={inputCls}
        />
        <datalist id="fis-modalita">
          {[...new Set(["F24", "QR CODE", "RID", ...modalitaNote])].map((m) => (
            <option key={m} value={m} />
          ))}
        </datalist>
      </label>
      <label className="flex flex-col gap-1">
        <span className="text-xs text-muted-foreground">{t("fis.colNote")}</span>
        <input
          value={form.note}
          onChange={(e) => setForm({ ...form, note: e.target.value })}
          className={inputCls}
        />
      </label>
      <div className="flex items-end gap-2">
        <label className="flex items-center gap-1.5 pb-1 text-xs">
          <input
            type="checkbox"
            checked={form.pagato}
            onChange={(e) => setForm({ ...form, pagato: e.target.checked })}
            className="accent-primary"
          />
          {t("fis.colPagato")}
        </label>
        <button
          type="button"
          disabled={busy}
          onClick={() => void salva(form)}
          className="rounded-lg bg-primary px-3 py-1 font-medium text-primary-foreground disabled:opacity-40"
        >
          {form.id ? t("common.save") : t("fis.aggiungi")}
        </button>
        <button
          type="button"
          onClick={() => setForm(null)}
          className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
        >
          {t("common.cancel")}
        </button>
      </div>
    </div>
  );

  const apriDettaglio = (voce: string, colonna: string, pagate: boolean) => {
    setVista("scadenze");
    setStatoF(pagate ? "pagate" : "aperte");
    setVoceF(voce);
    setMeseF(colonna);
    setQ("");
  };

  const pivotUI = (p: ReturnType<typeof pivot>, nomeCsv: string, pagate: boolean) => (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[700px] text-[13px]">
        <thead>
          <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
            <th className="py-1.5 pr-3">{t("fis.colVoce")}</th>
            {p.cols.map((c) => (
              <th key={c.chiave} className="py-1.5 pr-3 text-right whitespace-nowrap">
                {c.label}
              </th>
            ))}
            <th className="py-1.5 text-right">{t("fis.colTotale")}</th>
          </tr>
        </thead>
        <tbody>
          {p.vociOrd.map((v) => {
            const r = p.perVoce.get(v)!;
            const tot = [...r.values()].reduce((a, x) => a + x, 0);
            return (
              <tr key={v} className="border-b border-border/40">
                <td className="py-1 pr-3 whitespace-nowrap">{v}</td>
                {p.cols.map((c) => (
                  <td key={c.chiave} className="py-1 pr-3 text-right tabular-nums">
                    {r.get(c.chiave) ? (
                      <button
                        type="button"
                        onClick={() => apriDettaglio(v, c.chiave, pagate)}
                        title={t("fis.dettaglioTip")}
                        className="rounded px-1 hover:bg-muted hover:text-primary"
                      >
                        {eur(r.get(c.chiave)!)}
                      </button>
                    ) : (
                      "—"
                    )}
                  </td>
                ))}
                <td className="py-1 text-right font-medium tabular-nums">{eur(tot)}</td>
              </tr>
            );
          })}
          <tr className="border-t border-border font-semibold">
            <td className="py-1.5 pr-3">{t("fis.colTotale")}</td>
            {p.cols.map((c) => (
              <td key={c.chiave} className="py-1.5 pr-3 text-right tabular-nums">
                {eur(p.vociOrd.reduce((a, v) => a + (p.perVoce.get(v)!.get(c.chiave) ?? 0), 0))}
              </td>
            ))}
            <td className="py-1.5 text-right tabular-nums">
              {eur(
                p.vociOrd.reduce(
                  (a, v) => a + [...p.perVoce.get(v)!.values()].reduce((x, y) => x + y, 0),
                  0,
                ),
              )}
            </td>
          </tr>
        </tbody>
      </table>
      <button
        type="button"
        onClick={() => esportaPivot(p, nomeCsv)}
        className="mt-2 rounded-lg border border-border px-3 py-1 text-[13px] hover:bg-muted"
      >
        {t("common.exportCsv")}
      </button>
    </div>
  );

  if (errore)
    return <p className="rounded-xl border border-border bg-card p-4 text-sm">{errore}</p>;
  if (!db)
    return (
      <div className="flex items-center gap-2 rounded-xl border border-border bg-card p-4 text-sm">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );

  return (
    <div className="space-y-4">
      {/* Datalist sempre nel DOM: lo usano sia il form scadenze sia il
          "da registrare" (dentro il form non esisteva a form chiuso). */}
      <datalist id="fis-voci">
        {voci.map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>
      {/* Riepilogo */}
      <div className="grid grid-cols-2 gap-3 lg:grid-cols-5">
        {[
          [t("fis.cardTotale"), riepilogo.totale - riepilogo.finanziamenti],
          [t("fis.cardRate"), riepilogo.rate],
          [t("fis.cardCorrenti"), riepilogo.correnti],
          [t("fis.cardFin"), riepilogo.finanziamenti],
        ].map(([lbl, v]) => (
          <div
            key={String(lbl)}
            className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]"
          >
            <p className="text-[11px] uppercase text-muted-foreground">{lbl}</p>
            <p className="mt-1 text-lg font-semibold tabular-nums">{eur(Number(v))} €</p>
          </div>
        ))}
        <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
          <p className="text-[11px] uppercase text-muted-foreground">{t("fis.cardUltima")}</p>
          <p className="mt-1 text-lg font-semibold">{dataIt(riepilogo.ultima) || "—"}</p>
        </div>
      </div>

      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-1 flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold">{t("fis.titolo")}</span>
          {busy && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
        </div>
        <p className="mb-3 text-xs text-muted-foreground">
          {t("fis.desc")}
          {db.fonteFile ? ` ${t("fis.fonte")} ${db.fonteFile}.` : ""}
        </p>
        <div className="mb-3 rounded-xl border border-border/60 p-3">
          <p className="mb-1 text-xs font-semibold">{t("fis.drTitolo")}</p>
          <p className="mb-2 text-[11px] text-muted-foreground">{t("fis.drNota")}</p>
          <div className="mb-2 flex flex-wrap gap-2">
            {db.daRateizzare.map((d, i) => (
              <span
                key={`${d.voce}|${d.anno ?? ""}|${d.periodo ?? ""}|${d.importo}|${i}`}
                className="inline-flex items-center gap-1.5 rounded-full bg-muted px-2.5 py-1 text-xs"
              >
                {d.voce} {d.periodo ?? ""} {d.anno ?? ""} — {eur(d.importo)} €
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void salvaDaRateizzare(db.daRateizzare.filter((_, j) => j !== i))}
                  title={t("common.delete")}
                >
                  <Trash2 className="h-3 w-3 text-muted-foreground hover:text-destructive" />
                </button>
              </span>
            ))}
            {db.daRateizzare.length === 0 && (
              <span className="text-xs text-muted-foreground">{t("fis.drVuoto")}</span>
            )}
          </div>
          <div className="flex flex-wrap items-end gap-2 text-[13px]">
            <input
              list="fis-voci"
              value={drVoce}
              onChange={(e) => setDrVoce(e.target.value)}
              placeholder={t("fis.colVoce")}
              className={`${inputCls} w-32`}
            />
            <input
              value={drAnno}
              onChange={(e) => setDrAnno(e.target.value)}
              placeholder={t("fis.colAnno")}
              className={`${inputCls} w-20`}
            />
            <input
              value={drPeriodo}
              onChange={(e) => setDrPeriodo(e.target.value)}
              placeholder={t("fis.colPeriodo")}
              className={`${inputCls} w-28`}
            />
            <input
              value={drImporto}
              onChange={(e) => setDrImporto(e.target.value)}
              placeholder="0,00"
              className={`${inputCls} w-28 text-right`}
            />
            <button
              type="button"
              disabled={busy}
              onClick={() => {
                const imp = numDaTesto(drImporto);
                if (!drVoce.trim() || !Number.isFinite(imp) || imp <= 0) {
                  toast.error(t("fis.errImporto"));
                  return;
                }
                const annoTxt = drAnno.trim();
                if (annoTxt && !/^\d{4}$/.test(annoTxt)) {
                  toast.error(t("fis.errAnno"));
                  return;
                }
                void salvaDaRateizzare([
                  ...db.daRateizzare,
                  {
                    voce: drVoce.trim().toUpperCase(),
                    anno: annoTxt ? Number(annoTxt) : undefined,
                    periodo: drPeriodo.trim() || undefined,
                    importo: Math.round(imp * 100) / 100,
                  },
                ]).then((ok) => {
                  if (!ok) return; // a salvataggio fallito i campi restano
                  setDrVoce("");
                  setDrAnno("");
                  setDrPeriodo("");
                  setDrImporto("");
                });
              }}
              className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
            >
              {t("fis.drAggiungi")}
            </button>
          </div>
        </div>

        {/* Selettore vista + filtri */}
        <div className="mb-3 flex flex-wrap items-center gap-2 text-[13px]">
          <div className="flex rounded-lg border border-border overflow-hidden">
            {(["scadenze", "dapagare", "pagato"] as const).map((v) => (
              <button
                key={v}
                type="button"
                onClick={() => setVista(v)}
                className={`px-3 py-1 ${vista === v ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
              >
                {t(
                  v === "scadenze"
                    ? "fis.vistaScadenze"
                    : v === "dapagare"
                      ? "fis.vistaDaPagare"
                      : "fis.vistaPagato",
                )}
              </button>
            ))}
          </div>
          {vista === "scadenze" && (
            <>
              <input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder={t("fis.cercaPh")}
                className={`${inputCls} w-52`}
              />
              <select
                value={statoF}
                onChange={(e) => setStatoF(e.target.value as typeof statoF)}
                className={inputCls}
              >
                <option value="aperte">{t("fis.filtroAperte")}</option>
                <option value="pagate">{t("fis.filtroPagate")}</option>
                <option value="tutte">{t("fis.filtroTutte")}</option>
              </select>
              <select value={voceF} onChange={(e) => setVoceF(e.target.value)} className={inputCls}>
                <option value="">{t("fis.filtroVoci")}</option>
                {voci.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              {meseF && (
                <span className="inline-flex items-center gap-1.5 rounded-full bg-primary/10 px-2.5 py-1 text-xs text-primary">
                  {meseF}
                  <button type="button" onClick={() => setMeseF("")} title={t("common.delete")}>
                    ×
                  </button>
                </span>
              )}
              <button
                type="button"
                onClick={() => setForm({ ...FORM_VUOTO, dataPagamento: oggiISO })}
                className="inline-flex items-center gap-1 rounded-lg bg-primary px-3 py-1 font-medium text-primary-foreground"
              >
                <Plus className="h-4 w-4" /> {t("fis.nuova")}
              </button>
              <button
                type="button"
                onClick={esportaScadenze}
                className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
              >
                {t("common.exportCsv")}
              </button>
            </>
          )}
        </div>

        {vista === "scadenze" && (
          <>
            {formUI}
            <div className="overflow-x-auto">
              <table className="w-full min-w-[900px] text-[13px]">
                <thead>
                  <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
                    <th className="py-1.5 pr-3">{t("fis.colData")}</th>
                    <th className="py-1.5 pr-3">{t("fis.colVoce")}</th>
                    <th className="py-1.5 pr-3">{t("fis.colDettaglio")}</th>
                    <th className="py-1.5 pr-3 text-right">{t("fis.colAnno")}</th>
                    <th className="py-1.5 pr-3">{t("fis.colPeriodo")}</th>
                    <th className="py-1.5 pr-3 text-right">{t("fis.colImporto")}</th>
                    <th className="py-1.5 pr-3">{t("fis.colModalita")}</th>
                    <th className="py-1.5 pr-3">{t("fis.colCategoria")}</th>
                    <th className="py-1.5 pr-3 text-center">{t("fis.colPagato")}</th>
                    <th className="py-1.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {filtrate.map((s) => (
                    <tr
                      key={s.id ?? `${s.voce}|${s.dataPagamento}|${s.importo}`}
                      className="border-b border-border/40"
                    >
                      <td className="py-1 pr-3 whitespace-nowrap">{dataIt(s.dataPagamento)}</td>
                      <td className="py-1 pr-3 whitespace-nowrap font-medium">{s.voce}</td>
                      <td
                        className="max-w-52 truncate py-1 pr-3 text-muted-foreground"
                        title={`${s.voceOld ?? ""} ${s.note ?? ""}`.trim()}
                      >
                        {s.voceOld ?? s.note ?? "—"}
                      </td>
                      <td className="py-1 pr-3 text-right tabular-nums">{s.anno ?? "—"}</td>
                      <td className="py-1 pr-3 whitespace-nowrap">{s.periodo ?? "—"}</td>
                      <td className="py-1 pr-3 text-right font-medium tabular-nums">
                        {eur(s.importo)}
                      </td>
                      <td className="py-1 pr-3">{s.modalita ?? "—"}</td>
                      <td className="py-1 pr-3">{badgeCat(s.categoria)}</td>
                      <td className="py-1 pr-3 text-center">
                        <input
                          type="checkbox"
                          checked={s.pagato}
                          disabled={busy || !s.id}
                          onChange={() => void togglePagato(s)}
                          title={t("fis.pagatoTip")}
                          className="accent-primary"
                        />
                      </td>
                      <td className="py-1 whitespace-nowrap">
                        <button
                          type="button"
                          disabled={!s.id}
                          onClick={() =>
                            setForm({
                              id: s.id,
                              voce: s.voce,
                              anno: s.anno != null ? String(s.anno) : "",
                              periodo: s.periodo ?? "",
                              importo: String(s.importo).replace(".", ","),
                              dataPagamento: s.dataPagamento,
                              modalita: s.modalita ?? "",
                              note: s.note ?? "",
                              pagato: s.pagato,
                            })
                          }
                          title={t("fis.modificaTip")}
                          className="mr-1 rounded p-1 hover:bg-muted disabled:opacity-30"
                        >
                          <Pencil className="h-3.5 w-3.5 text-muted-foreground" />
                        </button>
                        {confermaElimina === s.id ? (
                          <button
                            type="button"
                            disabled={busy}
                            onClick={() => void elimina(s.id!)}
                            className="rounded bg-destructive px-2 py-0.5 text-[11px] font-medium text-destructive-foreground"
                          >
                            {t("fis.confermaElimina")}
                          </button>
                        ) : (
                          <button
                            type="button"
                            disabled={!s.id}
                            onClick={() => setConfermaElimina(s.id!)}
                            title={t("common.delete")}
                            className="rounded p-1 hover:bg-muted disabled:opacity-30"
                          >
                            <Trash2 className="h-3.5 w-3.5 text-muted-foreground" />
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                  {filtrate.length === 0 && (
                    <tr>
                      <td colSpan={10} className="py-4 text-center text-muted-foreground">
                        {t("fis.vuoto")}
                      </td>
                    </tr>
                  )}
                </tbody>
                {filtrate.length > 0 && (
                  <tfoot>
                    <tr className="border-t border-border font-semibold">
                      <td colSpan={5} className="py-1.5 pr-3">
                        {filtrate.length} {t("fis.scadenzeLbl")}
                      </td>
                      <td className="py-1.5 pr-3 text-right tabular-nums">{eur(totFiltrate)}</td>
                      <td colSpan={4}></td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </>
        )}

        {vista === "dapagare" && pivotUI(pivotDaPagare, "fiscale-da-pagare", false)}
        {vista === "pagato" && pivotUI(pivotPagato, "fiscale-pagato", true)}
      </div>
    </div>
  );
}
