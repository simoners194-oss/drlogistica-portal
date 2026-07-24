// DR Portal — Finanza → tab Fatture (solo direttore DR005 + admin).
// Scadenzario fatture emesse: import dall'export Aruba, stato incasso
// calcolato dagli abbinamenti coi movimenti bancari, ritardi per termini di
// pagamento cliente, riconciliazione automatica + abbinamento manuale.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  KeyRound,
  Link2,
  Loader2,
  Plug,
  Trash2,
  Upload,
  Wand2,
} from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";
import {
  computeStatoFattura,
  individuaReinvii,
  parseFattureMatrice,
  parseFatturaPA,
  proponiAbbinamenti,
  isNotaCredito,
  isEsclusaDalCredito,
  TERMINI_DEFAULT_GIORNI,
  type AbbinamentoIncasso,
  type DirezioneFattura,
  type FatturaRaw,
  type TerminePagamento,
} from "@/lib/fatture-logic";
import { clienteGroupKey } from "@/lib/finanza-logic";
import {
  spGetFatture,
  spImportFatture,
  spGetTerminiPagamento,
  spGetAbbinamenti,
  spCreateAbbinamenti,
  spDeleteAbbinamento,
  spGetMovimenti,
  spGetArubaStato,
  spSetArubaCredenziali,
  spArubaProvaConnessione,
} from "@/lib/sharepoint.functions";
import type { SpFattura, SpMovimento, ArubaStato } from "@/lib/sharepoint.server";
import type { ArubaProbeResult } from "@/lib/aruba.server";

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}
function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function csvNum(n: number): string {
  return (Math.round(n * 100) / 100).toString().replace(".", ",");
}

const CHUNK = 100;

type StatoFiltro = "tutte" | "ritardo" | "nonIncassata" | "parziale" | "pagata";

export function FattureTab() {
  const { t } = useLang();
  // Direzione corrente della vista: Emesse (crediti) o Ricevute (debiti).
  const [dir, setDir] = useState<DirezioneFattura>("Emessa");
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const fatture = dir === "Emessa" ? fattureEm : fattureRic;
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [abbinamenti, setAbbinamenti] = useState<AbbinamentoIncasso[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);

  // Filtri
  const [annoF, setAnnoF] = useState(new Date().getFullYear());
  const [clienteF, setClienteF] = useState("");
  const [statoF, setStatoF] = useState<StatoFiltro>("tutte");

  // Dettaglio espanso + abbinamento manuale
  const [openFile, setOpenFile] = useState<string | null>(null);
  const [abbMov, setAbbMov] = useState("");
  const [abbImporto, setAbbImporto] = useState("");
  const [abbBusy, setAbbBusy] = useState(false);

  // Riconciliazione automatica
  const [reconciling, setReconciling] = useState(false);

  // Import fatture: ZIP/XML FatturaPA (emesse E ricevute, direzione automatica
  // dalla P.IVA) oppure xlsx dell'export "Check fatture inviate" (emesse).
  const [showImport, setShowImport] = useState(false);
  const [impParsing, setImpParsing] = useState(false);
  const [impBusy, setImpBusy] = useState(false);
  const [previewImp, setPreviewImp] = useState<{
    descrizione: string;
    emesse: FatturaRaw[];
    ricevute: FatturaRaw[];
    scartate: number;
  } | null>(null);

  // Collegamento Aruba (API, sola lettura)
  const [aruba, setAruba] = useState<ArubaStato | null>(null);
  const [showAruba, setShowAruba] = useState(false);
  const [arubaUser, setArubaUser] = useState("");
  const [arubaPass, setArubaPass] = useState("");
  const [arubaSaving, setArubaSaving] = useState(false);
  const [arubaTesting, setArubaTesting] = useState(false);
  const [probe, setProbe] = useState<ArubaProbeResult | null>(null);

  const load = () => {
    spGetFatture({ data: { direzione: "Emessa" } })
      .then((l) => setFattureEm(l as SpFattura[]))
      .catch((err) => {
        setFattureEm([]);
        toast.error(t("ft.errLoad"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    spGetAbbinamenti()
      .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
      .catch(() => setAbbinamenti([]));
    spGetMovimenti({ data: {} })
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
    spGetArubaStato()
      .then((s) => setAruba(s as ArubaStato))
      .catch(() => setAruba(null));
  };
  useEffect(load, []); // eslint-disable-line react-hooks/exhaustive-deps

  const oggiISO = new Date().toISOString().slice(0, 10);

  // Incassato per fattura dagli abbinamenti registrati.
  const incassatoPerFattura = useMemo(() => {
    const m = new Map<string, number>();
    for (const a of abbinamenti ?? [])
      m.set(a.fatturaFile, (m.get(a.fatturaFile) ?? 0) + a.importo);
    return m;
  }, [abbinamenti]);

  // Reinvii (stessa fattura scartata dallo SdI e rispedita): se ne conta una
  // sola, le altre sono escluse dal credito.
  const reinvii = useMemo(() => individuaReinvii(fatture ?? []), [fatture]);

  // Fatture con stato calcolato.
  const conStato = useMemo(
    () =>
      (fatture ?? []).map((f) => {
        const s = computeStatoFattura(
          f,
          incassatoPerFattura.get(f.nomeFile) ?? 0,
          termini,
          oggiISO,
        );
        return reinvii.has(f.nomeFile)
          ? { f, s: { ...s, stato: "NC" as const, residuo: 0, inRitardo: false, giorniRitardo: 0 } }
          : { f, s };
      }),
    [fatture, incassatoPerFattura, termini, oggiISO, reinvii],
  );

  const filtrate = useMemo(() => {
    let out = conStato;
    if (annoF > 0) out = out.filter((x) => Number(x.f.dataDocumento.slice(0, 4)) === annoF);
    if (clienteF.trim()) {
      const q = clienteF.trim().toLowerCase();
      out = out.filter(
        (x) => x.f.cliente.toLowerCase().includes(q) || x.f.numero.toLowerCase().includes(q),
      );
    }
    if (statoF === "ritardo") out = out.filter((x) => x.s.inRitardo);
    else if (statoF === "nonIncassata") out = out.filter((x) => x.s.stato === "Non incassata");
    else if (statoF === "parziale") out = out.filter((x) => x.s.stato === "Parziale");
    else if (statoF === "pagata") out = out.filter((x) => x.s.stato === "Pagata");
    return out;
  }, [conStato, annoF, clienteF, statoF]);

  // Riepilogo (sull'anno filtrato, tutte le fatture non escluse).
  const riepilogo = useMemo(() => {
    const base = conStato.filter(
      (x) => (annoF <= 0 || Number(x.f.dataDocumento.slice(0, 4)) === annoF) && x.s.stato !== "NC",
    );
    const residuo = base.reduce((s, x) => s + x.s.residuo, 0);
    const inRitardo = base.filter((x) => x.s.inRitardo);
    const ritardoImporto = inRitardo.reduce((s, x) => s + x.s.residuo, 0);
    const incassato = base.reduce((s, x) => s + x.s.incassato, 0);
    return { n: base.length, residuo, nRitardo: inRitardo.length, ritardoImporto, incassato };
  }, [conStato, annoF]);

  // Riepilogo per cliente (come l'OVERVIEW del direttore, compattata).
  const perCliente = useMemo(() => {
    const m = new Map<
      string,
      { cliente: string; aperte: number; residuo: number; ritardo: number }
    >();
    for (const x of conStato) {
      if (x.s.stato === "NC" || x.s.residuo <= 0) continue;
      if (annoF > 0 && Number(x.f.dataDocumento.slice(0, 4)) !== annoF) continue;
      const key = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const row = m.get(key) ?? { cliente: x.f.cliente, aperte: 0, residuo: 0, ritardo: 0 };
      row.aperte++;
      row.residuo += x.s.residuo;
      if (x.s.inRitardo) row.ritardo += x.s.residuo;
      m.set(key, row);
    }
    return [...m.values()].sort((a, b) => b.residuo - a.residuo);
  }, [conStato, annoF]);

  // Movimenti con residuo da allocare per l'abbinamento manuale: incassi per
  // le emesse, uscite (in valore assoluto) per le ricevute.
  const incassiDisponibili = useMemo(() => {
    const allocato = new Map<string, number>();
    for (const a of abbinamenti ?? [])
      allocato.set(a.movimentoChiave, (allocato.get(a.movimentoChiave) ?? 0) + a.importo);
    return (movimenti ?? [])
      .filter((m) =>
        dir === "Emessa" ? m.importo > 0 && m.tipologia === "Incasso" : m.importo < 0,
      )
      .map((m) => ({
        m,
        residuo: Math.round((Math.abs(m.importo) - (allocato.get(m.chiave) ?? 0)) * 100) / 100,
      }))
      .filter((x) => x.residuo > 0.01)
      .sort((a, b) => b.m.dataContabile.localeCompare(a.m.dataContabile));
  }, [movimenti, abbinamenti, dir]);

  // --- Riconciliazione automatica ------------------------------------------
  const riconcilia = async () => {
    if (!fatture || !movimenti || !abbinamenti) return;
    setReconciling(true);
    try {
      const proposte = proponiAbbinamenti(
        fatture.filter((f) => !reinvii.has(f.nomeFile)),
        movimenti.map((m) => ({
          chiave: m.chiave,
          dataContabile: m.dataContabile,
          importo: m.importo,
          tipologia: m.tipologia,
          cliente: m.cliente,
          descrizione: m.descrizione,
          nrFattura: m.nrFattura,
        })),
        abbinamenti,
        dir,
      );
      if (proposte.length === 0) {
        toast(t("ft.reconcileNone"));
        return;
      }
      let creati = 0;
      for (let i = 0; i < proposte.length; i += CHUNK) {
        const res = (await spCreateAbbinamenti({
          data: {
            rows: proposte.slice(i, i + CHUNK).map(({ motivo: _m, ...r }) => r),
          },
        })) as { creati: number };
        creati += res.creati;
      }
      toast.success(t("ft.reconcileDone"), {
        description: `${creati} ${t("ft.reconcileCount")}`,
      });
      spGetAbbinamenti()
        .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setReconciling(false);
    }
  };

  // --- Abbinamento manuale --------------------------------------------------
  const abbinaManuale = async (fatturaFile: string, residuoFattura: number) => {
    const inc = incassiDisponibili.find((x) => x.m.chiave === abbMov);
    if (!inc) return toast.error(t("ft.abbSelect"));
    const importo = Number(abbImporto.replace(",", "."));
    if (!Number.isFinite(importo) || importo <= 0) return toast.error(t("ft.abbImporto"));
    if (importo > inc.residuo + 0.01 || importo > residuoFattura + 0.01)
      return toast.error(t("ft.abbTroppo"));
    setAbbBusy(true);
    try {
      await spCreateAbbinamenti({
        data: {
          rows: [{ fatturaFile, movimentoChiave: inc.m.chiave, importo, origine: "Manuale" }],
        },
      });
      toast.success(t("ft.abbDone"));
      setAbbMov("");
      setAbbImporto("");
      spGetAbbinamenti()
        .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setAbbBusy(false);
    }
  };

  const rimuoviAbbinamento = async (a: AbbinamentoIncasso) => {
    if (!a.id) return;
    if (!window.confirm(t("ft.abbDeleteConfirm"))) return;
    try {
      await spDeleteAbbinamento({ data: { abbinamentoId: a.id } });
      setAbbinamenti((prev) => (prev ?? []).filter((x) => x.id !== a.id));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    }
  };

  // --- Import fatture (ZIP/XML FatturaPA + xlsx emesse) ---------------------
  const onFiles = async (files: File[]) => {
    setPreviewImp(null);
    setImpParsing(true);
    try {
      const rows: FatturaRaw[] = [];
      let scartate = 0;
      const decoder = new TextDecoder("utf-8");
      const daXml = (testo: string, nome: string) => {
        const res = parseFatturaPA(testo, nome);
        rows.push(...res.rows);
        scartate += res.scartati.length;
      };
      for (const file of files) {
        const nome = file.name;
        if (/\.zip$/i.test(nome)) {
          const { unzipSync } = await import("fflate");
          const entries = unzipSync(new Uint8Array(await file.arrayBuffer()));
          for (const [entry, bytes] of Object.entries(entries)) {
            if (!/\.xml$/i.test(entry)) continue;
            daXml(decoder.decode(bytes), entry.split("/").pop() ?? entry);
          }
        } else if (/\.xml$/i.test(nome)) {
          daXml(await file.text(), nome);
        } else if (/\.xlsx?$/i.test(nome)) {
          const XLSX = await import("xlsx");
          const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
          let trovato = false;
          for (const sheet of wb.SheetNames) {
            const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
              header: 1,
              raw: true,
            }) as unknown[][];
            const res = parseFattureMatrice(matrix);
            if (res && res.rows.length) {
              rows.push(...res.rows);
              scartate += res.scartate;
              trovato = true;
              break;
            }
          }
          if (!trovato) scartate++;
        } else {
          scartate++;
        }
      }
      if (rows.length === 0) {
        toast.error(t("ft.errFile"), { description: t("ft.errFileDesc") });
        return;
      }
      // Dedup nel caricamento stesso (es. ZIP + xlsx insieme): prima vince.
      const visti = new Set<string>();
      const univoche = rows.filter((r) => {
        if (visti.has(r.nomeFile)) {
          scartate++;
          return false;
        }
        visti.add(r.nomeFile);
        return true;
      });
      setPreviewImp({
        descrizione: files.length === 1 ? files[0].name : `${files.length} file`,
        emesse: univoche.filter((r) => r.direzione === "Emessa"),
        ricevute: univoche.filter((r) => r.direzione === "Ricevuta"),
        scartate,
      });
    } catch (err) {
      toast.error(t("ft.errFile"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImpParsing(false);
    }
  };

  const eseguiImport = async () => {
    if (!previewImp) return;
    setImpBusy(true);
    try {
      let nuove = 0;
      let aggiornate = 0;
      let doppioni = 0;
      const errori: string[] = [];
      const importaGruppo = async (gruppo: FatturaRaw[], direzione: DirezioneFattura) => {
        for (let i = 0; i < gruppo.length; i += CHUNK) {
          const res = (await spImportFatture({
            data: { rows: gruppo.slice(i, i + CHUNK), direzione },
          })) as { importate: number; aggiornate: number; doppioni: number; errori: string[] };
          nuove += res.importate;
          aggiornate += res.aggiornate;
          doppioni += res.doppioni;
          errori.push(...res.errori);
        }
      };
      if (previewImp.emesse.length) await importaGruppo(previewImp.emesse, "Emessa");
      if (previewImp.ricevute.length) await importaGruppo(previewImp.ricevute, "Ricevuta");
      toast.success(t("ft.importDone"), {
        description: `${nuove} ${t("ft.importNew")} · ${aggiornate} ${t("ft.importUpd")} · ${doppioni} ${t("ft.importDup")}${errori.length ? ` · ${errori.length} ${t("common.error").toLowerCase()}` : ""}`,
      });
      setPreviewImp(null);
      setShowImport(false);
      load();
    } catch (err) {
      toast.error(t("ft.errImport"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImpBusy(false);
    }
  };

  // --- Collegamento Aruba ---------------------------------------------------
  const salvaCredenziali = async () => {
    if (!arubaUser.trim() || !arubaPass) return toast.error(t("ft.arCredMancanti"));
    setArubaSaving(true);
    try {
      await spSetArubaCredenziali({ data: { username: arubaUser.trim(), password: arubaPass } });
      toast.success(t("ft.arCredSalvate"));
      setArubaPass("");
      spGetArubaStato()
        .then((s) => setAruba(s as ArubaStato))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setArubaSaving(false);
    }
  };

  const provaConnessione = async () => {
    setArubaTesting(true);
    setProbe(null);
    try {
      const res = (await spArubaProvaConnessione()) as ArubaProbeResult;
      setProbe(res);
      toast.success(t("ft.arProvaOk"));
    } catch (err) {
      setProbe({
        ok: false,
        messaggio: err instanceof Error ? err.message : String(err),
      });
      toast.error(t("ft.arProvaKo"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setArubaTesting(false);
    }
  };

  const esporta = () => {
    esportaCsvFile(
      `fatture-${dir === "Ricevuta" ? "ricevute" : "emesse"}-${annoF > 0 ? annoF : "tutte"}`,
      [
        "Numero",
        "Data documento",
        "Cliente",
        "Tipo",
        "Totale",
        "Incassato",
        "Residuo",
        "Scadenza",
        "Stato",
        "Ritardo gg",
        "Stato SdI",
        "Nome file",
      ],
      filtrate.map(({ f, s }) => [
        f.numero,
        fmtData(f.dataDocumento),
        f.cliente,
        f.tipoDocumento,
        csvNum(f.totale),
        csvNum(s.incassato),
        csvNum(s.residuo),
        fmtData(s.scadenza),
        s.stato,
        s.inRitardo ? s.giorniRitardo : "",
        f.statoSdI,
        f.nomeFile,
      ]),
    );
  };

  const badge = (x: (typeof conStato)[number]) => {
    if (x.s.stato === "NC")
      return (
        <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
          {reinvii.has(x.f.nomeFile)
            ? t("ft.reinvio")
            : isNotaCredito(x.f.tipoDocumento)
              ? t("ft.nc")
              : isEsclusaDalCredito(x.f)
                ? x.f.statoSdI
                : "—"}
        </span>
      );
    if (x.s.stato === "Pagata")
      return (
        <span className="rounded-full bg-status-present/15 px-2 py-0.5 text-[11px] font-medium text-status-present">
          {t("ft.pagata")}
        </span>
      );
    return (
      <span
        className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${x.s.inRitardo ? "bg-status-absent/15 text-status-absent" : "bg-status-break/15 text-status-break"}`}
      >
        {x.s.stato === "Parziale"
          ? t("ft.parziale")
          : x.f.direzione === "Ricevuta"
            ? t("ft.nonPagata")
            : t("ft.nonIncassata")}
        {x.s.inRitardo ? ` · +${x.s.giorniRitardo}gg` : ""}
      </span>
    );
  };

  const loading = fatture == null || abbinamenti == null || movimenti == null;

  const ricevute = dir === "Ricevuta";

  return (
    <div className="space-y-4">
      {/* Direzione: emesse (crediti) / ricevute (debiti) */}
      <div className="inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => {
            setDir("Emessa");
            setOpenFile(null);
          }}
          className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${!ricevute ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          {t("ft.dirEmesse")}
        </button>
        <button
          type="button"
          onClick={() => {
            setDir("Ricevuta");
            setOpenFile(null);
          }}
          className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${ricevute ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          {t("ft.dirRicevute")}
        </button>
      </div>

      {/* Riepilogo */}
      <div className="grid gap-3 sm:grid-cols-3">
        <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {ricevute ? t("ft.kpiDaPagare") : t("ft.kpiAperto")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-foreground">
            {fmtImporto(riepilogo.residuo)} €
          </div>
        </div>
        <div className="rounded-2xl border border-status-absent/40 bg-status-absent/5 p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {t("ft.kpiRitardo")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-status-absent">
            {fmtImporto(riepilogo.ritardoImporto)} €
          </div>
          <div className="text-[11px] text-muted-foreground mt-0.5">
            {riepilogo.nRitardo} {t("ft.kpiRitardoN")}
          </div>
        </div>
        <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {ricevute ? t("ft.kpiPagato") : t("ft.kpiIncassato")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-status-present">
            {fmtImporto(riepilogo.incassato)} €
          </div>
        </div>
      </div>

      {/* Azioni + filtri */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-3 mb-4">
          <div className="w-28">
            <label className="text-xs text-muted-foreground">{t("ft.anno")}</label>
            <select
              value={annoF}
              onChange={(e) => setAnnoF(Number(e.target.value))}
              className={inputCls}
            >
              <option value={0}>{t("fin.allYears")}</option>
              {Array.from({ length: 4 }, (_, i) => new Date().getFullYear() - i).map((a) => (
                <option key={a} value={a}>
                  {a}
                </option>
              ))}
            </select>
          </div>
          <div className="w-44">
            <label className="text-xs text-muted-foreground">{t("common.status")}</label>
            <select
              value={statoF}
              onChange={(e) => setStatoF(e.target.value as StatoFiltro)}
              className={inputCls}
            >
              <option value="tutte">{t("common.allF")}</option>
              <option value="ritardo">{t("ft.fRitardo")}</option>
              <option value="nonIncassata">
                {ricevute ? t("ft.nonPagata") : t("ft.nonIncassata")}
              </option>
              <option value="parziale">{t("ft.parziale")}</option>
              <option value="pagata">{t("ft.pagata")}</option>
            </select>
          </div>
          <div className="flex-1 min-w-44">
            <label className="text-xs text-muted-foreground">{t("ft.cerca")}</label>
            <input
              value={clienteF}
              onChange={(e) => setClienteF(e.target.value)}
              placeholder={t("ft.cercaPh")}
              className={inputCls}
            />
          </div>
          <button
            type="button"
            onClick={() => void riconcilia()}
            disabled={loading || reconciling}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
            title={t("ft.reconcileTip")}
          >
            {reconciling ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <Wand2 className="h-4 w-4" />
            )}
            {t("ft.reconcile")}
          </button>
          <button
            type="button"
            onClick={() => {
              setShowImport((v) => !v);
              setShowAruba(false);
            }}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("ft.import")}
          </button>
          <button
            type="button"
            onClick={() => {
              setShowAruba((v) => !v);
              setShowImport(false);
            }}
            className={`inline-flex items-center gap-2 rounded-lg border px-3 py-2 text-sm hover:bg-muted ${aruba?.configurato ? "border-status-present/40 text-status-present" : "border-border text-foreground"}`}
          >
            <Plug className="h-4 w-4" /> Aruba
            {aruba?.configurato && <CheckCircle2 className="h-3.5 w-3.5" />}
          </button>
          <button
            type="button"
            onClick={esporta}
            disabled={filtrate.length === 0}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
          >
            <Download className="h-4 w-4" /> {t("common.exportCsv")}
          </button>
        </div>

        {/* Import export Aruba (a scomparsa) */}
        {showImport && (
          <div className="mb-4 rounded-xl border border-border p-4">
            <div className="flex items-center gap-2 text-sm font-medium text-foreground mb-1">
              <FileSpreadsheet className="h-4 w-4 text-primary" /> {t("ft.importTitle")}
            </div>
            <p className="text-xs text-muted-foreground mb-3">{t("ft.importDesc")}</p>
            <input
              type="file"
              accept=".zip,.xml,.xlsx,.xls"
              multiple
              disabled={impParsing || impBusy}
              onChange={(e) => {
                const files = Array.from(e.target.files ?? []);
                if (files.length) void onFiles(files);
                e.target.value = "";
              }}
              className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
            />
            {impParsing && (
              <p className="mt-2 text-sm text-muted-foreground inline-flex items-center gap-2">
                <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
              </p>
            )}
            {previewImp && (
              <div className="mt-3 text-[13px] text-muted-foreground">
                <b className="text-foreground">{previewImp.descrizione}</b> —{" "}
                {previewImp.emesse.length + previewImp.ricevute.length} {t("ft.importRows")} (
                {previewImp.emesse.length} {t("ft.dirEmesse").toLowerCase()},{" "}
                {previewImp.ricevute.length} {t("ft.dirRicevute").toLowerCase()})
                {previewImp.scartate > 0 && ` · ${previewImp.scartate} ${t("fin.previewSkipped")}`}
                <div className="mt-2 flex gap-2">
                  <button
                    type="button"
                    onClick={() => void eseguiImport()}
                    disabled={impBusy}
                    className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                  >
                    {impBusy ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <Upload className="h-4 w-4" />
                    )}
                    {t("fin.importBtn")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setPreviewImp(null)}
                    className="rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                  >
                    {t("common.cancel")}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Collegamento Aruba (a scomparsa) */}
        {showAruba && (
          <div className="mb-4 rounded-xl border border-border p-4">
            <div className="flex items-center gap-2 text-sm font-medium text-foreground mb-1">
              <Plug className="h-4 w-4 text-primary" /> {t("ft.arTitle")}
            </div>
            <p className="text-xs text-muted-foreground mb-3">
              {aruba == null
                ? t("common.loading")
                : !aruba.listaPresente
                  ? t("ft.arNoLista")
                  : aruba.configurato
                    ? `${t("ft.arConfigurato")} (${aruba.username})`
                    : t("ft.arDaConfigurare")}
            </p>
            {aruba?.listaPresente && (
              <div className="flex flex-wrap items-end gap-3">
                <div className="min-w-52">
                  <label className="text-xs text-muted-foreground">{t("ft.arUser")}</label>
                  <input
                    value={arubaUser}
                    onChange={(e) => setArubaUser(e.target.value)}
                    autoComplete="off"
                    className={inputCls}
                  />
                </div>
                <div className="min-w-52">
                  <label className="text-xs text-muted-foreground">{t("ft.arPass")}</label>
                  <input
                    type="password"
                    value={arubaPass}
                    onChange={(e) => setArubaPass(e.target.value)}
                    autoComplete="new-password"
                    className={inputCls}
                  />
                </div>
                <button
                  type="button"
                  onClick={() => void salvaCredenziali()}
                  disabled={arubaSaving || !arubaUser.trim() || !arubaPass}
                  className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
                >
                  {arubaSaving ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <KeyRound className="h-4 w-4" />
                  )}
                  {t("ft.arSalva")}
                </button>
                <button
                  type="button"
                  onClick={() => void provaConnessione()}
                  disabled={arubaTesting || !aruba.configurato}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {arubaTesting ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Plug className="h-4 w-4" />
                  )}
                  {t("ft.arProva")}
                </button>
              </div>
            )}
            {probe && (
              <div
                className={`mt-3 rounded-lg p-3 text-[13px] ${probe.ok ? "bg-status-present/10 text-foreground" : "bg-status-absent/10 text-status-absent"}`}
              >
                <div className="font-medium">{probe.messaggio}</div>
                {probe.ok && probe.campiEsempio && (
                  <div className="mt-2 text-muted-foreground">
                    <div className="text-xs font-medium text-foreground mb-1">
                      {t("ft.arCampi")}
                    </div>
                    <div className="overflow-x-auto">
                      <table className="text-xs">
                        <tbody>
                          {Object.entries(probe.campiEsempio).map(([k, v]) => (
                            <tr key={k}>
                              <td className="pr-3 py-0.5 font-mono text-foreground whitespace-nowrap">
                                {k}
                              </td>
                              <td className="py-0.5 break-all">{v}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
                {probe.ok && probe.elementi === 0 && (
                  <p className="mt-1 text-muted-foreground">{t("ft.arVuoto")}</p>
                )}
              </div>
            )}
            <p className="mt-3 text-[11px] text-muted-foreground">{t("ft.arNota")}</p>
          </div>
        )}

        {/* Elenco */}
        {loading ? (
          <div className="py-10 text-center text-sm text-muted-foreground">
            <Loader2 className="h-5 w-5 animate-spin inline-block" />
          </div>
        ) : filtrate.length === 0 ? (
          <p className="py-8 text-center text-sm text-muted-foreground">{t("ft.empty")}</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-muted-foreground border-b border-border">
                  <th className="py-2 pr-3">{t("ft.numero")}</th>
                  <th className="py-2 pr-3">{t("common.date")}</th>
                  <th className="py-2 pr-3">{ricevute ? t("ft.fornitore") : t("fin.cliente")}</th>
                  <th className="py-2 pr-3 text-right">{t("common.total")}</th>
                  <th className="py-2 pr-3 text-right">
                    {ricevute ? t("ft.pagato") : t("ft.incassato")}
                  </th>
                  <th className="py-2 pr-3 text-right">{t("ft.residuo")}</th>
                  <th className="py-2 pr-3">{t("ft.scadenza")}</th>
                  <th className="py-2 pr-3">{t("common.status")}</th>
                </tr>
              </thead>
              <tbody>
                {filtrate.slice(0, 400).map((x) => {
                  const abbFat = (abbinamenti ?? []).filter((a) => a.fatturaFile === x.f.nomeFile);
                  const aperta = openFile === x.f.nomeFile;
                  return [
                    <tr
                      key={x.f.nomeFile}
                      onClick={() => {
                        setOpenFile(aperta ? null : x.f.nomeFile);
                        setAbbMov("");
                        setAbbImporto("");
                      }}
                      className={`border-b border-border/50 cursor-pointer hover:bg-muted/40 ${aperta ? "bg-muted/30" : ""}`}
                      title={x.f.nomeFile}
                    >
                      <td className="py-1.5 pr-3 whitespace-nowrap font-medium">{x.f.numero}</td>
                      <td className="py-1.5 pr-3 whitespace-nowrap">
                        {fmtData(x.f.dataDocumento)}
                      </td>
                      <td className="py-1.5 pr-3 max-w-52 truncate">{x.f.cliente}</td>
                      <td className="py-1.5 pr-3 text-right whitespace-nowrap">
                        {fmtImporto(x.f.totale)}
                      </td>
                      <td className="py-1.5 pr-3 text-right whitespace-nowrap text-status-present">
                        {x.s.incassato ? fmtImporto(x.s.incassato) : ""}
                      </td>
                      <td className="py-1.5 pr-3 text-right whitespace-nowrap font-medium">
                        {x.s.stato === "NC" ? "" : fmtImporto(x.s.residuo)}
                      </td>
                      <td
                        className={`py-1.5 pr-3 whitespace-nowrap ${x.s.inRitardo ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                      >
                        {x.s.stato === "NC" ? "—" : fmtData(x.s.scadenza)}
                      </td>
                      <td className="py-1.5 pr-3">{badge(x)}</td>
                    </tr>,
                    aperta && (
                      <tr key={`${x.f.nomeFile}-det`} className="border-b border-border/50">
                        <td colSpan={8} className="py-3 px-3 bg-muted/20">
                          <div className="text-xs text-muted-foreground mb-2">
                            {x.f.tipoDocumento} · SdI {x.f.statoSdI || "—"} · {t("ft.terminiGg")}{" "}
                            {termini.length
                              ? `${(x.s.scadenza && x.f.dataDocumento && Math.round((new Date(x.s.scadenza).getTime() - new Date(x.f.dataDocumento).getTime()) / 86400000)) || TERMINI_DEFAULT_GIORNI}gg`
                              : `${TERMINI_DEFAULT_GIORNI}gg (default)`}
                          </div>
                          {abbFat.length > 0 ? (
                            <ul className="space-y-1 mb-3">
                              {abbFat.map((a) => {
                                const mov = (movimenti ?? []).find(
                                  (m) => m.chiave === a.movimentoChiave,
                                );
                                return (
                                  <li key={a.id} className="flex items-center gap-3 text-[13px]">
                                    <Link2 className="h-3.5 w-3.5 text-status-present shrink-0" />
                                    <span className="tabular-nums font-medium">
                                      {fmtImporto(a.importo)} €
                                    </span>
                                    <span className="text-muted-foreground truncate">
                                      {mov
                                        ? `${fmtData(mov.dataContabile)} · ${mov.cliente || mov.descrizione.slice(0, 50)}`
                                        : a.movimentoChiave.slice(0, 60)}
                                    </span>
                                    <span className="text-[11px] rounded-full bg-muted px-1.5">
                                      {a.origine}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        void rimuoviAbbinamento(a);
                                      }}
                                      className="ml-auto rounded-md p-1 text-muted-foreground hover:text-status-absent"
                                      title={t("ft.abbDelete")}
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </li>
                                );
                              })}
                            </ul>
                          ) : (
                            <p className="text-[13px] text-muted-foreground mb-3">
                              {t("ft.abbNone")}
                            </p>
                          )}
                          {x.s.stato !== "NC" && x.s.residuo > 0.01 && (
                            <div
                              className="flex flex-wrap items-end gap-2"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <div className="min-w-72 flex-1">
                                <label className="text-xs text-muted-foreground">
                                  {ricevute ? t("ft.abbPagamento") : t("ft.abbMovimento")}
                                </label>
                                <select
                                  value={abbMov}
                                  onChange={(e) => {
                                    setAbbMov(e.target.value);
                                    const inc = incassiDisponibili.find(
                                      (i) => i.m.chiave === e.target.value,
                                    );
                                    if (inc)
                                      setAbbImporto(
                                        String(Math.min(inc.residuo, x.s.residuo)).replace(
                                          ".",
                                          ",",
                                        ),
                                      );
                                  }}
                                  className={inputCls}
                                >
                                  <option value="">{t("common.select")}</option>
                                  {incassiDisponibili.slice(0, 200).map((i) => (
                                    <option key={i.m.chiave} value={i.m.chiave}>
                                      {fmtData(i.m.dataContabile)} · {fmtImporto(i.residuo)} € ·{" "}
                                      {(i.m.cliente || i.m.descrizione).slice(0, 60)}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="w-32">
                                <label className="text-xs text-muted-foreground">
                                  {t("common.amount")}
                                </label>
                                <input
                                  value={abbImporto}
                                  onChange={(e) => setAbbImporto(e.target.value)}
                                  className={inputCls}
                                />
                              </div>
                              <button
                                type="button"
                                onClick={() => void abbinaManuale(x.f.nomeFile, x.s.residuo)}
                                disabled={abbBusy || !abbMov}
                                className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                              >
                                {abbBusy ? (
                                  <Loader2 className="h-4 w-4 animate-spin" />
                                ) : (
                                  <Link2 className="h-4 w-4" />
                                )}
                                {t("ft.abbina")}
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    ),
                  ];
                })}
              </tbody>
            </table>
            {filtrate.length > 400 && (
              <div className="mt-2 text-xs text-muted-foreground">
                {t("fin.first500")} {filtrate.length}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Riepilogo per cliente */}
      {perCliente.length > 0 && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-sm font-semibold text-foreground mb-3">
            <AlertTriangle className="h-4 w-4 text-status-absent" />{" "}
            {ricevute ? t("ft.perFornitoreTitle") : t("ft.perClienteTitle")}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-muted-foreground border-b border-border">
                  <th className="py-2 pr-3">{t("fin.cliente")}</th>
                  <th className="py-2 pr-3 text-right">{t("ft.aperte")}</th>
                  <th className="py-2 pr-3 text-right">{t("ft.residuo")}</th>
                  <th className="py-2 pr-3 text-right">{t("ft.diCuiRitardo")}</th>
                </tr>
              </thead>
              <tbody>
                {perCliente.map((r) => (
                  <tr key={r.cliente} className="border-b border-border/50">
                    <td className="py-1.5 pr-3 max-w-64 truncate">{r.cliente}</td>
                    <td className="py-1.5 pr-3 text-right">{r.aperte}</td>
                    <td className="py-1.5 pr-3 text-right whitespace-nowrap font-medium">
                      {fmtImporto(r.residuo)}
                    </td>
                    <td
                      className={`py-1.5 pr-3 text-right whitespace-nowrap ${r.ritardo > 0 ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                    >
                      {r.ritardo > 0 ? fmtImporto(r.ritardo) : "—"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="flex items-start gap-2 text-xs text-muted-foreground">
        <CheckCircle2 className="h-4 w-4 shrink-0 mt-0.5" />
        <p>{t("ft.nota")}</p>
      </div>
    </div>
  );
}
