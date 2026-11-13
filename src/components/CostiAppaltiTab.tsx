// DR Portal — Finanze → "Costi appalti" (Simone 13/11/2026: "dobbiamo
// integrare un nuovo tab con il foglio in allegato, che posso filtrare i vari
// clienti per vedere i costi").
//
// Il file settimanale dei costi ("Costi WK45.xlsx": un foglio di Excel per
// appalto, più il Totale) si carica qui e diventa una pagina: i quattro
// numeri della settimana (ricavi, costi, MOL 1, margine), i clienti da
// accendere e spegnere, il conto voce per voce con le note del file, e
// l'andamento settimana dopo settimana man mano che i file si accumulano.
// I totali si RICALCOLANO dalle voci; quelli scritti nel file servono solo da
// controllo al caricamento.
import { useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { AlertTriangle, Calculator, Loader2, Trash2, Upload } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import {
  chiaveAppalto,
  controlliAppalto,
  leggiFoglioCosti,
  leggiPeriodo,
  righeDaCelle,
  settimanaIso,
  totaliAppalto,
  totaliInsieme,
  voceDi,
  vociDi,
  type AppaltoCosti,
  type CostiAppaltiDb,
  type SettimanaCosti,
  type TotaliAppalto,
} from "@/lib/costi-appalti-logic";
import {
  spCostiAppaltiElimina,
  spCostiAppaltiGet,
  spCostiAppaltiSalva,
} from "@/lib/costi-appalti.functions";

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmt0 = (n: number) => n.toLocaleString("it-IT", { maximumFractionDigits: 0 });
const pct = (n: number | null) =>
  n == null
    ? "—"
    : `${(n * 100).toLocaleString("it-IT", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
const ddmm = (iso: string) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}`;
const dataIt = (iso: string) => `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`;
const oggiISO = () => new Date().toISOString().slice(0, 10);

/** Verde se il margine regge, ambra se è sottile, rosso se si perde. */
const tintaMargine = (m: number | null) =>
  m == null
    ? "text-muted-foreground"
    : m < 0
      ? "text-rose-600 dark:text-rose-400"
      : m < 0.05
        ? "text-amber-600 dark:text-amber-400"
        : "text-emerald-700 dark:text-emerald-400";
const puntoMargine = (m: number | null) =>
  m == null ? "bg-slate-400" : m < 0 ? "bg-rose-500" : m < 0.05 ? "bg-amber-500" : "bg-emerald-500";

/** Il file appena letto, in attesa di conferma. */
export interface Bozza {
  file: string;
  etichetta: string;
  dal: string;
  al: string;
  appalti: AppaltoCosti[];
  /** Fogli del file che non sono appalti (il "Totale"). */
  saltati: string[];
}

const inputCls =
  "rounded-lg border border-border bg-background px-2 py-1 text-sm text-foreground focus:border-primary focus:outline-none";
const card = "rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]";

export function CostiAppaltiTab({
  prova,
}: {
  /** Solo per l'anteprima locale della pagina (senza server): i dati già
   *  pronti, i clienti accesi e un file "appena letto". Il portale non la usa. */
  prova?: { db: CostiAppaltiDb; scelti?: string[]; bozza?: Bozza };
} = {}) {
  const { t } = useLang();
  const [db, setDb] = useState<CostiAppaltiDb | null>(prova?.db ?? null);
  const [errore, setErrore] = useState<string | null>(null);
  const [settimanaId, setSettimanaId] = useState<string>("");
  /** Clienti accesi (chiavi); vuoto = tutti. */
  const [scelti, setScelti] = useState<Set<string>>(
    () => new Set((prova?.scelti ?? []).map(chiaveAppalto)),
  );
  const [vista, setVista] = useState<"importi" | "percento">("importi");
  const [ancheZero, setAncheZero] = useState(false);
  const [bozza, setBozza] = useState<Bozza | null>(prova?.bozza ?? null);
  const [busy, setBusy] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);

  const carica = () => {
    setErrore(null);
    spCostiAppaltiGet()
      .then((d) => setDb(d))
      .catch((e) => setErrore(e instanceof Error ? e.message : String(e)));
  };
  useEffect(() => {
    if (!prova) carica();
    // Una volta sola, all'apertura.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const settimane = useMemo(() => db?.settimane ?? [], [db]);
  const settimana: SettimanaCosti | undefined =
    settimane.find((s) => s.id === settimanaId) ?? settimane[0];

  // I clienti della settimana, e quelli accesi (vuoto = tutti).
  const appalti = useMemo(() => settimana?.appalti ?? [], [settimana]);
  const visibili = useMemo(() => {
    const l = appalti.filter((a) => scelti.has(chiaveAppalto(a.nome)));
    return l.length ? l : appalti;
  }, [appalti, scelti]);
  const tutti = visibili.length === appalti.length;
  const totali = useMemo(() => totaliInsieme(visibili), [visibili]);
  const uno = visibili.length === 1 ? visibili[0] : undefined;

  const toggle = (nome: string) =>
    setScelti((prima) => {
      const k = chiaveAppalto(nome);
      // Da "tutti" il primo clic isola quel cliente; poi si aggiungono.
      const base = new Set(
        [...prima].filter((x) => appalti.some((a) => chiaveAppalto(a.nome) === x)),
      );
      if (base.has(k)) base.delete(k);
      else base.add(k);
      return base.size === appalti.length ? new Set() : base;
    });

  // --- caricamento del file ---------------------------------------------------
  const leggiFile = async (file: File) => {
    setBusy(true);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await file.arrayBuffer());
      const letti: AppaltoCosti[] = [];
      const saltati: string[] = [];
      let etichetta = "";
      for (const nome of wb.SheetNames) {
        const letto = leggiFoglioCosti(
          nome,
          righeDaCelle(wb.Sheets[nome] as unknown as Record<string, unknown>),
        );
        if (!letto) {
          saltati.push(nome);
          continue;
        }
        letti.push(letto.appalto);
        etichetta ||= letto.etichetta;
      }
      if (!letti.length) {
        toast.error(t("ca.errNessunFoglio"), { description: file.name });
        return;
      }
      const periodo = leggiPeriodo(etichetta, file.name, oggiISO());
      setBozza({
        file: file.name,
        etichetta,
        dal: periodo?.dal ?? "",
        al: periodo?.al ?? "",
        appalti: letti,
        saltati,
      });
    } catch (e) {
      toast.error(t("ca.errLettura"), { description: e instanceof Error ? e.message : String(e) });
    } finally {
      setBusy(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };
  const dateOk = Boolean(
    bozza &&
    /^\d{4}-\d{2}-\d{2}$/.test(bozza.dal) &&
    /^\d{4}-\d{2}-\d{2}$/.test(bozza.al) &&
    bozza.al >= bozza.dal,
  );
  const giaCaricata = Boolean(bozza && dateOk && settimane.some((s) => s.id === bozza.dal));
  const salvaBozza = async () => {
    if (!bozza || !dateOk) return;
    setBusy(true);
    try {
      const salvato = await spCostiAppaltiSalva({
        data: {
          settimana: {
            id: bozza.dal,
            wk: settimanaIso(bozza.dal).wk,
            dal: bozza.dal,
            al: bozza.al,
            etichetta: bozza.etichetta,
            file: bozza.file,
            caricataIl: "",
            caricataDa: "",
            appalti: bozza.appalti,
          },
        },
      });
      setDb(salvato);
      setSettimanaId(bozza.dal);
      setBozza(null);
      toast.success(t("ca.salvata"));
    } catch (e) {
      toast.error(t("ca.errSalva"), { description: e instanceof Error ? e.message : String(e) });
    } finally {
      setBusy(false);
    }
  };
  const elimina = async () => {
    if (!settimana) return;
    if (!window.confirm(`${t("ca.eliminaConferma")} WK${settimana.wk} (${dataIt(settimana.dal)})?`))
      return;
    setBusy(true);
    try {
      setDb(await spCostiAppaltiElimina({ data: { id: settimana.id } }));
      setSettimanaId("");
    } catch (e) {
      toast.error(t("ca.errSalva"), { description: e instanceof Error ? e.message : String(e) });
    } finally {
      setBusy(false);
    }
  };

  // --- il conto voce per voce ---------------------------------------------------
  const righeLato = (lato: "ricavi" | "costi") =>
    vociDi(visibili, lato)
      .map((voce) => {
        const celle = visibili.map((a) => voceDi(a, lato, voce));
        const totale = celle.reduce((s, c) => s + (c?.importo ?? 0), 0);
        return { voce, celle, totale };
      })
      .filter((r) => ancheZero || r.celle.some((c) => Math.abs(c?.importo ?? 0) > 0.005));
  const righeRicavi = righeLato("ricavi");
  const righeCosti = righeLato("costi");
  const totPer = useMemo(() => visibili.map((a) => totaliAppalto(a)), [visibili]);

  /** Una cella della tabella: importo, oppure la quota sui ricavi di quel cliente. */
  const valore = (importo: number, ricavi: number) =>
    vista === "percento" ? pct(ricavi > 0.005 ? importo / ricavi : null) : fmt(importo);

  // --- andamento ----------------------------------------------------------------
  const andamento = useMemo(() => {
    const chiavi = new Set(visibili.map((a) => chiaveAppalto(a.nome)));
    return settimane
      .map((s) => {
        const l = tutti ? s.appalti : s.appalti.filter((a) => chiavi.has(chiaveAppalto(a.nome)));
        return { s, n: l.length, t: totaliInsieme(l) };
      })
      .filter((x) => x.n > 0);
  }, [settimane, visibili, tutti]);
  const maxRicavi = Math.max(1, ...andamento.map((x) => x.t.ricavi));

  const etichettaScelta = uno
    ? uno.nome
    : tutti
      ? t("ca.totaleDr")
      : `${visibili.length} ${t("ca.clientiScelti")}`;

  const csvRighe = () => {
    const out: (string | number)[][] = [];
    const n = (x: number) => x.toFixed(2).replace(".", ",");
    const riga = (
      sezione: string,
      voce: string,
      tot: number,
      per: (number | undefined)[],
      note = "",
    ) => out.push([sezione, voce, n(tot), ...per.map((x) => (x == null ? "" : n(x))), note]);
    for (const r of righeRicavi)
      riga(
        t("ca.ricavi"),
        r.voce,
        r.totale,
        r.celle.map((c) => c?.importo),
        r.celle
          .map((c, i) => (c?.nota ? `${visibili[i].nome}: ${c.nota}` : ""))
          .filter(Boolean)
          .join(" · "),
      );
    riga(
      t("ca.ricavi"),
      t("ca.totRicavi"),
      totali.ricavi,
      totPer.map((x) => x.ricavi),
    );
    for (const r of righeCosti)
      riga(
        t("ca.costi"),
        r.voce,
        r.totale,
        r.celle.map((c) => c?.importo),
        r.celle
          .map((c, i) => (c?.nota ? `${visibili[i].nome}: ${c.nota}` : ""))
          .filter(Boolean)
          .join(" · "),
      );
    riga(
      t("ca.costi"),
      t("ca.totCosti"),
      totali.costi,
      totPer.map((x) => x.costi),
    );
    riga(
      "",
      t("ca.mol"),
      totali.mol,
      totPer.map((x) => x.mol),
    );
    out.push(["", t("ca.margine"), pct(totali.margine), ...totPer.map((x) => pct(x.margine)), ""]);
    return out;
  };

  const scegliFile = (
    <>
      <input
        ref={fileRef}
        type="file"
        accept=".xlsx,.xlsm,.xls"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void leggiFile(f);
        }}
      />
      <button
        type="button"
        disabled={busy}
        onClick={() => fileRef.current?.click()}
        className="no-print inline-flex items-center gap-2 rounded-lg border border-primary bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground shadow-sm hover:opacity-90 disabled:opacity-50"
      >
        {busy ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
        {t("ca.caricaFile")}
      </button>
    </>
  );

  if (errore && !db)
    return (
      <div
        role="alert"
        className="rounded-xl border-2 border-rose-400 bg-rose-50 px-4 py-3 text-rose-900 shadow-sm dark:border-rose-700 dark:bg-rose-950/50 dark:text-rose-100"
      >
        <div className="text-sm font-bold uppercase tracking-wide">{t("fc.nonCaricatiTitolo")}</div>
        <p className="mt-1.5 text-[13px]">{t("ca.errCarica")}</p>
        <p className="mt-1 text-[11px] opacity-70">{errore}</p>
        <button
          type="button"
          onClick={carica}
          className="mt-2 rounded-lg border border-rose-500 bg-white px-3 py-1 text-[13px] font-semibold text-rose-800 hover:bg-rose-100 dark:bg-rose-900 dark:text-rose-50"
        >
          {t("fc.riprova")}
        </button>
      </div>
    );
  if (!db)
    return (
      <div className="flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );

  return (
    <div className="space-y-4">
      {/* Testata: cosa è, quale settimana, carica il file */}
      <div className={`${card} px-5 py-4`}>
        <div className="flex flex-wrap items-start gap-x-6 gap-y-3">
          <div className="min-w-64 flex-1">
            <div className="flex items-center gap-2 text-sm font-semibold text-foreground">
              <Calculator className="h-4 w-4 text-primary" /> {t("ca.titolo")}
            </div>
            <p className="mt-1 max-w-3xl text-xs text-muted-foreground">{t("ca.intro")}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {settimane.length > 0 && (
              <label className="no-print flex items-center gap-2 text-[13px] text-muted-foreground">
                {t("ca.settimana")}
                <select
                  value={settimana?.id ?? ""}
                  onChange={(e) => setSettimanaId(e.target.value)}
                  className={inputCls}
                >
                  {settimane.map((s) => (
                    <option key={s.id} value={s.id}>
                      WK{s.wk} · {ddmm(s.dal)} – {dataIt(s.al)}
                    </option>
                  ))}
                </select>
              </label>
            )}
            {scegliFile}
            {settimana && (
              <>
                <EsportaStampa
                  nomeFile={`costi-appalti-wk${settimana.wk}-${settimana.dal}`}
                  testata={[
                    t("ca.colSezione"),
                    t("ca.colVoce"),
                    etichettaScelta,
                    ...visibili.map((a) => a.nome),
                    t("ca.colNote"),
                  ]}
                  righe={csvRighe}
                />
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => void elimina()}
                  title={t("ca.eliminaTip")}
                  className="no-print rounded-lg border border-border p-2 text-muted-foreground hover:bg-muted hover:text-rose-600 disabled:opacity-50"
                >
                  <Trash2 className="h-4 w-4" />
                </button>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Il file appena letto: si guarda e si conferma */}
      {bozza && (
        <div className="overflow-hidden rounded-2xl border-2 border-primary/50 bg-card shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-center justify-between gap-2 bg-primary px-4 py-2 text-primary-foreground">
            <span className="text-[12px] font-semibold uppercase tracking-wide">
              {t("ca.bozzaTitolo")} · {bozza.file}
            </span>
            <span className="text-[13px] font-bold tabular-nums">
              {bozza.appalti.length} {t("ca.appalti")} · MOL {fmt(totaliInsieme(bozza.appalti).mol)}{" "}
              €
            </span>
          </div>
          <div className="space-y-3 px-4 py-3">
            <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-[13px]">
              <span className="text-muted-foreground">
                {t("ca.bozzaIntestazione")}{" "}
                <b className="text-foreground">{bozza.etichetta || "—"}</b>
              </span>
              <label className="flex items-center gap-2 text-muted-foreground">
                {t("ca.dal")}
                <input
                  type="date"
                  value={bozza.dal}
                  onChange={(e) => setBozza({ ...bozza, dal: e.target.value })}
                  className={inputCls}
                />
              </label>
              <label className="flex items-center gap-2 text-muted-foreground">
                {t("ca.al")}
                <input
                  type="date"
                  value={bozza.al}
                  onChange={(e) => setBozza({ ...bozza, al: e.target.value })}
                  className={inputCls}
                />
              </label>
              {dateOk && (
                <span className="rounded-full bg-primary/10 px-2 py-0.5 text-[12px] font-bold text-primary">
                  WK{settimanaIso(bozza.dal).wk}
                </span>
              )}
            </div>
            {!dateOk && (
              <p className="flex items-center gap-2 text-[12.5px] font-medium text-amber-700 dark:text-amber-400">
                <AlertTriangle className="h-4 w-4" /> {t("ca.bozzaDate")}
              </p>
            )}
            {giaCaricata && (
              <p className="flex items-center gap-2 text-[12.5px] font-medium text-amber-700 dark:text-amber-400">
                <AlertTriangle className="h-4 w-4" /> {t("ca.bozzaSostituisce")}
              </p>
            )}
            <div className="overflow-x-auto">
              <table className="w-full text-[13px]">
                <thead>
                  <tr className="border-b border-border text-left text-[11px] uppercase text-muted-foreground">
                    <th className="py-1.5 pr-3">{t("ca.colAppalto")}</th>
                    <th className="px-2 py-1.5 text-right">{t("ca.ricavi")}</th>
                    <th className="px-2 py-1.5 text-right">{t("ca.costi")}</th>
                    <th className="px-2 py-1.5 text-right">{t("ca.mol")}</th>
                    <th className="px-2 py-1.5 text-right">{t("ca.margine")}</th>
                    <th className="px-2 py-1.5">{t("ca.colControllo")}</th>
                  </tr>
                </thead>
                <tbody>
                  {bozza.appalti.map((a) => {
                    const x = totaliAppalto(a);
                    const diversi = controlliAppalto(a);
                    return (
                      <tr key={a.nome} className="border-b border-border/40">
                        <td className="py-1.5 pr-3 font-medium text-foreground">{a.nome}</td>
                        <td className="px-2 py-1.5 text-right tabular-nums">{fmt(x.ricavi)}</td>
                        <td className="px-2 py-1.5 text-right tabular-nums">{fmt(x.costi)}</td>
                        <td
                          className={`px-2 py-1.5 text-right font-semibold tabular-nums ${tintaMargine(x.margine)}`}
                        >
                          {fmt(x.mol)}
                        </td>
                        <td
                          className={`px-2 py-1.5 text-right tabular-nums ${tintaMargine(x.margine)}`}
                        >
                          {pct(x.margine)}
                        </td>
                        <td className="px-2 py-1.5 text-[12px]">
                          {diversi.length === 0 ? (
                            <span className="text-emerald-700 dark:text-emerald-400">
                              ✓ {t("ca.controlloOk")}
                            </span>
                          ) : (
                            <span className="text-amber-700 dark:text-amber-400">
                              {t("ca.controlloDiverso")} {diversi.join(", ")}
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {bozza.saltati.length > 0 && (
              <p className="text-[11.5px] text-muted-foreground">
                {t("ca.bozzaSaltati")} {bozza.saltati.join(", ")}.
              </p>
            )}
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                disabled={busy || !dateOk}
                onClick={() => void salvaBozza()}
                className="inline-flex items-center gap-2 rounded-lg border border-primary bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground shadow-sm hover:opacity-90 disabled:opacity-50"
              >
                {busy && <Loader2 className="h-4 w-4 animate-spin" />}
                {t(giaCaricata ? "ca.bozzaSostituisciBtn" : "ca.bozzaSalva")}
              </button>
              <button
                type="button"
                disabled={busy}
                onClick={() => setBozza(null)}
                className="rounded-lg border border-border px-3 py-1.5 text-sm text-muted-foreground hover:bg-muted"
              >
                {t("common.cancel")}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Nessuna settimana caricata */}
      {!settimana && !bozza && (
        <div className={`${card} px-5 py-8 text-center`}>
          <Calculator className="mx-auto h-8 w-8 text-muted-foreground" />
          <p className="mt-2 text-sm font-semibold text-foreground">{t("ca.vuotoTitolo")}</p>
          <p className="mx-auto mt-1 max-w-xl text-[13px] text-muted-foreground">{t("ca.vuoto")}</p>
        </div>
      )}

      {settimana && (
        <>
          {/* I quattro numeri della settimana, per i clienti accesi */}
          <div className={`${card} px-5 py-4`}>
            <div className="flex flex-wrap items-end gap-x-8 gap-y-4">
              <div>
                <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  WK{settimana.wk} · {ddmm(settimana.dal)} – {dataIt(settimana.al)}
                </div>
                <div className="mt-0.5 text-xl font-bold text-foreground">{etichettaScelta}</div>
                <div className="mt-1 text-[11px] text-muted-foreground">
                  {settimana.file}
                  {settimana.caricataDa ? ` · ${settimana.caricataDa}` : ""}
                  {settimana.caricataIl ? ` · ${dataIt(settimana.caricataIl.slice(0, 10))}` : ""}
                </div>
              </div>
              <Numero
                titolo={t("ca.ricavi")}
                valore={`${fmt(totali.ricavi)} €`}
                tinta="text-emerald-700 dark:text-emerald-400"
              />
              <Numero
                titolo={t("ca.costi")}
                valore={`${fmt(totali.costi)} €`}
                tinta="text-rose-700 dark:text-rose-400"
              />
              <Numero
                titolo={t("ca.mol")}
                valore={`${fmt(totali.mol)} €`}
                tinta={tintaMargine(totali.margine)}
                grande
              />
              <Numero
                titolo={t("ca.margine")}
                valore={pct(totali.margine)}
                tinta={tintaMargine(totali.margine)}
                grande
              />
            </div>
          </div>

          {/* I clienti: si accendono e si spengono */}
          <div className={`${card} no-print px-4 py-3`}>
            <div className="flex flex-wrap items-center gap-2">
              <span className="mr-1 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                {t("ca.clienti")}
              </span>
              <button
                type="button"
                onClick={() => setScelti(new Set())}
                className={`rounded-full border px-3 py-1 text-[12.5px] font-medium transition-colors ${tutti ? "border-primary bg-primary text-primary-foreground" : "border-border text-muted-foreground hover:bg-muted"}`}
              >
                {t("ca.tutti")}
              </button>
              {appalti.map((a) => {
                const x = totaliAppalto(a);
                const acceso = !tutti && visibili.includes(a);
                return (
                  <button
                    key={a.nome}
                    type="button"
                    onClick={() => toggle(a.nome)}
                    title={`${t("ca.ricavi")} ${fmt(x.ricavi)} · ${t("ca.costi")} ${fmt(x.costi)} · ${t("ca.mol")} ${fmt(x.mol)}`}
                    className={`inline-flex items-center gap-2 rounded-full border px-3 py-1 text-[12.5px] font-medium transition-colors ${acceso ? "border-primary bg-primary/10 text-foreground ring-1 ring-primary" : "border-border text-foreground hover:bg-muted"}`}
                  >
                    <span className={`h-2 w-2 rounded-full ${puntoMargine(x.margine)}`} />
                    {a.nome}
                    <span className={`tabular-nums ${tintaMargine(x.margine)}`}>
                      {pct(x.margine)}
                    </span>
                  </button>
                );
              })}
              <span className="ml-auto flex flex-wrap items-center gap-3 text-[12px] text-muted-foreground">
                <span className="inline-flex overflow-hidden rounded-lg border border-border">
                  {(["importi", "percento"] as const).map((v) => (
                    <button
                      key={v}
                      type="button"
                      onClick={() => setVista(v)}
                      className={`px-2.5 py-1 ${vista === v ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
                    >
                      {t(v === "importi" ? "ca.vistaImporti" : "ca.vistaPercento")}
                    </button>
                  ))}
                </span>
                <label className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={ancheZero}
                    onChange={(e) => setAncheZero(e.target.checked)}
                  />
                  {t("ca.ancheZero")}
                </label>
              </span>
            </div>
          </div>

          {/* Il conto voce per voce */}
          <div className={`${card} overflow-hidden`}>
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-[13px]">
                <thead>
                  <tr className="border-b border-border bg-muted/50 text-[11px] uppercase tracking-wide text-muted-foreground">
                    <th className="px-4 py-2 text-left">{t("ca.colVoce")}</th>
                    {!uno && (
                      <th className="border-l border-border px-3 py-2 text-right text-foreground">
                        {etichettaScelta}
                      </th>
                    )}
                    {visibili.map((a) => (
                      <th key={a.nome} className="border-l border-border px-3 py-2 text-right">
                        {a.nome}
                      </th>
                    ))}
                    {uno && (
                      <>
                        <th className="border-l border-border px-3 py-2 text-right">
                          {t("ca.colQuota")}
                        </th>
                        <th className="border-l border-border px-3 py-2 text-left">
                          {t("ca.colNote")}
                        </th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody>
                  <Fascia
                    nome={t("ca.ricavi")}
                    cls="bg-emerald-100/80 text-emerald-900 dark:bg-emerald-900/40 dark:text-emerald-200"
                    colonne={visibili.length + (uno ? 3 : 2)}
                  />
                  {righeRicavi.map((r) => (
                    <RigaVoce
                      key={`r|${r.voce}`}
                      voce={r.voce}
                      totale={uno ? undefined : valore(r.totale, totali.ricavi)}
                      celle={r.celle.map((c, i) => ({
                        testo: c ? valore(c.importo, totPer[i].ricavi) : "",
                        nota: c?.nota,
                        zero: Math.abs(c?.importo ?? 0) < 0.005,
                      }))}
                      uno={
                        uno
                          ? { quota: pct(totali.ricavi > 0.005 ? r.totale / totali.ricavi : null) }
                          : undefined
                      }
                    />
                  ))}
                  <RigaTotale
                    nome={t("ca.totRicavi")}
                    cls="text-emerald-800 dark:text-emerald-300"
                    totale={uno ? undefined : fmt(totali.ricavi)}
                    celle={totPer.map((x) => fmt(x.ricavi))}
                    uno={uno ? "100,0%" : undefined}
                  />
                  <Fascia
                    nome={t("ca.costi")}
                    cls="bg-rose-100/90 text-rose-900 dark:bg-rose-900/40 dark:text-rose-200"
                    colonne={visibili.length + (uno ? 3 : 2)}
                  />
                  {righeCosti.map((r) => (
                    <RigaVoce
                      key={`c|${r.voce}`}
                      voce={r.voce}
                      totale={uno ? undefined : valore(r.totale, totali.ricavi)}
                      celle={r.celle.map((c, i) => ({
                        testo: c ? valore(c.importo, totPer[i].ricavi) : "",
                        nota: c?.nota,
                        zero: Math.abs(c?.importo ?? 0) < 0.005,
                      }))}
                      uno={
                        uno
                          ? { quota: pct(totali.ricavi > 0.005 ? r.totale / totali.ricavi : null) }
                          : undefined
                      }
                    />
                  ))}
                  <RigaTotale
                    nome={t("ca.totCosti")}
                    cls="text-rose-800 dark:text-rose-300"
                    totale={uno ? undefined : valore(totali.costi, totali.ricavi)}
                    celle={totPer.map((x) => valore(x.costi, x.ricavi))}
                    uno={
                      uno
                        ? pct(totali.ricavi > 0.005 ? totali.costi / totali.ricavi : null)
                        : undefined
                    }
                  />
                  {/* MOL 1 e margine: il risultato, in evidenza */}
                  <tr className="border-t-2 border-foreground/70 bg-muted/40">
                    <td className="px-4 py-2 text-[12px] font-bold uppercase tracking-wide text-foreground">
                      {t("ca.mol")}
                    </td>
                    {!uno && (
                      <td
                        className={`border-l border-border px-3 py-2 text-right text-[15px] font-bold tabular-nums ${tintaMargine(totali.margine)}`}
                      >
                        {fmt(totali.mol)}
                      </td>
                    )}
                    {totPer.map((x, i) => (
                      <td
                        key={visibili[i].nome}
                        className={`border-l border-border px-3 py-2 text-right text-[15px] font-bold tabular-nums ${tintaMargine(x.margine)}`}
                      >
                        {fmt(x.mol)}
                      </td>
                    ))}
                    {uno && (
                      <>
                        <td
                          className={`border-l border-border px-3 py-2 text-right text-[15px] font-bold tabular-nums ${tintaMargine(totali.margine)}`}
                        >
                          {pct(totali.margine)}
                        </td>
                        <td className="border-l border-border px-3 py-2 text-[12px] text-muted-foreground">
                          {uno.giorni ? `${uno.giorni} ${t("ca.giorni")}` : ""}
                        </td>
                      </>
                    )}
                  </tr>
                  {!uno && (
                    <tr className="bg-muted/40">
                      <td className="px-4 py-1.5 text-[12px] font-bold uppercase tracking-wide text-foreground">
                        {t("ca.margine")}
                      </td>
                      <td
                        className={`border-l border-border px-3 py-1.5 text-right font-bold tabular-nums ${tintaMargine(totali.margine)}`}
                      >
                        {pct(totali.margine)}
                      </td>
                      {totPer.map((x, i) => (
                        <td
                          key={visibili[i].nome}
                          className={`border-l border-border px-3 py-1.5 text-right font-bold tabular-nums ${tintaMargine(x.margine)}`}
                        >
                          {pct(x.margine)}
                        </td>
                      ))}
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <p className="border-t border-border px-4 py-1.5 text-[11px] text-muted-foreground">
              {t("ca.piede")}
            </p>
          </div>

          {/* Andamento: una riga per settimana caricata */}
          {andamento.length > 1 && (
            <div className={`${card} overflow-hidden`}>
              <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-700 px-4 py-2 text-white">
                <span className="text-[12px] font-semibold uppercase tracking-wide">
                  {t("ca.andamento")} · {etichettaScelta}
                </span>
                <span className="text-[12px] opacity-80">
                  {andamento.length} {t("ca.settimane")}
                </span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-[13px]">
                  <thead>
                    <tr className="border-b border-border text-left text-[11px] uppercase text-muted-foreground">
                      <th className="px-4 py-1.5">{t("ca.settimana")}</th>
                      <th className="px-2 py-1.5 text-right">{t("ca.ricavi")}</th>
                      <th className="px-2 py-1.5 text-right">{t("ca.costi")}</th>
                      <th className="px-2 py-1.5 text-right">{t("ca.mol")}</th>
                      <th className="px-2 py-1.5 text-right">{t("ca.margine")}</th>
                      <th className="w-1/3 px-4 py-1.5">{t("ca.colBarra")}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {andamento.map(({ s, t: x }) => (
                      <tr
                        key={s.id}
                        className={`border-b border-border/40 ${s.id === settimana.id ? "bg-primary/5" : ""}`}
                      >
                        <td className="whitespace-nowrap px-4 py-1.5">
                          <button
                            type="button"
                            onClick={() => setSettimanaId(s.id)}
                            className="font-medium text-foreground underline decoration-dotted decoration-border underline-offset-4 hover:text-primary"
                          >
                            WK{s.wk}
                          </button>{" "}
                          <span className="text-muted-foreground">
                            {ddmm(s.dal)} – {dataIt(s.al)}
                          </span>
                        </td>
                        <td className="px-2 py-1.5 text-right tabular-nums">{fmt(x.ricavi)}</td>
                        <td className="px-2 py-1.5 text-right tabular-nums">{fmt(x.costi)}</td>
                        <td
                          className={`px-2 py-1.5 text-right font-semibold tabular-nums ${tintaMargine(x.margine)}`}
                        >
                          {fmt(x.mol)}
                        </td>
                        <td
                          className={`px-2 py-1.5 text-right tabular-nums ${tintaMargine(x.margine)}`}
                        >
                          {pct(x.margine)}
                        </td>
                        <td className="px-4 py-1.5">
                          <Barra
                            x={x}
                            max={maxRicavi}
                            etichetta={`${fmt0(x.costi)} / ${fmt0(x.ricavi)}`}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/** Un numero grande della banda in alto. */
function Numero({
  titolo,
  valore,
  tinta,
  grande,
}: {
  titolo: string;
  valore: string;
  tinta: string;
  grande?: boolean;
}) {
  return (
    <div className="border-l-2 border-border pl-4">
      <div className="text-[11px] uppercase tracking-wider text-muted-foreground">{titolo}</div>
      <div
        className={`mt-0.5 whitespace-nowrap font-semibold tabular-nums ${grande ? "text-2xl" : "text-xl"} ${tinta}`}
      >
        {valore}
      </div>
    </div>
  );
}

/** La fascia colorata che apre un blocco della tabella (ricavi, costi). */
function Fascia({ nome, cls, colonne }: { nome: string; cls: string; colonne: number }) {
  return (
    <tr>
      <td
        colSpan={colonne}
        className={`px-4 py-1 text-[11px] font-bold uppercase tracking-wider ${cls}`}
      >
        {nome}
      </td>
    </tr>
  );
}

function RigaVoce({
  voce,
  totale,
  celle,
  uno,
}: {
  voce: string;
  /** Colonna del totale (assente quando c'è un cliente solo). */
  totale?: string;
  celle: { testo: string; nota?: string; zero: boolean }[];
  /** Con un cliente solo: la quota sui ricavi e la nota per esteso. */
  uno?: { quota: string };
}) {
  return (
    <tr className="border-b border-border/40 hover:bg-primary/5">
      <td className="px-4 py-1.5 text-foreground">{voce}</td>
      {totale !== undefined && (
        <td className="border-l border-border px-3 py-1.5 text-right font-medium tabular-nums text-foreground">
          {totale}
        </td>
      )}
      {celle.map((c, i) => (
        <td
          key={i}
          title={uno ? undefined : c.nota}
          className={`border-l border-border px-3 py-1.5 text-right tabular-nums ${c.zero ? "text-muted-foreground/50" : "text-foreground"}`}
        >
          {c.zero ? "—" : c.testo}
          {!uno && c.nota && (
            <span className="ml-1 cursor-help align-super text-[9px] font-bold text-primary">
              ●
            </span>
          )}
        </td>
      ))}
      {uno && (
        <>
          <td className="border-l border-border px-3 py-1.5 text-right tabular-nums text-muted-foreground">
            {celle[0]?.zero ? "" : uno.quota}
          </td>
          <td className="border-l border-border px-3 py-1.5 text-[12px] text-muted-foreground">
            {celle[0]?.nota ?? ""}
          </td>
        </>
      )}
    </tr>
  );
}

function RigaTotale({
  nome,
  cls,
  totale,
  celle,
  uno,
}: {
  nome: string;
  cls: string;
  totale?: string;
  celle: string[];
  uno?: string;
}) {
  return (
    <tr className="border-b border-border bg-muted/30">
      <td className={`px-4 py-1.5 text-[12px] font-bold uppercase tracking-wide ${cls}`}>{nome}</td>
      {totale !== undefined && (
        <td
          className={`border-l border-border px-3 py-1.5 text-right font-bold tabular-nums ${cls}`}
        >
          {totale}
        </td>
      )}
      {celle.map((c, i) => (
        <td
          key={i}
          className={`border-l border-border px-3 py-1.5 text-right font-bold tabular-nums ${cls}`}
        >
          {c}
        </td>
      ))}
      {uno !== undefined && (
        <>
          <td
            className={`border-l border-border px-3 py-1.5 text-right font-bold tabular-nums ${cls}`}
          >
            {uno}
          </td>
          <td className="border-l border-border" />
        </>
      )}
    </tr>
  );
}

/** Costi dentro i ricavi: la parte piena sono i costi, il resto è il MOL. Se
 *  i costi superano i ricavi la barra diventa rossa. */
function Barra({ x, max, etichetta }: { x: TotaliAppalto; max: number; etichetta: string }) {
  const larghezza = Math.max(2, Math.round((Math.max(x.ricavi, x.costi) / max) * 100));
  const quotaCosti = x.ricavi > 0.005 ? Math.min(100, Math.round((x.costi / x.ricavi) * 100)) : 100;
  const perdita = x.mol < 0;
  return (
    <div className="flex items-center gap-2" title={etichetta}>
      <div className="h-3 flex-1 overflow-hidden rounded-full bg-muted">
        <div
          className={`flex h-full overflow-hidden rounded-full ${perdita ? "bg-rose-200 dark:bg-rose-900/60" : "bg-emerald-300 dark:bg-emerald-700/70"}`}
          style={{ width: `${larghezza}%` }}
        >
          <div
            className={perdita ? "h-full bg-rose-500" : "h-full bg-slate-400 dark:bg-slate-500"}
            style={{ width: `${quotaCosti}%` }}
          />
        </div>
      </div>
    </div>
  );
}
