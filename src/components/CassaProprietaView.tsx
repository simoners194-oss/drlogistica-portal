// DR Portal — Finanza → Flussi → vista "Cassa proprietà".
// La fotografia della cassa a una data, sul modello del commercialista:
// saldo in banca, entrate attese (certe / standby), uscite in tre secchi
// (ora / prossimi giorni / standby) con nota e stato per riga, righe manuali
// della proprietà e spese passate dalla carta Amex, totali come nel suo
// foglio, fotografie salvate. Logica in lib/cassa-logic; dati in cassa.json.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Camera, ChevronDown, ChevronRight, Loader2, Plus, Trash2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
type TFn = ReturnType<typeof useLang>["t"];
import { EsportaStampa } from "@/components/EsportaStampa";
import type { FatturaConStato } from "@/lib/fatture-logic";
import type { CostoFisso, FornitoreInfo } from "@/lib/flussi-logic";
import type { ScadenzaFiscale } from "@/lib/fiscale-logic";
import type { SpMovimento } from "@/lib/sharepoint.server";
import { spEbSaldo } from "@/lib/sharepoint.functions";
import { spStipendiGet } from "@/lib/stipendi.functions";
import { spCassaGet, spCassaSalva } from "@/lib/cassa.functions";
import type { StipendiDb } from "@/lib/stipendi-logic";
import {
  costruisciCassa,
  creaFotografia,
  emptyCassaDb,
  limiteDefault,
  stipendiDaPagare,
  type CassaDb,
  type FotografiaCassa,
  type RigaCassa,
  type RigaManualeCassa,
  type SecchioCassa,
  type StatoRigaCassa,
} from "@/lib/cassa-logic";

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtData = (iso: string) => (iso ? iso.slice(0, 10).split("-").reverse().join("/") : "—");
const oggiISO = () => new Date().toISOString().slice(0, 10);
const inputCls =
  "rounded-lg border border-border bg-background px-2 py-1 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";
const th = "px-2 py-1.5 text-left text-[11px] uppercase tracking-wider text-muted-foreground";
const thR = `${th} text-right`;
const td = "px-2 py-1 text-sm text-foreground align-top";
const tdR = `${td} text-right tabular-nums whitespace-nowrap`;

export function CassaProprietaView({
  attive,
  passive,
  esclusioni,
  infoFornitori,
  fiscale,
  fissi,
  movimenti,
}: {
  attive: readonly FatturaConStato[];
  passive: readonly FatturaConStato[];
  esclusioni: readonly { nome: string; mese?: string; meseFine?: string }[];
  infoFornitori: ReadonlyMap<string, FornitoreInfo>;
  fiscale: readonly ScadenzaFiscale[];
  fissi: readonly CostoFisso[];
  movimenti: readonly SpMovimento[];
}) {
  const { t } = useLang();
  const [db, setDb] = useState<CassaDb | null>(null);
  const [stipendiDb, setStipendiDb] = useState<StipendiDb | null>(null);
  const [saldoBanca, setSaldoBanca] = useState<{ saldo: number; al?: string } | null>(null);
  const [saldoErr, setSaldoErr] = useState(false);
  const [oggi, setOggi] = useState(oggiISO());
  const [limite, setLimite] = useState("");
  const [saldoOverride, setSaldoOverride] = useState("");
  const [salvando, setSalvando] = useState(false);
  const [filtro, setFiltro] = useState<"tutte" | SecchioCassa>("tutte");
  const [mostraOltre, setMostraOltre] = useState(false);
  const [fotoAperta, setFotoAperta] = useState<FotografiaCassa | null>(null);
  const [nuova, setNuova] = useState<Partial<RigaManualeCassa>>({
    tipo: "manuale",
    secchio: "ora",
    data: oggiISO(),
  });

  useEffect(() => {
    spCassaGet()
      .then((d) => {
        setDb(d);
        setLimite(limiteDefault(oggiISO(), d.giornoLimite));
      })
      .catch((e) => {
        setDb(emptyCassaDb());
        setLimite(limiteDefault(oggiISO(), 5));
        toast.error(t("common.error"), { description: e instanceof Error ? e.message : String(e) });
      });
    spStipendiGet()
      .then((d) => setStipendiDb(d))
      .catch(() => setStipendiDb(null));
    spEbSaldo()
      .then((s) => {
        if (s) setSaldoBanca({ saldo: s.saldo, al: s.aggiornatoAl });
        else setSaldoErr(true);
      })
      .catch(() => setSaldoErr(true));
  }, [t]);

  const saldo = useMemo(() => {
    if (saldoOverride.trim()) {
      const n = Number(saldoOverride.replace(/\./g, "").replace(",", "."));
      if (Number.isFinite(n)) return n;
    }
    if (saldoBanca) return saldoBanca.saldo;
    return db?.saldoManuale?.importo ?? 0;
  }, [saldoOverride, saldoBanca, db]);

  const stipendi = useMemo(
    () => (stipendiDb ? stipendiDaPagare(stipendiDb, (limite || oggi).slice(0, 7)) : []),
    [stipendiDb, limite, oggi],
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
      movimenti: movimenti.map((m) => ({
        cliente: m.cliente,
        descrizione: m.descrizione,
        importo: m.importo,
        dataContabile: m.dataContabile,
      })),
      stipendi,
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
    movimenti,
    stipendi,
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
  const aggiungiManuale = () => {
    if (!db) return;
    const importo = Number(
      String(nuova.importo ?? "")
        .replace(/\./g, "")
        .replace(",", "."),
    );
    if (!nuova.descrizione?.trim() || !Number.isFinite(importo) || importo === 0) {
      toast.error(t("cpr.manualeIncompleta"));
      return;
    }
    const riga: RigaManualeCassa = {
      id: `M${Date.now().toString(36)}`,
      tipo: nuova.tipo === "amex" ? "amex" : "manuale",
      descrizione: nuova.descrizione.trim(),
      importo: Math.abs(importo),
      data: nuova.data || oggi,
      secchio: nuova.secchio ?? "ora",
      nota: nuova.nota?.trim() || undefined,
      entrata: nuova.entrata ? true : undefined,
    };
    void salva({ ...db, manuali: [...db.manuali, riga] });
    setNuova({ tipo: nuova.tipo, secchio: "ora", data: oggi });
  };
  const eliminaManuale = (id: string) => {
    if (!db) return;
    const stati = { ...db.stati };
    delete stati[`man|${id}`];
    void salva({ ...db, manuali: db.manuali.filter((m) => m.id !== id), stati });
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
    if (fotoAperta?.id === id) setFotoAperta(null);
  };
  const cambiaGiornoLimite = (giorno: number) => {
    if (!db) return;
    setLimite(limiteDefault(oggi, giorno));
    void salva({ ...db, giornoLimite: giorno });
  };

  const righeCsv = () => {
    if (!esito) return [];
    const out: (string | number)[][] = [];
    for (const r of esito.righe)
      out.push([
        r.entrata ? t("cpr.entrata") : t("cpr.uscita"),
        secchioLabel(r.secchio, t),
        r.controparte,
        r.descrizione,
        fmtData(r.scadenza),
        r.importo,
        r.certa ? t("cpr.certa") : "",
        r.standbyIncasso ? t("cpr.standbyIncasso") : "",
        r.contestata ? t("cpr.contestata") : "",
        r.nota,
      ]);
    return out;
  };

  if (!db || !esito) {
    return (
      <div className="flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );
  }
  const tot = esito.totali;
  const entrate = esito.righe.filter((r) => r.entrata);
  const uscite = esito.righe.filter(
    (r) => !r.entrata && (filtro === "tutte" || r.secchio === filtro),
  );

  return (
    <div className="space-y-4">
      {/* Barra: data, limite, saldo, fotografia */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-3 text-[13px]">
          <label className="flex flex-col gap-1 text-xs text-muted-foreground">
            {t("cpr.situazioneAl")}
            <input
              type="date"
              value={oggi}
              onChange={(e) => setOggi(e.target.value || oggiISO())}
              className={inputCls}
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-muted-foreground">
            {t("cpr.prossimiFinoAl")}
            <input
              type="date"
              value={limite}
              onChange={(e) => setLimite(e.target.value)}
              className={inputCls}
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-muted-foreground">
            {t("cpr.giornoDefault")}
            <input
              type="number"
              min={1}
              max={28}
              value={db.giornoLimite}
              onChange={(e) =>
                cambiaGiornoLimite(Math.min(28, Math.max(1, Number(e.target.value) || 5)))
              }
              className={`${inputCls} w-20`}
            />
          </label>
          <label className="flex flex-col gap-1 text-xs text-muted-foreground">
            {t("cpr.saldoBanca")}
            <input
              value={saldoOverride}
              onChange={(e) => setSaldoOverride(e.target.value)}
              placeholder={
                saldoBanca
                  ? fmt(saldoBanca.saldo)
                  : db.saldoManuale
                    ? fmt(db.saldoManuale.importo)
                    : "—"
              }
              className={`${inputCls} w-40 text-right`}
            />
          </label>
          <span className="pb-1 text-[11px] text-muted-foreground">
            {saldoBanca
              ? `${t("cpr.saldoDaBanca")} ${saldoBanca.al ? fmtData(saldoBanca.al) : ""}`
              : saldoErr
                ? t("cpr.saldoNonDisponibile")
                : ""}
          </span>
          <div className="ml-auto flex items-center gap-2">
            {salvando && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
            <button
              type="button"
              onClick={salvaFotografia}
              className="inline-flex items-center gap-1 rounded-lg border border-primary bg-primary px-3 py-1 text-primary-foreground"
            >
              <Camera className="h-4 w-4" /> {t("cpr.salvaFoto")}
            </button>
            <EsportaStampa
              nomeFile={`cassa-proprieta-${oggi}`}
              testata={[
                t("cpr.tipo"),
                t("cpr.secchio"),
                t("cpr.controparte"),
                t("cpr.descrizione"),
                t("cpr.scadenza"),
                t("cpr.importo"),
                t("cpr.certa"),
                t("cpr.standbyIncasso"),
                t("cpr.contestata"),
                t("cpr.nota"),
              ]}
              righe={righeCsv}
            />
          </div>
        </div>
        <p className="mt-2 text-[12px] text-muted-foreground">{t("cpr.intro")}</p>
      </div>

      {/* Totali del commercialista */}
      <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-5">
        <Kpi label={t("cpr.saldoBanca")} value={tot.saldo} />
        <Kpi label={t("cpr.usciteOra")} value={tot.usciteOra} neg />
        <Kpi label={t("cpr.saldoDopoOra")} value={tot.saldoDopoOra} bold />
        <Kpi label={t("cpr.usciteProssimi")} value={tot.usciteProssimi} neg />
        <Kpi label={t("cpr.saldoDopoProssimi")} value={tot.saldoDopoProssimi} bold />
        <Kpi label={t("cpr.entrateCerte")} value={tot.entrateCerte} />
        <Kpi label={t("cpr.cassaFinale")} value={tot.cassaFinale} bold />
        <Kpi label={t("cpr.entrateStandby")} value={tot.entrateStandby} />
        <Kpi label={t("cpr.usciteStandby")} value={tot.usciteStandby} neg />
        <Kpi label={t("cpr.cassaConStandby")} value={tot.cassaConStandby} bold />
      </div>

      {/* Entrate */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-2 text-[15px] font-semibold text-foreground">
          {t("cpr.entrate")}{" "}
          <span className="text-sm font-normal text-muted-foreground">({entrate.length})</span>
        </div>
        <TabellaRighe
          righe={entrate}
          oggi={oggi}
          onSecchio={(r, b) => setStato(r.chiave, { b: b === r.secchioAuto ? undefined : b })}
          onStandbyIncasso={(r, v) => setStato(r.chiave, { s: v })}
          onNota={(r, n) => setStato(r.chiave, { n })}
          onContestata={(r, v) => setStato(r.chiave, { c: v })}
          onImporto={(r, i) => setStato(r.chiave, { i: i === r.importoOrigine ? undefined : i })}
          onTogli={(r) => setStato(r.chiave, { x: true })}
          onElimina={(r) => eliminaManuale(r.chiave.slice(4))}
          entrate
        />
      </div>

      {/* Uscite */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-2 flex flex-wrap items-center gap-2">
          <span className="text-[15px] font-semibold text-foreground">{t("cpr.uscite")}</span>
          <div className="ml-2 flex rounded-lg border border-border overflow-hidden text-[12px]">
            {(["tutte", "ora", "prossimi", "standby"] as const).map((f) => (
              <button
                key={f}
                type="button"
                onClick={() => setFiltro(f)}
                className={`px-2.5 py-1 ${filtro === f ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
              >
                {f === "tutte" ? t("common.all") : secchioLabel(f, t)}
              </button>
            ))}
          </div>
          <span className="text-[12px] text-muted-foreground">
            {t("cpr.ora")} {fmt(tot.usciteOra)} · {t("cpr.prossimi")} {fmt(tot.usciteProssimi)} ·{" "}
            {t("cpr.standby")} {fmt(tot.usciteStandby)}
          </span>
        </div>
        <TabellaRighe
          righe={uscite}
          oggi={oggi}
          onSecchio={(r, b) => setStato(r.chiave, { b: b === r.secchioAuto ? undefined : b })}
          onNota={(r, n) => setStato(r.chiave, { n })}
          onContestata={(r, v) => setStato(r.chiave, { c: v })}
          onImporto={(r, i) => setStato(r.chiave, { i: i === r.importoOrigine ? undefined : i })}
          onTogli={(r) => setStato(r.chiave, { x: true })}
          onElimina={(r) => eliminaManuale(r.chiave.slice(4))}
        />
        {esito.oltre.length > 0 && (
          <div className="mt-3 text-[12px] text-muted-foreground">
            <button
              type="button"
              onClick={() => setMostraOltre((v) => !v)}
              className="inline-flex items-center gap-1 underline-offset-2 hover:underline"
            >
              {mostraOltre ? (
                <ChevronDown className="h-3.5 w-3.5" />
              ) : (
                <ChevronRight className="h-3.5 w-3.5" />
              )}
              {t("cpr.oltreOrizzonte")} {esito.oltre.length} · {fmt(tot.usciteOltre)}
            </button>
            {mostraOltre && (
              <ul className="mt-2 space-y-0.5">
                {esito.oltre.map((r) => (
                  <li key={r.chiave} className="flex flex-wrap items-center gap-2">
                    <span>{fmtData(r.scadenza)}</span>
                    <span className="text-foreground">{r.controparte}</span>
                    <span>{r.descrizione}</span>
                    <span className="tabular-nums text-foreground">{fmt(r.importo)}</span>
                    <button
                      type="button"
                      className="rounded border border-border px-1.5 hover:bg-muted"
                      onClick={() => setStato(r.chiave, { b: "prossimi" })}
                    >
                      {t("cpr.portaNeiProssimi")}
                    </button>
                    <button
                      type="button"
                      className="rounded border border-border px-1.5 hover:bg-muted"
                      onClick={() => setStato(r.chiave, { b: "standby" })}
                    >
                      {t("cpr.parcheggia")}
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
      </div>

      {/* Righe manuali / Amex */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-2 text-[15px] font-semibold text-foreground">{t("cpr.manuali")}</div>
        <p className="mb-3 text-[12px] text-muted-foreground">{t("cpr.manualiHelp")}</p>
        <div className="flex flex-wrap items-end gap-2 text-[13px]">
          <select
            value={nuova.tipo ?? "manuale"}
            onChange={(e) =>
              setNuova((n) => ({ ...n, tipo: e.target.value as "manuale" | "amex" }))
            }
            className={inputCls}
          >
            <option value="manuale">{t("cpr.tipoManuale")}</option>
            <option value="amex">{t("cpr.tipoAmex")}</option>
          </select>
          <input
            placeholder={t("cpr.descrizione")}
            value={nuova.descrizione ?? ""}
            onChange={(e) => setNuova((n) => ({ ...n, descrizione: e.target.value }))}
            className={`${inputCls} w-64`}
          />
          <input
            placeholder={t("cpr.importo")}
            value={String(nuova.importo ?? "")}
            onChange={(e) =>
              setNuova((n) => ({ ...n, importo: e.target.value as unknown as number }))
            }
            className={`${inputCls} w-28 text-right`}
          />
          <input
            type="date"
            value={nuova.data ?? oggi}
            onChange={(e) => setNuova((n) => ({ ...n, data: e.target.value }))}
            className={inputCls}
          />
          <select
            value={nuova.secchio ?? "ora"}
            onChange={(e) => setNuova((n) => ({ ...n, secchio: e.target.value as SecchioCassa }))}
            className={inputCls}
          >
            {(["ora", "prossimi", "standby"] as const).map((s) => (
              <option key={s} value={s}>
                {secchioLabel(s, t)}
              </option>
            ))}
          </select>
          <label className="flex items-center gap-1 text-xs text-muted-foreground">
            <input
              type="checkbox"
              checked={Boolean(nuova.entrata)}
              onChange={(e) => setNuova((n) => ({ ...n, entrata: e.target.checked }))}
            />{" "}
            {t("cpr.entrata")}
          </label>
          <input
            placeholder={t("cpr.nota")}
            value={nuova.nota ?? ""}
            onChange={(e) => setNuova((n) => ({ ...n, nota: e.target.value }))}
            className={`${inputCls} w-56`}
          />
          <button
            type="button"
            onClick={aggiungiManuale}
            className="inline-flex items-center gap-1 rounded-lg border border-border px-3 py-1 hover:bg-muted"
          >
            <Plus className="h-4 w-4" /> {t("cpr.aggiungi")}
          </button>
        </div>
      </div>

      {/* Fotografie */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-2 text-[15px] font-semibold text-foreground">
          {t("cpr.fotografie")}{" "}
          <span className="text-sm font-normal text-muted-foreground">
            ({db.fotografie.length})
          </span>
        </div>
        {db.fotografie.length === 0 && (
          <p className="text-[12px] text-muted-foreground">{t("cpr.nessunaFoto")}</p>
        )}
        {db.fotografie.length > 0 && (
          <table className="w-full border-collapse">
            <thead>
              <tr className="border-b border-border">
                <th className={th}>{t("cpr.situazioneAl")}</th>
                <th className={th}>{t("cpr.prossimiFinoAl")}</th>
                <th className={thR}>{t("cpr.saldoBanca")}</th>
                <th className={thR}>{t("cpr.usciteOra")}</th>
                <th className={thR}>{t("cpr.saldoDopoProssimi")}</th>
                <th className={thR}>{t("cpr.cassaFinale")}</th>
                <th className={th}>{t("cpr.nota")}</th>
                <th className={th}></th>
              </tr>
            </thead>
            <tbody>
              {db.fotografie.map((f) => (
                <tr key={f.id} className="border-b border-border/60">
                  <td className={td}>
                    <button
                      type="button"
                      className="underline-offset-2 hover:underline"
                      onClick={() => setFotoAperta(fotoAperta?.id === f.id ? null : f)}
                    >
                      {fmtData(f.data)}
                    </button>
                    <span className="ml-1 text-[11px] text-muted-foreground">
                      {f.creataIl.slice(11, 16)}
                    </span>
                  </td>
                  <td className={td}>{fmtData(f.limite)}</td>
                  <td className={tdR}>{fmt(f.saldo)}</td>
                  <td className={tdR}>{fmt(f.totali.usciteOra)}</td>
                  <td className={tdR}>{fmt(f.totali.saldoDopoProssimi)}</td>
                  <td className={`${tdR} font-semibold`}>{fmt(f.totali.cassaFinale)}</td>
                  <td className={td}>{f.nota ?? ""}</td>
                  <td className={td}>
                    <button
                      type="button"
                      title={t("common.delete")}
                      onClick={() => eliminaFotografia(f.id)}
                      className="text-muted-foreground hover:text-status-absent"
                    >
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        {fotoAperta && (
          <div className="mt-3 rounded-xl border border-border bg-background/60 p-3">
            <div className="mb-1 text-[13px] font-semibold text-foreground">
              {t("cpr.fotoDel")} {fmtData(fotoAperta.data)} · {fotoAperta.righe.length}{" "}
              {t("cpr.righe")}
            </div>
            <table className="w-full border-collapse text-[12px]">
              <tbody>
                {fotoAperta.righe.map((r) => (
                  <tr key={r.k} className="border-b border-border/40">
                    <td className="px-2 py-0.5">{r.e ? "↑" : "↓"}</td>
                    <td className="px-2 py-0.5">{secchioLabel(r.b, t)}</td>
                    <td className="px-2 py-0.5 text-foreground">{r.c}</td>
                    <td className="px-2 py-0.5 text-muted-foreground">{r.d}</td>
                    <td className="px-2 py-0.5">{fmtData(r.s)}</td>
                    <td className="px-2 py-0.5 text-right tabular-nums text-foreground">
                      {fmt(r.i)}
                    </td>
                    <td className="px-2 py-0.5 text-muted-foreground">
                      {r.si ? t("cpr.standbyIncasso") : ""} {r.n ?? ""}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function secchioLabel(s: SecchioCassa, t: TFn): string {
  return s === "ora" ? t("cpr.ora") : s === "prossimi" ? t("cpr.prossimi") : t("cpr.standby");
}

function Kpi({
  label,
  value,
  neg,
  bold,
}: {
  label: string;
  value: number;
  neg?: boolean;
  bold?: boolean;
}) {
  return (
    <div className="rounded-xl border border-border bg-card px-4 py-3 shadow-[var(--shadow-card)]">
      <div className="text-[11px] uppercase tracking-wider text-muted-foreground">{label}</div>
      <div
        className={`mt-1 tabular-nums ${bold ? "text-lg font-semibold" : "text-base"} ${neg ? "text-status-absent" : value < 0 ? "text-status-absent" : "text-foreground"}`}
      >
        {neg && value > 0 ? "−" : ""}
        {fmt(Math.abs(value) < 0.005 ? 0 : value)}
      </div>
    </div>
  );
}

function TabellaRighe({
  righe,
  oggi,
  entrate,
  onSecchio,
  onStandbyIncasso,
  onNota,
  onContestata,
  onImporto,
  onTogli,
  onElimina,
}: {
  righe: readonly RigaCassa[];
  oggi: string;
  entrate?: boolean;
  onSecchio: (r: RigaCassa, b: SecchioCassa) => void;
  onStandbyIncasso?: (r: RigaCassa, v: boolean) => void;
  onNota: (r: RigaCassa, n: string) => void;
  onContestata: (r: RigaCassa, v: boolean) => void;
  onImporto: (r: RigaCassa, i: number) => void;
  onTogli: (r: RigaCassa) => void;
  onElimina: (r: RigaCassa) => void;
}) {
  const { t } = useLang();
  const [notaIn, setNotaIn] = useState<Record<string, string>>({});
  const [impIn, setImpIn] = useState<Record<string, string>>({});
  if (!righe.length)
    return <p className="text-[12px] text-muted-foreground">{t("cpr.nessunaRiga")}</p>;
  return (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[980px] border-collapse">
        <thead>
          <tr className="border-b border-border">
            <th className={th}>{t("cpr.scadenza")}</th>
            <th className={th}>{t("cpr.controparte")}</th>
            <th className={th}>{t("cpr.descrizione")}</th>
            <th className={thR}>{t("cpr.importo")}</th>
            <th className={th}>{t("cpr.secchio")}</th>
            <th className={th}>{t("cpr.nota")}</th>
            <th className={th}></th>
          </tr>
        </thead>
        <tbody>
          {righe.map((r) => {
            const scaduta = r.scadenza < oggi;
            const editabile = r.tipo === "stipendi" || r.tipo === "fisso";
            return (
              <tr
                key={r.chiave}
                className={`border-b border-border/60 ${r.secchio === "standby" ? "bg-amber-500/5" : ""} ${r.contestata ? "opacity-70" : ""}`}
              >
                <td className={`${td} whitespace-nowrap ${scaduta ? "text-status-absent" : ""}`}>
                  {fmtData(r.scadenza)}
                  {r.certa && (
                    <span className="ml-1 rounded bg-secondary px-1 text-[10px] uppercase text-muted-foreground">
                      {t("cpr.certaBreve")}
                    </span>
                  )}
                </td>
                <td className={td}>
                  {r.controparte}
                  {r.contestata && (
                    <span className="ml-1 rounded bg-status-absent/10 px-1 text-[10px] uppercase text-status-absent">
                      {t("cpr.contestata")}
                    </span>
                  )}
                </td>
                <td className={`${td} text-muted-foreground`}>{r.descrizione}</td>
                <td className={tdR}>
                  {editabile ? (
                    <input
                      value={impIn[r.chiave] ?? fmt(r.importo)}
                      onChange={(e) => setImpIn((s) => ({ ...s, [r.chiave]: e.target.value }))}
                      onBlur={() => {
                        const v = impIn[r.chiave];
                        if (v == null) return;
                        const n = Number(v.replace(/\./g, "").replace(",", "."));
                        setImpIn((s) => {
                          const c = { ...s };
                          delete c[r.chiave];
                          return c;
                        });
                        if (Number.isFinite(n) && n >= 0) onImporto(r, Math.round(n * 100) / 100);
                      }}
                      className={`${inputCls} w-28 text-right`}
                      title={
                        r.importo !== r.importoOrigine
                          ? `${t("cpr.importoOrigine")} ${fmt(r.importoOrigine)}`
                          : ""
                      }
                    />
                  ) : (
                    fmt(r.importo)
                  )}
                  {r.importo !== r.importoOrigine && (
                    <span className="ml-1 text-[10px] text-muted-foreground">*</span>
                  )}
                </td>
                <td className={`${td} whitespace-nowrap`}>
                  {entrate && onStandbyIncasso ? (
                    <label className="flex items-center gap-1 text-xs">
                      <input
                        type="checkbox"
                        checked={r.standbyIncasso}
                        onChange={(e) => onStandbyIncasso(r, e.target.checked)}
                      />{" "}
                      {t("cpr.standbyIncasso")}
                    </label>
                  ) : (
                    <div className="flex rounded-lg border border-border overflow-hidden text-[11px]">
                      {(["ora", "prossimi", "standby"] as const).map((b) => (
                        <button
                          key={b}
                          type="button"
                          onClick={() => onSecchio(r, b)}
                          className={`px-2 py-0.5 ${r.secchio === b ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
                        >
                          {secchioLabel(b, t)}
                        </button>
                      ))}
                    </div>
                  )}
                </td>
                <td className={td}>
                  <input
                    value={notaIn[r.chiave] ?? r.nota}
                    onChange={(e) => setNotaIn((s) => ({ ...s, [r.chiave]: e.target.value }))}
                    onBlur={() => {
                      const v = notaIn[r.chiave];
                      if (v == null) return;
                      setNotaIn((s) => {
                        const c = { ...s };
                        delete c[r.chiave];
                        return c;
                      });
                      if (v.trim() !== r.nota) onNota(r, v.trim());
                    }}
                    placeholder={t("cpr.notaPh")}
                    className={`${inputCls} w-48`}
                  />
                </td>
                <td className={`${td} whitespace-nowrap text-[11px]`}>
                  <button
                    type="button"
                    onClick={() => onContestata(r, !r.contestata)}
                    className="mr-2 text-muted-foreground underline-offset-2 hover:underline"
                  >
                    {r.contestata ? t("cpr.nonContestata") : t("cpr.contesta")}
                  </button>
                  {r.tipo === "manuale" || r.tipo === "amex" ? (
                    <button
                      type="button"
                      onClick={() => onElimina(r)}
                      className="text-muted-foreground hover:text-status-absent"
                      title={t("common.delete")}
                    >
                      <Trash2 className="inline h-3.5 w-3.5" />
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => onTogli(r)}
                      className="text-muted-foreground underline-offset-2 hover:underline"
                      title={t("cpr.togliTip")}
                    >
                      {t("cpr.togli")}
                    </button>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
