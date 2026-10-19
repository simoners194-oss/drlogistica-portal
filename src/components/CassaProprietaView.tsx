// DR Portal — Finanza → Flussi → vista "Cassa proprietà".
// IL FOGLIO del commercialista, tale e quale: colonne Entrate | Note |
// Saldo | Uscite | 1-5 del mese dopo | Standby | Note | Note2, i totali in
// fondo con le sue etichette, le fotografie come le linguette dei fogli
// Excel ("1210"/"1510"). Niente cruscotti: la proprietà legge solo questo.
// Si lavora dentro le celle: clic su una cella vuota delle tre colonne per
// spostare un'uscita, clic sull'importo per correggerlo (stipendi, fiscale,
// fissi, attesi), Note2 è una cella scrivibile, la × a bordo riga toglie la
// riga. Logica in lib/cassa-logic; dati in FinanzaData/cassa.json.
import { useEffect, useMemo, useState, type ReactNode } from "react";
import { toast } from "sonner";
import { Camera, Loader2, Plus, X } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import type { FatturaConStato } from "@/lib/fatture-logic";
import type { CostoFisso, FornitoreInfo } from "@/lib/flussi-logic";
import type { ScadenzaFiscale } from "@/lib/fiscale-logic";
import type { SpMovimento } from "@/lib/sharepoint.server";
import { spEbSaldoCache } from "@/lib/sharepoint.functions";
import { spStipendiGet } from "@/lib/stipendi.functions";
import { spCassaGet, spCassaSalva } from "@/lib/cassa.functions";
import type { StipendiDb } from "@/lib/stipendi-logic";
import {
  costruisciCassa,
  creaFotografia,
  emptyCassaDb,
  etichettaProssimi,
  etichettaProssimiBreve,
  limiteDefault,
  stipendiDaPagare,
  type CassaDb,
  type FotografiaCassa,
  type RigaCassa,
  type RigaManualeCassa,
  type SecchioCassa,
  type StatoRigaCassa,
  type TipoRigaCassa,
  type TotaliCassa,
} from "@/lib/cassa-logic";

type TFn = ReturnType<typeof useLang>["t"];

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtData = (iso: string) => (iso ? iso.slice(0, 10).split("-").reverse().join("/") : "—");
const ddmm = (iso: string) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : "");
const oggiISO = () => new Date().toISOString().slice(0, 10);
const parseNum = (s: string): number | null => {
  const n = Number(s.trim().replace(/\./g, "").replace(",", "."));
  return s.trim() && Number.isFinite(n) ? n : null;
};
const inputCls =
  "rounded-lg border border-border bg-background px-2 py-1 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";
/** Ultimo saldo banca visto in questa sessione: si dipinge subito al rientro
 *  nella vista, poi si riaggiorna dalla cache del server. */
let saldoMemo: { saldo: number; al?: string } | null = null;

/** Riga del foglio, uguale per la vista viva e per una fotografia. */
interface Cella {
  chiave: string;
  tipo: TipoRigaCassa;
  nome: string;
  dettaglio: string;
  importo: number;
  importoOrigine: number;
  secchio: SecchioCassa;
  entrata: boolean;
  attesa: boolean;
  nota: string;
  data: string;
  certa: boolean;
  documenti: number;
  editabile: boolean;
}

const daRiga = (r: RigaCassa): Cella => ({
  chiave: r.chiave,
  tipo: r.tipo,
  nome: r.controparte,
  dettaglio: r.descrizione,
  importo: r.importo,
  importoOrigine: r.importoOrigine,
  secchio: r.secchio,
  entrata: r.entrata,
  attesa: r.standbyIncasso,
  nota: r.nota,
  data: r.scadenza,
  certa: r.certa,
  documenti: r.documenti,
  editabile: r.editabile,
});

/** La colonna Entrate in due blocchi col loro totale: "Già entrate" (in
 *  banca, dentro il saldo) ed "Entranti" (attese = standby incasso) —
 *  Simone 19/10/2026. */
type SlotE = { kind: "head"; label: string; value: number } | { kind: "riga"; c: Cella };
function slotsEntrate(entrate: readonly Cella[], totali: TotaliCassa, t: TFn): SlotE[] {
  return [
    { kind: "head", label: t("cpr.giaEntrate"), value: totali.entrateRicevute },
    ...entrate.filter((c) => !c.attesa).map((c): SlotE => ({ kind: "riga", c })),
    { kind: "head", label: t("cpr.entranti"), value: totali.entrateStandby },
    ...entrate.filter((c) => c.attesa).map((c): SlotE => ({ kind: "riga", c })),
  ];
}

function daFoto(r: FotografiaCassa["righe"][number]): Cella {
  return {
    chiave: r.k,
    tipo: r.t,
    nome: r.c,
    dettaglio: r.d,
    importo: r.i,
    importoOrigine: r.i,
    secchio: r.b,
    entrata: r.e,
    attesa: Boolean(r.si),
    nota: r.n ?? "",
    data: r.s,
    certa: false,
    documenti: r.nd ?? 1,
    editabile: false,
  };
}

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
  const [saldoBanca, setSaldoBanca] = useState<{ saldo: number; al?: string } | null>(saldoMemo);
  const [saldoErr, setSaldoErr] = useState(false);
  const [oggi, setOggi] = useState(oggiISO());
  const [limite, setLimite] = useState("");
  const [saldoIn, setSaldoIn] = useState<string | null>(null);
  const [salvando, setSalvando] = useState(false);
  const [mostraOltre, setMostraOltre] = useState(false);
  const [fotoId, setFotoId] = useState<string | null>(null);

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
    // Saldo dalla cache del server (una riga SharePoint, niente banca e
    // niente rilettura dei movimenti): arriva subito.
    spEbSaldoCache()
      .then((s) => {
        if (s) {
          saldoMemo = { saldo: s.saldo, al: s.aggiornatoAl };
          setSaldoBanca(saldoMemo);
        } else setSaldoErr(true);
      })
      .catch(() => setSaldoErr(true));
  }, [t]);

  // Saldo: scritto a mano (vince), altrimenti dalla banca, altrimenti
  // l'ultimo scritto a mano e salvato.
  const saldo = useMemo(() => {
    if (saldoIn != null) {
      const n = parseNum(saldoIn);
      if (n != null) return n;
    }
    if (saldoBanca) return saldoBanca.saldo;
    return db?.saldoManuale?.importo ?? 0;
  }, [saldoIn, saldoBanca, db]);

  const stipendi = useMemo(
    () =>
      stipendiDb
        ? stipendiDaPagare(stipendiDb, (limite || oggi).slice(0, 7), oggi.slice(0, 7))
        : [],
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
        id: m.id,
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
  const aggiungiManuale = (riga: Omit<RigaManualeCassa, "id">) => {
    if (!db) return;
    void salva({
      ...db,
      manuali: [...db.manuali, { ...riga, id: `M${Date.now().toString(36)}` }],
    });
  };
  const aggiornaManuale = (id: string, patch: Partial<RigaManualeCassa>) => {
    if (!db) return;
    void salva({
      ...db,
      manuali: db.manuali.map((m) => (m.id === id ? { ...m, ...patch } : m)),
    });
  };
  const eliminaManuale = (id: string) => {
    if (!db) return;
    const stati = { ...db.stati };
    delete stati[`man|${id}`];
    void salva({ ...db, manuali: db.manuali.filter((m) => m.id !== id), stati });
  };
  const salvaSaldoManuale = () => {
    if (!db || saldoIn == null) return;
    const n = parseNum(saldoIn);
    if (n == null) {
      setSaldoIn(null);
      return;
    }
    void salva({ ...db, saldoManuale: { importo: Math.round(n * 100) / 100, data: oggi } });
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
    if (fotoId === id) setFotoId(null);
  };
  const cambiaLimite = (iso: string) => {
    if (!iso) return;
    setLimite(iso);
    const g = Number(iso.slice(8, 10));
    if (db && iso.slice(0, 7) === limiteDefault(oggi, 1).slice(0, 7) && g >= 1 && g <= 28)
      void salva({ ...db, giornoLimite: g });
  };

  // Spostare un'uscita in una colonna: se è la colonna automatica, si toglie
  // la forzatura; per le righe manuali si aggiorna la riga.
  const sposta = (c: Cella, col: SecchioCassa) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { secchio: col });
      return;
    }
    const auto = esito?.righe.find((r) => r.chiave === c.chiave)?.secchioAuto;
    setStato(c.chiave, { b: col === auto ? undefined : col });
  };
  const correggi = (c: Cella, importo: number) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { importo });
      return;
    }
    setStato(c.chiave, { i: Math.abs(importo - c.importoOrigine) < 0.005 ? undefined : importo });
  };
  const annota = (c: Cella, nota: string) => {
    if (c.tipo === "manuale" || c.tipo === "amex") {
      aggiornaManuale(c.chiave.slice(4), { nota: nota || undefined });
      return;
    }
    setStato(c.chiave, { n: nota });
  };
  const togli = (c: Cella) => {
    if (c.tipo === "manuale" || c.tipo === "amex") eliminaManuale(c.chiave.slice(4));
    else setStato(c.chiave, { x: true });
  };

  if (!db || !esito) {
    return (
      <div className="flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );
  }

  const foto = fotoId ? (db.fotografie.find((f) => f.id === fotoId) ?? null) : null;
  const entrate = foto
    ? foto.righe.filter((r) => r.e).map(daFoto)
    : esito.righe.filter((r) => r.entrata).map(daRiga);
  const uscite = foto
    ? foto.righe.filter((r) => !r.e).map(daFoto)
    : esito.righe.filter((r) => !r.entrata).map(daRiga);
  const totali = foto ? foto.totali : esito.totali;
  const limiteVista = foto ? foto.limite : limite;
  const saldoVista = foto ? foto.saldo : saldo;
  const dataVista = foto ? foto.data : oggi;

  const righeCsv = () => {
    const out: (string | number)[][] = [];
    const sl = slotsEntrate(entrate, totali, t);
    const n = Math.max(sl.length, uscite.length);
    out.push([totali.entrate, "", saldoVista, "", "", "", "", ""]);
    for (let i = 0; i < n; i++) {
      const s = sl[i];
      const e = s?.kind === "riga" ? s.c : undefined;
      const u = uscite[i];
      out.push([
        s?.kind === "head" ? s.value : e ? e.importo : "",
        s?.kind === "head"
          ? s.label
          : e
            ? `${e.nome}${e.attesa ? ` (${t("cpr.attesa")})` : ""}`
            : "",
        "",
        u && u.secchio === "ora" ? u.importo : "",
        u && u.secchio === "prossimi" ? u.importo : "",
        u && u.secchio === "standby" ? u.importo : "",
        u ? u.nome : "",
        u ? [u.dettaglio === "CDC" ? "CDC" : "", u.nota].filter(Boolean).join(" ") : "",
      ]);
    }
    out.push([
      totali.entrateStandby,
      "",
      "",
      totali.usciteOra,
      totali.usciteProssimi,
      totali.usciteStandby,
      totali.totaleDaPagare,
      totali.cassaConStandby,
    ]);
    out.push([
      t("cpr.lblStandbyIncasso"),
      "",
      "",
      t("cpr.lblDaPagare"),
      `${t("cpr.lblDaPagare")} ${etichettaProssimiBreve(limiteVista)}`,
      t("cpr.lblStandbyDaPagare"),
      t("cpr.lblTotaleDaPagare"),
      t("cpr.lblCassa"),
    ]);
    out.push(["", "", "", totali.saldoDopoOra, totali.saldoDopoProssimi, "", "", ""]);
    out.push(["", "", "", t("cpr.lblSaldoDopo"), "", "", "", ""]);
    return out;
  };

  const saldoTesto =
    saldoIn ??
    (saldoBanca ? fmt(saldoBanca.saldo) : db.saldoManuale ? fmt(db.saldoManuale.importo) : "");
  const saldoNota = saldoBanca
    ? `${t("cpr.saldoDaBanca")} ${saldoBanca.al ? fmtData(saldoBanca.al) : ""}`
    : saldoErr && db.saldoManuale
      ? `${t("cpr.saldoScritto")} ${fmtData(db.saldoManuale.data)}`
      : saldoErr
        ? t("cpr.saldoNonDisponibile")
        : "";

  return (
    <div className="space-y-4">
      {/* Banda in alto: saldo grande, i tre numeri del foglio, data e limite */}
      <div className="rounded-2xl border border-border bg-card px-5 py-4 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-x-8 gap-y-4">
          <div>
            <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
              {t("cpr.saldoBanca")}
              {foto && <span className="ml-1 normal-case">· {fmtData(foto.data)}</span>}
            </div>
            {foto ? (
              <div className="mt-0.5 text-3xl font-bold tabular-nums text-foreground">
                {fmt(foto.saldo)} €
              </div>
            ) : (
              <div className="mt-0.5 flex items-baseline gap-1">
                <input
                  value={saldoTesto}
                  onChange={(e) => setSaldoIn(e.target.value)}
                  onBlur={salvaSaldoManuale}
                  placeholder="0,00"
                  title={t("cpr.saldoBanca")}
                  className="w-52 border-b-2 border-dotted border-border bg-transparent text-right text-3xl font-bold tabular-nums text-foreground focus:border-primary focus:outline-none"
                />
                <span className="text-xl font-semibold text-muted-foreground">€</span>
              </div>
            )}
            <div className="mt-1 text-[11px] text-muted-foreground">{saldoNota}</div>
          </div>
          <Stat label={t("cpr.lblDaPagare")} value={totali.usciteOra} tono="rosso" />
          <Stat label={t("cpr.lblSaldoDopo")} value={totali.saldoDopoOra} tono="blu" />
          <Stat label={t("cpr.lblCassa")} value={totali.cassaConStandby} tono="segno" />
          <div className="ml-auto flex flex-wrap items-center gap-x-4 gap-y-2 text-[13px]">
            <label className="flex items-center gap-2 text-muted-foreground">
              {t("cpr.situazioneAl")}
              <input
                type="date"
                value={oggi}
                onChange={(e) => {
                  const v = e.target.value || oggiISO();
                  setOggi(v);
                  setLimite(limiteDefault(v, db.giornoLimite));
                }}
                className={inputCls}
              />
            </label>
            <label className="flex items-center gap-2 text-muted-foreground">
              {t("cpr.prossimiFinoAl")}
              <input
                type="date"
                value={limite}
                onChange={(e) => cambiaLimite(e.target.value)}
                className={inputCls}
              />
            </label>
            {salvando && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
            <button
              type="button"
              onClick={salvaFotografia}
              className="inline-flex items-center gap-1 rounded-lg border border-primary bg-primary px-3 py-1.5 text-primary-foreground shadow-sm hover:opacity-90"
            >
              <Camera className="h-4 w-4" /> {t("cpr.salvaFoto")}
            </button>
            <EsportaStampa
              nomeFile={`cassa-proprieta-${dataVista}`}
              testata={[
                t("cpr.colEntrate"),
                t("cpr.colNote"),
                t("cpr.colSaldo"),
                t("cpr.colUscite"),
                etichettaProssimi(limiteVista),
                t("cpr.colStandby"),
                t("cpr.colNote"),
                t("cpr.colNote2"),
              ]}
              righe={righeCsv}
            />
          </div>
        </div>
      </div>

      {/* Il foglio */}
      <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]">
        <Foglio
          t={t}
          entrate={entrate}
          uscite={uscite}
          totali={totali}
          saldo={saldoVista}
          limite={limiteVista}
          readOnly={Boolean(foto)}
          onSposta={sposta}
          onImporto={correggi}
          onNota={annota}
          onTogli={togli}
          onAggiungi={aggiungiManuale}
          oggi={oggi}
        />
        {/* Linguette come i fogli di Excel */}
        <div className="flex flex-wrap items-center gap-1 border-t border-border bg-muted/50 px-3 pt-1.5 pb-2 text-[12px]">
          <button
            type="button"
            onClick={() => setFotoId(null)}
            className={`rounded-b-lg border border-t-0 px-3 py-1 ${!foto ? "border-primary/40 bg-card font-semibold text-primary shadow-sm" : "border-transparent text-muted-foreground hover:bg-card/70"}`}
          >
            {t("cpr.oggi")} · {ddmm(oggi)}
          </button>
          {db.fotografie.map((f) => (
            <span
              key={f.id}
              className={`group inline-flex items-center gap-1 rounded-b-lg border border-t-0 px-3 py-1 ${foto?.id === f.id ? "border-primary/40 bg-card font-semibold text-primary shadow-sm" : "border-transparent text-muted-foreground hover:bg-card/70"}`}
            >
              <button type="button" onClick={() => setFotoId(f.id)} title={f.nota ?? ""}>
                {ddmm(f.data)}
                {f.nota ? ` · ${f.nota.slice(0, 18)}` : ""}
              </button>
              <button
                type="button"
                onClick={() => eliminaFotografia(f.id)}
                className="opacity-0 group-hover:opacity-100 hover:text-status-absent"
                title={t("common.delete")}
              >
                <X className="h-3 w-3" />
              </button>
            </span>
          ))}
        </div>
      </div>

      {!foto && esito.oltre.length > 0 && (
        <div className="text-[12px] text-muted-foreground">
          <button
            type="button"
            onClick={() => setMostraOltre((v) => !v)}
            className="underline-offset-2 hover:underline"
          >
            {mostraOltre ? "▾" : "▸"} {t("cpr.oltreIl")} {ddmm(limite)}: {esito.oltre.length}{" "}
            {t("cpr.voci")} · {fmt(esito.totali.usciteOltre)}
          </button>
          {mostraOltre && (
            <table className="mt-1 border-collapse text-[12px]">
              <tbody>
                {esito.oltre.map((r) => (
                  <tr key={r.chiave} className="border-b border-border/40">
                    <td className="py-0.5 pr-3 whitespace-nowrap">{ddmm(r.scadenza)}</td>
                    <td className="py-0.5 pr-3 text-foreground">
                      {r.controparte}
                      {r.documenti > 1 && r.tipo !== "stipendi" && (
                        <span className="ml-1 text-muted-foreground">×{r.documenti}</span>
                      )}
                    </td>
                    <td className="py-0.5 pr-3 text-right tabular-nums text-foreground">
                      {fmt(r.importo)}
                    </td>
                    <td className="py-0.5 pr-2">
                      <button
                        type="button"
                        className="rounded border border-border px-1.5 hover:bg-muted"
                        onClick={() => setStato(r.chiave, { b: "prossimi" })}
                      >
                        {t("cpr.portaNeiProssimi")}
                      </button>
                    </td>
                    <td className="py-0.5">
                      <button
                        type="button"
                        className="rounded border border-border px-1.5 hover:bg-muted"
                        onClick={() => setStato(r.chiave, { b: "standby" })}
                      >
                        {t("cpr.parcheggia")}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </div>
  );
}

function Stat({
  label,
  value,
  tono,
}: {
  label: string;
  value: number;
  tono: "rosso" | "blu" | "segno";
}) {
  const colore =
    tono === "rosso"
      ? "text-rose-600 dark:text-rose-400"
      : tono === "blu"
        ? "text-primary"
        : value < 0
          ? "text-rose-600 dark:text-rose-400"
          : "text-emerald-700 dark:text-emerald-400";
  return (
    <div>
      <div className="max-w-[220px] text-[11px] uppercase tracking-wider text-muted-foreground">
        {label}
      </div>
      <div className={`mt-0.5 text-xl font-semibold tabular-nums ${colore}`}>{fmt(value)} €</div>
    </div>
  );
}

// --- il foglio ---------------------------------------------------------------
// Tinte delle colonne (chiare in body, piene in testa): entrate verde,
// uscite rosso, prossimi ambra, standby ardesia, note neutre.
const TINTA = {
  entrate: {
    testa: "bg-emerald-600 text-white",
    col: "bg-emerald-50/60 dark:bg-emerald-950/25",
    num: "text-emerald-800 dark:text-emerald-300",
  },
  ora: {
    testa: "bg-rose-600 text-white",
    col: "bg-rose-50/60 dark:bg-rose-950/25",
    num: "text-rose-700 dark:text-rose-300",
  },
  prossimi: {
    testa: "bg-amber-500 text-white",
    col: "bg-amber-50/70 dark:bg-amber-950/25",
    num: "text-amber-700 dark:text-amber-300",
  },
  standby: {
    testa: "bg-slate-500 text-white",
    col: "bg-slate-100/70 dark:bg-slate-800/40",
    num: "text-slate-600 dark:text-slate-300",
  },
  note: { testa: "bg-muted text-foreground", col: "", num: "" },
} as const;
const cel = "border-b border-r border-border/60 px-3 py-1.5 align-middle last:border-r-0";
const celN = `${cel} text-right tabular-nums whitespace-nowrap text-[13.5px]`;
const celH = "px-3 py-2 text-left text-[12px] font-semibold uppercase tracking-wide";
const celT = `${celN} py-2.5 text-[15px] font-bold`;
const celL = `${cel} py-1 text-[10.5px] uppercase tracking-wider text-muted-foreground`;

function Foglio({
  t,
  entrate,
  uscite,
  totali,
  saldo,
  limite,
  readOnly,
  oggi,
  onSposta,
  onImporto,
  onNota,
  onTogli,
  onAggiungi,
}: {
  t: TFn;
  entrate: readonly Cella[];
  uscite: readonly Cella[];
  totali: TotaliCassa;
  saldo: number;
  limite: string;
  readOnly: boolean;
  oggi: string;
  onSposta: (c: Cella, col: SecchioCassa) => void;
  onImporto: (c: Cella, importo: number) => void;
  onNota: (c: Cella, nota: string) => void;
  onTogli: (c: Cella) => void;
  onAggiungi: (r: Omit<RigaManualeCassa, "id">) => void;
}) {
  const [nuova, setNuova] = useState<null | {
    entrata: boolean;
    descrizione: string;
    importo: string;
    data: string;
    secchio: SecchioCassa;
    amex: boolean;
  }>(null);
  const slotsE = slotsEntrate(entrate, totali, t);
  const n = Math.max(slotsE.length, uscite.length);
  const righe = Array.from({ length: n }, (_, i) => i);
  const breve = etichettaProssimiBreve(limite);

  const conferma = () => {
    if (!nuova) return;
    const imp = parseNum(nuova.importo);
    if (!nuova.descrizione.trim() || imp == null || imp === 0) {
      toast.error(t("cpr.manualeIncompleta"));
      return;
    }
    onAggiungi({
      tipo: nuova.amex ? "amex" : "manuale",
      descrizione: nuova.descrizione.trim(),
      importo: Math.abs(imp),
      data: nuova.data || oggi,
      secchio: nuova.entrata ? "ora" : nuova.secchio,
      entrata: nuova.entrata || undefined,
    });
    setNuova(null);
  };
  const apri = (entrata: boolean) =>
    setNuova({ entrata, descrizione: "", importo: "", data: oggi, secchio: "ora", amex: false });

  const cellaUscita = (u: Cella | undefined, col: SecchioCassa) => {
    const tinta = TINTA[col];
    if (!u) return <td className={`${celN} ${tinta.col}`} />;
    if (u.secchio === col)
      return (
        <td className={`${celN} ${tinta.col} ${tinta.num} font-medium`}>
          <Importo c={u} readOnly={readOnly} onImporto={onImporto} t={t} />
        </td>
      );
    if (readOnly) return <td className={`${celN} ${tinta.col}`} />;
    return (
      <td className={`${celN} ${tinta.col} p-0`}>
        <button
          type="button"
          onClick={() => onSposta(u, col)}
          title={t("cpr.spostaQui")}
          className="h-full min-h-[34px] w-full px-3 text-right text-transparent hover:bg-primary/10 hover:text-muted-foreground"
        >
          →
        </button>
      </td>
    );
  };

  return (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[1120px] border-collapse text-[13px]">
        <colgroup>
          <col className="w-[125px]" />
          <col className="w-[215px]" />
          <col className="w-[125px]" />
          <col className="w-[125px]" />
          <col className="w-[125px]" />
          <col className="w-[125px]" />
          <col className="w-[300px]" />
          <col />
        </colgroup>
        <thead>
          <tr>
            <th
              colSpan={3}
              className={`${celH} ${TINTA.entrate.testa} border-r border-white/30 text-center text-[11px]`}
            >
              {t("cpr.colEntrate")}
            </th>
            <th
              colSpan={5}
              className={`${celH} bg-primary text-center text-[11px] text-primary-foreground`}
            >
              {t("cpr.colUscite")}
            </th>
          </tr>
          <tr className="border-b border-border">
            <th
              className={`${celH} ${TINTA.entrate.col} ${TINTA.entrate.num} border-r border-border/60 text-right`}
            >
              {t("cpr.colEntrate")}
            </th>
            <th
              className={`${celH} ${TINTA.entrate.col} border-r border-border/60 text-foreground`}
            >
              {t("cpr.colNote")}
            </th>
            <th
              className={`${celH} border-r border-border/60 bg-muted/60 text-right text-foreground`}
            >
              {t("cpr.colSaldo")}
            </th>
            <th
              className={`${celH} ${TINTA.ora.col} ${TINTA.ora.num} border-r border-border/60 text-right`}
            >
              {t("cpr.colUscite")}
            </th>
            <th
              className={`${celH} ${TINTA.prossimi.col} ${TINTA.prossimi.num} border-r border-border/60 text-right normal-case`}
            >
              {etichettaProssimi(limite)}
            </th>
            <th
              className={`${celH} ${TINTA.standby.col} ${TINTA.standby.num} border-r border-border/60 text-right`}
            >
              {t("cpr.colStandby")}
            </th>
            <th className={`${celH} border-r border-border/60 bg-muted/60 text-foreground`}>
              {t("cpr.colNote")}
            </th>
            <th className={`${celH} bg-muted/60 text-foreground`}>{t("cpr.colNote2")}</th>
          </tr>
        </thead>
        <tbody>
          {/* riga 2: totale entrate e saldo, come nel foglio */}
          <tr className="bg-muted/30">
            <td className={`${celT} ${TINTA.entrate.num}`}>{fmt(totali.entrate)}</td>
            <td className={`${cel} text-[11px] uppercase tracking-wider text-muted-foreground`}>
              {t("cpr.lblTotaleEntrate")}
            </td>
            <td className={`${celT} text-foreground`}>{fmt(saldo)}</td>
            <td className={`${cel} ${TINTA.ora.col}`} />
            <td className={`${cel} ${TINTA.prossimi.col}`} />
            <td className={`${cel} ${TINTA.standby.col}`} />
            <td className={cel} />
            <td className={cel} />
          </tr>
          {righe.map((i) => {
            const slot = slotsE[i];
            const e = slot?.kind === "riga" ? slot.c : undefined;
            const u = uscite[i];
            return (
              <tr key={i} className="group/r odd:bg-background/40 hover:bg-primary/5">
                {slot?.kind === "head" ? (
                  <td
                    colSpan={2}
                    className={`${cel} bg-emerald-100/70 dark:bg-emerald-900/30 ${TINTA.entrate.num}`}
                  >
                    <span className="flex items-center justify-between gap-2">
                      <span className="text-[11px] font-semibold uppercase tracking-wider">
                        {slot.label}
                      </span>
                      <span className="text-[13.5px] font-bold tabular-nums">
                        {fmt(slot.value)}
                      </span>
                    </span>
                  </td>
                ) : (
                  <>
                    <td
                      className={`${celN} ${TINTA.entrate.col} ${e?.attesa ? "italic text-muted-foreground" : `${TINTA.entrate.num} font-medium`}`}
                    >
                      {e && <Importo c={e} readOnly={readOnly} onImporto={onImporto} t={t} />}
                    </td>
                    <td
                      className={`${cel} ${TINTA.entrate.col} ${e?.attesa ? "text-muted-foreground" : "text-foreground"}`}
                    >
                      {e && (
                        <span className="flex items-center gap-1.5">
                          <span className="truncate" title={e.dettaglio}>
                            {e.nome}
                          </span>
                          {e.attesa ? (
                            <span className="shrink-0 rounded-full border border-emerald-300 px-1.5 text-[10px] text-emerald-700 dark:border-emerald-700 dark:text-emerald-300">
                              {t("cpr.attesa")}
                            </span>
                          ) : (
                            <span className="shrink-0 text-[10.5px] text-muted-foreground">
                              {ddmm(e.data)}
                            </span>
                          )}
                          {!readOnly && <Togli onClick={() => onTogli(e)} t={t} />}
                        </span>
                      )}
                    </td>
                  </>
                )}
                <td className={`${cel} bg-muted/20`} />
                {cellaUscita(u, "ora")}
                {cellaUscita(u, "prossimi")}
                {cellaUscita(u, "standby")}
                <td className={`${cel} text-foreground`}>
                  {u && (
                    <span className="flex items-center gap-1.5">
                      <span className="truncate font-medium" title={u.dettaglio}>
                        {u.nome}
                      </span>
                      {u.documenti > 1 && u.tipo !== "stipendi" && (
                        <span className="shrink-0 text-[10.5px] text-muted-foreground">
                          ×{u.documenti}
                        </span>
                      )}
                      {u.tipo === "passiva" && u.certa && (
                        <span className="shrink-0 rounded bg-secondary px-1 text-[9px] uppercase text-muted-foreground">
                          auto
                        </span>
                      )}
                      {(u.tipo === "stipendi" || u.tipo === "fisso" || u.tipo === "fiscale") && (
                        <span className="shrink-0 truncate text-[10.5px] text-muted-foreground">
                          {u.dettaglio}
                        </span>
                      )}
                    </span>
                  )}
                </td>
                <td className={`${cel} p-0`}>
                  {u && (
                    <span className="flex items-center">
                      {u.dettaglio === "CDC" && (
                        <span className="pl-3 text-[11px] font-semibold text-muted-foreground">
                          CDC
                        </span>
                      )}
                      <Nota c={u} readOnly={readOnly} onNota={onNota} />
                      {!readOnly && (
                        <span className="pr-2">
                          <Togli onClick={() => onTogli(u)} t={t} />
                        </span>
                      )}
                    </span>
                  )}
                </td>
              </tr>
            );
          })}
          {/* riga per aggiungere */}
          {!readOnly && (
            <tr className="bg-muted/10">
              {nuova?.entrata ? (
                <>
                  <td className={`${celN} ${TINTA.entrate.col} p-1`}>
                    <input
                      autoFocus
                      value={nuova.importo}
                      onChange={(e) => setNuova({ ...nuova, importo: e.target.value })}
                      onKeyDown={(e) => e.key === "Enter" && conferma()}
                      placeholder={t("cpr.importo")}
                      className={`${inputCls} w-full text-right`}
                    />
                  </td>
                  <td className={`${cel} ${TINTA.entrate.col} p-1`}>
                    <span className="flex items-center gap-1">
                      <input
                        value={nuova.descrizione}
                        onChange={(e) => setNuova({ ...nuova, descrizione: e.target.value })}
                        onKeyDown={(e) => e.key === "Enter" && conferma()}
                        placeholder={t("cpr.descrizione")}
                        className={`${inputCls} w-full`}
                      />
                      <Ok onOk={conferma} onNo={() => setNuova(null)} t={t} />
                    </span>
                  </td>
                </>
              ) : (
                <>
                  <td className={`${cel} ${TINTA.entrate.col}`} />
                  <td className={`${cel} ${TINTA.entrate.col} py-1`}>
                    <Aggiungi onClick={() => apri(true)} t={t} />
                  </td>
                </>
              )}
              <td className={`${cel} bg-muted/20`} />
              {nuova && !nuova.entrata ? (
                <>
                  <td className={`${celN} p-1`} colSpan={3}>
                    <span className="flex items-center gap-1">
                      <input
                        autoFocus
                        value={nuova.importo}
                        onChange={(e) => setNuova({ ...nuova, importo: e.target.value })}
                        onKeyDown={(e) => e.key === "Enter" && conferma()}
                        placeholder={t("cpr.importo")}
                        className={`${inputCls} w-28 text-right`}
                      />
                      <select
                        value={nuova.secchio}
                        onChange={(e) =>
                          setNuova({ ...nuova, secchio: e.target.value as SecchioCassa })
                        }
                        className={inputCls}
                      >
                        <option value="ora">{t("cpr.colUscite")}</option>
                        <option value="prossimi">{breve}</option>
                        <option value="standby">{t("cpr.colStandby")}</option>
                      </select>
                      <input
                        type="date"
                        value={nuova.data}
                        onChange={(e) => setNuova({ ...nuova, data: e.target.value })}
                        className={inputCls}
                      />
                    </span>
                  </td>
                  <td className={`${cel} p-1`}>
                    <input
                      value={nuova.descrizione}
                      onChange={(e) => setNuova({ ...nuova, descrizione: e.target.value })}
                      onKeyDown={(e) => e.key === "Enter" && conferma()}
                      placeholder={t("cpr.descrizione")}
                      className={`${inputCls} w-full`}
                    />
                  </td>
                  <td className={`${cel} p-1`}>
                    <span className="flex items-center gap-2">
                      <label className="flex items-center gap-1 text-[12px] text-muted-foreground">
                        <input
                          type="checkbox"
                          checked={nuova.amex}
                          onChange={(e) => setNuova({ ...nuova, amex: e.target.checked })}
                        />
                        CDC
                      </label>
                      <Ok onOk={conferma} onNo={() => setNuova(null)} t={t} />
                    </span>
                  </td>
                </>
              ) : (
                <>
                  <td className={`${cel} ${TINTA.ora.col}`} />
                  <td className={`${cel} ${TINTA.prossimi.col}`} />
                  <td className={`${cel} ${TINTA.standby.col}`} />
                  <td className={`${cel} py-1`}>
                    <Aggiungi onClick={() => apri(false)} t={t} />
                  </td>
                  <td className={cel} />
                </>
              )}
            </tr>
          )}
        </tbody>
        <tfoot>
          {/* totali come nel foglio: valori e sotto le sue etichette */}
          <tr className="border-t-2 border-border bg-muted/50">
            <td className={`${celT} ${TINTA.entrate.num}`}>{fmt(totali.entrateStandby)}</td>
            <td className={cel} />
            <td className={cel} />
            <td className={`${celT} ${TINTA.ora.num}`}>{fmt(totali.usciteOra)}</td>
            <td className={`${celT} ${TINTA.prossimi.num}`}>{fmt(totali.usciteProssimi)}</td>
            <td className={`${celT} ${TINTA.standby.num}`}>{fmt(totali.usciteStandby)}</td>
            <td className={`${celT} text-foreground`}>{fmt(totali.totaleDaPagare)}</td>
            <td
              className={`${celT} ${totali.cassaConStandby < 0 ? "text-rose-600 dark:text-rose-400" : "text-emerald-700 dark:text-emerald-400"}`}
            >
              {fmt(totali.cassaConStandby)}
            </td>
          </tr>
          <tr className="bg-muted/50">
            <td className={`${celL} text-right`}>{t("cpr.lblStandbyIncasso")}</td>
            <td className={cel} />
            <td className={cel} />
            <td className={`${celL} text-right`}>{t("cpr.lblDaPagare")}</td>
            <td className={`${celL} text-right`}>
              {t("cpr.lblDaPagare")} {breve}
            </td>
            <td className={`${celL} text-right`}>{t("cpr.lblStandbyDaPagare")}</td>
            <td className={`${celL} text-right`}>{t("cpr.lblTotaleDaPagare")}</td>
            <td className={`${celL} text-right`}>{t("cpr.lblCassa")}</td>
          </tr>
          <tr>
            <td className={`${cel} py-1.5`} colSpan={8} />
          </tr>
          <tr className="bg-primary/5">
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
            <td
              className={`${celT} ${totali.saldoDopoOra < 0 ? "text-rose-600 dark:text-rose-400" : "text-primary"}`}
            >
              {fmt(totali.saldoDopoOra)}
            </td>
            <td
              className={`${celT} ${totali.saldoDopoProssimi < 0 ? "text-rose-600 dark:text-rose-400" : "text-primary"}`}
            >
              {fmt(totali.saldoDopoProssimi)}
            </td>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
          </tr>
          <tr className="bg-primary/5">
            <td className={`${cel} border-b-0`} />
            <td className={`${cel} border-b-0`} />
            <td className={`${cel} border-b-0`} />
            <td className={`${celL} border-b-0 text-right`} colSpan={2}>
              {t("cpr.lblSaldoDopo")} · {t("cpr.lblDopoProssimi")} {breve}
            </td>
            <td className={`${cel} border-b-0`} />
            <td className={`${cel} border-b-0`} />
            <td className={`${cel} border-b-0`} />
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

function Importo({
  c,
  readOnly,
  onImporto,
  t,
}: {
  c: Cella;
  readOnly: boolean;
  onImporto: (c: Cella, importo: number) => void;
  t: TFn;
}) {
  const [v, setV] = useState<string | null>(null);
  const corretto = Math.abs(c.importo - c.importoOrigine) > 0.005;
  const titolo = corretto ? `${t("cpr.importoOrigine")} ${fmt(c.importoOrigine)}` : c.dettaglio;
  if (readOnly || !c.editabile)
    return (
      <span title={titolo}>
        {fmt(c.importo)}
        {corretto && <span className="ml-0.5 text-[10px] text-muted-foreground">*</span>}
      </span>
    );
  if (v == null)
    return (
      <button
        type="button"
        onClick={() => setV(fmt(c.importo))}
        title={titolo}
        className="w-full text-right underline decoration-dotted decoration-border underline-offset-4 hover:decoration-foreground"
      >
        {fmt(c.importo)}
        {corretto && <span className="ml-0.5 text-[10px] text-muted-foreground">*</span>}
      </button>
    );
  const commit = () => {
    const n = parseNum(v);
    setV(null);
    if (n != null && n >= 0 && Math.abs(n - c.importo) > 0.005)
      onImporto(c, Math.round(n * 100) / 100);
  };
  return (
    <input
      autoFocus
      value={v}
      onChange={(e) => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") commit();
        if (e.key === "Escape") setV(null);
      }}
      className={`${inputCls} w-full text-right`}
    />
  );
}

function Nota({
  c,
  readOnly,
  onNota,
}: {
  c: Cella;
  readOnly: boolean;
  onNota: (c: Cella, nota: string) => void;
}) {
  const [v, setV] = useState<string | null>(null);
  if (readOnly) return <span className="px-3 py-1.5 text-foreground">{c.nota}</span>;
  return (
    <input
      value={v ?? c.nota}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => {
        if (v != null && v.trim() !== c.nota) onNota(c, v.trim());
        setV(null);
      }}
      onKeyDown={(e) => e.key === "Enter" && (e.target as HTMLInputElement).blur()}
      className="w-full bg-transparent px-3 py-1.5 text-[13px] text-foreground focus:bg-background focus:outline-none"
    />
  );
}

function Togli({ onClick, t }: { onClick: () => void; t: TFn }) {
  return (
    <button
      type="button"
      onClick={onClick}
      title={t("cpr.togliTip")}
      className="ml-auto shrink-0 text-muted-foreground opacity-0 group-hover/r:opacity-100 hover:text-status-absent"
    >
      <X className="h-3.5 w-3.5" />
    </button>
  );
}

function Aggiungi({ onClick, t }: { onClick: () => void; t: TFn }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="inline-flex items-center gap-1 text-[12px] text-muted-foreground hover:text-primary"
    >
      <Plus className="h-3.5 w-3.5" /> {t("cpr.aggiungi")}
    </button>
  );
}

function Ok({ onOk, onNo, t }: { onOk: () => void; onNo: () => void; t: TFn }): ReactNode {
  return (
    <span className="flex shrink-0 items-center gap-1">
      <button
        type="button"
        onClick={onOk}
        className="rounded border border-primary bg-primary px-2 py-0.5 text-[12px] text-primary-foreground"
      >
        {t("common.save")}
      </button>
      <button
        type="button"
        onClick={onNo}
        className="rounded border border-border px-2 py-0.5 text-[12px] hover:bg-muted"
      >
        {t("common.cancel")}
      </button>
    </span>
  );
}
