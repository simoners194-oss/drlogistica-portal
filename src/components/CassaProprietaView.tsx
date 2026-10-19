// DR Portal — Finanza → Flussi → vista "Cassa proprietà".
// IL FOGLIO del commercialista, tale e quale: colonne Entrate | Note |
// Saldo | Uscite | 1-5 del mese dopo | Standby | Note | Note2, i totali in
// fondo con le sue etichette, le fotografie come i fogli "1210"/"1510".
// Niente cruscotti: la proprietà legge solo questo. Si lavora dentro le
// celle: clic su una cella vuota delle tre colonne per spostare un'uscita,
// clic sull'importo per correggerlo (stipendi, fiscale, fissi, attesi),
// Note2 è una cella scrivibile, la × a bordo riga toglie la riga.
// Logica in lib/cassa-logic; dati in FinanzaData/cassa.json.
import { useEffect, useMemo, useState, type ReactNode } from "react";
import { toast } from "sonner";
import { Camera, Loader2, X } from "lucide-react";
import { useLang } from "@/lib/i18n";
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
    spEbSaldo()
      .then((s) => {
        if (s) setSaldoBanca({ saldo: s.saldo, al: s.aggiornatoAl });
        else setSaldoErr(true);
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
  const entrateVive = esito.righe.filter((r) => r.entrata).map(daRiga);
  const usciteVive = esito.righe.filter((r) => !r.entrata).map(daRiga);
  const entrate = foto ? foto.righe.filter((r) => r.e).map(daFoto) : entrateVive;
  const uscite = foto ? foto.righe.filter((r) => !r.e).map(daFoto) : usciteVive;
  const totali = foto ? foto.totali : esito.totali;
  const limiteVista = foto ? foto.limite : limite;
  const saldoVista = foto ? foto.saldo : saldo;

  const righeCsv = () => {
    const out: (string | number)[][] = [];
    const n = Math.max(entrate.length, uscite.length);
    out.push([totali.entrate, "", saldoVista, "", "", "", "", ""]);
    for (let i = 0; i < n; i++) {
      const e = entrate[i];
      const u = uscite[i];
      out.push([
        e ? e.importo : "",
        e ? `${e.nome}${e.attesa ? ` (${t("cpr.attesa")})` : ""}` : "",
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

  return (
    <div className="space-y-3">
      {/* Barra: data, saldo, limite, fotografia, CSV */}
      <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-[13px]">
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
          {t("cpr.saldoBanca")}
          <input
            value={
              saldoIn ??
              (saldoBanca
                ? fmt(saldoBanca.saldo)
                : db.saldoManuale
                  ? fmt(db.saldoManuale.importo)
                  : "")
            }
            onChange={(e) => setSaldoIn(e.target.value)}
            onBlur={salvaSaldoManuale}
            placeholder="0,00"
            className={`${inputCls} w-36 text-right tabular-nums`}
          />
          <span className="text-[11px]">
            {saldoBanca
              ? `${t("cpr.saldoDaBanca")} ${saldoBanca.al ? fmtData(saldoBanca.al) : ""}`
              : saldoErr && db.saldoManuale
                ? `${t("cpr.saldoScritto")} ${fmtData(db.saldoManuale.data)}`
                : saldoErr
                  ? t("cpr.saldoNonDisponibile")
                  : ""}
          </span>
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
            nomeFile={`cassa-proprieta-${foto ? foto.data : oggi}`}
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

      {/* Linguette come i fogli di Excel: Oggi + fotografie salvate */}
      <div className="flex flex-wrap items-center gap-1 text-[12px]">
        <button
          type="button"
          onClick={() => setFotoId(null)}
          className={`rounded-t-lg border border-b-0 border-border px-3 py-1 ${!foto ? "bg-card font-semibold text-foreground" : "bg-muted/40 text-muted-foreground hover:bg-muted"}`}
        >
          {t("cpr.oggi")}
        </button>
        {db.fotografie.map((f) => (
          <span
            key={f.id}
            className={`group inline-flex items-center gap-1 rounded-t-lg border border-b-0 border-border px-3 py-1 ${foto?.id === f.id ? "bg-card font-semibold text-foreground" : "bg-muted/40 text-muted-foreground hover:bg-muted"}`}
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
                      {r.documenti > 1 && (
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

// --- il foglio ---------------------------------------------------------------
const cel = "border border-border/70 px-2 py-1 align-middle";
const celN = `${cel} text-right tabular-nums whitespace-nowrap`;
const celH = `${cel} bg-muted/40 text-left font-semibold text-foreground`;
const celT = `${celN} bg-muted/30 font-semibold`;
const celL = `${cel} text-[11px] text-muted-foreground`;

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
  const n = Math.max(entrate.length, uscite.length);
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
    if (!u) return <td className={celN} />;
    if (u.secchio === col)
      return (
        <td className={`${celN} text-foreground`}>
          <Importo c={u} readOnly={readOnly} onImporto={onImporto} t={t} />
        </td>
      );
    if (readOnly) return <td className={celN} />;
    return (
      <td className={`${celN} p-0`}>
        <button
          type="button"
          onClick={() => onSposta(u, col)}
          title={t("cpr.spostaQui")}
          className="h-full w-full px-2 py-1 text-right text-transparent hover:bg-primary/10 hover:text-muted-foreground"
        >
          ·
        </button>
      </td>
    );
  };

  return (
    <div className="overflow-x-auto rounded-xl border border-border bg-card shadow-[var(--shadow-card)]">
      <table className="w-full min-w-[1080px] border-collapse text-[13px]">
        <colgroup>
          <col className="w-[115px]" />
          <col className="w-[190px]" />
          <col className="w-[115px]" />
          <col className="w-[115px]" />
          <col className="w-[115px]" />
          <col className="w-[115px]" />
          <col className="w-[230px]" />
          <col />
        </colgroup>
        <thead>
          <tr>
            <th className={celH}>{t("cpr.colEntrate")}</th>
            <th className={celH}>{t("cpr.colNote")}</th>
            <th className={celH}>{t("cpr.colSaldo")}</th>
            <th className={celH}>{t("cpr.colUscite")}</th>
            <th className={celH}>{etichettaProssimi(limite)}</th>
            <th className={celH}>{t("cpr.colStandby")}</th>
            <th className={celH}>{t("cpr.colNote")}</th>
            <th className={celH}>{t("cpr.colNote2")}</th>
          </tr>
        </thead>
        <tbody>
          {/* riga 2: totale entrate e saldo */}
          <tr>
            <td className={`${celN} font-semibold text-foreground`}>{fmt(totali.entrate)}</td>
            <td className={cel} />
            <td className={`${celN} font-semibold text-foreground`}>{fmt(saldo)}</td>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
          </tr>
          {righe.map((i) => {
            const e = entrate[i];
            const u = uscite[i];
            return (
              <tr key={i} className="group/r hover:bg-muted/20">
                <td
                  className={`${celN} ${e?.attesa ? "italic text-muted-foreground" : "text-foreground"}`}
                >
                  {e && <Importo c={e} readOnly={readOnly} onImporto={onImporto} t={t} />}
                </td>
                <td
                  className={`${cel} ${e?.attesa ? "italic text-muted-foreground" : "text-foreground"}`}
                >
                  {e && (
                    <span className="flex items-center gap-1">
                      <span className="truncate" title={e.dettaglio}>
                        {e.nome}
                      </span>
                      <span className="shrink-0 text-[10px] not-italic text-muted-foreground">
                        {e.attesa ? t("cpr.attesa") : ddmm(e.data)}
                      </span>
                      {!readOnly && <Togli onClick={() => onTogli(e)} t={t} />}
                    </span>
                  )}
                </td>
                <td className={cel} />
                {cellaUscita(u, "ora")}
                {cellaUscita(u, "prossimi")}
                {cellaUscita(u, "standby")}
                <td className={`${cel} text-foreground`}>
                  {u && (
                    <span className="flex items-center gap-1">
                      <span className="truncate" title={u.dettaglio}>
                        {u.nome}
                      </span>
                      {u.documenti > 1 && (
                        <span className="shrink-0 text-[10px] text-muted-foreground">
                          ×{u.documenti}
                        </span>
                      )}
                      {u.tipo === "passiva" && u.certa && (
                        <span className="shrink-0 rounded bg-secondary px-1 text-[9px] uppercase text-muted-foreground">
                          auto
                        </span>
                      )}
                      {(u.tipo === "stipendi" || u.tipo === "fisso" || u.tipo === "fiscale") && (
                        <span className="shrink-0 truncate text-[10px] text-muted-foreground">
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
                        <span className="pl-2 text-[11px] font-semibold text-muted-foreground">
                          CDC
                        </span>
                      )}
                      <Nota c={u} readOnly={readOnly} onNota={onNota} />
                      {!readOnly && (
                        <span className="pr-1">
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
            <tr>
              {nuova?.entrata ? (
                <>
                  <td className={`${celN} p-1`}>
                    <input
                      autoFocus
                      value={nuova.importo}
                      onChange={(e) => setNuova({ ...nuova, importo: e.target.value })}
                      onKeyDown={(e) => e.key === "Enter" && conferma()}
                      placeholder={t("cpr.importo")}
                      className={`${inputCls} w-full text-right`}
                    />
                  </td>
                  <td className={`${cel} p-1`}>
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
                  <td className={`${cel} p-1`} />
                  <td className={`${cel} p-1`}>
                    <Aggiungi onClick={() => apri(true)} t={t} />
                  </td>
                </>
              )}
              <td className={cel} />
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
                  <td className={cel} />
                  <td className={cel} />
                  <td className={cel} />
                  <td className={`${cel} p-1`}>
                    <Aggiungi onClick={() => apri(false)} t={t} />
                  </td>
                  <td className={cel} />
                </>
              )}
            </tr>
          )}
          {/* totali come nel foglio */}
          <tr>
            <td className={celT}>{fmt(totali.entrateStandby)}</td>
            <td className={cel} />
            <td className={cel} />
            <td className={celT}>{fmt(totali.usciteOra)}</td>
            <td className={celT}>{fmt(totali.usciteProssimi)}</td>
            <td className={celT}>{fmt(totali.usciteStandby)}</td>
            <td className={celT}>{fmt(totali.totaleDaPagare)}</td>
            <td className={`${celT} ${totali.cassaConStandby < 0 ? "text-status-absent" : ""}`}>
              {fmt(totali.cassaConStandby)}
            </td>
          </tr>
          <tr>
            <td className={celL}>{t("cpr.lblStandbyIncasso")}</td>
            <td className={cel} />
            <td className={cel} />
            <td className={celL}>{t("cpr.lblDaPagare")}</td>
            <td className={celL}>
              {t("cpr.lblDaPagare")} {breve}
            </td>
            <td className={celL}>{t("cpr.lblStandbyDaPagare")}</td>
            <td className={celL}>{t("cpr.lblTotaleDaPagare")}</td>
            <td className={celL}>{t("cpr.lblCassa")}</td>
          </tr>
          <tr>
            <td className={cel} colSpan={8} />
          </tr>
          <tr>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
            <td className={`${celT} ${totali.saldoDopoOra < 0 ? "text-status-absent" : ""}`}>
              {fmt(totali.saldoDopoOra)}
            </td>
            <td className={`${celT} ${totali.saldoDopoProssimi < 0 ? "text-status-absent" : ""}`}>
              {fmt(totali.saldoDopoProssimi)}
            </td>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
          </tr>
          <tr>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
            <td className={celL} colSpan={2}>
              {t("cpr.lblSaldoDopo")}
            </td>
            <td className={cel} />
            <td className={cel} />
            <td className={cel} />
          </tr>
        </tbody>
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
  if (readOnly) return <span className="px-2 py-1 text-foreground">{c.nota}</span>;
  return (
    <input
      value={v ?? c.nota}
      onChange={(e) => setV(e.target.value)}
      onBlur={() => {
        if (v != null && v.trim() !== c.nota) onNota(c, v.trim());
        setV(null);
      }}
      onKeyDown={(e) => e.key === "Enter" && (e.target as HTMLInputElement).blur()}
      className="w-full bg-transparent px-2 py-1 text-[13px] text-foreground focus:bg-background focus:outline-none"
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
      className="text-[12px] text-muted-foreground underline-offset-2 hover:underline"
    >
      {t("cpr.aggiungi")}
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
