// DR Portal — Finanza → tab "Estratto conto" (direttore).
// L'estratto conto di un GRUPPO di controparti come lo vede il cliente:
// per società e mese di competenza, con bonifici e compensazioni, accanto
// alla vista Aruba e alla differenza; il semaforo mensile banca/Aruba/
// compensazioni; la previsione degli incassi. Logica in lib/estratto-logic.
import { useEffect, useMemo, useState, type ReactElement } from "react";
import { toast } from "sonner";
import { ChevronDown, ChevronRight, Loader2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import {
  calcolaEstratto,
  etichettaMese,
  parseConfigGruppo,
  serializeConfigGruppo,
  type EstrattoConto,
  type FatturaEC,
} from "@/lib/estratto-logic";
import {
  spGetFatture,
  spGetFlussiCassa,
  spGetGruppiControparti,
  spGetMovimenti,
  spGetTerminiPagamento,
  spUpsertFlussoCassa,
} from "@/lib/sharepoint.functions";
import type { TerminePagamento } from "@/lib/fatture-logic";
import type {
  FlussoCassaRiga,
  GruppoControparti,
  SpFattura,
  SpMovimento,
} from "@/lib/sharepoint.server";

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtData = (iso: string) => (iso ? iso.slice(0, 10).split("-").reverse().join("/") : "—");
const oggiISO = () => new Date().toISOString().slice(0, 10);

const inputCls =
  "rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";
const th = "px-2 py-1.5 text-left text-[11px] uppercase tracking-wider text-muted-foreground";
const thR = `${th} text-right`;
const td = "px-2 py-1.5 text-sm text-foreground";
const tdR = `${td} text-right tabular-nums`;

export function EstrattoContoTab() {
  const { t } = useLang();
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [gruppi, setGruppi] = useState<GruppoControparti[] | null>(null);
  const [flussi, setFlussi] = useState<FlussoCassaRiga[]>([]);
  const [gruppoId, setGruppoId] = useState("");
  const [dataRif, setDataRif] = useState(oggiISO());
  const [aperte, setAperte] = useState<Set<string>>(new Set());
  const [passiveAperte, setPassiveAperte] = useState(false);
  const [salvando, setSalvando] = useState(false);

  useEffect(() => {
    const err = (e: unknown) =>
      toast.error(t("ft.errLoad"), { description: e instanceof Error ? e.message : String(e) });
    spGetFatture({ data: { direzione: "Emessa" } })
      .then((l) => setFattureEm(l as SpFattura[]))
      .catch((e) => {
        setFattureEm([]);
        err(e);
      });
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetMovimenti()
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    spGetGruppiControparti()
      .then((l) => setGruppi(l as GruppoControparti[]))
      .catch(() => setGruppi([]));
    spGetFlussiCassa()
      .then((l) => setFlussi(l as FlussoCassaRiga[]))
      .catch(() => setFlussi([]));
  }, [t]);

  // Solo i gruppi veri: le viste salvate del Resoconto vivono sulla stessa
  // lista come JSON e non sono gruppi.
  const gruppiVeri = useMemo(
    () => (gruppi ?? []).filter((g) => g.membri && !g.membri.trim().startsWith("{")),
    [gruppi],
  );
  useEffect(() => {
    if (gruppoId || !gruppiVeri.length) return;
    const univex = gruppiVeri.find((g) => /univex/i.test(g.nome));
    setGruppoId((univex ?? gruppiVeri[0]).id);
  }, [gruppiVeri, gruppoId]);
  const gruppo = gruppiVeri.find((g) => g.id === gruppoId) ?? null;

  const config = useMemo(() => {
    if (!gruppo) return {};
    const riga = flussi.find((r) => r.genere === "gruppo" && r.nome === gruppo.nome);
    return parseConfigGruppo(riga?.note);
  }, [flussi, gruppo]);

  const ec: EstrattoConto | null = useMemo(() => {
    if (!gruppo || !fattureEm || !fattureRic || !movimenti) return null;
    return calcolaEstratto({
      gruppo: { nome: gruppo.nome, membri: gruppo.membri },
      attive: fattureEm,
      passive: fattureRic,
      movimenti: movimenti.map((m) => ({
        dataContabile: m.dataContabile,
        importo: m.importo,
        cliente: m.cliente,
      })),
      termini,
      oggi: dataRif,
      assorbe: config.assorbe,
    });
  }, [gruppo, fattureEm, fattureRic, movimenti, termini, dataRif, config.assorbe]);

  const nomeSocieta = (chiave: string) =>
    ec?.societa.find((s) => s.chiave === chiave)?.nome ?? chiave;

  const salvaAssorbe = async (chiave: string) => {
    if (!gruppo) return;
    setSalvando(true);
    try {
      await spUpsertFlussoCassa({
        data: {
          nome: gruppo.nome,
          genere: "gruppo",
          importo: 0,
          note: serializeConfigGruppo({ assorbe: chiave || undefined }),
        },
      });
      const l = (await spGetFlussiCassa()) as FlussoCassaRiga[];
      setFlussi(l);
      toast.success(t("ec.salvato"));
    } catch (e) {
      toast.error(t("common.error"), { description: e instanceof Error ? e.message : String(e) });
    } finally {
      setSalvando(false);
    }
  };

  const toggle = (k: string) =>
    setAperte((s) => {
      const n = new Set(s);
      if (n.has(k)) n.delete(k);
      else n.add(k);
      return n;
    });

  const caricamento = !fattureEm || !fattureRic || !movimenti || !gruppi;

  const righeCsv = () => {
    if (!ec) return [];
    const out: (string | number)[][] = [];
    for (const r of ec.righe) {
      out.push([
        nomeSocieta(r.societa),
        etichettaMese(r.competenza),
        "",
        "",
        r.fatturato,
        r.bonifici,
        r.compensazioni,
        r.apertoCliente,
        r.scadutoCliente,
        r.apertoAruba,
        r.apertoAruba - r.apertoCliente,
      ]);
      for (const f of r.fatture) {
        out.push([
          nomeSocieta(r.societa),
          etichettaMese(r.competenza),
          f.numero,
          fmtData(f.data),
          f.netto,
          f.bonifici,
          f.compensazioni,
          f.residuoCliente,
          f.scadenza <= dataRif ? f.residuoCliente : 0,
          f.residuoAruba,
          f.residuoAruba - f.residuoCliente,
        ]);
      }
    }
    return out;
  };

  const Delta = ({ v }: { v: number }) => (
    <span
      className={Math.abs(v) < 0.5 ? "text-muted-foreground" : "font-semibold text-status-absent"}
    >
      {Math.abs(v) < 0.5 ? "—" : fmt(v)}
    </span>
  );

  return (
    <div className="space-y-5">
      <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-3">
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">
              {t("ec.gruppo")}
            </label>
            <select
              className={`${inputCls} mt-1 block`}
              value={gruppoId}
              onChange={(e) => {
                setGruppoId(e.target.value);
                setAperte(new Set());
              }}
            >
              {gruppiVeri.map((g) => (
                <option key={g.id} value={g.id}>
                  {g.nome}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">
              {t("ec.assorbe")}
            </label>
            <select
              className={`${inputCls} mt-1 block`}
              value={ec?.assorbe ?? ""}
              disabled={!ec || salvando}
              onChange={(e) => void salvaAssorbe(e.target.value)}
            >
              {(ec?.societa ?? [])
                .filter((s) => s.ciFattura)
                .map((s) => (
                  <option key={s.chiave} value={s.chiave}>
                    {s.nome}
                  </option>
                ))}
            </select>
          </div>
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">
              {t("ec.situazioneAl")}
            </label>
            <input
              type="date"
              className={`${inputCls} mt-1 block`}
              value={dataRif}
              onChange={(e) => setDataRif(e.target.value || oggiISO())}
            />
          </div>
          <div className="ml-auto">
            <EsportaStampa
              nomeFile={`estratto-conto-${(gruppo?.nome ?? "gruppo").toLowerCase().replace(/\s+/g, "-")}-${dataRif}`}
              testata={[
                t("ec.societa"),
                t("ec.competenza"),
                t("ec.fatturaNumero"),
                t("common.date"),
                t("ec.fatturato"),
                t("ec.bonifici"),
                t("ec.compensazioni"),
                t("ec.apertoCliente"),
                t("ec.scaduto"),
                t("ec.apertoAruba"),
                t("ec.delta"),
              ]}
              righe={righeCsv}
              disabled={!ec}
            />
          </div>
        </div>
        <p className="mt-3 text-[13px] text-muted-foreground">{t("ec.intro")}</p>
      </div>

      {caricamento && (
        <div className="flex items-center gap-2 text-sm text-muted-foreground">
          <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
        </div>
      )}
      {!caricamento && !gruppo && (
        <div className="text-sm text-muted-foreground">{t("ec.nessunGruppo")}</div>
      )}

      {ec && (
        <>
          <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
            <Kpi label={t("ec.apertoCliente")} value={ec.totaleGruppo.apertoCliente} />
            <Kpi label={t("ec.scaduto")} value={ec.totaleGruppo.scadutoCliente} rosso />
            <Kpi label={t("ec.apertoAruba")} value={ec.totaleGruppo.apertoAruba} />
            <Kpi label={t("ec.delta")} value={ec.totaleGruppo.delta} delta />
            <Kpi label={t("ec.noiDobbiamo")} value={ec.passivoAperto} />
            <Kpi label={t("ec.nonAllocati")} value={ec.bonificiNonAllocati} delta />
          </div>

          {/* Per società e competenza */}
          <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <div className="mb-3 text-[15px] font-semibold text-foreground">
              {t("ec.perSocieta")}
            </div>
            <div className="overflow-x-auto">
              <table className="w-full min-w-[900px] border-collapse">
                <thead>
                  <tr className="border-b border-border">
                    <th className={th}>{t("ec.competenza")}</th>
                    <th className={thR}>{t("ec.fatturato")}</th>
                    <th className={thR}>{t("ec.bonifici")}</th>
                    <th className={thR}>{t("ec.compensazioni")}</th>
                    <th className={thR}>{t("ec.apertoCliente")}</th>
                    <th className={thR}>{t("ec.scaduto")}</th>
                    <th className={thR}>{t("ec.apertoAruba")}</th>
                    <th className={thR}>{t("ec.delta")}</th>
                  </tr>
                </thead>
                <tbody>
                  {ec.totali.map((tot) => {
                    const righe = ec.righe.filter((r) => r.societa === tot.societa);
                    return (
                      <RigheSocieta
                        key={tot.societa}
                        nome={nomeSocieta(tot.societa)}
                        tot={tot}
                        righe={righe}
                        aperte={aperte}
                        toggle={toggle}
                        dataRif={dataRif}
                        Delta={Delta}
                      />
                    );
                  })}
                  <tr className="border-t-2 border-border font-semibold">
                    <td className={td}>{t("ec.totaleGruppo")}</td>
                    <td className={tdR}>{fmt(ec.totaleGruppo.fatturato)}</td>
                    <td className={tdR}>{fmt(ec.totaleGruppo.bonifici)}</td>
                    <td className={tdR}>{fmt(ec.totaleGruppo.compensazioni)}</td>
                    <td className={tdR}>{fmt(ec.totaleGruppo.apertoCliente)}</td>
                    <td className={`${tdR} text-status-absent`}>
                      {fmt(ec.totaleGruppo.scadutoCliente)}
                    </td>
                    <td className={tdR}>{fmt(ec.totaleGruppo.apertoAruba)}</td>
                    <td className={tdR}>
                      <Delta v={ec.totaleGruppo.delta} />
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <p className="mt-2 text-[12px] text-muted-foreground">{t("ec.deltaHelp")}</p>
          </div>

          <div className="grid gap-5 lg:grid-cols-2">
            {/* Semaforo */}
            <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
              <div className="mb-1 text-[15px] font-semibold text-foreground">
                {t("ec.semaforo")}
              </div>
              <p className="mb-3 text-[12px] text-muted-foreground">{t("ec.semaforoHelp")}</p>
              <table className="w-full border-collapse">
                <thead>
                  <tr className="border-b border-border">
                    <th className={th}>{t("ec.mese")}</th>
                    <th className={thR}>{t("ec.arubaReg")}</th>
                    <th className={thR}>{t("ec.banca")}</th>
                    <th className={thR}>{t("ec.compensazioni")}</th>
                    <th className={thR}>{t("ec.delta")}</th>
                  </tr>
                </thead>
                <tbody>
                  {ec.semaforo.map((s) => (
                    <tr key={s.mese} className="border-b border-border/60">
                      <td className={td}>
                        <span
                          className={`mr-2 inline-block h-2.5 w-2.5 rounded-full ${
                            s.stato === "ok"
                              ? "bg-status-present"
                              : s.stato === "attenzione"
                                ? "bg-amber-500"
                                : "bg-status-absent"
                          }`}
                        />
                        {etichettaMese(s.mese)}
                      </td>
                      <td className={tdR}>{fmt(s.arubaIncassato)}</td>
                      <td className={tdR}>{fmt(s.banca)}</td>
                      <td className={tdR}>{fmt(s.compensazioni)}</td>
                      <td className={tdR}>
                        <Delta v={s.delta} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Previsione */}
            <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
              <div className="mb-1 text-[15px] font-semibold text-foreground">
                {t("ec.previsione")}
              </div>
              <p className="mb-3 text-[12px] text-muted-foreground">{t("ec.previsioneHelp")}</p>
              <table className="w-full border-collapse">
                <thead>
                  <tr className="border-b border-border">
                    <th className={th}>{t("ec.mese")}</th>
                    <th className={thR}>{t("ec.daIncassare")}</th>
                    <th className={thR}>{t("ec.daTrattenere")}</th>
                    <th className={thR}>{t("ec.nettoAtteso")}</th>
                  </tr>
                </thead>
                <tbody>
                  {ec.previsione.map((p) => (
                    <tr key={p.mese} className="border-b border-border/60">
                      <td
                        className={`${td} ${p.mese === "scaduto" ? "font-semibold text-status-absent" : ""}`}
                      >
                        {p.mese === "scaduto" ? t("ec.scaduto") : etichettaMese(p.mese)}
                      </td>
                      <td className={tdR}>{fmt(p.daIncassare)}</td>
                      <td className={tdR}>{p.daTrattenere ? `−${fmt(p.daTrattenere)}` : "—"}</td>
                      <td className={`${tdR} font-semibold`}>{fmt(p.netto)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Passive del gruppo */}
          <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <button
              type="button"
              onClick={() => setPassiveAperte((v) => !v)}
              className="flex items-center gap-2 text-[15px] font-semibold text-foreground"
            >
              {passiveAperte ? (
                <ChevronDown className="h-4 w-4" />
              ) : (
                <ChevronRight className="h-4 w-4" />
              )}
              {t("ec.passive")} ({ec.passive.length})
            </button>
            {passiveAperte && (
              <table className="mt-3 w-full border-collapse">
                <thead>
                  <tr className="border-b border-border">
                    <th className={th}>{t("ec.fatturaNumero")}</th>
                    <th className={th}>{t("common.date")}</th>
                    <th className={th}>{t("ec.societa")}</th>
                    <th className={th}>{t("ec.competenza")}</th>
                    <th className={thR}>{t("ec.totale")}</th>
                    <th className={th}>{t("ec.compensataSu")}</th>
                  </tr>
                </thead>
                <tbody>
                  {ec.passive.map((p) => (
                    <tr key={p.nomeFile} className="border-b border-border/60">
                      <td className={td}>{p.numero}</td>
                      <td className={td}>{fmtData(p.data)}</td>
                      <td className={td}>{nomeSocieta(p.societa)}</td>
                      <td className={td}>{etichettaMese(p.competenza)}</td>
                      <td className={tdR}>{fmt(p.netto)}</td>
                      <td className={td}>
                        {p.nonAllocata > 0.005
                          ? `${t("ec.daCompensare")} ${fmt(p.nonAllocata)}`
                          : p.compensataSu === "bonifico"
                            ? t("ec.pagataBonifico")
                            : nomeSocieta(p.compensataSu ?? "")}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function Kpi({
  label,
  value,
  rosso,
  delta,
}: {
  label: string;
  value: number;
  rosso?: boolean;
  delta?: boolean;
}) {
  const zero = Math.abs(value) < 0.5;
  return (
    <div className="rounded-xl border border-border bg-card px-4 py-3 shadow-[var(--shadow-card)]">
      <div className="text-[11px] uppercase tracking-wider text-muted-foreground">{label}</div>
      <div
        className={`mt-1 text-lg font-semibold tabular-nums ${
          rosso && !zero
            ? "text-status-absent"
            : delta && !zero
              ? "text-status-absent"
              : "text-foreground"
        }`}
      >
        {delta && zero ? "—" : fmt(value)}
      </div>
    </div>
  );
}

function RigheSocieta({
  nome,
  tot,
  righe,
  aperte,
  toggle,
  dataRif,
  Delta,
}: {
  nome: string;
  tot: EstrattoConto["totali"][number];
  righe: EstrattoConto["righe"];
  aperte: Set<string>;
  toggle: (k: string) => void;
  dataRif: string;
  Delta: (p: { v: number }) => ReactElement;
}) {
  const { t } = useLang();
  return (
    <>
      <tr className="border-t border-border bg-secondary/40 font-semibold">
        <td className={td}>{nome}</td>
        <td className={tdR}>{fmt(tot.fatturato)}</td>
        <td className={tdR}>{fmt(tot.bonifici)}</td>
        <td className={tdR}>{fmt(tot.compensazioni)}</td>
        <td className={tdR}>{fmt(tot.apertoCliente)}</td>
        <td className={`${tdR} text-status-absent`}>{fmt(tot.scadutoCliente)}</td>
        <td className={tdR}>{fmt(tot.apertoAruba)}</td>
        <td className={tdR}>
          <Delta v={tot.delta} />
        </td>
      </tr>
      {righe.map((r) => {
        const k = `${r.societa}|${r.competenza}`;
        const aperta = aperte.has(k);
        return (
          <RigaMese
            key={k}
            r={r}
            aperta={aperta}
            onToggle={() => toggle(k)}
            dataRif={dataRif}
            Delta={Delta}
          />
        );
      })}
    </>
  );
}

function RigaMese({
  r,
  aperta,
  onToggle,
  dataRif,
  Delta,
}: {
  r: EstrattoConto["righe"][number];
  aperta: boolean;
  onToggle: () => void;
  dataRif: string;
  Delta: (p: { v: number }) => ReactElement;
}) {
  const { t } = useLang();
  return (
    <>
      <tr
        className="cursor-pointer border-b border-border/60 hover:bg-secondary/30"
        onClick={onToggle}
      >
        <td className={`${td} pl-6`}>
          <span className="mr-1 inline-block w-3 text-muted-foreground">{aperta ? "▾" : "▸"}</span>
          {etichettaMese(r.competenza)}
          <span className="ml-2 text-[11px] text-muted-foreground">
            {r.fatture.length} {t("ec.fattureN")}
          </span>
        </td>
        <td className={tdR}>{fmt(r.fatturato)}</td>
        <td className={tdR}>{r.bonifici ? fmt(r.bonifici) : "—"}</td>
        <td className={tdR}>{r.compensazioni ? fmt(r.compensazioni) : "—"}</td>
        <td className={tdR}>{fmt(r.apertoCliente)}</td>
        <td className={`${tdR} ${r.scadutoCliente ? "text-status-absent" : ""}`}>
          {r.scadutoCliente ? fmt(r.scadutoCliente) : "—"}
        </td>
        <td className={tdR}>{fmt(r.apertoAruba)}</td>
        <td className={tdR}>
          <Delta v={r.apertoAruba - r.apertoCliente} />
        </td>
      </tr>
      {aperta &&
        r.fatture.map((f: FatturaEC) => (
          <tr key={f.nomeFile} className="border-b border-border/40 bg-background/60 text-[13px]">
            <td className={`${td} pl-12 text-muted-foreground`}>
              {f.numero} · {fmtData(f.data)} · {t("ec.scadenza")} {fmtData(f.scadenza)}
              {f.scadenza <= dataRif && f.residuoCliente > 0.5 ? " · " + t("ec.scadutaBadge") : ""}
            </td>
            <td className={`${tdR} text-muted-foreground`}>{fmt(f.netto)}</td>
            <td className={`${tdR} text-muted-foreground`}>{f.bonifici ? fmt(f.bonifici) : "—"}</td>
            <td className={`${tdR} text-muted-foreground`}>
              {f.compensazioni ? fmt(f.compensazioni) : "—"}
            </td>
            <td className={`${tdR} text-muted-foreground`}>{fmt(f.residuoCliente)}</td>
            <td className={`${tdR} text-muted-foreground`}>
              {f.scadenza <= dataRif && f.residuoCliente > 0.5 ? fmt(f.residuoCliente) : "—"}
            </td>
            <td className={`${tdR} text-muted-foreground`}>{fmt(f.residuoAruba)}</td>
            <td className={`${tdR}`}>
              <Delta v={f.residuoAruba - f.residuoCliente} />
            </td>
          </tr>
        ))}
    </>
  );
}
