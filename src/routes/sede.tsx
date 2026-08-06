// DR Portal — Vista di sede del PREPOSTO (es. capo appalto), in SOLA LETTURA.
// Mostra i turni del giorno e le ore del periodo dei soli dipendenti della
// propria sede: nessuna modifica, nessuna approvazione. Il perimetro è imposto
// dal server sulla sede del record SharePoint, non su quanto dice il client.
import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { CalendarSearch, Loader2, Lock, Users, Hourglass } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import { spGetSedeGiorno, spGetSedeOre } from "@/lib/sharepoint.functions";
import type { ResocontoGiornoRiga, RendicontoRiga } from "@/lib/sharepoint.server";
import { lunediDellaSettimana, ymd } from "@/lib/rendiconto-logic";
import { useLang } from "@/lib/i18n";
import { formatOra } from "@/lib/mock-data";

export const Route = createFileRoute("/sede")({
  head: () => ({ meta: [{ title: "La mia sede — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: SedePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function SedePage() {
  const { t } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [tab, setTab] = useState<"giorno" | "ore">("giorno");
  const [data, setData] = useState(() => new Date().toISOString().slice(0, 10));
  const [sede, setSede] = useState("");
  const [righe, setRighe] = useState<ResocontoGiornoRiga[] | null>(null);
  // Periodo per le ore: settimana corrente (lun–dom).
  const [from, setFrom] = useState(() =>
    lunediDellaSettimana(new Date().toISOString().slice(0, 10)),
  );
  const [to, setTo] = useState(() => {
    const d = new Date(`${lunediDellaSettimana(new Date().toISOString().slice(0, 10))}T00:00:00`);
    d.setDate(d.getDate() + 6);
    return ymd(d);
  });
  const [ore, setOre] = useState<RendicontoRiga[] | null>(null);
  // Filtro sede: ha senso solo per chi vede TUTTE le sedi (operatore/admin);
  // per il preposto il server restituisce già solo la sua.
  const [sedeF, setSedeF] = useState("tutte");

  const abilitato =
    session != null &&
    (Boolean(session.preposto) || session.operatore || session.ruolo === "amministratore_sistema");

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
  }, []);

  const caricaGiorno = () => {
    setRighe(null);
    spGetSedeGiorno({ data: { data } })
      .then((r) => {
        setSede(r.sede);
        setRighe(r.righe as ResocontoGiornoRiga[]);
      })
      .catch((err) => {
        setRighe([]);
        toast.error(t("sede.err"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  };
  const caricaOre = () => {
    setOre(null);
    spGetSedeOre({ data: { from, to } })
      .then((r) => setOre(r as RendicontoRiga[]))
      .catch((err) => {
        setOre([]);
        toast.error(t("sede.err"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  };

  useEffect(() => {
    if (!abilitato) return;
    if (tab === "giorno") caricaGiorno();
    else caricaOre();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [abilitato, tab, data, from, to]);

  const righeVis = useMemo(
    () => (righe ?? []).filter((r) => sedeF === "tutte" || r.sede === sedeF),
    [righe, sedeF],
  );
  const oreVis = useMemo(
    () => (ore ?? []).filter((r) => sedeF === "tutte" || r.sede === sedeF),
    [ore, sedeF],
  );
  const sediViste = useMemo(() => {
    const set = new Set<string>();
    for (const r of righe ?? []) if (r.sede) set.add(r.sede);
    for (const r of ore ?? []) if (r.sede) set.add(r.sede);
    return [...set].sort();
  }, [righe, ore]);
  const totaleOre = useMemo(() => oreVis.reduce((s, r) => s + r.oreLavorate, 0), [oreVis]);

  if (session && !abilitato) {
    return (
      <AppShell title={t("sede.title")}>
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">{t("sede.negato")}</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">{t("sede.negatoMsg")}</p>
          </div>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell
      title={t("sede.title")}
      subtitle={sede ? `${t("sede.sede")}: ${sede}` : t("sede.subtitle")}
    >
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setTab("giorno")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "giorno" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <CalendarSearch className="h-4 w-4" /> {t("sede.tabGiorno")}
        </button>
        <button
          type="button"
          onClick={() => setTab("ore")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "ore" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <Hourglass className="h-4 w-4" /> {t("sede.tabOre")}
        </button>
      </div>

      <p className="mb-4 text-[13px] text-muted-foreground">{t("sede.soloLettura")}</p>

      {tab === "giorno" ? (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-4 flex flex-wrap items-end gap-3">
            <div className="w-48">
              <label className="text-xs text-muted-foreground">{t("sede.data")}</label>
              <input
                type="date"
                value={data}
                onChange={(e) => setData(e.target.value)}
                className={inputCls}
              />
            </div>
            {sede === "tutte" && (
              <div className="w-48">
                <label className="text-xs text-muted-foreground">{t("common.site")}</label>
                <select
                  value={sedeF}
                  onChange={(e) => setSedeF(e.target.value)}
                  className={inputCls}
                >
                  <option value="tutte">{t("common.allF")}</option>
                  {sediViste.map((x) => (
                    <option key={x} value={x}>
                      {x}
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>
          {righe == null ? (
            <div className="py-8 text-center">
              <Loader2 className="h-5 w-5 animate-spin inline-block text-muted-foreground" />
            </div>
          ) : righeVis.length === 0 ? (
            <p className="py-6 text-center text-sm text-muted-foreground">{t("sede.vuoto")}</p>
          ) : (
            <ul className="space-y-2">
              {righeVis.map((r) => (
                <li key={r.dipendenteId} className="rounded-xl border border-border p-3">
                  <div className="flex flex-wrap items-center gap-x-3 gap-y-2">
                    <span className="inline-flex items-center gap-2 text-sm font-medium text-foreground min-w-48">
                      <Users className="h-4 w-4 text-muted-foreground" />
                      {r.nomeCompleto}
                      <span className="text-xs text-muted-foreground">{r.codice}</span>
                    </span>
                    {r.malattia && (
                      <span className="rounded-full bg-status-break/15 px-2 py-0.5 text-[11px] font-medium text-status-break">
                        {t("gt.malattiaBadge")}
                      </span>
                    )}
                    {r.ferie && (
                      <span className="rounded-full bg-primary/10 px-2 py-0.5 text-[11px] font-medium text-primary">
                        {t("gt.ferieBadge")}
                      </span>
                    )}
                    {r.senzaTimbrature && !r.malattia && !r.ferie ? (
                      <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
                        {t("sede.nessuna")}
                      </span>
                    ) : (
                      <span className="flex flex-wrap gap-1.5">
                        {r.eventi.map((e) => (
                          <span
                            key={e.id}
                            className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground"
                          >
                            {t(`evento.${e.evento}`)} {formatOra(e.dataOra)}
                          </span>
                        ))}
                      </span>
                    )}
                    {r.anomalie.map((a) => (
                      <span
                        key={a}
                        className="rounded-full bg-status-absent/15 px-2 py-0.5 text-[11px] font-medium text-status-absent"
                      >
                        {t(`anomalia.${a}`)}
                      </span>
                    ))}
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-4 flex flex-wrap items-end gap-3">
            <div className="w-44">
              <label className="text-xs text-muted-foreground">{t("common.from")}</label>
              <input
                type="date"
                value={from}
                onChange={(e) => setFrom(e.target.value)}
                className={inputCls}
              />
            </div>
            <div className="w-44">
              <label className="text-xs text-muted-foreground">{t("common.to")}</label>
              <input
                type="date"
                value={to}
                onChange={(e) => setTo(e.target.value)}
                className={inputCls}
              />
            </div>
            {sede === "tutte" && (
              <div className="w-44">
                <label className="text-xs text-muted-foreground">{t("common.site")}</label>
                <select
                  value={sedeF}
                  onChange={(e) => setSedeF(e.target.value)}
                  className={inputCls}
                >
                  <option value="tutte">{t("common.allF")}</option>
                  {sediViste.map((x) => (
                    <option key={x} value={x}>
                      {x}
                    </option>
                  ))}
                </select>
              </div>
            )}
            <div className="ml-auto rounded-xl border border-border bg-secondary/40 px-4 py-2 text-right">
              <div className="text-[11px] text-muted-foreground">{t("sede.totaleOre")}</div>
              <div className="text-lg font-semibold tabular-nums text-foreground">
                {totaleOre.toLocaleString("it-IT", { maximumFractionDigits: 2 })}
              </div>
            </div>
          </div>
          {ore == null ? (
            <div className="py-8 text-center">
              <Loader2 className="h-5 w-5 animate-spin inline-block text-muted-foreground" />
            </div>
          ) : oreVis.length === 0 ? (
            <p className="py-6 text-center text-sm text-muted-foreground">{t("sede.vuoto")}</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">{t("sede.dipendente")}</th>
                    <th className="py-2 pr-3 text-right">{t("sede.oreLavorate")}</th>
                    <th className="py-2 pr-3 text-right">{t("sede.straordinario")}</th>
                    <th className="py-2 pr-3 text-right">{t("sede.giorniAperti")}</th>
                  </tr>
                </thead>
                <tbody>
                  {oreVis.map((r) => (
                    <tr key={r.dipendenteId} className="border-b border-border/50">
                      <td className="py-1.5 pr-3">{r.nomeCompleto}</td>
                      <td className="py-1.5 pr-3 text-right tabular-nums">
                        {r.oreLavorate.toLocaleString("it-IT", { maximumFractionDigits: 2 })}
                      </td>
                      <td className="py-1.5 pr-3 text-right tabular-nums">
                        {r.straordinarioCalcolato
                          ? r.straordinarioCalcolato.toLocaleString("it-IT", {
                              maximumFractionDigits: 2,
                            })
                          : "—"}
                      </td>
                      <td className="py-1.5 pr-3 text-right tabular-nums">
                        {r.giorniNonChiusi || "—"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </AppShell>
  );
}
