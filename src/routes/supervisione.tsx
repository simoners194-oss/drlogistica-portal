import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import {
  ShieldCheck,
  Lock,
  CheckCircle2,
  ClipboardList,
  CalendarDays,
  Clock,
  Receipt,
  Download,
} from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetRichieste,
  spGetDipendenti,
  spGetTimbratureManuali,
} from "@/lib/sharepoint.functions";
import type { SpRichiesta, SpDipendente, TimbraturaManualeItem } from "@/lib/sharepoint.server";
import { type SedeId } from "@/lib/mock-data";

export const Route = createFileRoute("/supervisione")({
  head: () => ({ meta: [{ title: "Supervisione — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: SupervisionePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}
function fmtDataOra(iso: string): string {
  const d = fmtData(iso);
  const t = iso.slice(11, 16);
  return t ? `${d} ${t}` : d;
}
// La sede è già il suo nome reale: nessuna mappatura id→nome.
function sedeNome(id: string): string {
  return id;
}

// Elenco sedi distinte presenti nei dati (richieste + dipendenti), ordinato.
function sediDistinte(nomi: (string | undefined)[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const n of nomi) {
    const s = (n ?? "").trim();
    if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
      seen.add(s.toLowerCase());
      out.push(s);
    }
  }
  return out.sort((a, b) => a.localeCompare(b));
}

const STATO_BADGE: Record<string, string> = {
  Approvata: "bg-status-present/15 text-status-present",
  Respinta: "bg-status-absent/15 text-status-absent",
};
function StatoBadge({ stato, label }: { stato: string; label?: string }) {
  return (
    <span
      className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${STATO_BADGE[stato] ?? "bg-muted text-muted-foreground"}`}
    >
      {label ?? stato}
    </span>
  );
}

function SupervisionePage() {
  const { t, tStato, tVal } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [tab, setTab] = useState<"approvate" | "rimborsi" | "manuali">("approvate");

  const [decise, setDecise] = useState<SpRichiesta[] | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[]>([]);
  const [manuali, setManuali] = useState<TimbraturaManualeItem[] | null>(null);

  // Filtri report richieste decise (approvate + rifiutate)
  const [sedeF, setSedeF] = useState<SedeId | "tutte">("tutte");
  const [dipF, setDipF] = useState("");
  const [statoF, setStatoF] = useState<"tutte" | "Approvata" | "Respinta">("tutte");
  const [dal, setDal] = useState("");
  const [al, setAl] = useState("");

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    const puoVedere =
      s.autorizza ||
      s.operatore ||
      s.ruolo === "amministratore_sistema" ||
      s.ruolo === "responsabile";
    if (!puoVedere) return;
    Promise.all([
      spGetRichieste({ data: { stato: "Approvata" } }),
      spGetRichieste({ data: { stato: "Respinta" } }),
    ])
      .then(([a, r]) => setDecise([...(a as SpRichiesta[]), ...(r as SpRichiesta[])]))
      .catch((err) => {
        setDecise([]);
        toast.error(t("sup.errRequests"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
    spGetDipendenti()
      .then((l) => setDipendenti(l as SpDipendente[]))
      .catch(() => {});
    spGetTimbratureManuali({ data: { giorni: 30 } })
      .then((l) => setManuali(l as TimbraturaManualeItem[]))
      .catch((err) => {
        setManuali([]);
        toast.error(t("sup.errEntries"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  }, []);

  const nomeById = useMemo(() => {
    const m = new Map<string, string>();
    for (const d of dipendenti) m.set(d.id, d.nomeCompleto || `${d.cognome} ${d.nome}`);
    return m;
  }, [dipendenti]);

  const sediOptions = useMemo(
    () =>
      sediDistinte([
        ...(decise ?? []).map((r) => r.sedeRichiedente),
        ...dipendenti.map((d) => d.sede),
      ]),
    [decise, dipendenti],
  );

  const filtrate = useMemo(() => {
    return (decise ?? []).filter((r) => {
      if (statoF !== "tutte" && r.stato !== statoF) return false;
      if (sedeF !== "tutte" && r.sedeRichiedente !== sedeNome(sedeF)) return false;
      if (dipF && r.richiedenteId !== dipF) return false;
      const d = r.dataInizio.slice(0, 10);
      if (dal && d < dal) return false;
      if (al && d > al) return false;
      return true;
    });
  }, [decise, statoF, sedeF, dipF, dal, al]);

  const rimborsi = useMemo(() => filtrate.filter((r) => r.tipo === "Rimborso spese"), [filtrate]);
  // Totale sui soli rimborsi APPROVATI (i respinti non concorrono alla spesa).
  const totaleImporto = useMemo(
    () => rimborsi.reduce((s, r) => s + (r.stato === "Approvata" ? (r.importo ?? 0) : 0), 0),
    [rimborsi],
  );

  const puoVedere =
    session != null &&
    (session.autorizza ||
      session.operatore ||
      session.ruolo === "amministratore_sistema" ||
      session.ruolo === "responsabile");

  if (session && !puoVedere) {
    return (
      <AppShell title={t("sup.title")}>
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">{t("common.restricted")}</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">{t("sup.restrictedMsg")}</p>
          </div>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell title={t("sup.title")} subtitle={t("sup.subtitle")}>
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setTab("approvate")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "approvate" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <CheckCircle2 className="h-4 w-4" /> {t("sup.tabDecise")}
        </button>
        <button
          type="button"
          onClick={() => setTab("manuali")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "manuali" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <ClipboardList className="h-4 w-4" /> {t("sup.tabManuali")}
        </button>
        <button
          type="button"
          onClick={() => setTab("rimborsi")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "rimborsi" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <Receipt className="h-4 w-4" /> {t("sup.tabRimborsi")}
        </button>
      </div>

      {tab === "approvate" ? (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
              <ShieldCheck className="h-4 w-4 text-primary" /> {t("sup.deciseTitle")}
            </div>
            {filtrate.length > 0 && (
              <button
                type="button"
                onClick={() =>
                  esportaCsvFile(
                    "richieste-decise",
                    ["Richiesta", "Stato", "Dipendente", "Tipo", "Dal", "Al", "Sede", "Documento"],
                    filtrate.map((r) => [
                      r.title || r.id,
                      r.stato,
                      nomeById.get(r.richiedenteId) || r.codiceRichiedente,
                      r.tipo + (r.modalita ? ` (${r.modalita})` : ""),
                      fmtData(r.dataInizio),
                      fmtData(r.dataFine || r.dataInizio),
                      r.sedeRichiedente,
                      r.giustificativo ?? "",
                    ]),
                  )
                }
                className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors"
              >
                <Download className="h-4 w-4" /> {t("common.exportCsv")}
              </button>
            )}
          </div>

          {/* Filtri */}
          <div className="grid gap-3 sm:grid-cols-5 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Stato
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={statoF}
                onChange={(e) => setStatoF(e.target.value as "tutte" | "Approvata" | "Respinta")}
              >
                <option value="tutte">{t("common.allF")}</option>
                <option value="Approvata">{t("sup.approvate")}</option>
                <option value="Respinta">{t("sup.rifiutate")}</option>
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.site")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={sedeF}
                onChange={(e) => setSedeF(e.target.value as SedeId | "tutte")}
              >
                <option value="tutte">{t("common.allF")}</option>
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.employee")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={dipF}
                onChange={(e) => setDipF(e.target.value)}
              >
                <option value="">{t("common.all")}</option>
                {[...dipendenti]
                  .sort((a, b) => `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`))
                  .map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.cognome} {d.nome}
                    </option>
                  ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.from")}
              </label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={dal}
                onChange={(e) => setDal(e.target.value)}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.to")}
              </label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={al}
                onChange={(e) => setAl(e.target.value)}
              />
            </div>
          </div>

          {decise === null ? (
            <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
          ) : filtrate.length === 0 ? (
            <div className="text-sm text-muted-foreground">{t("sup.deciseNone")}</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">{t("sup.colRichiesta")}</th>
                    <th className="py-2 pr-3">{t("common.status")}</th>
                    <th className="py-2 pr-3">{t("common.employee")}</th>
                    <th className="py-2 pr-3">{t("common.type")}</th>
                    <th className="py-2 pr-3">{t("sup.colPeriodo")}</th>
                    <th className="py-2 pr-3">{t("common.site")}</th>
                    <th className="py-2 pr-3">{t("common.doc")}</th>
                  </tr>
                </thead>
                <tbody>
                  {filtrate.map((r) => {
                    const periodo =
                      r.dataFine && r.dataFine.slice(0, 10) !== r.dataInizio.slice(0, 10)
                        ? `${fmtData(r.dataInizio)} → ${fmtData(r.dataFine)}`
                        : fmtData(r.dataInizio);
                    const ore = r.oraInizio && r.oraFine ? ` · ${r.oraInizio}–${r.oraFine}` : "";
                    return (
                      <tr key={r.id} className="border-b border-border/60">
                        <td className="py-2 pr-3 text-muted-foreground">{r.title || `#${r.id}`}</td>
                        <td className="py-2 pr-3">
                          <StatoBadge stato={r.stato} label={tStato(r.stato)} />
                        </td>
                        <td className="py-2 pr-3 text-foreground">
                          {nomeById.get(r.richiedenteId) ||
                            r.codiceRichiedente ||
                            `#${r.richiedenteId}`}
                        </td>
                        <td className="py-2 pr-3">
                          {tVal("tipoR", r.tipo)}
                          {r.modalita ? ` (${tVal("mod", r.modalita)})` : ""}
                        </td>
                        <td className="py-2 pr-3 whitespace-nowrap">
                          {periodo}
                          {ore}
                        </td>
                        <td className="py-2 pr-3 text-muted-foreground">
                          {r.sedeRichiedente || "—"}
                        </td>
                        <td className="py-2 pr-3">
                          {r.giustificativo ? (
                            <a
                              href={r.giustificativo}
                              target="_blank"
                              rel="noreferrer"
                              className="text-primary underline"
                            >
                              {t("common.open")}
                            </a>
                          ) : (
                            "—"
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <div className="mt-3 text-[12px] text-muted-foreground">
                {filtrate.length} {filtrate.length === 1 ? t("sup.request") : t("sup.requests")}.
              </div>
            </div>
          )}
        </div>
      ) : tab === "rimborsi" ? (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
              <Receipt className="h-4 w-4 text-primary" /> {t("sup.rimborsiTitle")}
            </div>
            {rimborsi.length > 0 && (
              <button
                type="button"
                onClick={() =>
                  esportaCsvFile(
                    "rimborsi",
                    ["Dipendente", "Stato", "Sede", "Data", "Tipologia", "Importo", "Documento"],
                    rimborsi.map((r) => [
                      nomeById.get(r.richiedenteId) || r.codiceRichiedente,
                      r.stato,
                      r.sedeRichiedente,
                      fmtData(r.dataInizio),
                      r.tipoAcquisto ?? "",
                      r.importo ?? "",
                      r.giustificativo ?? "",
                    ]),
                  )
                }
                className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors"
              >
                <Download className="h-4 w-4" /> {t("common.exportCsv")}
              </button>
            )}
          </div>

          <div className="grid gap-3 sm:grid-cols-4 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.site")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={sedeF}
                onChange={(e) => setSedeF(e.target.value as SedeId | "tutte")}
              >
                <option value="tutte">{t("common.allF")}</option>
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.employee")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={dipF}
                onChange={(e) => setDipF(e.target.value)}
              >
                <option value="">{t("common.all")}</option>
                {[...dipendenti]
                  .sort((a, b) => `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`))
                  .map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.cognome} {d.nome}
                    </option>
                  ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.from")}
              </label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={dal}
                onChange={(e) => setDal(e.target.value)}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.to")}
              </label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={al}
                onChange={(e) => setAl(e.target.value)}
              />
            </div>
          </div>

          {decise === null ? (
            <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
          ) : rimborsi.length === 0 ? (
            <div className="text-sm text-muted-foreground">Nessun rimborso con questi filtri.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">{t("common.employee")}</th>
                    <th className="py-2 pr-3">{t("common.status")}</th>
                    <th className="py-2 pr-3">{t("common.site")}</th>
                    <th className="py-2 pr-3">{t("common.date")}</th>
                    <th className="py-2 pr-3">{t("sup.colTipologia")}</th>
                    <th className="py-2 pr-3 text-right">{t("common.amount")}</th>
                    <th className="py-2 pr-3">{t("common.doc")}</th>
                  </tr>
                </thead>
                <tbody>
                  {rimborsi.map((r) => (
                    <tr key={r.id} className="border-b border-border/60">
                      <td className="py-2 pr-3 text-foreground">
                        {nomeById.get(r.richiedenteId) ||
                          r.codiceRichiedente ||
                          `#${r.richiedenteId}`}
                      </td>
                      <td className="py-2 pr-3">
                        <StatoBadge stato={r.stato} label={tStato(r.stato)} />
                      </td>
                      <td className="py-2 pr-3 text-muted-foreground">
                        {r.sedeRichiedente || "—"}
                      </td>
                      <td className="py-2 pr-3 whitespace-nowrap">{fmtData(r.dataInizio)}</td>
                      <td className="py-2 pr-3">{r.tipoAcquisto || "—"}</td>
                      <td className="py-2 pr-3 text-right tabular-nums">
                        € {(r.importo ?? 0).toFixed(2)}
                      </td>
                      <td className="py-2 pr-3">
                        {r.giustificativo ? (
                          <a
                            href={r.giustificativo}
                            target="_blank"
                            rel="noreferrer"
                            className="text-primary underline"
                          >
                            {t("common.open")}
                          </a>
                        ) : (
                          "—"
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="border-t border-border font-semibold">
                    <td className="py-2 pr-3" colSpan={5}>
                      {t("sup.totApproved")} (
                      {rimborsi.filter((r) => r.stato === "Approvata").length})
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums">
                      € {totaleImporto.toFixed(2)}
                    </td>
                    <td />
                  </tr>
                </tfoot>
              </table>
            </div>
          )}
        </div>
      ) : (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <ClipboardList className="h-4 w-4 text-primary" /> {t("sup.manualiTitle")}
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">{t("sup.manualiDesc")}</p>

          {manuali === null ? (
            <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
          ) : manuali.length === 0 ? (
            <div className="text-sm text-muted-foreground">{t("sup.manualiNone")}</div>
          ) : (
            <ul className="space-y-2">
              {manuali.map((t) => (
                <li
                  key={t.id}
                  className="flex items-center justify-between gap-3 rounded-xl border border-border bg-background p-3"
                >
                  <div className="min-w-0">
                    <div className="font-medium text-foreground truncate">{t.nomeCompleto}</div>
                    <div className="mt-0.5 flex items-center gap-2 text-[13px] text-muted-foreground">
                      <Clock className="h-3.5 w-3.5" />
                      {tVal("evento", t.evento)} · {fmtDataOra(t.dataOra)}
                      {t.sede ? ` · ${sedeNome(t.sede)}` : ""}
                    </div>
                    {t.note && (
                      <div className="mt-0.5 text-[12px] text-muted-foreground/80 italic">
                        “{t.note}”
                      </div>
                    )}
                  </div>
                  <CalendarDays className="h-4 w-4 shrink-0 text-muted-foreground/50" />
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
    </AppShell>
  );
}
