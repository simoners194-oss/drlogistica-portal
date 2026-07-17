import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import {
  ShoppingCart,
  Send,
  Loader2,
  CheckCircle2,
  XCircle,
  Inbox,
  Lock,
  Download,
} from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetVoci,
  spGetAcquisti,
  spCreateAcquisto,
  spDecideAcquisto,
} from "@/lib/sharepoint.functions";
import type { SpVoce, SpAcquisto } from "@/lib/sharepoint.server";
import { isSedeStorica } from "@/lib/richieste-logic";
import { useLang } from "@/lib/i18n";

export const Route = createFileRoute("/procurement")({
  head: () => ({ meta: [{ title: "Procurement — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: ProcurementPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

const STATO_STYLE: Record<string, string> = {
  Inviata: "bg-primary/10 text-primary",
  Approvata: "bg-status-present/15 text-status-present",
  Respinta: "bg-status-absent/15 text-status-absent",
};

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

function esportaCsv(righe: SpAcquisto[]): void {
  const header = [
    "Richiesta",
    "Stato",
    "Richiedente",
    "Sede",
    "Macro",
    "Dettaglio",
    "Descrizione",
    "Importo",
    "Data richiesta",
    "Note decisione",
  ];
  const esc = (v: string | number): string => {
    const s = String(v ?? "");
    return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const body = righe.map((r) =>
    [
      r.title || r.id,
      r.stato,
      r.codiceRichiedente,
      r.sedeRichiedente,
      r.macro,
      r.dettaglio,
      r.descrizione,
      r.importo ?? "",
      fmtData(r.dataRichiesta || r.createdAt),
      r.noteDecisione ?? "",
    ]
      .map(esc)
      .join(";"),
  );
  const csv = [header.join(";"), ...body].join("\r\n");
  const blob = new Blob([String.fromCharCode(0xfeff) + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `richieste-acquisto.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function ProcurementPage() {
  const { t, tStato } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [voci, setVoci] = useState<SpVoce[]>([]);
  const [mie, setMie] = useState<SpAcquisto[] | null>(null);
  const [tutte, setTutte] = useState<SpAcquisto[] | null>(null);
  const [view, setView] = useState<"mie" | "coda">("mie");

  // Form
  const [macro, setMacro] = useState("");
  const [dettaglio, setDettaglio] = useState("");
  const [descrizione, setDescrizione] = useState("");
  const [importo, setImporto] = useState("");
  const [submitting, setSubmitting] = useState(false);

  // Decisioni
  const [decidingId, setDecidingId] = useState<string | null>(null);
  const [rejectingId, setRejectingId] = useState<string | null>(null);
  const [rejectNote, setRejectNote] = useState("");

  const isApprovatore =
    session != null && (session.autorizza || session.ruolo === "amministratore_sistema");
  const puoRichiedere =
    session != null &&
    (isSedeStorica(String(session.sede)) || session.ruolo === "amministratore_sistema");

  const loadMie = () => {
    spGetAcquisti({ data: { mie: true } })
      .then((l) => setMie(l as SpAcquisto[]))
      .catch((err) => {
        setMie([]);
        toast.error(t("sup.errRequests"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  };
  const loadTutte = () => {
    spGetAcquisti({ data: {} })
      .then((l) => setTutte(l as SpAcquisto[]))
      .catch(() => setTutte([]));
  };

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    spGetVoci({ data: { ambito: "Acquisto" } })
      .then((l) => setVoci(l as SpVoce[]))
      .catch(() => {});
    loadMie();
    if (s.autorizza || s.ruolo === "amministratore_sistema") loadTutte();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const macros = useMemo(() => [...new Set(voci.map((v) => v.macro))], [voci]);
  const dettagli = useMemo(
    () => voci.filter((v) => v.macro === macro && v.dettaglio).map((v) => v.dettaglio),
    [voci, macro],
  );

  const pending = useMemo(() => (tutte ?? []).filter((r) => r.stato === "Inviata"), [tutte]);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!macro) {
      toast.error(t("proc.selectItem"));
      return;
    }
    if (dettagli.length > 0 && !dettaglio) {
      toast.error(t("proc.selectItemDetail"));
      return;
    }
    if (!descrizione.trim()) {
      toast.error(t("proc.needDesc"));
      return;
    }
    setSubmitting(true);
    try {
      await spCreateAcquisto({
        data: {
          macro,
          dettaglio,
          descrizione: descrizione.trim(),
          importo: importo ? Number(importo.replace(",", ".")) : undefined,
        },
      });
      toast.success(t("proc.submitted"));
      setMacro("");
      setDettaglio("");
      setDescrizione("");
      setImporto("");
      loadMie();
      if (isApprovatore) loadTutte();
    } catch (err) {
      toast.error(t("proc.submitErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  };

  const decide = async (r: SpAcquisto, decisione: "Approvata" | "Respinta", nota?: string) => {
    setDecidingId(r.id);
    try {
      await spDecideAcquisto({
        data: { acquistoId: r.id, decisione, noteDecisione: nota?.trim() || undefined },
      });
      toast.success(decisione === "Approvata" ? t("rich.approvedToast") : t("rich.rejectedToast"));
      setRejectingId(null);
      setRejectNote("");
      loadTutte();
      loadMie();
    } catch (err) {
      toast.error(t("proc.decideErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setDecidingId(null);
    }
  };

  if (session && !puoRichiedere && !isApprovatore) {
    return (
      <AppShell title={t("proc.title")}>
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">{t("proc.notActive")}</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">{t("proc.notActiveMsg")}</p>
          </div>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell title={t("proc.title")} subtitle={t("proc.subtitle")}>
      {isApprovatore && (
        <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
          <button
            type="button"
            onClick={() => setView("mie")}
            className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${view === "mie" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
          >
            {t("proc.myRequests")}
          </button>
          <button
            type="button"
            onClick={() => setView("coda")}
            className={`rounded-lg px-3 py-1.5 font-medium transition-colors inline-flex items-center gap-2 ${view === "coda" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
          >
            {t("proc.toApprove")}
            {pending.length > 0 && (
              <span
                className={`rounded-full px-1.5 text-[11px] ${view === "coda" ? "bg-primary-foreground/20" : "bg-primary/10 text-primary"}`}
              >
                {pending.length}
              </span>
            )}
          </button>
        </div>
      )}

      {view === "coda" && isApprovatore ? (
        <div className="space-y-6">
          <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
              <Inbox className="h-4 w-4 text-primary" /> {t("proc.queueTitle")}
            </div>
            {tutte === null ? (
              <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
            ) : pending.length === 0 ? (
              <div className="text-sm text-muted-foreground">{t("proc.nonePending")}</div>
            ) : (
              <ul className="space-y-3">
                {pending.map((r) => (
                  <li key={r.id} className="rounded-xl border border-border bg-background p-4">
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="font-semibold text-foreground">{r.macro}</span>
                          {r.dettaglio && (
                            <span className="text-[13px] text-muted-foreground">
                              › {r.dettaglio}
                            </span>
                          )}
                          {r.importo != null && (
                            <span className="font-medium text-foreground tabular-nums">
                              € {r.importo.toFixed(2)}
                            </span>
                          )}
                        </div>
                        <div className="mt-1 text-[13px] text-foreground/80">{r.descrizione}</div>
                        <div className="mt-1 text-[11px] text-muted-foreground">
                          {r.codiceRichiedente} · {r.sedeRichiedente} ·{" "}
                          {fmtData(r.dataRichiesta || r.createdAt)} · {r.title || `#${r.id}`}
                        </div>
                      </div>
                      <div className="flex shrink-0 flex-col gap-2">
                        <Button
                          type="button"
                          size="sm"
                          className="bg-status-present hover:bg-status-present/90"
                          disabled={decidingId === r.id}
                          onClick={() => decide(r, "Approvata")}
                        >
                          <CheckCircle2 className="h-4 w-4" /> {t("common.approve")}
                        </Button>
                        <Button
                          type="button"
                          size="sm"
                          variant="ghost"
                          className="text-status-absent hover:text-status-absent"
                          disabled={decidingId === r.id}
                          onClick={() => setRejectingId((cur) => (cur === r.id ? null : r.id))}
                        >
                          <XCircle className="h-4 w-4" /> {t("common.reject")}
                        </Button>
                      </div>
                    </div>
                    {rejectingId === r.id && (
                      <div className="mt-3 rounded-lg bg-status-absent/5 p-3">
                        <label className="text-xs uppercase tracking-wider text-muted-foreground">
                          {t("common.rejectReason")}{" "}
                          <span className="normal-case">{t("common.required")}</span>
                        </label>
                        <textarea
                          className={`${inputCls} mt-1 min-h-[60px] resize-y`}
                          value={rejectNote}
                          onChange={(e) => setRejectNote(e.target.value)}
                          autoFocus
                        />
                        <div className="mt-2 flex justify-end gap-2">
                          <Button
                            type="button"
                            size="sm"
                            variant="ghost"
                            onClick={() => {
                              setRejectingId(null);
                              setRejectNote("");
                            }}
                          >
                            {t("common.cancel")}
                          </Button>
                          <Button
                            type="button"
                            size="sm"
                            className="bg-status-absent hover:bg-status-absent/90"
                            disabled={decidingId === r.id || !rejectNote.trim()}
                            onClick={() => decide(r, "Respinta", rejectNote)}
                          >
                            {t("common.confirmReject")}
                          </Button>
                        </div>
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            )}
          </div>

          <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
              <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
                <ShoppingCart className="h-4 w-4 text-primary" /> {t("proc.allRequests")}
              </div>
              {(tutte ?? []).length > 0 && (
                <button
                  type="button"
                  onClick={() => esportaCsv(tutte ?? [])}
                  className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors"
                >
                  <Download className="h-4 w-4" /> {t("common.exportCsv")}
                </button>
              )}
            </div>
            {tutte === null ? (
              <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
            ) : (tutte ?? []).length === 0 ? (
              <div className="text-sm text-muted-foreground">{t("proc.none")}</div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                      <th className="py-2 pr-3">{t("sup.colRichiesta")}</th>
                      <th className="py-2 pr-3">{t("common.status")}</th>
                      <th className="py-2 pr-3">{t("common.employee")}</th>
                      <th className="py-2 pr-3">{t("proc.item")}</th>
                      <th className="py-2 pr-3 text-right">{t("common.amount")}</th>
                      <th className="py-2 pr-3">{t("common.date")}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(tutte ?? []).map((r) => (
                      <tr key={r.id} className="border-b border-border/60">
                        <td className="py-2 pr-3 text-muted-foreground">{r.title || `#${r.id}`}</td>
                        <td className="py-2 pr-3">
                          <span
                            className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${STATO_STYLE[r.stato] ?? "bg-muted text-muted-foreground"}`}
                          >
                            {tStato(r.stato)}
                          </span>
                        </td>
                        <td className="py-2 pr-3 text-foreground">{r.codiceRichiedente}</td>
                        <td className="py-2 pr-3">
                          {r.macro}
                          {r.dettaglio ? ` › ${r.dettaglio}` : ""}
                        </td>
                        <td className="py-2 pr-3 text-right tabular-nums">
                          {r.importo != null ? `€ ${r.importo.toFixed(2)}` : "—"}
                        </td>
                        <td className="py-2 pr-3 whitespace-nowrap">
                          {fmtData(r.dataRichiesta || r.createdAt)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      ) : (
        <div className="space-y-6">
          {puoRichiedere && (
            <form
              onSubmit={submit}
              className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]"
            >
              <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
                <ShoppingCart className="h-4 w-4 text-primary" /> {t("proc.newTitle")}
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("proc.item")}
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={macro}
                    onChange={(e) => {
                      setMacro(e.target.value);
                      setDettaglio("");
                    }}
                  >
                    <option value="">{t("common.select")}</option>
                    {macros.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                  {macros.length === 0 && (
                    <p className="mt-1 text-[11px] text-status-absent">{t("proc.noVoci")}</p>
                  )}
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("proc.detail")}
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={dettaglio}
                    onChange={(e) => setDettaglio(e.target.value)}
                    disabled={dettagli.length === 0}
                  >
                    <option value="">
                      {dettagli.length === 0 ? t("proc.noDetail") : t("common.select")}
                    </option>
                    {dettagli.map((d) => (
                      <option key={d} value={d}>
                        {d}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="sm:col-span-2">
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("proc.desc")}
                  </label>
                  <textarea
                    className={`${inputCls} mt-1 min-h-[70px] resize-y`}
                    value={descrizione}
                    onChange={(e) => setDescrizione(e.target.value)}
                    placeholder={t("proc.descPh")}
                  />
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("proc.estAmount")}
                  </label>
                  <input
                    className={`${inputCls} mt-1`}
                    inputMode="decimal"
                    value={importo}
                    onChange={(e) => setImporto(e.target.value)}
                    placeholder="0,00"
                  />
                </div>
              </div>
              <div className="mt-4 flex justify-end">
                <Button type="submit" disabled={submitting}>
                  {submitting ? (
                    <Loader2 className="h-4 w-4 mr-2 animate-spin" />
                  ) : (
                    <Send className="h-4 w-4 mr-2" />
                  )}
                  {t("proc.submit")}
                </Button>
              </div>
            </form>
          )}

          <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
              <ShoppingCart className="h-4 w-4 text-primary" /> {t("proc.myRequests")}
            </div>
            {mie === null ? (
              <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
            ) : mie.length === 0 ? (
              <div className="text-sm text-muted-foreground">{t("proc.mineNone")}</div>
            ) : (
              <ul className="space-y-2">
                {mie.map((r) => (
                  <li
                    key={r.id}
                    className="flex items-start justify-between gap-3 rounded-xl border border-border bg-background p-3"
                  >
                    <div className="min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span
                          className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${STATO_STYLE[r.stato] ?? "bg-muted text-muted-foreground"}`}
                        >
                          {tStato(r.stato)}
                        </span>
                        <span className="font-medium text-foreground">{r.macro}</span>
                        {r.dettaglio && (
                          <span className="text-[13px] text-muted-foreground">› {r.dettaglio}</span>
                        )}
                        {r.importo != null && (
                          <span className="tabular-nums text-[13px]">€ {r.importo.toFixed(2)}</span>
                        )}
                      </div>
                      <div className="mt-0.5 text-[13px] text-foreground/80">{r.descrizione}</div>
                      <div className="mt-0.5 text-[11px] text-muted-foreground">
                        {fmtData(r.dataRichiesta || r.createdAt)} · {r.title || `#${r.id}`}
                      </div>
                      {r.noteDecisione && r.stato === "Respinta" && (
                        <div className="mt-1 text-[12px] text-status-absent italic">
                          “{r.noteDecisione}”
                        </div>
                      )}
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
      )}
    </AppShell>
  );
}
