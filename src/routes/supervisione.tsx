import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { ShieldCheck, Lock, CheckCircle2, ClipboardList, CalendarDays, Clock } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetRichieste,
  spGetDipendenti,
  spGetTimbratureManuali,
} from "@/lib/sharepoint.functions";
import type { SpRichiesta, SpDipendente, TimbraturaManualeItem } from "@/lib/sharepoint.server";
import { labelTipo } from "@/lib/mock-data";
import { SEDI, type SedeId } from "@/lib/mock-data";

export const Route = createFileRoute("/supervisione")({
  head: () => ({ meta: [{ title: "Supervisione — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    let hasSession = false;
    try {
      const raw = window.sessionStorage.getItem("dr:currentUser");
      if (raw) hasSession = Boolean((JSON.parse(raw) as { id?: string } | null)?.id);
    } catch {
      hasSession = false;
    }
    if (!hasSession) throw redirect({ to: "/", search: { redirect: location.href } });
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
function sedeNome(id: string): string {
  return SEDI.find((s) => s.id === id)?.nome ?? id;
}

function SupervisionePage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [tab, setTab] = useState<"approvate" | "manuali">("approvate");

  const [approvate, setApprovate] = useState<SpRichiesta[] | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[]>([]);
  const [manuali, setManuali] = useState<TimbraturaManualeItem[] | null>(null);

  // Filtri report approvate
  const [sedeF, setSedeF] = useState<SedeId | "tutte">("tutte");
  const [dipF, setDipF] = useState("");
  const [dal, setDal] = useState("");
  const [al, setAl] = useState("");

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    if (!s.autorizza) return;
    spGetRichieste({ data: { stato: "Approvata" } })
      .then((l) => setApprovate(l as SpRichiesta[]))
      .catch((err) => {
        setApprovate([]);
        toast.error("Errore richieste", {
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
        toast.error("Errore timbrature", {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  }, []);

  const nomeById = useMemo(() => {
    const m = new Map<string, string>();
    for (const d of dipendenti) m.set(d.id, d.nomeCompleto || `${d.cognome} ${d.nome}`);
    return m;
  }, [dipendenti]);

  const filtrate = useMemo(() => {
    return (approvate ?? []).filter((r) => {
      if (sedeF !== "tutte" && r.sedeRichiedente !== sedeNome(sedeF)) return false;
      if (dipF && r.richiedenteId !== dipF) return false;
      const d = r.dataInizio.slice(0, 10);
      if (dal && d < dal) return false;
      if (al && d > al) return false;
      return true;
    });
  }, [approvate, sedeF, dipF, dal, al]);

  if (session && !session.autorizza) {
    return (
      <AppShell title="Supervisione">
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">Accesso riservato</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">
              Questa sezione è riservata ai supervisori.
            </p>
          </div>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell title="Supervisione" subtitle="Approvazioni e timbrature manuali">
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setTab("approvate")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "approvate" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <CheckCircle2 className="h-4 w-4" /> Richieste approvate
        </button>
        <button
          type="button"
          onClick={() => setTab("manuali")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "manuali" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <ClipboardList className="h-4 w-4" /> Timbrature manuali
        </button>
      </div>

      {tab === "approvate" ? (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <ShieldCheck className="h-4 w-4 text-primary" /> Richieste approvate
          </div>

          {/* Filtri */}
          <div className="grid gap-3 sm:grid-cols-4 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Sede</label>
              <select
                className={`${inputCls} mt-1`}
                value={sedeF}
                onChange={(e) => setSedeF(e.target.value as SedeId | "tutte")}
              >
                <option value="tutte">Tutte</option>
                {SEDI.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.nome}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Dipendente
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={dipF}
                onChange={(e) => setDipF(e.target.value)}
              >
                <option value="">Tutti</option>
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
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Dal</label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={dal}
                onChange={(e) => setDal(e.target.value)}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Al</label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={al}
                onChange={(e) => setAl(e.target.value)}
              />
            </div>
          </div>

          {approvate === null ? (
            <div className="text-sm text-muted-foreground">Caricamento…</div>
          ) : filtrate.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessuna richiesta approvata con questi filtri.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">Richiesta</th>
                    <th className="py-2 pr-3">Dipendente</th>
                    <th className="py-2 pr-3">Tipo</th>
                    <th className="py-2 pr-3">Periodo</th>
                    <th className="py-2 pr-3">Sede</th>
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
                        <td className="py-2 pr-3 text-foreground">
                          {nomeById.get(r.richiedenteId) ||
                            r.codiceRichiedente ||
                            `#${r.richiedenteId}`}
                        </td>
                        <td className="py-2 pr-3">
                          {r.tipo}
                          {r.modalita ? ` (${r.modalita})` : ""}
                        </td>
                        <td className="py-2 pr-3 whitespace-nowrap">
                          {periodo}
                          {ore}
                        </td>
                        <td className="py-2 pr-3 text-muted-foreground">
                          {r.sedeRichiedente || "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <div className="mt-3 text-[12px] text-muted-foreground">
                {filtrate.length} richiest{filtrate.length === 1 ? "a" : "e"}.
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <ClipboardList className="h-4 w-4 text-primary" /> Timbrature manuali (ultimi 30 giorni)
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">
            Inserimenti fatti dall'operatore (origine Manuale), per visione.
          </p>

          {manuali === null ? (
            <div className="text-sm text-muted-foreground">Caricamento…</div>
          ) : manuali.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessuna timbratura manuale nel periodo.
            </div>
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
                      {labelTipo(t.evento)} · {fmtDataOra(t.dataOra)}
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
