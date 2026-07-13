import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { BarChart3, Lock, AlertTriangle } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import { spGetRendiconto } from "@/lib/sharepoint.functions";
import type { RendicontoRiga } from "@/lib/sharepoint.server";
import { SEDI, type SedeId } from "@/lib/mock-data";

export const Route = createFileRoute("/report")({
  head: () => ({ meta: [{ title: "Rendiconto — DR Portal" }] }),
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
  component: RendicontoPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function sedeNome(id: string): string {
  return SEDI.find((s) => s.id === id)?.nome ?? id;
}
function h(n: number): string {
  return n > 0 ? `${n} h` : "—";
}
function gg(n: number): string {
  return n > 0 ? `${n}` : "—";
}

function currentPeriodo(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function RendicontoPage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [periodo, setPeriodo] = useState<string>(currentPeriodo());
  const [righe, setRighe] = useState<RendicontoRiga[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [sedeF, setSedeF] = useState<SedeId | "tutte">("tutte");
  const [dipF, setDipF] = useState("");

  const canView =
    session != null &&
    (session.operatore ||
      session.autorizza ||
      session.ruolo === "amministratore_sistema" ||
      session.ruolo === "responsabile");

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
  }, []);

  useEffect(() => {
    if (!canView) return;
    const anno = Number(periodo.slice(0, 4));
    const mese = Number(periodo.slice(5, 7));
    if (!anno || !mese) return;
    setLoading(true);
    setRighe(null);
    spGetRendiconto({ data: { anno, mese } })
      .then((l) => setRighe(l as RendicontoRiga[]))
      .catch((err) => {
        setRighe([]);
        toast.error("Errore rendiconto", {
          description: err instanceof Error ? err.message : String(err),
        });
      })
      .finally(() => setLoading(false));
  }, [periodo, canView]);

  const filtrate = useMemo(() => {
    return (righe ?? []).filter((r) => {
      if (sedeF !== "tutte" && r.sede !== sedeF) return false;
      if (dipF && r.dipendenteId !== dipF) return false;
      return true;
    });
  }, [righe, sedeF, dipF]);

  if (session && !canView) {
    return (
      <AppShell title="Rendiconto">
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">Accesso riservato</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">
              Il rendiconto è riservato a operatori, supervisori e amministratori.
            </p>
          </div>
        </div>
      </AppShell>
    );
  }

  return (
    <AppShell title="Rendiconto" subtitle="Ore mensili per dipendente">
      <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
          <BarChart3 className="h-4 w-4 text-primary" /> Rendiconto mensile
        </div>

        {/* Filtri */}
        <div className="grid gap-3 sm:grid-cols-3 mb-4">
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">Mese</label>
            <input
              type="month"
              className={`${inputCls} mt-1`}
              value={periodo}
              onChange={(e) => setPeriodo(e.target.value)}
            />
          </div>
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
              {(righe ?? []).map((r) => (
                <option key={r.dipendenteId} value={r.dipendenteId}>
                  {r.nomeCompleto}
                </option>
              ))}
            </select>
          </div>
        </div>

        {loading || righe === null ? (
          <div className="text-sm text-muted-foreground">Calcolo in corso…</div>
        ) : filtrate.length === 0 ? (
          <div className="text-sm text-muted-foreground">
            Nessun dato per il periodo/filtri selezionati.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                  <th className="py-2 pr-3">Dipendente</th>
                  <th className="py-2 pr-3">Sede</th>
                  <th className="py-2 pr-3 text-right">Ore lav.</th>
                  <th className="py-2 pr-3 text-right">Str. calc.</th>
                  <th className="py-2 pr-3 text-right">Str. autor.</th>
                  <th className="py-2 pr-3 text-right">Permessi</th>
                  <th className="py-2 pr-3 text-right">Ferie</th>
                  <th className="py-2 pr-3 text-right">Malattie</th>
                </tr>
              </thead>
              <tbody>
                {filtrate.map((r) => (
                  <tr key={r.dipendenteId} className="border-b border-border/60">
                    <td className="py-2 pr-3 text-foreground">
                      <span className="inline-flex items-center gap-1.5">
                        {r.nomeCompleto}
                        {r.giorniNonChiusi > 0 && (
                          <span
                            title={`${r.giorniNonChiusi} giorno/i con turno non chiuso: correggere nelle Anomalie`}
                            className="inline-flex items-center text-status-absent"
                          >
                            <AlertTriangle className="h-3.5 w-3.5" />
                          </span>
                        )}
                      </span>
                    </td>
                    <td className="py-2 pr-3 text-muted-foreground">{sedeNome(r.sede)}</td>
                    <td className="py-2 pr-3 text-right tabular-nums text-foreground">
                      {h(r.oreLavorate)}
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums">
                      {h(r.straordinarioCalcolato)}
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums text-muted-foreground">
                      {h(r.straordinarioAutorizzato)}
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums">{h(r.permessiOre)}</td>
                    <td className="py-2 pr-3 text-right tabular-nums">{gg(r.ferieGiorni)}</td>
                    <td className="py-2 pr-3 text-right tabular-nums">{gg(r.malattiaGiorni)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div className="mt-4 text-[12px] text-muted-foreground leading-relaxed">
          <strong>Str. calc.</strong> = straordinario dalle timbrature (ore Lun–Sab oltre il monte
          ore settimanale + tutta la domenica). <strong>Str. autor.</strong> = ore da richieste di
          straordinario approvate, per confronto. Il previsto settimanale è ridotto da
          ferie/malattia ({"OreSettimanali/5"} per giorno) e dai permessi. Le settimane a cavallo di
          due mesi contano nel mese del loro lunedì. ⚠️ = giorni con turno non chiuso (ore non
          conteggiate: correggere nelle Anomalie).
        </div>
      </div>
    </AppShell>
  );
}
