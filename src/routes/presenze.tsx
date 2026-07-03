import { createFileRoute, redirect, useNavigate } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { PresenzeSkeleton } from "@/components/skeletons/PresenzeSkeleton";
import { LogIn, Coffee, PlayCircle, LogOut, Clock, Timer, ListChecks, Hourglass, TrendingUp } from "lucide-react";
import { formatOra, labelTipo, type Dipendente, type Timbratura } from "@/lib/mock-data";
import { dataService, oreLavorateOggi, displayStato, DISPLAY_DOT, DISPLAY_LABEL } from "@/lib/data-service";
import {
  computeOreOggi,
  formatDurata,
  isTransitionAllowed,
  lastEvento,
  reasonNotAllowed,
  type EventoTimbratura,
} from "@/lib/presenze-logic";

export const Route = createFileRoute("/presenze")({
  head: () => ({ meta: [{ title: "Modulo Presenze — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    // La sessione vive in sessionStorage (client-only): su SSR/prerender
    // non c'è nulla da controllare, la guardia scatta solo nel browser.
    if (typeof window === "undefined") return;
    let hasSession = false;
    try {
      const raw = window.sessionStorage.getItem("dr:currentUser");
      if (raw) {
        const parsed = JSON.parse(raw) as { id?: string } | null;
        hasSession = Boolean(parsed?.id);
      }
    } catch {
      hasSession = false;
    }
    if (!hasSession) {
      throw redirect({
        to: "/",
        search: { redirect: location.href },
      });
    }
  },
  component: PresenzePage,
});

function PresenzePage() {
  const navigate = useNavigate();
  const [now, setNow] = useState(new Date());
  const [me, setMe] = useState<Dipendente | undefined>(undefined);
  const [errore, setErrore] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let currentId: string | null = null;
    try {
      const raw = sessionStorage.getItem("dr:currentUser");
      if (raw) currentId = (JSON.parse(raw) as { id?: string }).id ?? null;
    } catch {
      /* ignore */
    }
    if (!currentId) {
      toast.error("Sessione scaduta. Effettua di nuovo l'accesso.");
      navigate({ to: "/" });
      return;
    }
    dataService
      .getDipendente(currentId)
      .then((d) => {
        if (d) setMe(d);
        else setErrore("Dipendente non trovato su SharePoint.");
      })
      .catch((err) => setErrore(err instanceof Error ? err.message : String(err)));
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, [navigate]);

  const timbra = async (tipo: Timbratura["tipo"]) => {
    if (!me || busy) return;
    const last = lastEvento(me.eventiOggi ?? []);
    if (!isTransitionAllowed(tipo, last)) {
      const motivo = reasonNotAllowed(tipo, last) ?? "Timbratura non consentita in questo momento.";
      toast.error("Timbratura non consentita in questo momento.", { description: motivo });
      return;
    }
    setBusy(true);
    try {
      const updated = await dataService.timbra(me.id, tipo);
      setMe(updated);
      toast.success(`Timbratura registrata: ${labelTipo(tipo)}`, {
        description: `Ore ${formatOra(updated.ultimaTimbratura?.ora)} · Stato: ${DISPLAY_LABEL[displayStato(updated)]}`,
      });
    } catch (err) {
      toast.error("Timbratura non salvata", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setBusy(false);
    }
  };

  if (errore) {
    return (
      <AppShell title="Le mie presenze">
        <div className="text-sm text-status-absent">{errore}</div>
      </AppShell>
    );
  }

  if (!me) {
    return (
      <AppShell title="Le mie presenze" subtitle="Caricamento in corso…">
        <PresenzeSkeleton />
      </AppShell>
    );
  }

  const ds = displayStato(me);
  const eventiOggi = me.eventiOggi ?? [];
  const last = lastEvento(eventiOggi);
  const ore = computeOreOggi(eventiOggi, now);
  const azioni: {
    tipo: EventoTimbratura;
    label: string;
    Icon: typeof LogIn;
    enabled: boolean;
    reason: string | null;
    tone: "primary" | "warn" | "ok" | "danger";
  }[] = (
    [
      { tipo: "entrata", label: "Entrata", Icon: LogIn, tone: "primary" },
      { tipo: "inizio-pausa", label: "Inizio pausa", Icon: Coffee, tone: "warn" },
      { tipo: "fine-pausa", label: "Fine pausa", Icon: PlayCircle, tone: "ok" },
      { tipo: "uscita", label: "Uscita", Icon: LogOut, tone: "danger" },
    ] as const
  ).map((a) => ({
    ...a,
    enabled: isTransitionAllowed(a.tipo, last),
    reason: reasonNotAllowed(a.tipo, last),
  }));

  return (
    <AppShell title="Le mie presenze" subtitle={`${me.nome} ${me.cognome} · ${me.ruolo}`}>
      <div className="grid gap-4 md:gap-5 lg:grid-cols-3">
        <div
          className="lg:col-span-2 rounded-2xl p-5 sm:p-6 text-primary-foreground shadow-[var(--shadow-elegant)]"
          style={{ background: "var(--gradient-hero)" }}
        >
          <div className="text-[13px] sm:text-sm text-white/80 capitalize">
            {now.toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
          </div>
          <div className="mt-1 text-[3.5rem] leading-none sm:text-6xl font-semibold tabular-nums tracking-tight">
            {now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
          </div>
          <div className="mt-4 flex flex-wrap items-center gap-x-5 gap-y-2 text-[13px] sm:text-sm">
            <span className="inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1">
              <span className={`h-2.5 w-2.5 rounded-full ${DISPLAY_DOT[ds]} ring-2 ring-white/30`} />
              <span className="text-white/90">Stato: <strong>{DISPLAY_LABEL[ds]}</strong></span>
            </span>
            <span className="inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1 text-white/90">
              <Timer className="h-4 w-4" /> Ore oggi: <strong className="tabular-nums">{oreLavorateOggi(me, now)}</strong>
            </span>
          </div>
        </div>

        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-muted-foreground text-[13px] sm:text-sm">
            <Clock className="h-4 w-4" /> Ultima timbratura
          </div>
          {me.ultimaTimbratura ? (
            <div className="mt-3">
              <div className="text-lg font-semibold text-foreground">{labelTipo(me.ultimaTimbratura.tipo)}</div>
              <div className="text-[2.5rem] sm:text-4xl leading-none font-semibold tabular-nums mt-1 text-primary">{formatOra(me.ultimaTimbratura.ora)}</div>
            </div>
          ) : (
            <div className="mt-4 text-sm text-muted-foreground">Nessuna timbratura registrata oggi.</div>
          )}
        </div>
      </div>

      {/* Riepilogo ore */}
      <div className="mt-5 md:mt-6 grid gap-3 sm:gap-4 grid-cols-2 lg:grid-cols-4">
        <RiepilogoCard Icon={LogIn} label="Entrata" value={ore.entrataOra ? formatOra(ore.entrataOra) : "—"} />
        <RiepilogoCard Icon={Coffee} label="Pausa totale" value={formatDurata(ore.pausaMinuti)} hint={ore.inPausa ? "In corso" : undefined} />
        <RiepilogoCard Icon={Hourglass} label="Ore lavorate" value={formatDurata(ore.oreLavorateMinuti)} />
        <RiepilogoCard
          Icon={TrendingUp}
          label="Oltre orario"
          value={ore.oltreOrarioMinuti > 0 ? `+${formatDurata(ore.oltreOrarioMinuti)}` : "—"}
          highlight={ore.oltreOrarioMinuti > 0}
        />
      </div>

      <div className="mt-5 md:mt-6 grid gap-3 sm:gap-4 grid-cols-2 lg:grid-cols-4">
        {azioni.map((a) => (
          <button
            key={a.tipo}
            disabled={!a.enabled || busy}
            onClick={() => timbra(a.tipo)}
            title={a.reason ?? undefined}
            aria-label={`${a.label}${a.reason ? ` — ${a.reason}` : ""}`}
            className={`group relative rounded-2xl border p-4 sm:p-6 text-left transition-all min-h-[156px] sm:min-h-[176px] flex flex-col justify-between touch-manipulation
              disabled:cursor-not-allowed
              ${a.enabled ? "border-border bg-card shadow-[var(--shadow-card)] hover:shadow-[var(--shadow-elegant)] hover:-translate-y-1 active:translate-y-0 active:scale-[0.98]" : "border-dashed border-border/70 bg-muted/50 opacity-70"}
            `}
          >
            <div
              className={`inline-flex h-12 w-12 sm:h-14 sm:w-14 items-center justify-center rounded-xl text-white shadow-sm ${
                !a.enabled ? "bg-muted-foreground/50" :
                a.tone === "primary" ? "bg-primary" :
                a.tone === "warn" ? "bg-status-break" :
                a.tone === "ok" ? "bg-status-present" :
                "bg-status-absent"
              }`}
            >
              <a.Icon className="h-6 w-6 sm:h-7 sm:w-7" />
            </div>
            <div>
              <div className="text-base sm:text-lg font-semibold text-foreground leading-tight">{a.label}</div>
              <div className={`text-[11px] sm:text-xs mt-1 leading-snug ${a.enabled ? "text-muted-foreground" : "text-muted-foreground/90"}`}>
                {a.enabled ? "Tocca per registrare" : a.reason ?? "Non disponibile ora"}
              </div>
            </div>
          </button>
        ))}
      </div>

      {/* Timeline timbrature di oggi */}
      <div className="mt-5 md:mt-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex items-center gap-2 text-[13px] sm:text-sm text-muted-foreground mb-4">
          <ListChecks className="h-4 w-4" /> Timbrature di oggi
        </div>
        {eventiOggi.length === 0 ? (
          <div className="text-sm text-muted-foreground">Nessuna timbratura registrata oggi.</div>
        ) : (
          <ol className="relative border-l border-border ml-2 space-y-4">
            {eventiOggi.map((e, i) => (
              <li key={`${e.tipo}-${e.ora}-${i}`} className="pl-5 relative">
                <span className={`absolute -left-[7px] top-1.5 h-3 w-3 rounded-full ring-2 ring-card ${dotForEvento(e.tipo)}`} />
                <div className="flex items-baseline justify-between gap-3">
                  <div className="text-[14px] sm:text-[15px] font-medium text-foreground">{labelTipo(e.tipo)}</div>
                  <div className="text-[15px] sm:text-base font-semibold tabular-nums text-primary">{formatOra(e.ora)}</div>
                </div>
              </li>
            ))}
          </ol>
        )}
      </div>
    </AppShell>
  );
}

function dotForEvento(t: EventoTimbratura): string {
  switch (t) {
    case "entrata":
      return "bg-primary";
    case "inizio-pausa":
      return "bg-status-break";
    case "fine-pausa":
      return "bg-status-present";
    case "uscita":
      return "bg-status-absent";
  }
}

function RiepilogoCard({
  Icon,
  label,
  value,
  hint,
  highlight,
}: {
  Icon: typeof LogIn;
  label: string;
  value: string;
  hint?: string;
  highlight?: boolean;
}) {
  return (
    <div className={`rounded-2xl border p-4 sm:p-5 shadow-[var(--shadow-card)] transition-all ${highlight ? "border-status-out/40 bg-status-out/5" : "border-border bg-card"}`}>
      <div className="flex items-center justify-between gap-2">
        <span className="text-[11px] sm:text-xs uppercase tracking-wider text-muted-foreground">{label}</span>
        <span className={`h-8 w-8 rounded-lg flex items-center justify-center ${highlight ? "bg-status-out/15 text-status-out" : "bg-primary/10 text-primary"}`}>
          <Icon className="h-4 w-4" />
        </span>
      </div>
      <div className={`mt-2 text-2xl sm:text-[26px] leading-none font-semibold tabular-nums tracking-tight ${highlight ? "text-status-out" : "text-foreground"}`}>{value}</div>
      {hint && <div className="mt-1 text-[11px] text-muted-foreground">{hint}</div>}
    </div>
  );
}