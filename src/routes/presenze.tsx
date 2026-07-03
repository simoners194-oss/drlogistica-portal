import { createFileRoute, redirect, useNavigate } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { LogIn, Coffee, PlayCircle, LogOut, Clock, Timer } from "lucide-react";
import { formatOra, labelTipo, type Dipendente, type Timbratura } from "@/lib/mock-data";
import { dataService, oreLavorateOggi, displayStato, DISPLAY_DOT, DISPLAY_LABEL } from "@/lib/data-service";

export const Route = createFileRoute("/presenze")({
  head: () => ({ meta: [{ title: "Le mie presenze — DR Portal" }] }),
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
      <AppShell title="Le mie presenze">
        <div className="text-sm text-muted-foreground">Caricamento…</div>
      </AppShell>
    );
  }

  const ds = displayStato(me);
  const azioni: {
    tipo: Timbratura["tipo"];
    label: string;
    Icon: typeof LogIn;
    enabled: boolean;
    tone: "primary" | "warn" | "ok" | "danger";
  }[] = [
    { tipo: "entrata", label: "Entrata", Icon: LogIn, enabled: me.stato === "non-timbrato" || me.stato === "uscito", tone: "primary" },
    { tipo: "inizio-pausa", label: "Inizio pausa", Icon: Coffee, enabled: me.stato === "presente", tone: "warn" },
    { tipo: "fine-pausa", label: "Fine pausa", Icon: PlayCircle, enabled: me.stato === "pausa", tone: "ok" },
    { tipo: "uscita", label: "Uscita", Icon: LogOut, enabled: me.stato === "presente" || me.stato === "pausa", tone: "danger" },
  ];

  return (
    <AppShell title="Le mie presenze" subtitle={`${me.nome} ${me.cognome} · ${me.ruolo}`}>
      <div className="grid gap-4 lg:grid-cols-3">
        <div
          className="lg:col-span-2 rounded-2xl p-6 text-primary-foreground shadow-[var(--shadow-elegant)]"
          style={{ background: "var(--gradient-hero)" }}
        >
          <div className="text-sm text-white/80 capitalize">
            {now.toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
          </div>
          <div className="mt-1 text-6xl font-semibold tabular-nums tracking-tight">
            {now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
          </div>
          <div className="mt-4 flex flex-wrap items-center gap-x-6 gap-y-2 text-sm">
            <span className="inline-flex items-center gap-2">
              <span className={`h-2.5 w-2.5 rounded-full ${DISPLAY_DOT[ds]} ring-2 ring-white/30`} />
              <span className="text-white/90">Stato: <strong>{DISPLAY_LABEL[ds]}</strong></span>
            </span>
            <span className="inline-flex items-center gap-2 text-white/90">
              <Timer className="h-4 w-4" /> Ore oggi: <strong className="tabular-nums">{oreLavorateOggi(me, now)}</strong>
            </span>
          </div>
        </div>

        <div className="rounded-2xl border border-border bg-card p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-muted-foreground text-sm">
            <Clock className="h-4 w-4" /> Ultima timbratura
          </div>
          {me.ultimaTimbratura ? (
            <div className="mt-3">
              <div className="text-lg font-semibold text-foreground">{labelTipo(me.ultimaTimbratura.tipo)}</div>
              <div className="text-4xl font-semibold tabular-nums mt-1 text-primary">{formatOra(me.ultimaTimbratura.ora)}</div>
            </div>
          ) : (
            <div className="mt-6 text-sm text-muted-foreground">Nessuna timbratura registrata oggi.</div>
          )}
        </div>
      </div>

      <div className="mt-6 grid gap-4 grid-cols-2 lg:grid-cols-4">
        {azioni.map((a) => (
          <button
            key={a.tipo}
            disabled={!a.enabled || busy}
            onClick={() => timbra(a.tipo)}
            className={`group relative rounded-2xl border p-6 text-left transition-all min-h-[160px] flex flex-col justify-between
              disabled:opacity-40 disabled:cursor-not-allowed
              ${a.enabled ? "border-border bg-card hover:shadow-[var(--shadow-elegant)] hover:-translate-y-1 active:translate-y-0" : "border-border bg-muted"}
            `}
          >
            <div
              className={`inline-flex h-14 w-14 items-center justify-center rounded-xl text-white ${
                a.tone === "primary" ? "bg-primary" :
                a.tone === "warn" ? "bg-status-break" :
                a.tone === "ok" ? "bg-status-present" :
                "bg-status-absent"
              }`}
            >
              <a.Icon className="h-7 w-7" />
            </div>
            <div>
              <div className="text-lg font-semibold text-foreground">{a.label}</div>
              <div className="text-xs text-muted-foreground mt-0.5">
                {a.enabled ? "Tocca per registrare" : "Non disponibile"}
              </div>
            </div>
          </button>
        ))}
      </div>
    </AppShell>
  );
}