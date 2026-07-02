import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { AppShell } from "@/components/AppShell";
import { LogIn, Coffee, PlayCircle, LogOut, Clock } from "lucide-react";
import { formatOra, labelTipo, STATO_COLOR, STATO_LABEL, type StatoTimbratura, type Timbratura } from "@/lib/mock-data";

export const Route = createFileRoute("/dipendente")({
  head: () => ({ meta: [{ title: "Area Dipendente — DR Portal" }] }),
  component: DipendentePage,
});

const UTENTE = { nome: "Marco Rossi", ruolo: "Magazziniere", sede: "Roma" };

function DipendentePage() {
  const [now, setNow] = useState(new Date());
  const [stato, setStato] = useState<StatoTimbratura>("non-timbrato");
  const [ultima, setUltima] = useState<Timbratura | undefined>();

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const timbra = (tipo: Timbratura["tipo"]) => {
    const nuova: Timbratura = { tipo, ora: new Date().toISOString() };
    setUltima(nuova);
    setStato(
      tipo === "entrata" || tipo === "fine-pausa"
        ? "presente"
        : tipo === "inizio-pausa"
          ? "pausa"
          : "uscito",
    );
  };

  const azioni: { tipo: Timbratura["tipo"]; label: string; Icon: typeof LogIn; enabled: boolean; variant: "primary" | "warn" | "ok" | "danger" }[] = [
    { tipo: "entrata", label: "Entrata", Icon: LogIn, enabled: stato === "non-timbrato" || stato === "uscito", variant: "primary" },
    { tipo: "inizio-pausa", label: "Inizio pausa", Icon: Coffee, enabled: stato === "presente", variant: "warn" },
    { tipo: "fine-pausa", label: "Fine pausa", Icon: PlayCircle, enabled: stato === "pausa", variant: "ok" },
    { tipo: "uscita", label: "Uscita", Icon: LogOut, enabled: stato === "presente" || stato === "pausa", variant: "danger" },
  ];

  return (
    <AppShell role="dipendente">
      <h1 className="text-2xl font-semibold text-foreground">Ciao, {UTENTE.nome.split(" ")[0]}</h1>
      <p className="text-sm text-muted-foreground">{UTENTE.ruolo} · Sede {UTENTE.sede}</p>

      <div className="mt-6 grid gap-4 md:grid-cols-3">
        <div className="md:col-span-2 rounded-xl p-6 text-primary-foreground shadow-[var(--shadow-elegant)]" style={{ background: "var(--gradient-hero)" }}>
          <div className="text-sm text-white/80 capitalize">
            {now.toLocaleDateString("it-IT", { weekday: "long", day: "numeric", month: "long", year: "numeric" })}
          </div>
          <div className="mt-2 text-5xl font-semibold tabular-nums tracking-tight">
            {now.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
          </div>
          <div className="mt-4 flex items-center gap-2">
            <span className={`inline-block h-2.5 w-2.5 rounded-full ${STATO_COLOR[stato]}`} />
            <span className="text-sm text-white/90">Stato: <strong>{STATO_LABEL[stato]}</strong></span>
          </div>
        </div>

        <div className="rounded-xl border border-border bg-card p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-muted-foreground text-sm">
            <Clock className="h-4 w-4" /> Ultima timbratura
          </div>
          {ultima ? (
            <div className="mt-3">
              <div className="text-lg font-semibold text-foreground">{labelTipo(ultima.tipo)}</div>
              <div className="text-3xl font-semibold tabular-nums mt-1 text-primary">{formatOra(ultima.ora)}</div>
            </div>
          ) : (
            <div className="mt-6 text-sm text-muted-foreground">Nessuna timbratura registrata oggi.</div>
          )}
        </div>
      </div>

      <div className="mt-6 grid gap-3 grid-cols-2 md:grid-cols-4">
        {azioni.map((a) => (
          <button
            key={a.tipo}
            disabled={!a.enabled}
            onClick={() => timbra(a.tipo)}
            className={`group relative rounded-xl border p-5 text-left transition-all disabled:opacity-40 disabled:cursor-not-allowed
              ${a.enabled ? "border-border bg-card hover:shadow-[var(--shadow-elegant)] hover:-translate-y-0.5" : "border-border bg-muted"}
            `}
          >
            <div
              className={`inline-flex h-10 w-10 items-center justify-center rounded-lg text-white ${
                a.variant === "primary" ? "bg-primary" :
                a.variant === "warn" ? "bg-status-break" :
                a.variant === "ok" ? "bg-status-present" :
                "bg-status-absent"
              }`}
            >
              <a.Icon className="h-5 w-5" />
            </div>
            <div className="mt-3 font-medium text-foreground">{a.label}</div>
            <div className="text-xs text-muted-foreground mt-0.5">Registra timbratura</div>
          </button>
        ))}
      </div>
    </AppShell>
  );
}