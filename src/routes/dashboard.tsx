import { createFileRoute, redirect } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { AppShell } from "@/components/AppShell";
import {
  Users,
  UserCheck,
  Coffee,
  UserX,
  TrendingUp,
  RefreshCw,
  AlertTriangle,
  Clock,
  Building2,
} from "lucide-react";
import {
  aggregate,
  bySede,
  displayStato,
  DISPLAY_DOT,
  DISPLAY_LABEL,
  type DisplayStato,
} from "@/lib/data-service";
import { SEDI, formatOra, labelTipo, type SedeId, type Dipendente } from "@/lib/mock-data";
import { useLivePresenze } from "@/lib/use-live-presenze";

export const Route = createFileRoute("/dashboard")({
  head: () => ({ meta: [{ title: "Dashboard — DR Portal" }] }),
  beforeLoad: ({ location }) => {
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
  component: DashboardPage,
});

function DashboardPage() {
  const { data, lastUpdate, error, refresh } = useLivePresenze(15000);
  const totals = useMemo(() => aggregate(data), [data]);
  const [refreshing, setRefreshing] = useState(false);

  const sediStats = useMemo(
    () =>
      SEDI.map((s) => {
        const dip = bySede(data, s.id);
        const presenti = dip.filter((d) => displayStato(d) !== "assente").length;
        return { ...s, presenti, totale: dip.length };
      }),
    [data],
  );

  const handleRefresh = async () => {
    setRefreshing(true);
    await refresh();
    setRefreshing(false);
  };

  return (
    <AppShell title="Dashboard presenze" subtitle="Monitoraggio live sedi DR Logistica">
      {error && (
        <div className="mb-4 rounded-lg border border-status-absent/40 bg-status-absent/5 p-3 flex items-start gap-2">
          <AlertTriangle className="h-4 w-4 text-status-absent mt-0.5 shrink-0" />
          <div className="text-xs">
            <p className="font-medium text-status-absent">Sistema momentaneamente non disponibile</p>
            <p className="text-muted-foreground mt-0.5">
              Impossibile leggere i dati da SharePoint. Le timbrature sono temporaneamente
              disabilitate. Dettagli tecnici in Amministrazione.
            </p>
          </div>
        </div>
      )}

      {/* Sintesi presenze — riepilogo pulito per la direzione */}
      <section className="mb-6 rounded-2xl border border-border bg-card p-4 sm:p-5 md:p-6 shadow-[var(--shadow-card)] animate-fade-in">
        <div className="grid grid-cols-[minmax(0,1fr)_auto] items-start gap-4 mb-5 sm:flex sm:items-center sm:justify-between">
          <div className="min-w-0">
            <h2 className="text-[17px] sm:text-lg md:text-xl font-semibold text-foreground tracking-tight leading-tight">Sintesi presenze per sede</h2>
            <p className="text-xs text-muted-foreground mt-1 flex items-center gap-2">
              <span className="relative flex h-1.5 w-1.5 shrink-0">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-status-present opacity-70" />
                <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-status-present" />
              </span>
              <span className="truncate">
                Aggiornato alle{" "}
                {lastUpdate.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" })}
              </span>
            </p>
          </div>
          <button
            onClick={handleRefresh}
            disabled={refreshing}
            aria-label="Aggiorna sintesi presenze"
            className="shrink-0 inline-flex items-center gap-2 rounded-full bg-primary px-4 h-11 sm:h-9 min-w-11 text-sm font-medium text-primary-foreground shadow-sm hover:bg-primary/90 hover:shadow-[var(--shadow-elegant)] transition-all active:scale-[0.97] disabled:opacity-60 disabled:cursor-not-allowed touch-manipulation"
          >
            <RefreshCw className={`h-4 w-4 ${refreshing ? "animate-spin" : ""}`} />
            <span className="hidden sm:inline">Aggiorna ora</span>
            <span className="sm:hidden">Aggiorna</span>
          </button>
        </div>

        <div className="grid gap-3 md:gap-4 grid-cols-1 sm:grid-cols-2">
          {sediStats.map((s) => (
            <div
              key={s.id}
              className="rounded-xl border border-border bg-secondary/40 p-4 min-h-[84px] flex items-center justify-between gap-3 transition-all hover:bg-secondary/70 hover:shadow-sm"
            >
              <div className="flex items-center gap-3 min-w-0">
                <span className="h-12 w-12 shrink-0 rounded-xl bg-primary/10 text-primary flex items-center justify-center">
                  <Building2 className="h-[22px] w-[22px]" />
                </span>
                <div className="min-w-0">
                  <div className="text-[15px] font-semibold text-foreground truncate leading-tight">Sede {s.nome}</div>
                  <div className="text-xs text-muted-foreground mt-0.5">{s.totale} dipendenti totali</div>
                </div>
              </div>
              <div className="text-right shrink-0">
                <div className="text-[28px] sm:text-3xl font-semibold tabular-nums text-foreground tracking-tight leading-none">
                  {s.presenti}<span className="text-muted-foreground text-base sm:text-lg font-normal">/{s.totale}</span>
                </div>
                <div className="text-[10px] uppercase tracking-wider text-muted-foreground mt-1">Presenti</div>
              </div>
            </div>
          ))}
        </div>
      </section>

      {/* Top KPI cards */}
      <div className="grid gap-3 grid-cols-2 md:grid-cols-5">
        <KpiCard label="Dipendenti attivi" value={totals.attivi} Icon={Users} tone="primary" />
        <KpiCard label="Presenti" value={totals.presenti} Icon={UserCheck} tone="present" />
        <KpiCard label="In pausa" value={totals.pausa} Icon={Coffee} tone="break" />
        <KpiCard label="Assenti" value={totals.assenti} Icon={UserX} tone="absent" />
        <KpiCard label="Straordinari" value={totals.straordinari} Icon={TrendingUp} tone="out" className="col-span-2 md:col-span-1" />
      </div>


      {/* Live view: two side-by-side sede panels */}
      <div className="mt-6 grid gap-4 lg:grid-cols-2">
        {SEDI.map((s) => (
          <SedePanel key={s.id} sedeId={s.id} sedeName={s.nome} dipendenti={bySede(data, s.id)} />
        ))}
      </div>

      {/* Alerts */}
      <div className="mt-6 grid gap-4 lg:grid-cols-3">
        <AlertPanel
          Icon={AlertTriangle}
          tone="warn"
          title="Ritardi"
          items={data.filter((d) => (d.ritardoMinuti ?? 0) > 0)}
          renderMeta={(d) => `Atteso ${d.orarioAtteso} · ${sedeLabel(d.sede)}`}
          renderValue={(d) => `+${d.ritardoMinuti} min`}
        />
        <AlertPanel
          Icon={TrendingUp}
          tone="ok"
          title="Straordinari"
          items={data.filter((d) => (d.straordinariMinuti ?? 0) > 0)}
          renderMeta={(d) => `${d.ruolo} · ${sedeLabel(d.sede)}`}
          renderValue={(d) => `+${d.straordinariMinuti} min`}
        />
        <AlertPanel
          Icon={UserX}
          tone="danger"
          title="Non timbrati"
          items={data.filter((d) => d.stato === "non-timbrato")}
          renderMeta={(d) => `${d.ruolo} · ${sedeLabel(d.sede)}`}
          renderValue={(d) => `orario ${d.orarioAtteso}`}
        />
      </div>
    </AppShell>
  );
}

function sedeLabel(id: SedeId) {
  return SEDI.find((s) => s.id === id)?.nome ?? id;
}

function KpiCard({
  label,
  value,
  Icon,
  tone,
  className = "",
}: {
  label: string;
  value: number;
  Icon: typeof Users;
  tone: "primary" | "present" | "break" | "absent" | "out";
  className?: string;
}) {
  const bg = {
    primary: "bg-primary/10 text-primary",
    present: "bg-status-present/15 text-status-present",
    break: "bg-status-break/15 text-status-break",
    absent: "bg-status-absent/15 text-status-absent",
    out: "bg-status-out/15 text-status-out",
  }[tone];
  return (
    <div className={`rounded-xl border border-border bg-card p-4 shadow-[var(--shadow-card)] transition-all hover:shadow-[var(--shadow-elegant)] hover:-translate-y-0.5 ${className}`}>
      <div className="flex items-center justify-between gap-2">
        <span className="text-[11px] sm:text-xs text-muted-foreground uppercase tracking-wider leading-tight">{label}</span>
        <span className={`h-9 w-9 shrink-0 rounded-lg flex items-center justify-center ${bg}`}>
          <Icon className="h-4 w-4" />
        </span>
      </div>
      <div className="mt-2 text-[26px] sm:text-3xl leading-none font-semibold tabular-nums text-foreground tracking-tight">{value}</div>
    </div>
  );
}

function SedePanel({
  sedeId,
  sedeName,
  dipendenti,
}: {
  sedeId: SedeId;
  sedeName: string;
  dipendenti: Dipendente[];
}) {
  const [filter, setFilter] = useState<DisplayStato | "tutti">("tutti");
  const filtered = filter === "tutti" ? dipendenti : dipendenti.filter((d) => displayStato(d) === filter);
  const presenti = dipendenti.filter((d) => displayStato(d) === "presente" || displayStato(d) === "oltre" || displayStato(d) === "pausa").length;

  return (
    <section className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)] flex flex-col">
      <header className="flex items-center justify-between px-5 py-4 border-b border-border">
        <div>
          <h2 className="text-base font-semibold text-foreground">Sede {sedeName}</h2>
          <p className="text-xs text-muted-foreground">{dipendenti.length} dipendenti totali</p>
        </div>
        <div className="text-right">
          <div className="text-2xl font-semibold tabular-nums text-foreground">
            {presenti}<span className="text-muted-foreground text-base">/{dipendenti.length}</span>
          </div>
          <div className="text-[10px] uppercase tracking-wider text-muted-foreground">In sede</div>
        </div>
      </header>

      <div className="flex items-center gap-1 px-3 py-2 border-b border-border overflow-x-auto">
        {(["tutti", "presente", "pausa", "oltre", "assente"] as const).map((f) => (
          <button
            key={f}
            onClick={() => setFilter(f)}
            className={`text-xs px-2.5 py-1 rounded-md whitespace-nowrap transition-colors ${
              filter === f
                ? "bg-secondary text-secondary-foreground font-medium"
                : "text-muted-foreground hover:bg-secondary/60"
            }`}
          >
            {f === "tutti" ? "Tutti" : DISPLAY_LABEL[f]}
          </button>
        ))}
      </div>

      <ul className="divide-y divide-border flex-1">
        {filtered.map((d) => {
          const ds = displayStato(d);
          return (
            <li key={d.id} className="flex items-center gap-3 px-5 py-3 hover:bg-secondary/40">
              <div className="h-9 w-9 rounded-full bg-primary/10 text-primary text-xs font-semibold flex items-center justify-center shrink-0">
                {d.nome[0]}{d.cognome[0]}
              </div>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-medium text-foreground truncate">
                  {d.nome} {d.cognome}
                </div>
                <div className="text-xs text-muted-foreground truncate">
                  {d.ruolo}
                  {d.ultimaTimbratura && (
                    <> · {labelTipo(d.ultimaTimbratura.tipo)} {formatOra(d.ultimaTimbratura.ora)}</>
                  )}
                </div>
              </div>
              <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium bg-secondary text-secondary-foreground`}>
                <span className={`h-2 w-2 rounded-full ${DISPLAY_DOT[ds]}`} />
                {DISPLAY_LABEL[ds]}
              </span>
            </li>
          );
        })}
        {filtered.length === 0 && (
          <li className="text-center text-sm text-muted-foreground py-8">Nessun dipendente in questo stato</li>
        )}
      </ul>

      <footer className="px-5 py-3 border-t border-border bg-secondary/30 text-xs text-muted-foreground flex items-center justify-between">
        <span>Presenti <strong className="text-foreground tabular-nums">{presenti}</strong> / Totale <strong className="text-foreground tabular-nums">{dipendenti.length}</strong></span>
        <span className="flex items-center gap-1"><Clock className="h-3 w-3" /> {sedeId}</span>
      </footer>
    </section>
  );
}

function AlertPanel({
  Icon,
  tone,
  title,
  items,
  renderMeta,
  renderValue,
}: {
  Icon: typeof Users;
  tone: "warn" | "ok" | "danger";
  title: string;
  items: Dipendente[];
  renderMeta: (d: Dipendente) => string;
  renderValue: (d: Dipendente) => string;
}) {
  const color = tone === "warn" ? "text-status-break" : tone === "ok" ? "text-status-present" : "text-status-absent";
  const bg = tone === "warn" ? "bg-status-break/10" : tone === "ok" ? "bg-status-present/10" : "bg-status-absent/10";
  return (
    <div className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)]">
      <div className="flex items-center justify-between px-4 py-3 border-b border-border">
        <div className="flex items-center gap-2">
          <span className={`h-8 w-8 rounded-lg ${bg} ${color} flex items-center justify-center`}><Icon className="h-4 w-4" /></span>
          <h3 className="font-semibold text-foreground">{title}</h3>
        </div>
        <span className="text-xs text-muted-foreground tabular-nums">{items.length}</span>
      </div>
      <div className="p-2">
        {items.length === 0 ? (
          <div className="px-3 py-6 text-center text-sm text-muted-foreground">Nessun elemento</div>
        ) : (
          items.map((d) => (
            <div key={d.id} className="flex items-center justify-between px-2 py-2 rounded-md hover:bg-secondary/50">
              <div>
                <div className="text-sm font-medium text-foreground">{d.nome} {d.cognome}</div>
                <div className="text-xs text-muted-foreground">{renderMeta(d)}</div>
              </div>
              <div className={`text-sm font-semibold tabular-nums ${color}`}>{renderValue(d)}</div>
            </div>
          ))
        )}
      </div>
    </div>
  );
}