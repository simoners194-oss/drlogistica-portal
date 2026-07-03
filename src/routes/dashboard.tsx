import { createFileRoute } from "@tanstack/react-router";
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
  head: () => ({ meta: [{ title: "Dashboard HR — DR Portal" }] }),
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

      <div className="flex items-center justify-between mb-4">
        <div className="inline-flex items-center gap-2 text-xs text-muted-foreground">
          <span className="relative flex h-2 w-2">
            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-status-present opacity-75" />
            <span className="relative inline-flex rounded-full h-2 w-2 bg-status-present" />
          </span>
          Live · aggiornato alle {lastUpdate.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
        </div>
        <button
          className="inline-flex items-center gap-1.5 text-xs text-muted-foreground hover:text-foreground"
          onClick={() => window.location.reload()}
        >
          <RefreshCw className="h-3.5 w-3.5" /> Aggiorna
        </button>
      </div>

      {/* Demo domani — riepilogo pulito per riunione */}
      <section className="mb-6 rounded-xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="flex items-start justify-between gap-4 mb-5">
          <div>
            <h2 className="text-lg font-semibold text-foreground">Demo domani</h2>
            <p className="text-xs text-muted-foreground">
              Presenze per sede · aggiornato alle{" "}
              {lastUpdate.toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit", second: "2-digit" })}
            </p>
          </div>
          <button
            onClick={handleRefresh}
            disabled={refreshing}
            className="inline-flex items-center gap-1.5 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 disabled:opacity-60 disabled:cursor-not-allowed"
          >
            <RefreshCw className={`h-4 w-4 ${refreshing ? "animate-spin" : ""}`} />
            Aggiorna ora
          </button>
        </div>

        <div className="grid gap-4 grid-cols-1 sm:grid-cols-2">
          {sediStats.map((s) => (
            <div
              key={s.id}
              className="rounded-lg border border-border bg-secondary/40 p-4 flex items-center justify-between"
            >
              <div className="flex items-center gap-3">
                <span className="h-10 w-10 rounded-lg bg-primary/10 text-primary flex items-center justify-center">
                  <Building2 className="h-5 w-5" />
                </span>
                <div>
                  <div className="text-sm font-medium text-foreground">Sede {s.nome}</div>
                  <div className="text-xs text-muted-foreground">{s.totale} dipendenti totali</div>
                </div>
              </div>
              <div className="text-right">
                <div className="text-3xl font-semibold tabular-nums text-foreground">
                  {s.presenti}<span className="text-muted-foreground text-lg font-normal">/{s.totale}</span>
                </div>
                <div className="text-[10px] uppercase tracking-wider text-muted-foreground">Presenti</div>
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
        <KpiCard label="Straordinari" value={totals.straordinari} Icon={TrendingUp} tone="out" />
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
}: {
  label: string;
  value: number;
  Icon: typeof Users;
  tone: "primary" | "present" | "break" | "absent" | "out";
}) {
  const bg = {
    primary: "bg-primary/10 text-primary",
    present: "bg-status-present/15 text-status-present",
    break: "bg-status-break/15 text-status-break",
    absent: "bg-status-absent/15 text-status-absent",
    out: "bg-status-out/15 text-status-out",
  }[tone];
  return (
    <div className="rounded-xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
      <div className="flex items-center justify-between">
        <span className="text-xs text-muted-foreground uppercase tracking-wider">{label}</span>
        <span className={`h-8 w-8 rounded-lg flex items-center justify-center ${bg}`}>
          <Icon className="h-4 w-4" />
        </span>
      </div>
      <div className="mt-2 text-3xl font-semibold tabular-nums text-foreground">{value}</div>
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