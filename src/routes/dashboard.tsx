import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { AppShell } from "@/components/AppShell";
import {
  Users,
  UserCheck,
  Coffee,
  UserX,
  TrendingUp,
  LogOut,
  RefreshCw,
  AlertTriangle,
  Clock,
  Building2,
  ChevronRight,
  BarChart3,
  FileText,
  Settings,
  Activity,
} from "lucide-react";
import {
  aggregate,
  bySede,
  displayStato,
  DISPLAY_DOT,
  DISPLAY_LABEL,
  type DisplayStato,
} from "@/lib/data-service";
import { sedeTimbra, formatOra, labelTipo, type SedeId, type Dipendente } from "@/lib/mock-data";
import { useLivePresenze } from "@/lib/use-live-presenze";
import { formatDurata } from "@/lib/presenze-logic";
import { DashboardSkeleton } from "@/components/skeletons/DashboardSkeleton";
import { DettaglioDipendenteDialog } from "@/components/DettaglioDipendenteDialog";
import { readSession, type Ruolo, type SessionUser } from "@/lib/session";
import { QuickAccess, type QuickAccessItem } from "@/components/QuickAccess";

export const Route = createFileRoute("/dashboard")({
  head: () => ({ meta: [{ title: "Dashboard — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    const session = readSession();
    if (!session) {
      throw redirect({
        to: "/",
        search: { redirect: location.href },
      });
    }
    // Dashboard riservata a Responsabili e Amministratori: il Dipendente
    // viene reindirizzato alle proprie Presenze.
    if (session.ruolo === "dipendente") {
      throw redirect({ to: "/presenze" });
    }
  },
  component: DashboardPage,
});

function DashboardPage() {
  const { data, lastUpdate, error, refresh, loading } = useLivePresenze(15000);
  const [refreshing, setRefreshing] = useState(false);
  const [session, setSession] = useState<SessionUser | null>(null);
  const [selected, setSelected] = useState<Dipendente | null>(null);
  useEffect(() => {
    setSession(readSession());
  }, []);

  const ruolo: Ruolo = session?.ruolo ?? "amministratore_sistema";
  const isResponsabile = ruolo === "responsabile";
  const isAdmin = ruolo === "amministratore_sistema";

  // Timbratura attiva se almeno un dipendente appartiene a una sede timbrante.
  const timbraturaAttiva = data.some((d) => sedeTimbra(d.sede));

  const quickItems: QuickAccessItem[] = isAdmin
    ? [
        {
          label: "Presenze",
          to: "/presenze",
          Icon: Clock,
          ready: timbraturaAttiva,
          disabledNote: "Sede senza timbratura",
        },
        { label: "Report", to: "/report", Icon: BarChart3, ready: false },
        { label: "Dipendenti", Icon: Users, ready: false },
        { label: "Amministrazione", to: "/amministrazione", Icon: Settings, ready: true },
        {
          label: "Diagnostica",
          to: "/amministrazione",
          Icon: Activity,
          ready: true,
          description: "Health & self-test",
        },
      ]
    : [
        {
          label: "Presenze",
          to: "/presenze",
          Icon: Clock,
          ready: timbraturaAttiva,
          disabledNote: "Sede senza timbratura",
        },
        { label: "Report", to: "/report", Icon: BarChart3, ready: false },
        { label: "Dipendenti", Icon: Users, ready: false },
        { label: "Richieste", to: "/richieste", Icon: FileText, ready: false },
      ];

  // Sia il Responsabile sia l'Amministratore di sistema vedono i dati di
  // TUTTE le sedi. Il Responsabile mantiene però una vista in sola lettura
  // (niente pannelli operativi di alert azionabili).
  const scopedData = data;

  // Solo dipendenti di sedi che timbrano entrano nelle viste presenze.
  const presenzeData = useMemo(() => scopedData.filter((d) => sedeTimbra(d.sede)), [scopedData]);
  const totals = useMemo(() => aggregate(presenzeData), [presenzeData]);
  // Mantiene sincronizzato il dettaglio con l'ultimo snapshot: chi è aperto
  // vede gli aggiornamenti live delle proprie timbrature.
  const selectedLive = useMemo(
    () => (selected ? (scopedData.find((d) => d.id === selected.id) ?? selected) : null),
    [scopedData, selected],
  );

  // Sedi timbranti effettivamente presenti nei dati (distinte, ordinate).
  const sediStats = useMemo(() => {
    const seen = new Set<string>();
    const nomi: string[] = [];
    for (const d of presenzeData) {
      const s = (d.sede ?? "").trim();
      if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
        seen.add(s.toLowerCase());
        nomi.push(s);
      }
    }
    nomi.sort((a, b) => a.localeCompare(b));
    return nomi.map((nome) => {
      const dip = bySede(presenzeData, nome);
      const presenti = dip.filter((d) => displayStato(d) !== "assente").length;
      return { id: nome, nome, presenti, totale: dip.length };
    });
  }, [presenzeData]);

  const visibleSedi = sediStats;

  const title = isResponsabile ? "Dashboard responsabili" : "Dashboard presenze";
  const subtitle = isResponsabile
    ? "Panoramica live · sola lettura · tutte le sedi"
    : "Monitoraggio live sedi DR Logistica";

  const handleRefresh = async () => {
    setRefreshing(true);
    await refresh();
    setRefreshing(false);
  };

  return (
    <AppShell title={title} subtitle={subtitle}>
      {error && (
        <div className="mb-4 rounded-lg border border-status-absent/40 bg-status-absent/5 p-3 flex items-start gap-2">
          <AlertTriangle className="h-4 w-4 text-status-absent mt-0.5 shrink-0" />
          <div className="text-xs">
            <p className="font-medium text-status-absent">
              Sistema momentaneamente non disponibile
            </p>
            <p className="text-muted-foreground mt-0.5">
              Impossibile leggere i dati da SharePoint. Le timbrature sono temporaneamente
              disabilitate. Dettagli tecnici in Amministrazione.
            </p>
          </div>
        </div>
      )}

      {loading && scopedData.length === 0 ? (
        <DashboardSkeleton />
      ) : !timbraturaAttiva ? (
        <div className="rounded-2xl border border-border bg-card p-6 text-center shadow-[var(--shadow-card)]">
          <div className="text-sm font-semibold text-foreground">Timbrature non attive</div>
          <p className="text-[13px] text-muted-foreground mt-1 max-w-md mx-auto">
            Nessuna sede ha la timbratura attiva: la panoramica presenze comparirà automaticamente
            quando una sede abiliterà la timbratura.
          </p>
        </div>
      ) : (
        <>
          {/* Sintesi presenze — riepilogo pulito per la direzione */}
          <section className="mb-6 rounded-2xl border border-border bg-card p-4 sm:p-5 md:p-6 shadow-[var(--shadow-card)] animate-fade-in">
            <div className="grid grid-cols-[minmax(0,1fr)_auto] items-start gap-4 mb-5 sm:flex sm:items-center sm:justify-between">
              <div className="min-w-0">
                <h2 className="text-[17px] sm:text-lg md:text-xl font-semibold text-foreground tracking-tight leading-tight">
                  Sintesi presenze per sede
                </h2>
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
                      <div className="text-[15px] font-semibold text-foreground truncate leading-tight">
                        Sede {s.nome}
                      </div>
                      <div className="text-xs text-muted-foreground mt-0.5">
                        {s.totale} dipendenti totali
                      </div>
                    </div>
                  </div>
                  <div className="text-right shrink-0">
                    <div className="text-[28px] sm:text-3xl font-semibold tabular-nums text-foreground tracking-tight leading-none">
                      {s.presenti}
                      <span className="text-muted-foreground text-base sm:text-lg font-normal">
                        /{s.totale}
                      </span>
                    </div>
                    <div className="text-[10px] uppercase tracking-wider text-muted-foreground mt-1">
                      Presenti
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </section>

          {/* KPI: presenti, in pausa, usciti, non timbrati, in straordinario */}
          <div className="grid gap-3 grid-cols-2 md:grid-cols-5">
            <KpiCard label="Presenti" value={totals.presenti} Icon={UserCheck} tone="present" />
            <KpiCard label="In pausa" value={totals.pausa} Icon={Coffee} tone="break" />
            <KpiCard label="Usciti" value={totals.usciti} Icon={LogOut} tone="primary" />
            <KpiCard label="Non timbrati" value={totals.assenti} Icon={UserX} tone="absent" />
            <KpiCard
              label="In straordinario"
              value={totals.oltre}
              Icon={TrendingUp}
              tone="out"
              className="col-span-2 md:col-span-1"
            />
          </div>

          {/* Elenco dipendenti — clic su una riga apre il dettaglio giornaliero */}
          <div className={`mt-6 grid gap-4 ${visibleSedi.length > 1 ? "lg:grid-cols-2" : ""}`}>
            {visibleSedi.map((s) => (
              <SedePanel
                key={s.id}
                sedeId={s.id}
                sedeName={s.nome}
                dipendenti={bySede(scopedData, s.id)}
                onSelect={setSelected}
              />
            ))}
          </div>

          {/* Alerts: nascosti al Responsabile per mantenere il focus operativo */}
          {!isResponsabile && (
            <div className="mt-6 grid gap-4 lg:grid-cols-3">
              <AlertPanel
                Icon={AlertTriangle}
                tone="warn"
                title="Ritardi"
                items={presenzeData.filter((d) => (d.ritardoMinuti ?? 0) > 0)}
                renderMeta={(d) => `Atteso ${d.orarioAtteso} · ${sedeLabel(d.sede)}`}
                renderValue={(d) => `+${d.ritardoMinuti} min`}
              />
              <AlertPanel
                Icon={TrendingUp}
                tone="ok"
                title="In straordinario"
                items={presenzeData.filter((d) => (d.oltreOrarioMinuti ?? 0) > 0)}
                renderMeta={(d) => `${d.ruolo} · ${sedeLabel(d.sede)}`}
                renderValue={(d) => `+${formatDurata(d.oltreOrarioMinuti ?? 0)}`}
              />
              <AlertPanel
                Icon={UserX}
                tone="danger"
                title="Non timbrati"
                items={presenzeData.filter((d) => d.stato === "non-timbrato")}
                renderMeta={(d) => `${d.ruolo} · ${sedeLabel(d.sede)}`}
                renderValue={(d) => `orario ${d.orarioAtteso}`}
              />
            </div>
          )}
        </>
      )}

      <DettaglioDipendenteDialog
        dipendente={selectedLive}
        open={selected !== null}
        onOpenChange={(o) => {
          if (!o) setSelected(null);
        }}
      />
      <QuickAccess items={quickItems} />
    </AppShell>
  );
}

function sedeLabel(id: SedeId) {
  return id;
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
    <div
      className={`rounded-xl border border-border bg-card p-4 shadow-[var(--shadow-card)] transition-all hover:shadow-[var(--shadow-elegant)] hover:-translate-y-0.5 ${className}`}
    >
      <div className="flex items-center justify-between gap-2">
        <span className="text-[11px] sm:text-xs text-muted-foreground uppercase tracking-wider leading-tight">
          {label}
        </span>
        <span className={`h-9 w-9 shrink-0 rounded-lg flex items-center justify-center ${bg}`}>
          <Icon className="h-4 w-4" />
        </span>
      </div>
      <div className="mt-2 text-[26px] sm:text-3xl leading-none font-semibold tabular-nums text-foreground tracking-tight">
        {value}
      </div>
    </div>
  );
}

function SedePanel({
  sedeId,
  sedeName,
  dipendenti,
  onSelect,
}: {
  sedeId: SedeId;
  sedeName: string;
  dipendenti: Dipendente[];
  onSelect?: (d: Dipendente) => void;
}) {
  const [filter, setFilter] = useState<DisplayStato | "tutti">("tutti");
  const filtered =
    filter === "tutti" ? dipendenti : dipendenti.filter((d) => displayStato(d) === filter);
  const presenti = dipendenti.filter(
    (d) =>
      displayStato(d) === "presente" || displayStato(d) === "oltre" || displayStato(d) === "pausa",
  ).length;

  return (
    <section className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)] flex flex-col">
      <header className="flex items-center justify-between gap-3 px-4 sm:px-5 py-4 border-b border-border">
        <div className="min-w-0">
          <h2 className="text-[15px] sm:text-base font-semibold text-foreground truncate">
            Sede {sedeName}
          </h2>
          <p className="text-xs text-muted-foreground">{dipendenti.length} dipendenti totali</p>
        </div>
        <div className="text-right shrink-0">
          <div className="text-2xl font-semibold tabular-nums text-foreground leading-none">
            {presenti}
            <span className="text-muted-foreground text-base">/{dipendenti.length}</span>
          </div>
          <div className="text-[10px] uppercase tracking-wider text-muted-foreground mt-1">
            In sede
          </div>
        </div>
      </header>

      <div className="flex items-center gap-1.5 px-3 py-2.5 border-b border-border overflow-x-auto scrollbar-none">
        {(["tutti", "presente", "pausa", "oltre", "assente"] as const).map((f) => (
          <button
            key={f}
            onClick={() => setFilter(f)}
            className={`text-xs px-3 py-1.5 min-h-8 rounded-full whitespace-nowrap transition-colors touch-manipulation ${
              filter === f
                ? "bg-primary text-primary-foreground font-medium shadow-sm"
                : "text-muted-foreground bg-secondary/50 hover:bg-secondary"
            }`}
          >
            {f === "tutti" ? "Tutti" : DISPLAY_LABEL[f]}
          </button>
        ))}
      </div>

      <ul className="divide-y divide-border flex-1">
        {filtered.map((d) => {
          const ds = displayStato(d);
          const clickable = Boolean(onSelect);
          return (
            <li key={d.id}>
              <button
                type="button"
                onClick={() => onSelect?.(d)}
                disabled={!clickable}
                aria-label={
                  clickable ? `Apri dettaglio giornaliero di ${d.nome} ${d.cognome}` : undefined
                }
                className={`w-full flex items-center gap-3 px-4 sm:px-5 py-3.5 text-left transition-colors ${clickable ? "hover:bg-secondary/60 focus-visible:bg-secondary/70 focus:outline-none active:bg-secondary" : ""}`}
              >
                <div className="h-10 w-10 rounded-full bg-primary/10 text-primary text-xs font-semibold flex items-center justify-center shrink-0">
                  {d.nome[0]}
                  {d.cognome[0]}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="text-[14px] font-medium text-foreground truncate">
                    {d.nome} {d.cognome}
                  </div>
                  <div className="text-[12px] text-muted-foreground truncate mt-0.5">
                    {d.ruolo}
                    {d.ultimaTimbratura && (
                      <>
                        {" "}
                        · {labelTipo(d.ultimaTimbratura.tipo)} {formatOra(d.ultimaTimbratura.ora)}
                      </>
                    )}
                  </div>
                </div>
                <span
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] sm:text-xs font-medium bg-secondary text-secondary-foreground shrink-0`}
                >
                  <span className={`h-2 w-2 rounded-full ${DISPLAY_DOT[ds]}`} />
                  <span className="hidden sm:inline">{DISPLAY_LABEL[ds]}</span>
                  <span className="sm:hidden">{DISPLAY_LABEL[ds].slice(0, 3)}.</span>
                </span>
                {clickable && <ChevronRight className="h-4 w-4 text-muted-foreground shrink-0" />}
              </button>
            </li>
          );
        })}
        {filtered.length === 0 && (
          <li className="text-center text-sm text-muted-foreground py-8">
            Nessun dipendente in questo stato
          </li>
        )}
      </ul>

      <footer className="px-4 sm:px-5 py-3 border-t border-border bg-secondary/30 text-xs text-muted-foreground flex items-center justify-between gap-2">
        <span>
          Presenti <strong className="text-foreground tabular-nums">{presenti}</strong> / Totale{" "}
          <strong className="text-foreground tabular-nums">{dipendenti.length}</strong>
        </span>
        <span className="flex items-center gap-1 shrink-0">
          <Clock className="h-3 w-3" /> {sedeId}
        </span>
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
  const color =
    tone === "warn"
      ? "text-status-break"
      : tone === "ok"
        ? "text-status-present"
        : "text-status-absent";
  const bg =
    tone === "warn"
      ? "bg-status-break/10"
      : tone === "ok"
        ? "bg-status-present/10"
        : "bg-status-absent/10";
  return (
    <div className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)]">
      <div className="flex items-center justify-between px-4 py-3.5 border-b border-border">
        <div className="flex items-center gap-2.5">
          <span className={`h-9 w-9 rounded-lg ${bg} ${color} flex items-center justify-center`}>
            <Icon className="h-4 w-4" />
          </span>
          <h3 className="text-[15px] font-semibold text-foreground">{title}</h3>
        </div>
        <span className="text-xs font-medium text-muted-foreground tabular-nums bg-secondary rounded-full px-2 py-0.5 min-w-6 text-center">
          {items.length}
        </span>
      </div>
      <div className="p-2">
        {items.length === 0 ? (
          <div className="px-3 py-6 text-center text-sm text-muted-foreground">Nessun elemento</div>
        ) : (
          items.map((d) => (
            <div
              key={d.id}
              className="flex items-center justify-between gap-3 px-3 py-2.5 rounded-lg hover:bg-secondary/50 transition-colors"
            >
              <div className="min-w-0">
                <div className="text-sm font-medium text-foreground truncate">
                  {d.nome} {d.cognome}
                </div>
                <div className="text-xs text-muted-foreground truncate">{renderMeta(d)}</div>
              </div>
              <div className={`text-sm font-semibold tabular-nums shrink-0 ${color}`}>
                {renderValue(d)}
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
}
