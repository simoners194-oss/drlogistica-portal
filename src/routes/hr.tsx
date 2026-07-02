import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";
import { AppShell } from "@/components/AppShell";
import { DIPENDENTI, SEDI, STATO_COLOR, STATO_LABEL, formatOra, labelTipo, type SedeId, type Dipendente } from "@/lib/mock-data";
import { Users, Coffee, UserX, AlertTriangle, TrendingUp } from "lucide-react";

export const Route = createFileRoute("/hr")({
  head: () => ({ meta: [{ title: "Dashboard HR — DR Portal" }] }),
  component: HrPage,
});

function HrPage() {
  const [sede, setSede] = useState<SedeId>("roma");
  const perSede = (id: SedeId) => DIPENDENTI.filter((d) => d.sede === id);
  const attivi = perSede(sede);

  const conta = (list: Dipendente[]) => ({
    presenti: list.filter((d) => d.stato === "presente").length,
    pausa: list.filter((d) => d.stato === "pausa").length,
    nonTimbrati: list.filter((d) => d.stato === "non-timbrato").length,
    usciti: list.filter((d) => d.stato === "uscito").length,
  });

  const ritardi = attivi.filter((d) => (d.ritardoMinuti ?? 0) > 0);
  const straordinari = attivi.filter((d) => (d.straordinariMinuti ?? 0) > 0);
  const nonTimbrati = attivi.filter((d) => d.stato === "non-timbrato");

  return (
    <AppShell role="hr">
      <div className="flex items-end justify-between gap-4 flex-wrap">
        <div>
          <h1 className="text-2xl font-semibold text-foreground">Dashboard presenze</h1>
          <p className="text-sm text-muted-foreground">Monitoraggio live sedi DR Logistica</p>
        </div>
        <div className="inline-flex rounded-lg border border-border bg-card p-1 shadow-sm">
          {SEDI.map((s) => (
            <button
              key={s.id}
              onClick={() => setSede(s.id)}
              className={`px-4 py-1.5 text-sm rounded-md transition-colors ${
                sede === s.id ? "bg-primary text-primary-foreground font-medium" : "text-muted-foreground hover:text-foreground"
              }`}
            >
              {s.nome}
            </button>
          ))}
        </div>
      </div>

      {/* Overview both sedi */}
      <div className="mt-6 grid gap-4 md:grid-cols-2">
        {SEDI.map((s) => {
          const c = conta(perSede(s.id));
          const active = s.id === sede;
          return (
            <button
              key={s.id}
              onClick={() => setSede(s.id)}
              className={`text-left rounded-xl border p-5 transition-all ${
                active ? "border-primary shadow-[var(--shadow-elegant)] bg-card" : "border-border bg-card hover:border-primary/40"
              }`}
            >
              <div className="flex items-center justify-between">
                <div className="text-base font-semibold text-foreground">Sede {s.nome}</div>
                <div className="text-xs text-muted-foreground">{perSede(s.id).length} dipendenti</div>
              </div>
              <div className="mt-4 grid grid-cols-3 gap-3">
                <Stat label="Presenti" value={c.presenti} tone="present" />
                <Stat label="In pausa" value={c.pausa} tone="break" />
                <Stat label="Non timbrati" value={c.nonTimbrati} tone="absent" />
              </div>
            </button>
          );
        })}
      </div>

      {/* Employee list */}
      <section className="mt-8">
        <div className="flex items-center justify-between mb-3">
          <h2 className="text-lg font-semibold text-foreground">Dipendenti — {SEDI.find((s) => s.id === sede)?.nome}</h2>
          <span className="text-xs text-muted-foreground">{attivi.length} totali</span>
        </div>
        <div className="rounded-xl border border-border bg-card overflow-hidden shadow-[var(--shadow-card)]">
          <table className="w-full text-sm">
            <thead className="bg-secondary/60 text-muted-foreground">
              <tr className="text-left">
                <th className="px-4 py-2.5 font-medium">Dipendente</th>
                <th className="px-4 py-2.5 font-medium">Ruolo</th>
                <th className="px-4 py-2.5 font-medium">Stato</th>
                <th className="px-4 py-2.5 font-medium">Ultima timbratura</th>
              </tr>
            </thead>
            <tbody>
              {attivi.map((d) => (
                <tr key={d.id} className="border-t border-border hover:bg-secondary/40">
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-3">
                      <Avatar nome={d.nome} cognome={d.cognome} />
                      <div className="font-medium text-foreground">{d.nome} {d.cognome}</div>
                    </div>
                  </td>
                  <td className="px-4 py-3 text-muted-foreground">{d.ruolo}</td>
                  <td className="px-4 py-3"><StatoBadge stato={d.stato} /></td>
                  <td className="px-4 py-3 text-muted-foreground tabular-nums">
                    {d.ultimaTimbratura ? `${labelTipo(d.ultimaTimbratura.tipo)} · ${formatOra(d.ultimaTimbratura.ora)}` : "—"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Sections */}
      <div className="mt-8 grid gap-4 lg:grid-cols-3">
        <Panel Icon={AlertTriangle} tone="warn" title="Ritardi" count={ritardi.length}>
          {ritardi.length === 0 ? <Empty text="Nessun ritardo oggi" /> : ritardi.map((d) => (
            <Row key={d.id} nome={`${d.nome} ${d.cognome}`} meta={`Atteso ${d.orarioAtteso}`} value={`+${d.ritardoMinuti} min`} tone="warn" />
          ))}
        </Panel>
        <Panel Icon={TrendingUp} tone="ok" title="Straordinari" count={straordinari.length}>
          {straordinari.length === 0 ? <Empty text="Nessuno straordinario" /> : straordinari.map((d) => (
            <Row key={d.id} nome={`${d.nome} ${d.cognome}`} meta={d.ruolo} value={`+${d.straordinariMinuti} min`} tone="ok" />
          ))}
        </Panel>
        <Panel Icon={UserX} tone="danger" title="Non timbrati" count={nonTimbrati.length}>
          {nonTimbrati.length === 0 ? <Empty text="Tutti hanno timbrato" /> : nonTimbrati.map((d) => (
            <Row key={d.id} nome={`${d.nome} ${d.cognome}`} meta={d.ruolo} value={`orario ${d.orarioAtteso}`} tone="danger" />
          ))}
        </Panel>
      </div>
    </AppShell>
  );
}

function Stat({ label, value, tone }: { label: string; value: number; tone: "present" | "break" | "absent" }) {
  const dot = tone === "present" ? "bg-status-present" : tone === "break" ? "bg-status-break" : "bg-status-absent";
  return (
    <div className="rounded-lg bg-secondary/50 px-3 py-2.5">
      <div className="flex items-center gap-1.5 text-[11px] text-muted-foreground">
        <span className={`h-1.5 w-1.5 rounded-full ${dot}`} />
        {label}
      </div>
      <div className="text-2xl font-semibold tabular-nums text-foreground mt-0.5">{value}</div>
    </div>
  );
}

function StatoBadge({ stato }: { stato: Dipendente["stato"] }) {
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-medium bg-secondary text-secondary-foreground">
      <span className={`h-2 w-2 rounded-full ${STATO_COLOR[stato]}`} />
      {STATO_LABEL[stato]}
    </span>
  );
}

function Avatar({ nome, cognome }: { nome: string; cognome: string }) {
  return (
    <div className="h-8 w-8 rounded-full bg-primary/10 text-primary text-xs font-semibold flex items-center justify-center">
      {nome[0]}{cognome[0]}
    </div>
  );
}

function Panel({ Icon, tone, title, count, children }: { Icon: typeof Users; tone: "warn" | "ok" | "danger"; title: string; count: number; children: React.ReactNode }) {
  const color = tone === "warn" ? "text-status-break" : tone === "ok" ? "text-status-present" : "text-status-absent";
  const bg = tone === "warn" ? "bg-status-break/10" : tone === "ok" ? "bg-status-present/10" : "bg-status-absent/10";
  return (
    <div className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)]">
      <div className="flex items-center justify-between px-4 py-3 border-b border-border">
        <div className="flex items-center gap-2">
          <span className={`h-8 w-8 rounded-lg ${bg} ${color} flex items-center justify-center`}><Icon className="h-4 w-4" /></span>
          <h3 className="font-semibold text-foreground">{title}</h3>
        </div>
        <span className="text-xs text-muted-foreground tabular-nums">{count}</span>
      </div>
      <div className="p-2">{children}</div>
    </div>
  );
}

function Row({ nome, meta, value, tone }: { nome: string; meta: string; value: string; tone: "warn" | "ok" | "danger" }) {
  const color = tone === "warn" ? "text-status-break" : tone === "ok" ? "text-status-present" : "text-status-absent";
  return (
    <div className="flex items-center justify-between px-2 py-2 rounded-md hover:bg-secondary/50">
      <div>
        <div className="text-sm font-medium text-foreground">{nome}</div>
        <div className="text-xs text-muted-foreground">{meta}</div>
      </div>
      <div className={`text-sm font-semibold tabular-nums ${color}`}>{value}</div>
    </div>
  );
}

function Empty({ text }: { text: string }) {
  return <div className="px-3 py-6 text-center text-sm text-muted-foreground">{text}</div>;
}