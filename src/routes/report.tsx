import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { BarChart3, Lock, AlertTriangle, Download, CalendarDays } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetRendiconto,
  spGetRendicontoPeriodo,
  spGetSaldoFerie,
} from "@/lib/sharepoint.functions";
import type { RendicontoRiga, SaldoFerieRiga } from "@/lib/sharepoint.server";
import { type SedeId } from "@/lib/mock-data";

export const Route = createFileRoute("/report")({
  head: () => ({ meta: [{ title: "Rendiconto — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: RendicontoPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

// La sede è già il suo nome reale: nessuna mappatura id→nome.
function sedeNome(id: string): string {
  return id;
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

// --- Settimane (lun-dom) -----------------------------------------------------
function ymdLocal(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function addDays(d: Date, n: number): Date {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}
function mondayOf(d: Date): Date {
  const x = new Date(d);
  const dow = x.getDay() === 0 ? 7 : x.getDay();
  x.setDate(x.getDate() - (dow - 1));
  return x;
}
// Lunedì della settimana fiscale (ISO): la settimana 1 è quella che contiene
// il 4 gennaio.
function mondayOfFiscalWeek(anno: number, week: number): Date {
  return addDays(mondayOf(new Date(anno, 0, 4)), (week - 1) * 7);
}
// Lunedì della settimana del mese (lun-dom): la week1 è quella che contiene il
// giorno 1 del mese; le week ripartono da 1 ogni mese.
function mondayOfMonthWeek(anno: number, mese: number, week: number): Date {
  return addDays(mondayOf(new Date(anno, mese - 1, 1)), (week - 1) * 7);
}

// Esporta il rendiconto in CSV (separatore ";", BOM UTF-8) → apribile in Excel.
function esportaCsv(righe: RendicontoRiga[], periodo: string): void {
  const header = [
    "Dipendente",
    "Sede",
    "Ore lavorate",
    "Straordinario calcolato",
    "Straordinario autorizzato",
    "Permessi (ore)",
    "Ferie (giorni)",
    "Malattie (giorni)",
    "Giorni non chiusi",
  ];
  const esc = (v: string | number): string => {
    const s = String(v ?? "");
    return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const body = righe.map((r) =>
    [
      r.nomeCompleto,
      r.sede,
      r.oreLavorate,
      r.straordinarioCalcolato,
      r.straordinarioAutorizzato,
      r.permessiOre,
      r.ferieGiorni,
      r.malattiaGiorni,
      r.giorniNonChiusi,
    ]
      .map(esc)
      .join(";"),
  );
  const csv = [header.join(";"), ...body].join("\r\n");
  const blob = new Blob([String.fromCharCode(0xfeff) + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `rendiconto-${periodo}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function RendicontoPage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [periodo, setPeriodo] = useState<string>(currentPeriodo());
  const [righe, setRighe] = useState<RendicontoRiga[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [sedeF, setSedeF] = useState<SedeId | "tutte">("tutte");
  const [dipF, setDipF] = useState("");
  const [vista, setVista] = useState<"rendiconto" | "ferie">("rendiconto");
  // Granularità del periodo: mese solare, settimana fiscale (dell'anno) o
  // settimana del mese (lun-dom, riparte da week1 ogni mese).
  const [periodoModo, setPeriodoModo] = useState<"mese" | "fiscal" | "mensile">("mese");
  const [weekNum, setWeekNum] = useState(1);
  const [saldo, setSaldo] = useState<SaldoFerieRiga[] | null>(null);
  const [saldoLoading, setSaldoLoading] = useState(false);

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

  // Intervallo effettivo del periodo selezionato (per le viste settimanali).
  const rangeSettimana = useMemo(() => {
    const anno = Number(periodo.slice(0, 4));
    const mese = Number(periodo.slice(5, 7));
    if (!anno || !mese) return null;
    if (periodoModo === "fiscal") {
      const lun = mondayOfFiscalWeek(anno, weekNum);
      return { from: ymdLocal(lun), to: ymdLocal(addDays(lun, 6)) };
    }
    if (periodoModo === "mensile") {
      const lun = mondayOfMonthWeek(anno, mese, weekNum);
      return { from: ymdLocal(lun), to: ymdLocal(addDays(lun, 6)) };
    }
    return null;
  }, [periodo, periodoModo, weekNum]);

  useEffect(() => {
    if (!canView) return;
    const anno = Number(periodo.slice(0, 4));
    const mese = Number(periodo.slice(5, 7));
    if (!anno || !mese) return;
    setLoading(true);
    setRighe(null);
    const req =
      periodoModo === "mese" || !rangeSettimana
        ? spGetRendiconto({ data: { anno, mese } })
        : spGetRendicontoPeriodo({ data: rangeSettimana });
    req
      .then((l) => setRighe(l as RendicontoRiga[]))
      .catch((err) => {
        setRighe([]);
        toast.error("Errore rendiconto", {
          description: err instanceof Error ? err.message : String(err),
        });
      })
      .finally(() => setLoading(false));
  }, [periodo, periodoModo, rangeSettimana, canView]);

  useEffect(() => {
    if (!canView || vista !== "ferie") return;
    const anno = Number(periodo.slice(0, 4));
    if (!anno) return;
    setSaldoLoading(true);
    spGetSaldoFerie({ data: { anno } })
      .then((l) => setSaldo(l as SaldoFerieRiga[]))
      .catch((err) => {
        setSaldo([]);
        toast.error("Errore saldo ferie", {
          description: err instanceof Error ? err.message : String(err),
        });
      })
      .finally(() => setSaldoLoading(false));
  }, [vista, periodo, canView]);

  const saldoFiltrato = useMemo(() => {
    return (saldo ?? []).filter((r) => {
      if (sedeF !== "tutte" && r.sede !== sedeF) return false;
      if (dipF && r.dipendenteId !== dipF) return false;
      return true;
    });
  }, [saldo, sedeF, dipF]);

  const sediOptions = useMemo(() => {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const r of [...(righe ?? []), ...(saldo ?? [])]) {
      const s = (r.sede ?? "").trim();
      if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
        seen.add(s.toLowerCase());
        out.push(s);
      }
    }
    return out.sort((a, b) => a.localeCompare(b));
  }, [righe, saldo]);

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
    <AppShell title="Rendiconto" subtitle="Ore mensili e saldo ferie">
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setVista("rendiconto")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${vista === "rendiconto" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <BarChart3 className="h-4 w-4" /> Rendiconto
        </button>
        <button
          type="button"
          onClick={() => setVista("ferie")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${vista === "ferie" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <CalendarDays className="h-4 w-4" /> Saldo ferie
        </button>
      </div>

      {vista === "rendiconto" && (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <BarChart3 className="h-4 w-4 text-primary" /> Rendiconto mensile
          </div>

          {/* Filtri */}
          <div className="grid gap-3 sm:grid-cols-4 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Periodo
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={periodoModo}
                onChange={(e) => {
                  setPeriodoModo(e.target.value as "mese" | "fiscal" | "mensile");
                  setWeekNum(1);
                }}
              >
                <option value="mese">Mese</option>
                <option value="fiscal">Settimana fiscale (anno)</option>
                <option value="mensile">Settimana del mese</option>
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {periodoModo === "fiscal" ? "Anno (dal mese)" : "Mese"}
              </label>
              <input
                type="month"
                className={`${inputCls} mt-1`}
                value={periodo}
                onChange={(e) => setPeriodo(e.target.value)}
              />
            </div>
            {periodoModo !== "mese" && (
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  {periodoModo === "fiscal" ? "Week fiscale (1-53)" : "Week del mese (1-6)"}
                </label>
                <input
                  type="number"
                  min={1}
                  max={periodoModo === "fiscal" ? 53 : 6}
                  className={`${inputCls} mt-1`}
                  value={weekNum}
                  onChange={(e) => setWeekNum(Math.max(1, Number(e.target.value) || 1))}
                />
                {rangeSettimana && (
                  <p className="mt-1 text-[11px] text-muted-foreground">
                    dal {rangeSettimana.from.split("-").reverse().join("/")} al{" "}
                    {rangeSettimana.to.split("-").reverse().join("/")}
                  </p>
                )}
              </div>
            )}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Sede</label>
              <select
                className={`${inputCls} mt-1`}
                value={sedeF}
                onChange={(e) => setSedeF(e.target.value as SedeId | "tutte")}
              >
                <option value="tutte">Tutte</option>
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
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

          {filtrate.length > 0 && (
            <div className="mb-3 flex justify-end">
              <button
                type="button"
                onClick={() => esportaCsv(filtrate, periodo)}
                className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors"
              >
                <Download className="h-4 w-4" /> Esporta CSV
              </button>
            </div>
          )}

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
            ferie/malattia ({"OreSettimanali/5"} per giorno) e dai permessi. Le settimane a cavallo
            di due mesi contano nel mese del loro lunedì. ⚠️ = giorni con turno non chiuso (ore non
            conteggiate: correggere nelle Anomalie).
          </div>
        </div>
      )}

      {vista === "ferie" && (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <CalendarDays className="h-4 w-4 text-primary" /> Saldo ferie {periodo.slice(0, 4)}
          </div>

          <div className="grid gap-3 sm:grid-cols-3 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Anno</label>
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
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
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
                {(saldo ?? []).map((r) => (
                  <option key={r.dipendenteId} value={r.dipendenteId}>
                    {r.nomeCompleto}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {saldoLoading || saldo === null ? (
            <div className="text-sm text-muted-foreground">Calcolo in corso…</div>
          ) : saldoFiltrato.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessun dato per i filtri selezionati.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">Dipendente</th>
                    <th className="py-2 pr-3">Sede</th>
                    <th className="py-2 pr-3 text-right">Spettanti</th>
                    <th className="py-2 pr-3 text-right">Godute</th>
                    <th className="py-2 pr-3 text-right">Residue</th>
                  </tr>
                </thead>
                <tbody>
                  {saldoFiltrato.map((r) => (
                    <tr key={r.dipendenteId} className="border-b border-border/60">
                      <td className="py-2 pr-3 text-foreground">{r.nomeCompleto}</td>
                      <td className="py-2 pr-3 text-muted-foreground">{sedeNome(r.sede)}</td>
                      <td className="py-2 pr-3 text-right tabular-nums">{r.spettanti}</td>
                      <td className="py-2 pr-3 text-right tabular-nums">{r.godute}</td>
                      <td
                        className={`py-2 pr-3 text-right tabular-nums font-semibold ${r.residui < 0 ? "text-status-absent" : "text-foreground"}`}
                      >
                        {r.residui}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className="mt-3 text-[12px] text-muted-foreground">
                <strong>Spettanti</strong> dalla colonna GiorniFerieAnnui del dipendente (default 26
                se non impostata). <strong>Godute</strong> = giorni di Ferie approvate nell'anno.
              </div>
            </div>
          )}
        </div>
      )}
    </AppShell>
  );
}
