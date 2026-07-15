import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import {
  ShieldCheck,
  Lock,
  CheckCircle2,
  ClipboardList,
  CalendarDays,
  Clock,
  Receipt,
} from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetRichieste,
  spGetDipendenti,
  spGetTimbratureManuali,
} from "@/lib/sharepoint.functions";
import type { SpRichiesta, SpDipendente, TimbraturaManualeItem } from "@/lib/sharepoint.server";
import { labelTipo, type SedeId } from "@/lib/mock-data";

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
// La sede è già il suo nome reale: nessuna mappatura id→nome.
function sedeNome(id: string): string {
  return id;
}

// Elenco sedi distinte presenti nei dati (richieste + dipendenti), ordinato.
function sediDistinte(nomi: (string | undefined)[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const n of nomi) {
    const s = (n ?? "").trim();
    if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
      seen.add(s.toLowerCase());
      out.push(s);
    }
  }
  return out.sort((a, b) => a.localeCompare(b));
}

const STATO_BADGE: Record<string, string> = {
  Approvata: "bg-status-present/15 text-status-present",
  Respinta: "bg-status-absent/15 text-status-absent",
};
function StatoBadge({ stato }: { stato: string }) {
  return (
    <span
      className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${STATO_BADGE[stato] ?? "bg-muted text-muted-foreground"}`}
    >
      {stato}
    </span>
  );
}

function SupervisionePage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [tab, setTab] = useState<"approvate" | "rimborsi" | "manuali">("approvate");

  const [decise, setDecise] = useState<SpRichiesta[] | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[]>([]);
  const [manuali, setManuali] = useState<TimbraturaManualeItem[] | null>(null);

  // Filtri report richieste decise (approvate + rifiutate)
  const [sedeF, setSedeF] = useState<SedeId | "tutte">("tutte");
  const [dipF, setDipF] = useState("");
  const [statoF, setStatoF] = useState<"tutte" | "Approvata" | "Respinta">("tutte");
  const [dal, setDal] = useState("");
  const [al, setAl] = useState("");

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    const puoVedere =
      s.autorizza ||
      s.operatore ||
      s.ruolo === "amministratore_sistema" ||
      s.ruolo === "responsabile";
    if (!puoVedere) return;
    Promise.all([
      spGetRichieste({ data: { stato: "Approvata" } }),
      spGetRichieste({ data: { stato: "Respinta" } }),
    ])
      .then(([a, r]) => setDecise([...(a as SpRichiesta[]), ...(r as SpRichiesta[])]))
      .catch((err) => {
        setDecise([]);
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

  const sediOptions = useMemo(
    () =>
      sediDistinte([
        ...(decise ?? []).map((r) => r.sedeRichiedente),
        ...dipendenti.map((d) => d.sede),
      ]),
    [decise, dipendenti],
  );

  const filtrate = useMemo(() => {
    return (decise ?? []).filter((r) => {
      if (statoF !== "tutte" && r.stato !== statoF) return false;
      if (sedeF !== "tutte" && r.sedeRichiedente !== sedeNome(sedeF)) return false;
      if (dipF && r.richiedenteId !== dipF) return false;
      const d = r.dataInizio.slice(0, 10);
      if (dal && d < dal) return false;
      if (al && d > al) return false;
      return true;
    });
  }, [decise, statoF, sedeF, dipF, dal, al]);

  const rimborsi = useMemo(() => filtrate.filter((r) => r.tipo === "Rimborso spese"), [filtrate]);
  // Totale sui soli rimborsi APPROVATI (i respinti non concorrono alla spesa).
  const totaleImporto = useMemo(
    () => rimborsi.reduce((s, r) => s + (r.stato === "Approvata" ? (r.importo ?? 0) : 0), 0),
    [rimborsi],
  );

  const puoVedere =
    session != null &&
    (session.autorizza ||
      session.operatore ||
      session.ruolo === "amministratore_sistema" ||
      session.ruolo === "responsabile");

  if (session && !puoVedere) {
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
          <CheckCircle2 className="h-4 w-4" /> Richieste decise
        </button>
        <button
          type="button"
          onClick={() => setTab("manuali")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "manuali" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <ClipboardList className="h-4 w-4" /> Timbrature manuali
        </button>
        <button
          type="button"
          onClick={() => setTab("rimborsi")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${tab === "rimborsi" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <Receipt className="h-4 w-4" /> Rimborsi
        </button>
      </div>

      {tab === "approvate" ? (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <ShieldCheck className="h-4 w-4 text-primary" /> Richieste decise (approvate e
            rifiutate)
          </div>

          {/* Filtri */}
          <div className="grid gap-3 sm:grid-cols-5 mb-4">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Stato
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={statoF}
                onChange={(e) => setStatoF(e.target.value as "tutte" | "Approvata" | "Respinta")}
              >
                <option value="tutte">Tutte</option>
                <option value="Approvata">Approvate</option>
                <option value="Respinta">Rifiutate</option>
              </select>
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

          {decise === null ? (
            <div className="text-sm text-muted-foreground">Caricamento…</div>
          ) : filtrate.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessuna richiesta decisa con questi filtri.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">Richiesta</th>
                    <th className="py-2 pr-3">Stato</th>
                    <th className="py-2 pr-3">Dipendente</th>
                    <th className="py-2 pr-3">Tipo</th>
                    <th className="py-2 pr-3">Periodo</th>
                    <th className="py-2 pr-3">Sede</th>
                    <th className="py-2 pr-3">Doc.</th>
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
                        <td className="py-2 pr-3">
                          <StatoBadge stato={r.stato} />
                        </td>
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
                        <td className="py-2 pr-3">
                          {r.giustificativo ? (
                            <a
                              href={r.giustificativo}
                              target="_blank"
                              rel="noreferrer"
                              className="text-primary underline"
                            >
                              apri
                            </a>
                          ) : (
                            "—"
                          )}
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
      ) : tab === "rimborsi" ? (
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <Receipt className="h-4 w-4 text-primary" /> Rimborsi spese (decisi)
          </div>

          <div className="grid gap-3 sm:grid-cols-4 mb-4">
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

          {decise === null ? (
            <div className="text-sm text-muted-foreground">Caricamento…</div>
          ) : rimborsi.length === 0 ? (
            <div className="text-sm text-muted-foreground">Nessun rimborso con questi filtri.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-[11px] uppercase tracking-wider text-muted-foreground border-b border-border">
                    <th className="py-2 pr-3">Dipendente</th>
                    <th className="py-2 pr-3">Stato</th>
                    <th className="py-2 pr-3">Sede</th>
                    <th className="py-2 pr-3">Data</th>
                    <th className="py-2 pr-3">Tipologia</th>
                    <th className="py-2 pr-3 text-right">Importo</th>
                    <th className="py-2 pr-3">Doc.</th>
                  </tr>
                </thead>
                <tbody>
                  {rimborsi.map((r) => (
                    <tr key={r.id} className="border-b border-border/60">
                      <td className="py-2 pr-3 text-foreground">
                        {nomeById.get(r.richiedenteId) ||
                          r.codiceRichiedente ||
                          `#${r.richiedenteId}`}
                      </td>
                      <td className="py-2 pr-3">
                        <StatoBadge stato={r.stato} />
                      </td>
                      <td className="py-2 pr-3 text-muted-foreground">
                        {r.sedeRichiedente || "—"}
                      </td>
                      <td className="py-2 pr-3 whitespace-nowrap">{fmtData(r.dataInizio)}</td>
                      <td className="py-2 pr-3">{r.tipoAcquisto || "—"}</td>
                      <td className="py-2 pr-3 text-right tabular-nums">
                        € {(r.importo ?? 0).toFixed(2)}
                      </td>
                      <td className="py-2 pr-3">
                        {r.giustificativo ? (
                          <a
                            href={r.giustificativo}
                            target="_blank"
                            rel="noreferrer"
                            className="text-primary underline"
                          >
                            apri
                          </a>
                        ) : (
                          "—"
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="border-t border-border font-semibold">
                    <td className="py-2 pr-3" colSpan={5}>
                      Totale approvati ({rimborsi.filter((r) => r.stato === "Approvata").length})
                    </td>
                    <td className="py-2 pr-3 text-right tabular-nums">
                      € {totaleImporto.toFixed(2)}
                    </td>
                    <td />
                  </tr>
                </tfoot>
              </table>
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
