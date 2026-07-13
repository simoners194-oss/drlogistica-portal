import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { AlertTriangle, ClipboardList, Lock, PenLine, PlusCircle } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetDipendenti,
  spGetAnomalie,
  spCreateTimbraturaManuale,
  spCreateTurnoManuale,
} from "@/lib/sharepoint.functions";
import type { SpDipendente, AnomaliaItem } from "@/lib/sharepoint.server";
import { EVENTI, LABEL_ANOMALIA, type EventoTimbratura } from "@/lib/presenze-logic";
import { SEDI, labelTipo, type SedeId } from "@/lib/mock-data";

export const Route = createFileRoute("/gestione-timbrature")({
  head: () => ({ meta: [{ title: "Gestione timbrature — DR Portal" }] }),
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
  component: GestioneTimbraturePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function toIso(data: string, ora: string): string {
  return new Date(`${data}T${ora}`).toISOString();
}

function sedeNome(id: string): string {
  return SEDI.find((s) => s.id === id)?.nome ?? id;
}

function fmtData(iso: string): string {
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

function GestioneTimbraturePage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [topTab, setTopTab] = useState<"inserimento" | "anomalie">("inserimento");
  const [anomalie, setAnomalie] = useState<AnomaliaItem[] | null>(null);

  const [sedeFilter, setSedeFilter] = useState<SedeId | "tutte">("tutte");
  const [dipendenteId, setDipendenteId] = useState("");
  const [mode, setMode] = useState<"singola" | "turno">("singola");
  const [data, setData] = useState("");
  const [note, setNote] = useState("");
  // Singola
  const [evento, setEvento] = useState<EventoTimbratura>("entrata");
  const [ora, setOra] = useState("");
  // Turno
  const [entrataOra, setEntrataOra] = useState("");
  const [uscitaOra, setUscitaOra] = useState("");
  const [pausaInizio, setPausaInizio] = useState("");
  const [pausaFine, setPausaFine] = useState("");

  const [submitting, setSubmitting] = useState(false);

  function loadAnomalie() {
    spGetAnomalie({ data: { giorni: 14 } })
      .then((list) => setAnomalie(list as AnomaliaItem[]))
      .catch((err) => {
        setAnomalie([]);
        toast.error("Errore nel calcolo anomalie", {
          description: err instanceof Error ? err.message : String(err),
        });
      });
  }

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    if (!s.operatore) return;
    spGetDipendenti()
      .then((list) => setDipendenti(list as SpDipendente[]))
      .catch((err) => setLoadError(err instanceof Error ? err.message : String(err)));
    loadAnomalie();
  }, []);

  const filteredDip = useMemo(() => {
    const arr = (dipendenti ?? []).filter((d) =>
      sedeFilter === "tutte" ? true : d.sede === sedeFilter,
    );
    return [...arr].sort((a, b) =>
      `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`),
    );
  }, [dipendenti, sedeFilter]);

  function resetTimes() {
    setOra("");
    setEntrataOra("");
    setUscitaOra("");
    setPausaInizio("");
    setPausaFine("");
    setNote("");
  }

  // Precompila il form dall'anomalia e porta l'operatore all'inserimento.
  function correggi(a: AnomaliaItem) {
    setTopTab("inserimento");
    setMode("singola");
    setSedeFilter("tutte");
    setDipendenteId(a.dipendenteId);
    setData(a.data);
    setEvento(a.tipo === "turno-non-chiuso" ? "uscita" : "fine-pausa");
    setOra("");
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!session || submitting) return;
    if (!dipendenteId) return toast.error("Seleziona un dipendente");
    if (!data) return toast.error("Inserisci la data");

    setSubmitting(true);
    try {
      const dip = dipendenti?.find((d) => d.id === dipendenteId);
      const chi = dip ? `${dip.cognome} ${dip.nome} · ` : "";
      if (mode === "singola") {
        if (!ora) return toast.error("Inserisci l'ora");
        await spCreateTimbraturaManuale({
          data: {
            operatoreId: session.id,
            dipendenteId,
            evento,
            dataOra: toIso(data, ora),
            note: note.trim() || undefined,
          },
        });
        toast.success("Timbratura inserita", {
          description: `${chi}${labelTipo(evento)} · ${data} ${ora}`,
        });
      } else {
        if (!entrataOra || !uscitaOra) return toast.error("Inserisci entrata e uscita");
        const conPausa = Boolean(pausaInizio || pausaFine);
        if (conPausa && (!pausaInizio || !pausaFine))
          return toast.error("Per la pausa servono sia inizio sia fine");
        const res = (await spCreateTurnoManuale({
          data: {
            operatoreId: session.id,
            dipendenteId,
            entrata: toIso(data, entrataOra),
            uscita: toIso(data, uscitaOra),
            inizioPausa: conPausa ? toIso(data, pausaInizio) : undefined,
            finePausa: conPausa ? toIso(data, pausaFine) : undefined,
            note: note.trim() || undefined,
          },
        })) as unknown[];
        toast.success("Turno inserito", {
          description: `${chi}${res.length} timbrature · ${data}`,
        });
      }
      resetTimes();
      loadAnomalie(); // la correzione può aver risolto un'anomalia
    } catch (err) {
      toast.error("Inserimento non riuscito", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  }

  if (session && !session.operatore) {
    return (
      <AppShell title="Gestione timbrature">
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">Accesso riservato</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">
              Questa sezione è riservata agli operatori abilitati.
            </p>
          </div>
        </div>
      </AppShell>
    );
  }

  const anomalieCount = anomalie?.length ?? 0;

  return (
    <AppShell title="Gestione timbrature" subtitle="Inserimento e anomalie (operatore)">
      {/* Tab: inserimento / anomalie */}
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setTopTab("inserimento")}
          className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "inserimento" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          Inserimento
        </button>
        <button
          type="button"
          onClick={() => setTopTab("anomalie")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "anomalie" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          Anomalie
          {anomalieCount > 0 && (
            <span
              className={`rounded-full px-1.5 text-[11px] ${topTab === "anomalie" ? "bg-primary-foreground/20" : "bg-status-absent/15 text-status-absent"}`}
            >
              {anomalieCount}
            </span>
          )}
        </button>
      </div>

      {topTab === "anomalie" ? (
        /* ---------------- Anomalie ---------------- */
        <div className="max-w-2xl rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <AlertTriangle className="h-4 w-4 text-status-absent" /> Anomalie (ultimi 14 giorni)
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">
            Giornate passate con turno o pausa non chiusi. Il giorno corrente è escluso.
          </p>

          {anomalie === null ? (
            <div className="text-sm text-muted-foreground">Calcolo in corso…</div>
          ) : anomalie.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessuna anomalia rilevata. Tutto in ordine. 🎉
            </div>
          ) : (
            <ul className="space-y-2">
              {anomalie.map((a, i) => (
                <li
                  key={`${a.dipendenteId}-${a.data}-${a.tipo}-${i}`}
                  className="flex items-center justify-between gap-3 rounded-xl border border-border bg-background p-3"
                >
                  <div className="min-w-0">
                    <div className="font-medium text-foreground truncate">{a.nomeCompleto}</div>
                    <div className="text-[13px] text-muted-foreground">
                      {fmtData(a.data)}
                      {a.sede ? ` · ${sedeNome(a.sede)}` : ""}
                    </div>
                    <div className="mt-0.5 text-[12px] text-status-absent">
                      {LABEL_ANOMALIA[a.tipo]}
                    </div>
                  </div>
                  <Button
                    type="button"
                    size="sm"
                    variant="ghost"
                    className="shrink-0"
                    onClick={() => correggi(a)}
                  >
                    <PenLine className="h-4 w-4" />
                    Correggi
                  </Button>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : (
        /* ---------------- Inserimento ---------------- */
        <div className="max-w-xl">
          <form
            onSubmit={handleSubmit}
            className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)] space-y-4"
          >
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
              <ClipboardList className="h-4 w-4 text-primary" /> Nuovo inserimento manuale
            </div>

            {loadError && (
              <div className="rounded-lg bg-status-absent/10 p-3 text-[13px] text-status-absent">
                Errore nel caricamento dipendenti: {loadError}
              </div>
            )}

            {/* Filtro sede + dipendente */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  Sede
                </label>
                <select
                  className={`${inputCls} mt-1`}
                  value={sedeFilter}
                  onChange={(e) => {
                    setSedeFilter(e.target.value as SedeId | "tutte");
                    setDipendenteId("");
                  }}
                >
                  <option value="tutte">Tutte le sedi</option>
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
                  value={dipendenteId}
                  onChange={(e) => setDipendenteId(e.target.value)}
                  disabled={dipendenti === null}
                >
                  <option value="">{dipendenti === null ? "Caricamento…" : "— seleziona —"}</option>
                  {filteredDip.map((d) => (
                    <option key={d.id} value={d.id}>
                      {d.cognome} {d.nome}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Modalità: singola vs turno intero */}
            <div className="inline-flex rounded-lg border border-border bg-background p-1 text-sm">
              <button
                type="button"
                onClick={() => setMode("singola")}
                className={`rounded-md px-3 py-1.5 font-medium transition-colors ${mode === "singola" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
              >
                Timbratura singola
              </button>
              <button
                type="button"
                onClick={() => setMode("turno")}
                className={`rounded-md px-3 py-1.5 font-medium transition-colors ${mode === "turno" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
              >
                Turno intero
              </button>
            </div>

            {/* Data (comune) */}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Data</label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={data}
                onChange={(e) => setData(e.target.value)}
              />
            </div>

            {mode === "singola" ? (
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Evento
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={evento}
                    onChange={(e) => setEvento(e.target.value as EventoTimbratura)}
                  >
                    {EVENTI.map((ev) => (
                      <option key={ev} value={ev}>
                        {labelTipo(ev)}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Ora
                  </label>
                  <input
                    type="time"
                    className={`${inputCls} mt-1`}
                    value={ora}
                    onChange={(e) => setOra(e.target.value)}
                  />
                </div>
              </div>
            ) : (
              <div className="space-y-3">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Entrata
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={entrataOra}
                      onChange={(e) => setEntrataOra(e.target.value)}
                    />
                  </div>
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Uscita
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={uscitaOra}
                      onChange={(e) => setUscitaOra(e.target.value)}
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Inizio pausa{" "}
                      <span className="normal-case text-muted-foreground/70">(facolt.)</span>
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={pausaInizio}
                      onChange={(e) => setPausaInizio(e.target.value)}
                    />
                  </div>
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Fine pausa{" "}
                      <span className="normal-case text-muted-foreground/70">(facolt.)</span>
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={pausaFine}
                      onChange={(e) => setPausaFine(e.target.value)}
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Note */}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Nota <span className="normal-case text-muted-foreground/70">(facoltativa)</span>
              </label>
              <textarea
                className={`${inputCls} mt-1 min-h-[60px] resize-y`}
                value={note}
                onChange={(e) => setNote(e.target.value)}
                placeholder="Motivo (es. telefono non funzionante)"
              />
            </div>

            <div className="rounded-lg bg-primary/5 p-3 text-[12px] text-muted-foreground">
              {mode === "turno"
                ? "Verranno inserite entrata, eventuale pausa e uscita, tutte con origine "
                : "La timbratura verrà registrata con origine "}
              <strong>Manuale</strong> e resa visibile al supervisore.
            </div>

            <Button type="submit" className="w-full" disabled={submitting}>
              <PlusCircle className="h-4 w-4" />
              {submitting
                ? "Inserimento…"
                : mode === "turno"
                  ? "Inserisci turno"
                  : "Inserisci timbratura"}
            </Button>
          </form>
        </div>
      )}
    </AppShell>
  );
}
