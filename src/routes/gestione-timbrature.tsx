import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { ClipboardList, Lock, PlusCircle } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import { spGetDipendenti, spCreateTimbraturaManuale } from "@/lib/sharepoint.functions";
import type { SpDipendente } from "@/lib/sharepoint.server";
import { EVENTI, type EventoTimbratura } from "@/lib/presenze-logic";
import { labelTipo } from "@/lib/mock-data";

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

function GestioneTimbraturePage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [dipendenteId, setDipendenteId] = useState("");
  const [evento, setEvento] = useState<EventoTimbratura>("entrata");
  const [data, setData] = useState("");
  const [ora, setOra] = useState("");
  const [note, setNote] = useState("");
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    if (!s.operatore) return; // non operatore: non carichiamo nulla
    spGetDipendenti()
      .then((list) => {
        const arr = list as SpDipendente[];
        arr.sort((a, b) => `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`));
        setDipendenti(arr);
      })
      .catch((err) => setLoadError(err instanceof Error ? err.message : String(err)));
  }, []);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!session || submitting) return;
    if (!dipendenteId) {
      toast.error("Seleziona un dipendente");
      return;
    }
    if (!data || !ora) {
      toast.error("Inserisci data e ora");
      return;
    }
    const dataOra = new Date(`${data}T${ora}`);
    if (Number.isNaN(dataOra.getTime())) {
      toast.error("Data/ora non valida");
      return;
    }
    setSubmitting(true);
    try {
      await spCreateTimbraturaManuale({
        data: {
          operatoreId: session.id,
          dipendenteId,
          evento,
          dataOra: dataOra.toISOString(),
          note: note.trim() || undefined,
        },
      });
      const dip = dipendenti?.find((d) => d.id === dipendenteId);
      toast.success("Timbratura manuale inserita", {
        description: `${dip ? `${dip.cognome} ${dip.nome} · ` : ""}${labelTipo(evento)} · ${data} ${ora}`,
      });
      setData("");
      setOra("");
      setNote("");
    } catch (err) {
      toast.error("Inserimento non riuscito", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  }

  // Accesso riservato agli operatori (gating server-side comunque presente).
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

  return (
    <AppShell title="Gestione timbrature" subtitle="Inserimento manuale (operatore)">
      <div className="max-w-xl">
        <form
          onSubmit={handleSubmit}
          className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)] space-y-4"
        >
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
            <ClipboardList className="h-4 w-4 text-primary" /> Nuova timbratura manuale
          </div>

          {loadError && (
            <div className="rounded-lg bg-status-absent/10 p-3 text-[13px] text-status-absent">
              Errore nel caricamento dipendenti: {loadError}
            </div>
          )}

          {/* Dipendente */}
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
              {dipendenti?.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.cognome} {d.nome}
                </option>
              ))}
            </select>
          </div>

          {/* Evento */}
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">Evento</label>
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

          {/* Data + ora */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Data</label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={data}
                onChange={(e) => setData(e.target.value)}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Ora</label>
              <input
                type="time"
                className={`${inputCls} mt-1`}
                value={ora}
                onChange={(e) => setOra(e.target.value)}
              />
            </div>
          </div>

          {/* Note */}
          <div>
            <label className="text-xs uppercase tracking-wider text-muted-foreground">
              Nota <span className="normal-case text-muted-foreground/70">(facoltativa)</span>
            </label>
            <textarea
              className={`${inputCls} mt-1 min-h-[60px] resize-y`}
              value={note}
              onChange={(e) => setNote(e.target.value)}
              placeholder="Motivo della correzione / inserimento"
            />
          </div>

          <div className="rounded-lg bg-primary/5 p-3 text-[12px] text-muted-foreground">
            La timbratura verrà registrata con origine <strong>Manuale</strong> e resa visibile al
            supervisore.
          </div>

          <Button type="submit" className="w-full" disabled={submitting}>
            <PlusCircle className="h-4 w-4" />
            {submitting ? "Inserimento…" : "Inserisci timbratura"}
          </Button>
        </form>
      </div>
    </AppShell>
  );
}
