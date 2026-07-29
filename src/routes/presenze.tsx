import { createFileRoute, redirect, useNavigate } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { PresenzeSkeleton } from "@/components/skeletons/PresenzeSkeleton";
import {
  LogIn,
  Coffee,
  PlayCircle,
  LogOut,
  Clock,
  Timer,
  ListChecks,
  Hourglass,
  TrendingUp,
} from "lucide-react";
import { FileText, History, User } from "lucide-react";
import { QuickAccess } from "@/components/QuickAccess";
import { formatOra, type Dipendente, type Timbratura } from "@/lib/mock-data";
import { dataService, displayStato, DISPLAY_DOT } from "@/lib/data-service";
import { readSession } from "@/lib/session";
import { useLang } from "@/lib/i18n";
import {
  aperturaTurnoCorrente,
  computeOreOggi,
  confermaRichiesta,
  formatDurata,
  MAX_TURNO_ORE,
  isTransitionAllowed,
  lastEvento,
  reasonNotAllowed,
  UNDO_TIMBRATURA_MINUTI,
  type EventoTimbratura,
} from "@/lib/presenze-logic";
import { spAnnullaUltimaTimbratura, spCreateTimbratura } from "@/lib/sharepoint.functions";
import { accoda, isErroreRete, leggiCoda, salvaCoda, svuotaCoda } from "@/lib/timbratura-offline";
import { Undo2 } from "lucide-react";

export const Route = createFileRoute("/presenze")({
  head: () => ({ meta: [{ title: "Modulo Presenze — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    // La sessione è client-only: su SSR/prerender non c'è nulla da
    // controllare, la guardia scatta solo nel browser.
    if (typeof window === "undefined") return;
    if (!readSession()) {
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
  const { t, lang } = useLang();
  const locale = lang === "it" ? "it-IT" : "en-GB";
  const [now, setNow] = useState(new Date());
  const [me, setMe] = useState<Dipendente | undefined>(undefined);
  const [errore, setErrore] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [oreSett, setOreSett] = useState<number | null>(null);
  const avvisoOrarioRef = useRef(false);
  // Timbrature in coda offline (salvate sul dispositivo, in attesa di rete).
  const [codaCount, setCodaCount] = useState(0);
  const flushingRef = useRef(false);

  useEffect(() => {
    const s = readSession();
    const currentId = s?.id ?? null;
    if (!currentId) {
      toast.error("Sessione scaduta. Effettua di nuovo l'accesso.");
      navigate({ to: "/" });
      return;
    }
    setOreSett(s?.oreSettimanali ?? null);
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

  // Avviso "monte ore giornaliero": una volta, quando il dipendente è ancora al
  // lavoro e raggiunge le ore previste per la giornata (OreSettimanali/5).
  useEffect(() => {
    if (!me || oreSett == null || oreSett <= 0) return;
    const o = computeOreOggi(me.eventiOggi ?? [], now);
    const quotaMin = (oreSett / 5) * 60;
    if (!o.chiusa && o.oreLavorateMinuti >= quotaMin && !avvisoOrarioRef.current) {
      avvisoOrarioRef.current = true;
      toast(t("presenze.dailyQuotaTitle"), {
        description: t("presenze.dailyQuotaMsg"),
        duration: 8000,
      });
    }
  }, [me, now, oreSett]);

  // Promemoria "manca l'uscita": turno ancora aperto oltre le ore previste,
  // ripetuto ogni 30 minuti finché la giornata non viene chiusa. È l'avviso
  // che il dipendente vede con la pagina aperta; la notifica push a telefono
  // spento arriverà con la sincronizzazione programmata.
  const ultimoPromemoriaRef = useRef(0);
  useEffect(() => {
    if (!me || oreSett == null || oreSett <= 0) return;
    const eventi = (me.eventiOggi ?? []).map((e) => ({ evento: e.tipo, ora: e.ora }));
    const apertura = aperturaTurnoCorrente(eventi, now);
    if (!apertura) return;
    const oreTurno = (now.getTime() - new Date(apertura).getTime()) / 3600_000;
    const soglia = Math.min(oreSett / 5 + 1, MAX_TURNO_ORE);
    if (oreTurno < soglia) return;
    if (now.getTime() - ultimoPromemoriaRef.current < 30 * 60_000) return;
    ultimoPromemoriaRef.current = now.getTime();
    toast.warning(t("presenze.mancaUscitaTitle"), {
      description: t("presenze.mancaUscitaMsg"),
      duration: 12000,
    });
  }, [me, now, oreSett]);

  // Invia la coda offline: al ritorno della rete, all'apertura della pagina e
  // periodicamente. Gli orari inviati sono quelli REALI della pressione.
  const inviaCoda = async () => {
    if (flushingRef.current || !me) return;
    if (!leggiCoda().length) return;
    flushingRef.current = true;
    try {
      const esito = await svuotaCoda(async (item) => {
        await spCreateTimbratura({
          data: {
            dipendenteId: me.id,
            evento: item.evento,
            origine: "Web",
            note: "Recuperata offline",
            dataOraClient: item.dataOra,
          },
        });
      });
      setCodaCount(esito.rimaste);
      if (esito.inviate > 0) {
        toast.success(t("presenze.offlineSent"), { description: `${esito.inviate}` });
        await dataService.getDipendenti().catch(() => null);
        const updated = await dataService.getDipendente(me.id).catch(() => undefined);
        if (updated) setMe(updated);
      }
    } finally {
      flushingRef.current = false;
    }
  };

  useEffect(() => {
    if (!me?.id) return;
    setCodaCount(leggiCoda().length);
    void inviaCoda();
    const onOnline = () => void inviaCoda();
    window.addEventListener("online", onOnline);
    const iv = setInterval(onOnline, 30_000);
    return () => {
      window.removeEventListener("online", onOnline);
      clearInterval(iv);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [me?.id]);

  // Annulla l'ultima timbratura ("tasto sbagliato"): possibile solo entro i
  // minuti di finestra; la verifica vera è comunque lato server.
  const annullaUltima = async () => {
    if (!me || busy) return;
    const ultima = me.ultimaTimbratura;
    if (!ultima) return;
    if (
      !window.confirm(
        `${t("presenze.undoConfirm")} (${t(`evento.${ultima.tipo}`)} ${formatOra(ultima.ora)})`,
      )
    )
      return;
    // L'ultima timbratura è ancora in coda offline: si toglie dal dispositivo,
    // il server non l'ha mai vista.
    const coda = leggiCoda();
    if (coda.length > 0) {
      coda.pop();
      salvaCoda(coda);
      setCodaCount(coda.length);
      const eventi = (me.eventiOggi ?? []).slice(0, -1);
      const prev = eventi[eventi.length - 1];
      setMe({ ...me, eventiOggi: eventi, ultimaTimbratura: prev });
      toast.success(t("presenze.undoDone"), {
        description: `${t(`evento.${ultima.tipo}`)} · ${formatOra(ultima.ora)}`,
      });
      return;
    }
    setBusy(true);
    try {
      await spAnnullaUltimaTimbratura();
      // Refresh dello snapshot e del proprio record (come dopo una timbratura).
      await dataService.getDipendenti();
      const updated = await dataService.getDipendente(me.id);
      if (updated) setMe(updated);
      toast.success(t("presenze.undoDone"), {
        description: `${t(`evento.${ultima.tipo}`)} · ${formatOra(ultima.ora)}`,
      });
    } catch (err) {
      toast.error(t("presenze.undoErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setBusy(false);
    }
  };

  const timbra = async (tipo: Timbratura["tipo"]) => {
    if (!me || busy) return;
    const last = lastEvento(me.eventiOggi ?? []);
    if (!isTransitionAllowed(tipo, last)) {
      toast.error(t("presenze.notAllowedNow"), {
        description: reasonNotAllowed(tipo, last) ?? t("presenze.notAllowedNow"),
      });
      return;
    }
    // Conferme informative (mai bloccanti): seconda pausa, o uscita dopo una
    // giornata lunga senza pause registrate.
    const conferma = confermaRichiesta(
      tipo,
      (me.eventiOggi ?? []).map((e) => ({ evento: e.tipo, ora: e.ora })),
      now,
    );
    if (conferma && !window.confirm(t(`presenze.ask.${conferma}`))) return;
    setBusy(true);
    try {
      const updated = await dataService.timbra(me.id, tipo);
      setMe(updated);
      toast.success(`${t("presenze.entryRecorded")} ${t(`evento.${tipo}`)}`, {
        description: `${formatOra(updated.ultimaTimbratura?.ora)} · ${t("presenze.statusLabel")} ${t(`dstato.${displayStato(updated)}`)}`,
      });
    } catch (err) {
      if (isErroreRete(err)) {
        // Niente rete: l'evento va in coda sul dispositivo con l'ora REALE
        // della pressione e la pagina si aggiorna come se fosse registrato.
        const oraLocale = new Date().toISOString();
        const coda = accoda(tipo, oraLocale);
        setCodaCount(coda.length);
        setMe({
          ...me,
          eventiOggi: [...(me.eventiOggi ?? []), { tipo, ora: oraLocale }],
          ultimaTimbratura: { tipo, ora: oraLocale },
        });
        toast(t("presenze.offlineQueuedTitle"), {
          description: t("presenze.offlineQueuedMsg"),
          duration: 8000,
        });
      } else {
        toast.error(t("presenze.entryNotSaved"), {
          description: err instanceof Error ? err.message : String(err),
        });
      }
    } finally {
      setBusy(false);
    }
  };

  if (errore) {
    return (
      <AppShell title={t("presenze.title")}>
        <div className="text-sm text-status-absent">{errore}</div>
      </AppShell>
    );
  }

  if (!me) {
    return (
      <AppShell title={t("presenze.title")} subtitle={t("common.loading")}>
        <PresenzeSkeleton />
      </AppShell>
    );
  }

  const ds = displayStato(me);
  const eventiOggi = me.eventiOggi ?? [];
  const last = lastEvento(eventiOggi);
  const ore = computeOreOggi(eventiOggi, now);
  const azioni: {
    tipo: EventoTimbratura;
    label: string;
    Icon: typeof LogIn;
    enabled: boolean;
    reason: string | null;
    tone: "primary" | "warn" | "ok" | "danger";
  }[] = // Quattro tasti: la pausa si può ripetere più volte nella giornata e
    // restano ammessi più turni entrata→uscita (turni spezzati).
    (
      [
        { tipo: "entrata", Icon: LogIn, tone: "primary" },
        { tipo: "inizio-pausa", Icon: Coffee, tone: "warn" },
        { tipo: "fine-pausa", Icon: PlayCircle, tone: "ok" },
        { tipo: "uscita", Icon: LogOut, tone: "danger" },
      ] as const
    ).map((a) => ({
      ...a,
      label: t(`evento.${a.tipo}`),
      enabled: isTransitionAllowed(a.tipo, last),
      reason: reasonNotAllowed(a.tipo, last),
    }));

  return (
    <AppShell title={t("presenze.title")} subtitle={`${me.nome} ${me.cognome} · ${me.ruolo}`}>
      <div className="grid gap-4 md:gap-5 lg:grid-cols-3">
        <div
          className="lg:col-span-2 rounded-2xl p-5 sm:p-6 text-primary-foreground shadow-[var(--shadow-elegant)]"
          style={{ background: "var(--gradient-hero)" }}
        >
          <div className="text-[13px] sm:text-sm text-white/80 capitalize">
            {now.toLocaleDateString(locale, {
              weekday: "long",
              day: "numeric",
              month: "long",
              year: "numeric",
            })}
          </div>
          <div className="mt-1 text-[3.5rem] leading-none sm:text-6xl font-semibold tabular-nums tracking-tight">
            {now.toLocaleTimeString(locale, {
              hour: "2-digit",
              minute: "2-digit",
              second: "2-digit",
            })}
          </div>
          <div className="mt-4 flex flex-wrap items-center gap-x-5 gap-y-2 text-[13px] sm:text-sm">
            <span className="inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1">
              <span
                className={`h-2.5 w-2.5 rounded-full ${DISPLAY_DOT[ds]} ring-2 ring-white/30`}
              />
              <span className="text-white/90">
                {t("presenze.statusLabel")} <strong>{t(`dstato.${ds}`)}</strong>
              </span>
            </span>
            <span className="inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1 text-white/90">
              <Timer className="h-4 w-4" /> {t("presenze.hoursToday")}{" "}
              <strong className="tabular-nums">{formatDurata(ore.oreLavorateMinuti)}</strong>
            </span>
          </div>
        </div>

        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-muted-foreground text-[13px] sm:text-sm">
            <Clock className="h-4 w-4" /> {t("presenze.lastEntry")}
          </div>
          {me.ultimaTimbratura ? (
            <div className="mt-3">
              <div className="text-lg font-semibold text-foreground">
                {t(`evento.${me.ultimaTimbratura.tipo}`)}
              </div>
              <div className="text-[2.5rem] sm:text-4xl leading-none font-semibold tabular-nums mt-1 text-primary">
                {formatOra(me.ultimaTimbratura.ora)}
              </div>
              {(() => {
                // "Tasto sbagliato": annullabile finché la finestra non scade.
                const scadenza =
                  new Date(me.ultimaTimbratura.ora).getTime() + UNDO_TIMBRATURA_MINUTI * 60_000;
                const restanteMs = scadenza - now.getTime();
                if (restanteMs <= 0) return null;
                const mm = Math.floor(restanteMs / 60_000);
                const ss = Math.floor((restanteMs % 60_000) / 1000);
                return (
                  <button
                    type="button"
                    onClick={annullaUltima}
                    disabled={busy}
                    className="mt-3 inline-flex items-center gap-2 rounded-lg border border-border px-3 py-1.5 text-[13px] font-medium text-foreground hover:bg-muted disabled:opacity-50"
                  >
                    <Undo2 className="h-4 w-4" />
                    {t("presenze.undoLast")}
                    <span className="tabular-nums text-muted-foreground">
                      {mm}:{String(ss).padStart(2, "0")}
                    </span>
                  </button>
                );
              })()}
            </div>
          ) : (
            <div className="mt-4 text-sm text-muted-foreground">{t("presenze.noneToday")}</div>
          )}
        </div>
      </div>

      {/* Riepilogo ore */}
      {/* Fuori servizio: informativo, non bloccante (si può rientrare). */}
      {ore.chiusa && (
        <div
          role="status"
          className="mt-5 md:mt-6 flex items-start gap-3 rounded-2xl border border-border bg-muted/50 p-4 sm:p-5 animate-fade-in"
        >
          <span className="h-9 w-9 shrink-0 rounded-lg bg-secondary text-muted-foreground flex items-center justify-center">
            <LogOut className="h-4 w-4" />
          </span>
          <div className="min-w-0">
            <div className="text-sm sm:text-[15px] font-semibold text-foreground">
              {t("presenze.dayClosedTitle")}
            </div>
            <p className="text-[13px] text-muted-foreground mt-0.5 leading-snug">
              {t("presenze.dayClosedMsg")}
            </p>
          </div>
        </div>
      )}

      <div className="mt-5 md:mt-6 grid gap-3 sm:gap-4 grid-cols-2 lg:grid-cols-4">
        <RiepilogoCard
          Icon={LogIn}
          label={t("evento.entrata")}
          value={ore.entrataOra ? formatOra(ore.entrataOra) : "—"}
        />
        <RiepilogoCard
          Icon={Coffee}
          label={t("presenze.totalBreak")}
          value={formatDurata(ore.pausaMinuti)}
          hint={ore.inPausa ? t("presenze.inProgress") : undefined}
        />
        <RiepilogoCard
          Icon={Hourglass}
          label={t("presenze.workedHours")}
          value={formatDurata(ore.oreLavorateMinuti)}
        />
        <RiepilogoCard
          Icon={TrendingUp}
          label={t("presenze.overtime")}
          value={ore.oltreOrarioMinuti > 0 ? `+${formatDurata(ore.oltreOrarioMinuti)}` : "—"}
          highlight={ore.oltreOrarioMinuti > 0}
        />
      </div>

      <div className="mt-5 md:mt-6 grid gap-3 sm:gap-4 grid-cols-2 lg:grid-cols-4">
        {azioni.map((a) => (
          <button
            key={a.tipo}
            disabled={!a.enabled || busy}
            onClick={() => timbra(a.tipo)}
            title={a.reason ?? undefined}
            aria-label={`${a.label}${a.reason ? ` — ${a.reason}` : ""}`}
            className={`group relative rounded-2xl border p-4 sm:p-6 text-left transition-all min-h-[156px] sm:min-h-[176px] flex flex-col justify-between touch-manipulation
              disabled:cursor-not-allowed
              ${a.enabled ? "border-border bg-card shadow-[var(--shadow-card)] hover:shadow-[var(--shadow-elegant)] hover:-translate-y-1 active:translate-y-0 active:scale-[0.98]" : "border-dashed border-border/70 bg-muted/50 opacity-70"}
            `}
          >
            <div
              className={`inline-flex h-12 w-12 sm:h-14 sm:w-14 items-center justify-center rounded-xl text-white shadow-sm ${
                !a.enabled
                  ? "bg-muted-foreground/50"
                  : a.tone === "primary"
                    ? "bg-primary"
                    : a.tone === "warn"
                      ? "bg-status-break"
                      : a.tone === "ok"
                        ? "bg-status-present"
                        : "bg-status-absent"
              }`}
            >
              <a.Icon className="h-6 w-6 sm:h-7 sm:w-7" />
            </div>
            <div>
              <div className="text-base sm:text-lg font-semibold text-foreground leading-tight">
                {a.label}
              </div>
              <div
                className={`text-[11px] sm:text-xs mt-1 leading-snug ${a.enabled ? "text-muted-foreground" : "text-muted-foreground/90"}`}
              >
                {a.enabled ? t("presenze.tapToRecord") : (a.reason ?? t("presenze.notAvailable"))}
              </div>
            </div>
          </button>
        ))}
      </div>
      <p className="mt-2 text-[11px] text-muted-foreground">{t("presenze.hintPausa")}</p>
      {codaCount > 0 && (
        <p className="mt-1 text-[12px] font-medium text-status-break">
          {codaCount} {t("presenze.offlinePending")}
        </p>
      )}

      {/* Timeline timbrature di oggi */}
      <div className="mt-5 md:mt-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex items-center gap-2 text-[13px] sm:text-sm text-muted-foreground mb-4">
          <ListChecks className="h-4 w-4" /> {t("presenze.todayEntries")}
        </div>
        {eventiOggi.length === 0 ? (
          <div className="text-sm text-muted-foreground">{t("presenze.noneToday")}</div>
        ) : (
          <ol className="relative border-l border-border ml-2 space-y-4">
            {eventiOggi.map((e, i) => (
              <li key={`${e.tipo}-${e.ora}-${i}`} className="pl-5 relative">
                <span
                  className={`absolute -left-[7px] top-1.5 h-3 w-3 rounded-full ring-2 ring-card ${dotForEvento(e.tipo)}`}
                />
                <div className="flex items-baseline justify-between gap-3">
                  <div className="text-[14px] sm:text-[15px] font-medium text-foreground">
                    {t(`evento.${e.tipo}`)}
                  </div>
                  <div className="text-[15px] sm:text-base font-semibold tabular-nums text-primary">
                    {formatOra(e.ora)}
                  </div>
                </div>
              </li>
            ))}
          </ol>
        )}
      </div>

      {/* Accesso rapido — link ai moduli disponibili per il Dipendente */}
      <QuickAccess
        items={[
          {
            label: t("presenze.quickAttendance"),
            to: "/presenze",
            Icon: Clock,
            ready: true,
            description: t("presenze.quickMyEntries"),
          },
          {
            label: t("presenze.quickRequests"),
            to: "/richieste",
            Icon: FileText,
            ready: true,
            description: t("presenze.quickRequestsDesc"),
          },
          { label: t("presenze.quickHistory"), Icon: History, ready: false },
          { label: t("presenze.quickProfile"), Icon: User, ready: false },
        ]}
      />
    </AppShell>
  );
}

function dotForEvento(t: EventoTimbratura): string {
  switch (t) {
    case "entrata":
      return "bg-primary";
    case "inizio-pausa":
      return "bg-status-break";
    case "fine-pausa":
      return "bg-status-present";
    case "uscita":
      return "bg-status-absent";
  }
}

function RiepilogoCard({
  Icon,
  label,
  value,
  hint,
  highlight,
}: {
  Icon: typeof LogIn;
  label: string;
  value: string;
  hint?: string;
  highlight?: boolean;
}) {
  return (
    <div
      className={`rounded-2xl border p-4 sm:p-5 shadow-[var(--shadow-card)] transition-all ${highlight ? "border-status-out/40 bg-status-out/5" : "border-border bg-card"}`}
    >
      <div className="flex items-center justify-between gap-2">
        <span className="text-[11px] sm:text-xs uppercase tracking-wider text-muted-foreground">
          {label}
        </span>
        <span
          className={`h-8 w-8 rounded-lg flex items-center justify-center ${highlight ? "bg-status-out/15 text-status-out" : "bg-primary/10 text-primary"}`}
        >
          <Icon className="h-4 w-4" />
        </span>
      </div>
      <div
        className={`mt-2 text-2xl sm:text-[26px] leading-none font-semibold tabular-nums tracking-tight ${highlight ? "text-status-out" : "text-foreground"}`}
      >
        {value}
      </div>
      {hint && <div className="mt-1 text-[11px] text-muted-foreground">{hint}</div>}
    </div>
  );
}
