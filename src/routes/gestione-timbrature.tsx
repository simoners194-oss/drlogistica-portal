import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import {
  AlertTriangle,
  CalendarSearch,
  ClipboardList,
  Loader2,
  Lock,
  PenLine,
  PlusCircle,
  Trash2,
} from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetCorrezioni,
  spDecideCorrezione,
  spCronTurniToken,
  spGetDipendenti,
  spGetAnomalie,
  spCreateTimbraturaManuale,
  spCreateTurnoManuale,
  spGetResocontoGiorno,
  spDeleteTimbratura,
} from "@/lib/sharepoint.functions";
import type {
  SpCorrezione,
  SpDipendente,
  AnomaliaItem,
  SpTimbratura,
  ResocontoGiornoRiga,
} from "@/lib/sharepoint.server";
import { EVENTI_ATTIVI, type EventoTimbratura } from "@/lib/presenze-logic";
import { formatOra, type SedeId } from "@/lib/mock-data";
import { useLang } from "@/lib/i18n";

export const Route = createFileRoute("/gestione-timbrature")({
  head: () => ({ meta: [{ title: "Gestione timbrature — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: GestioneTimbraturePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function toIso(data: string, ora: string): string {
  return new Date(`${data}T${ora}`).toISOString();
}

// La sede è già il suo nome reale: nessuna mappatura id→nome.
function sedeNome(id: string): string {
  return id;
}

function fmtData(iso: string): string {
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

function GestioneTimbraturePage() {
  const { t, tVal } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [topTab, setTopTab] = useState<"inserimento" | "giornata" | "anomalie" | "correzioni">(
    "inserimento",
  );
  // Coda delle correzioni chieste dai dipendenti (fase D).
  const [correzioni, setCorrezioni] = useState<SpCorrezione[] | null>(null);
  const [corrBusy, setCorrBusy] = useState<string | null>(null);
  // Indirizzo del promemoria push "manca l'uscita" (flusso schedulato).
  const [cronUrl, setCronUrl] = useState<string | null>(null);
  const [anomalie, setAnomalie] = useState<AnomaliaItem[] | null>(null);

  // Turni del giorno (resoconto per sede): sede + data → TUTTI i dipendenti
  // con i loro eventi, compresi quelli senza timbrature. Codice/dipendente
  // sono filtri client-side sull'elenco caricato.
  const [gSede, setGSede] = useState<SedeId | "tutte">("tutte");
  const [gCodice, setGCodice] = useState("");
  const [gDipId, setGDipId] = useState("");
  const [gData, setGData] = useState(() => new Date().toISOString().slice(0, 10));
  const [gReso, setGReso] = useState<ResocontoGiornoRiga[] | null>(null);
  const [gLoading, setGLoading] = useState(false);
  const [gDeleting, setGDeleting] = useState<string | null>(null);

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

  function loadCorrezioni() {
    spGetCorrezioni({ data: { tutte: true } })
      .then((l) => setCorrezioni(l as SpCorrezione[]))
      .catch(() => setCorrezioni([]));
  }

  async function decidiCorrezione(c: SpCorrezione, approvata: boolean) {
    setCorrBusy(c.id);
    try {
      await spDecideCorrezione({ data: { correzioneId: c.id, approvata } });
      toast.success(t("gt.corrFatto"));
      loadCorrezioni();
      loadAnomalie();
    } catch (err) {
      toast.error(t("gt.corrErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setCorrBusy(null);
    }
  }

  function loadAnomalie() {
    spGetAnomalie({ data: { giorni: 14 } })
      .then((list) => setAnomalie(list as AnomaliaItem[]))
      .catch((err) => {
        setAnomalie([]);
        toast.error(t("gt.anomErr"), {
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
    if (!s.operatore && s.ruolo !== "amministratore_sistema") return;
    spGetDipendenti()
      .then((list) => setDipendenti(list as SpDipendente[]))
      .catch((err) => setLoadError(err instanceof Error ? err.message : String(err)));
    loadAnomalie();
    loadCorrezioni();
  }, []);

  const corrInAttesa = useMemo(
    () => (correzioni ?? []).filter((c) => c.stato === "In attesa").length,
    [correzioni],
  );

  // Turni rimasti aperti IERI: il "giro del mattino" dell'operatore.
  const turniApertiIeri = useMemo(() => {
    const ieri = new Date();
    ieri.setDate(ieri.getDate() - 1);
    const g = ieri.toISOString().slice(0, 10);
    return (anomalie ?? []).filter((a) => a.data === g && a.tipo === "turno-non-chiuso");
  }, [anomalie]);

  const sediOptions = useMemo(() => {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const d of dipendenti ?? []) {
      const s = (d.sede ?? "").trim();
      if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
        seen.add(s.toLowerCase());
        out.push(s);
      }
    }
    return out.sort((a, b) => a.localeCompare(b));
  }, [dipendenti]);

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

  // Dipendenti della vista "Turni del giorno", filtrati per sede e per
  // codice/nome. Se il filtro riduce a UN solo dipendente, viene selezionato
  // da solo (es. digitando "DR0018").
  const gFilteredDip = useMemo(() => {
    const q = gCodice.trim().toLowerCase();
    const arr = (dipendenti ?? []).filter((d) => {
      if (gSede !== "tutte" && d.sede !== gSede) return false;
      if (!q) return true;
      return (
        d.codice.toLowerCase().includes(q) ||
        `${d.cognome} ${d.nome}`.toLowerCase().includes(q) ||
        `${d.nome} ${d.cognome}`.toLowerCase().includes(q)
      );
    });
    return [...arr].sort((a, b) =>
      `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`),
    );
  }, [dipendenti, gSede, gCodice]);

  useEffect(() => {
    if (gFilteredDip.length === 1) {
      if (gDipId !== gFilteredDip[0].id) setGDipId(gFilteredDip[0].id);
    } else if (gDipId && !gFilteredDip.some((d) => d.id === gDipId)) {
      // Il dipendente selezionato non rientra più nei filtri.
      setGDipId("");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [gFilteredDip]);

  // Carica il resoconto del giorno per la sede selezionata.
  async function caricaGiornata(giorno = gData) {
    if (!giorno) return toast.error(t("gt.needDate"));
    setGLoading(true);
    try {
      const list = (await spGetResocontoGiorno({
        data: { sede: gSede, data: giorno },
      })) as ResocontoGiornoRiga[];
      setGReso(list);
    } catch (err) {
      setGReso([]);
      toast.error(t("gt.dayLoadErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setGLoading(false);
    }
  }

  // Righe del resoconto dopo i filtri client (codice/dipendente).
  const gResoVisibile = useMemo(() => {
    let out = gReso ?? [];
    if (gDipId) out = out.filter((r) => r.dipendenteId === gDipId);
    else if (gCodice.trim()) {
      const q = gCodice.trim().toLowerCase();
      out = out.filter(
        (r) => r.codice.toLowerCase().includes(q) || r.nomeCompleto.toLowerCase().includes(q),
      );
    }
    return out;
  }, [gReso, gDipId, gCodice]);

  // Elimina una timbratura errata: lo stato del dipendente si ricalcola dagli
  // eventi rimasti, quindi i pulsanti corretti si riabilitano da soli.
  async function eliminaTimbratura(riga: ResocontoGiornoRiga, tim: SpTimbratura) {
    if (
      !window.confirm(
        `${t("gt.deleteConfirm")}\n${riga.nomeCompleto} · ${tVal("evento", tim.evento)} ${formatOra(tim.dataOra)}`,
      )
    )
      return;
    setGDeleting(tim.id);
    try {
      await spDeleteTimbratura({ data: { timbraturaId: tim.id } });
      setGReso((prev) =>
        (prev ?? []).map((r) =>
          r.dipendenteId === riga.dipendenteId
            ? { ...r, eventi: r.eventi.filter((e) => e.id !== tim.id) }
            : r,
        ),
      );
      toast.success(t("gt.deleted"), {
        description: `${tVal("evento", tim.evento)} · ${formatOra(tim.dataOra)}`,
      });
      loadAnomalie();
    } catch (err) {
      toast.error(t("gt.deleteErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setGDeleting(null);
    }
  }

  // Dalla giornata al form di inserimento, precompilato per QUEL dipendente
  // (per aggiungere gli eventi mancanti dopo aver eliminato quello sbagliato).
  function inserisciPerGiornata(dipId: string, evento_: EventoTimbratura) {
    setTopTab("inserimento");
    setMode("singola");
    setSedeFilter("tutte");
    setDipendenteId(dipId);
    setData(gData);
    setEvento(evento_);
    setOra("");
  }

  // Precompila il form dall'anomalia e porta l'operatore all'inserimento.
  function correggi(a: AnomaliaItem) {
    setTopTab("inserimento");
    setMode("singola");
    setSedeFilter("tutte");
    setDipendenteId(a.dipendenteId);
    setData(a.data);
    // La pausa non chiusa si sana con la fine pausa, il turno con l'uscita.
    setEvento(a.tipo === "pausa-non-chiusa" ? "fine-pausa" : "uscita");
    setOra("");
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!session || submitting) return;
    if (!dipendenteId) return toast.error(t("gt.selectEmployee"));
    if (!data) return toast.error(t("gt.needDate"));

    setSubmitting(true);
    try {
      const dip = dipendenti?.find((d) => d.id === dipendenteId);
      const chi = dip ? `${dip.cognome} ${dip.nome} · ` : "";
      if (mode === "singola") {
        if (!ora) return toast.error(t("gt.needTime"));
        await spCreateTimbraturaManuale({
          data: {
            operatoreId: session.id,
            dipendenteId,
            evento,
            dataOra: toIso(data, ora),
            note: note.trim() || undefined,
          },
        });
        toast.success(t("gt.entryInserted"), {
          description: `${chi}${tVal("evento", evento)} · ${data} ${ora}`,
        });
      } else {
        if (!entrataOra || !uscitaOra) return toast.error(t("gt.needInOut"));
        const conPausa = Boolean(pausaInizio || pausaFine);
        if (conPausa && (!pausaInizio || !pausaFine)) return toast.error(t("gt.needBreakBoth"));
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
        toast.success(t("gt.shiftInserted"), {
          description: `${chi}${res.length} timbrature · ${data}`,
        });
      }
      resetTimes();
      loadAnomalie(); // la correzione può aver risolto un'anomalia
    } catch (err) {
      toast.error(t("gt.insertFailed"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  }

  if (session && !session.operatore && session.ruolo !== "amministratore_sistema") {
    return (
      <AppShell title={t("gt.title")}>
        <div className="flex items-start gap-3 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <span className="h-9 w-9 shrink-0 rounded-lg bg-muted text-muted-foreground flex items-center justify-center">
            <Lock className="h-4 w-4" />
          </span>
          <div>
            <div className="text-sm font-semibold text-foreground">{t("common.restricted")}</div>
            <p className="text-[13px] text-muted-foreground mt-0.5">{t("gt.restrictedMsg")}</p>
          </div>
        </div>
      </AppShell>
    );
  }

  const anomalieCount = anomalie?.length ?? 0;

  return (
    <AppShell title={t("gt.title")} subtitle={t("gt.subtitle")}>
      {/* Riepilogo del mattino: chi IERI non ha chiuso il turno. Sta in cima
          perché è la correzione più urgente (blocca il conteggio ore). */}
      {turniApertiIeri.length > 0 && (
        <div className="mb-4 rounded-2xl border border-status-absent/40 bg-status-absent/5 p-4 sm:p-5">
          <div className="flex items-start gap-3">
            <span className="h-9 w-9 shrink-0 rounded-lg bg-status-absent/15 text-status-absent flex items-center justify-center">
              <AlertTriangle className="h-4 w-4" />
            </span>
            <div className="min-w-0">
              <div className="text-sm font-semibold text-foreground">
                {t("gt.ieriTitle")} ({turniApertiIeri.length})
              </div>
              <p className="text-[13px] text-muted-foreground mt-0.5">{t("gt.ieriMsg")}</p>
              <div className="mt-2 flex flex-wrap gap-2">
                {turniApertiIeri.map((a) => (
                  <button
                    key={`${a.dipendenteId}-${a.data}`}
                    type="button"
                    onClick={() => correggi(a)}
                    className="rounded-full border border-border bg-background px-3 py-1 text-xs text-foreground hover:bg-muted"
                  >
                    {a.nomeCompleto}
                    {a.sede ? ` · ${a.sede}` : ""}
                  </button>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Tab: inserimento / anomalie */}
      <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
        <button
          type="button"
          onClick={() => setTopTab("inserimento")}
          className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "inserimento" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          {t("gt.tabInsert")}
        </button>
        <button
          type="button"
          onClick={() => setTopTab("giornata")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "giornata" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <CalendarSearch className="h-4 w-4" /> {t("gt.tabDay")}
        </button>
        <button
          type="button"
          onClick={() => setTopTab("anomalie")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "anomalie" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          {t("gt.tabAnomalies")}
          {anomalieCount > 0 && (
            <span
              className={`rounded-full px-1.5 text-[11px] ${topTab === "anomalie" ? "bg-primary-foreground/20" : "bg-status-absent/15 text-status-absent"}`}
            >
              {anomalieCount}
            </span>
          )}
        </button>
        <button
          type="button"
          onClick={() => setTopTab("correzioni")}
          className={`inline-flex items-center gap-2 rounded-lg px-3 py-1.5 font-medium transition-colors ${topTab === "correzioni" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
        >
          <PenLine className="h-4 w-4" /> {t("gt.tabCorrezioni")}
          {corrInAttesa > 0 && (
            <span
              className={`rounded-full px-1.5 text-[11px] ${topTab === "correzioni" ? "bg-primary-foreground/20" : "bg-status-break/15 text-status-break"}`}
            >
              {corrInAttesa}
            </span>
          )}
        </button>
      </div>

      {topTab === "correzioni" ? (
        /* ---------------- Correzioni chieste dai dipendenti ---------------- */
        <div className="max-w-3xl rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-3">
            <PenLine className="h-4 w-4 text-muted-foreground" /> {t("gt.tabCorrezioni")}
          </div>
          {correzioni == null ? (
            <div className="py-6 text-center">
              <Loader2 className="h-5 w-5 animate-spin inline-block text-muted-foreground" />
            </div>
          ) : correzioni.length === 0 ? (
            <p className="py-4 text-sm text-muted-foreground">{t("gt.corrVuote")}</p>
          ) : (
            <ul className="space-y-3">
              {correzioni.map((c) => (
                <li key={c.id} className="rounded-xl border border-border p-3">
                  <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
                    <span className="text-sm font-medium text-foreground">{c.nomeDipendente}</span>
                    <span className="text-xs text-muted-foreground">
                      {c.codiceDipendente}
                      {c.sede ? ` · ${c.sede}` : ""}
                    </span>
                    <span className="text-sm tabular-nums">{fmtData(c.giorno)}</span>
                    <span
                      className={`ml-auto rounded-full px-2 py-0.5 text-[11px] font-medium ${
                        c.stato === "Approvata"
                          ? "bg-status-present/15 text-status-present"
                          : c.stato === "Respinta"
                            ? "bg-status-absent/15 text-status-absent"
                            : "bg-status-break/15 text-status-break"
                      }`}
                    >
                      {t(`mie.stato.${c.stato}`)}
                    </span>
                  </div>
                  <div className="mt-2 grid gap-1 text-[13px] sm:grid-cols-2">
                    <div>
                      <span className="text-muted-foreground">{t("gt.corrAttuali")}: </span>
                      {c.orariAttuali || "—"}
                    </div>
                    <div>
                      <span className="text-muted-foreground">{t("gt.corrProposti")}: </span>
                      <b>{c.orariProposti}</b>
                    </div>
                  </div>
                  {c.motivo && (
                    <div className="mt-1 text-[13px]">
                      <span className="text-muted-foreground">{t("gt.corrMotivo")}: </span>
                      {c.motivo}
                    </div>
                  )}
                  {c.stato === "In attesa" && (
                    <div className="mt-3 flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => void decidiCorrezione(c, true)}
                        disabled={corrBusy != null}
                        className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                      >
                        {corrBusy === c.id && <Loader2 className="h-3.5 w-3.5 animate-spin" />}
                        {t("gt.corrApprova")}
                      </button>
                      <button
                        type="button"
                        onClick={() => void decidiCorrezione(c, false)}
                        disabled={corrBusy != null}
                        className="rounded-lg border border-border px-3 py-1.5 text-xs text-foreground hover:bg-muted disabled:opacity-50"
                      >
                        {t("gt.corrRespingi")}
                      </button>
                    </div>
                  )}
                </li>
              ))}
            </ul>
          )}

          {/* Promemoria push programmato: indirizzo per il flusso. */}
          <div className="mt-5 rounded-xl border border-border p-3">
            <div className="text-sm font-medium text-foreground">{t("gt.cronTitle")}</div>
            <p className="mt-1 text-[11px] text-muted-foreground">{t("gt.cronDesc")}</p>
            {cronUrl ? (
              <code className="mt-2 block break-all rounded-lg bg-muted px-2 py-1.5 text-[11px]">
                {cronUrl}
              </code>
            ) : (
              <button
                type="button"
                onClick={() => {
                  spCronTurniToken()
                    .then((r) =>
                      setCronUrl(`${window.location.origin}/cron-turni?token=${r.token}`),
                    )
                    .catch((err) =>
                      toast.error(t("gt.corrErr"), {
                        description: err instanceof Error ? err.message : String(err),
                      }),
                    );
                }}
                className="mt-2 rounded-lg border border-border px-3 py-1.5 text-xs text-foreground hover:bg-muted"
              >
                {t("gt.cronMostra")}
              </button>
            )}
          </div>
        </div>
      ) : topTab === "anomalie" ? (
        /* ---------------- Anomalie ---------------- */
        <div className="max-w-2xl rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <AlertTriangle className="h-4 w-4 text-status-absent" /> {t("gt.anomaliesTitle")}
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">{t("gt.anomaliesDesc")}</p>

          {anomalie === null ? (
            <div className="text-sm text-muted-foreground">{t("rep.calculating")}</div>
          ) : anomalie.length === 0 ? (
            <div className="text-sm text-muted-foreground">{t("gt.anomaliesNone")}</div>
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
                      {tVal("anomalia", a.tipo)}
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
                    {t("gt.fix")}
                  </Button>
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : topTab === "giornata" ? (
        /* ------- Turni del giorno: resoconto per sede + correzione ------- */
        <div className="max-w-4xl rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <CalendarSearch className="h-4 w-4 text-primary" /> {t("gt.dayTitle")}
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">{t("gt.dayDesc")}</p>

          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 items-end">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.site")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={gSede}
                onChange={(e) => {
                  setGSede(e.target.value as SedeId | "tutte");
                  setGReso(null);
                }}
              >
                <option value="tutte">{t("common.allSites")}</option>
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("gt.code")}
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={gCodice}
                onChange={(e) => setGCodice(e.target.value)}
                placeholder="DR0018"
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("common.employee")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={gDipId}
                onChange={(e) => setGDipId(e.target.value)}
                disabled={dipendenti === null}
              >
                <option value="">
                  {dipendenti === null ? t("common.loading") : t("common.all")}
                </option>
                {gFilteredDip.map((d) => (
                  <option key={d.id} value={d.id}>
                    {d.codice ? `${d.codice} · ` : ""}
                    {d.cognome} {d.nome}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("gt.date")}
              </label>
              <input
                type="date"
                className={`${inputCls} mt-1`}
                value={gData}
                onChange={(e) => {
                  setGData(e.target.value);
                  setGReso(null);
                }}
              />
            </div>
            <Button type="button" onClick={() => void caricaGiornata()} disabled={gLoading}>
              {gLoading ? <Loader2 className="h-4 w-4 animate-spin" /> : t("gt.dayLoad")}
            </Button>
          </div>

          {gReso != null && (
            <div className="mt-5">
              {gResoVisibile.length === 0 ? (
                <div className="text-sm text-muted-foreground">{t("gt.dayEmpty")}</div>
              ) : (
                <ul className="space-y-2">
                  {gResoVisibile.map((r) => (
                    <li
                      key={r.dipendenteId}
                      className="rounded-xl border border-border bg-background p-3"
                    >
                      <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
                        <span className="font-medium text-foreground">
                          {r.codice ? `${r.codice} · ` : ""}
                          {r.nomeCompleto}
                        </span>
                        <span className="text-[12px] text-muted-foreground">{r.sede}</span>
                        {r.senzaTimbrature && (
                          <span className="rounded-full bg-status-absent/15 px-2 py-0.5 text-[11px] font-medium text-status-absent">
                            {t("gt.noEntriesBadge")}
                          </span>
                        )}
                        {r.anomalie.map((a) => (
                          <span
                            key={a}
                            className="rounded-full bg-status-break/15 px-2 py-0.5 text-[11px] font-medium text-status-break"
                          >
                            {tVal("anomalia", a)}
                          </span>
                        ))}
                      </div>
                      {r.eventi.length > 0 && (
                        <div className="mt-2 flex flex-wrap items-center gap-1.5">
                          {r.eventi.map((tim) => (
                            <span
                              key={tim.id}
                              className="inline-flex items-center gap-1.5 rounded-lg border border-border bg-card px-2 py-1 text-[12px]"
                              title={`${tim.origine || "Web"}${tim.note ? ` · ${tim.note}` : ""}`}
                            >
                              {tVal("evento", tim.evento)}
                              <b className="tabular-nums text-primary">{formatOra(tim.dataOra)}</b>
                              {tim.origine === "Manuale" && (
                                <span className="text-[10px] text-muted-foreground">(M)</span>
                              )}
                              <button
                                type="button"
                                onClick={() => void eliminaTimbratura(r, tim)}
                                disabled={gDeleting != null}
                                title={t("gt.delete")}
                                className="rounded p-0.5 text-muted-foreground hover:text-status-absent"
                              >
                                {gDeleting === tim.id ? (
                                  <Loader2 className="h-3 w-3 animate-spin" />
                                ) : (
                                  <Trash2 className="h-3 w-3" />
                                )}
                              </button>
                            </span>
                          ))}
                        </div>
                      )}
                      <div className="mt-2 flex flex-wrap items-center gap-1.5">
                        <span className="text-[11px] text-muted-foreground">
                          {t("gt.dayAddMissing")}
                        </span>
                        {EVENTI_ATTIVI.map((ev) => (
                          <button
                            key={ev}
                            type="button"
                            onClick={() => inserisciPerGiornata(r.dipendenteId, ev)}
                            className="inline-flex items-center gap-1 rounded-md border border-border px-1.5 py-0.5 text-[11px] text-muted-foreground hover:text-foreground hover:bg-muted"
                          >
                            <PlusCircle className="h-3 w-3" /> {tVal("evento", ev)}
                          </button>
                        ))}
                      </div>
                    </li>
                  ))}
                </ul>
              )}
              <div className="mt-3 rounded-lg bg-primary/5 p-3 text-[12px] text-muted-foreground">
                {t("gt.dayHint")}
              </div>
            </div>
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
                {t("gt.loadEmpErr")} {loadError}
              </div>
            )}

            {/* Filtro sede + dipendente */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  {t("common.site")}
                </label>
                <select
                  className={`${inputCls} mt-1`}
                  value={sedeFilter}
                  onChange={(e) => {
                    setSedeFilter(e.target.value as SedeId | "tutte");
                    setDipendenteId("");
                  }}
                >
                  <option value="tutte">{t("common.allSites")}</option>
                  {sediOptions.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  {t("common.employee")}
                </label>
                <select
                  className={`${inputCls} mt-1`}
                  value={dipendenteId}
                  onChange={(e) => setDipendenteId(e.target.value)}
                  disabled={dipendenti === null}
                >
                  <option value="">
                    {dipendenti === null ? t("common.loading") : t("common.select")}
                  </option>
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
                {t("gt.singleEntry")}
              </button>
              <button
                type="button"
                onClick={() => setMode("turno")}
                className={`rounded-md px-3 py-1.5 font-medium transition-colors ${mode === "turno" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
              >
                {t("gt.fullShift")}
              </button>
            </div>

            {/* Data (comune) */}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("gt.date")}
              </label>
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
                    {t("gt.event")}
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={evento}
                    onChange={(e) => setEvento(e.target.value as EventoTimbratura)}
                  >
                    {EVENTI_ATTIVI.map((ev) => (
                      <option key={ev} value={ev}>
                        {tVal("evento", ev)}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("gt.time")}
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
                      {t("gt.entryTime")}
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
                      {t("gt.exitTime")}
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
                      {t("gt.breakStart")}{" "}
                      <span className="normal-case text-muted-foreground/70">
                        {t("gt.optShort")}
                      </span>
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
                      {t("gt.breakEnd")}{" "}
                      <span className="normal-case text-muted-foreground/70">
                        {t("gt.optShort")}
                      </span>
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
                {t("gt.note")}{" "}
                <span className="normal-case text-muted-foreground/70">{t("common.optional")}</span>
              </label>
              <textarea
                className={`${inputCls} mt-1 min-h-[60px] resize-y`}
                value={note}
                onChange={(e) => setNote(e.target.value)}
                placeholder={t("gt.notePh")}
              />
            </div>

            <div className="rounded-lg bg-primary/5 p-3 text-[12px] text-muted-foreground">
              {mode === "turno" ? t("gt.shiftHint") : t("gt.singleHint")}
              <strong>Manuale</strong> e resa visibile al supervisore.
            </div>

            <Button type="submit" className="w-full" disabled={submitting}>
              <PlusCircle className="h-4 w-4" />
              {submitting
                ? t("gt.inserting")
                : mode === "turno"
                  ? t("gt.insertShift")
                  : t("gt.insertEntry")}
            </Button>
          </form>
        </div>
      )}
    </AppShell>
  );
}
