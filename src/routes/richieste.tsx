import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import {
  FileText,
  Plus,
  Send,
  XCircle,
  CalendarDays,
  Clock,
  Info,
  Inbox,
  CheckCircle2,
  Receipt,
} from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetRichieste,
  spCreateRichiesta,
  spCancelRichiesta,
  spDecideRichiesta,
  spUploadGiustificativo,
} from "@/lib/sharepoint.functions";
import type { SpRichiesta } from "@/lib/sharepoint.server";
import {
  TIPI_RICHIESTA,
  MODALITA,
  TIPI_ACQUISTO,
  misuraInGiorni,
  richiedeApprovazione,
  isRimborso,
  validateRichiesta,
  canCancel,
  parseStato,
  type TipoRichiesta,
  type ModalitaStraordinario,
  type TipoAcquisto,
  type DecisioneRichiesta,
  type RichiestaInput,
} from "@/lib/richieste-logic";

export const Route = createFileRoute("/richieste")({
  head: () => ({ meta: [{ title: "Richieste — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    // Guardia client-only, come le altre route protette.
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
  component: RichiestePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

const STATO_STYLE: Record<string, string> = {
  Bozza: "bg-muted text-muted-foreground",
  Inviata: "bg-primary/10 text-primary",
  Comunicata: "bg-primary/10 text-primary",
  Approvata: "bg-status-present/15 text-status-present",
  Respinta: "bg-status-absent/15 text-status-absent",
  Annullata: "bg-muted text-muted-foreground",
};

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

// Legge un File come data URL base64 ("data:...;base64,XXXX"). Il server ne
// estrae la parte base64 e la scrive nella libreria documenti.
function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Lettura file fallita"));
    reader.readAsDataURL(file);
  });
}

function periodoLabel(r: SpRichiesta): string {
  const range =
    r.dataFine && r.dataFine.slice(0, 10) !== r.dataInizio.slice(0, 10)
      ? `${fmtData(r.dataInizio)} → ${fmtData(r.dataFine)}`
      : fmtData(r.dataInizio);
  const ore = r.oraInizio && r.oraFine ? ` · ${r.oraInizio}–${r.oraFine}` : "";
  return range + ore;
}

function RichiestePage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [view, setView] = useState<"mie" | "coda">("mie");

  // Le mie richieste
  const [richieste, setRichieste] = useState<SpRichiesta[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [cancelingId, setCancelingId] = useState<string | null>(null);

  // Coda approvatore
  const [pending, setPending] = useState<SpRichiesta[] | null>(null);
  const [pendingError, setPendingError] = useState<string | null>(null);
  const [decidingId, setDecidingId] = useState<string | null>(null);
  const [rejectingId, setRejectingId] = useState<string | null>(null);
  const [rejectNote, setRejectNote] = useState("");

  // Stato del form
  const [tipo, setTipo] = useState<TipoRichiesta>("Ferie");
  const [dataInizio, setDataInizio] = useState("");
  const [dataFine, setDataFine] = useState("");
  const [oraInizio, setOraInizio] = useState("");
  const [oraFine, setOraFine] = useState("");
  const [motivazione, setMotivazione] = useState("");
  const [modalita, setModalita] = useState<ModalitaStraordinario | "">("");
  const [protocolloInps, setProtocolloInps] = useState("");
  const [importo, setImporto] = useState("");
  const [tipoAcquisto, setTipoAcquisto] = useState<TipoAcquisto | "">("");
  const [giustificativo, setGiustificativo] = useState("");
  const [giustFile, setGiustFile] = useState<File | null>(null);
  const [formErrors, setFormErrors] = useState<string[]>([]);
  const [submitting, setSubmitting] = useState(false);

  const isRimb = isRimborso(tipo);
  const perGiorni = misuraInGiorni(tipo);
  const motivObbligatoria = tipo === "Permesso" || tipo === "Straordinario";
  const senzaApprovazione = !richiedeApprovazione(tipo);
  const isApprovatore = session?.autorizza === true;

  async function loadRichieste(id: string) {
    try {
      const list = (await spGetRichieste({ data: { richiedenteId: id } })) as SpRichiesta[];
      setRichieste(list);
      setLoadError(null);
    } catch (err) {
      setLoadError(err instanceof Error ? err.message : String(err));
      setRichieste([]);
    }
  }

  async function loadPending() {
    try {
      const list = (await spGetRichieste({ data: { stato: "Inviata" } })) as SpRichiesta[];
      setPending(list);
      setPendingError(null);
    } catch (err) {
      setPendingError(err instanceof Error ? err.message : String(err));
      setPending([]);
    }
  }

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    void loadRichieste(s.id);
    if (s.autorizza) void loadPending();
  }, []);

  function resetForm() {
    setDataInizio("");
    setDataFine("");
    setOraInizio("");
    setOraFine("");
    setMotivazione("");
    setModalita("");
    setProtocolloInps("");
    setImporto("");
    setTipoAcquisto("");
    setGiustificativo("");
    setGiustFile(null);
    setFormErrors([]);
  }

  function buildInput(): RichiestaInput {
    if (isRimb) {
      return {
        tipo,
        dataInizio, // data acquisto
        dataFine: dataInizio,
        motivazione: motivazione.trim() || undefined,
        importo: importo ? Number(importo.replace(",", ".")) : undefined,
        tipoAcquisto: tipoAcquisto || undefined,
        giustificativo: giustificativo.trim() || undefined,
      };
    }
    return {
      tipo,
      dataInizio,
      dataFine: perGiorni ? dataFine : dataInizio,
      oraInizio: perGiorni ? undefined : oraInizio || undefined,
      oraFine: perGiorni ? undefined : oraFine || undefined,
      motivazione: motivazione.trim() || undefined,
      modalita: tipo === "Straordinario" ? modalita || undefined : undefined,
      protocolloInps: tipo === "Malattia" ? protocolloInps.trim() || undefined : undefined,
    };
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!session || submitting) return;
    const input = buildInput();
    const v = validateRichiesta(input);
    if (!v.ok) {
      setFormErrors(v.errors);
      toast.error("Controlla i campi della richiesta", { description: v.errors[0] });
      return;
    }
    setFormErrors([]);
    setSubmitting(true);
    try {
      // Rimborso con file allegato: prima carico il giustificativo nella
      // libreria documenti e uso il webUrl restituito come "Giustificativo".
      if (isRimb && giustFile) {
        if (giustFile.size > 8 * 1024 * 1024) {
          throw new Error("File troppo grande: il limite è 8 MB.");
        }
        const contentBase64 = await fileToDataUrl(giustFile);
        const up = await spUploadGiustificativo({
          data: { filename: giustFile.name, contentBase64 },
        });
        input.giustificativo = up.webUrl || input.giustificativo;
      }
      await spCreateRichiesta({ data: { richiedenteId: session.id, ...input, submit: true } });
      toast.success(senzaApprovazione ? "Comunicazione inviata" : "Richiesta inviata", {
        description: `${tipo} · ${fmtData(input.dataInizio)}`,
      });
      resetForm();
      await loadRichieste(session.id);
    } catch (err) {
      toast.error("Invio non riuscito", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  }

  async function handleCancel(r: SpRichiesta) {
    if (!session || cancelingId) return;
    setCancelingId(r.id);
    try {
      await spCancelRichiesta({ data: { richiestaId: r.id, richiedenteId: session.id } });
      toast.success(`Richiesta ${r.title || `#${r.id}`} annullata`);
      await loadRichieste(session.id);
    } catch (err) {
      toast.error("Annullamento non riuscito", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setCancelingId(null);
    }
  }

  async function decide(r: SpRichiesta, decisione: DecisioneRichiesta, note?: string) {
    if (!session || decidingId) return;
    if (decisione === "Respinta" && (!note || !note.trim())) {
      toast.error("Serve una nota per respingere la richiesta");
      return;
    }
    setDecidingId(r.id);
    try {
      await spDecideRichiesta({
        data: {
          richiestaId: r.id,
          approvatoreId: session.id,
          decisione,
          noteDecisione: note?.trim() || undefined,
        },
      });
      toast.success(decisione === "Approvata" ? "Richiesta approvata" : "Richiesta respinta");
      setRejectingId(null);
      setRejectNote("");
      await loadPending();
    } catch (err) {
      toast.error("Operazione non riuscita", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setDecidingId(null);
    }
  }

  const pendingCount = pending?.length ?? 0;

  return (
    <AppShell title="Richieste" subtitle="Ferie, permessi e giustificativi">
      {/* Tab di navigazione — solo per gli approvatori */}
      {isApprovatore && (
        <div className="mb-4 inline-flex rounded-xl border border-border bg-card p-1 text-sm shadow-[var(--shadow-card)]">
          <button
            type="button"
            onClick={() => setView("mie")}
            className={`rounded-lg px-3 py-1.5 font-medium transition-colors ${view === "mie" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
          >
            Le mie richieste
          </button>
          <button
            type="button"
            onClick={() => setView("coda")}
            className={`rounded-lg px-3 py-1.5 font-medium transition-colors inline-flex items-center gap-2 ${view === "coda" ? "bg-primary text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}
          >
            Da approvare
            {pendingCount > 0 && (
              <span
                className={`rounded-full px-1.5 text-[11px] ${view === "coda" ? "bg-primary-foreground/20" : "bg-primary/10 text-primary"}`}
              >
                {pendingCount}
              </span>
            )}
          </button>
        </div>
      )}

      {view === "coda" && isApprovatore ? (
        /* ---------------- Coda approvatore ---------------- */
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <Inbox className="h-4 w-4 text-primary" /> Richieste da approvare
          </div>

          {pendingError && (
            <div className="rounded-lg bg-status-absent/10 p-3 text-[13px] text-status-absent mb-3">
              Errore nel caricamento: {pendingError}
            </div>
          )}

          {pending === null ? (
            <div className="text-sm text-muted-foreground">Caricamento in corso…</div>
          ) : pending.length === 0 ? (
            <div className="text-sm text-muted-foreground">
              Nessuna richiesta in attesa di approvazione.
            </div>
          ) : (
            <ul className="space-y-3">
              {pending.map((r) => (
                <li key={r.id} className="rounded-xl border border-border bg-background p-4">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-semibold text-foreground">{r.tipo || "—"}</span>
                        {r.codiceRichiedente && (
                          <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] font-medium text-muted-foreground">
                            {r.codiceRichiedente}
                          </span>
                        )}
                        {r.modalita && (
                          <span className="text-[11px] text-muted-foreground">({r.modalita})</span>
                        )}
                        {r.sedeRichiedente && (
                          <span className="text-[11px] text-muted-foreground">
                            · {r.sedeRichiedente}
                          </span>
                        )}
                      </div>
                      <div className="mt-1 flex items-center gap-2 text-[13px] text-muted-foreground">
                        {r.oraInizio ? (
                          <Clock className="h-3.5 w-3.5" />
                        ) : (
                          <CalendarDays className="h-3.5 w-3.5" />
                        )}
                        <span className="truncate">{periodoLabel(r)}</span>
                      </div>
                      {r.motivazione && (
                        <div className="mt-1 text-[13px] text-foreground/80 italic">
                          “{r.motivazione}”
                        </div>
                      )}
                      {r.tipo === "Rimborso spese" && (
                        <div className="mt-1.5 flex flex-wrap items-center gap-2 text-[13px]">
                          <span className="font-medium text-foreground tabular-nums">
                            € {(r.importo ?? 0).toFixed(2)}
                          </span>
                          {r.tipoAcquisto && (
                            <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
                              {r.tipoAcquisto}
                            </span>
                          )}
                          {r.giustificativo && (
                            <a
                              href={r.giustificativo}
                              target="_blank"
                              rel="noreferrer"
                              className="inline-flex items-center gap-1 text-primary underline"
                            >
                              <Receipt className="h-3.5 w-3.5" /> Giustificativo
                            </a>
                          )}
                        </div>
                      )}
                      <div className="mt-0.5 text-[11px] text-muted-foreground/70">
                        {r.title || `#${r.id}`}
                        {r.protocolloInps ? ` · INPS ${r.protocolloInps}` : ""}
                      </div>
                    </div>
                    <div className="flex shrink-0 flex-col gap-2">
                      <Button
                        type="button"
                        size="sm"
                        className="bg-status-present hover:bg-status-present/90"
                        disabled={decidingId === r.id}
                        onClick={() => decide(r, "Approvata")}
                      >
                        <CheckCircle2 className="h-4 w-4" />
                        Approva
                      </Button>
                      <Button
                        type="button"
                        size="sm"
                        variant="ghost"
                        className="text-status-absent hover:text-status-absent"
                        disabled={decidingId === r.id}
                        onClick={() => setRejectingId((cur) => (cur === r.id ? null : r.id))}
                      >
                        <XCircle className="h-4 w-4" />
                        Respingi
                      </Button>
                    </div>
                  </div>

                  {rejectingId === r.id && (
                    <div className="mt-3 rounded-lg bg-status-absent/5 p-3">
                      <label className="text-xs uppercase tracking-wider text-muted-foreground">
                        Motivo del rifiuto <span className="normal-case">(obbligatorio)</span>
                      </label>
                      <textarea
                        className={`${inputCls} mt-1 min-h-[60px] resize-y`}
                        value={rejectNote}
                        onChange={(e) => setRejectNote(e.target.value)}
                        autoFocus
                      />
                      <div className="mt-2 flex justify-end gap-2">
                        <Button
                          type="button"
                          size="sm"
                          variant="ghost"
                          onClick={() => {
                            setRejectingId(null);
                            setRejectNote("");
                          }}
                        >
                          Annulla
                        </Button>
                        <Button
                          type="button"
                          size="sm"
                          className="bg-status-absent hover:bg-status-absent/90"
                          disabled={decidingId === r.id || !rejectNote.trim()}
                          onClick={() => decide(r, "Respinta", rejectNote)}
                        >
                          Conferma rifiuto
                        </Button>
                      </div>
                    </div>
                  )}
                </li>
              ))}
            </ul>
          )}
        </div>
      ) : (
        /* ---------------- Le mie richieste (crea + elenco) ---------------- */
        <div className="grid gap-4 md:gap-5 lg:grid-cols-5">
          {/* Nuova richiesta */}
          <form
            onSubmit={handleSubmit}
            className="lg:col-span-2 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)] space-y-4"
          >
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
              <Plus className="h-4 w-4 text-primary" /> Nuova richiesta
            </div>

            {/* Tipo */}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Tipo</label>
              <select
                className={`${inputCls} mt-1`}
                value={tipo}
                onChange={(e) => {
                  setTipo(e.target.value as TipoRichiesta);
                  setFormErrors([]);
                }}
              >
                {TIPI_RICHIESTA.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>

            {senzaApprovazione && (
              <div className="flex items-start gap-2 rounded-lg bg-primary/5 p-3 text-[13px] text-muted-foreground">
                <Info className="h-4 w-4 shrink-0 text-primary mt-0.5" />
                Questa è una <strong className="text-foreground">comunicazione</strong>: non
                richiede approvazione, viene registrata direttamente.
              </div>
            )}

            {/* Campi per tipo */}
            {isRimb ? (
              <div className="space-y-3">
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Data acquisto
                  </label>
                  <input
                    type="date"
                    className={`${inputCls} mt-1`}
                    value={dataInizio}
                    onChange={(e) => setDataInizio(e.target.value)}
                  />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Importo (€)
                    </label>
                    <input
                      type="number"
                      step="0.01"
                      min="0"
                      className={`${inputCls} mt-1`}
                      value={importo}
                      onChange={(e) => setImporto(e.target.value)}
                      placeholder="0,00"
                    />
                  </div>
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Tipologia
                    </label>
                    <select
                      className={`${inputCls} mt-1`}
                      value={tipoAcquisto}
                      onChange={(e) => setTipoAcquisto(e.target.value as TipoAcquisto | "")}
                    >
                      <option value="">— seleziona —</option>
                      {TIPI_ACQUISTO.map((t) => (
                        <option key={t} value={t}>
                          {t}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Documento giustificativo{" "}
                    <span className="normal-case text-muted-foreground/70">(file, opzionale)</span>
                  </label>
                  <input
                    type="file"
                    accept="image/*,application/pdf"
                    className={`${inputCls} mt-1 file:mr-3 file:rounded-md file:border-0 file:bg-secondary file:px-3 file:py-1 file:text-xs file:font-medium file:text-foreground`}
                    onChange={(e) => setGiustFile(e.target.files?.[0] ?? null)}
                  />
                  <p className="mt-1 text-[11px] text-muted-foreground">
                    Foto o PDF dello scontrino/fattura · max 8 MB. In alternativa incolla un link
                    qui sotto.
                  </p>
                  <input
                    type="text"
                    className={`${inputCls} mt-2`}
                    value={giustificativo}
                    onChange={(e) => setGiustificativo(e.target.value)}
                    placeholder="Oppure link al documento (opzionale)"
                  />
                </div>
              </div>
            ) : perGiorni ? (
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Dal
                  </label>
                  <input
                    type="date"
                    className={`${inputCls} mt-1`}
                    value={dataInizio}
                    onChange={(e) => setDataInizio(e.target.value)}
                  />
                </div>
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Al
                  </label>
                  <input
                    type="date"
                    className={`${inputCls} mt-1`}
                    value={dataFine}
                    onChange={(e) => setDataFine(e.target.value)}
                  />
                </div>
              </div>
            ) : (
              <div className="space-y-3">
                <div>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Giorno
                  </label>
                  <input
                    type="date"
                    className={`${inputCls} mt-1`}
                    value={dataInizio}
                    onChange={(e) => setDataInizio(e.target.value)}
                  />
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Dalle
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={oraInizio}
                      onChange={(e) => setOraInizio(e.target.value)}
                    />
                  </div>
                  <div>
                    <label className="text-xs uppercase tracking-wider text-muted-foreground">
                      Alle
                    </label>
                    <input
                      type="time"
                      className={`${inputCls} mt-1`}
                      value={oraFine}
                      onChange={(e) => setOraFine(e.target.value)}
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Modalità (solo Straordinario) */}
            {tipo === "Straordinario" && (
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  Modalità
                </label>
                <select
                  className={`${inputCls} mt-1`}
                  value={modalita}
                  onChange={(e) => setModalita(e.target.value as ModalitaStraordinario | "")}
                >
                  <option value="">— seleziona —</option>
                  {MODALITA.map((m) => (
                    <option key={m} value={m}>
                      {m}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Protocollo INPS (solo Malattia) */}
            {tipo === "Malattia" && (
              <div>
                <label className="text-xs uppercase tracking-wider text-muted-foreground">
                  Protocollo INPS{" "}
                  <span className="normal-case text-muted-foreground/70">(facoltativo)</span>
                </label>
                <input
                  type="text"
                  className={`${inputCls} mt-1`}
                  value={protocolloInps}
                  onChange={(e) => setProtocolloInps(e.target.value)}
                  placeholder="Numero di protocollo del certificato"
                />
              </div>
            )}

            {/* Motivazione */}
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Motivazione{" "}
                {motivObbligatoria ? (
                  <span className="text-status-absent normal-case">(obbligatoria)</span>
                ) : (
                  <span className="normal-case text-muted-foreground/70">(facoltativa)</span>
                )}
              </label>
              <textarea
                className={`${inputCls} mt-1 min-h-[72px] resize-y`}
                value={motivazione}
                onChange={(e) => setMotivazione(e.target.value)}
              />
            </div>

            {formErrors.length > 0 && (
              <ul className="rounded-lg bg-status-absent/10 p-3 text-[13px] text-status-absent space-y-1">
                {formErrors.map((err, i) => (
                  <li key={i}>• {err}</li>
                ))}
              </ul>
            )}

            <Button type="submit" className="w-full" disabled={submitting}>
              <Send className="h-4 w-4" />
              {submitting ? "Invio in corso…" : senzaApprovazione ? "Comunica" : "Invia richiesta"}
            </Button>
          </form>

          {/* Le mie richieste */}
          <div className="lg:col-span-3 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
              <FileText className="h-4 w-4 text-primary" /> Le mie richieste
            </div>

            {loadError && (
              <div className="rounded-lg bg-status-absent/10 p-3 text-[13px] text-status-absent mb-3">
                Errore nel caricamento: {loadError}
              </div>
            )}

            {richieste === null ? (
              <div className="text-sm text-muted-foreground">Caricamento in corso…</div>
            ) : richieste.length === 0 ? (
              <div className="text-sm text-muted-foreground">
                Nessuna richiesta ancora. Creane una dal riquadro a sinistra.
              </div>
            ) : (
              <ul className="space-y-3">
                {richieste.map((r) => {
                  const stato = String(r.stato);
                  const annullabile = canCancel(parseStato(stato));
                  return (
                    <li
                      key={r.id}
                      className="rounded-xl border border-border bg-background p-4 flex items-start justify-between gap-3"
                    >
                      <div className="min-w-0">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className="font-semibold text-foreground">{r.tipo || "—"}</span>
                          <span
                            className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${STATO_STYLE[stato] ?? "bg-muted text-muted-foreground"}`}
                          >
                            {stato}
                          </span>
                          {r.modalita && (
                            <span className="text-[11px] text-muted-foreground">
                              ({r.modalita})
                            </span>
                          )}
                        </div>
                        <div className="mt-1 flex items-center gap-2 text-[13px] text-muted-foreground">
                          {r.oraInizio ? (
                            <Clock className="h-3.5 w-3.5" />
                          ) : (
                            <CalendarDays className="h-3.5 w-3.5" />
                          )}
                          <span className="truncate">{periodoLabel(r)}</span>
                        </div>
                        {r.noteDecisione && stato === "Respinta" && (
                          <div className="mt-1 text-[12px] text-status-absent">
                            Motivo: {r.noteDecisione}
                          </div>
                        )}
                        <div className="mt-0.5 text-[11px] text-muted-foreground/70">
                          {r.title || `#${r.id}`}
                          {r.protocolloInps ? ` · INPS ${r.protocolloInps}` : ""}
                        </div>
                      </div>
                      {annullabile && (
                        <Button
                          type="button"
                          variant="ghost"
                          size="sm"
                          className="shrink-0 text-status-absent hover:text-status-absent"
                          disabled={cancelingId === r.id}
                          onClick={() => handleCancel(r)}
                        >
                          <XCircle className="h-4 w-4" />
                          {cancelingId === r.id ? "…" : "Annulla"}
                        </Button>
                      )}
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        </div>
      )}
    </AppShell>
  );
}
