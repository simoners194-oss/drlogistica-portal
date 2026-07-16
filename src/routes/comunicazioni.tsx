import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import {
  Megaphone,
  Send,
  Loader2,
  CheckCircle2,
  Eye,
  Users,
  CalendarDays,
  Paperclip,
  BellRing,
  Download,
} from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { checkPushSupport, enablePushNotifications, pushPermission } from "@/lib/push-client";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetComunicazioni,
  spCreateComunicazione,
  spGetMiePreseVisione,
  spMarkPresaVisione,
  spGetPreseVisione,
  spUploadFile,
  spGetDipendenti,
} from "@/lib/sharepoint.functions";
import type { SpComunicazione, SpDipendente, SpPresaVisione } from "@/lib/sharepoint.server";

export const Route = createFileRoute("/comunicazioni")({
  head: () => ({ meta: [{ title: "Comunicazioni — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: ComunicazioniPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function fmtDataOra(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  const t = iso.slice(11, 16);
  const d = y && m && g ? `${g}/${m}/${y}` : iso;
  return t ? `${d} ${t}` : d;
}
function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Lettura file fallita"));
    reader.readAsDataURL(file);
  });
}

function ComunicazioniPage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [comunicazioni, setComunicazioni] = useState<SpComunicazione[] | null>(null);
  const [mieViste, setMieViste] = useState<string[]>([]);
  const [dipendenti, setDipendenti] = useState<SpDipendente[]>([]);
  const [confermando, setConfermando] = useState<string | null>(null);

  // Chi ha letto (per pubblicatori): mappa comunicazioneId -> lista prese visione
  const [letture, setLetture] = useState<Record<string, SpPresaVisione[]>>({});

  // Form
  const [titolo, setTitolo] = useState("");
  const [testo, setTesto] = useState("");
  const [tipo, setTipo] = useState<"Comunicazione" | "Riunione">("Comunicazione");
  const [sede, setSede] = useState("Tutte");
  const [richiedePresaVisione, setRichiedePresaVisione] = useState(false);
  const [allegato, setAllegato] = useState<File | null>(null);
  const [destEmail, setDestEmail] = useState("");
  const [submitting, setSubmitting] = useState(false);

  // Banner attivazione notifiche push (client-only, dopo il mount).
  const [pushBanner, setPushBanner] = useState(false);
  const [pushBusy, setPushBusy] = useState(false);

  const canPubblicare =
    session != null &&
    (session.ruolo === "responsabile" ||
      session.ruolo === "amministratore_sistema" ||
      session.operatore);

  const load = () => {
    spGetComunicazioni()
      .then((l) => setComunicazioni(l as SpComunicazione[]))
      .catch((err) => {
        setComunicazioni([]);
        toast.error("Errore comunicazioni", {
          description: err instanceof Error ? err.message : String(err),
        });
      });
    spGetMiePreseVisione()
      .then((l) => setMieViste(l as string[]))
      .catch(() => {});
  };

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    load();
    if (s.ruolo === "responsabile" || s.ruolo === "amministratore_sistema" || s.operatore) {
      spGetDipendenti()
        .then((l) => setDipendenti(l as SpDipendente[]))
        .catch(() => {});
    }
    // Mostra il banner push se il dispositivo può riceverle e non sono attive
    // (o se su iOS serve prima installare la PWA). Con permesso già concesso
    // verifica che esista DAVVERO una subscription: se la prima attivazione è
    // fallita a metà, il banner deve ricomparire per riprovare.
    const support = checkPushSupport();
    if (support === "ios-not-installed") setPushBanner(true);
    else if (support === "ok") {
      if (pushPermission() !== "granted") setPushBanner(true);
      else {
        navigator.serviceWorker
          .getRegistration("/sw.js")
          .then((reg) => reg?.pushManager.getSubscription() ?? null)
          .then((sub) => {
            if (!sub) setPushBanner(true);
          })
          .catch(() => setPushBanner(true));
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const attivaPush = async () => {
    setPushBusy(true);
    try {
      const err = await enablePushNotifications();
      if (err) {
        toast.error("Notifiche non attivate", { description: err });
      } else {
        toast.success("Notifiche attivate su questo dispositivo");
        setPushBanner(false);
      }
    } catch (e) {
      toast.error("Errore attivazione notifiche", {
        description: e instanceof Error ? e.message : String(e),
      });
    } finally {
      setPushBusy(false);
    }
  };

  const sediOptions = useMemo(() => {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const d of dipendenti) {
      const s = (d.sede ?? "").trim();
      if (s && s.toLowerCase() !== "tutte" && !seen.has(s.toLowerCase())) {
        seen.add(s.toLowerCase());
        out.push(s);
      }
    }
    return out.sort((a, b) => a.localeCompare(b));
  }, [dipendenti]);

  const nomeById = useMemo(() => {
    const m = new Map<string, string>();
    for (const d of dipendenti) m.set(d.id, d.nomeCompleto || `${d.cognome} ${d.nome}`);
    return m;
  }, [dipendenti]);

  const conferma = async (c: SpComunicazione) => {
    setConfermando(c.id);
    try {
      await spMarkPresaVisione({ data: { comunicazioneId: c.id } });
      setMieViste((v) => (v.includes(c.id) ? v : [...v, c.id]));
      toast.success("Presa visione registrata");
    } catch (err) {
      toast.error("Errore", { description: err instanceof Error ? err.message : String(err) });
    } finally {
      setConfermando(null);
    }
  };

  const mostraLetture = async (c: SpComunicazione) => {
    if (letture[c.id]) {
      setLetture((l) => {
        const { [c.id]: _omit, ...rest } = l;
        return rest;
      });
      return;
    }
    try {
      const pv = (await spGetPreseVisione({ data: { comunicazioneId: c.id } })) as SpPresaVisione[];
      setLetture((l) => ({ ...l, [c.id]: pv }));
    } catch (err) {
      toast.error("Errore", { description: err instanceof Error ? err.message : String(err) });
    }
  };

  const submit = async () => {
    if (!titolo.trim() || !testo.trim()) {
      toast.error("Titolo e testo sono obbligatori");
      return;
    }
    if (allegato && allegato.size > 8 * 1024 * 1024) {
      toast.error("Allegato troppo grande: il limite è 8 MB.");
      return;
    }
    setSubmitting(true);
    try {
      let allegatoUrl: string | undefined;
      if (allegato) {
        const contentBase64 = await fileToDataUrl(allegato);
        const up = await spUploadFile({
          data: { subfolder: "Comunicazioni", filename: allegato.name, contentBase64 },
        });
        allegatoUrl = up.webUrl;
      }
      const res = (await spCreateComunicazione({
        data: {
          titolo: titolo.trim(),
          testo: testo.trim(),
          tipo,
          sede,
          allegato: allegatoUrl,
          richiedePresaVisione,
          destinatariEmail: destEmail.trim() || undefined,
        },
      })) as SpComunicazione & { pushEsito?: string };
      toast.success("Comunicazione pubblicata", {
        description: res.pushEsito || undefined,
        duration: 8000,
      });
      setTitolo("");
      setTesto("");
      setAllegato(null);
      setRichiedePresaVisione(false);
      setDestEmail("");
      load();
    } catch (err) {
      toast.error("Errore nella pubblicazione", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <AppShell title="Comunicazioni" subtitle="Comunicazioni interne e avvisi">
      {pushBanner && (
        <div className="mb-6 flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-primary/30 bg-primary/5 p-4 shadow-[var(--shadow-card)]">
          <div className="flex items-start gap-3 min-w-0">
            <span className="h-9 w-9 shrink-0 rounded-lg bg-primary/10 text-primary flex items-center justify-center">
              <BellRing className="h-4 w-4" />
            </span>
            <div className="min-w-0">
              <div className="text-sm font-semibold text-foreground">
                Ricevi le comunicazioni sul telefono
              </div>
              <p className="text-[13px] text-muted-foreground mt-0.5">
                Attiva le notifiche per essere avvisato quando esce una nuova comunicazione, anche
                ad app chiusa.
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <Button type="button" size="sm" onClick={attivaPush} disabled={pushBusy}>
              {pushBusy ? (
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <BellRing className="h-4 w-4 mr-2" />
              )}
              Attiva notifiche
            </Button>
            <Button type="button" size="sm" variant="ghost" onClick={() => setPushBanner(false)}>
              Non ora
            </Button>
          </div>
        </div>
      )}

      {canPubblicare && (
        <div className="mb-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <Send className="h-4 w-4 text-primary" /> Nuova comunicazione
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            <div className="sm:col-span-2">
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Titolo
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={titolo}
                onChange={(e) => setTitolo(e.target.value)}
                placeholder="Es. Riunione mensile, Uso obbligatorio DPI…"
              />
            </div>
            <div className="sm:col-span-2">
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Testo
              </label>
              <textarea
                className={`${inputCls} mt-1 min-h-[90px] resize-y`}
                value={testo}
                onChange={(e) => setTesto(e.target.value)}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">Tipo</label>
              <select
                className={`${inputCls} mt-1`}
                value={tipo}
                onChange={(e) => setTipo(e.target.value as "Comunicazione" | "Riunione")}
              >
                <option value="Comunicazione">Comunicazione</option>
                <option value="Riunione">Riunione</option>
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Destinatari
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={sede}
                onChange={(e) => setSede(e.target.value)}
              >
                <option value="Tutte">Tutte le sedi</option>
                {sediOptions.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>
            <div className="sm:col-span-2">
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Invia anche via email a{" "}
                <span className="normal-case text-muted-foreground/70">(opzionale)</span>
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={destEmail}
                onChange={(e) => setDestEmail(e.target.value)}
                placeholder="email1@esempio.it; email2@esempio.it — anche esterni, separati da ; o ,"
              />
            </div>
            <div className="sm:col-span-2 flex flex-wrap items-center gap-4">
              <label className="flex items-center gap-2 text-sm text-foreground">
                <input
                  type="checkbox"
                  checked={richiedePresaVisione}
                  onChange={(e) => setRichiedePresaVisione(e.target.checked)}
                />
                Richiedi presa visione
              </label>
              <label className="flex items-center gap-2 text-sm text-muted-foreground">
                <Paperclip className="h-4 w-4" />
                <input
                  type="file"
                  accept="image/*,application/pdf"
                  className="text-xs"
                  onChange={(e) => setAllegato(e.target.files?.[0] ?? null)}
                />
              </label>
            </div>
          </div>
          <div className="mt-4 flex justify-end">
            <Button type="button" onClick={submit} disabled={submitting}>
              {submitting ? (
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Send className="h-4 w-4 mr-2" />
              )}
              Pubblica
            </Button>
          </div>
        </div>
      )}

      <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
            <Megaphone className="h-4 w-4 text-primary" /> Bacheca
          </div>
          {canPubblicare && (comunicazioni ?? []).length > 0 && (
            <button
              type="button"
              onClick={() =>
                esportaCsvFile(
                  "comunicazioni",
                  ["Tipologia", "Titolo", "Testo", "Sede", "Data", "Autore", "Presa visione"],
                  (comunicazioni ?? []).map((c) => [
                    c.tipo,
                    c.titolo,
                    c.testo,
                    c.sede || "Tutte",
                    fmtDataOra(c.dataComunicazione || c.createdAt),
                    c.autore,
                    c.richiedePresaVisione ? "Sì" : "No",
                  ]),
                )
              }
              className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors"
            >
              <Download className="h-4 w-4" /> Esporta CSV
            </button>
          )}
        </div>

        {comunicazioni === null ? (
          <div className="text-sm text-muted-foreground">Caricamento…</div>
        ) : comunicazioni.length === 0 ? (
          <div className="text-sm text-muted-foreground">Nessuna comunicazione.</div>
        ) : (
          <ul className="space-y-3">
            {comunicazioni.map((c) => {
              const vista = mieViste.includes(c.id);
              return (
                <li key={c.id} className="rounded-xl border border-border bg-background p-4">
                  <div className="flex flex-wrap items-center gap-2">
                    <span
                      className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${c.tipo === "Riunione" ? "bg-status-break/15 text-status-break" : "bg-primary/10 text-primary"}`}
                    >
                      {c.tipo || "Comunicazione"}
                    </span>
                    <span className="font-semibold text-foreground">{c.titolo}</span>
                    <span className="text-[11px] text-muted-foreground">· {c.sede || "Tutte"}</span>
                    {c.richiedePresaVisione && (
                      <span className="rounded-full bg-status-out/15 text-status-out px-2 py-0.5 text-[10px] font-medium uppercase tracking-wider">
                        presa visione
                      </span>
                    )}
                  </div>
                  <p className="mt-2 text-[13px] text-foreground/90 whitespace-pre-wrap leading-relaxed">
                    {c.testo}
                  </p>
                  <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] text-muted-foreground">
                    <span className="inline-flex items-center gap-1">
                      <CalendarDays className="h-3.5 w-3.5" />
                      {fmtDataOra(c.dataComunicazione || c.createdAt)}
                    </span>
                    {c.autore && <span>· {c.autore}</span>}
                    {c.allegato && (
                      <a
                        href={c.allegato}
                        target="_blank"
                        rel="noreferrer"
                        className="inline-flex items-center gap-1 text-primary underline"
                      >
                        <Paperclip className="h-3.5 w-3.5" /> allegato
                      </a>
                    )}
                  </div>

                  {c.richiedePresaVisione && (
                    <div className="mt-3 flex flex-wrap items-center gap-3">
                      {vista ? (
                        <span className="inline-flex items-center gap-1 text-[13px] text-status-present">
                          <CheckCircle2 className="h-4 w-4" /> Presa visione
                        </span>
                      ) : (
                        <Button
                          type="button"
                          size="sm"
                          onClick={() => conferma(c)}
                          disabled={confermando === c.id}
                        >
                          {confermando === c.id ? (
                            <Loader2 className="h-4 w-4 mr-2 animate-spin" />
                          ) : (
                            <CheckCircle2 className="h-4 w-4 mr-2" />
                          )}
                          Conferma presa visione
                        </Button>
                      )}
                      {canPubblicare && (
                        <button
                          type="button"
                          onClick={() => mostraLetture(c)}
                          className="inline-flex items-center gap-1 text-[13px] text-primary hover:underline"
                        >
                          <Eye className="h-3.5 w-3.5" />
                          {letture[c.id] ? "Nascondi letture" : "Chi ha letto"}
                        </button>
                      )}
                    </div>
                  )}

                  {canPubblicare && letture[c.id] && (
                    <div className="mt-2 rounded-lg bg-secondary/40 p-2 text-[12px]">
                      <div className="flex items-center gap-1 text-muted-foreground mb-1">
                        <Users className="h-3.5 w-3.5" /> {letture[c.id].length} letture
                      </div>
                      {letture[c.id].length === 0 ? (
                        <span className="text-muted-foreground">Nessuno ha ancora letto.</span>
                      ) : (
                        <ul className="space-y-0.5">
                          {letture[c.id].map((p) => (
                            <li key={p.id} className="text-foreground">
                              {nomeById.get(p.dipendenteId) || p.codiceDipendente || p.dipendenteId}
                              <span className="text-muted-foreground">
                                {" "}
                                · {fmtDataOra(p.dataLettura)}
                              </span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  )}
                </li>
              );
            })}
          </ul>
        )}
      </div>
    </AppShell>
  );
}
