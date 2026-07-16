import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { FileText, Upload, Loader2, Building2, User, Download } from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetDocumenti,
  spCreateDocumento,
  spUploadFile,
  spGetDipendenti,
} from "@/lib/sharepoint.functions";
import type { SpDocumento, SpDipendente } from "@/lib/sharepoint.server";

export const Route = createFileRoute("/documenti")({
  head: () => ({ meta: [{ title: "Documenti — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: DocumentiPage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

const CATEGORIE = ["Contratto", "Busta paga", "DPI", "Certificato corso", "Altro"] as const;

const CAT_STYLE: Record<string, string> = {
  Contratto: "bg-primary/10 text-primary",
  "Busta paga": "bg-status-present/15 text-status-present",
  DPI: "bg-status-out/15 text-status-out",
  "Certificato corso": "bg-status-break/15 text-status-break",
  Altro: "bg-muted text-muted-foreground",
};

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

// Estrae un codice fiscale italiano dal nome file (16 caratteri, pattern
// standard), anche in nomi "sporchi" (es. "CedolinoGiugno_RSSMRA80A01H501U.pdf").
const CF_RE = /[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]/i;
function estraiCF(filename: string): string | null {
  const m = filename.toUpperCase().match(CF_RE);
  return m ? m[0] : null;
}
function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Lettura file fallita"));
    reader.readAsDataURL(file);
  });
}

function DocumentiPage() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [documenti, setDocumenti] = useState<SpDocumento[] | null>(null);
  const [dipendenti, setDipendenti] = useState<SpDipendente[]>([]);

  // Form (solo pubblicatori)
  const [categoria, setCategoria] = useState<(typeof CATEGORIE)[number]>("Contratto");
  const [titolo, setTitolo] = useState("");
  const [ambito, setAmbito] = useState<"Personale" | "Generale">("Personale");
  const [destinatarioId, setDestinatarioId] = useState("");
  const [sedeDestinatario, setSedeDestinatario] = useState("Tutte");
  const [file, setFile] = useState<File | null>(null);
  const [submitting, setSubmitting] = useState(false);

  // Filtri elenco
  const [catF, setCatF] = useState<string>("tutte");

  // Buste paga (multi-upload con abbinamento per codice fiscale).
  // Il periodo è precompilato col MESE SCORSO (es. a gennaio 2026 → "Dicembre
  // 2025"), modificabile.
  const [bpFiles, setBpFiles] = useState<File[]>([]);
  const [bpPeriodo, setBpPeriodo] = useState(() => {
    const d = new Date();
    d.setMonth(d.getMonth() - 1);
    const mese = d.toLocaleDateString("it-IT", { month: "long" });
    return `${mese.charAt(0).toUpperCase()}${mese.slice(1)} ${d.getFullYear()}`;
  });
  const [bpBusy, setBpBusy] = useState(false);
  const [bpEsiti, setBpEsiti] = useState<{ nome: string; esito: string; ok: boolean }[]>([]);

  const canPubblicare =
    session != null &&
    (session.ruolo === "responsabile" ||
      session.ruolo === "amministratore_sistema" ||
      session.operatore);

  const load = () => {
    spGetDocumenti()
      .then((l) => setDocumenti(l as SpDocumento[]))
      .catch((err) => {
        setDocumenti([]);
        toast.error("Errore documenti", {
          description: err instanceof Error ? err.message : String(err),
        });
      });
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const nomeById = useMemo(() => {
    const m = new Map<string, string>();
    for (const d of dipendenti) m.set(d.id, d.nomeCompleto || `${d.cognome} ${d.nome}`);
    return m;
  }, [dipendenti]);

  const dipByCf = useMemo(() => {
    const m = new Map<string, SpDipendente>();
    for (const d of dipendenti) if (d.cf) m.set(d.cf.toUpperCase(), d);
    return m;
  }, [dipendenti]);

  // Anteprima abbinamento buste paga: file → dipendente (per CF nel nome file).
  const bpAnteprima = useMemo(
    () =>
      bpFiles.map((f) => {
        const cf = estraiCF(f.name);
        const dip = cf ? dipByCf.get(cf) : undefined;
        return { file: f, cf, dip };
      }),
    [bpFiles, dipByCf],
  );

  const isOperatoreOAdmin =
    session != null && (session.operatore || session.ruolo === "amministratore_sistema");

  const caricaBustePaga = async () => {
    if (!bpPeriodo.trim()) {
      toast.error("Indica il periodo (es. Giugno 2026)");
      return;
    }
    const abbinati = bpAnteprima.filter((a) => a.dip);
    if (abbinati.length === 0) {
      toast.error("Nessun file abbinato a un dipendente");
      return;
    }
    setBpBusy(true);
    const esiti: { nome: string; esito: string; ok: boolean }[] = [];
    for (const a of bpAnteprima) {
      if (!a.dip) {
        esiti.push({
          nome: a.file.name,
          esito: a.cf ? `CF ${a.cf} non trovato tra i dipendenti` : "Nessun CF nel nome file",
          ok: false,
        });
        continue;
      }
      try {
        if (a.file.size > 8 * 1024 * 1024) throw new Error("oltre 8 MB");
        const contentBase64 = await fileToDataUrl(a.file);
        const up = await spUploadFile({
          data: { subfolder: "Documenti", filename: a.file.name, contentBase64 },
        });
        await spCreateDocumento({
          data: {
            categoria: "Busta paga",
            titolo: `Busta paga ${bpPeriodo.trim()}`,
            ambito: "Personale",
            destinatarioId: a.dip.id,
            file: up.webUrl,
            nomeFile: up.fileName,
          },
        });
        esiti.push({
          nome: a.file.name,
          esito: `→ ${a.dip.nomeCompleto || a.dip.cognome} ✓`,
          ok: true,
        });
      } catch (err) {
        esiti.push({
          nome: a.file.name,
          esito: err instanceof Error ? err.message : String(err),
          ok: false,
        });
      }
    }
    setBpEsiti(esiti);
    setBpFiles([]);
    setBpBusy(false);
    const okN = esiti.filter((e) => e.ok).length;
    toast[okN > 0 ? "success" : "error"](
      `Buste paga: ${okN} caricate${esiti.length - okN ? `, ${esiti.length - okN} non abbinate` : ""}`,
    );
    load();
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

  const filtrati = useMemo(
    () => (documenti ?? []).filter((d) => catF === "tutte" || d.categoria === catF),
    [documenti, catF],
  );

  const resetForm = () => {
    setTitolo("");
    setFile(null);
    setDestinatarioId("");
    setSedeDestinatario("Tutte");
  };

  const submit = async () => {
    if (!file) {
      toast.error("Allega un file");
      return;
    }
    if (!titolo.trim()) {
      toast.error("Inserisci un titolo");
      return;
    }
    if (ambito === "Personale" && !destinatarioId) {
      toast.error("Seleziona il dipendente destinatario");
      return;
    }
    if (file.size > 8 * 1024 * 1024) {
      toast.error("File troppo grande: il limite è 8 MB.");
      return;
    }
    setSubmitting(true);
    try {
      const contentBase64 = await fileToDataUrl(file);
      const up = await spUploadFile({
        data: { subfolder: "Documenti", filename: file.name, contentBase64 },
      });
      await spCreateDocumento({
        data: {
          categoria,
          titolo: titolo.trim(),
          ambito,
          destinatarioId: ambito === "Personale" ? destinatarioId : undefined,
          sedeDestinatario: ambito === "Generale" ? sedeDestinatario : undefined,
          file: up.webUrl,
          nomeFile: up.fileName,
        },
      });
      toast.success("Documento caricato");
      resetForm();
      load();
    } catch (err) {
      toast.error("Errore nel caricamento", {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <AppShell title="Documenti" subtitle="Documenti dei dipendenti">
      {canPubblicare && (
        <div className="mb-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <Upload className="h-4 w-4 text-primary" /> Carica documento
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Categoria
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={categoria}
                onChange={(e) => setCategoria(e.target.value as (typeof CATEGORIE)[number])}
              >
                {CATEGORIE.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Titolo
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={titolo}
                onChange={(e) => setTitolo(e.target.value)}
                placeholder="Es. Contratto 2026, Busta paga giugno…"
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Destinatario
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={ambito}
                onChange={(e) => setAmbito(e.target.value as "Personale" | "Generale")}
              >
                <option value="Personale">Un dipendente specifico</option>
                <option value="Generale">Generale (tutti / una sede)</option>
              </select>
            </div>
            <div>
              {ambito === "Personale" ? (
                <>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Dipendente
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={destinatarioId}
                    onChange={(e) => setDestinatarioId(e.target.value)}
                  >
                    <option value="">— seleziona —</option>
                    {[...dipendenti]
                      .sort((a, b) =>
                        `${a.cognome} ${a.nome}`.localeCompare(`${b.cognome} ${b.nome}`),
                      )
                      .map((d) => (
                        <option key={d.id} value={d.id}>
                          {d.cognome} {d.nome}
                        </option>
                      ))}
                  </select>
                </>
              ) : (
                <>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    Sede
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={sedeDestinatario}
                    onChange={(e) => setSedeDestinatario(e.target.value)}
                  >
                    <option value="Tutte">Tutte le sedi</option>
                    {sediOptions.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </>
              )}
            </div>
            <div className="sm:col-span-2">
              <label className="text-xs uppercase tracking-wider text-muted-foreground">File</label>
              <input
                type="file"
                accept="image/*,application/pdf"
                className={`${inputCls} mt-1`}
                onChange={(e) => setFile(e.target.files?.[0] ?? null)}
              />
              <p className="mt-1 text-[11px] text-muted-foreground">PDF o immagine · max 8 MB</p>
            </div>
          </div>
          <div className="mt-4 flex justify-end">
            <Button type="button" onClick={submit} disabled={submitting}>
              {submitting ? (
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Upload className="h-4 w-4 mr-2" />
              )}
              Carica
            </Button>
          </div>
        </div>
      )}

      {isOperatoreOAdmin && (
        <div className="mb-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <Upload className="h-4 w-4 text-primary" /> Carica buste paga (multiplo)
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">
            Seleziona tutti i PDF insieme: l'abbinamento al dipendente avviene dal{" "}
            <strong>codice fiscale nel nome del file</strong> (colonna CF su Dipendenti). Ogni
            dipendente riceve il documento personale e una notifica.
          </p>
          <div className="grid gap-3 sm:grid-cols-2">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                Periodo
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={bpPeriodo}
                onChange={(e) => setBpPeriodo(e.target.value)}
                placeholder="Es. Giugno 2026"
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                File PDF
              </label>
              <input
                type="file"
                multiple
                accept="application/pdf"
                className={`${inputCls} mt-1`}
                onChange={(e) => {
                  setBpFiles(Array.from(e.target.files ?? []));
                  setBpEsiti([]);
                }}
              />
            </div>
          </div>

          {bpAnteprima.length > 0 && (
            <div className="mt-3 rounded-lg border border-border p-3 text-[12px]">
              <p className="text-muted-foreground mb-1">
                Anteprima abbinamenti ({bpAnteprima.filter((a) => a.dip).length}/
                {bpAnteprima.length} riconosciuti):
              </p>
              <ul className="space-y-0.5 max-h-48 overflow-auto font-mono">
                {bpAnteprima.map((a, i) => (
                  <li key={i} className={a.dip ? "text-foreground" : "text-status-absent"}>
                    {a.file.name} →{" "}
                    {a.dip
                      ? a.dip.nomeCompleto || a.dip.cognome
                      : a.cf
                        ? `CF ${a.cf} non trovato`
                        : "nessun CF nel nome"}
                  </li>
                ))}
              </ul>
            </div>
          )}
          {bpEsiti.length > 0 && (
            <div className="mt-3 rounded-lg border border-border p-3 text-[12px]">
              <ul className="space-y-0.5 max-h-48 overflow-auto font-mono">
                {bpEsiti.map((e, i) => (
                  <li key={i} className={e.ok ? "text-status-present" : "text-status-absent"}>
                    {e.nome}: {e.esito}
                  </li>
                ))}
              </ul>
            </div>
          )}

          <div className="mt-4 flex justify-end">
            <Button
              type="button"
              onClick={caricaBustePaga}
              disabled={bpBusy || bpAnteprima.filter((a) => a.dip).length === 0}
            >
              {bpBusy ? (
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Upload className="h-4 w-4 mr-2" />
              )}
              Carica {bpAnteprima.filter((a) => a.dip).length} buste paga
            </Button>
          </div>
        </div>
      )}

      <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
            <FileText className="h-4 w-4 text-primary" />
            {canPubblicare ? "Tutti i documenti" : "I miei documenti"}
          </div>
          <div className="flex items-center gap-2">
            <select
              className="rounded-lg border border-border bg-background px-2 py-1 text-xs"
              value={catF}
              onChange={(e) => setCatF(e.target.value)}
            >
              <option value="tutte">Tutte le categorie</option>
              {CATEGORIE.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
            {canPubblicare && filtrati.length > 0 && (
              <button
                type="button"
                onClick={() =>
                  esportaCsvFile(
                    "documenti",
                    [
                      "Categoria",
                      "Titolo",
                      "Ambito",
                      "Destinatario",
                      "Sede",
                      "Data",
                      "Caricato da",
                      "Link",
                    ],
                    filtrati.map((d) => [
                      d.categoria,
                      d.titolo,
                      d.ambito,
                      d.ambito === "Personale"
                        ? nomeById.get(d.destinatarioId) || d.codiceDestinatario
                        : "",
                      d.sedeDestinatario,
                      fmtData(d.dataDocumento || d.createdAt),
                      d.caricatoDa,
                      d.file,
                    ]),
                  )
                }
                className="inline-flex items-center gap-1.5 rounded-lg border border-border bg-background px-2 py-1 text-xs text-foreground hover:bg-secondary transition-colors"
              >
                <Download className="h-3.5 w-3.5" /> CSV
              </button>
            )}
          </div>
        </div>

        {documenti === null ? (
          <div className="text-sm text-muted-foreground">Caricamento…</div>
        ) : filtrati.length === 0 ? (
          <div className="text-sm text-muted-foreground">Nessun documento disponibile.</div>
        ) : (
          <ul className="space-y-2">
            {filtrati.map((d) => (
              <li
                key={d.id}
                className="flex items-start justify-between gap-3 rounded-xl border border-border bg-background p-3"
              >
                <div className="min-w-0">
                  <div className="flex flex-wrap items-center gap-2">
                    <span
                      className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${CAT_STYLE[d.categoria] ?? "bg-muted text-muted-foreground"}`}
                    >
                      {d.categoria || "—"}
                    </span>
                    <span className="font-medium text-foreground truncate">{d.titolo || "—"}</span>
                  </div>
                  <div className="mt-1 flex items-center gap-2 text-[12px] text-muted-foreground">
                    {d.ambito === "Personale" ? (
                      <>
                        <User className="h-3.5 w-3.5" />
                        {nomeById.get(d.destinatarioId) || d.codiceDestinatario || "dipendente"}
                      </>
                    ) : (
                      <>
                        <Building2 className="h-3.5 w-3.5" />
                        {d.sedeDestinatario || "Tutte"}
                      </>
                    )}
                    <span>· {fmtData(d.dataDocumento || d.createdAt)}</span>
                    {canPubblicare && d.caricatoDa && <span>· da {d.caricatoDa}</span>}
                  </div>
                </div>
                {d.file ? (
                  <a
                    href={d.file}
                    target="_blank"
                    rel="noreferrer"
                    className="shrink-0 text-primary underline text-sm"
                  >
                    apri
                  </a>
                ) : (
                  <span className="shrink-0 text-muted-foreground text-sm">—</span>
                )}
              </li>
            ))}
          </ul>
        )}
      </div>
    </AppShell>
  );
}
