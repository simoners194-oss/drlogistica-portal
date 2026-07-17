import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { Button } from "@/components/ui/button";
import { FileText, Upload, Loader2, Building2, User, Download } from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";
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
  const { t, tVal } = useLang();
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
        toast.error(t("doc.errDocs"), {
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
      toast.error(t("doc.bpNeedPeriod"));
      return;
    }
    const abbinati = bpAnteprima.filter((a) => a.dip);
    if (abbinati.length === 0) {
      toast.error(t("doc.bpNoneMatched"));
      return;
    }
    setBpBusy(true);
    const esiti: { nome: string; esito: string; ok: boolean }[] = [];
    for (const a of bpAnteprima) {
      if (!a.dip) {
        esiti.push({
          nome: a.file.name,
          esito: a.cf ? `CF ${a.cf}: ${t("doc.bpCfNotFound")}` : t("doc.bpNoCf"),
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
      `${t("doc.bpDone")} ${okN} ${t("doc.bpDoneUploaded")}${esiti.length - okN ? `, ${esiti.length - okN} ${t("doc.bpDoneUnmatched")}` : ""}`,
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
      toast.error(t("doc.errAttach"));
      return;
    }
    if (!titolo.trim()) {
      toast.error(t("doc.errTitle"));
      return;
    }
    if (ambito === "Personale" && !destinatarioId) {
      toast.error(t("doc.errRecipient"));
      return;
    }
    if (file.size > 8 * 1024 * 1024) {
      toast.error(t("rich.fileTooBig"));
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
      toast.success(t("doc.uploaded"));
      resetForm();
      load();
    } catch (err) {
      toast.error(t("doc.uploadErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <AppShell title={t("doc.title")} subtitle={t("doc.subtitle")}>
      {canPubblicare && (
        <div className="mb-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-4">
            <Upload className="h-4 w-4 text-primary" /> {t("doc.uploadTitle")}
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.category")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={categoria}
                onChange={(e) => setCategoria(e.target.value as (typeof CATEGORIE)[number])}
              >
                {CATEGORIE.map((c) => (
                  <option key={c} value={c}>
                    {tVal("cat", c)}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.docTitle")}
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={titolo}
                onChange={(e) => setTitolo(e.target.value)}
                placeholder={t("doc.titlePh")}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.recipient")}
              </label>
              <select
                className={`${inputCls} mt-1`}
                value={ambito}
                onChange={(e) => setAmbito(e.target.value as "Personale" | "Generale")}
              >
                <option value="Personale">{t("doc.recipientPersonal")}</option>
                <option value="Generale">{t("doc.recipientGeneral")}</option>
              </select>
            </div>
            <div>
              {ambito === "Personale" ? (
                <>
                  <label className="text-xs uppercase tracking-wider text-muted-foreground">
                    {t("common.employee")}
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={destinatarioId}
                    onChange={(e) => setDestinatarioId(e.target.value)}
                  >
                    <option value="">{t("common.select")}</option>
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
                    {t("common.site")}
                  </label>
                  <select
                    className={`${inputCls} mt-1`}
                    value={sedeDestinatario}
                    onChange={(e) => setSedeDestinatario(e.target.value)}
                  >
                    <option value="Tutte">{t("common.allSites")}</option>
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
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.file")}
              </label>
              <input
                type="file"
                accept="image/*,application/pdf"
                className={`${inputCls} mt-1`}
                onChange={(e) => setFile(e.target.files?.[0] ?? null)}
              />
              <p className="mt-1 text-[11px] text-muted-foreground">{t("doc.fileHint")}</p>
            </div>
          </div>
          <div className="mt-4 flex justify-end">
            <Button type="button" onClick={submit} disabled={submitting}>
              {submitting ? (
                <Loader2 className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Upload className="h-4 w-4 mr-2" />
              )}
              {t("doc.upload")}
            </Button>
          </div>
        </div>
      )}

      {isOperatoreOAdmin && (
        <div className="mb-6 rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground mb-1">
            <Upload className="h-4 w-4 text-primary" /> {t("doc.bpTitle")}
          </div>
          <p className="text-[12px] text-muted-foreground mb-4">{t("doc.bpDesc")}</p>
          <div className="grid gap-3 sm:grid-cols-2">
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.bpPeriod")}
              </label>
              <input
                className={`${inputCls} mt-1`}
                value={bpPeriodo}
                onChange={(e) => setBpPeriodo(e.target.value)}
                placeholder={t("doc.bpPeriodPh")}
              />
            </div>
            <div>
              <label className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("doc.bpFiles")}
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
                {t("doc.bpPreview")} ({bpAnteprima.filter((a) => a.dip).length}/{bpAnteprima.length}{" "}
                {t("doc.bpRecognized")}):
              </p>
              <ul className="space-y-0.5 max-h-48 overflow-auto font-mono">
                {bpAnteprima.map((a, i) => (
                  <li key={i} className={a.dip ? "text-foreground" : "text-status-absent"}>
                    {a.file.name} →{" "}
                    {a.dip
                      ? a.dip.nomeCompleto || a.dip.cognome
                      : a.cf
                        ? `CF ${a.cf}: ${t("doc.bpCfNotFound")}`
                        : t("doc.bpNoCf")}
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
              {t("doc.bpUploadN")} ({bpAnteprima.filter((a) => a.dip).length})
            </Button>
          </div>
        </div>
      )}

      <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
            <FileText className="h-4 w-4 text-primary" />
            {canPubblicare ? t("doc.listAll") : t("doc.listMine")}
          </div>
          <div className="flex items-center gap-2">
            <select
              className="rounded-lg border border-border bg-background px-2 py-1 text-xs"
              value={catF}
              onChange={(e) => setCatF(e.target.value)}
            >
              <option value="tutte">{t("doc.allCategories")}</option>
              {CATEGORIE.map((c) => (
                <option key={c} value={c}>
                  {tVal("cat", c)}
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
          <div className="text-sm text-muted-foreground">{t("common.loading")}</div>
        ) : filtrati.length === 0 ? (
          <div className="text-sm text-muted-foreground">{t("doc.none")}</div>
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
                      {d.categoria ? tVal("cat", d.categoria) : "—"}
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
                    {canPubblicare && d.caricatoDa && (
                      <span>
                        · {t("doc.by")} {d.caricatoDa}
                      </span>
                    )}
                  </div>
                </div>
                {d.file ? (
                  <a
                    href={d.file}
                    target="_blank"
                    rel="noreferrer"
                    className="shrink-0 text-primary underline text-sm"
                  >
                    {t("common.open")}
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
