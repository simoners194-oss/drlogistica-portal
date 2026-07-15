import { createFileRoute, redirect } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import {
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  KeyRound,
  Loader2,
  PlayCircle,
  XCircle,
  Upload,
  Users,
} from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { useEffect, useState, useCallback } from "react";
import {
  dataService,
  getIntegrationStatus,
  refreshIntegrationDiagnostics,
  runSpSelfTest,
  type IntegrationStatus,
} from "@/lib/data-service";
import type { SpSelfTestResult, SpHealth, ImportDipendentiResult } from "@/lib/sharepoint.server";
import { spImportDipendenti } from "@/lib/sharepoint.functions";
import { readSession } from "@/lib/session";
import {
  microsoftAuthConfig,
  isMicrosoftAuthConfigured,
  testMicrosoftAuth,
  type MicrosoftAuthTestResult,
} from "@/lib/auth-config";

export const Route = createFileRoute("/amministrazione")({
  head: () => ({ meta: [{ title: "Modulo Amministrazione — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    const s = readSession();
    if (!s) {
      throw redirect({ to: "/", search: { redirect: location.href } });
    }
    if (s.ruolo !== "amministratore_sistema") {
      // Fallback pulito: rimanda alla landing del ruolo (Presenze o Dashboard).
      throw redirect({ to: s.ruolo === "dipendente" ? "/presenze" : "/dashboard" });
    }
  },
  component: AmministrazionePage,
});

function AmministrazionePage() {
  const [status, setStatus] = useState<IntegrationStatus>(() => getIntegrationStatus());
  const [loading, setLoading] = useState(false);
  const [selfTest, setSelfTest] = useState<SpSelfTestResult | null>(null);
  const [selfTestLoading, setSelfTestLoading] = useState(false);

  const refresh = useCallback(async (force = false) => {
    setLoading(true);
    try {
      await refreshIntegrationDiagnostics(force);
      await dataService.getDipendenti();
    } catch {
      /* già tracciato in integrationStatus */
    } finally {
      setStatus(getIntegrationStatus());
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
  }, [refresh]);

  const runTest = async () => {
    setSelfTestLoading(true);
    try {
      const r = (await runSpSelfTest()) as SpSelfTestResult;
      setSelfTest(r);
      setStatus(getIntegrationStatus());
    } finally {
      setSelfTestLoading(false);
    }
  };

  const discovered = status.diagnostics?.discovered ?? null;
  const spError = status.diagnostics?.error ?? status.ultimoErrore;
  const connected = Boolean(discovered) && !spError;
  const health: SpHealth | null = status.diagnostics?.health ?? null;
  const lastSyncAt = status.diagnostics?.lastSyncAt ?? null;

  return (
    <AppShell title="Amministrazione" subtitle="Configurazione sedi, ruoli e integrazioni">
      <div className="space-y-6">
        <Card>
          <CardHeader className="flex flex-row items-start justify-between gap-4 space-y-0">
            <div>
              <CardTitle className="flex items-center gap-2 text-base">
                {connected ? (
                  <CheckCircle2 className="h-5 w-5 text-status-present" />
                ) : (
                  <AlertTriangle className="h-5 w-5 text-status-absent" />
                )}
                Stato integrazione SharePoint
              </CardTitle>
              <p className="text-xs text-muted-foreground mt-1">
                Sito e liste rilevati automaticamente tramite il connettore Microsoft SharePoint.
              </p>
            </div>
            <Button size="sm" variant="outline" onClick={() => refresh(true)} disabled={loading}>
              <RefreshCw className={`h-4 w-4 mr-2 ${loading ? "animate-spin" : ""}`} />
              Riscopri
            </Button>
          </CardHeader>
          <CardContent className="grid gap-4 sm:grid-cols-2">
            <InfoRow label="Modalità attuale">
              <Badge
                variant={connected ? "default" : "secondary"}
                className={connected ? "bg-status-present text-white" : ""}
              >
                {connected ? "SharePoint (live)" : "Non connesso"}
              </Badge>
            </InfoRow>
            <InfoRow label="Dipendenti caricati">
              <span className="font-semibold">{status.dipendentiCaricati}</span>
            </InfoRow>
            <InfoRow label="Sito SharePoint">
              {discovered ? (
                <a
                  href={discovered.siteWebUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="text-sm underline underline-offset-2 hover:text-primary"
                >
                  {discovered.siteName}
                </a>
              ) : (
                <span className="text-sm text-muted-foreground">—</span>
              )}
            </InfoRow>
            <InfoRow label="Liste rilevate">
              {discovered ? (
                <div className="flex flex-wrap gap-1">
                  <Badge variant="secondary" className="font-mono text-[10px]">
                    {discovered.listDipendentiName}
                  </Badge>
                  <Badge variant="secondary" className="font-mono text-[10px]">
                    {discovered.listTimbratureName}
                  </Badge>
                </div>
              ) : (
                <span className="text-sm text-muted-foreground">—</span>
              )}
            </InfoRow>
            <InfoRow label="Ultimo aggiornamento">
              <span className="text-sm">
                {status.ultimoAggiornamento
                  ? status.ultimoAggiornamento.toLocaleTimeString("it-IT")
                  : "—"}
              </span>
            </InfoRow>
            <InfoRow label="Errori integrazione">
              {spError ? (
                <span className="text-sm text-status-absent break-all">{spError}</span>
              ) : (
                <span className="text-sm text-muted-foreground">Nessuno</span>
              )}
            </InfoRow>
            {status.diagnostics &&
              (!status.diagnostics.hasLovableKey || !status.diagnostics.hasConnectionKey) && (
                <div className="sm:col-span-2 rounded-md border border-dashed border-status-absent/50 p-3 text-xs text-status-absent">
                  <p className="font-medium mb-1">Credenziali connettore mancanti sul server</p>
                  <ul className="list-disc list-inside space-y-0.5">
                    {!status.diagnostics.hasLovableKey && <li>LOVABLE_API_KEY non disponibile</li>}
                    {!status.diagnostics.hasConnectionKey && (
                      <li>
                        MICROSOFT_SHAREPOINT_API_KEY non disponibile — riconnetti il connettore.
                      </li>
                    )}
                  </ul>
                </div>
              )}
            {status.log.length > 0 && (
              <div className="sm:col-span-2 rounded-md border border-border p-3">
                <p className="text-[11px] uppercase tracking-wider text-muted-foreground mb-2">
                  Log integrazione (ultime {status.log.length} operazioni)
                </p>
                <ul className="space-y-1 max-h-56 overflow-auto text-xs font-mono">
                  {status.log.map((entry, i) => (
                    <li key={i} className="flex gap-2">
                      <span className="text-muted-foreground shrink-0">
                        {entry.ts.toLocaleTimeString("it-IT")}
                      </span>
                      <span
                        className={
                          entry.level === "error"
                            ? "text-status-absent shrink-0"
                            : entry.level === "warn"
                              ? "text-status-break shrink-0"
                              : "text-status-present shrink-0"
                        }
                      >
                        {entry.level.toUpperCase()}
                      </span>
                      <span className="text-foreground shrink-0">{entry.operation}</span>
                      <span className="text-muted-foreground break-all">{entry.message}</span>
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </CardContent>
        </Card>

        <HealthCard
          health={health}
          lastSyncAt={lastSyncAt}
          onSelfTest={runTest}
          selfTest={selfTest}
          selfTestLoading={selfTestLoading}
        />

        <ImportDipendentiCard onDone={() => refresh(true)} />

        <MicrosoftAuthTestCard />
      </div>
    </AppShell>
  );
}

function HealthCard({
  health,
  lastSyncAt,
  onSelfTest,
  selfTest,
  selfTestLoading,
}: {
  health: SpHealth | null;
  lastSyncAt: string | null;
  onSelfTest: () => void;
  selfTest: SpSelfTestResult | null;
  selfTestLoading: boolean;
}) {
  const items: { label: string; ok: boolean; detail?: string }[] = health
    ? [
        { label: "Connessione Graph", ok: health.graphOk },
        { label: "Token valido", ok: health.tokenOk },
        { label: "Permessi Sites.Read", ok: health.permissionsOk },
        { label: "Sito trovato", ok: health.siteFound, detail: health.siteName ?? undefined },
        { label: "Lista Dipendenti", ok: health.dipendentiListFound },
        { label: "Lista Timbrature", ok: health.timbratureListFound },
        {
          label: "Colonne Dipendenti",
          ok: health.dipendentiColumnsOk,
          detail: health.dipendentiMissing.length
            ? `mancanti: ${health.dipendentiMissing.join(", ")}`
            : undefined,
        },
        {
          label: "Colonne Timbrature",
          ok: health.timbratureColumnsOk,
          detail: health.timbratureMissing.length
            ? `mancanti: ${health.timbratureMissing.join(", ")}`
            : undefined,
        },
      ]
    : [];

  return (
    <Card>
      <CardHeader className="flex flex-row items-start justify-between gap-4 space-y-0">
        <div>
          <CardTitle className="text-base flex items-center gap-2">
            <CheckCircle2 className="h-5 w-5 text-primary" />
            Salute integrazione
          </CardTitle>
          <p className="text-xs text-muted-foreground mt-1">
            Checklist automatica + self-test end-to-end con rollback.
          </p>
        </div>
        <Button size="sm" onClick={onSelfTest} disabled={selfTestLoading}>
          {selfTestLoading ? (
            <Loader2 className="h-4 w-4 mr-2 animate-spin" />
          ) : (
            <PlayCircle className="h-4 w-4 mr-2" />
          )}
          Esegui test integrazione
        </Button>
      </CardHeader>
      <CardContent className="grid gap-4 sm:grid-cols-2">
        <div className="sm:col-span-2 grid gap-1.5">
          {items.map((i) => (
            <div key={i.label} className="flex items-center gap-2 text-sm">
              {i.ok ? (
                <CheckCircle2 className="h-4 w-4 text-status-present shrink-0" />
              ) : (
                <XCircle className="h-4 w-4 text-status-absent shrink-0" />
              )}
              <span className="text-foreground">{i.label}</span>
              {i.detail && <span className="text-xs text-muted-foreground">· {i.detail}</span>}
            </div>
          ))}
          {!health && (
            <p className="text-xs text-muted-foreground">In attesa dei dati diagnostici…</p>
          )}
        </div>
        <InfoRow label="Ultima sincronizzazione">
          <span className="text-sm">
            {lastSyncAt ? new Date(lastSyncAt).toLocaleTimeString("it-IT") : "—"}
          </span>
        </InfoRow>
        <InfoRow label="Tempo risposta Graph">
          <span className="text-sm tabular-nums">
            {health?.graphResponseMs ? `${health.graphResponseMs} ms` : "—"}
          </span>
        </InfoRow>
        <InfoRow label="Scadenza cache discovery">
          <span className="text-sm">
            {health?.cacheExpiresAt
              ? new Date(health.cacheExpiresAt).toLocaleTimeString("it-IT")
              : "—"}
          </span>
        </InfoRow>
        <InfoRow label="Site ID">
          <span className="font-mono text-[10px] break-all text-muted-foreground">
            {health?.siteId ?? "—"}
          </span>
        </InfoRow>

        {selfTest && (
          <div className="sm:col-span-2 rounded-md border border-border p-3">
            <div className="flex items-center justify-between mb-2">
              <p className="text-sm font-medium">Risultato self-test</p>
              <Badge
                className={
                  selfTest.score === 100
                    ? "bg-status-present text-white"
                    : selfTest.score >= 70
                      ? "bg-status-break text-white"
                      : "bg-status-absent text-white"
                }
              >
                Salute {selfTest.score}%
              </Badge>
            </div>
            <ul className="space-y-1 text-xs">
              {selfTest.checks.map((c) => (
                <li key={c.key} className="flex items-start gap-2">
                  {c.ok ? (
                    <CheckCircle2 className="h-3.5 w-3.5 text-status-present mt-0.5 shrink-0" />
                  ) : (
                    <XCircle className="h-3.5 w-3.5 text-status-absent mt-0.5 shrink-0" />
                  )}
                  <span className="font-medium text-foreground w-56 shrink-0">{c.label}</span>
                  <span className="text-muted-foreground tabular-nums shrink-0">
                    {c.durataMs ? `${c.durataMs}ms` : ""}
                  </span>
                  {c.message && (
                    <span
                      className={c.ok ? "text-muted-foreground" : "text-status-absent break-all"}
                    >
                      {c.message}
                    </span>
                  )}
                </li>
              ))}
            </ul>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function ImportDipendentiCard({ onDone }: { onDone: () => void }) {
  const [csv, setCsv] = useState("");
  const [result, setResult] = useState<ImportDipendentiResult | null>(null);
  const [loading, setLoading] = useState<"preview" | "import" | null>(null);
  const [error, setError] = useState<string | null>(null);

  const run = async (dryRun: boolean) => {
    setError(null);
    setLoading(dryRun ? "preview" : "import");
    try {
      const r = (await spImportDipendenti({
        data: { csv, dryRun },
      })) as ImportDipendentiResult;
      setResult(r);
      if (!dryRun && r.created > 0) onDone();
    } catch (e) {
      setResult(null);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(null);
    }
  };

  const hasMissing = Boolean(result && result.missingColumns.length > 0);
  const canImport =
    csv.trim().length > 0 && !hasMissing && result != null && result.dryRun && result.totalRows > 0;

  return (
    <Card>
      <CardHeader className="space-y-1">
        <CardTitle className="flex items-center gap-2 text-base">
          <Users className="h-5 w-5 text-primary" />
          Import massivo dipendenti
        </CardTitle>
        <p className="text-xs text-muted-foreground">
          Incolla dati da Excel (TAB) o CSV, prima riga con le intestazioni uguali ai nomi colonna
          della lista SharePoint. Usa le credenziali del portale: nessun secret da configurare. Fai
          sempre l'<strong>Anteprima</strong> prima di importare.
        </p>
      </CardHeader>
      <CardContent className="space-y-3">
        <textarea
          value={csv}
          onChange={(e) => {
            setCsv(e.target.value);
            setResult(null);
          }}
          placeholder={
            "Nome,Cognome,Nome Completo,Email,Sede,Attivo,Ruolo,Codice,PIN,Visibile,Autorizza,Operatore,OreSettimanali\nMario,Rossi,Mario Rossi,mario@dr.it,Milano,Sì,Dipendente,DR100,1234,Sì,No,No,40"
          }
          spellCheck={false}
          className="w-full min-h-[140px] resize-y rounded-lg border border-border bg-background px-3 py-2 font-mono text-xs text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40"
        />

        <div className="flex flex-wrap items-center gap-3">
          <Button
            size="sm"
            variant="outline"
            onClick={() => run(true)}
            disabled={loading !== null || csv.trim().length === 0}
          >
            {loading === "preview" ? (
              <Loader2 className="h-4 w-4 mr-2 animate-spin" />
            ) : (
              <PlayCircle className="h-4 w-4 mr-2" />
            )}
            Anteprima (non scrive)
          </Button>
          <Button size="sm" onClick={() => run(false)} disabled={loading !== null || !canImport}>
            {loading === "import" ? (
              <Loader2 className="h-4 w-4 mr-2 animate-spin" />
            ) : (
              <Upload className="h-4 w-4 mr-2" />
            )}
            Importa {result?.dryRun ? `${result.totalRows} righe` : ""}
          </Button>
          {result?.dryRun && !hasMissing && (
            <span className="text-xs text-muted-foreground">
              Anteprima ok · premi <strong>Importa</strong> per creare i record.
            </span>
          )}
        </div>

        {error && (
          <div className="rounded-md border border-status-absent/40 bg-status-absent/5 p-3 text-xs text-status-absent break-all">
            {error}
          </div>
        )}

        {hasMissing && (
          <div className="rounded-md border border-status-absent/40 bg-status-absent/5 p-3 text-xs text-status-absent">
            <p className="font-medium">Intestazioni non riconosciute (import bloccato):</p>
            <p className="mt-1 break-all">{result!.missingColumns.join(", ")}</p>
            {result!.availableColumns.length > 0 && (
              <p className="mt-2 text-muted-foreground break-all">
                <span className="font-medium text-foreground">
                  Colonne scrivibili su SharePoint:
                </span>{" "}
                {result!.availableColumns.join(", ")}
              </p>
            )}
            <p className="mt-2 text-muted-foreground">
              L'intestazione deve combaciare ESATTAMENTE con una colonna qui sopra. Se
              "Inquadramento" non è nell'elenco, la colonna è di sola lettura/nascosta o ha un nome
              diverso: correggi su SharePoint (o l'intestazione nel file), poi rifai l'anteprima.
            </p>
          </div>
        )}

        {result && !hasMissing && (
          <div className="rounded-md border border-border p-3 text-xs">
            <div className="flex items-center gap-2 mb-2">
              {result.dryRun ? (
                <Badge variant="secondary">Anteprima · {result.totalRows} righe</Badge>
              ) : (
                <Badge className="bg-status-present text-white">
                  Creati {result.created}
                  {result.failed ? ` · Errori ${result.failed}` : ""}
                </Badge>
              )}
              <span className="text-muted-foreground">
                Colonne riconosciute: {result.matchedColumns.length}
              </span>
            </div>
            <ul className="space-y-1 max-h-60 overflow-auto font-mono">
              {result.rows.map((r, i) => (
                <li key={i} className="flex items-start gap-2">
                  {r.ok ? (
                    <CheckCircle2 className="h-3.5 w-3.5 text-status-present mt-0.5 shrink-0" />
                  ) : (
                    <XCircle className="h-3.5 w-3.5 text-status-absent mt-0.5 shrink-0" />
                  )}
                  <span className="text-foreground shrink-0">{r.label}</span>
                  {r.error && <span className="text-status-absent break-all">{r.error}</span>}
                </li>
              ))}
            </ul>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function InfoRow({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1">
      <span className="text-[11px] uppercase tracking-wider text-muted-foreground">{label}</span>
      <div>{children}</div>
    </div>
  );
}

function MicrosoftAuthTestCard() {
  const configured = isMicrosoftAuthConfigured();
  const hasTenant = Boolean(microsoftAuthConfig.TENANT_ID);
  const hasClient = Boolean(microsoftAuthConfig.CLIENT_ID);
  const [result, setResult] = useState<MicrosoftAuthTestResult | null>(null);
  const [loading, setLoading] = useState(false);

  const runTest = async () => {
    setLoading(true);
    try {
      setResult(await testMicrosoftAuth());
    } finally {
      setLoading(false);
    }
  };

  return (
    <Card>
      <CardHeader className="space-y-1">
        <CardTitle className="flex items-center gap-2 text-base">
          <KeyRound className="h-5 w-5 text-primary" />
          Test autenticazione Microsoft
        </CardTitle>
        <p className="text-xs text-muted-foreground">
          Verifica opzionale di Entra ID. L'app funziona regolarmente anche senza configurazione.
        </p>
      </CardHeader>
      <CardContent className="grid gap-4 sm:grid-cols-2">
        <InfoRow label="Stato configurazione">
          <Badge
            variant={configured ? "default" : "secondary"}
            className={configured ? "bg-status-present text-white" : ""}
          >
            {configured ? "Configurato" : "Non configurato"}
          </Badge>
        </InfoRow>
        <InfoRow label="Redirect URI">
          <span className="font-mono text-xs break-all">
            {microsoftAuthConfig.REDIRECT_URI || "—"}
          </span>
        </InfoRow>
        <InfoRow label="Tenant ID">
          <FieldStatus present={hasTenant} envKey="VITE_MS_TENANT_ID" />
        </InfoRow>
        <InfoRow label="Client ID">
          <FieldStatus present={hasClient} envKey="VITE_MS_CLIENT_ID" />
        </InfoRow>

        <div className="sm:col-span-2 flex flex-wrap items-center gap-3">
          <Button size="sm" onClick={runTest} disabled={loading}>
            {loading ? (
              <Loader2 className="h-4 w-4 mr-2 animate-spin" />
            ) : (
              <KeyRound className="h-4 w-4 mr-2" />
            )}
            Test login Microsoft
          </Button>
          <span className="text-xs text-muted-foreground">
            Il test verifica la validità del tenant senza aprire una finestra di login.
          </span>
        </div>

        {result && (
          <div
            className={`sm:col-span-2 rounded-md border p-3 text-xs ${
              result.ok
                ? "border-status-present/40 bg-status-present/5"
                : "border-status-absent/40 bg-status-absent/5"
            }`}
          >
            <p className="font-medium text-foreground flex items-center gap-2">
              {result.ok ? (
                <CheckCircle2 className="h-4 w-4 text-status-present" />
              ) : (
                <AlertTriangle className="h-4 w-4 text-status-absent" />
              )}
              {result.message}
            </p>
            {result.detail && (
              <p className="mt-1 text-muted-foreground break-all">{result.detail}</p>
            )}
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function FieldStatus({ present, envKey }: { present: boolean; envKey: string }) {
  return (
    <div className="flex items-center gap-2">
      <Badge
        variant={present ? "default" : "outline"}
        className={present ? "bg-status-present text-white" : "text-muted-foreground"}
      >
        {present ? "Presente" : "Mancante"}
      </Badge>
      <span className="font-mono text-[11px] text-muted-foreground">{envKey}</span>
    </div>
  );
}
