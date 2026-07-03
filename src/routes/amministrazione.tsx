import { createFileRoute } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import { PlaceholderSection } from "@/components/PlaceholderSection";
import { Settings, CheckCircle2, AlertTriangle, RefreshCw, KeyRound, Loader2 } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { useEffect, useState, useCallback } from "react";
import { dataService, getIntegrationStatus, type IntegrationStatus } from "@/lib/data-service";
import {
  microsoftAuthConfig,
  isMicrosoftAuthConfigured,
  testMicrosoftAuth,
  type MicrosoftAuthTestResult,
} from "@/lib/auth-config";

export const Route = createFileRoute("/amministrazione")({
  head: () => ({ meta: [{ title: "Amministrazione — DR Portal" }] }),
  component: AmministrazionePage,
});

function AmministrazionePage() {
  const [status, setStatus] = useState<IntegrationStatus>(() => getIntegrationStatus());
  const [loading, setLoading] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true);
    try {
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

  const isSharePoint = status.mode === "sharepoint";

  return (
    <AppShell title="Amministrazione" subtitle="Configurazione sedi, ruoli e integrazioni">
      <div className="space-y-6">
        <Card>
          <CardHeader className="flex flex-row items-start justify-between gap-4 space-y-0">
            <div>
              <CardTitle className="flex items-center gap-2 text-base">
                {isSharePoint ? (
                  <CheckCircle2 className="h-5 w-5 text-status-present" />
                ) : (
                  <AlertTriangle className="h-5 w-5 text-status-break" />
                )}
                Stato integrazione SharePoint
              </CardTitle>
              <p className="text-xs text-muted-foreground mt-1">
                Origine dati corrente di DR Portal e diagnostica connessione a Microsoft 365.
              </p>
            </div>
            <Button size="sm" variant="outline" onClick={refresh} disabled={loading}>
              <RefreshCw className={`h-4 w-4 mr-2 ${loading ? "animate-spin" : ""}`} />
              Aggiorna
            </Button>
          </CardHeader>
          <CardContent className="grid gap-4 sm:grid-cols-2">
            <InfoRow label="Modalità attuale">
              <Badge
                variant={isSharePoint ? "default" : "secondary"}
                className={isSharePoint ? "bg-status-present text-white" : ""}
              >
                {isSharePoint ? "SharePoint" : "Mock (dati di esempio)"}
              </Badge>
            </InfoRow>
            <InfoRow label="Dipendenti caricati">
              <span className="font-semibold">{status.dipendentiCaricati}</span>
            </InfoRow>
            <InfoRow label="Ultimo aggiornamento">
              <span className="text-sm">
                {status.ultimoAggiornamento
                  ? status.ultimoAggiornamento.toLocaleTimeString("it-IT")
                  : "—"}
              </span>
            </InfoRow>
            <InfoRow label="Errori integrazione">
              {status.ultimoErrore ? (
                <span className="text-sm text-status-absent">{status.ultimoErrore}</span>
              ) : (
                <span className="text-sm text-muted-foreground">Nessuno</span>
              )}
            </InfoRow>
            {!isSharePoint && status.campiMancanti.length > 0 && (
              <div className="sm:col-span-2 rounded-md border border-dashed border-border p-3 text-xs text-muted-foreground">
                <p className="font-medium text-foreground mb-1">
                  Configurazione SharePoint incompleta
                </p>
                Variabili mancanti:{" "}
                <span className="font-mono">{status.campiMancanti.join(", ")}</span>.
                <br />
                Impostale nel file <span className="font-mono">.env</span> (prefisso{" "}
                <span className="font-mono">VITE_SP_*</span>) per attivare la modalità reale.
              </div>
            )}
          </CardContent>
        </Card>

        <PlaceholderSection
          Icon={Settings}
          title="Pannello amministrativo in arrivo"
          description="Gestione anagrafica dipendenti, sedi, turni e integrazione con Azure AD / SharePoint."
        />

        <MicrosoftAuthTestCard />
      </div>
    </AppShell>
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