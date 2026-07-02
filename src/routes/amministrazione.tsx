import { createFileRoute } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import { PlaceholderSection } from "@/components/PlaceholderSection";
import { Settings, CheckCircle2, AlertTriangle, RefreshCw } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { useEffect, useState, useCallback } from "react";
import { dataService, getIntegrationStatus, type IntegrationStatus } from "@/lib/data-service";

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