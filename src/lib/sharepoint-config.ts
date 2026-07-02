// DR Portal — SharePoint / Microsoft Graph configuration
// -----------------------------------------------------------------------------
// Questo file centralizza i parametri necessari per collegare DR Portal alle
// liste SharePoint di Microsoft 365. Finché almeno uno dei valori richiesti è
// vuoto, l'app continua a funzionare in modalità mock (vedi `data-service.ts`).
//
// Nel prossimo step queste variabili verranno lette da import.meta.env /
// process.env (per il lato server) e usate da un client Microsoft Graph.
// Per ora restano placeholder in modo che il progetto compili e funzioni con i
// dati di esempio.
// -----------------------------------------------------------------------------

export interface SharePointConfig {
  TENANT_ID: string;
  CLIENT_ID: string;
  SITE_ID: string;
  LIST_DIPENDENTI_ID: string;
  LIST_TIMBRATURE_ID: string;
}

// Legge una variabile d'ambiente Vite in modo tollerante (senza rompere il
// build se non è definita). In produzione queste verranno impostate dal
// tenant IT di DR Logistica.
function envVar(key: string): string {
  try {
    // import.meta.env è iniettato da Vite in build time.
    const v = (import.meta as unknown as { env?: Record<string, string> }).env?.[key];
    return typeof v === "string" ? v.trim() : "";
  } catch {
    return "";
  }
}

export const sharepointConfig: SharePointConfig = {
  TENANT_ID: envVar("VITE_SP_TENANT_ID"),
  CLIENT_ID: envVar("VITE_SP_CLIENT_ID"),
  SITE_ID: envVar("VITE_SP_SITE_ID"),
  LIST_DIPENDENTI_ID: envVar("VITE_SP_LIST_DIPENDENTI_ID"),
  LIST_TIMBRATURE_ID: envVar("VITE_SP_LIST_TIMBRATURE_ID"),
};

// La configurazione è considerata "attiva" solo se TUTTI i campi obbligatori
// sono valorizzati. In caso contrario il data-service resta in modalità mock.
export function isSharePointConfigured(cfg: SharePointConfig = sharepointConfig): boolean {
  return Boolean(
    cfg.TENANT_ID &&
      cfg.CLIENT_ID &&
      cfg.SITE_ID &&
      cfg.LIST_DIPENDENTI_ID &&
      cfg.LIST_TIMBRATURE_ID,
  );
}

// Elenco dei campi richiesti — utile per la UI di diagnostica in Amministrazione.
export function missingSharePointFields(
  cfg: SharePointConfig = sharepointConfig,
): (keyof SharePointConfig)[] {
  return (Object.keys(cfg) as (keyof SharePointConfig)[]).filter((k) => !cfg[k]);
}

export type IntegrationMode = "mock" | "sharepoint";

export function currentIntegrationMode(): IntegrationMode {
  return isSharePointConfigured() ? "sharepoint" : "mock";
}