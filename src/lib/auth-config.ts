// DR Portal — Microsoft 365 / Entra ID auth configuration
// -----------------------------------------------------------------------------
// L'autenticazione Microsoft 365 è opzionale e serve solo come TEST TECNICO:
// l'app funziona sempre in modalità mock o con accesso semplificato
// (selezione dipendente + PIN) anche senza queste variabili.
// -----------------------------------------------------------------------------

export interface MicrosoftAuthConfig {
  TENANT_ID: string;
  CLIENT_ID: string;
  REDIRECT_URI: string;
}

function envVar(key: string): string {
  try {
    const v = (import.meta as unknown as { env?: Record<string, string> }).env?.[key];
    return typeof v === "string" ? v.trim() : "";
  } catch {
    return "";
  }
}

export const microsoftAuthConfig: MicrosoftAuthConfig = {
  TENANT_ID: envVar("VITE_MS_TENANT_ID"),
  CLIENT_ID: envVar("VITE_MS_CLIENT_ID"),
  REDIRECT_URI:
    envVar("VITE_MS_REDIRECT_URI") ||
    (typeof window !== "undefined" ? window.location.origin : ""),
};

export function isMicrosoftAuthConfigured(
  cfg: MicrosoftAuthConfig = microsoftAuthConfig,
): boolean {
  return Boolean(cfg.TENANT_ID && cfg.CLIENT_ID);
}

export function missingMicrosoftAuthFields(
  cfg: MicrosoftAuthConfig = microsoftAuthConfig,
): (keyof MicrosoftAuthConfig)[] {
  const required: (keyof MicrosoftAuthConfig)[] = ["TENANT_ID", "CLIENT_ID"];
  return required.filter((k) => !cfg[k]);
}

export type AuthMode = "mock" | "simple" | "microsoft";

export interface MicrosoftAuthTestResult {
  ok: boolean;
  message: string;
  detail?: string;
}

// Test "soft" della configurazione MS: NON esegue un vero login (MSAL non è
// ancora integrato). Verifica che i parametri siano formalmente validi e che
// l'endpoint OpenID del tenant risponda. Serve solo come diagnostica in
// Amministrazione — in produzione verrà sostituito da un vero MSAL popup.
export async function testMicrosoftAuth(
  cfg: MicrosoftAuthConfig = microsoftAuthConfig,
): Promise<MicrosoftAuthTestResult> {
  if (!isMicrosoftAuthConfigured(cfg)) {
    return {
      ok: false,
      message: "Configurazione mancante",
      detail: `Variabili richieste: ${missingMicrosoftAuthFields(cfg)
        .map((f) => `VITE_MS_${f}`)
        .join(", ")}`,
    };
  }
  const guidRe = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guidRe.test(cfg.TENANT_ID) || !guidRe.test(cfg.CLIENT_ID)) {
    return {
      ok: false,
      message: "Formato non valido",
      detail: "TENANT_ID e CLIENT_ID devono essere GUID Microsoft (xxxxxxxx-xxxx-…).",
    };
  }
  try {
    const url = `https://login.microsoftonline.com/${cfg.TENANT_ID}/v2.0/.well-known/openid-configuration`;
    const res = await fetch(url, { method: "GET" });
    if (!res.ok) {
      return {
        ok: false,
        message: `Endpoint tenant non raggiungibile (HTTP ${res.status})`,
        detail: "Verifica il TENANT_ID nel portale Entra ID.",
      };
    }
    const data = (await res.json()) as { issuer?: string };
    return {
      ok: true,
      message: "Configurazione Microsoft valida",
      detail: data.issuer ? `Issuer: ${data.issuer}` : undefined,
    };
  } catch (err) {
    return {
      ok: false,
      message: "Errore di rete durante il test",
      detail: err instanceof Error ? err.message : String(err),
    };
  }
}