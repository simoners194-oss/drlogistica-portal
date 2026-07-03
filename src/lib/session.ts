// DR Portal — sessione client-side e ruoli.
//
// La sessione dell'utente vive in `sessionStorage` (chiave `dr:currentUser`).
// Contiene solo dati non sensibili — id lista SharePoint, nome, cognome, sede
// e ruolo — sufficienti per pilotare navigazione, sidebar e filtri di UI.
// Le operazioni sensibili (creazione timbrature, letture privilegiate) sono
// comunque validate lato server tramite le server function SharePoint.

import type { SedeId } from "./mock-data";

export type Ruolo = "dipendente" | "responsabile" | "amministratore";

export const RUOLO_LABEL: Record<Ruolo, string> = {
  dipendente: "Dipendente",
  responsabile: "Responsabile",
  amministratore: "Amministratore",
};

// Normalizza il valore SharePoint "Ruolo" in uno dei ruoli gestiti.
// Tollerante a maiuscole, spazi e varianti (es. "Resp.", "Admin").
export function normalizeRuolo(raw: string | null | undefined): Ruolo {
  const s = (raw ?? "").toString().trim().toLowerCase();
  if (!s) return "dipendente";
  if (s.startsWith("ammin") || s.startsWith("admin")) return "amministratore";
  if (s.startsWith("resp")) return "responsabile";
  return "dipendente";
}

export interface SessionUser {
  id: string;
  nome: string;
  cognome: string;
  sede: SedeId;
  ruolo: Ruolo;
}

const KEY = "dr:currentUser";

export function readSession(): SessionUser | null {
  if (typeof window === "undefined") return null;
  try {
    const raw = window.sessionStorage.getItem(KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<SessionUser> | null;
    if (!parsed?.id) return null;
    return {
      id: String(parsed.id),
      nome: parsed.nome ?? "",
      cognome: parsed.cognome ?? "",
      sede: (parsed.sede as SedeId) ?? "roma",
      ruolo: normalizeRuolo(parsed.ruolo),
    };
  } catch {
    return null;
  }
}

export function writeSession(u: SessionUser) {
  if (typeof window === "undefined") return;
  try {
    window.sessionStorage.setItem(KEY, JSON.stringify(u));
  } catch {
    /* ignore */
  }
}

export function clearSession() {
  if (typeof window === "undefined") return;
  try {
    window.sessionStorage.removeItem(KEY);
  } catch {
    /* ignore */
  }
}

// Landing di default in base al ruolo. I dipendenti aprono direttamente le
// proprie presenze; responsabili e amministratori atterrano sulla dashboard.
export function defaultLandingFor(ruolo: Ruolo): "/presenze" | "/dashboard" {
  return ruolo === "dipendente" ? "/presenze" : "/dashboard";
}

export function canAccess(module: { roles?: readonly Ruolo[] }, ruolo: Ruolo): boolean {
  if (!module.roles || module.roles.length === 0) return true;
  return module.roles.includes(ruolo);
}