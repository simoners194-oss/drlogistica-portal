// DR Portal — sessione client-side e ruoli.
//
// La sessione dell'utente vive in `sessionStorage` (chiave `dr:currentUser`).
// Contiene solo dati non sensibili — id lista SharePoint, nome, cognome, sede
// e ruolo — sufficienti per pilotare navigazione, sidebar e filtri di UI.
// Le operazioni sensibili (creazione timbrature, letture privilegiate) sono
// comunque validate lato server tramite le server function SharePoint.

import type { SedeId } from "./mock-data";

// I tre ruoli ufficiali di DR Portal.
// - `dipendente`             — utente operativo, vede solo le proprie presenze
// - `responsabile`           — vede in sola lettura tutte le sedi
// - `amministratore_sistema` — accesso completo, gestione tecnica del portale
export type Ruolo = "dipendente" | "responsabile" | "amministratore_sistema";

export const RUOLO_LABEL: Record<Ruolo, string> = {
  dipendente: "Dipendente",
  responsabile: "Responsabile",
  amministratore_sistema: "Amministratore di sistema",
};

// Normalizza il valore SharePoint "Ruolo" in uno dei ruoli gestiti.
// Tollerante a maiuscole, spazi e varianti (es. "Resp.", "Admin",
// "Amministratore di sistema").
export function normalizeRuolo(raw: string | null | undefined): Ruolo {
  const s = (raw ?? "").toString().trim().toLowerCase();
  if (!s) return "dipendente";
  if (
    s.includes("sistema") ||
    s.startsWith("ammin") ||
    s.startsWith("admin") ||
    s === "sysadmin" ||
    s === "system"
  ) {
    return "amministratore_sistema";
  }
  if (s.startsWith("resp")) return "responsabile";
  return "dipendente";
}

// La sede della sessione può assumere il valore speciale "tutte" per
// l'account Amministratore di sistema (ADM001), che non è vincolato ad una
// sede operativa.
export type SessionSede = SedeId | "tutte";

export const SEDE_LABEL_TUTTE = "Tutte le sedi";

export interface SessionUser {
  id: string;
  nome: string;
  cognome: string;
  sede: SessionSede;
  ruolo: Ruolo;
  // Abilita all'approvazione di richieste (ferie/permessi/straordinari).
  // Usato SOLO per pilotare la UI (mostrare la coda approvatore). L'effettiva
  // autorizzazione a decidere è comunque ri-verificata lato server contro
  // SharePoint. Retrocompatibile: sessioni vecchie senza il campo → false.
  autorizza: boolean;
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
      sede: (parsed.sede as SessionSede) ?? "roma",
      ruolo: normalizeRuolo(parsed.ruolo),
      autorizza: Boolean(parsed.autorizza),
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