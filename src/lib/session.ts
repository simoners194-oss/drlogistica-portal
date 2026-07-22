// DR Portal — sessione client-side e ruoli.
//
// La sessione dell'utente vive in `localStorage` (chiave `dr:currentUser`),
// così sopravvive alla chiusura del browser/PWA (fondamentale sul telefono:
// tap su una notifica → si entra senza rifare login). Contiene solo dati non
// sensibili — id lista SharePoint, nome, cognome, sede e ruolo — sufficienti
// per pilotare navigazione, sidebar e filtri di UI. Le operazioni sensibili
// sono comunque validate lato server (cookie firmato httpOnly) tramite le
// server function SharePoint.

import { sedeTimbra, type SedeId } from "./mock-data";

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
  // Operatore/back-office (DR000): abilita l'inserimento di timbrature manuali.
  // Solo gating UI; l'autorizzazione reale è ri-verificata lato server.
  operatore: boolean;
  // Ore contrattuali settimanali (per l'avviso "monte ore giornaliero" in
  // Presenze). Opzionale: sessioni vecchie senza il campo → undefined.
  oreSettimanali?: number | null;
  // Codice dipendente (es. DR005): gating UI dei moduli riservati al direttore.
  // Sessioni vecchie senza il campo → null (serve re-login per vederli).
  codice?: string | null;
}

const KEY = "dr:currentUser";

export function readSession(): SessionUser | null {
  if (typeof window === "undefined") return null;
  try {
    // localStorage è la sede primaria; sessionStorage resta come fallback di
    // migrazione per le sessioni create prima di questo cambio.
    const raw = window.localStorage.getItem(KEY) ?? window.sessionStorage.getItem(KEY);
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
      operatore: Boolean(parsed.operatore),
      oreSettimanali: typeof parsed.oreSettimanali === "number" ? parsed.oreSettimanali : null,
      codice: typeof parsed.codice === "string" && parsed.codice ? parsed.codice : null,
    };
  } catch {
    return null;
  }
}

export function writeSession(u: SessionUser) {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(KEY, JSON.stringify(u));
    // Rimuovi l'eventuale copia legacy per evitare disallineamenti.
    window.sessionStorage.removeItem(KEY);
  } catch {
    /* ignore */
  }
}

export function clearSession() {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.removeItem(KEY);
    window.sessionStorage.removeItem(KEY);
  } catch {
    /* ignore */
  }
}

// Landing di default in base al ruolo. I dipendenti aprono direttamente le
// proprie presenze; responsabili e amministratori atterrano sulla dashboard.
export function defaultLandingFor(
  ruolo: Ruolo,
  sede?: SessionSede,
): "/presenze" | "/dashboard" | "/richieste" {
  if (ruolo === "dipendente") {
    // Dipendenti di sedi che non timbrano atterrano sulle Richieste.
    return sede && !sedeTimbra(sede) ? "/richieste" : "/presenze";
  }
  return "/dashboard";
}

export function canAccess(module: { roles?: readonly Ruolo[] }, ruolo: Ruolo): boolean {
  if (!module.roles || module.roles.length === 0) return true;
  return module.roles.includes(ruolo);
}
