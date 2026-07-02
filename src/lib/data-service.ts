// DR Portal — data service
// -----------------------------------------------------------------------------
// Questo modulo è l'unico punto di accesso ai dati di presenza.
// Attualmente restituisce dati mock in memoria; è già strutturato per essere
// sostituito da un client SharePoint / Microsoft Graph senza toccare le UI.
//
// Per collegare SharePoint in futuro basterà:
//   1. Implementare le stesse funzioni (getDipendenti, getSedi, timbra…)
//      leggendo/scrivendo su liste SharePoint via Microsoft Graph.
//   2. Sostituire l'export `dataService` con l'implementazione reale
//      (per esempio in base a import.meta.env.VITE_DR_DATA_SOURCE).
// -----------------------------------------------------------------------------

import {
  DIPENDENTI as MOCK_DIPENDENTI,
  SEDI,
  type Dipendente,
  type SedeId,
  type Timbratura,
} from "./mock-data";

// Stato in memoria (mock). In produzione verrà rimpiazzato da chiamate a
// SharePoint. Cloniamo l'array per poter mutare le timbrature durante la
// sessione.
let state: Dipendente[] = MOCK_DIPENDENTI.map((d) => ({ ...d }));

export interface DataService {
  getSedi(): Promise<typeof SEDI>;
  getDipendenti(): Promise<Dipendente[]>;
  getDipendente(id: string): Promise<Dipendente | undefined>;
  timbra(dipendenteId: string, tipo: Timbratura["tipo"]): Promise<Dipendente>;
}

export const dataService: DataService = {
  async getSedi() {
    return SEDI;
  },
  async getDipendenti() {
    return state.map((d) => ({ ...d }));
  },
  async getDipendente(id) {
    const d = state.find((x) => x.id === id);
    return d ? { ...d } : undefined;
  },
  async timbra(dipendenteId, tipo) {
    const idx = state.findIndex((d) => d.id === dipendenteId);
    if (idx < 0) throw new Error("Dipendente non trovato");
    const nuova: Timbratura = { tipo, ora: new Date().toISOString() };
    const prev = state[idx];
    const stato =
      tipo === "entrata" || tipo === "fine-pausa"
        ? "presente"
        : tipo === "inizio-pausa"
          ? "pausa"
          : "uscito";
    const entrataOra =
      tipo === "entrata" ? nuova.ora : prev.entrataOra;
    state[idx] = { ...prev, stato, ultimaTimbratura: nuova, entrataOra };
    return { ...state[idx] };
  },
};

// Filtra per sede senza duplicare la logica nelle pagine.
export function bySede(list: Dipendente[], sede: SedeId) {
  return list.filter((d) => d.sede === sede);
}

// Statistiche aggregate riusate sia dalla dashboard sia dai widget live.
export function aggregate(list: Dipendente[]) {
  const attivi = list.length;
  const presenti = list.filter((d) => d.stato === "presente").length;
  const pausa = list.filter((d) => d.stato === "pausa").length;
  const usciti = list.filter((d) => d.stato === "uscito").length;
  const assenti = list.filter((d) => d.stato === "non-timbrato").length;
  const straordinari = list.filter(
    (d) => (d.straordinariMinuti ?? 0) > 0 && d.stato !== "uscito",
  ).length;
  const ritardi = list.filter((d) => (d.ritardoMinuti ?? 0) > 0).length;
  return { attivi, presenti, pausa, usciti, assenti, straordinari, ritardi };
}

// Stato "visivo" richiesto dalla dashboard live:
//   verde = presente · giallo = in pausa · rosso = assente · blu = oltre orario
export type DisplayStato = "presente" | "pausa" | "assente" | "oltre";

export function displayStato(d: Dipendente): DisplayStato {
  if (d.stato === "pausa") return "pausa";
  if (d.stato === "non-timbrato" || d.stato === "uscito") return "assente";
  if ((d.straordinariMinuti ?? 0) > 0) return "oltre";
  return "presente";
}

export const DISPLAY_LABEL: Record<DisplayStato, string> = {
  presente: "Presente",
  pausa: "In pausa",
  assente: "Assente",
  oltre: "Oltre orario",
};

export const DISPLAY_DOT: Record<DisplayStato, string> = {
  presente: "bg-status-present",
  pausa: "bg-status-break",
  assente: "bg-status-absent",
  oltre: "bg-status-out",
};

// Calcola le ore lavorate oggi. Semplice: from entrataOra a "ora"
// (in un client reale useremo la somma degli intervalli tra entrate/uscite).
export function oreLavorateOggi(d: Dipendente, now = new Date()): string {
  if (!d.entrataOra) return "0h 00m";
  const start = new Date(d.entrataOra).getTime();
  const end =
    d.stato === "uscito" && d.ultimaTimbratura
      ? new Date(d.ultimaTimbratura.ora).getTime()
      : now.getTime();
  const min = Math.max(0, Math.floor((end - start) / 60000));
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${h}h ${m.toString().padStart(2, "0")}m`;
}