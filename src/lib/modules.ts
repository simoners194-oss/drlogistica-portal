// DR Portal — Registro moduli della piattaforma.
//
// Ogni voce del menu laterale è un "modulo" della piattaforma DR Portal.
// Per aggiungere un nuovo modulo in futuro, aggiungere una entry qui
// (con la relativa route) senza modificare AppSidebar o il layout.
//
//   - `ready: true`  → modulo attivo, accessibile dall'utente.
//   - `ready: false` → modulo predisposto, mostra il badge "In arrivo".

import {
  LayoutDashboard,
  Clock,
  FileText,
  BarChart3,
  Settings,
  Sparkles,
  ClipboardList,
  ShieldCheck,
  type LucideIcon,
} from "lucide-react";
import type { Ruolo } from "./session";

export type AppModule = {
  /** Identificatore stabile del modulo. */
  id: string;
  /** Etichetta visualizzata nel menu laterale. */
  title: string;
  /** Route TanStack associata. */
  url: string;
  /** Icona Lucide mostrata accanto al titolo. */
  icon: LucideIcon;
  /** Se `false`, la voce mostra il badge "In arrivo". */
  ready: boolean;
  /** Descrizione breve, usabile in tooltip o pagine sommario. */
  description?: string;
  /** Ruoli abilitati a vedere il modulo. Se omesso: visibile a tutti. */
  roles?: readonly Ruolo[];
  /** Se true, la voce è visibile solo agli utenti con capability operatore
   *  (flag `operatore` in sessione), a prescindere dal ruolo. */
  requiresOperatore?: boolean;
  /** Se true, la voce è visibile solo agli utenti con capability autorizza
   *  (approvatore/supervisore — flag `autorizza` in sessione). */
  requiresAutorizza?: boolean;
  /** Capability alternative: se presenti, la voce è visibile quando il ruolo è
   *  ammesso OPPURE l'utente possiede almeno una di queste capability. */
  orCapabilities?: readonly ("operatore" | "autorizza")[];
  /** Se true, la voce è visibile solo agli utenti la cui sede timbra
   *  (`sedeTimbra`). Le sedi che non timbrano non vedono il modulo. */
  richiedeTimbratura?: boolean;
};

export const APP_NAME = "DR Portal";
export const APP_TAGLINE = "Il portale aziendale di DR Logistica";

export const MODULES: readonly AppModule[] = [
  {
    id: "dashboard",
    title: "Dashboard",
    url: "/dashboard",
    icon: LayoutDashboard,
    ready: true,
    description: "Panoramica live delle presenze per sede.",
    roles: ["responsabile", "amministratore_sistema"],
  },
  {
    id: "presenze",
    title: "Presenze",
    url: "/presenze",
    icon: Clock,
    ready: true,
    description: "Timbrature, pause e stato personale.",
    roles: ["dipendente", "responsabile", "amministratore_sistema"],
    richiedeTimbratura: true,
  },
  {
    id: "richieste",
    title: "Richieste",
    url: "/richieste",
    icon: FileText,
    ready: true,
    description: "Ferie, permessi, straordinari, smart working, malattia e reperibilità.",
    roles: ["dipendente", "responsabile", "amministratore_sistema"],
  },
  {
    id: "gestione-timbrature",
    title: "Gestione timbrature",
    url: "/gestione-timbrature",
    icon: ClipboardList,
    ready: true,
    description: "Inserimento e correzione timbrature manuali (operatore).",
    roles: ["dipendente", "responsabile", "amministratore_sistema"],
    requiresOperatore: true,
  },
  {
    id: "supervisione",
    title: "Supervisione",
    url: "/supervisione",
    icon: ShieldCheck,
    ready: true,
    description: "Richieste approvate e timbrature manuali (supervisore).",
    roles: ["dipendente", "responsabile", "amministratore_sistema"],
    requiresAutorizza: true,
  },
  {
    id: "report",
    title: "Rendiconto",
    url: "/report",
    icon: BarChart3,
    ready: true,
    description: "Rendiconto mensile ore per dipendente.",
    roles: ["responsabile", "amministratore_sistema"],
    orCapabilities: ["operatore", "autorizza"],
  },
  {
    id: "amministrazione",
    title: "Amministrazione",
    url: "/amministrazione",
    icon: Settings,
    ready: true,
    description: "Configurazione integrazioni e utenti.",
    roles: ["amministratore_sistema"],
  },
  {
    id: "novita",
    title: "Novità",
    url: "/novita",
    icon: Sparkles,
    ready: true,
    description: "Registro delle versioni e delle novità di DR Portal.",
  },
] as const;
