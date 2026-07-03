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
  type LucideIcon,
} from "lucide-react";

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
  },
  {
    id: "presenze",
    title: "Presenze",
    url: "/presenze",
    icon: Clock,
    ready: true,
    description: "Timbrature, pause e stato personale.",
  },
  {
    id: "richieste",
    title: "Richieste",
    url: "/richieste",
    icon: FileText,
    ready: false,
    description: "Ferie, permessi e giustificativi.",
  },
  {
    id: "report",
    title: "Report",
    url: "/report",
    icon: BarChart3,
    ready: false,
    description: "Reportistica mensile e annuale.",
  },
  {
    id: "amministrazione",
    title: "Amministrazione",
    url: "/amministrazione",
    icon: Settings,
    ready: false,
    description: "Configurazione integrazioni e utenti.",
  },
] as const;