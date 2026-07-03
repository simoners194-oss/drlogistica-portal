// DR Portal — hook e bus di stato integrazione SharePoint.
//
// Fornisce al footer (e ad altri componenti) uno stato "online/offline"
// aggiornato dagli errori/successi del data-service, senza obbligare ogni
// pagina a ripolling la diagnostica. Le pagine che caricano dati reali
// (dashboard, presenze) alimentano automaticamente lo stato tramite
// `markSpOnline` / `markSpOffline` in data-service.

import { useEffect, useState } from "react";

export type SpConnectivity = "unknown" | "online" | "offline";

let current: SpConnectivity = "unknown";
let currentMessage: string | null = null;
const listeners = new Set<() => void>();

function emit() {
  for (const l of listeners) l();
}

export function setSpStatus(next: SpConnectivity, message?: string | null) {
  if (current === next && (message ?? null) === currentMessage) return;
  current = next;
  currentMessage = message ?? null;
  emit();
}

export function getSpStatus(): { status: SpConnectivity; message: string | null } {
  return { status: current, message: currentMessage };
}

export function useSpStatus(): { status: SpConnectivity; message: string | null } {
  const [state, setState] = useState(getSpStatus);
  useEffect(() => {
    const l = () => setState(getSpStatus());
    listeners.add(l);
    return () => {
      listeners.delete(l);
    };
  }, []);
  return state;
}