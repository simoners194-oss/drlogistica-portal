import { useEffect, useState } from "react";
import { dataService, getIntegrationStatus } from "./data-service";
import type { Dipendente } from "./mock-data";

// Hook di polling: rilegge i dipendenti ogni `intervalMs` millisecondi.
// Quando l'origine dati sarà SharePoint sarà sufficiente sostituire
// `dataService.getDipendenti()` con una subscription reale (Graph webhooks).
export function useLivePresenze(intervalMs = 15000) {
  const [data, setData] = useState<Dipendente[]>([]);
  const [lastUpdate, setLastUpdate] = useState<Date>(new Date());
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let mounted = true;
    const load = async () => {
      const list = await dataService.getDipendenti();
      if (!mounted) return;
      setData(list);
      setLastUpdate(new Date());
      setError(getIntegrationStatus().ultimoErrore);
    };
    load();
    const t = setInterval(load, intervalMs);
    return () => {
      mounted = false;
      clearInterval(t);
    };
  }, [intervalMs]);

  return { data, lastUpdate, error };
}