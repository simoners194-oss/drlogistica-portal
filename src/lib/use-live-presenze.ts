import { useEffect, useRef, useState, useCallback } from "react";
import { dataService, getIntegrationStatus } from "./data-service";
import type { Dipendente } from "./mock-data";

export function useLivePresenze(intervalMs = 15000) {
  const [data, setData] = useState<Dipendente[]>([]);
  const [lastUpdate, setLastUpdate] = useState<Date>(new Date());
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [tick, setTick] = useState(0);
  const mountedRef = useRef(true);

  const refresh = useCallback(async () => {
    setTick((t) => t + 1);
    const list = await dataService.getDipendenti();
    if (!mountedRef.current) return;
    setData(list);
    setLastUpdate(new Date());
    setError(getIntegrationStatus().ultimoErrore);
  }, []);

  useEffect(() => {
    mountedRef.current = true;
    const load = async () => {
      const list = await dataService.getDipendenti();
      if (!mountedRef.current) return;
      setData(list);
      setLastUpdate(new Date());
      setError(getIntegrationStatus().ultimoErrore);
      setLoading(false);
    };
    load();
    const t = setInterval(load, intervalMs);
    return () => {
      mountedRef.current = false;
      clearInterval(t);
    };
  }, [intervalMs, tick]);

  return { data, lastUpdate, error, refresh, loading };
}
