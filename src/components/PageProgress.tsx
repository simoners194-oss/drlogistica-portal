import { useRouterState } from "@tanstack/react-router";
import { useEffect, useState } from "react";

/**
 * Barra di avanzamento sottile mostrata durante le transizioni tra pagine.
 * Segue il router state (`pending`) con un piccolo debounce per evitare
 * flash su navigazioni istantanee.
 */
export function PageProgress() {
  const status = useRouterState({ select: (s) => s.status });
  const isLoading = status === "pending";
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    let t: ReturnType<typeof setTimeout> | undefined;
    if (isLoading) {
      t = setTimeout(() => setVisible(true), 120);
    } else {
      // piccolo delay per far percepire il completamento
      t = setTimeout(() => setVisible(false), 180);
    }
    return () => {
      if (t) clearTimeout(t);
    };
  }, [isLoading]);

  return (
    <div
      aria-hidden
      className={`pointer-events-none fixed left-0 right-0 top-0 z-[60] h-0.5 transition-opacity duration-200 ${
        visible ? "opacity-100" : "opacity-0"
      }`}
    >
      <div
        className="h-full w-1/3 rounded-r-full bg-gradient-to-r from-primary via-[color:var(--primary-glow)] to-primary shadow-[0_0_10px_var(--primary-glow)] animate-[pageProgress_1.2s_ease-in-out_infinite]"
      />
    </div>
  );
}