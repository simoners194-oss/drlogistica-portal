// DR Portal — menu a discesa MULTI-selezione con caselle di spunta, per i
// filtri incrociati senza affollare la barra: lista vuota = nessun filtro.
// Con ricerca incorporata: le voci sono tante (centinaia di controparti),
// si scrive "im" e resta IMILE.
import { useEffect, useRef, useState } from "react";
import { ChevronDown } from "lucide-react";

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

export function MultiSelect<T extends string | number>({
  label,
  tuttiLabel,
  selLabel,
  cercaLabel,
  opzioni,
  valori,
  onChange,
  className,
}: {
  label: string;
  tuttiLabel: string;
  /** Testo per "N selezionati" quando le voci scelte sono più di due. */
  selLabel: string;
  /** Segnaposto del campo di ricerca (es. "Cerca…"). */
  cercaLabel?: string;
  opzioni: { v: T; label: string }[];
  valori: T[];
  onChange: (v: T[]) => void;
  className?: string;
}) {
  const [open, setOpen] = useState(false);
  const [cerca, setCerca] = useState("");
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    setCerca("");
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, [open]);
  const riepilogo =
    valori.length === 0
      ? tuttiLabel
      : valori.length <= 2
        ? opzioni
            .filter((o) => valori.includes(o.v))
            .map((o) => o.label)
            .join(", ")
        : `${valori.length} ${selLabel}`;
  const q = cerca.trim().toLowerCase();
  const visibili = q ? opzioni.filter((o) => o.label.toLowerCase().includes(q)) : opzioni;
  return (
    <div ref={ref} className={`relative ${className ?? ""}`}>
      <label className="text-xs text-muted-foreground">{label}</label>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className={`${inputCls} flex items-center justify-between gap-2 text-left`}
      >
        <span className="truncate">{riepilogo}</span>
        <ChevronDown className="h-4 w-4 shrink-0 text-muted-foreground" />
      </button>
      {open && (
        <div className="absolute z-30 mt-1 w-full min-w-44 rounded-lg border border-border bg-card p-1 shadow-[var(--shadow-elegant)]">
          <input
            autoFocus
            value={cerca}
            onChange={(e) => setCerca(e.target.value)}
            placeholder={cercaLabel ?? "…"}
            className="mb-1 w-full rounded-md border border-border bg-background px-2 py-1 text-xs text-foreground"
          />
          <div className="max-h-64 overflow-auto">
            {!q && (
              <label className="flex cursor-pointer items-center gap-2 rounded-md px-2 py-1.5 text-sm text-foreground hover:bg-muted">
                <input
                  type="checkbox"
                  className="accent-primary"
                  checked={valori.length === 0}
                  onChange={() => onChange([])}
                />
                {tuttiLabel}
              </label>
            )}
            {visibili.map((o) => (
              <label
                key={String(o.v)}
                className="flex cursor-pointer items-center gap-2 rounded-md px-2 py-1.5 text-sm text-foreground hover:bg-muted"
              >
                <input
                  type="checkbox"
                  className="accent-primary"
                  checked={valori.includes(o.v)}
                  onChange={() =>
                    onChange(
                      valori.includes(o.v) ? valori.filter((x) => x !== o.v) : [...valori, o.v],
                    )
                  }
                />
                {o.label}
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
