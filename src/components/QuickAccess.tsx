import { Link } from "@tanstack/react-router";
import type { LucideIcon } from "lucide-react";

export interface QuickAccessItem {
  label: string;
  to?: string;
  Icon: LucideIcon;
  ready?: boolean;
  description?: string;
}

// Sezione "Accesso rapido" mostrata nelle landing per ciascun ruolo.
// I bottoni disabilitati (ready === false) mostrano il badge "In arrivo".
export function QuickAccess({
  title = "Accesso rapido",
  items,
}: {
  title?: string;
  items: QuickAccessItem[];
}) {
  return (
    <section
      aria-labelledby="quick-access-title"
      className="mt-6 rounded-2xl border border-border bg-card p-4 sm:p-5 shadow-[var(--shadow-card)]"
    >
      <div className="flex items-center justify-between mb-4">
        <h2
          id="quick-access-title"
          className="text-[15px] font-semibold text-foreground tracking-tight"
        >
          {title}
        </h2>
      </div>
      <div className="grid gap-3 grid-cols-2 md:grid-cols-4 lg:grid-cols-5">
        {items.map((it) => {
          const disabled = !it.ready || !it.to;
          const content = (
            <>
              <span className="h-10 w-10 shrink-0 rounded-xl bg-primary/10 text-primary flex items-center justify-center transition-colors group-hover:bg-primary group-hover:text-primary-foreground">
                <it.Icon className="h-5 w-5" />
              </span>
              <span className="min-w-0 text-left">
                <span className="block text-[13px] font-medium text-foreground truncate">
                  {it.label}
                </span>
                {!it.ready && (
                  <span className="mt-0.5 inline-block text-[9px] font-medium uppercase tracking-wider text-primary/80 bg-primary/10 px-1.5 py-0.5 rounded-full">
                    In arrivo
                  </span>
                )}
                {it.ready && it.description && (
                  <span className="block text-[11px] text-muted-foreground truncate mt-0.5">
                    {it.description}
                  </span>
                )}
              </span>
            </>
          );
          const base =
            "group flex items-center gap-3 rounded-xl border border-border bg-secondary/40 p-3 min-h-[60px] transition-all";
          if (disabled) {
            return (
              <button
                key={it.label}
                type="button"
                disabled
                aria-disabled="true"
                title="In arrivo — disponibile in una prossima versione"
                className={`${base} opacity-60 cursor-not-allowed`}
              >
                {content}
              </button>
            );
          }
          return (
            <Link
              key={it.label}
              to={it.to!}
              className={`${base} hover:bg-secondary hover:shadow-sm hover:-translate-y-0.5 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40`}
            >
              {content}
            </Link>
          );
        })}
      </div>
    </section>
  );
}