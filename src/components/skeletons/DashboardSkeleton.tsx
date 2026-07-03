import { Skeleton } from "@/components/ui/skeleton";

/**
 * Skeleton che riproduce fedelmente la struttura della Dashboard HR
 * mentre i dati SharePoint vengono caricati.
 */
export function DashboardSkeleton() {
  return (
    <div className="animate-fade-in" aria-busy="true" aria-live="polite">
      <span className="sr-only">Caricamento dashboard…</span>

      {/* Sintesi presenze */}
      <section className="mb-6 rounded-2xl border border-border bg-card p-4 sm:p-5 md:p-6 shadow-[var(--shadow-card)]">
        <div className="flex items-start justify-between gap-4 mb-5">
          <div className="space-y-2 min-w-0 flex-1">
            <Skeleton className="skeleton-shimmer h-6 w-64 max-w-full rounded-md" />
            <Skeleton className="skeleton-shimmer h-3 w-40 max-w-full rounded-md" />
          </div>
          <Skeleton className="skeleton-shimmer h-11 sm:h-9 w-11 sm:w-32 rounded-full" />
        </div>
        <div className="grid gap-3 md:gap-4 grid-cols-1 sm:grid-cols-2">
          {Array.from({ length: 2 }).map((_, i) => (
            <div
              key={i}
              className="rounded-xl border border-border bg-secondary/40 p-4 min-h-[84px] flex items-center justify-between gap-3"
            >
              <div className="flex items-center gap-3 min-w-0 flex-1">
                <Skeleton className="skeleton-shimmer h-12 w-12 shrink-0 rounded-xl" />
                <div className="space-y-2 flex-1 min-w-0">
                  <Skeleton className="skeleton-shimmer h-4 w-28 rounded-md" />
                  <Skeleton className="skeleton-shimmer h-3 w-40 max-w-full rounded-md" />
                </div>
              </div>
              <div className="flex flex-col items-end gap-2 shrink-0">
                <Skeleton className="skeleton-shimmer h-8 w-16 rounded-md" />
                <Skeleton className="skeleton-shimmer h-2.5 w-14 rounded-md" />
              </div>
            </div>
          ))}
        </div>
      </section>

      {/* KPI grid */}
      <div className="grid gap-3 grid-cols-2 md:grid-cols-5">
        {Array.from({ length: 5 }).map((_, i) => (
          <div
            key={i}
            className={`rounded-xl border border-border bg-card p-4 shadow-[var(--shadow-card)] ${
              i === 4 ? "col-span-2 md:col-span-1" : ""
            }`}
          >
            <div className="flex items-center justify-between gap-2">
              <Skeleton className="skeleton-shimmer h-3 w-20 rounded-md" />
              <Skeleton className="skeleton-shimmer h-9 w-9 rounded-lg" />
            </div>
            <Skeleton className="skeleton-shimmer mt-3 h-8 w-14 rounded-md" />
          </div>
        ))}
      </div>

      {/* Sede panels */}
      <div className="mt-6 grid gap-4 lg:grid-cols-2">
        {Array.from({ length: 2 }).map((_, i) => (
          <section
            key={i}
            className="rounded-xl border border-border bg-card shadow-[var(--shadow-card)] overflow-hidden"
          >
            <header className="flex items-center justify-between gap-3 px-4 sm:px-5 py-4 border-b border-border">
              <div className="space-y-2 min-w-0 flex-1">
                <Skeleton className="skeleton-shimmer h-4 w-32 rounded-md" />
                <Skeleton className="skeleton-shimmer h-3 w-40 max-w-full rounded-md" />
              </div>
              <div className="space-y-2 items-end flex flex-col">
                <Skeleton className="skeleton-shimmer h-6 w-14 rounded-md" />
                <Skeleton className="skeleton-shimmer h-2.5 w-10 rounded-md" />
              </div>
            </header>
            <div className="flex items-center gap-2 px-3 py-2.5 border-b border-border">
              {Array.from({ length: 4 }).map((__, j) => (
                <Skeleton key={j} className="skeleton-shimmer h-7 w-16 rounded-full" />
              ))}
            </div>
            <ul className="divide-y divide-border">
              {Array.from({ length: 4 }).map((__, j) => (
                <li key={j} className="flex items-center gap-3 px-4 sm:px-5 py-3.5">
                  <Skeleton className="skeleton-shimmer h-10 w-10 rounded-full shrink-0" />
                  <div className="flex-1 min-w-0 space-y-2">
                    <Skeleton className="skeleton-shimmer h-3.5 w-40 max-w-full rounded-md" />
                    <Skeleton className="skeleton-shimmer h-3 w-56 max-w-full rounded-md" />
                  </div>
                  <Skeleton className="skeleton-shimmer h-6 w-20 rounded-full shrink-0" />
                </li>
              ))}
            </ul>
          </section>
        ))}
      </div>
    </div>
  );
}