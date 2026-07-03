import { Skeleton } from "@/components/ui/skeleton";

/**
 * Skeleton per la pagina Presenze mentre viene recuperato il dipendente
 * corrente da SharePoint.
 */
export function PresenzeSkeleton() {
  return (
    <div className="animate-fade-in" aria-busy="true" aria-live="polite">
      <span className="sr-only">Caricamento presenze…</span>

      <div className="grid gap-4 md:gap-5 lg:grid-cols-3">
        <div
          className="lg:col-span-2 rounded-2xl p-5 sm:p-6 shadow-[var(--shadow-elegant)] text-primary-foreground"
          style={{ background: "var(--gradient-hero)" }}
        >
          <Skeleton className="h-4 w-40 rounded-md bg-white/25" />
          <Skeleton className="mt-3 h-14 sm:h-16 w-56 rounded-lg bg-white/25" />
          <div className="mt-4 flex flex-wrap gap-2">
            <Skeleton className="h-7 w-32 rounded-full bg-white/20" />
            <Skeleton className="h-7 w-28 rounded-full bg-white/20" />
          </div>
        </div>
        <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
          <Skeleton className="skeleton-shimmer h-4 w-32 rounded-md" />
          <Skeleton className="skeleton-shimmer mt-4 h-5 w-24 rounded-md" />
          <Skeleton className="skeleton-shimmer mt-2 h-10 w-32 rounded-md" />
        </div>
      </div>

      <div className="mt-5 md:mt-6 grid gap-3 sm:gap-4 grid-cols-2 lg:grid-cols-4">
        {Array.from({ length: 4 }).map((_, i) => (
          <div
            key={i}
            className="rounded-2xl border border-border bg-card p-4 sm:p-6 min-h-[148px] sm:min-h-[168px] flex flex-col justify-between shadow-[var(--shadow-card)]"
          >
            <Skeleton className="skeleton-shimmer h-12 w-12 sm:h-14 sm:w-14 rounded-xl" />
            <div className="space-y-2">
              <Skeleton className="skeleton-shimmer h-4 w-24 rounded-md" />
              <Skeleton className="skeleton-shimmer h-3 w-32 max-w-full rounded-md" />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}