import type { LucideIcon } from "lucide-react";

export function PlaceholderSection({
  Icon,
  title,
  description,
}: {
  Icon: LucideIcon;
  title: string;
  description: string;
}) {
  return (
    <div className="rounded-2xl border border-border bg-card p-10 md:p-14 text-center shadow-[var(--shadow-card)] animate-fade-in">
      <div className="inline-flex h-16 w-16 items-center justify-center rounded-2xl bg-primary/10 text-primary mb-5 shadow-sm">
        <Icon className="h-7 w-7" />
      </div>
      <h2 className="text-xl font-semibold text-foreground tracking-tight">{title}</h2>
      <p className="mt-2 text-sm text-muted-foreground max-w-md mx-auto leading-relaxed">{description}</p>
      <div className="mt-6 inline-flex items-center gap-2 rounded-full bg-secondary px-3 py-1 text-[11px] font-medium uppercase tracking-wider text-muted-foreground">
        <span className="h-1.5 w-1.5 rounded-full bg-primary" />
        Disponibile a breve
      </div>
    </div>
  );
}