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
    <div className="rounded-2xl border border-dashed border-border bg-card p-10 text-center shadow-[var(--shadow-card)]">
      <div className="inline-flex h-14 w-14 items-center justify-center rounded-2xl bg-primary/10 text-primary mb-4">
        <Icon className="h-6 w-6" />
      </div>
      <h2 className="text-lg font-semibold text-foreground">{title}</h2>
      <p className="mt-2 text-sm text-muted-foreground max-w-md mx-auto">{description}</p>
      <div className="mt-6 inline-flex items-center gap-2 text-[11px] uppercase tracking-wider text-muted-foreground">
        <span className="h-1.5 w-1.5 rounded-full bg-status-break" />
        Sezione predisposta · integrazione SharePoint in arrivo
      </div>
    </div>
  );
}