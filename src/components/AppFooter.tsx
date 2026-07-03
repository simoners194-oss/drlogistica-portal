import { APP_INFO } from "@/lib/version";
import { useSpStatus } from "@/lib/use-sp-status";

// Footer discreto e responsive presente in ogni schermata dell'app.
// Mostra a sinistra identità+versione+copyright, a destra lo stato
// dell'integrazione SharePoint (Online / Offline / Verifica…).
export function AppFooter({ variant = "app" }: { variant?: "app" | "auth" }) {
  const { status, message } = useSpStatus();

  const statusLabel =
    status === "online" ? "Online" : status === "offline" ? "Offline" : "Verifica…";
  const statusDotClass =
    status === "online"
      ? "bg-status-present"
      : status === "offline"
        ? "bg-status-absent"
        : "bg-muted-foreground";
  const statusTextClass =
    status === "online"
      ? "text-status-present"
      : status === "offline"
        ? "text-status-absent"
        : "text-muted-foreground";

  return (
    <footer
      role="contentinfo"
      className={
        variant === "app"
          ? "border-t border-border bg-card/70 backdrop-blur-sm text-[11px] text-muted-foreground"
          : "text-[11px] text-muted-foreground/80 px-4 py-3 border-t border-border/60"
      }
    >
      <div
        className={
          variant === "app"
            ? "max-w-[1400px] mx-auto px-5 md:px-8 py-3 grid grid-cols-[minmax(0,1fr)_auto] items-center gap-3"
            : "grid grid-cols-[minmax(0,1fr)_auto] items-center gap-3 max-w-md mx-auto"
        }
      >
        <div className="min-w-0 flex flex-wrap items-center gap-x-2 gap-y-0.5 leading-tight">
          <span className="font-medium text-foreground">{APP_INFO.name}</span>
          <span className="text-muted-foreground/70">·</span>
          <span className="tabular-nums">Versione {APP_INFO.version}</span>
          <span className="text-muted-foreground/70 hidden sm:inline">·</span>
          <span className="hidden sm:inline">{APP_INFO.copyright}</span>
        </div>
        <div
          className="shrink-0 inline-flex items-center gap-1.5"
          title={message ?? undefined}
          aria-label={`Stato integrazione SharePoint: ${statusLabel}`}
        >
          <span className="uppercase tracking-wider text-[10px] text-muted-foreground/80">
            Stato
          </span>
          <span className="relative inline-flex h-2 w-2">
            {status === "online" && (
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-status-present opacity-60" />
            )}
            <span className={`relative inline-flex h-2 w-2 rounded-full ${statusDotClass}`} />
          </span>
          <span className={`font-medium ${statusTextClass}`}>{statusLabel}</span>
        </div>
      </div>
    </footer>
  );
}