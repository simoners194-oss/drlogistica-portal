import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { formatOra, type Dipendente } from "@/lib/mock-data";
import { useLang } from "@/lib/i18n";
import { computeOreOggi, formatDurata, type EventoTimbratura } from "@/lib/presenze-logic";
import { LogIn, Coffee, Hourglass, TrendingUp, ListChecks, Lock } from "lucide-react";

// Dettaglio giornaliero del dipendente per il portale Responsabili.
// Sola lettura: nessuna modifica alle timbrature è ancora consentita.
export function DettaglioDipendenteDialog({
  dipendente,
  open,
  onOpenChange,
}: {
  dipendente: Dipendente | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const { t, tVal } = useLang();
  const eventi = dipendente?.eventiOggi ?? [];
  const ore = computeOreOggi(eventi);

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-lg">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-3">
            <span className="h-10 w-10 rounded-full bg-primary/10 text-primary text-xs font-semibold flex items-center justify-center">
              {dipendente?.nome?.[0]}
              {dipendente?.cognome?.[0]}
            </span>
            <span className="min-w-0">
              <span className="block text-base font-semibold text-foreground truncate">
                {dipendente ? `${dipendente.nome} ${dipendente.cognome}` : ""}
              </span>
              <span className="block text-xs font-normal text-muted-foreground truncate">
                {dipendente?.ruolo}
              </span>
            </span>
          </DialogTitle>
          <DialogDescription>{t("dlg.subtitle")}</DialogDescription>
        </DialogHeader>

        {ore.chiusa && (
          <div className="flex items-start gap-2.5 rounded-xl border border-status-out/40 bg-status-out/5 p-3">
            <Lock className="h-4 w-4 text-status-out mt-0.5 shrink-0" />
            <div className="text-xs text-muted-foreground">{t("dlg.dayClosed")}</div>
          </div>
        )}

        <div className="grid grid-cols-2 gap-3">
          <MiniStat
            Icon={LogIn}
            label={t("evento.entrata")}
            value={ore.entrataOra ? formatOra(ore.entrataOra) : "—"}
          />
          <MiniStat
            Icon={LogIn}
            label={t("evento.uscita")}
            value={ore.uscitaOra ? formatOra(ore.uscitaOra) : "—"}
          />
          <MiniStat
            Icon={Coffee}
            label={t("presenze.totalBreak")}
            value={formatDurata(ore.pausaMinuti)}
            hint={ore.inPausa ? t("presenze.inProgress") : undefined}
          />
          <MiniStat
            Icon={Hourglass}
            label={t("presenze.workedHours")}
            value={formatDurata(ore.oreLavorateMinuti)}
          />
          <MiniStat
            Icon={TrendingUp}
            label={t("dash.kpiOvertime")}
            value={ore.oltreOrarioMinuti > 0 ? `+${formatDurata(ore.oltreOrarioMinuti)}` : "—"}
            highlight={ore.oltreOrarioMinuti > 0}
            className="col-span-2"
          />
        </div>

        <div className="rounded-xl border border-border bg-card p-4">
          <div className="flex items-center gap-2 text-[13px] text-muted-foreground mb-3">
            <ListChecks className="h-4 w-4" /> {t("presenze.todayEntries")}
          </div>
          {eventi.length === 0 ? (
            <div className="text-sm text-muted-foreground">{t("presenze.noneToday")}</div>
          ) : (
            <ol className="relative border-l border-border ml-2 space-y-3">
              {eventi.map((e, i) => (
                <li key={`${e.tipo}-${e.ora}-${i}`} className="pl-4 relative">
                  <span
                    className={`absolute -left-[7px] top-1.5 h-3 w-3 rounded-full ring-2 ring-card ${dotForEvento(e.tipo)}`}
                  />
                  <div className="flex items-baseline justify-between gap-3">
                    <div className="text-[14px] font-medium text-foreground">
                      {tVal("evento", e.tipo)}
                    </div>
                    <div className="text-[15px] font-semibold tabular-nums text-primary">
                      {formatOra(e.ora)}
                    </div>
                  </div>
                </li>
              ))}
            </ol>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}

function dotForEvento(t: EventoTimbratura): string {
  switch (t) {
    case "entrata":
      return "bg-primary";
    case "inizio-pausa":
      return "bg-status-break";
    case "fine-pausa":
      return "bg-status-present";
    case "uscita":
      return "bg-status-absent";
  }
}

function MiniStat({
  Icon,
  label,
  value,
  hint,
  highlight,
  className = "",
}: {
  Icon: typeof LogIn;
  label: string;
  value: string;
  hint?: string;
  highlight?: boolean;
  className?: string;
}) {
  return (
    <div
      className={`rounded-xl border p-3 ${highlight ? "border-status-out/40 bg-status-out/5" : "border-border bg-secondary/40"} ${className}`}
    >
      <div className="flex items-center justify-between gap-2">
        <span className="text-[10px] uppercase tracking-wider text-muted-foreground">{label}</span>
        <span
          className={`h-7 w-7 rounded-lg flex items-center justify-center ${highlight ? "bg-status-out/15 text-status-out" : "bg-primary/10 text-primary"}`}
        >
          <Icon className="h-3.5 w-3.5" />
        </span>
      </div>
      <div
        className={`mt-1.5 text-xl leading-none font-semibold tabular-nums tracking-tight ${highlight ? "text-status-out" : "text-foreground"}`}
      >
        {value}
      </div>
      {hint && <div className="mt-1 text-[11px] text-muted-foreground">{hint}</div>}
    </div>
  );
}
