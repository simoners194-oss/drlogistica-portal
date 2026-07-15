import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect } from "react";
import { AppShell } from "@/components/AppShell";
import { Sparkles, ShieldCheck, Wrench, TrendingUp } from "lucide-react";
import { readSession } from "@/lib/session";
import { APP_INFO, formatReleaseDate } from "@/lib/version";
import { RELEASES, TAG_LABEL, type ReleaseTag } from "@/lib/releases";

export const Route = createFileRoute("/novita")({
  head: () => ({
    meta: [
      { title: "Novità — DR Portal" },
      {
        name: "description",
        content: "Registro delle novità e delle versioni pubblicate su DR Portal.",
      },
    ],
  }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    const s = readSession();
    if (!s) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: NovitaPage,
});

const TAG_ICON: Record<ReleaseTag, typeof Sparkles> = {
  feature: Sparkles,
  improvement: TrendingUp,
  fix: Wrench,
  security: ShieldCheck,
};

const TAG_CLASS: Record<ReleaseTag, string> = {
  feature: "bg-primary/10 text-primary",
  improvement: "bg-status-present/15 text-status-present",
  fix: "bg-status-break/15 text-status-break",
  security: "bg-status-out/15 text-status-out",
};

function NovitaPage() {
  const sorted = [...RELEASES].sort((a, b) => b.date.localeCompare(a.date));
  const latest = sorted[0];

  // Segna le novità come "viste" per la versione corrente: spegne il popup.
  useEffect(() => {
    try {
      window.localStorage.setItem("dr:novita:lastVersion", APP_INFO.version);
    } catch {
      /* localStorage non disponibile */
    }
  }, []);

  return (
    <AppShell title="Novità" subtitle={`Registro delle versioni pubblicate di ${APP_INFO.name}`}>
      {latest && (
        <div className="mb-6 rounded-2xl border border-primary/30 bg-primary/5 p-4 sm:p-5 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-primary text-[11px] font-semibold uppercase tracking-wider">
            <Sparkles className="h-3.5 w-3.5" />
            Ultima versione
          </div>
          <div className="mt-1 flex flex-wrap items-baseline gap-x-2 gap-y-0.5">
            <span className="text-2xl font-semibold text-foreground tracking-tight tabular-nums">
              v{latest.version}
            </span>
            {latest.codename && (
              <span className="text-sm text-muted-foreground">· {latest.codename}</span>
            )}
            {latest.author && (
              <span className="text-sm text-muted-foreground">· a cura di {latest.author}</span>
            )}
            <span className="text-xs text-muted-foreground ml-auto">
              Pubblicata il {formatReleaseDate(latest.date)}
            </span>
          </div>
        </div>
      )}

      <ol className="relative space-y-8 border-l border-border pl-5 sm:pl-6">
        {sorted.map((r) => (
          <li key={r.version} className="relative">
            <span className="absolute -left-[27px] sm:-left-[31px] top-1 h-3 w-3 rounded-full bg-primary ring-4 ring-background" />
            <header className="flex flex-wrap items-baseline gap-x-2 gap-y-0.5">
              <h2 className="text-lg font-semibold text-foreground tracking-tight tabular-nums">
                Versione {r.version}
              </h2>
              {r.codename && <span className="text-sm text-muted-foreground">· {r.codename}</span>}
              {r.author && (
                <span className="text-sm text-muted-foreground">· a cura di {r.author}</span>
              )}
              <span className="text-xs text-muted-foreground ml-auto">
                {formatReleaseDate(r.date)}
              </span>
            </header>
            <ul className="mt-3 space-y-2">
              {r.entries.map((e, i) => {
                const tag = e.tag ?? "feature";
                const Icon = TAG_ICON[tag];
                return (
                  <li
                    key={i}
                    className="rounded-xl border border-border bg-card p-3 sm:p-4 shadow-[var(--shadow-card)]"
                  >
                    <div className="flex items-start gap-3">
                      <span
                        className={`h-8 w-8 shrink-0 rounded-lg flex items-center justify-center ${TAG_CLASS[tag]}`}
                      >
                        <Icon className="h-4 w-4" />
                      </span>
                      <div className="min-w-0">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-[14px] font-medium text-foreground">{e.title}</span>
                          <span className="text-[10px] font-medium uppercase tracking-wider text-muted-foreground bg-secondary rounded-full px-1.5 py-0.5">
                            {TAG_LABEL[tag]}
                          </span>
                        </div>
                        {e.description && (
                          <p className="mt-1 text-[13px] text-muted-foreground leading-relaxed">
                            {e.description}
                          </p>
                        )}
                      </div>
                    </div>
                  </li>
                );
              })}
            </ul>
          </li>
        ))}
      </ol>

      <p className="mt-8 text-center text-[11px] text-muted-foreground">
        Build {APP_INFO.build} — {APP_INFO.copyright}
      </p>
      <p className="mt-1 text-center text-[11px] text-muted-foreground/70">
        Realizzato da Simone Russo
      </p>
    </AppShell>
  );
}
