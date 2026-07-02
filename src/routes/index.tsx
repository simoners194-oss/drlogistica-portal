import { createFileRoute } from "@tanstack/react-router";
import { Link } from "@tanstack/react-router";
import { Logo } from "@/components/Logo";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "DR Portal — Accedi" },
      { name: "description", content: "Portale presenze DR Logistica. Accedi con il tuo account Microsoft 365." },
    ],
  }),
  component: Index,
});

function Index() {
  return (
    <div className="min-h-screen grid lg:grid-cols-2 bg-background">
      <div className="hidden lg:flex flex-col justify-between p-12 text-primary-foreground" style={{ background: "var(--gradient-hero)" }}>
        <Logo variant="light" size={64} />
        <div>
          <h1 className="text-4xl font-semibold tracking-tight leading-tight">
            Gestisci le presenze<br />in modo semplice.
          </h1>
          <p className="mt-4 text-white/80 max-w-md">
            DR Portal centralizza timbrature, pause e reportistica per tutte le sedi DR Logistica.
          </p>
        </div>
        <div className="text-xs text-white/60">© DR Logistica — Powered by Microsoft 365</div>
      </div>

      <div className="flex items-center justify-center p-6 lg:p-12">
        <div className="w-full max-w-sm">
          <div className="lg:hidden mb-10 flex justify-center"><Logo size={72} /></div>
          <h2 className="text-2xl font-semibold text-foreground">Accedi</h2>
          <p className="text-sm text-muted-foreground mt-1">Usa il tuo account aziendale Microsoft 365</p>

          <div className="mt-8 space-y-3">
            <Link
              to="/dipendente"
              className="w-full flex items-center justify-center gap-3 h-11 rounded-md border border-border bg-card hover:bg-secondary transition-colors text-sm font-medium text-foreground shadow-sm"
            >
              <MicrosoftLogo />
              Accedi come Dipendente
            </Link>
            <Link
              to="/hr"
              className="w-full flex items-center justify-center gap-3 h-11 rounded-md text-primary-foreground text-sm font-medium shadow-[var(--shadow-elegant)] transition-transform hover:scale-[1.01]"
              style={{ background: "var(--gradient-primary)" }}
            >
              <MicrosoftLogo light />
              Accedi come HR
            </Link>
          </div>

          <p className="mt-6 text-[11px] text-muted-foreground text-center">
            Login simulato · integrazione SharePoint / Microsoft 365 disponibile a breve
          </p>
        </div>
      </div>
    </div>
  );
}

function MicrosoftLogo({ light = false }: { light?: boolean }) {
  const c = light ? ["#fff", "#fff", "#fff", "#fff"] : ["#F25022", "#7FBA00", "#00A4EF", "#FFB900"];
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <rect x="0" y="0" width="7" height="7" fill={c[0]} />
      <rect x="9" y="0" width="7" height="7" fill={c[1]} />
      <rect x="0" y="9" width="7" height="7" fill={c[2]} />
      <rect x="9" y="9" width="7" height="7" fill={c[3]} />
    </svg>
  );
}
