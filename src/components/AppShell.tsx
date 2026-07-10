import { useNavigate } from "@tanstack/react-router";
import { LogOut } from "lucide-react";
import { useEffect, useState, type ReactNode } from "react";
import { AppSidebar } from "./AppSidebar";
import { Logo } from "./Logo";
import { PageProgress } from "./PageProgress";
import { AppFooter } from "./AppFooter";
import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar";
import { clearSession, readSession, RUOLO_LABEL, type SessionUser } from "@/lib/session";

export function AppShell({
  children,
  title,
  subtitle,
}: {
  children: ReactNode;
  title: string;
  subtitle?: string;
}) {
  const navigate = useNavigate();
  // Deferred to client only to avoid SSR/CSR hydration mismatch on the
  // user pill (sessionStorage is not available during prerender).
  const [user, setUser] = useState<SessionUser | null>(null);
  useEffect(() => {
    setUser(readSession());
  }, []);

  const handleLogout = () => {
    if (typeof window !== "undefined" && !window.confirm("Sei sicuro di voler uscire?")) {
      return;
    }
    clearSession();
    navigate({ to: "/", replace: true });
  };

  return (
    <SidebarProvider>
      <PageProgress />
      <div className="min-h-screen flex w-full bg-background">
        <AppSidebar />
        <div className="flex-1 flex flex-col min-w-0">
          <header className="h-18 border-b border-border bg-card/85 backdrop-blur-md sticky top-0 z-10 flex items-center gap-3 px-4 md:px-6">
            <SidebarTrigger className="text-muted-foreground" />
            <div className="hidden md:block h-6 w-px bg-border" />
            <div className="md:hidden flex-1"><Logo size={48} /></div>
            <div className="hidden md:block flex-1 min-w-0">
              <h1 className="text-[15px] font-semibold text-foreground truncate tracking-tight">{title}</h1>
              {subtitle && <p className="text-xs text-muted-foreground truncate">{subtitle}</p>}
            </div>
            <div className="flex items-center gap-2">
              {user && (
                <div className="hidden sm:flex items-center gap-2 px-3 h-9 rounded-full bg-secondary text-sm animate-fade-in">
                  <div className="h-6 w-6 rounded-full bg-gradient-to-br from-primary to-[color:var(--primary-glow)] text-primary-foreground text-[10px] font-semibold flex items-center justify-center shadow-sm">
                    {user.nome?.[0] ?? ""}{user.cognome?.[0] ?? ""}
                  </div>
                  <span className="text-secondary-foreground">{user.nome} {user.cognome}</span>
                  <span className="text-[10px] uppercase tracking-wider text-muted-foreground border-l border-border pl-2 ml-0.5">
                    {RUOLO_LABEL[user.ruolo]}
                  </span>
                </div>
              )}
              <button
                type="button"
                onClick={handleLogout}
                className="inline-flex items-center gap-1.5 px-2.5 h-9 rounded-full text-muted-foreground hover:bg-secondary hover:text-foreground transition-colors"
                aria-label="Esci"
              >
                <LogOut className="h-4 w-4" />
                <span className="hidden sm:inline text-sm">Esci</span>
              </button>
            </div>
          </header>
          <main className="flex-1 w-full max-w-[1400px] mx-auto px-5 md:px-8 py-6 md:py-8 animate-fade-in">
            <div className="md:hidden mb-5">
              <h1 className="text-[26px] leading-tight font-semibold text-foreground tracking-tight">{title}</h1>
              {subtitle && <p className="text-xs text-muted-foreground mt-0.5">{subtitle}</p>}
            </div>
            {children}
          </main>
          <AppFooter />
        </div>
      </div>
    </SidebarProvider>
  );
}