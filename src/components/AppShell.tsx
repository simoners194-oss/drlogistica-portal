import { useNavigate } from "@tanstack/react-router";
import { LogOut } from "lucide-react";
import { useMemo, type ReactNode } from "react";
import { AppSidebar } from "./AppSidebar";
import { Logo } from "./Logo";
import { SidebarProvider, SidebarTrigger } from "@/components/ui/sidebar";

function getCurrentUser() {
  if (typeof window === "undefined") return null;
  try {
    const raw = window.sessionStorage.getItem("dr:currentUser");
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { id?: string; nome?: string; cognome?: string } | null;
    return parsed?.id ? parsed : null;
  } catch {
    return null;
  }
}

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
  const user = useMemo(() => getCurrentUser(), []);

  const handleLogout = () => {
    try {
      window.sessionStorage.removeItem("dr:currentUser");
    } catch {
      /* ignore */
    }
    navigate({ to: "/", replace: true });
  };

  return (
    <SidebarProvider>
      <div className="min-h-screen flex w-full bg-background">
        <AppSidebar />
        <div className="flex-1 flex flex-col min-w-0">
          <header className="h-16 border-b border-border bg-card/80 backdrop-blur sticky top-0 z-10 flex items-center gap-3 px-4">
            <SidebarTrigger className="text-muted-foreground" />
            <div className="hidden md:block h-6 w-px bg-border" />
            <div className="md:hidden flex-1"><Logo size={36} /></div>
            <div className="hidden md:block flex-1 min-w-0">
              <h1 className="text-base font-semibold text-foreground truncate">{title}</h1>
              {subtitle && <p className="text-xs text-muted-foreground truncate">{subtitle}</p>}
            </div>
            <div className="flex items-center gap-2">
              {user && (
                <div className="hidden sm:flex items-center gap-2 px-3 h-9 rounded-md bg-secondary text-sm">
                  <div className="h-6 w-6 rounded-full bg-primary text-primary-foreground text-[10px] font-semibold flex items-center justify-center">
                    {user.nome?.[0] ?? ""}{user.cognome?.[0] ?? ""}
                  </div>
                  <span className="text-secondary-foreground">{user.nome} {user.cognome}</span>
                </div>
              )}
              <button
                type="button"
                onClick={handleLogout}
                className="inline-flex items-center gap-1.5 p-2 rounded-md text-muted-foreground hover:bg-secondary"
                aria-label="Esci"
              >
                <LogOut className="h-4 w-4" />
                <span className="hidden sm:inline text-sm">Esci</span>
              </button>
            </div>
          </header>
          <main className="flex-1 w-full max-w-[1400px] mx-auto px-4 md:px-6 py-6">
            <div className="md:hidden mb-4">
              <h1 className="text-xl font-semibold text-foreground">{title}</h1>
              {subtitle && <p className="text-xs text-muted-foreground">{subtitle}</p>}
            </div>
            {children}
          </main>
        </div>
      </div>
    </SidebarProvider>
  );
}