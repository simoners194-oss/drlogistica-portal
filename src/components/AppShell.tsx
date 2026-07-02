import { Link, useRouterState } from "@tanstack/react-router";
import { Logo } from "./Logo";
import { LogOut } from "lucide-react";
import type { ReactNode } from "react";

export function AppShell({ children, role }: { children: ReactNode; role: "dipendente" | "hr" }) {
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  return (
    <div className="min-h-screen flex flex-col bg-background">
      <header className="border-b border-border bg-card/80 backdrop-blur sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <Link to="/dipendente" className="flex items-center gap-2">
            <Logo />
          </Link>
          <nav className="flex items-center gap-1 text-sm">
            <Link
              to="/dipendente"
              className={`px-3 py-1.5 rounded-md transition-colors ${pathname === "/dipendente" ? "bg-secondary text-secondary-foreground font-medium" : "text-muted-foreground hover:bg-secondary/60"}`}
            >
              Dipendente
            </Link>
            <Link
              to="/hr"
              className={`px-3 py-1.5 rounded-md transition-colors ${pathname === "/hr" ? "bg-secondary text-secondary-foreground font-medium" : "text-muted-foreground hover:bg-secondary/60"}`}
            >
              HR
            </Link>
            <Link to="/" className="ml-2 p-2 rounded-md text-muted-foreground hover:bg-secondary/60" aria-label="Esci">
              <LogOut className="h-4 w-4" />
            </Link>
          </nav>
        </div>
      </header>
      <main className="flex-1 max-w-6xl w-full mx-auto px-4 py-6">
        <div className="text-xs uppercase tracking-wider text-muted-foreground mb-1">
          {role === "dipendente" ? "Area Dipendente" : "Area HR"}
        </div>
        {children}
      </main>
    </div>
  );
}