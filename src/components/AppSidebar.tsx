import { Link, useRouterState } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarHeader,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  useSidebar,
} from "@/components/ui/sidebar";
import { Logo } from "./Logo";
import { MODULES } from "@/lib/modules";
import { canAccess, readSession, type Ruolo } from "@/lib/session";

export function AppSidebar() {
  const { state } = useSidebar();
  const collapsed = state === "collapsed";
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  // Il ruolo viene letto client-side (sessionStorage) dopo il mount per
  // evitare mismatch di hydration con SSR/prerender.
  const [ruolo, setRuolo] = useState<Ruolo | null>(null);
  const [operatore, setOperatore] = useState(false);
  const [autorizza, setAutorizza] = useState(false);
  useEffect(() => {
    const s = readSession();
    setRuolo(s?.ruolo ?? null);
    setOperatore(s?.operatore ?? false);
    setAutorizza(s?.autorizza ?? false);
  }, [pathname]);

  // Finché il ruolo non è noto, mostra solo le voci pubbliche a tutti i
  // ruoli (Presenze, Richieste) per evitare "flash" del menu completo. Le voci
  // con `requiresOperatore` compaiono solo per gli operatori (DR000).
  const visibleModules = MODULES.filter((m) => {
    const roleOk = ruolo ? canAccess(m, ruolo) : canAccess(m, "dipendente");
    if (!roleOk) return false;
    if (m.requiresOperatore && !operatore) return false;
    if (m.requiresAutorizza && !autorizza) return false;
    return true;
  });

  return (
    <Sidebar collapsible="icon">
      <SidebarHeader className="h-20 border-b border-sidebar-border flex items-center justify-center px-3 py-2">
        {collapsed ? <img src="/favicon.png" alt="DR" className="h-10 w-10" /> : <Logo size={64} />}
      </SidebarHeader>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Moduli</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {visibleModules.map((item) => {
                const active = pathname === item.url;
                // Voci "In arrivo": non cliccabili, aspetto attenuato.
                if (!item.ready) {
                  return (
                    <SidebarMenuItem key={item.title}>
                      <SidebarMenuButton
                        tooltip={`${item.title} — In arrivo`}
                        aria-disabled="true"
                        className="opacity-60 cursor-not-allowed hover:bg-transparent"
                        onClick={(e) => e.preventDefault()}
                      >
                        <item.icon className="h-4 w-4 shrink-0" />
                        {!collapsed && (
                          <span className="flex-1 flex items-center justify-between">
                            {item.title}
                            <span className="text-[9px] font-medium uppercase tracking-wider text-primary/80 bg-primary/10 px-1.5 py-0.5 rounded-full">
                              In arrivo
                            </span>
                          </span>
                        )}
                      </SidebarMenuButton>
                    </SidebarMenuItem>
                  );
                }
                return (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton
                      asChild
                      isActive={active}
                      tooltip={item.title}
                      className={
                        active
                          ? "bg-primary/10 text-primary font-medium border-l-2 border-primary rounded-l-none pl-[calc(0.5rem-2px)]"
                          : ""
                      }
                    >
                      <Link to={item.url} className="flex items-center gap-2">
                        <item.icon className="h-4 w-4 shrink-0" />
                        {!collapsed && <span className="flex-1">{item.title}</span>}
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                );
              })}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>

        {!collapsed && (
          <div className="mt-auto px-4 py-3 text-[10px] text-muted-foreground border-t border-sidebar-border flex items-center gap-2">
            <span className="relative flex h-1.5 w-1.5">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-status-present opacity-70" />
              <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-status-present" />
            </span>
            Connesso a Microsoft 365
          </div>
        )}
      </SidebarContent>
    </Sidebar>
  );
}
