import { Link, useRouterState } from "@tanstack/react-router";
import {
  LayoutDashboard,
  Clock,
  FileText,
  BarChart3,
  Settings,
} from "lucide-react";
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

const items = [
  { title: "Dashboard", url: "/dashboard", icon: LayoutDashboard, ready: true },
  { title: "Presenze", url: "/presenze", icon: Clock, ready: true },
  { title: "Richieste", url: "/richieste", icon: FileText, ready: false },
  { title: "Report", url: "/report", icon: BarChart3, ready: false },
  { title: "Amministrazione", url: "/amministrazione", icon: Settings, ready: false },
] as const;

export function AppSidebar() {
  const { state } = useSidebar();
  const collapsed = state === "collapsed";
  const pathname = useRouterState({ select: (s) => s.location.pathname });

  return (
    <Sidebar collapsible="icon">
      <SidebarHeader className="h-16 border-b border-sidebar-border flex items-center justify-center px-2">
        {collapsed ? (
          <img src="/favicon.png" alt="DR" className="h-8 w-8" />
        ) : (
          <Logo size={40} />
        )}
      </SidebarHeader>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Menu</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {items.map((item) => {
                const active = pathname === item.url;
                return (
                  <SidebarMenuItem key={item.title}>
                    <SidebarMenuButton asChild isActive={active} tooltip={item.title}>
                      <Link
                        to={item.url}
                        className="flex items-center gap-2"
                        aria-disabled={!item.ready}
                        onClick={(e) => {
                          if (!item.ready) return; // le pagine placeholder mostrano comunque un messaggio
                        }}
                      >
                        <item.icon className="h-4 w-4 shrink-0" />
                        {!collapsed && (
                          <span className="flex-1 flex items-center justify-between">
                            {item.title}
                            {!item.ready && (
                              <span className="text-[9px] uppercase tracking-wider text-muted-foreground bg-muted px-1.5 py-0.5 rounded">
                                Presto
                              </span>
                            )}
                          </span>
                        )}
                      </Link>
                    </SidebarMenuButton>
                  </SidebarMenuItem>
                );
              })}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>

        {!collapsed && (
          <div className="mt-auto px-4 py-3 text-[10px] text-muted-foreground border-t border-sidebar-border">
            Dati sincronizzati da Microsoft 365 · SharePoint
          </div>
        )}
      </SidebarContent>
    </Sidebar>
  );
}