import { Link, useRouterState } from "@tanstack/react-router";
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
          <SidebarGroupLabel>Moduli</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {MODULES.map((item) => {
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
                              <span className="text-[9px] font-medium uppercase tracking-wider text-primary/80 bg-primary/10 px-1.5 py-0.5 rounded-full">
                                In arrivo
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