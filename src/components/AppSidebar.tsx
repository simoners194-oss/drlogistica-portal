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
// Versione RITAGLIATA del logo: il PNG originale è 1920×1920 con il marchio
// che occupa solo il 21% (enormi margini trasparenti) — con object-contain
// renderizzava minuscolo. Questo file è il solo contenuto reale (1802×492).
import logoTrim from "@/assets/dr-logistica-logo.png";
import { MODULES } from "@/lib/modules";
import { canAccess, readSession, type Ruolo, type SessionSede } from "@/lib/session";
import { sedeTimbra, anySedeTimbra } from "@/lib/mock-data";
import { isSedeStorica, isSupervisoreGlobale } from "@/lib/richieste-logic";
import { useLang } from "@/lib/i18n";

export function AppSidebar() {
  const { state } = useSidebar();
  const { t, tModule } = useLang();
  const collapsed = state === "collapsed";
  const pathname = useRouterState({ select: (s) => s.location.pathname });
  // Il ruolo viene letto client-side (sessionStorage) dopo il mount per
  // evitare mismatch di hydration con SSR/prerender.
  const [ruolo, setRuolo] = useState<Ruolo | null>(null);
  const [operatore, setOperatore] = useState(false);
  const [autorizza, setAutorizza] = useState(false);
  const [sede, setSede] = useState<SessionSede | null>(null);
  const [codice, setCodice] = useState<string>("");
  useEffect(() => {
    const s = readSession();
    setRuolo(s?.ruolo ?? null);
    setOperatore(s?.operatore ?? false);
    setAutorizza(s?.autorizza ?? false);
    setSede(s?.sede ?? null);
    setCodice(s?.codice ?? "");
  }, [pathname]);

  // Finché il ruolo non è noto, mostra solo le voci pubbliche a tutti i
  // ruoli (Presenze, Richieste) per evitare "flash" del menu completo. Le voci
  // con `requiresOperatore` compaiono solo per gli operatori (DR000).
  // Presenze attive per l'utente? Sede "tutte" (admin/resp.) → basta che UNA
  // sede timbri; sede reale → dipende dalla sede. I moduli che richiedono la
  // timbratura restano visibili ma GRIGI quando non attiva.
  const presenzeAttive =
    sede == null ? true : sede === "tutte" ? anySedeTimbra() : sedeTimbra(sede);

  const visibleModules = MODULES.filter((m) => {
    // Requisiti di capability obbligatori (AND col ruolo). L'amministratore
    // di sistema è esente: vede anche i moduli riservati all'operatore.
    const admin = ruolo === "amministratore_sistema";
    if (m.requiresOperatore && !operatore && !admin) return false;
    if (m.requiresAutorizza && !autorizza && !admin) return false;
    // Moduli riservati alle sedi storiche (es. Procurement): visibili anche
    // ad autorizzatori (DR005 approva) e a chi ha sede "tutte" (admin).
    if (
      m.soloSediStoriche &&
      sede != null &&
      sede !== "tutte" &&
      !isSedeStorica(sede) &&
      !autorizza
    )
      return false;
    const roleOk = ruolo ? canAccess(m, ruolo) : canAccess(m, "dipendente");
    // Moduli riservati al direttore (DR005): visibili anche all'admin (ruolo).
    if (m.soloDirettore) return roleOk || isSupervisoreGlobale(codice);
    // Capability alternative: visibile se ruolo ammesso OPPURE capability.
    if (m.orCapabilities && m.orCapabilities.length) {
      const capOk = m.orCapabilities.some(
        (c) => (c === "operatore" && operatore) || (c === "autorizza" && autorizza),
      );
      return roleOk || capOk;
    }
    return roleOk;
  });

  return (
    <Sidebar collapsible="icon">
      <SidebarHeader className="h-24 border-b border-sidebar-border overflow-hidden p-0">
        {collapsed ? (
          <div className="h-full w-full flex items-center justify-center">
            <img src="/favicon.png" alt="DR" className="h-10 w-10" />
          </div>
        ) : (
          // Il logo copre TUTTO il riquadro: header alto abbastanza da far
          // vincere il vincolo di larghezza (logo ~3.7:1 → tutta la sidebar).
          <img
            src={logoTrim}
            alt="DR Logistica"
            className="h-full w-full object-contain px-2 py-1.5"
          />
        )}
      </SidebarHeader>
      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>{t("nav.modules")}</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {visibleModules.map((item) => {
                const active = pathname === item.url;
                // Voci disabilitate: "In arrivo" (non pronte) oppure timbratura
                // non attiva per la sede. Non cliccabili, aspetto attenuato.
                const disabledTimbratura = item.richiedeTimbratura && !presenzeAttive;
                const titolo = tModule(item.id, item.title);
                if (!item.ready || disabledTimbratura) {
                  const badge = !item.ready ? t("nav.comingSoon") : t("nav.notActive");
                  const tip = `${titolo} — ${badge}`;
                  return (
                    <SidebarMenuItem key={item.title}>
                      <SidebarMenuButton
                        tooltip={tip}
                        aria-disabled="true"
                        className="opacity-60 cursor-not-allowed hover:bg-transparent"
                        onClick={(e) => e.preventDefault()}
                      >
                        <item.icon className="h-4 w-4 shrink-0" />
                        {!collapsed && (
                          <span className="flex-1 flex items-center justify-between">
                            {titolo}
                            <span className="text-[9px] font-medium uppercase tracking-wider text-muted-foreground bg-muted px-1.5 py-0.5 rounded-full">
                              {badge}
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
                      tooltip={titolo}
                      className={
                        active
                          ? "bg-primary/10 text-primary font-medium border-l-2 border-primary rounded-l-none pl-[calc(0.5rem-2px)]"
                          : ""
                      }
                    >
                      <Link to={item.url} className="flex items-center gap-2">
                        <item.icon className="h-4 w-4 shrink-0" />
                        {!collapsed && <span className="flex-1">{titolo}</span>}
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
