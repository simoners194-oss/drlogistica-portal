import { createFileRoute } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import { PlaceholderSection } from "@/components/PlaceholderSection";
import { Settings } from "lucide-react";

export const Route = createFileRoute("/amministrazione")({
  head: () => ({ meta: [{ title: "Amministrazione — DR Portal" }] }),
  component: () => (
    <AppShell title="Amministrazione" subtitle="Configurazione sedi, ruoli e permessi">
      <PlaceholderSection
        Icon={Settings}
        title="Pannello amministrativo in arrivo"
        description="Gestione anagrafica dipendenti, sedi, turni e integrazione con Azure AD / SharePoint."
      />
    </AppShell>
  ),
});