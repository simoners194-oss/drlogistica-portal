import { createFileRoute } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import { PlaceholderSection } from "@/components/PlaceholderSection";
import { FileText } from "lucide-react";

export const Route = createFileRoute("/richieste")({
  head: () => ({ meta: [{ title: "Modulo Richieste — DR Portal" }] }),
  component: () => (
    <AppShell title="Richieste" subtitle="Ferie, permessi e giustificativi">
      <PlaceholderSection
        Icon={FileText}
        title="Modulo Richieste in arrivo"
        description="Qui i dipendenti potranno inviare richieste di ferie, permessi e giustificativi. Il flusso di approvazione sarà collegato alla lista SharePoint 'Richieste' di Microsoft 365."
      />
    </AppShell>
  ),
});