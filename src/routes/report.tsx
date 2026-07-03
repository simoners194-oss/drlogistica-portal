import { createFileRoute } from "@tanstack/react-router";
import { AppShell } from "@/components/AppShell";
import { PlaceholderSection } from "@/components/PlaceholderSection";
import { BarChart3 } from "lucide-react";

export const Route = createFileRoute("/report")({
  head: () => ({ meta: [{ title: "Modulo Report — DR Portal" }] }),
  component: () => (
    <AppShell title="Report" subtitle="Statistiche presenze, ritardi e straordinari">
      <PlaceholderSection
        Icon={BarChart3}
        title="Report in arrivo"
        description="Reportistica mensile e annuale con esportazione in Excel e sincronizzazione automatica con SharePoint."
      />
    </AppShell>
  ),
});