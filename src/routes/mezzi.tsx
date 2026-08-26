// DR Portal — modulo Mezzi (gestione parco mezzi, Fase A).
// Scadenze e semafori, affidamenti, contratti, multe, ZTL, officina, km.

import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { Loader2 } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useLang } from "@/lib/i18n";
import { readSession, type SessionUser } from "@/lib/session";
import { haVistaDirezione } from "@/lib/richieste-logic";
import { spMezziGetDb } from "@/lib/mezzi.functions";
import type { MezziDb } from "@/lib/mezzi-types";
import { useMezzi } from "@/components/mezzi/shared";
import { ParcoTab } from "@/components/mezzi/ParcoTab";
import { ScadenzeTab } from "@/components/mezzi/ScadenzeTab";
import { AffidamentiTab } from "@/components/mezzi/AffidamentiTab";
import { ContrattiTab } from "@/components/mezzi/ContrattiTab";
import { MulteTab } from "@/components/mezzi/MulteTab";
import { ZtlTab } from "@/components/mezzi/ZtlTab";
import { OfficinaTab } from "@/components/mezzi/OfficinaTab";
import { KmCarburanteTab } from "@/components/mezzi/KmCarburanteTab";
import { ParametriTab } from "@/components/mezzi/ParametriTab";

export const Route = createFileRoute("/mezzi")({
  head: () => ({ meta: [{ title: "Mezzi — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    const s = readSession();
    if (!s) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: MezziPage,
});

function accessoMezzi(s: SessionUser | null): boolean {
  if (!s) return false;
  return (
    s.ruolo === "amministratore_sistema" ||
    s.ruolo === "responsabile" ||
    haVistaDirezione((s as { codice?: string }).codice ?? "")
  );
}

function MezziPage() {
  const { t } = useLang();
  const { m } = useMezzi();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [db, setDb] = useState<MezziDb | null>(null);
  const [errore, setErrore] = useState<string | null>(null);

  useEffect(() => {
    setSession(readSession());
  }, []);

  useEffect(() => {
    let vivo = true;
    spMezziGetDb()
      .then((res) => {
        if (vivo) setDb(res);
      })
      .catch((err) => {
        if (vivo) {
          const msg = err instanceof Error ? err.message : String(err);
          setErrore(msg);
          toast.error(msg);
        }
      });
    return () => {
      vivo = false;
    };
  }, []);

  const isAdmin = session?.ruolo === "amministratore_sistema";

  if (session && !accessoMezzi(session)) {
    return (
      <AppShell title={t("module.mezzi")}>
        <p className="text-sm text-muted-foreground">{t("common.restricted")}</p>
      </AppShell>
    );
  }

  return (
    <AppShell title={t("module.mezzi")} subtitle={t("mezzi.subtitle")} wide>
      {!db && !errore && (
        <div className="flex items-center gap-2 py-16 text-muted-foreground">
          <Loader2 className="h-5 w-5 animate-spin" /> {t("common.loading")}
        </div>
      )}
      {errore && !db && <p className="text-sm text-destructive">{errore}</p>}
      {db && (
        <Tabs defaultValue="parco">
          <TabsList className="mb-4 flex h-auto flex-wrap justify-start">
            <TabsTrigger value="parco">{m("tab.parco")}</TabsTrigger>
            <TabsTrigger value="scadenze">{m("tab.scadenze")}</TabsTrigger>
            <TabsTrigger value="affidamenti">{m("tab.affidamenti")}</TabsTrigger>
            <TabsTrigger value="contratti">{m("tab.contratti")}</TabsTrigger>
            <TabsTrigger value="multe">{m("tab.multe")}</TabsTrigger>
            <TabsTrigger value="ztl">{m("tab.ztl")}</TabsTrigger>
            <TabsTrigger value="officina">{m("tab.officina")}</TabsTrigger>
            <TabsTrigger value="km">{m("tab.km")}</TabsTrigger>
            <TabsTrigger value="parametri">{m("tab.parametri")}</TabsTrigger>
          </TabsList>
          <TabsContent value="parco">
            <ParcoTab db={db} onDb={setDb} isAdmin={isAdmin} />
          </TabsContent>
          <TabsContent value="scadenze">
            <ScadenzeTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="affidamenti">
            <AffidamentiTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="contratti">
            <ContrattiTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="multe">
            <MulteTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="ztl">
            <ZtlTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="officina">
            <OfficinaTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="km">
            <KmCarburanteTab db={db} onDb={setDb} />
          </TabsContent>
          <TabsContent value="parametri">
            <ParametriTab db={db} onDb={setDb} isAdmin={isAdmin} />
          </TabsContent>
        </Tabs>
      )}
    </AppShell>
  );
}
