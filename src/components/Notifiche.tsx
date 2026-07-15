import { useEffect } from "react";
import { useNavigate } from "@tanstack/react-router";
import { toast } from "sonner";
import { readSession } from "@/lib/session";
import { APP_INFO } from "@/lib/version";
import { spGetComunicazioni, spGetMiePreseVisione } from "@/lib/sharepoint.functions";
import type { SpComunicazione } from "@/lib/sharepoint.server";

// Mostra i popup di notifica UNA volta per caricamento pagina (non ad ogni
// navigazione interna). Renderizza null: è solo un effetto.
let notificheMostrate = false;

export function Notifiche() {
  const navigate = useNavigate();

  useEffect(() => {
    if (notificheMostrate) return;
    const s = readSession();
    if (!s) return;
    notificheMostrate = true;

    // 1) Novità: se la versione app è cambiata dall'ultima visita a /novita.
    try {
      const last = window.localStorage.getItem("dr:novita:lastVersion");
      if (last !== APP_INFO.version) {
        toast(`Novità in DR Portal (v${APP_INFO.version})`, {
          description: "Scopri cosa è cambiato in quest'aggiornamento.",
          action: { label: "Vedi", onClick: () => navigate({ to: "/novita" }) },
          duration: 8000,
        });
      }
    } catch {
      /* localStorage non disponibile */
    }

    // 2) Comunicazioni con presa visione non ancora confermata.
    Promise.all([spGetComunicazioni(), spGetMiePreseVisione()])
      .then(([comsRaw, visteRaw]) => {
        const coms = comsRaw as SpComunicazione[];
        const viste = new Set(visteRaw as string[]);
        const daLeggere = coms.filter((c) => c.richiedePresaVisione && !viste.has(c.id));
        if (daLeggere.length > 0) {
          const n = daLeggere.length;
          toast(`Hai ${n} comunicazion${n === 1 ? "e" : "i"} da leggere`, {
            description: daLeggere[0].titolo,
            action: { label: "Apri", onClick: () => navigate({ to: "/comunicazioni" }) },
            duration: 10000,
          });
        }
      })
      .catch(() => {
        /* silenzioso: le notifiche non devono disturbare se il modulo è assente */
      });
  }, [navigate]);

  return null;
}
