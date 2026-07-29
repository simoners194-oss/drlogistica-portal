// DR Portal — innesco della sincronizzazione bancaria PROGRAMMATA.
// -----------------------------------------------------------------------------
// Il Worker non ha uno scheduler proprio: un flusso Power Automate chiama
// questo indirizzo agli orari stabiliti (es. 07:00 e 18:00)
//   https://portal.drlogistica.it/cron-banca?token=<token>
// Il token è derivato dal segreto server e si legge in Amministrazione →
// Collegamento banca. Nessuna sessione, nessun dato esposto: la pagina
// risponde solo con l'esito dell'operazione.
//
// NB: le esecuzioni programmate NON sono "presidiate" (nessun utente davanti),
// quindi consumano il budget di accessi PSD2 concesso dalla banca.
import { createFileRoute } from "@tanstack/react-router";
import { spEbCron } from "@/lib/sharepoint.functions";

interface EsitoCron {
  esito: string;
  scritti: number;
  doppioni: number;
  pendenti: number;
}

export const Route = createFileRoute("/cron-banca")({
  head: () => ({
    meta: [{ title: "Sync banca — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  // Il loader gira lato server sulla richiesta HTTP del flusso.
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = (await spEbCron({ data: { token: deps.token } })) as EsitoCron;
      return {
        ok: true,
        messaggio: `${r.esito} — ${r.scritti} nuovi, ${r.doppioni} già presenti, ${r.pendenti} in attesa`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronBancaPage,
});

function CronBancaPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
