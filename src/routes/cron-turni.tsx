// DR Portal — innesco del promemoria "manca la timbratura di uscita".
// -----------------------------------------------------------------------------
// Stesso schema del sync bancario: un flusso Power Automate chiama questo
// indirizzo a intervalli regolari (es. ogni ora dalle 15 alle 22)
//   https://portal.drlogistica.it/cron-turni?token=<token>
// e il portale manda la notifica push a chi ha il turno ancora aperto oltre le
// ore previste. Il token si legge in Gestione timbrature. Nessun dato esposto:
// la pagina risponde solo con il conteggio degli avvisi inviati.
import { createFileRoute } from "@tanstack/react-router";
import { spCronTurni } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/cron-turni")({
  head: () => ({
    meta: [{ title: "Promemoria turni — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = await spCronTurni({ data: { token: deps.token } });
      return {
        ok: true,
        messaggio: `${r.avvisati} promemoria inviati (su ${r.controllati} in servizio oggi)`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronTurniPage,
});

function CronTurniPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
