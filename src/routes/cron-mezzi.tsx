// DR Portal — innesco degli alert scadenze del modulo Mezzi.
// -----------------------------------------------------------------------------
// Stesso schema dei cron esistenti: un flusso Power Automate chiama una volta
// al giorno
//   https://portal.drlogistica.it/cron-mezzi?token=<token>
// (il token è lo stesso dei cron fatture/incassi). Il portale controlla le
// scadenze aperte e accoda le email 90/60/30 giorni (+ scadute) sulla Coda
// Email. Nessun dato esposto: la pagina risponde solo con i conteggi.
import { createFileRoute } from "@tanstack/react-router";
import { spMezziCron } from "@/lib/mezzi.functions";

export const Route = createFileRoute("/cron-mezzi")({
  head: () => ({
    meta: [{ title: "Alert mezzi — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = await spMezziCron({ data: { token: deps.token } });
      return {
        ok: true,
        messaggio: `${r.notifiche} notifiche su ${r.controllate} scadenze aperte${r.emailInviata ? " (email accodata)" : ""}`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronMezziPage,
});

function CronMezziPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
