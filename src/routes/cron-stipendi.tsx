// DR Portal — innesco della lettura automatica dei file paghe.
// -----------------------------------------------------------------------------
// Stesso schema di cron-turni: un'attività programmata (il giro giornaliero
// sul PC di Simone, o un flusso Power Automate) chiama
//   https://portal.drlogistica.it/cron-stipendi?token=<token>
// e il portale legge Stipendi Dr.xlsx e i COSTI mensili da SharePoint
// aggiornando la tab Stipendi e la riga Stipendi dei Flussi. Vale il token
// "stipendi" oppure quello "fatture" già in uso sul PC. Risponde solo con
// un riepilogo: nessun dato esposto.
import { createFileRoute } from "@tanstack/react-router";
import { spCronStipendi } from "@/lib/stipendi.functions";

export const Route = createFileRoute("/cron-stipendi")({
  head: () => ({
    meta: [{ title: "Lettura stipendi — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = await spCronStipendi({ data: { token: deps.token } });
      return {
        ok: r.errori.length === 0,
        messaggio:
          `netti ${r.nettiMesi.length} mesi, costi ${r.costiMesi.length} mesi, flussi ${r.flussiAggiornati}, saltati ${r.saltati.length}, errori ${r.errori.length}` +
          (r.errori.length ? ` — ${r.errori.join(" | ")}` : "") +
          ` (${Math.round(r.durataMs / 1000)}s)`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronStipendiPage,
});

function CronStipendiPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
