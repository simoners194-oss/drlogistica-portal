// DR Portal — POLL del PC aziendale per il GIRO SU RICHIESTA.
// -----------------------------------------------------------------------------
// Il bottone "Sincronizza da Aruba" scrive una richiesta di giro completo;
// il PC aziendale chiama questo indirizzo ogni ~10 minuti:
//   /cron-giro?token=<token del cron fatture>
// Risposta "OK: ESEGUI" = c'era una richiesta fresca (viene consumata) e lo
// script locale deve partire; "OK: NIENTE" = nulla da fare.
import { createFileRoute } from "@tanstack/react-router";
import { spCronGiroPoll } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/cron-giro")({
  head: () => ({
    meta: [{ title: "Giro su richiesta — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = (await spCronGiroPoll({ data: { token: deps.token } })) as { esegui: boolean };
      return { ok: true, messaggio: r.esegui ? "ESEGUI" : "NIENTE" };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronGiroPage,
});

function CronGiroPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
