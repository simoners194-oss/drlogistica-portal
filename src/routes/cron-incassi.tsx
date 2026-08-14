// DR Portal — ricezione degli INCASSI aggregati dai ReportMovimenti.
// -----------------------------------------------------------------------------
// Lo script locale (scarica_aruba.py incassi) aggrega le rate per fattura e
// chiama questo indirizzo a blocchi:
//   /cron-incassi?token=<token del cron fatture>&dati=<base64 di [{numero,cliente,flusso,incassato,ultimaData}]>
// Il server applica SOLO AUMENTI: riduzioni contate e mai applicate,
// nessun azzeramento. Il resto resta all'import manuale con le conferme.
import { createFileRoute } from "@tanstack/react-router";
import { spCronIncassi } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/cron-incassi")({
  head: () => ({
    meta: [{ title: "Cron incassi — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
    dati: typeof search.dati === "string" ? search.dati : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token, dati: search.dati }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token || !deps.dati) return { ok: false, messaggio: "token o dati mancanti" };
    try {
      const r = (await spCronIncassi({ data: { token: deps.token, dati: deps.dati } })) as {
        aggiornate: number;
        invariate: number;
        riduzioniIgnorate: number;
        nonTrovate: number;
      };
      return {
        ok: true,
        messaggio: `${r.aggiornate} aggiornate, ${r.invariate} invariate, ${r.riduzioniIgnorate} riduzioni ignorate, ${r.nonTrovate} non in archivio`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronIncassiPage,
});

function CronIncassiPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
