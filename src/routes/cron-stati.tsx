// DR Portal — ricezione degli STATI DI PAGAMENTO letti dalla griglia Aruba.
// -----------------------------------------------------------------------------
// Lo script locale (scarica_aruba.py nclinks) legge la colonna "Pagamenti"
// (Pagata / Non pagata / Stornata / Non gestita) dal gestionale Aruba e
// chiama questo indirizzo a blocchi:
//   /cron-stati?token=<token del cron fatture>&dati=<base64 di [{file,stato,dir,dataPag?}]>
// Il server applica con le guardie: mai retrocessioni automatiche, i
// parziali non si scrivono (gli importi veri li porta il report incassi).
import { createFileRoute } from "@tanstack/react-router";
import { spCronStati } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/cron-stati")({
  head: () => ({
    meta: [{ title: "Stati pagamento — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
    dati: typeof search.dati === "string" ? search.dati : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token, dati: search.dati }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token || !deps.dati) return { ok: false, messaggio: "token o dati mancanti" };
    try {
      const r = (await spCronStati({ data: { token: deps.token, dati: deps.dati } })) as {
        aggiornate: number;
        invariate: number;
        retrocessioniIgnorate: number;
        nonTrovate: number;
      };
      return {
        ok: true,
        messaggio: `${r.aggiornate} aggiornate, ${r.invariate} invariate, ${r.retrocessioniIgnorate} retrocessioni ignorate, ${r.nonTrovate} non in archivio`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronStatiPage,
});

function CronStatiPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
