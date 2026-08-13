// DR Portal — ricezione dei COLLEGAMENTI NC estratti dal sito Aruba.
// -----------------------------------------------------------------------------
// Lo script locale (scarica_aruba.py nclinks) legge la colonna "Doc. coll."
// dal gestionale Aruba e chiama questo indirizzo a blocchi:
//   /cron-nc?token=<token del cron fatture>&dati=<base64 di [{file,numero,dir}]>
// Il server scrive RettificaNumero SOLO sulle NC ancora scollegate.
import { createFileRoute } from "@tanstack/react-router";
import { spCronNc } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/cron-nc")({
  head: () => ({
    meta: [{ title: "Collegamenti NC — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
    dati: typeof search.dati === "string" ? search.dati : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token, dati: search.dati }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token || !deps.dati) return { ok: false, messaggio: "token o dati mancanti" };
    try {
      const r = (await spCronNc({ data: { token: deps.token, dati: deps.dati } })) as {
        collegate: number;
        giaCollegate: number;
        nonTrovate: number;
      };
      return {
        ok: true,
        messaggio: `${r.collegate} collegate, ${r.giaCollegate} gia' collegate, ${r.nonTrovate} non in archivio`,
      };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronNcPage,
});

function CronNcPage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
