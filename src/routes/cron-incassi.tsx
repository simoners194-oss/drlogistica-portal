// DR Portal — ricezione degli INCASSI aggregati dai ReportMovimenti.
// -----------------------------------------------------------------------------
// Lo script locale (scarica_aruba.py incassi) aggrega le rate per fattura e
// chiama questo indirizzo a blocchi:
//   /cron-incassi?token=<token del cron fatture>&dati=<base64 di [{numero,cliente,flusso,incassato,ultimaData}]>
// Il server applica SOLO AUMENTI: riduzioni contate e mai applicate,
// nessun azzeramento dal singolo blocco. Con `&forza=1` (solo a mano:
// `scarica_aruba.py incassi forza`, dopo una correzione fatta su Aruba) le
// riduzioni si applicano.
// Dal 27/10/2026 lo script manda anche la CORSA (`&corsa=<id>&blocco=<n>
// &blocchi=<tot>&anni=2025,2026`): sull'ultimo blocco il server cerca le
// fatture con incassi nel portale ma sparite dal report (incassi tolti su
// Aruba): il giro le conta, `forza` le azzera.
import { createFileRoute } from "@tanstack/react-router";
import { spCronIncassi } from "@/lib/sharepoint.functions";

const intero = (v: unknown): number => {
  const n = typeof v === "number" ? v : typeof v === "string" ? Number(v) : NaN;
  return Number.isInteger(n) && n > 0 ? n : 0;
};

export const Route = createFileRoute("/cron-incassi")({
  head: () => ({
    meta: [{ title: "Cron incassi — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
    dati: typeof search.dati === "string" ? search.dati : "",
    forza: search.forza === "1" || search.forza === 1 || search.forza === true,
    // Il router legge "20261027160030" come numero: la corsa va accettata anche così.
    corsa:
      typeof search.corsa === "string"
        ? search.corsa
        : typeof search.corsa === "number" && Number.isSafeInteger(search.corsa)
          ? String(search.corsa)
          : "",
    blocco: intero(search.blocco),
    blocchi: intero(search.blocchi),
    anni: typeof search.anni === "string" ? search.anni : "",
  }),
  loaderDeps: ({ search }) => ({
    token: search.token,
    dati: search.dati,
    forza: search.forza,
    corsa: search.corsa,
    blocco: search.blocco,
    blocchi: search.blocchi,
    anni: search.anni,
  }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token || !deps.dati) return { ok: false, messaggio: "token o dati mancanti" };
    try {
      const anni = deps.anni
        .split(",")
        .map((a) => Number(a.trim()))
        .filter((a) => Number.isInteger(a));
      const corsa =
        deps.corsa && deps.blocco && deps.blocchi && anni.length
          ? { id: deps.corsa, blocco: deps.blocco, blocchi: deps.blocchi, anni }
          : undefined;
      const r = (await spCronIncassi({
        data: { token: deps.token, dati: deps.dati, forza: deps.forza, corsa },
      })) as {
        aggiornate: number;
        invariate: number;
        riduzioniIgnorate: number;
        nonTrovate: number;
        daAzzerare: number;
        azzerate: number;
        azzerateElenco: string[];
        azzerateNota?: string;
      };
      let coda = "";
      if (corsa && corsa.blocco === corsa.blocchi) {
        coda = deps.forza
          ? `, ${r.azzerate} azzerate${r.azzerateElenco.length ? ` [${r.azzerateElenco.join(" | ").slice(0, 600)}]` : ""}`
          : `, ${r.daAzzerare} con incassi assenti dal report`;
        if (r.azzerateNota) coda += ` (${r.azzerateNota.slice(0, 300)})`;
      }
      return {
        ok: true,
        messaggio: `${deps.forza ? "FORZATO — " : ""}${r.aggiornate} aggiornate, ${r.invariate} invariate, ${r.riduzioniIgnorate} riduzioni ignorate, ${r.nonTrovate} non in archivio${coda}`,
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
