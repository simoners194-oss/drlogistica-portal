// DR Portal — innesco del SYNC FATTURE Aruba programmato.
// -----------------------------------------------------------------------------
// GitHub Actions (o qualunque scheduler) chiama a orari fissi
//   https://portal.drlogistica.it/cron-fatture?token=<token>
// Il token è derivato dal segreto server e si legge in Amministrazione →
// Diagnostica Aruba. Nessuna sessione, nessun dato esposto: la pagina
// risponde solo con l'esito. Guardia anti-doppione: sotto i 20 minuti
// dall'ultimo sync completo non si riparte.
import { createFileRoute } from "@tanstack/react-router";
import { spArubaCron } from "@/lib/sharepoint.functions";
import type { ArubaSyncResult } from "@/lib/aruba.server";

export const Route = createFileRoute("/cron-fatture")({
  head: () => ({
    meta: [{ title: "Sync fatture — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  validateSearch: (search: Record<string, unknown>) => ({
    token: typeof search.token === "string" ? search.token : "",
  }),
  loaderDeps: ({ search }) => ({ token: search.token }),
  loader: async ({ deps }): Promise<{ ok: boolean; messaggio: string }> => {
    if (!deps.token) return { ok: false, messaggio: "token mancante" };
    try {
      const r = (await spArubaCron({ data: { token: deps.token } })) as
        ArubaSyncResult | { saltato: true; messaggio: string };
      if ("saltato" in r) return { ok: true, messaggio: r.messaggio };
      const righe = r.esiti.map(
        (e) =>
          `${e.direzione}: ${e.lotti} lotti, ${e.daScaricare} nuovi, ${e.importate} importate, ${e.aggiornate} aggiornate${e.errori.length ? `, errori: ${e.errori.length}` : ""}`,
      );
      return { ok: true, messaggio: `${r.finestraDa} → ${r.finestraA} — ${righe.join(" | ")}` };
    } catch (err) {
      return { ok: false, messaggio: err instanceof Error ? err.message : String(err) };
    }
  },
  component: CronFatturePage,
});

function CronFatturePage() {
  const { ok, messaggio } = Route.useLoaderData();
  return (
    <pre className="p-6 font-mono text-sm">
      {ok ? "OK" : "ERRORE"}: {messaggio}
    </pre>
  );
}
