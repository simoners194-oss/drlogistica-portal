// DR Portal — indirizzo FISSO che dichiara la versione pubblicata.
// La sentinella client (src/lib/versione-client.ts) lo interroga per capire
// se la scheda è rimasta indietro dopo una publish e deve ricaricarsi.
// Deve restare leggerissimo: nessun dato, nessuna chiamata.
// ATTENZIONE CACHE: oggi l'HTML del worker non passa dalla cache edge di
// Cloudflare e il client interroga con ?ts= + no-store. Se un giorno si
// attivasse una Cache Rule "cache everything" su portal.drlogistica.it,
// ESCLUDERE /versione (una risposta in cache renderebbe cieca la sentinella).
import { createFileRoute } from "@tanstack/react-router";
import { APP_INFO } from "@/lib/version";

export const Route = createFileRoute("/versione")({
  head: () => ({
    meta: [{ title: "Versione — DR Portal" }, { name: "robots", content: "noindex" }],
  }),
  component: VersionePage,
});

function VersionePage() {
  return (
    <pre className="p-6 font-mono text-sm" data-versione={APP_INFO.version}>
      {APP_INFO.name} v{APP_INFO.version} · build {APP_INFO.build}
    </pre>
  );
}
