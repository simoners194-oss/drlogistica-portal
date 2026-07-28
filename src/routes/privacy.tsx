// DR Portal — Informativa privacy (pagina PUBBLICA, senza login).
// Nata come requisito per la registrazione dell'app Enable Banking (PSD2),
// ma doverosa in generale: il portale tratta dati di dipendenti e dati
// bancari aziendali. Testo sintetico; per i dettagli si rimanda al Titolare.
import { createFileRoute, Link } from "@tanstack/react-router";
import { APP_NAME } from "@/lib/modules";

export const Route = createFileRoute("/privacy")({
  head: () => ({ meta: [{ title: "Privacy — DR Portal" }] }),
  component: PrivacyPage,
});

function PrivacyPage() {
  return (
    <div className="min-h-screen bg-background text-foreground">
      <div className="mx-auto max-w-2xl px-6 py-10">
        <h1 className="text-2xl font-semibold tracking-tight">Informativa privacy</h1>
        <p className="mt-1 text-sm text-muted-foreground">
          {APP_NAME} — portale aziendale di DR Logistica S.r.l. · aggiornata al 28/07/2026
        </p>

        <div className="mt-6 space-y-5 text-[15px] leading-relaxed">
          <section>
            <h2 className="font-semibold">Titolare del trattamento</h2>
            <p>
              DR Logistica S.r.l. (P.IVA 16935881009). Per ogni richiesta relativa ai dati
              personali:{" "}
              <a className="text-primary underline" href="mailto:segreteria@drlogistica.it">
                segreteria@drlogistica.it
              </a>
              .
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Quali dati e perché</h2>
            <p>
              Il portale è riservato al personale autorizzato di DR Logistica e tratta: dati
              identificativi e organizzativi dei dipendenti (anagrafica, sede, ruolo), presenze e
              timbrature, richieste (ferie, permessi, malattia, rimborsi), documenti di lavoro (es.
              buste paga, contratti) e comunicazioni interne. Finalità: gestione del rapporto di
              lavoro e dell'organizzazione aziendale.
            </p>
            <p className="mt-2">
              La sezione finanziaria, accessibile alla sola direzione, tratta i movimenti del conto
              corrente aziendale e le fatture elettroniche della società — anche, previo consenso
              esplicito della direzione presso la banca, tramite servizi di accesso ai conti
              (PSD2/open banking) in <b>sola lettura</b>. Non vengono trattati conti personali di
              dipendenti o terzi.
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Dove sono i dati</h2>
            <p>
              I dati risiedono nel tenant Microsoft 365 aziendale (SharePoint) e sono raggiunti
              tramite connessioni cifrate. L'accesso è protetto da autenticazione personale; le
              operazioni sensibili sono verificate lato server.
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Conservazione e diritti</h2>
            <p>
              I dati sono conservati per il tempo necessario alle finalità indicate e agli obblighi
              di legge. Gli interessati possono esercitare i diritti previsti dagli artt. 15-22 del
              GDPR (accesso, rettifica, cancellazione, limitazione, opposizione) scrivendo al
              Titolare all'indirizzo sopra indicato.
            </p>
          </section>
        </div>

        <div className="mt-8 text-sm">
          <Link to="/" className="text-primary underline">
            ← Torna al portale
          </Link>
        </div>
      </div>
    </div>
  );
}
