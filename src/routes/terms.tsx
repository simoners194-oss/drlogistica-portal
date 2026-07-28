// DR Portal — Termini d'uso (pagina PUBBLICA, senza login).
// Requisito della registrazione Enable Banking; testo sintetico adeguato a
// uno strumento interno aziendale.
import { createFileRoute, Link } from "@tanstack/react-router";
import { APP_NAME } from "@/lib/modules";

export const Route = createFileRoute("/terms")({
  head: () => ({ meta: [{ title: "Termini d'uso — DR Portal" }] }),
  component: TermsPage,
});

function TermsPage() {
  return (
    <div className="min-h-screen bg-background text-foreground">
      <div className="mx-auto max-w-2xl px-6 py-10">
        <h1 className="text-2xl font-semibold tracking-tight">Termini d'uso</h1>
        <p className="mt-1 text-sm text-muted-foreground">
          {APP_NAME} — portale aziendale di DR Logistica S.r.l. · aggiornati al 28/07/2026
        </p>

        <div className="mt-6 space-y-5 text-[15px] leading-relaxed">
          <section>
            <h2 className="font-semibold">Natura del servizio</h2>
            <p>
              {APP_NAME} è uno strumento interno di DR Logistica S.r.l. (P.IVA 16935881009),
              riservato al personale e ai collaboratori autorizzati, per la gestione di presenze,
              richieste, documenti, comunicazioni e — per la sola direzione — dei dati
              amministrativo-finanziari della società.
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Accesso e credenziali</h2>
            <p>
              L'accesso avviene con credenziali personali (codice dipendente e PIN), che non devono
              essere condivise. Ogni operazione effettuata con le proprie credenziali è attribuita
              al loro titolare. Smarrimenti o sospetti di uso improprio vanno segnalati subito a{" "}
              <a className="text-primary underline" href="mailto:segreteria@drlogistica.it">
                segreteria@drlogistica.it
              </a>
              .
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Uso consentito</h2>
            <p>
              Il portale va utilizzato esclusivamente per finalità lavorative, secondo le procedure
              aziendali. Non è consentito accedere a dati non di propria competenza, estrarre dati
              per fini estranei al lavoro o interferire con il funzionamento del servizio.
            </p>
          </section>
          <section>
            <h2 className="font-semibold">Disponibilità e modifiche</h2>
            <p>
              Il servizio è fornito così com'è, senza garanzia di disponibilità continuativa;
              funzionalità e presenti termini possono essere aggiornati da DR Logistica in qualsiasi
              momento. Il trattamento dei dati personali è descritto nell'{" "}
              <Link to="/privacy" className="text-primary underline">
                informativa privacy
              </Link>
              .
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
