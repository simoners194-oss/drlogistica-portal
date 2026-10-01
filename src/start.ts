import { createStart, createMiddleware } from "@tanstack/react-start";

import { renderErrorPage } from "./lib/error-page";
import { MSG_VERSIONE_VECCHIA, ricaricaSeAggiornato, rispostaStantia } from "./lib/versione-client";

const errorMiddleware = createMiddleware().server(async ({ next }) => {
  try {
    return await next();
  } catch (error) {
    if (error != null && typeof error === "object" && "statusCode" in error) {
      throw error;
    }
    console.error(error);
    return new Response(renderErrorPage(), {
      status: 500,
      headers: { "content-type": "text/html; charset=utf-8" },
    });
  }
});

// SCHEDA VECCHIA DOPO UNA PUBLISH (Zingali 01/10/2026, "non riescono a
// timbrare"): l'hosting risponde a un indirizzo server non più esistente con
// 200 e un file JavaScript, non con un errore; il client di TanStack consegna
// allora la Response grezza come risultato e OGNI salvataggio del portale
// (timbrature, correzioni del preposto, schede finanza…) sembra riuscito
// senza esserlo. Questo middleware, sul client, trasforma quel finto successo
// in un errore riconoscibile e fa ripartire la sentinella che ricarica la
// pagina (vedi lib/versione-client).
const versioneMiddleware = createMiddleware({ type: "function" }).client(async ({ next }) => {
  const res = await next();
  if (rispostaStantia((res as unknown as { result?: unknown }).result)) {
    void ricaricaSeAggiornato(undefined, true);
    throw new Error(MSG_VERSIONE_VECCHIA);
  }
  return res;
});

export const startInstance = createStart(() => ({
  requestMiddleware: [errorMiddleware],
  functionMiddleware: [versioneMiddleware],
}));
