// DR Portal — informazioni di versione centralizzate.
// Ogni pagina, footer, pannello diagnostico e centro notifiche
// deve leggere queste costanti — non ripetere mai a mano
// versione, build o data di rilascio nel codice o nella UI.

export const APP_INFO = {
  name: "DR Portal",
  tagline: "Il portale aziendale di DR Logistica",
  vendor: "DR Logistica",
  version: "1.1.0",
  build: "2026.07.13",
  releaseDate: "2026-07-13",
  copyright: `© ${new Date().getFullYear()} DR Logistica`,
} as const;

export function formatReleaseDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString("it-IT", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
  } catch {
    return iso;
  }
}
