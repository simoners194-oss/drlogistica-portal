// DR Portal — Finanza → Storico estratti: A QUALE MOVIMENTO di banca è
// abbinata una distinta (Simone 20/11/2026: "ho bisogno di vedere a quale
// movimento sono collegati, devono essere cliccabili e riportarmi la lista
// dei movimenti abbinati").
//
// Una distinta è il gruppo delle disposizioni dello stesso giorno e dello
// stesso tipo nel report "Esiti pagamenti". Si abbina in due modi:
//  - A MANO o PER TRANCHE: le chiavi dei movimenti sono salvate sulle righe
//    della distinta (`movChiavi`). Vince su tutto, anche se la somma non torna;
//  - DA SOLA: nessuna chiave salvata, ma in banca c'è un'uscita con la stessa
//    somma (±1 €) entro 6 giorni. Se ce n'è più di una si prende LA PIÙ VICINA
//    per data (fino alla 1.125.0 era la prima trovata nell'elenco, che poteva
//    essere quella di sei giorni dopo anche con una dello stesso giorno).
//
// Logica pura: niente React, niente rete.

/** Quello che serve di un movimento di banca. */
export interface MovimentoPerDistinta {
  chiave: string;
  /** YYYY-MM-DD */
  dataContabile: string;
  /** Negativo = uscita. */
  importo: number;
}

/** Quello che serve di una distinta. */
export interface DistintaPerAggancio {
  /** Data di esecuzione, YYYY-MM-DD. */
  data: string;
  /** Somma delle disposizioni (positiva). */
  somma: number;
  /** Chiavi dei movimenti agganciati a mano o per tranche. */
  movChiavi: readonly string[];
}

/** Giorni entro cui un'uscita con la stessa somma si abbina da sola. */
export const GIORNI_AGGANCIO = 6;
/** Scarto ammesso sulla somma, in euro. */
export const SCARTO_AGGANCIO = 1;

const r2 = (n: number) => Math.round(n * 100) / 100;
const giorno = (iso: string) => Date.parse(`${iso.slice(0, 10)}T00:00:00Z`);
/** Giorni tra due date (sempre positivo). */
export const giorniTra = (a: string, b: string) => Math.abs((giorno(a) - giorno(b)) / 86400000);

/** Le uscite che per somma e data possono essere l'addebito della distinta,
 *  dalla più vicina per data (a pari data, quella con la somma più vicina). */
export function candidatiDistinta<M extends MovimentoPerDistinta>(
  g: Pick<DistintaPerAggancio, "data" | "somma">,
  movimenti: readonly M[],
): M[] {
  return movimenti
    .filter(
      (m) =>
        m.importo < 0 &&
        Math.abs(r2(-m.importo) - g.somma) <= SCARTO_AGGANCIO &&
        giorniTra(m.dataContabile, g.data) <= GIORNI_AGGANCIO,
    )
    .sort(
      (a, b) =>
        giorniTra(a.dataContabile, g.data) - giorniTra(b.dataContabile, g.data) ||
        Math.abs(-a.importo - g.somma) - Math.abs(-b.importo - g.somma) ||
        a.dataContabile.localeCompare(b.dataContabile) ||
        a.chiave.localeCompare(b.chiave),
    );
}

export interface AbbinatiDistinta<M> {
  /** "mano" = chiavi salvate (a mano o per tranche); "auto" = trovato da solo
   *  per somma e data; "nessuno" = non abbinata. */
  modo: "mano" | "auto" | "nessuno";
  /** I movimenti abbinati, dal più vecchio. */
  movimenti: M[];
  /** Solo con "auto": le altre uscite con la stessa somma nei 6 giorni, non
   *  abbinate. Servono a vedere quando l'abbinamento è una coincidenza. */
  altri: M[];
  /** Quanto coprono i movimenti abbinati (positivo). */
  coperto: number;
  /** La copertura torna con la somma della distinta (±1 €). */
  completo: boolean;
}

/** I movimenti di banca abbinati a una distinta, e come. */
export function movimentiAbbinati<M extends MovimentoPerDistinta>(
  g: DistintaPerAggancio,
  movimenti: readonly M[],
): AbbinatiDistinta<M> {
  const collegati = movimenti
    .filter((m) => g.movChiavi.includes(m.chiave))
    .sort(
      (a, b) => a.dataContabile.localeCompare(b.dataContabile) || a.chiave.localeCompare(b.chiave),
    );
  if (collegati.length) {
    const coperto = r2(collegati.reduce((s, m) => s - m.importo, 0));
    return {
      modo: "mano",
      movimenti: collegati,
      altri: [],
      coperto,
      completo: Math.abs(coperto - g.somma) <= SCARTO_AGGANCIO,
    };
  }
  const candidati = candidatiDistinta(g, movimenti);
  if (!candidati.length)
    return { modo: "nessuno", movimenti: [], altri: [], coperto: 0, completo: false };
  return {
    modo: "auto",
    movimenti: [candidati[0]],
    altri: candidati.slice(1),
    coperto: r2(-candidati[0].importo),
    completo: true,
  };
}
