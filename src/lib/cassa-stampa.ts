// DR Portal — Finanza → Flussi → Cassa proprietà: STAMPA ed EXPORT delle
// singole sezioni del foglio (Simone 11/11/2026: "ho bisogno di
// stampare/esportare le varie voci di cassa proprietà"):
//   1) uscite imminenti   2) uscite della seconda colonna (1-5 del mese dopo)
//   3) entrati            4) entranti (dentro anche le "da sollecitare")
//   5) standby entrate    6) standby uscite    7) standby 2 uscite
//
// Logica pura (le sezioni, le righe del CSV, l'HTML della stampa) più una
// sola funzione che tocca il browser: apriStampa, che apre il documento in
// una finestra sua e lancia la stampa (da lì si salva anche in PDF).
import type { PrioritaCassa, SecchioCassa } from "./cassa-logic";
import { bloccoEntrata, sommaCelle, type BloccoEntrate, type VoceFiltrabile } from "./cassa-filtri";

export type SezioneCassaId =
  | "uscite_ora"
  | "uscite_prossimi"
  | "entrati"
  | "entranti"
  | "standby_entrate"
  | "standby_uscite"
  | "standby2_uscite";

/** Una riga del foglio, con quello che serve per stamparla. */
export interface VoceStampa extends VoceFiltrabile {
  /** Scadenza / data attesa / data dell'incasso (YYYY-MM-DD). */
  data: string;
}

export interface SezioneCassa<T extends VoceStampa = VoceStampa> {
  id: SezioneCassaId;
  entrata: boolean;
  voci: T[];
  totale: number;
}

/** Le sette sezioni, nell'ordine chiesto, ognuna ordinata per data (poi per
 *  importo, dal più grande) e col suo totale. */
export function sezioniCassa<T extends VoceStampa>(
  entrate: readonly T[],
  uscite: readonly T[],
): SezioneCassa<T>[] {
  const ordina = (l: T[]) =>
    [...l].sort(
      (a, b) =>
        a.data.localeCompare(b.data) || b.importo - a.importo || a.nome.localeCompare(b.nome, "it"),
    );
  const u = (s: SecchioCassa) => ordina(uscite.filter((x) => x.secchio === s));
  const e = (blocchi: BloccoEntrate[]) =>
    ordina(entrate.filter((x) => blocchi.includes(bloccoEntrata(x))));
  const sez = (id: SezioneCassaId, entrata: boolean, voci: T[]): SezioneCassa<T> => ({
    id,
    entrata,
    voci,
    totale: sommaCelle(voci),
  });
  return [
    sez("uscite_ora", false, u("ora")),
    sez("uscite_prossimi", false, u("prossimi")),
    sez("entrati", true, e(["gia"])),
    sez("entranti", true, e(["entranti", "sollecitare"])),
    sez("standby_entrate", true, e(["standby"])),
    sez("standby_uscite", false, u("standby")),
    sez("standby2_uscite", false, u("standby2")),
  ];
}

/** Testi della stampa e dell'export (dalla lingua della vista). */
export interface EtichetteStampa {
  /** "Cassa proprietà" */
  titolo: string;
  /** "situazione al 11/11/2026 · stampato il …" (+ eventuale avviso filtri) */
  sottotitolo: string;
  sezioni: Record<SezioneCassaId, string>;
  colSezione: string;
  colVoce: string;
  colDettaglio: string;
  colData: string;
  colImporto: string;
  colTipologia: string;
  colCliente: string;
  colPriorita: string;
  colSollecitare: string;
  colNota: string;
  totale: string;
  voci: string;
  si: string;
  nessuna: string;
  priorita: Record<PrioritaCassa, string>;
  /** Indirizzo del logo (facoltativo). */
  logo?: string;
}

const eur = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dataIt = (iso: string) =>
  /^\d{4}-\d{2}-\d{2}/.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : "";
const csvNum = (n: number) => (Math.round(n * 100) / 100).toFixed(2).replace(".", ",");
const daSollecitare = (v: VoceStampa) => v.attesa && !v.parcheggiata && v.daSollecitare;

/** Testata e righe del CSV: una riga per voce, poi il totale di ogni sezione. */
export function csvSezioni(
  sezioni: readonly SezioneCassa[],
  e: EtichetteStampa,
): { testata: string[]; righe: (string | number)[][] } {
  const testata = [
    e.colSezione,
    e.colVoce,
    e.colDettaglio,
    e.colData,
    e.colImporto,
    e.colTipologia,
    e.colCliente,
    e.colPriorita,
    e.colSollecitare,
    e.colNota,
  ];
  const righe: (string | number)[][] = [];
  for (const s of sezioni) {
    const nome = e.sezioni[s.id];
    for (const v of s.voci)
      righe.push([
        nome,
        v.nome,
        v.dettaglio,
        /^\d{4}-\d{2}-\d{2}/.test(v.data) ? v.data.slice(0, 10) : "",
        csvNum(v.importo),
        v.tipologia?.trim() ?? "",
        v.cliente?.trim() ?? "",
        v.priorita ? e.priorita[v.priorita] : "",
        daSollecitare(v) ? e.si : "",
        v.nota,
      ]);
    righe.push([
      nome,
      `${e.totale} (${s.voci.length} ${e.voci})`,
      "",
      "",
      csvNum(s.totale),
      "",
      "",
      "",
      "",
      "",
    ]);
  }
  return { testata, righe };
}

const esc = (v: unknown) =>
  String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

/** Colore della testata di ogni sezione, come le colonne del foglio. */
const COLORE: Record<SezioneCassaId, string> = {
  uscite_ora: "#e11d48",
  uscite_prossimi: "#d97706",
  entrati: "#047857",
  entranti: "#059669",
  standby_entrate: "#64748b",
  standby_uscite: "#64748b",
  standby2_uscite: "#334155",
};

/** Il documento da stampare: intestazione, poi una tabella per sezione con
 *  voce, dettaglio, data, importo e nota, e il totale in fondo. */
export function htmlStampaSezioni(sezioni: readonly SezioneCassa[], e: EtichetteStampa): string {
  const corpo = sezioni
    .map((s) => {
      const righe = s.voci.length
        ? s.voci
            .map((v) => {
              const nota = [v.priorita ? `[${e.priorita[v.priorita]}]` : "", v.nota]
                .filter(Boolean)
                .join(" ");
              return `<tr><td class="v">${esc(v.nome)}${
                daSollecitare(v) ? ` <span class="tag">${esc(e.colSollecitare)}</span>` : ""
              }</td><td class="d">${esc(v.dettaglio)}</td><td class="c">${esc(dataIt(v.data))}</td><td class="n${
                v.importo < 0 ? " neg" : ""
              }">${esc(eur(v.importo))}</td><td class="d">${esc(nota)}</td></tr>`;
            })
            .join("")
        : `<tr><td colspan="5" class="d">${esc(e.nessuna)}</td></tr>`;
      return `<section><h2 style="background:${COLORE[s.id]}"><span>${esc(e.sezioni[s.id])}</span><span>${
        s.voci.length
      } ${esc(e.voci)} · ${esc(eur(s.totale))} €</span></h2><table><thead><tr><th>${esc(
        e.colVoce,
      )}</th><th>${esc(e.colDettaglio)}</th><th class="c">${esc(e.colData)}</th><th class="n">${esc(
        e.colImporto,
      )}</th><th>${esc(e.colNota)}</th></tr></thead><tbody>${righe}</tbody><tfoot><tr><td colspan="3">${esc(
        e.totale,
      )} ${esc(e.sezioni[s.id])}</td><td class="n">${esc(eur(s.totale))}</td><td></td></tr></tfoot></table></section>`;
    })
    .join("");
  const titoloPagina =
    sezioni.length === 1 ? `${e.titolo} · ${e.sezioni[sezioni[0].id]}` : e.titolo;
  return `<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(
    titoloPagina,
  )}</title><style>
@page{margin:12mm}
*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
body{font-family:"Century Gothic","Segoe UI",Arial,sans-serif;color:#111;font-size:11px;margin:0;padding:16px}
header{display:flex;align-items:center;gap:12px;border-bottom:2px solid #111;padding-bottom:8px;margin-bottom:4px}
header img{height:34px;width:auto}
h1{font-size:17px;margin:0}
.sub{color:#555;font-size:10.5px;margin-top:2px}
section{margin-top:14px}
h2{font-size:12.5px;margin:0;padding:5px 8px;color:#fff;display:flex;justify-content:space-between;gap:12px;break-after:avoid}
table{width:100%;border-collapse:collapse}
th{text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.04em;color:#555;border-bottom:1.5px solid #111;padding:4px 6px}
td{padding:3.5px 6px;border-bottom:1px solid #ddd;vertical-align:top}
tr{break-inside:avoid}
.v{font-weight:600;width:26%}
.d{color:#555;font-size:10px}
.c{white-space:nowrap;text-align:center;width:9%}
.n{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;width:12%}
.neg{color:#be123c}
.tag{font-size:8.5px;font-weight:700;text-transform:uppercase;color:#b45309;border:1px solid #b45309;border-radius:3px;padding:0 3px}
tfoot td{font-weight:700;font-size:12px;border-top:1.5px solid #111;border-bottom:none}
@media screen{body{max-width:1000px;margin:0 auto}}
</style></head><body><header>${
    e.logo ? `<img src="${esc(e.logo)}" alt="">` : ""
  }<div><h1>${esc(titoloPagina)}</h1><div class="sub">${esc(e.sottotitolo)}</div></div></header>${corpo}</body></html>`;
}

/** Apre il documento in una finestra sua e lancia la stampa (dal riquadro di
 *  stampa si salva anche in PDF). Falso se il browser blocca la finestra. */
export function apriStampa(html: string): boolean {
  const w = window.open("", "_blank");
  if (!w) return false;
  w.document.open();
  w.document.write(html);
  w.document.close();
  let fatto = false;
  const stampa = () => {
    if (fatto) return;
    fatto = true;
    w.focus();
    w.print();
  };
  // Si aspetta il logo (se c'è), ma non più di un secondo e mezzo.
  const img = w.document.images[0];
  if (img && !img.complete) {
    img.addEventListener("load", stampa);
    img.addEventListener("error", stampa);
    window.setTimeout(stampa, 1500);
  } else {
    window.setTimeout(stampa, 100);
  }
  return true;
}
