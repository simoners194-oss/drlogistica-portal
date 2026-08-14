// Vista PIVOT stile Excel richiesta dalla direzione: righe raggruppate per
// classificazione, colonne = mesi, valori = somma importi. Due sezioni:
// SINTESI (allocazione primaria × tipologia, con subtotali) chiusa di
// default — tenerla chiusa velocizza l'apertura — e DETTAGLIO (tutti i
// livelli, con colonne attivabili) aperto. Subtotali sul primo livello in
// ENTRAMBE, comprimibili con un click (e comprimi/espandi tutto). Export
// CSV di ogni sezione, subtotali compresi. Usata da Movimenti e Fatture
// passive (dove il mese è quello di competenza).
import { useMemo, useState } from "react";
import { useLang } from "../lib/i18n";
import { esportaCsvFile } from "../lib/csv";

export interface RigaPivot {
  allocPrimaria: string;
  allocSecondaria: string;
  tipologia: string;
  sottocategoria: string;
  cliente: string;
  /** "YYYY-MM" (mese contabile o mese di competenza). */
  mese: string;
  importo: number;
}

const TUTTI_CAMPI = [
  "allocPrimaria",
  "allocSecondaria",
  "tipologia",
  "sottocategoria",
  "cliente",
] as const;
type Campo = (typeof TUTTI_CAMPI)[number];

function fmtN(x: number): string {
  return x.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

interface Gruppo {
  vals: string[];
  perMese: Map<string, number>;
  tot: number;
}

/** Aggregazione pura: la usano sia la tabella sia l'export CSV. */
function aggrega(righe: RigaPivot[], livelli: readonly Campo[]) {
  const mesiSet = new Set<string>();
  const map = new Map<string, Gruppo>();
  const totMese = new Map<string, number>();
  let totale = 0;
  for (const r of righe) {
    const m = r.mese || "—";
    mesiSet.add(m);
    const vals = livelli.map((l) => r[l] || "");
    const k = vals.join("");
    const g = map.get(k) ?? { vals, perMese: new Map<string, number>(), tot: 0 };
    g.perMese.set(m, (g.perMese.get(m) ?? 0) + r.importo);
    g.tot += r.importo;
    map.set(k, g);
    totMese.set(m, (totMese.get(m) ?? 0) + r.importo);
    totale += r.importo;
  }
  const gruppi = [...map.values()].sort((a, b) => {
    for (let i = 0; i < a.vals.length; i++) {
      const c = a.vals[i].localeCompare(b.vals[i]);
      if (c) return c;
    }
    return 0;
  });
  // Subtotali per PRIMO livello, nell'ordine dei gruppi.
  const sub = new Map<string, Gruppo>();
  for (const g of gruppi) {
    const k = g.vals[0];
    const s = sub.get(k) ?? { vals: [k], perMese: new Map<string, number>(), tot: 0 };
    for (const [m, v] of g.perMese) s.perMese.set(m, (s.perMese.get(m) ?? 0) + v);
    s.tot += g.tot;
    sub.set(k, s);
  }
  return { mesi: [...mesiSet].sort(), gruppi, sub, totMese, totale };
}

function Tabella({
  righe,
  livelli,
  etichette,
  vuotaLabel,
  totaleLabel,
  comprimiLabel,
  espandiLabel,
}: {
  righe: RigaPivot[];
  livelli: readonly Campo[];
  etichette: Record<Campo, string>;
  vuotaLabel: string;
  totaleLabel: string;
  comprimiLabel: string;
  espandiLabel: string;
}) {
  const { mesi, gruppi, sub, totMese, totale } = useMemo(
    () => aggrega(righe, livelli),
    [righe, livelli],
  );
  // Gruppi del primo livello COMPRESSI: click sul subtotale (o comprimi
  // tutto) e restano solo le righe di totale — lettura veloce del direttore.
  const [chiusi, setChiusi] = useState<Set<string>>(new Set());
  const conSub = livelli.length > 1;

  if (!righe.length) return null;
  const cellaNum = (v: number | undefined, extra = "") => (
    <td className={`whitespace-nowrap py-1 pl-3 text-right tabular-nums ${extra}`}>
      {v == null || Math.abs(v) < 0.005 ? "" : fmtN(v)}
    </td>
  );
  const corpo: React.ReactNode[] = [];
  let gruppoCorr: string | null = null;
  const emettiSub = (nome: string) => {
    const s = sub.get(nome);
    if (!s) return;
    const chiuso = chiusi.has(nome);
    corpo.push(
      <tr
        key={`sub-${nome}`}
        onClick={() =>
          setChiusi((prev) => {
            const next = new Set(prev);
            if (next.has(nome)) next.delete(nome);
            else next.add(nome);
            return next;
          })
        }
        className="cursor-pointer border-t border-border bg-muted/50 font-semibold hover:bg-muted"
      >
        <td colSpan={livelli.length} className="py-1 pr-2">
          {chiuso ? "▸" : "▾"} {nome || vuotaLabel} — {totaleLabel}
        </td>
        {mesi.map((m) => cellaNum(s.perMese.get(m)))}
        {cellaNum(s.tot)}
      </tr>,
    );
  };
  for (const g of gruppi) {
    if (conSub && g.vals[0] !== gruppoCorr) {
      if (gruppoCorr != null) emettiSub(gruppoCorr);
      gruppoCorr = g.vals[0];
    }
    if (conSub && chiusi.has(g.vals[0])) continue;
    corpo.push(
      <tr key={g.vals.join("|")} className="border-t border-border/40">
        {g.vals.map((v, i) => (
          <td
            key={livelli[i]}
            className="max-w-52 truncate py-1 pr-2 text-foreground"
            title={v || vuotaLabel}
          >
            {v || <span className="text-muted-foreground">{vuotaLabel}</span>}
          </td>
        ))}
        {mesi.map((m) => cellaNum(g.perMese.get(m)))}
        {cellaNum(g.tot, "font-medium")}
      </tr>,
    );
  }
  if (conSub && gruppoCorr != null) emettiSub(gruppoCorr);
  return (
    <div>
      {conSub && (
        <div className="mb-1 flex gap-3 text-[11px]">
          <button
            type="button"
            onClick={() => setChiusi(new Set(sub.keys()))}
            className="text-muted-foreground underline underline-offset-2 hover:text-foreground"
          >
            {comprimiLabel}
          </button>
          <button
            type="button"
            onClick={() => setChiusi(new Set())}
            className="text-muted-foreground underline underline-offset-2 hover:text-foreground"
          >
            {espandiLabel}
          </button>
        </div>
      )}
      <div className="max-h-[70vh] overflow-auto">
        <table className="w-full text-[12px] leading-tight">
          <thead className="sticky top-0 z-10 bg-card">
            <tr className="border-b border-border text-left text-[11px] text-muted-foreground">
              {livelli.map((l) => (
                <th key={l} className="py-1.5 pr-2">
                  {etichette[l]}
                </th>
              ))}
              {mesi.map((m) => (
                <th key={m} className="whitespace-nowrap py-1.5 pl-3 text-right">
                  {m}
                </th>
              ))}
              <th className="py-1.5 pl-3 text-right">{totaleLabel}</th>
            </tr>
          </thead>
          <tbody>
            {corpo}
            <tr className="border-t-2 border-border bg-muted/60 font-semibold">
              <td colSpan={livelli.length} className="py-1.5 pr-2">
                {totaleLabel}
              </td>
              {mesi.map((m) => cellaNum(totMese.get(m), "font-semibold"))}
              {cellaNum(totale, "font-semibold")}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

export function PivotClassificazione({
  righe,
  nome = "pivot",
}: {
  righe: RigaPivot[];
  nome?: string;
}) {
  const { t } = useLang();
  const [sintesiAperta, setSintesiAperta] = useState(false);
  // Il DETTAGLIO parte aperto (richiesta direzione: i dati subito in
  // vista); la SINTESI parte chiusa per velocizzare l'apertura.
  const [dettaglioAperto, setDettaglioAperto] = useState(true);
  const [attivi, setAttivi] = useState<Campo[]>([...TUTTI_CAMPI]);
  const etichette: Record<Campo, string> = {
    allocPrimaria: t("fin.allocPri"),
    allocSecondaria: t("fin.allocSec"),
    tipologia: t("ft.colTipologia"),
    sottocategoria: t("fin.sottocat"),
    cliente: t("fin.cliente"),
  };
  const livelliAttivi = TUTTI_CAMPI.filter((c) => attivi.includes(c));

  // Export CSV di una sezione: stesse righe della tabella, subtotali del
  // primo livello e Totale compresi; numeri con la virgola per Excel.
  const esporta = (livelli: readonly Campo[], suffisso: string) => {
    const { mesi, gruppi, sub, totale, totMese } = aggrega(righe, livelli);
    const num = (v: number | undefined) =>
      v == null || Math.abs(v) < 0.005 ? "" : v.toFixed(2).replace(".", ",");
    const out: (string | number)[][] = [];
    let corr: string | null = null;
    const vuota = t("ft.classVuota");
    const rigaSub = (nomeG: string) => {
      const s = sub.get(nomeG);
      if (!s) return;
      out.push([
        `${nomeG || vuota} — ${t("fin.pivotTotale")}`,
        ...Array(livelli.length - 1).fill(""),
        ...mesi.map((m) => num(s.perMese.get(m))),
        num(s.tot),
      ]);
    };
    for (const g of gruppi) {
      if (livelli.length > 1 && g.vals[0] !== corr) {
        if (corr != null) rigaSub(corr);
        corr = g.vals[0];
      }
      out.push([
        ...g.vals.map((v) => v || vuota),
        ...mesi.map((m) => num(g.perMese.get(m))),
        num(g.tot),
      ]);
    }
    if (livelli.length > 1 && corr != null) rigaSub(corr);
    out.push([
      t("fin.pivotTotale"),
      ...Array(livelli.length - 1).fill(""),
      ...mesi.map((m) => num(totMese.get(m))),
      num(totale),
    ]);
    esportaCsvFile(
      `pivot-${nome}-${suffisso}`,
      [...livelli.map((l) => etichette[l]), ...mesi, t("fin.pivotTotale")],
      out,
    );
  };

  const bottoneCsv = (livelli: readonly Campo[], suffisso: string) => (
    <button
      type="button"
      onClick={() => esporta(livelli, suffisso)}
      className="ml-2 rounded border border-border px-2 py-0.5 text-[11px] text-muted-foreground hover:bg-muted"
    >
      CSV
    </button>
  );

  return (
    <div className="mb-4 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      <div className="flex items-center">
        <button
          type="button"
          onClick={() => setSintesiAperta((x) => !x)}
          className="text-sm font-semibold text-foreground"
        >
          {sintesiAperta ? "▾" : "▸"} {t("fin.pivotSintesi")}
        </button>
        {bottoneCsv(["allocPrimaria", "tipologia"], "sintesi")}
      </div>
      {sintesiAperta && (
        <div className="mt-2">
          <Tabella
            righe={righe}
            livelli={["allocPrimaria", "tipologia"]}
            etichette={etichette}
            vuotaLabel={t("ft.classVuota")}
            totaleLabel={t("fin.pivotTotale")}
            comprimiLabel={t("fin.pivotComprimi")}
            espandiLabel={t("fin.pivotEspandi")}
          />
        </div>
      )}
      <div className="mt-2 border-t border-border/50 pt-2">
        <div className="flex items-center">
          <button
            type="button"
            onClick={() => setDettaglioAperto((x) => !x)}
            className="text-sm font-semibold text-foreground"
          >
            {dettaglioAperto ? "▾" : "▸"} {t("fin.pivotDettaglio")}
          </button>
          {bottoneCsv(livelliAttivi, "dettaglio")}
        </div>
        {dettaglioAperto && (
          <>
            <div className="mt-2 flex flex-wrap items-center gap-2 text-[11px]">
              <span className="text-muted-foreground">{t("fin.pivotColonne")}</span>
              {TUTTI_CAMPI.map((c) => (
                <button
                  key={c}
                  type="button"
                  onClick={() =>
                    setAttivi((prev) => {
                      const senza = prev.filter((x) => x !== c);
                      if (senza.length === prev.length) return [...prev, c];
                      return senza.length ? senza : prev; // almeno un livello
                    })
                  }
                  className={`rounded-full border px-2.5 py-0.5 font-medium ${
                    attivi.includes(c)
                      ? "border-primary bg-primary/10 text-primary"
                      : "border-border text-muted-foreground"
                  }`}
                >
                  {etichette[c]}
                </button>
              ))}
            </div>
            <div className="mt-2">
              <Tabella
                righe={righe}
                livelli={livelliAttivi}
                etichette={etichette}
                vuotaLabel={t("ft.classVuota")}
                totaleLabel={t("fin.pivotTotale")}
                comprimiLabel={t("fin.pivotComprimi")}
                espandiLabel={t("fin.pivotEspandi")}
              />
            </div>
          </>
        )}
      </div>
    </div>
  );
}
