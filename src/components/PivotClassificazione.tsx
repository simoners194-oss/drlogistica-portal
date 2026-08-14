// Vista PIVOT stile Excel richiesta dalla direzione: righe raggruppate per
// classificazione, colonne = mesi, valori = somma importi. Due sezioni:
// SINTESI (allocazione primaria × tipologia, con subtotali) chiusa di
// default — tenerla chiusa velocizza l'apertura — e DETTAGLIO (tutti i
// livelli, con colonne attivabili/disattivabili) aperto. Usata sia dai
// Movimenti sia dalle Fatture passive (dove il mese è quello di competenza).
import { useMemo, useState } from "react";
import { useLang } from "../lib/i18n";

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

function Tabella({
  righe,
  livelli,
  etichette,
  sottototali,
  vuotaLabel,
  totaleLabel,
}: {
  righe: RigaPivot[];
  livelli: readonly Campo[];
  etichette: Record<Campo, string>;
  /** Righe di subtotale per il PRIMO livello (come la pivot del direttore). */
  sottototali: boolean;
  vuotaLabel: string;
  totaleLabel: string;
}) {
  const { mesi, gruppi, totMese, totale } = useMemo(() => {
    const mesiSet = new Set<string>();
    const map = new Map<string, { vals: string[]; perMese: Map<string, number>; tot: number }>();
    const totMese2 = new Map<string, number>();
    let totale2 = 0;
    for (const r of righe) {
      const m = r.mese || "—";
      mesiSet.add(m);
      const vals = livelli.map((l) => r[l] || "");
      const k = vals.join("");
      const g = map.get(k) ?? { vals, perMese: new Map<string, number>(), tot: 0 };
      g.perMese.set(m, (g.perMese.get(m) ?? 0) + r.importo);
      g.tot += r.importo;
      map.set(k, g);
      totMese2.set(m, (totMese2.get(m) ?? 0) + r.importo);
      totale2 += r.importo;
    }
    const ordinati = [...map.values()].sort((a, b) => {
      for (let i = 0; i < a.vals.length; i++) {
        const c = a.vals[i].localeCompare(b.vals[i]);
        if (c) return c;
      }
      return 0;
    });
    return {
      mesi: [...mesiSet].sort(),
      gruppi: ordinati,
      totMese: totMese2,
      totale: totale2,
    };
  }, [righe, livelli]);

  if (!righe.length) return null;
  const cellaNum = (v: number | undefined, extra = "") => (
    <td className={`whitespace-nowrap py-1 pl-3 text-right tabular-nums ${extra}`}>
      {v == null || Math.abs(v) < 0.005 ? "" : fmtN(v)}
    </td>
  );
  // Righe con eventuali subtotali intercalati sul primo livello.
  const corpo: React.ReactNode[] = [];
  let gruppoCorr = "";
  let sub: { perMese: Map<string, number>; tot: number } | null = null;
  const chiudiSub = () => {
    if (!sottototali || !sub) return;
    corpo.push(
      <tr key={`sub-${gruppoCorr}`} className="border-t border-border bg-muted/50 font-semibold">
        <td colSpan={livelli.length} className="py-1 pr-2">
          {gruppoCorr || vuotaLabel} — {totaleLabel}
        </td>
        {mesi.map((m) => cellaNum(sub!.perMese.get(m)))}
        {cellaNum(sub!.tot)}
      </tr>,
    );
    sub = null;
  };
  for (const g of gruppi) {
    if (sottototali && g.vals[0] !== gruppoCorr && sub) chiudiSub();
    if (g.vals[0] !== gruppoCorr) gruppoCorr = g.vals[0];
    if (sottototali) {
      sub = sub ?? { perMese: new Map(), tot: 0 };
      for (const [m, v] of g.perMese) sub.perMese.set(m, (sub.perMese.get(m) ?? 0) + v);
      sub.tot += g.tot;
    }
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
  chiudiSub();
  return (
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
  );
}

export function PivotClassificazione({ righe }: { righe: RigaPivot[] }) {
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
  return (
    <div className="mb-4 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      {/* SINTESI: chiusa di default (apertura pagina piu' veloce). */}
      <button
        type="button"
        onClick={() => setSintesiAperta((x) => !x)}
        className="text-sm font-semibold text-foreground"
      >
        {sintesiAperta ? "▾" : "▸"} {t("fin.pivotSintesi")}
      </button>
      {sintesiAperta && (
        <div className="mt-2">
          <Tabella
            righe={righe}
            livelli={["allocPrimaria", "tipologia"]}
            etichette={etichette}
            sottototali
            vuotaLabel={t("ft.classVuota")}
            totaleLabel={t("fin.pivotTotale")}
          />
        </div>
      )}
      <div className="mt-2 border-t border-border/50 pt-2">
        <button
          type="button"
          onClick={() => setDettaglioAperto((x) => !x)}
          className="text-sm font-semibold text-foreground"
        >
          {dettaglioAperto ? "▾" : "▸"} {t("fin.pivotDettaglio")}
        </button>
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
                sottototali={false}
                vuotaLabel={t("ft.classVuota")}
                totaleLabel={t("fin.pivotTotale")}
              />
            </div>
          </>
        )}
      </div>
    </div>
  );
}
