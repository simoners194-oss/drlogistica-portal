// DR Portal — Finanza → Storico estratti: i MOVIMENTI DI BANCA abbinati a una
// distinta (Simone 20/11/2026: "ho bisogno di vedere a quale movimento sono
// collegati, devono essere cliccabili e riportarmi la lista dei movimenti
// abbinati"). Sta in cima alla finestra che si apre dal badge della distinta.
import { AlertTriangle, Landmark } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { giorniTra, type AbbinatiDistinta } from "@/lib/distinte-aggancio";

/** Quello che si mostra di un movimento di banca. */
export interface MovimentoAbbinato {
  chiave: string;
  dataContabile: string;
  importo: number;
  descrizione: string;
  cliente: string;
  conto: string;
}

const soldi = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dataIt = (iso: string) =>
  /^\d{4}-\d{2}-\d{2}/.test(iso)
    ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`
    : "—";

export function DistintaMovimenti({
  abbinati,
  data,
  somma,
}: {
  abbinati: AbbinatiDistinta<MovimentoAbbinato>;
  /** Data di esecuzione della distinta (YYYY-MM-DD). */
  data: string;
  /** Somma della distinta. */
  somma: number;
}) {
  const { t } = useLang();
  const { modo, movimenti, altri, coperto, completo } = abbinati;
  // Quanti giorni passano tra la distinta e l'addebito: sei giorni di distanza
  // su una somma uguale sono spesso un altro pagamento.
  const distanza = (m: MovimentoAbbinato) => {
    const gg = Math.round(giorniTra(m.dataContabile, data));
    if (gg === 0) return t("fin.distMovStessoGiorno");
    return `${gg} ${t(m.dataContabile > data ? "fin.distMovGgDopo" : "fin.distMovGgPrima")}`;
  };
  const riga = (m: MovimentoAbbinato, spenta = false) => (
    <tr key={m.chiave} className="border-t border-border/50 align-top">
      <td className="py-1.5 pr-3 whitespace-nowrap">
        <div className={`tabular-nums ${spenta ? "text-muted-foreground" : "font-medium"}`}>
          {dataIt(m.dataContabile)}
        </div>
        <div className="text-[10.5px] text-muted-foreground">{distanza(m)}</div>
      </td>
      <td
        className={`py-1.5 pr-3 text-right tabular-nums whitespace-nowrap ${spenta ? "text-muted-foreground" : "font-semibold text-foreground"}`}
      >
        {soldi(m.importo)} €
      </td>
      <td className="py-1.5 pr-2">
        {m.cliente && (
          <div className={spenta ? "text-muted-foreground" : "font-medium text-foreground"}>
            {m.cliente}
          </div>
        )}
        <div className="text-[11px] leading-snug break-words text-muted-foreground">
          {m.descrizione || "—"}
        </div>
      </td>
      <td className="py-1.5 text-right text-[11px] whitespace-nowrap text-muted-foreground">
        {m.conto}
      </td>
    </tr>
  );
  const pillola =
    modo === "nessuno"
      ? "bg-status-absent/10 text-status-absent"
      : modo === "auto"
        ? "bg-status-present/10 text-status-present"
        : completo
          ? "bg-status-present/10 text-status-present"
          : "bg-primary/10 text-primary";
  return (
    <section className="mb-4 overflow-hidden rounded-xl border border-border">
      <header className="flex flex-wrap items-center gap-2 bg-muted/50 px-3 py-2">
        <Landmark className="h-4 w-4 text-primary" />
        <span className="text-[11px] font-semibold tracking-wider text-foreground uppercase">
          {t("fin.distMovTitolo")}
        </span>
        <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${pillola}`}>
          {t(
            modo === "nessuno"
              ? "fin.distMovNessuno"
              : modo === "auto"
                ? "fin.distMovAuto"
                : "fin.distMovMano",
          )}
        </span>
        {movimenti.length > 0 && (
          <span className="ml-auto text-right text-[12px] text-muted-foreground">
            {movimenti.length} mov · {t("fin.distMovCoperto")}{" "}
            <b className="text-[13.5px] text-foreground tabular-nums">{soldi(coperto)} €</b>
            {!completo && (
              <>
                {" "}
                {t("fin.distMovDi")} <span className="tabular-nums">{soldi(somma)} €</span>
              </>
            )}
          </span>
        )}
      </header>
      {movimenti.length > 0 && (
        <div className="px-3 pb-1">
          <table className="w-full text-[12.5px]">
            <tbody>{movimenti.map((m) => riga(m))}</tbody>
          </table>
        </div>
      )}
      {modo === "auto" && (
        <p className="border-t border-border/50 px-3 py-1.5 text-[11px] text-muted-foreground">
          {t("fin.distMovAutoNota")}
        </p>
      )}
      {altri.length > 0 && (
        <div className="border-t border-amber-400/50 bg-amber-50 px-3 py-2 dark:bg-amber-950/30">
          <p className="mb-0.5 flex items-center gap-1.5 text-[11px] font-medium text-amber-800 dark:text-amber-300">
            <AlertTriangle className="h-3.5 w-3.5" /> {t("fin.distMovAltri")} ({altri.length})
          </p>
          <table className="w-full text-[12.5px]">
            <tbody>{altri.map((m) => riga(m, true))}</tbody>
          </table>
        </div>
      )}
    </section>
  );
}
