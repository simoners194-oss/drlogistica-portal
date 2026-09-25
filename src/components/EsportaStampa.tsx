// DR Portal — "Esporta CSV" + "Stampa" per i riepiloghi delle ore (Simone
// 25/09: "in ogni pagina con riepiloghi delle ore metti la possibilità di
// stampare o esportare"). Un solo componente, stessi due tasti ovunque: il
// CSV apre in Excel (separatore ";", BOM), la stampa usa il foglio di stile
// di stampa globale (styles.css) che nasconde menù, filtri e bottoni.
import { Download, Printer } from "lucide-react";
import { esportaCsvFile } from "@/lib/csv";
import { useLang } from "@/lib/i18n";

export function EsportaStampa({
  nomeFile,
  testata,
  righe,
  disabled,
  className,
}: {
  /** Nome del file CSV (senza estensione). */
  nomeFile: string;
  testata: string[];
  /** Le righe si calcolano al click: la pagina non ricostruisce nulla a ogni render. */
  righe: () => (string | number | null | undefined)[][];
  disabled?: boolean;
  className?: string;
}) {
  const { t } = useLang();
  const btn =
    "inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-1.5 text-sm text-foreground hover:bg-secondary transition-colors disabled:opacity-50 disabled:cursor-not-allowed";
  return (
    <div className={`no-print flex flex-wrap items-center gap-2 ${className ?? ""}`}>
      <button
        type="button"
        disabled={disabled}
        onClick={() => esportaCsvFile(nomeFile, testata, righe())}
        className={btn}
      >
        <Download className="h-4 w-4" /> {t("common.exportCsv")}
      </button>
      <button
        type="button"
        disabled={disabled}
        onClick={() => window.print()}
        className={btn}
        title={t("common.printTip")}
      >
        <Printer className="h-4 w-4" /> {t("common.print")}
      </button>
    </div>
  );
}
