// DR Portal — export CSV condiviso (apribile in Excel).
// Separatore ";" (convenzione Excel italiano) + BOM UTF-8 per gli accenti.

/** Data per il CSV: ISO `aaaa-mm-gg`, l'unico formato che Excel riconosce come
 *  DATA in qualunque lingua. Scritta come "gg/mm/aaaa" finisce in colonna testo
 *  appena il computer non è in italiano, e su una colonna di testo la pivot non
 *  offre il raggruppamento per mese/trimestre. Vuoto = cella vuota, mai "—":
 *  un trattino basta a far degradare l'intera colonna a testo. */
export function csvData(iso?: string | null): string {
  const s = (iso ?? "").slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : "";
}

/** Anno, trimestre e mese di una data ISO, come colonne già pronte: così il
 *  raggruppamento della pivot non dipende da come Excel ha letto la data. */
export function csvPeriodo(iso?: string | null): [string, string, string] {
  const s = csvData(iso);
  if (!s) return ["", "", ""];
  const [anno, mese] = s.split("-");
  return [anno, `${anno}-T${Math.ceil(Number(mese) / 3)}`, `${anno}-${mese}`];
}

export function esportaCsvFile(
  nomeFile: string,
  header: string[],
  righe: (string | number | null | undefined)[][],
): void {
  const esc = (v: string | number | null | undefined): string => {
    const s = String(v ?? "");
    return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const body = righe.map((r) => r.map(esc).join(";"));
  const csv = [header.join(";"), ...body].join("\r\n");
  const blob = new Blob([String.fromCharCode(0xfeff) + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeFile.endsWith(".csv") ? nomeFile : `${nomeFile}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}
