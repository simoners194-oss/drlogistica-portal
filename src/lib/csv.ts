// DR Portal — export CSV condiviso (apribile in Excel).
// Separatore ";" (convenzione Excel italiano) + BOM UTF-8 per gli accenti.

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
