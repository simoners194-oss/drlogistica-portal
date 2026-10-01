// DR Portal — Rendiconto → scheda "Riepilogo presenze" (Simone 01/10/2026,
// v1.88.0): il foglio "PRESENZE <appalto> SUDDIVISO" di Monica fatto dal
// portale. Un giorno per colonna, i dipendenti raggruppati per sede con il
// numero progressivo; per ognuno la riga Tot (ore del giorno) e, se ne ha,
// la riga Not (ore nella fascia notturna 22–06); i codici FE/ML/PC/NC al
// posto delle ore quando non si è lavorato, SAB/DOM nei weekend vuoti;
// la colonna TOT e, per ogni sede, la riga "TOTALE ORE GIORNALIERE". Usa i
// filtri del Rendiconto (periodo, sede, dipendente) ed esce in CSV/stampa
// con la stessa griglia. Arrotondamento ai 15 minuti acceso di default.
import { Fragment, useMemo, useState } from "react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import { arrotondaQuarto, isoDow } from "@/lib/rendiconto-logic";
import type { PresenzeGiorno, PresenzeMatriceRiga } from "@/lib/sharepoint.server";

function eachDay(from: string, to: string): string[] {
  const out: string[] = [];
  const d = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`).getTime();
  while (d.getTime() <= end) {
    out.push(
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`,
    );
    d.setDate(d.getDate() + 1);
  }
  return out;
}
const ddmm = (g: string) => `${g.slice(8, 10)}/${g.slice(5, 7)}`;
const fmtOre = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
const csvNum = (n: number) => (n > 0 ? String(Math.round(n * 100) / 100).replace(".", ",") : "0");

export function RiepilogoPresenze({
  righe,
  loading,
  from,
  to,
  sedeF,
  dipF,
  nomeFile,
}: {
  righe: PresenzeMatriceRiga[] | null;
  loading: boolean;
  from: string;
  to: string;
  sedeF: string;
  dipF: string;
  nomeFile: string;
}) {
  const { t } = useLang();
  const [arrotonda, setArrotonda] = useState(true);
  const giorni = useMemo(() => eachDay(from, to), [from, to]);
  const r = (h: number) => (arrotonda ? arrotondaQuarto(h) : h);

  // Gruppi per sede, dipendenti in ordine alfabetico, filtri del Rendiconto.
  const gruppi = useMemo(() => {
    const per = new Map<string, PresenzeMatriceRiga[]>();
    for (const x of righe ?? []) {
      if (sedeF !== "tutte" && x.sede !== sedeF) continue;
      if (dipF && x.dipendenteId !== dipF) continue;
      const k = x.sede || "—";
      const l = per.get(k) ?? [];
      l.push(x);
      per.set(k, l);
    }
    return [...per.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], "it"))
      .map(([sede, dips]) => ({
        sede,
        dips: dips.sort((a, b) => a.nomeCompleto.localeCompare(b.nomeCompleto, "it")),
      }));
  }, [righe, sedeF, dipF]);

  // Cella: ore arrotondate, oppure il codice quando non si è lavorato.
  const cella = (g: string, c: PresenzeGiorno | undefined): { testo: string; tipo: string } => {
    if (c?.codice === "NC") return { testo: "NC", tipo: "nc" };
    const ore = r(c?.ore ?? 0);
    if (ore > 0) return { testo: fmtOre(ore), tipo: c?.permessoOre ? "orePerm" : "ore" };
    if (c?.codice) return { testo: c.codice, tipo: "codice" };
    const dow = isoDow(g);
    if (dow === 6) return { testo: "SAB", tipo: "we" };
    if (dow === 7) return { testo: "DOM", tipo: "we" };
    return { testo: "—", tipo: "vuoto" };
  };
  const totDip = (x: PresenzeMatriceRiga) =>
    giorni.reduce((s, g) => s + (x.giorni[g]?.codice === "NC" ? 0 : r(x.giorni[g]?.ore ?? 0)), 0);
  const totNotte = (x: PresenzeMatriceRiga) =>
    giorni.reduce((s, g) => s + (x.giorni[g]?.codice === "NC" ? 0 : r(x.giorni[g]?.notte ?? 0)), 0);
  const totGiornoSede = (dips: PresenzeMatriceRiga[], g: string) =>
    dips.reduce((s, x) => s + (x.giorni[g]?.codice === "NC" ? 0 : r(x.giorni[g]?.ore ?? 0)), 0);
  const tuttiDip = gruppi.flatMap((x) => x.dips);

  const testataCsv = [
    "N.",
    t("common.employee"),
    t("common.site"),
    t("rep.presRiga"),
    ...giorni.map(ddmm),
    "TOT",
  ];
  const righeCsv = () => {
    const out: (string | number)[][] = [];
    for (const grp of gruppi) {
      grp.dips.forEach((x, i) => {
        out.push([
          i + 1,
          x.nomeCompleto,
          grp.sede,
          "Tot",
          ...giorni.map((g) => {
            const c = cella(g, x.giorni[g]);
            return c.tipo === "ore" || c.tipo === "orePerm"
              ? csvNum(r(x.giorni[g]?.ore ?? 0))
              : c.testo;
          }),
          csvNum(totDip(x)),
        ]);
        if (totNotte(x) > 0)
          out.push([
            "",
            x.nomeCompleto,
            grp.sede,
            "Not",
            ...giorni.map((g) => {
              const c = cella(g, x.giorni[g]);
              return c.tipo === "ore" || c.tipo === "orePerm"
                ? csvNum(r(x.giorni[g]?.notte ?? 0))
                : c.testo;
            }),
            csvNum(totNotte(x)),
          ]);
      });
      out.push([
        "",
        `${t("rep.presTotGiorno")} ${grp.sede}`,
        grp.sede,
        "",
        ...giorni.map((g) => csvNum(totGiornoSede(grp.dips, g))),
        csvNum(grp.dips.reduce((s, x) => s + totDip(x), 0)),
      ]);
    }
    if (gruppi.length > 1)
      out.push([
        "",
        t("rep.presTotale"),
        "",
        "",
        ...giorni.map((g) => csvNum(totGiornoSede(tuttiDip, g))),
        csvNum(tuttiDip.reduce((s, x) => s + totDip(x), 0)),
      ]);
    return out;
  };

  const thDay = (g: string) => {
    const dow = isoDow(g);
    return (
      <th
        key={g}
        className={`px-1 py-1 text-center text-[10px] font-medium ${dow >= 6 ? "bg-muted/60 text-muted-foreground" : "text-muted-foreground"}`}
      >
        {ddmm(g)}
      </th>
    );
  };
  const tdCls = (tipo: string, g: string) => {
    const we = isoDow(g) >= 6 ? "bg-muted/40 " : "";
    switch (tipo) {
      case "nc":
        return `${we}text-status-absent font-semibold`;
      case "codice":
        return `${we}text-primary`;
      case "we":
      case "vuoto":
        return `${we}text-muted-foreground/60`;
      case "orePerm":
        return `${we}text-foreground underline decoration-dotted`;
      default:
        return `${we}text-foreground`;
    }
  };

  return (
    <div className="rounded-2xl border border-border bg-card p-5 sm:p-6 shadow-[var(--shadow-card)]">
      <div className="mb-3 flex flex-wrap items-center gap-3">
        <span className="text-[15px] font-semibold text-foreground">{t("rep.presTitle")}</span>
        <span className="text-xs text-muted-foreground">
          {ddmm(from)} – {ddmm(to)} · {tuttiDip.length} {t("rep.presDipendenti")}
        </span>
        <label className="no-print ml-auto flex cursor-pointer items-center gap-1.5 text-xs text-muted-foreground">
          <input
            type="checkbox"
            className="accent-primary"
            checked={arrotonda}
            onChange={(e) => setArrotonda(e.target.checked)}
          />
          {t("rep.presArrotonda")}
        </label>
        <EsportaStampa
          nomeFile={nomeFile}
          testata={testataCsv}
          righe={righeCsv}
          disabled={loading || tuttiDip.length === 0}
        />
      </div>
      <p className="mb-3 text-xs text-muted-foreground">{t("rep.presDesc")}</p>

      {loading || righe === null ? (
        <div className="text-sm text-muted-foreground">{t("rep.calculating")}</div>
      ) : tuttiDip.length === 0 ? (
        <div className="text-sm text-muted-foreground">{t("rep.noData")}</div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-[12px]">
            <thead>
              <tr className="border-b border-border">
                <th className="px-1 py-1 text-left text-[10px] font-medium text-muted-foreground">
                  N.
                </th>
                <th className="min-w-44 px-1 py-1 text-left text-[10px] font-medium uppercase text-muted-foreground">
                  {t("common.employee")}
                </th>
                <th className="px-1 py-1 text-left text-[10px] font-medium text-muted-foreground">
                  {t("rep.presRiga")}
                </th>
                {giorni.map(thDay)}
                <th className="px-1 py-1 text-right text-[10px] font-semibold text-foreground">
                  TOT
                </th>
              </tr>
            </thead>
            {gruppi.map((grp) => (
              <tbody key={grp.sede}>
                <tr className="border-t-2 border-border bg-muted/30">
                  <td
                    colSpan={giorni.length + 4}
                    className="px-1 py-1 text-[11px] font-semibold uppercase"
                  >
                    {grp.sede}
                  </td>
                </tr>
                {grp.dips.map((x, i) => {
                  const notte = totNotte(x);
                  return (
                    <Fragment key={x.dipendenteId}>
                      <tr className="border-t border-border/40">
                        <td className="px-1 py-0.5 text-muted-foreground">{i + 1}</td>
                        <td
                          className="max-w-56 truncate px-1 py-0.5 font-medium"
                          title={x.nomeCompleto}
                        >
                          {x.nomeCompleto}
                        </td>
                        <td className="px-1 py-0.5 text-muted-foreground">Tot</td>
                        {giorni.map((g) => {
                          const c = cella(g, x.giorni[g]);
                          return (
                            <td
                              key={g}
                              className={`whitespace-nowrap px-1 py-0.5 text-center tabular-nums ${tdCls(c.tipo, g)}`}
                              title={
                                x.giorni[g]?.permessoOre
                                  ? `${t("rep.presPermesso")} ${fmtOre(x.giorni[g].permessoOre ?? 0)} h`
                                  : x.giorni[g]?.codice === "NC"
                                    ? t("rep.presNcTip")
                                    : undefined
                              }
                            >
                              {c.testo}
                            </td>
                          );
                        })}
                        <td className="px-1 py-0.5 text-right font-semibold tabular-nums">
                          {fmtOre(totDip(x))}
                        </td>
                      </tr>
                      {notte > 0 && (
                        <tr className="text-muted-foreground">
                          <td />
                          <td className="px-1 py-0.5 text-[11px]" />
                          <td className="px-1 py-0.5">Not</td>
                          {giorni.map((g) => {
                            const c = cella(g, x.giorni[g]);
                            const n = r(x.giorni[g]?.notte ?? 0);
                            return (
                              <td
                                key={g}
                                className={`whitespace-nowrap px-1 py-0.5 text-center tabular-nums ${isoDow(g) >= 6 ? "bg-muted/40" : ""}`}
                              >
                                {c.tipo === "ore" || c.tipo === "orePerm"
                                  ? n > 0
                                    ? fmtOre(n)
                                    : "0"
                                  : c.testo}
                              </td>
                            );
                          })}
                          <td className="px-1 py-0.5 text-right tabular-nums">{fmtOre(notte)}</td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
                <tr className="border-t border-border font-semibold">
                  <td />
                  <td colSpan={2} className="px-1 py-1 text-[11px] uppercase">
                    {t("rep.presTotGiorno")} {grp.sede}
                  </td>
                  {giorni.map((g) => {
                    const v = totGiornoSede(grp.dips, g);
                    return (
                      <td
                        key={g}
                        className={`px-1 py-1 text-center tabular-nums ${isoDow(g) >= 6 ? "bg-muted/40" : ""}`}
                      >
                        {v > 0 ? fmtOre(v) : "0"}
                      </td>
                    );
                  })}
                  <td className="px-1 py-1 text-right tabular-nums">
                    {fmtOre(grp.dips.reduce((s, x) => s + totDip(x), 0))}
                  </td>
                </tr>
              </tbody>
            ))}
            {gruppi.length > 1 && (
              <tfoot>
                <tr className="border-t-2 border-border bg-muted/30 font-semibold">
                  <td />
                  <td colSpan={2} className="px-1 py-1 text-[11px] uppercase">
                    {t("rep.presTotale")}
                  </td>
                  {giorni.map((g) => {
                    const v = totGiornoSede(tuttiDip, g);
                    return (
                      <td key={g} className="px-1 py-1 text-center tabular-nums">
                        {v > 0 ? fmtOre(v) : "0"}
                      </td>
                    );
                  })}
                  <td className="px-1 py-1 text-right tabular-nums">
                    {fmtOre(tuttiDip.reduce((s, x) => s + totDip(x), 0))}
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      )}
      <p className="mt-3 text-[11px] text-muted-foreground">{t("rep.presLegenda")}</p>
    </div>
  );
}
