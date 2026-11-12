// DR Portal — Finanze → "Da correggere su Aruba" (Simone 12/11/2026: "metti
// una pagina di alert con tutte le fatture che a tuo avviso vanno corrette su
// Aruba").
//
// Aruba è la fonte dello stato delle fatture; la banca però arriva prima. Qui
// stanno le fatture in cui le due cose non tornano e la correzione va fatta
// SU ARUBA (il portale la rilegge col giro del pomeriggio):
//   A. passive pagate in banca ma ancora "da pagare" su Aruba;
//   B. attive incassate in banca ma ancora "da incassare" su Aruba;
//   C. passive pagate con la carta (flag della scheda fornitore), chiuse nel
//      portale in attesa di Aruba;
//   D. fatture date per pagate/incassate su Aruba con una data nel futuro.
// A e B vengono da lib/pagamenti-banca (stessa controparte, stesso importo al
// centesimo, movimento non prima della fattura): sono proposte, da guardare.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Loader2, RotateCcw, TriangleAlert } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import {
  collegaNoteCredito,
  computeStatoFattura,
  fattureEscluse,
  isNotaCredito,
  parseIncassoAruba,
  residuoAperto,
  type FatturaConStato,
  type TerminePagamento,
} from "@/lib/fatture-logic";
import { agganciaPagamenti, type AggancioBanca } from "@/lib/pagamenti-banca";
import type { SpFattura, SpMovimento } from "@/lib/sharepoint.server";
import {
  spGetFatture,
  spGetMovimenti,
  spGetTerminiPagamento,
  spSetPagataCdc,
} from "@/lib/sharepoint.functions";

type Tipo = "pagate" | "incassate" | "carta" | "future";
interface Riga {
  tipo: Tipo;
  nomeFile: string;
  controparte: string;
  numero: string;
  dataFattura: string;
  importo: number;
  /** Data da registrare su Aruba (o la data sbagliata, per "future"). */
  quando: string;
  come: string;
}

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dataIt = (iso: string) =>
  /^\d{4}-\d{2}-\d{2}/.test(iso)
    ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`
    : "—";
const ddmm = (iso: string) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}` : "");

const STILE: Record<Tipo, { testa: string; punto: string }> = {
  pagate: { testa: "bg-rose-600 text-white", punto: "bg-rose-600" },
  incassate: { testa: "bg-emerald-600 text-white", punto: "bg-emerald-600" },
  carta: { testa: "bg-violet-600 text-white", punto: "bg-violet-600" },
  future: { testa: "bg-amber-500 text-white", punto: "bg-amber-500" },
};
const ORDINE: Tipo[] = ["pagate", "incassate", "carta", "future"];

export function ArubaAlertTab() {
  const { t } = useLang();
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [busy, setBusy] = useState<string | null>(null);

  useEffect(() => {
    spGetFatture({ data: { direzione: "Emessa" } })
      .then((l) => setFattureEm(l as SpFattura[]))
      .catch(() => setFattureEm([]));
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    spGetMovimenti()
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
  }, []);

  const oggi = new Date().toISOString().slice(0, 10);
  const pronto = fattureEm && fattureRic && termini && movimenti;

  const righe = useMemo((): Riga[] => {
    if (!fattureEm || !fattureRic || !termini || !movimenti) return [];
    // Stessa preparazione dei Flussi: stati calcolati, note di credito collegate.
    const prepara = (fatture: SpFattura[]): FatturaConStato[] => {
      const escluse = fattureEscluse(fatture);
      const nc = collegaNoteCredito(fatture, escluse);
      return fatture
        .filter((f) => !escluse.has(f.nomeFile))
        .map((f) => ({
          f,
          s: computeStatoFattura(f, 0, termini, oggi, nc.get(f.nomeFile)?.importo ?? 0),
        }));
    };
    const attive = prepara(fattureEm);
    const passive = prepara(fattureRic);
    const come = (a: AggancioBanca, incasso: boolean): string => {
      const mov = t(incasso ? "aa.accredito" : "aa.bonifico");
      if (a.modo === "somma")
        return `${t("aa.unSolo")} ${mov} ${ddmm(a.data)} ${t("aa.insiemeA")} ${(a.insieme ?? []).join(", ")}`;
      if (a.modo === "rate")
        return `${a.movimenti.length} ${t("aa.movimenti")}: ${a.movimenti
          .map((m) => `${ddmm(m.data)} ${fmt(m.importo)}`)
          .join(" + ")}`;
      return `${mov} ${ddmm(a.data)}`;
    };
    const out: Riga[] = [];
    const daAgganci = (lista: FatturaConStato[], tipo: Tipo, incasso: boolean) => {
      const agganci = agganciaPagamenti(lista, movimenti, incasso ? "Emessa" : "Ricevuta");
      for (const x of lista) {
        const a = agganci.get(x.f.nomeFile);
        if (!a) continue;
        out.push({
          tipo,
          nomeFile: x.f.nomeFile,
          controparte: x.f.cliente,
          numero: x.f.numero,
          dataFattura: x.f.dataDocumento,
          importo: residuoAperto(x),
          quando: a.data,
          come: come(a, incasso),
        });
      }
    };
    daAgganci(passive, "pagate", false);
    daAgganci(attive, "incassate", true);
    for (const x of passive) {
      if (!x.f.pagataCdc || isNotaCredito(x.f.tipoDocumento)) continue;
      // Ancora aperta per Aruba? (il flag della carta, da solo, la chiude)
      const perAruba = residuoAperto({ ...x, f: { ...x.f, pagataCdc: undefined } });
      if (perAruba <= 1) continue;
      out.push({
        tipo: "carta",
        nomeFile: x.f.nomeFile,
        controparte: x.f.cliente,
        numero: x.f.numero,
        dataFattura: x.f.dataDocumento,
        importo: perAruba,
        quando: x.f.pagataCdc,
        come: t("aa.conCarta"),
      });
    }
    for (const x of [...passive, ...attive]) {
      const di = x.f.dataIncasso ?? "";
      if (di <= oggi || parseIncassoAruba(x.f.incassoAruba) !== "Incassata") continue;
      out.push({
        tipo: "future",
        nomeFile: x.f.nomeFile,
        controparte: x.f.cliente,
        numero: x.f.numero,
        dataFattura: x.f.dataDocumento,
        importo: Math.abs(x.f.totale),
        quando: di,
        come: t(x.f.direzione === "Emessa" ? "aa.futuraIncasso" : "aa.futuraPagamento"),
      });
    }
    return out.sort(
      (a, b) =>
        ORDINE.indexOf(a.tipo) - ORDINE.indexOf(b.tipo) ||
        b.quando.localeCompare(a.quando) ||
        a.controparte.localeCompare(b.controparte, "it"),
    );
  }, [fattureEm, fattureRic, termini, movimenti, oggi, t]);

  const riapri = async (r: Riga) => {
    setBusy(r.nomeFile);
    try {
      await spSetPagataCdc({ data: { nomeFile: r.nomeFile, data: "" } });
      setFattureRic((await spGetFatture({ data: { direzione: "Ricevuta" } })) as SpFattura[]);
      toast.success(t("aa.riaperta"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setBusy(null);
    }
  };

  const titolo: Record<Tipo, string> = {
    pagate: t("aa.secPagate"),
    incassate: t("aa.secIncassate"),
    carta: t("aa.secCarta"),
    future: t("aa.secFuture"),
  };
  const azione: Record<Tipo, string> = {
    pagate: t("aa.azPagata"),
    incassate: t("aa.azIncassata"),
    carta: t("aa.azCarta"),
    future: t("aa.azData"),
  };
  const spiega: Record<Tipo, string> = {
    pagate: t("aa.spPagate"),
    incassate: t("aa.spIncassate"),
    carta: t("aa.spCarta"),
    future: t("aa.spFuture"),
  };

  if (!pronto)
    return (
      <div className="flex items-center gap-2 text-sm text-muted-foreground">
        <Loader2 className="h-4 w-4 animate-spin" /> {t("common.loading")}
      </div>
    );

  const perTipo = (tipo: Tipo) => righe.filter((r) => r.tipo === tipo);
  const somma = (l: Riga[]) => l.reduce((s, r) => s + r.importo, 0);

  return (
    <div className="space-y-4">
      {/* Testata: cosa c'è da fare, in breve */}
      <div className="rounded-2xl border border-border bg-card px-5 py-4 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-start gap-x-6 gap-y-3">
          <div className="min-w-64 flex-1">
            <div className="flex items-center gap-2 text-sm font-semibold text-foreground">
              <TriangleAlert className="h-4 w-4 text-amber-500" /> {t("aa.titolo")}
              <span className="rounded-full bg-amber-500/15 px-2 py-0.5 text-[11px] font-bold text-amber-800 dark:text-amber-300">
                {righe.length}
              </span>
            </div>
            <p className="mt-1 max-w-3xl text-xs text-muted-foreground">{t("aa.intro")}</p>
          </div>
          <div className="flex flex-wrap items-center gap-x-5 gap-y-2">
            {ORDINE.map((tipo) => {
              const l = perTipo(tipo);
              return (
                <div key={tipo} className="border-l-2 border-border pl-3">
                  <div className="flex items-center gap-1.5 text-[11px] uppercase tracking-wider text-muted-foreground">
                    <span className={`h-2 w-2 rounded-full ${STILE[tipo].punto}`} />
                    {titolo[tipo]}
                  </div>
                  <div className="mt-0.5 text-lg font-semibold tabular-nums text-foreground">
                    {l.length}
                    <span className="ml-2 text-[12px] font-normal text-muted-foreground">
                      {fmt(somma(l))} €
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
          <EsportaStampa
            nomeFile={`da-correggere-su-aruba-${oggi}`}
            testata={[
              t("aa.colCosa"),
              t("aa.colControparte"),
              t("aa.colFattura"),
              t("aa.colDataFattura"),
              t("aa.colImporto"),
              t("aa.colAzione"),
              t("aa.colQuando"),
              t("aa.colCome"),
            ]}
            righe={() =>
              righe.map((r) => [
                titolo[r.tipo],
                r.controparte,
                r.numero,
                r.dataFattura,
                r.importo.toFixed(2).replace(".", ","),
                azione[r.tipo],
                r.quando,
                r.come,
              ])
            }
          />
        </div>
      </div>

      {righe.length === 0 && (
        <div className="rounded-2xl border border-border bg-card px-5 py-6 text-sm text-muted-foreground shadow-[var(--shadow-card)]">
          {t("aa.nessuna")}
        </div>
      )}

      {ORDINE.map((tipo) => {
        const l = perTipo(tipo);
        if (!l.length) return null;
        return (
          <div
            key={tipo}
            className="overflow-hidden rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]"
          >
            <div
              className={`flex flex-wrap items-center justify-between gap-2 px-4 py-2 ${STILE[tipo].testa}`}
            >
              <span className="text-[12px] font-semibold uppercase tracking-wide">
                {titolo[tipo]}
              </span>
              <span className="text-[13px] font-bold tabular-nums">
                {l.length} · {fmt(somma(l))} €
              </span>
            </div>
            <p className="border-b border-border/60 px-4 py-1.5 text-[11.5px] text-muted-foreground">
              {spiega[tipo]}
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-[13px]">
                <thead>
                  <tr className="border-b border-border text-left text-[11px] uppercase text-muted-foreground">
                    <th className="px-4 py-1.5">{t("aa.colControparte")}</th>
                    <th className="px-2 py-1.5">{t("aa.colFattura")}</th>
                    <th className="px-2 py-1.5">{t("aa.colDataFattura")}</th>
                    <th className="px-2 py-1.5 text-right">{t("aa.colImporto")}</th>
                    <th className="px-2 py-1.5">{t("aa.colAzione")}</th>
                    <th className="px-2 py-1.5">{t("aa.colCome")}</th>
                    {tipo === "carta" && <th className="px-2 py-1.5" />}
                  </tr>
                </thead>
                <tbody>
                  {l.map((r) => (
                    <tr key={`${r.tipo}|${r.nomeFile}`} className="border-b border-border/40">
                      <td className="px-4 py-1.5 font-medium text-foreground">{r.controparte}</td>
                      <td className="whitespace-nowrap px-2 py-1.5">{r.numero}</td>
                      <td className="whitespace-nowrap px-2 py-1.5 text-muted-foreground">
                        {dataIt(r.dataFattura)}
                      </td>
                      <td className="whitespace-nowrap px-2 py-1.5 text-right font-medium tabular-nums">
                        {fmt(r.importo)}
                      </td>
                      <td className="whitespace-nowrap px-2 py-1.5">
                        {azione[r.tipo]} <b>{dataIt(r.quando)}</b>
                      </td>
                      <td className="px-2 py-1.5 text-[12px] text-muted-foreground">{r.come}</td>
                      {tipo === "carta" && (
                        <td className="px-2 py-1.5 text-right">
                          <button
                            type="button"
                            disabled={busy === r.nomeFile}
                            onClick={() => void riapri(r)}
                            title={t("aa.riapriTip")}
                            className="inline-flex items-center gap-1 rounded-md border border-border px-2 py-0.5 text-[11px] text-muted-foreground hover:bg-muted hover:text-foreground disabled:opacity-40"
                          >
                            <RotateCcw className="h-3 w-3" /> {t("aa.riapri")}
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
    </div>
  );
}
