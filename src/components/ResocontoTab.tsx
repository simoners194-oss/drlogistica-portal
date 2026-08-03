// DR Portal — Finanza → tab Resoconto (direttore).
// Il resoconto per controparte: si scelgono uno o più CLIENTI e, col secondo
// filtro, uno o più FORNITORI, e si ottiene la riconciliazione del file del
// direttore — Estratto conto vs Fatture attive vs Fatture passive, con l'OK
// verde quando torna — più lo specchietto dei RITARDI (da incassare e da
// pagare) filtrabile per fasce di giorni/settimane, con le fatture in vista.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { CheckCircle2, AlertTriangle, Loader2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { MultiSelect } from "@/components/MultiSelect";
import {
  computeStatoFattura,
  collegaNoteCredito,
  fattureEscluse,
  isNotaCredito,
  parseIncassoAruba,
  type FatturaRaw,
  type TerminePagamento,
} from "@/lib/fatture-logic";
import { clienteGroupKey } from "@/lib/finanza-logic";
import { spGetFatture, spGetMovimenti, spGetTerminiPagamento } from "@/lib/sharepoint.functions";
import type { SpFattura, SpMovimento } from "@/lib/sharepoint.server";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

// Fasce di ritardo (giorni): parlano in settimane/mesi come chiede il
// direttore. null = nessun limite superiore.
const FASCE: { label: string; da: number; a: number | null }[] = [
  { label: "1–7 gg", da: 1, a: 7 },
  { label: "1–2 sett", da: 8, a: 14 },
  { label: "2–4 sett", da: 15, a: 30 },
  { label: "1–2 mesi", da: 31, a: 60 },
  { label: "2–3 mesi", da: 61, a: 90 },
  { label: "oltre 3 mesi", da: 91, a: null },
];

export function ResocontoTab() {
  const { t } = useLang();
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [clientiSel, setClientiSel] = useState<string[]>([]);
  const [fornitoriSel, setFornitoriSel] = useState<string[]>([]);
  const [fasceSel, setFasceSel] = useState<number[]>([]); // indici in FASCE

  useEffect(() => {
    spGetFatture({ data: { direzione: "Emessa" } })
      .then((l) => setFattureEm(l as SpFattura[]))
      .catch((err) => {
        setFattureEm([]);
        toast.error(t("ft.errLoad"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetMovimenti()
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const oggiISO = new Date().toISOString().slice(0, 10);

  // Stati calcolati per direzione (senza abbinamenti banca: qui contano le
  // letture di FATTURAZIONE — le stesse dei pivot del direttore).
  const prepara = (fatture: SpFattura[]) => {
    const escluse = fattureEscluse(fatture);
    const nc = collegaNoteCredito(fatture, escluse);
    return fatture
      .filter((f) => !escluse.has(f.nomeFile))
      .map((f) => ({
        f,
        s: computeStatoFattura(f, 0, termini, oggiISO, nc.get(f.nomeFile)?.importo ?? 0),
      }));
  };
  const attive = useMemo(
    () => prepara(fattureEm ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [fattureEm, termini],
  );
  const passive = useMemo(
    () => prepara(fattureRic ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [fattureRic, termini],
  );

  // Opzioni dei due filtri, per fatturato/costo decrescente.
  const opzioni = (righe: typeof attive) => {
    const somme = new Map<string, { nome: string; tot: number }>();
    for (const x of righe) {
      const k = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const v = somme.get(k) ?? { nome: x.f.cliente, tot: 0 };
      v.tot += Math.abs(x.f.totale);
      somme.set(k, v);
    }
    return [...somme.entries()]
      .sort((a, b) => b[1].tot - a[1].tot)
      .map(([k, v]) => ({ v: k, label: v.nome }));
  };
  const opzioniClienti = useMemo(() => opzioni(attive), [attive]);
  const opzioniFornitori = useMemo(() => opzioni(passive), [passive]);

  // Incassato/pagato con la semantica dell'export (quella dei pivot): le NC
  // compensate pesano in negativo, le fatture al valore registrato.
  const incassatoDi = (x: (typeof attive)[number]) =>
    isNotaCredito(x.f.tipoDocumento)
      ? x.s.statoIncassi === "Pagata" || x.s.statoFatturazione === "Pagata"
        ? -Math.abs(x.f.totale)
        : 0
      : (x.s.incassatoIncassi ??
        (parseIncassoAruba(x.f.incassoAruba) === "Incassata"
          ? Math.max(0, x.f.totale - x.s.notaCredito)
          : 0));

  const inSelezione = (nome: string, sel: string[]) =>
    sel.length === 0 || sel.includes(clienteGroupKey(nome) || nome);

  // --- Riconciliazione: Estratto conto vs Attive vs Passive -----------------
  const quadro = useMemo(() => {
    const attSel = attive.filter((x) => inSelezione(x.f.cliente, clientiSel));
    const pasSel = passive.filter((x) => inSelezione(x.f.cliente, fornitoriSel));
    const chiavi = new Set([
      ...(clientiSel.length ? clientiSel : opzioniClienti.map((o) => o.v)),
      ...(fornitoriSel.length ? fornitoriSel : []),
    ]);
    // Con nessun filtro il quadro è GLOBALE: tutti i movimenti con
    // controparte, tutte le fatture.
    const tuttoSelezionato = clientiSel.length === 0 && fornitoriSel.length === 0;
    const movSel = (movimenti ?? []).filter((m) => {
      if (!m.cliente) return false;
      const k = clienteGroupKey(m.cliente) || m.cliente;
      return tuttoSelezionato ? true : chiavi.has(k);
    });
    const estratto = movSel.reduce((s, m) => s + m.importo, 0);
    const incassatoAtt = attSel.reduce((s, x) => s + incassatoDi(x), 0);
    const pagatoPas = pasSel.reduce((s, x) => s + incassatoDi(x), 0);
    const differenza = Math.round((estratto - incassatoAtt + pagatoPas) * 100) / 100;
    return {
      estratto: Math.round(estratto * 100) / 100,
      incassatoAtt: Math.round(incassatoAtt * 100) / 100,
      pagatoPas: Math.round(pagatoPas * 100) / 100,
      differenza,
      nMov: movSel.length,
      ok: Math.abs(differenza) <= 1,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attive, passive, movimenti, clientiSel, fornitoriSel, opzioniClienti]);

  // --- Ritardi (da incassare e da pagare), con fasce -------------------------
  const inFascia = (gg: number) => {
    if (fasceSel.length === 0) return true;
    return fasceSel.some((i) => {
      const fscia = FASCE[i];
      return gg >= fscia.da && (fscia.a == null || gg <= fscia.a);
    });
  };
  const ritardi = (righe: typeof attive, sel: string[]) =>
    righe
      .filter(
        (x) =>
          x.s.inRitardo &&
          x.s.residuo > 1 &&
          (x.s.statoIncassi != null || x.s.statoFatturazione != null) &&
          inSelezione(x.f.cliente, sel) &&
          inFascia(x.s.giorniRitardo),
      )
      .sort((a, b) => b.s.giorniRitardo - a.s.giorniRitardo);
  const ritardiAtt = useMemo(
    () => ritardi(attive, clientiSel),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [attive, clientiSel, fasceSel],
  );
  const ritardiPas = useMemo(
    () => ritardi(passive, fornitoriSel),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, fornitoriSel, fasceSel],
  );

  const loading = fattureEm == null || fattureRic == null || movimenti == null;

  const cardRitardi = (titolo: string, righe: typeof attive, vuoto: string) => (
    <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      <div className="flex items-baseline gap-3 mb-2">
        <span className="text-sm font-semibold text-foreground">{titolo}</span>
        <span className="text-xs text-muted-foreground">
          {righe.length} · {fmtImporto(righe.reduce((s, x) => s + x.s.residuo, 0))} €
        </span>
      </div>
      {righe.length === 0 ? (
        <p className="text-sm text-muted-foreground">{vuoto}</p>
      ) : (
        <div className="overflow-x-auto max-h-80 overflow-y-auto">
          <table className="w-full text-[13px]">
            <thead>
              <tr className="text-left text-[11px] text-muted-foreground">
                <th className="py-1 pr-2">{t("ft.numero")}</th>
                <th className="py-1 pr-2">{t("fin.cliente")}</th>
                <th className="py-1 pr-2">{t("ft.scadenza")}</th>
                <th className="py-1 pr-2 text-right">{t("rt.gg")}</th>
                <th className="py-1 pr-2 text-right">{t("ft.residuo")}</th>
              </tr>
            </thead>
            <tbody>
              {righe.slice(0, 200).map((x) => (
                <tr key={x.f.nomeFile} className="border-t border-border/40">
                  <td className="py-0.5 pr-2 whitespace-nowrap font-medium">{x.f.numero}</td>
                  <td className="py-0.5 pr-2 max-w-52 truncate">{x.f.cliente}</td>
                  <td className="py-0.5 pr-2 whitespace-nowrap text-muted-foreground">
                    {fmtData(x.s.scadenza)}
                  </td>
                  <td className="py-0.5 pr-2 text-right tabular-nums text-status-absent">
                    {x.s.giorniRitardo}
                  </td>
                  <td className="py-0.5 pr-2 text-right tabular-nums font-medium">
                    {fmtImporto(x.s.residuo)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  return (
    <div className="space-y-4">
      {/* Filtri: uno o più clienti E uno o più fornitori. */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-3">
          <MultiSelect
            label={t("fin.cliente")}
            tuttiLabel={t("common.allF")}
            selLabel={t("fin.msSel")}
            opzioni={opzioniClienti}
            valori={clientiSel}
            onChange={setClientiSel}
            className="w-64"
          />
          <MultiSelect
            label={t("ft.fornitore")}
            tuttiLabel={t("common.allF")}
            selLabel={t("fin.msSel")}
            opzioni={opzioniFornitori}
            valori={fornitoriSel}
            onChange={setFornitoriSel}
            className="w-64"
          />
          {/* Fasce di ritardo: giorni e settimane, multi-selezione. */}
          <div>
            <label className="text-xs text-muted-foreground">{t("rt.fasce")}</label>
            <div className="flex flex-wrap gap-1.5 pt-1.5">
              {FASCE.map((fscia, i) => (
                <button
                  key={fscia.label}
                  type="button"
                  onClick={() =>
                    setFasceSel(
                      fasceSel.includes(i) ? fasceSel.filter((x) => x !== i) : [...fasceSel, i],
                    )
                  }
                  className={`rounded-full border px-3 py-1.5 text-xs font-medium transition-colors ${
                    fasceSel.includes(i)
                      ? "border-primary bg-primary text-primary-foreground"
                      : "border-border bg-background text-foreground hover:bg-muted"
                  }`}
                >
                  {fscia.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      </div>

      {loading ? (
        <div className="py-10 text-center">
          <Loader2 className="h-5 w-5 animate-spin inline-block text-muted-foreground" />
        </div>
      ) : (
        <>
          {/* Lo specchietto del direttore: Situazione fatture vs Estratto
              conto, con l'OK verde quando la differenza è zero. */}
          <div
            className={`rounded-2xl border p-5 shadow-[var(--shadow-card)] ${
              quadro.ok
                ? "border-status-present/40 bg-status-present/5"
                : "border-status-absent/40 bg-status-absent/5"
            }`}
          >
            <div className="flex items-center gap-2 mb-3">
              {quadro.ok ? (
                <CheckCircle2 className="h-5 w-5 text-status-present" />
              ) : (
                <AlertTriangle className="h-5 w-5 text-status-absent" />
              )}
              <span className="text-sm font-semibold text-foreground">{t("rt.titolo")}</span>
              <span
                className={`ml-auto rounded-full px-3 py-1 text-sm font-bold ${
                  quadro.ok
                    ? "bg-status-present/15 text-status-present"
                    : "bg-status-absent/15 text-status-absent"
                }`}
              >
                {quadro.ok ? "OK" : fmtImporto(quadro.differenza) + " €"}
              </span>
            </div>
            <div className="grid gap-3 sm:grid-cols-4">
              <div>
                <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.estratto")}
                </div>
                <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
                  {fmtImporto(quadro.estratto)} €
                </div>
                <div className="text-[11px] text-muted-foreground">{quadro.nMov} mov.</div>
              </div>
              <div>
                <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.attive")}
                </div>
                <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
                  {fmtImporto(quadro.incassatoAtt)} €
                </div>
              </div>
              <div>
                <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.passive")}
                </div>
                <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
                  {fmtImporto(quadro.pagatoPas)} €
                </div>
              </div>
              <div>
                <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.differenza")}
                </div>
                <div
                  className={`mt-0.5 text-xl font-semibold tabular-nums ${
                    quadro.ok ? "text-status-present" : "text-status-absent"
                  }`}
                >
                  {fmtImporto(quadro.differenza)} €
                </div>
              </div>
            </div>
            <p className="mt-2 text-[11px] text-muted-foreground">{t("rt.nota")}</p>
          </div>

          {/* I ritardi, nelle due direzioni, con le fatture in chiaro. */}
          <div className="grid gap-4 lg:grid-cols-2">
            {cardRitardi(t("rt.ritardiIncassare"), ritardiAtt, t("rt.nessunoIncassare"))}
            {cardRitardi(t("rt.ritardiPagare"), ritardiPas, t("rt.nessunoPagare"))}
          </div>
        </>
      )}
    </div>
  );
}
