// DR Portal — Finanza → tab Fornitori (Simone 29/09/2026, v1.84.0).
// "Una scheda in cui c'è la spiegazione delle fatture in entrata presenti
// nel flusso di cassa": una riga per fornitore con fatture passive aperte
// nei Flussi — nome, chi se ne occupa in DR, macrovoce, appalto (scritti a
// mano, salvati sulla lista FlussiCassa genere "fornitore") e, dalle
// fatture: scaduto, aperto totale, conteggi, prima/ultima fattura, prossima
// scadenza, fatture pagate e importi pagati ad oggi. Qui si forza anche la
// modalità di pagamento (RID / RiBa / nessuna) che divide le uscite dei
// Flussi; senza forzatura vale l'XML di ogni fattura.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Loader2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { EsportaStampa } from "@/components/EsportaStampa";
import {
  computeStatoFattura,
  collegaNoteCredito,
  fattureEscluse,
  residuoAperto,
  type TerminePagamento,
} from "@/lib/fatture-logic";
import {
  esclusaDaFlussi,
  mappaFornitori,
  riepilogoFornitori,
  serializeFornitoreInfo,
  type FornitoreInfo,
  type ModalitaPagamento,
  type RigaFornitore,
} from "@/lib/flussi-logic";
import {
  spGetFatture,
  spGetFlussiCassa,
  spGetTerminiPagamento,
  spUpsertFlussoCassa,
} from "@/lib/sharepoint.functions";
import { csvData } from "@/lib/csv";
import type { FlussoCassaRiga, SpFattura } from "@/lib/sharepoint.server";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
const dataIt = (iso: string) =>
  /^\d{4}-\d{2}-\d{2}/.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : "";

type CampoTesto = "referente" | "macrovoce" | "appalto";
type Ordine = "aperto" | "scaduto" | "nome" | "prossima";

const BADGE: Record<ModalitaPagamento, string> = { rid: "RID", riba: "RiBa", altro: "—" };

export function FornitoriFlussoTab() {
  const { t } = useLang();
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [flussi, setFlussi] = useState<FlussoCassaRiga[] | null>(null);
  const [errore, setErrore] = useState<string | null>(null);
  const [cerca, setCerca] = useState("");
  const [ordine, setOrdine] = useState<Ordine>("aperto");
  // Testo in modifica per fornitore (prima del salvataggio al blur).
  const [drafts, setDrafts] = useState<Record<string, Partial<Record<CampoTesto, string>>>>({});
  const [salvando, setSalvando] = useState<Set<string>>(new Set());

  const ricaricaFlussi = () =>
    spGetFlussiCassa()
      .then((l) => {
        setFlussi(l as FlussoCassaRiga[]);
        setErrore(null);
      })
      .catch((err) => {
        setFlussi([]);
        setErrore(err instanceof Error ? err.message : String(err));
      });
  useEffect(() => {
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    void ricaricaFlussi();
  }, []);

  const oggiISO = new Date().toISOString().slice(0, 10);
  const esclusioni = useMemo(
    () => (flussi ?? []).filter((x) => x.genere === "esclusione"),
    [flussi],
  );
  const info = useMemo(() => mappaFornitori(flussi ?? []), [flussi]);

  // Stessa preparazione dei Flussi: stati calcolati, NC collegate, DR
  // Logistics fuori (regola FR 08/09: a DR Logistics si deve solo la girata).
  const passive = useMemo(() => {
    const fatture = fattureRic ?? [];
    const escluse = fattureEscluse(fatture);
    const nc = collegaNoteCredito(fatture, escluse);
    return fatture
      .filter((f) => !escluse.has(f.nomeFile))
      .filter((f) => !f.cliente.toLowerCase().includes("dr logistics"))
      .map((f) => ({
        f,
        s: computeStatoFattura(f, 0, termini, oggiISO, nc.get(f.nomeFile)?.importo ?? 0),
      }));
  }, [fattureRic, termini, oggiISO]);

  const righe = useMemo(
    () =>
      riepilogoFornitori(
        passive,
        info,
        (x) =>
          residuoAperto(x) > 1 &&
          !!x.s.scadenza &&
          !esclusaDaFlussi(esclusioni, x.f.cliente, x.s.scadenza.slice(0, 7)),
      ),
    [passive, info, esclusioni],
  );

  const visibili = useMemo(() => {
    const c = cerca.trim().toLowerCase();
    const l = c
      ? righe.filter((r) =>
          [r.nome, r.referente, r.macrovoce, r.appalto].some((v) => v.toLowerCase().includes(c)),
        )
      : righe;
    const cmp: Record<Ordine, (a: RigaFornitore, b: RigaFornitore) => number> = {
      aperto: (a, b) => b.aperto - a.aperto,
      scaduto: (a, b) => b.scaduto - a.scaduto,
      nome: (a, b) => a.nome.localeCompare(b.nome, "it"),
      prossima: (a, b) =>
        (a.prossimaScadenza || "9999").localeCompare(b.prossimaScadenza || "9999"),
    };
    return [...l].sort((a, b) => cmp[ordine](a, b) || a.nome.localeCompare(b.nome, "it"));
  }, [righe, cerca, ordine]);

  const totali = useMemo(
    () =>
      visibili.reduce(
        (acc, r) => ({
          scaduto: acc.scaduto + r.scaduto,
          aperto: acc.aperto + r.aperto,
          nScadute: acc.nScadute + r.nScadute,
          nAperte: acc.nAperte + r.nAperte,
          nTotale: acc.nTotale + r.nTotale,
          nPagate: acc.nPagate + r.nPagate,
          pagato: acc.pagato + r.pagato,
        }),
        { scaduto: 0, aperto: 0, nScadute: 0, nAperte: 0, nTotale: 0, nPagate: 0, pagato: 0 },
      ),
    [visibili],
  );

  // Valori già usati (per i suggerimenti delle caselle Macrovoce/Appalto).
  const vocabolario = useMemo(() => {
    const macro = new Set<string>();
    const appalti = new Set<string>();
    const ref = new Set<string>();
    for (const r of righe) {
      if (r.macrovoce) macro.add(r.macrovoce);
      if (r.macrovoceAuto) macro.add(r.macrovoceAuto);
      if (r.appalto) appalti.add(r.appalto);
      if (r.appaltoAuto) appalti.add(r.appaltoAuto);
      if (r.referente) ref.add(r.referente);
    }
    const ord = (s: Set<string>) => [...s].sort((a, b) => a.localeCompare(b, "it"));
    return { macro: ord(macro), appalti: ord(appalti), ref: ord(ref) };
  }, [righe]);

  const valore = (r: RigaFornitore, campo: CampoTesto) => drafts[r.chiave]?.[campo] ?? r[campo];

  /** Salva la scheda del fornitore (una riga FlussiCassa genere "fornitore"):
   *  i testi in bozza più l'eventuale nuova modalità. */
  const salva = async (r: RigaFornitore, patch: Partial<FornitoreInfo>) => {
    const nuovo: FornitoreInfo = {
      modalita: r.override,
      referente: valore(r, "referente"),
      macrovoce: valore(r, "macrovoce"),
      appalto: valore(r, "appalto"),
      ...patch,
    };
    setSalvando((s) => new Set(s).add(r.chiave));
    try {
      await spUpsertFlussoCassa({
        data: {
          nome: r.nome,
          genere: "fornitore",
          importo: 0,
          note: serializeFornitoreInfo(nuovo),
        },
      });
      await ricaricaFlussi();
      setDrafts((d) => {
        const nd = { ...d };
        delete nd[r.chiave];
        return nd;
      });
      toast.success(t("for.salvato"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSalvando((s) => {
        const ns = new Set(s);
        ns.delete(r.chiave);
        return ns;
      });
    }
  };

  const scrivi = (r: RigaFornitore, campo: CampoTesto, v: string) =>
    setDrafts((d) => ({ ...d, [r.chiave]: { ...d[r.chiave], [campo]: v } }));
  const alBlur = (r: RigaFornitore) => {
    const d = drafts[r.chiave];
    if (!d) return;
    const cambiato = (["referente", "macrovoce", "appalto"] as const).some(
      (c) => d[c] != null && d[c].trim() !== r[c],
    );
    if (cambiato) void salva(r, {});
  };

  const testataCsv = [
    t("for.colFornitore"),
    t("for.colModalita"),
    t("for.colReferente"),
    t("for.colMacro"),
    t("for.colAppalto"),
    t("for.colScaduto"),
    t("for.colAperto"),
    t("for.colNScad"),
    t("for.colNAperte"),
    t("for.colNTot"),
    t("for.colPrima"),
    t("for.colProssima"),
    t("for.colUltima"),
    t("for.colNPagate"),
    t("for.colPagato"),
  ];
  const righeCsv = () =>
    visibili.map((r) => [
      r.nome,
      r.override ? BADGE[r.override] : `${BADGE[r.modalitaAuto]} (auto)`,
      valore(r, "referente"),
      valore(r, "macrovoce") || r.macrovoceAuto,
      valore(r, "appalto") || r.appaltoAuto,
      r.scaduto.toFixed(2).replace(".", ","),
      r.aperto.toFixed(2).replace(".", ","),
      r.nScadute,
      r.nAperte,
      r.nTotale,
      csvData(r.primaFattura),
      csvData(r.prossimaScadenza),
      csvData(r.ultimaFattura),
      r.nPagate,
      r.pagato.toFixed(2).replace(".", ","),
    ]);

  const loading = fattureRic == null || flussi == null;
  const inputCls =
    "w-full min-w-24 rounded border border-border bg-background px-1.5 py-0.5 text-[12px] text-foreground placeholder:italic placeholder:text-muted-foreground/70";
  const th = "py-1.5 pr-3 text-left text-[11px] font-medium uppercase text-muted-foreground";
  const thN = `${th} text-right`;
  const tdN = "py-1 pr-3 text-right tabular-nums whitespace-nowrap";

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-1 flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold text-foreground">{t("for.titolo")}</span>
          {!loading && (
            <span className="text-xs text-muted-foreground">
              {visibili.length} {t("for.fornitori")}
            </span>
          )}
          {salvando.size > 0 && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
        </div>
        <p className="mb-3 text-xs text-muted-foreground">{t("for.desc")}</p>

        <div className="no-print mb-3 flex flex-wrap items-center gap-2 text-[13px]">
          <input
            value={cerca}
            onChange={(e) => setCerca(e.target.value)}
            placeholder={t("for.cercaPh")}
            className="w-64 rounded-lg border border-border bg-background px-2 py-1 text-[13px]"
          />
          <label className="flex items-center gap-1.5 text-xs text-muted-foreground">
            {t("for.ordina")}
            <select
              value={ordine}
              onChange={(e) => setOrdine(e.target.value as Ordine)}
              className="rounded-lg border border-border bg-background px-2 py-1 text-[13px] text-foreground"
            >
              <option value="aperto">{t("for.ordAperto")}</option>
              <option value="scaduto">{t("for.ordScaduto")}</option>
              <option value="prossima">{t("for.ordProssima")}</option>
              <option value="nome">{t("for.ordNome")}</option>
            </select>
          </label>
          <EsportaStampa
            nomeFile="fornitori-flusso-di-cassa"
            testata={testataCsv}
            righe={righeCsv}
            disabled={loading}
            className="ml-auto"
          />
        </div>

        {errore && (
          <p className="mb-3 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-900">
            {errore}
          </p>
        )}

        {loading ? (
          <div className="py-10 text-center text-sm text-muted-foreground">
            <Loader2 className="inline-block h-5 w-5 animate-spin" />
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-[13px]">
              <thead>
                <tr className="border-b border-border">
                  <th className={`${th} min-w-44`}>{t("for.colFornitore")}</th>
                  <th className={th}>{t("for.colModalita")}</th>
                  <th className={th}>{t("for.colReferente")}</th>
                  <th className={th}>{t("for.colMacro")}</th>
                  <th className={th}>{t("for.colAppalto")}</th>
                  <th className={thN}>{t("for.colScaduto")}</th>
                  <th className={thN}>{t("for.colAperto")}</th>
                  <th className={thN}>{t("for.colNScad")}</th>
                  <th className={thN}>{t("for.colNAperte")}</th>
                  <th className={thN}>{t("for.colNTot")}</th>
                  <th className={th}>{t("for.colPrima")}</th>
                  <th className={th}>{t("for.colProssima")}</th>
                  <th className={th}>{t("for.colUltima")}</th>
                  <th className={thN}>{t("for.colNPagate")}</th>
                  <th className={thN}>{t("for.colPagato")}</th>
                </tr>
              </thead>
              <tbody>
                {visibili.map((r) => (
                  <tr key={r.chiave} className="border-b border-border/40 align-top">
                    <td className="max-w-64 py-1 pr-3">
                      <div className="truncate font-medium" title={r.nome}>
                        {r.nome}
                      </div>
                      {r.mista && (
                        <div className="text-[10px] text-muted-foreground">{t("for.mista")}</div>
                      )}
                    </td>
                    <td className="py-1 pr-3">
                      <select
                        value={r.override ?? ""}
                        disabled={salvando.has(r.chiave)}
                        title={t("for.modAutoTip")}
                        onChange={(e) =>
                          void salva(r, {
                            modalita: (e.target.value || undefined) as
                              ModalitaPagamento | undefined,
                          })
                        }
                        className={`rounded border bg-background px-1 py-0.5 text-[12px] text-foreground ${r.override ? "border-primary" : "border-border"}`}
                      >
                        <option value="">
                          {t("for.modAuto")} ({BADGE[r.modalitaAuto]}
                          {r.mista ? " ±" : ""})
                        </option>
                        <option value="rid">{t("for.modRid")}</option>
                        <option value="riba">{t("for.modRiba")}</option>
                        <option value="altro">{t("for.modAltro")}</option>
                      </select>
                    </td>
                    <td className="py-1 pr-3">
                      <input
                        list="for-ref"
                        value={valore(r, "referente")}
                        placeholder={t("for.refPh")}
                        disabled={salvando.has(r.chiave)}
                        onChange={(e) => scrivi(r, "referente", e.target.value)}
                        onBlur={() => alBlur(r)}
                        onKeyDown={(e) =>
                          e.key === "Enter" && (e.target as HTMLInputElement).blur()
                        }
                        className={inputCls}
                      />
                    </td>
                    <td className="py-1 pr-3">
                      <input
                        list="for-macro"
                        value={valore(r, "macrovoce")}
                        placeholder={r.macrovoceAuto || t("for.macroPh")}
                        title={r.macrovoceAuto}
                        disabled={salvando.has(r.chiave)}
                        onChange={(e) => scrivi(r, "macrovoce", e.target.value)}
                        onBlur={() => alBlur(r)}
                        onKeyDown={(e) =>
                          e.key === "Enter" && (e.target as HTMLInputElement).blur()
                        }
                        className={inputCls}
                      />
                    </td>
                    <td className="py-1 pr-3">
                      <input
                        list="for-appalti"
                        value={valore(r, "appalto")}
                        placeholder={r.appaltoAuto || t("for.appaltoPh")}
                        title={r.appaltoAuto}
                        disabled={salvando.has(r.chiave)}
                        onChange={(e) => scrivi(r, "appalto", e.target.value)}
                        onBlur={() => alBlur(r)}
                        onKeyDown={(e) =>
                          e.key === "Enter" && (e.target as HTMLInputElement).blur()
                        }
                        className={inputCls}
                      />
                    </td>
                    <td className={`${tdN} ${r.scaduto > 0.005 ? "text-status-absent" : ""}`}>
                      {r.scaduto > 0.005 ? fmtImporto(r.scaduto) : "—"}
                    </td>
                    <td className={`${tdN} font-medium`}>{fmtImporto(r.aperto)}</td>
                    <td className={tdN}>{r.nScadute || "—"}</td>
                    <td className={tdN}>{r.nAperte}</td>
                    <td className={tdN}>{r.nTotale}</td>
                    <td className="whitespace-nowrap py-1 pr-3 text-muted-foreground">
                      {dataIt(r.primaFattura)}
                    </td>
                    <td
                      className={`whitespace-nowrap py-1 pr-3 ${r.prossimaScadenza && r.prossimaScadenza < oggiISO ? "text-status-absent" : ""}`}
                    >
                      {dataIt(r.prossimaScadenza)}
                    </td>
                    <td className="whitespace-nowrap py-1 pr-3 text-muted-foreground">
                      {dataIt(r.ultimaFattura)}
                    </td>
                    <td className={tdN}>{r.nPagate || "—"}</td>
                    <td className={tdN}>
                      {Math.abs(r.pagato) > 0.005 ? fmtImporto(r.pagato) : "—"}
                    </td>
                  </tr>
                ))}
                {visibili.length === 0 && (
                  <tr>
                    <td colSpan={15} className="py-6 text-center text-muted-foreground">
                      {t("fc.simNessuna")}
                    </td>
                  </tr>
                )}
              </tbody>
              {visibili.length > 0 && (
                <tfoot>
                  <tr className="border-t-2 border-border font-semibold">
                    <td className="py-1.5 pr-3" colSpan={5}>
                      {t("for.totale")} ({visibili.length})
                    </td>
                    <td className={`${tdN} text-status-absent`}>{fmtImporto(totali.scaduto)}</td>
                    <td className={tdN}>{fmtImporto(totali.aperto)}</td>
                    <td className={tdN}>{totali.nScadute}</td>
                    <td className={tdN}>{totali.nAperte}</td>
                    <td className={tdN}>{totali.nTotale}</td>
                    <td colSpan={3} />
                    <td className={tdN}>{totali.nPagate}</td>
                    <td className={tdN}>{fmtImporto(totali.pagato)}</td>
                  </tr>
                </tfoot>
              )}
            </table>
            <datalist id="for-ref">
              {vocabolario.ref.map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
            <datalist id="for-macro">
              {vocabolario.macro.map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
            <datalist id="for-appalti">
              {vocabolario.appalti.map((v) => (
                <option key={v} value={v} />
              ))}
            </datalist>
          </div>
        )}
        <p className="mt-3 text-[11px] text-muted-foreground">{t("for.nota")}</p>
      </div>
    </div>
  );
}
