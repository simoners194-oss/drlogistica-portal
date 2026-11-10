// DR Portal — Finanza → Flussi → Cassa proprietà: le schede che si aprono
// con un clic sui nomi del foglio (Simone 04/11/2026: "me li link come su
// vista classica"). I fornitori aprono la scheda fornitore dei Flussi
// (SchedaFornitore, dal tab); qui vivono la scheda del CLIENTE (le sue
// fatture aperte, dalle attive) e l'elenco dei fornitori dentro un gruppo
// (uscite raggruppate per tipologia o per cliente).
import type { ReactNode } from "react";
import { useLang } from "@/lib/i18n";
import { residuoAperto, type FatturaConStato } from "@/lib/fatture-logic";
import { chiaveFornitore } from "@/lib/flussi-logic";
import {
  chiaveCliente,
  scadenzaEffettiva,
  type TipoRigaCassa,
  type VoceDettaglioCassa,
} from "@/lib/cassa-logic";

const fmt = (n: number) =>
  n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const ddmmyy = (iso: string) =>
  iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(2, 4)}` : "—";
const LINK_CLS =
  "text-left underline decoration-dotted decoration-border underline-offset-4 hover:text-primary hover:decoration-primary";

/** Finestra sopra il foglio, come la scheda fornitore dei Flussi. */
function Finestra({
  titolo,
  sottotitolo,
  onChiudi,
  children,
}: {
  titolo: string;
  sottotitolo?: string;
  onChiudi: () => void;
  children: ReactNode;
}) {
  const { t } = useLang();
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onChiudi}
    >
      <div
        className="max-h-[88vh] w-full max-w-3xl overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold">{titolo}</span>
          {sottotitolo && (
            <span className="rounded-full bg-muted px-2 py-0.5 text-[11px]">{sottotitolo}</span>
          )}
          <button
            type="button"
            onClick={onChiudi}
            className="ml-auto rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted"
          >
            {t("common.close")}
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

/** Le fatture aperte del cliente: residuo (incassi Aruba/banca e note di
 *  credito già tolti), scadenza con la regola di fine mese, stato Aruba,
 *  giorni di ritardo. Stessi dati e stesse regole del foglio. */
export function SchedaClienteCassa({
  nome,
  nomi,
  attive,
  oggi,
  onChiudi,
}: {
  nome: string;
  /** Riga unita per gruppo: le società dentro (tutte le loro fatture). */
  nomi?: readonly string[];
  attive: readonly FatturaConStato[];
  oggi: string;
  onChiudi: () => void;
}) {
  const { t } = useLang();
  // Stessa chiave del foglio (nome in fattura) e, in più, la chiave canonica:
  // gli incassi in banca portano un nome un po' diverso da quello in fattura.
  const tutti = nomi && nomi.length ? nomi : [nome];
  const k1 = new Set(tutti.map(chiaveCliente));
  const k2 = new Set(tutti.map(chiaveFornitore));
  const righe = attive
    .filter((x) => k1.has(chiaveCliente(x.f.cliente)) || k2.has(chiaveFornitore(x.f.cliente)))
    .map((x) => ({
      x,
      residuo: residuoAperto(x),
      scad: x.s.scadenza ? scadenzaEffettiva(x.s.scadenza) : "",
    }))
    .filter((r) => r.residuo > 1)
    .sort((a, b) => a.scad.localeCompare(b.scad) || a.x.f.numero.localeCompare(b.x.f.numero, "it"));
  const totale = righe.reduce((s, r) => s + r.residuo, 0);
  const scaduto = righe.filter((r) => r.scad && r.scad <= oggi).reduce((s, r) => s + r.residuo, 0);
  const giorni = (scad: string) =>
    Math.round(
      (new Date(`${oggi}T00:00:00`).getTime() - new Date(`${scad}T00:00:00`).getTime()) / 86400000,
    );
  const th = "py-1 pr-3 text-left text-[11px] uppercase text-muted-foreground";
  const tdN = "py-1 pr-3 text-right tabular-nums whitespace-nowrap";
  return (
    <Finestra
      titolo={nome}
      sottotitolo={`${t("cpr.schedaCliente")}: ${righe.length}`}
      onChiudi={onChiudi}
    >
      {righe.length === 0 ? (
        <p className="text-sm text-muted-foreground">{t("cpr.schedaClienteVuota")}</p>
      ) : (
        <table className="w-full text-[13px]">
          <thead>
            <tr className="border-b border-border">
              <th className={th}>{t("for.colNumero")}</th>
              <th className={th}>{t("fis.colData")}</th>
              <th className={th}>{t("for.colScad")}</th>
              <th className={`${th} text-right`}>{t("fis.colImporto")}</th>
              <th className={`${th} text-right`}>{t("for.colResiduo")}</th>
              <th className={th}>{t("cpr.colAruba")}</th>
            </tr>
          </thead>
          <tbody>
            {righe.map((r) => {
              const ritardo = r.scad && r.scad <= oggi ? giorni(r.scad) : 0;
              return (
                <tr
                  key={`${r.x.f.numero}|${r.x.f.dataDocumento}`}
                  className="border-b border-border/40"
                >
                  <td className="py-1 pr-3 font-medium">{r.x.f.numero}</td>
                  <td className="py-1 pr-3 whitespace-nowrap">{ddmmyy(r.x.f.dataDocumento)}</td>
                  <td
                    className={`py-1 pr-3 whitespace-nowrap ${ritardo > 0 ? "text-status-absent" : ""}`}
                  >
                    {ddmmyy(r.scad)}
                    {ritardo > 0 && (
                      <span className="ml-1 text-[11px]">
                        ({ritardo} gg {t("cpr.scaduto")})
                      </span>
                    )}
                  </td>
                  <td className={tdN}>{fmt(r.x.f.totale)}</td>
                  <td className={`${tdN} font-medium`}>{fmt(r.residuo)}</td>
                  <td className="py-1 pr-3 text-[12px] text-muted-foreground">
                    {r.x.s.aruba || "—"}
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="border-t border-border font-semibold">
              <td className="py-1.5 pr-3" colSpan={4}>
                {t("cpr.apertoTot")}
                {scaduto > 0 && (
                  <span className="ml-2 font-normal text-status-absent">
                    {t("cpr.diCuiScaduto")} {fmt(scaduto)}
                  </span>
                )}
              </td>
              <td className={tdN}>{fmt(totale)}</td>
              <td />
            </tr>
          </tfoot>
        </table>
      )}
      <p className="mt-3 text-[11px] text-muted-foreground">{t("cpr.schedaClienteNota")}</p>
    </Finestra>
  );
}

/** Le scadenze dentro una riga fiscale (Fiscale / IVA / Banca): voce,
 *  scadenza, modalità, importo — Simone 04/11/2026 ("cliccando sul Fiscale
 *  voglio che si apra il popup con i dettagli"). */
export function SchedaFiscaleCassa({
  nome,
  voci,
  importo,
  onChiudi,
}: {
  nome: string;
  voci: readonly VoceDettaglioCassa[];
  importo: number;
  onChiudi: () => void;
}) {
  const { t } = useLang();
  const righe = [...voci].sort(
    (a, b) => a.data.localeCompare(b.data) || a.nome.localeCompare(b.nome, "it"),
  );
  const somma = righe.reduce((s, v) => s + v.importo, 0);
  const th = "py-1 pr-3 text-left text-[11px] uppercase text-muted-foreground";
  return (
    <Finestra
      titolo={nome}
      sottotitolo={`${t("cpr.schedaFiscale")}: ${righe.length}`}
      onChiudi={onChiudi}
    >
      <table className="w-full text-[13px]">
        <thead>
          <tr className="border-b border-border">
            <th className={th}>{t("cpr.colScadenza")}</th>
            <th className={th}>{t("cpr.colVoce")}</th>
            <th className={th}>{t("cpr.colModalita")}</th>
            <th className={`${th} text-right`}>{t("fis.colImporto")}</th>
          </tr>
        </thead>
        <tbody>
          {righe.map((v, i) => (
            <tr key={`${v.data}|${v.nome}|${i}`} className="border-b border-border/40">
              <td className="py-1 pr-3 whitespace-nowrap">{ddmmyy(v.data)}</td>
              <td className="py-1 pr-3">{v.nome}</td>
              <td className="py-1 pr-3 text-[12px] text-muted-foreground">{v.extra ?? ""}</td>
              <td className="py-1 text-right tabular-nums whitespace-nowrap">{fmt(v.importo)}</td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr className="border-t border-border font-semibold">
            <td className="py-1.5 pr-3" colSpan={3}>
              {t("cpr.apertoTot")}
              {Math.abs(somma - importo) > 0.005 && (
                <span className="ml-2 font-normal text-muted-foreground">
                  ({t("cpr.importoOrigine")} {fmt(somma)})
                </span>
              )}
            </td>
            <td className="py-1.5 text-right tabular-nums">{fmt(importo)}</td>
          </tr>
        </tfoot>
      </table>
    </Finestra>
  );
}

/** I fornitori dentro un gruppo (uscite per tipologia o per cliente): un
 *  clic su un fornitore chiude l'elenco e apre la sua scheda dei Flussi. */
export function SchedaGruppoCassa({
  gruppo,
  onApriFornitore,
  onChiudi,
}: {
  gruppo: {
    nome: string;
    importo: number;
    membri?: readonly { nome: string; importo: number; tipo: TipoRigaCassa }[];
  };
  onApriFornitore?: (nome: string) => void;
  onChiudi: () => void;
}) {
  const { t } = useLang();
  const membri = [...(gruppo.membri ?? [])].sort(
    (a, b) => b.importo - a.importo || a.nome.localeCompare(b.nome, "it"),
  );
  return (
    <Finestra
      titolo={gruppo.nome}
      sottotitolo={`${t("cpr.schedaGruppo")}: ${membri.length}`}
      onChiudi={onChiudi}
    >
      <table className="w-full text-[13px]">
        <tbody>
          {membri.map((m, i) => (
            <tr key={`${m.nome}|${i}`} className="border-b border-border/40">
              <td className="py-1 pr-3">
                {onApriFornitore && m.tipo === "passiva" ? (
                  <button
                    type="button"
                    onClick={() => {
                      onChiudi();
                      onApriFornitore(m.nome);
                    }}
                    title={t("cpr.apriFornitore")}
                    className={LINK_CLS}
                  >
                    {m.nome}
                  </button>
                ) : (
                  m.nome
                )}
              </td>
              <td className="py-1 text-right tabular-nums whitespace-nowrap">{fmt(m.importo)}</td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr className="border-t border-border font-semibold">
            <td className="py-1.5 pr-3">{t("cpr.apertoTot")}</td>
            <td className="py-1.5 text-right tabular-nums">{fmt(gruppo.importo)}</td>
          </tr>
        </tfoot>
      </table>
    </Finestra>
  );
}
