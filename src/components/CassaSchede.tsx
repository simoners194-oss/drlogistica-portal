// DR Portal — Finanza → Flussi → Cassa proprietà: le schede che si aprono
// con un clic sui nomi del foglio (Simone 04/11/2026: "me li link come su
// vista classica"). I fornitori aprono la scheda fornitore dei Flussi
// (SchedaFornitore, dal tab); qui vivono la scheda del CLIENTE (le sue
// fatture aperte, dalle attive) e l'elenco dei fornitori dentro un gruppo
// (uscite raggruppate per tipologia o per cliente).
import { useState, type ReactNode } from "react";
import { useLang } from "@/lib/i18n";
import { residuoAperto, type FatturaConStato } from "@/lib/fatture-logic";
import { chiaveFornitore } from "@/lib/flussi-logic";
import {
  chiaveCliente,
  scadenzaEffettiva,
  type SceltaFattura,
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
  larga,
  onChiudi,
  children,
}: {
  titolo: string;
  sottotitolo?: string;
  /** Finestra più larga (la scheda del cliente con la scelta delle fatture). */
  larga?: boolean;
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
        className={`max-h-[88vh] w-full ${larga ? "max-w-4xl" : "max-w-3xl"} overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-xl`}
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

/** Dove può stare una fattura attesa: tra le entranti, o in uno dei due
 *  standby delle entrate. */
type DoveFattura = SceltaFattura | "";
const DOVE: readonly {
  v: DoveFattura;
  chiave: "cpr.entranti" | "cpr.colStandby" | "cpr.colStandby2";
}[] = [
  { v: "", chiave: "cpr.entranti" },
  { v: "standby", chiave: "cpr.colStandby" },
  { v: "standby2", chiave: "cpr.colStandby2" },
];
const DOVE_ACCESO: Record<DoveFattura, string> = {
  "": "border-emerald-600 bg-emerald-600 text-white",
  standby: "border-slate-500 bg-slate-500 text-white",
  standby2: "border-slate-700 bg-slate-700 text-white",
};
const DOVE_RIQUADRO: Record<DoveFattura, string> = {
  "": "bg-emerald-100/80 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300",
  standby: "bg-slate-200/80 text-slate-700 dark:bg-slate-700/60 dark:text-slate-200",
  standby2: "bg-slate-300/80 text-slate-800 dark:bg-slate-600/70 dark:text-slate-100",
};

/** Le fatture aperte del cliente: residuo (incassi Aruba/banca e note di
 *  credito già tolti), scadenza con la regola di fine mese, stato Aruba,
 *  giorni di ritardo. Stessi dati e stesse regole del foglio.
 *  Dalla 1.122.0 (Simone 19/11/2026: "devo poter selezionare per cliente
 *  quali fatture mettere in stand by e quali lasciare come entranti … solo
 *  alcune fatture per cliente") qui si sceglie, fattura per fattura, dove va
 *  ognuna: Entranti, Standby (da incassare più avanti) o Standby 2 (non le
 *  incasseremo). Una alla volta coi tre pulsanti della riga, oppure più
 *  insieme con le spunte. */
export function SchedaClienteCassa({
  nome,
  nomi,
  attive,
  oggi,
  limite,
  dove,
  onScelta,
  accordo,
  perAcconto,
  onTogliAcconto,
  onChiudi,
}: {
  nome: string;
  /** Riga unita per gruppo: le società dentro (tutte le loro fatture). */
  nomi?: readonly string[];
  attive: readonly FatturaConStato[];
  oggi: string;
  /** Fin dove guarda il foglio: le fatture che scadono dopo non ci sono ancora. */
  limite?: string;
  /** Dove sta ogni fattura adesso (nomeFile): "" = tra le entranti. */
  dove?: (nomeFile: string) => DoveFattura;
  /** Sposta una o più fatture; assente = sola lettura. */
  onScelta?: (file: string[], dove: DoveFattura) => void;
  /** L'acconto scritto in "Entra ora" per questo cliente, se c'è. */
  accordo?: number;
  /** Vero per le fatture che nel foglio stanno nello Standby perché la
   *  società ha un acconto: sono il resto che l'acconto non copre. */
  perAcconto?: (nomeFile: string) => boolean;
  /** Toglie l'acconto della società (assente = sola lettura). */
  onTogliAcconto?: () => void;
  onChiudi: () => void;
}) {
  const { t } = useLang();
  const [spuntate, setSpuntate] = useState<ReadonlySet<string>>(new Set());
  // Stessa chiave del foglio (nome in fattura) e, in più, la chiave canonica:
  // gli incassi in banca portano un nome un po' diverso da quello in fattura.
  const tutti = nomi && nomi.length ? nomi : [nome];
  const k1 = new Set(tutti.map(chiaveCliente));
  const k2 = new Set(tutti.map(chiaveFornitore));
  const righe = attive
    .filter((x) => k1.has(chiaveCliente(x.f.cliente)) || k2.has(chiaveFornitore(x.f.cliente)))
    .map((x) => ({
      x,
      file: x.f.nomeFile,
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
  // Dove sta la fattura NEL FOGLIO: la sua scelta, oppure lo Standby quando
  // è il resto dell'acconto della società (1.124.2, Simone 19/11/2026: "sono
  // in stand by ma sono segnati come entranti, ma non si spostano").
  const daAcconto = (file: string) => Boolean(perAcconto?.(file)) && !dove?.(file);
  const doveDi = (file: string): DoveFattura =>
    daAcconto(file) ? "standby" : (dove?.(file) ?? "");
  const fuoriFoglio = (scad: string) => Boolean(limite && scad && scad > limite);
  // Quanto c'è in ogni blocco (solo le fatture che il foglio guarda).
  const perBlocco = (v: DoveFattura) =>
    righe.filter((r) => !fuoriFoglio(r.scad) && doveDi(r.file) === v);
  const sommaDi = (l: typeof righe) => l.reduce((s, r) => s + r.residuo, 0);
  const selezionate = righe.filter((r) => spuntate.has(r.file));
  const tutteSpuntate = righe.length > 0 && selezionate.length === righe.length;
  const spunta = (file: string) =>
    setSpuntate((prima) => {
      const next = new Set(prima);
      if (next.has(file)) next.delete(file);
      else next.add(file);
      return next;
    });
  const sposta = (file: string[], v: DoveFattura) => {
    if (!onScelta || !file.length) return;
    onScelta(file, v);
    setSpuntate(new Set());
  };
  const th = "py-1 pr-3 text-left text-[11px] uppercase text-muted-foreground";
  const tdN = "py-1 pr-3 text-right tabular-nums whitespace-nowrap";
  const scelta = Boolean(dove);
  return (
    <Finestra
      titolo={nome}
      sottotitolo={`${t("cpr.schedaCliente")}: ${righe.length}`}
      larga={scelta}
      onChiudi={onChiudi}
    >
      {righe.length === 0 ? (
        <p className="text-sm text-muted-foreground">{t("cpr.schedaClienteVuota")}</p>
      ) : (
        <>
          {/* L'acconto della società comanda sulle scelte: va detto in cima,
              col pulsante per toglierlo. */}
          {scelta && accordo != null && (
            <div className="mb-3 flex flex-wrap items-center gap-3 rounded-xl border-2 border-amber-400 bg-amber-50 px-3 py-2 text-[12.5px] text-amber-900 dark:border-amber-600 dark:bg-amber-950/40 dark:text-amber-100">
              <span className="min-w-0 flex-1">
                <b>
                  {t("cpr.sceltaAccontoTitolo")} {fmt(accordo)}.
                </b>{" "}
                {t("cpr.sceltaAccontoTesto")}
              </span>
              {onTogliAcconto && (
                <button
                  type="button"
                  onClick={onTogliAcconto}
                  className="shrink-0 rounded-lg border border-amber-500 bg-background px-3 py-1 text-[12.5px] font-semibold text-foreground hover:bg-amber-100 dark:hover:bg-amber-900/40"
                >
                  {t("cpr.sceltaTogliAcconto")}
                </button>
              )}
            </div>
          )}
          {scelta && (
            <div className="mb-3 rounded-xl border border-border bg-muted/30 p-3">
              <div className="flex flex-wrap items-stretch gap-2">
                {DOVE.map((d) => {
                  const l = perBlocco(d.v);
                  return (
                    <div
                      key={d.v || "entranti"}
                      className={`flex min-w-[130px] flex-col rounded-lg px-3 py-1.5 leading-tight ${DOVE_RIQUADRO[d.v]}`}
                    >
                      <span className="text-[10px] font-semibold uppercase tracking-wider opacity-80">
                        {t(d.chiave)} · {l.length} {t("cpr.fattureBreve")}
                      </span>
                      <span className="text-[15px] font-bold tabular-nums">{fmt(sommaDi(l))}</span>
                    </div>
                  );
                })}
              </div>
              <p className="mt-2 text-[11px] text-muted-foreground">{t("cpr.sceltaGuida")}</p>
              {onScelta && (
                <div className="mt-2 flex flex-wrap items-center gap-2 text-[12px]">
                  <span className="text-muted-foreground">
                    {selezionate.length
                      ? `${selezionate.length} ${t("cpr.sceltaSelezionate")} · ${fmt(sommaDi(selezionate))}`
                      : t("cpr.sceltaNessuna")}
                  </span>
                  {DOVE.map((d) => (
                    <button
                      key={d.v || "entranti"}
                      type="button"
                      disabled={!selezionate.length}
                      onClick={() =>
                        sposta(
                          selezionate.map((r) => r.file),
                          d.v,
                        )
                      }
                      className="rounded-lg border border-border bg-background px-2.5 py-1 font-medium hover:bg-muted disabled:opacity-40"
                    >
                      → {t(d.chiave)}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
          <table className="w-full text-[13px]">
            <thead>
              <tr className="border-b border-border">
                {scelta && onScelta && (
                  <th className="w-7 py-1 pr-2">
                    <input
                      type="checkbox"
                      checked={tutteSpuntate}
                      onChange={() =>
                        setSpuntate(tutteSpuntate ? new Set() : new Set(righe.map((r) => r.file)))
                      }
                      title={t("cpr.sceltaTutte")}
                    />
                  </th>
                )}
                <th className={th}>{t("for.colNumero")}</th>
                <th className={th}>{t("cpr.sceltaData")}</th>
                <th className={th}>{t("for.colScad")}</th>
                <th className={`${th} text-right`}>{t("for.colResiduo")}</th>
                <th className={th}>{t("cpr.colAruba")}</th>
                {scelta && <th className={th}>{t("cpr.sceltaDove")}</th>}
              </tr>
            </thead>
            <tbody>
              {righe.map((r) => {
                const ritardo = r.scad && r.scad <= oggi ? giorni(r.scad) : 0;
                const qui = doveDi(r.file);
                const fuori = fuoriFoglio(r.scad);
                const perAcc = daAcconto(r.file);
                return (
                  <tr
                    key={r.file}
                    className={`border-b border-border/40 ${spuntate.has(r.file) ? "bg-primary/5" : ""}`}
                  >
                    {scelta && onScelta && (
                      <td className="py-1 pr-2">
                        <input
                          type="checkbox"
                          checked={spuntate.has(r.file)}
                          onChange={() => spunta(r.file)}
                        />
                      </td>
                    )}
                    <td className="py-1 pr-3 font-medium">
                      {r.x.f.numero}
                      {Math.abs(r.x.f.totale - r.residuo) > 0.005 && (
                        <span className="ml-1 text-[11px] font-normal text-muted-foreground">
                          ({t("cpr.sceltaSu")} {fmt(r.x.f.totale)})
                        </span>
                      )}
                    </td>
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
                    <td className={`${tdN} font-medium`}>{fmt(r.residuo)}</td>
                    <td className="py-1 pr-3 text-[12px] text-muted-foreground">
                      {r.x.s.aruba || "—"}
                    </td>
                    {scelta && (
                      <td className="py-1">
                        <span className="inline-flex items-center gap-1">
                          {DOVE.map((d) => (
                            <button
                              key={d.v || "entranti"}
                              type="button"
                              // Finché c'è l'acconto entra solo quello: una
                              // sua fattura non si riporta tra le entranti da qui.
                              disabled={!onScelta || (perAcc && d.v === "")}
                              title={
                                perAcc && d.v === "" ? t("cpr.sceltaPerAccontoTip") : undefined
                              }
                              onClick={() => qui !== d.v && sposta([r.file], d.v)}
                              className={`rounded border px-1.5 py-0.5 text-[10px] font-bold uppercase ${
                                qui === d.v
                                  ? DOVE_ACCESO[d.v]
                                  : "border-border text-muted-foreground hover:bg-muted"
                              } ${onScelta ? "" : "cursor-default"}`}
                            >
                              {t(d.chiave)}
                            </button>
                          ))}
                          {perAcc && (
                            <span
                              className="rounded border border-amber-400 px-1 text-[10px] font-semibold text-amber-800 dark:text-amber-200"
                              title={t("cpr.sceltaPerAccontoTip")}
                            >
                              {t("cpr.sceltaPerAcconto")}
                            </span>
                          )}
                          {fuori && (
                            <span
                              className="text-[10.5px] italic text-muted-foreground"
                              title={t("cpr.sceltaFuoriTip")}
                            >
                              {t("cpr.sceltaFuori")}
                            </span>
                          )}
                        </span>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="border-t border-border font-semibold">
                <td className="py-1.5 pr-3" colSpan={scelta && onScelta ? 4 : 3}>
                  {t("cpr.apertoTot")}
                  {scaduto > 0 && (
                    <span className="ml-2 font-normal text-status-absent">
                      {t("cpr.diCuiScaduto")} {fmt(scaduto)}
                    </span>
                  )}
                </td>
                <td className={tdN}>{fmt(totale)}</td>
                <td colSpan={scelta ? 2 : 1} />
              </tr>
            </tfoot>
          </table>
        </>
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
  fonte,
  onChiudi,
}: {
  nome: string;
  voci: readonly VoceDettaglioCassa[];
  importo: number;
  /** Da quale file viene lo scadenzario (1.120.0). */
  fonte?: string;
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
      {fonte && (
        <p className="mt-3 text-[11px] text-muted-foreground">
          {t("cpr.fiscaleFonte")} {fonte}
        </p>
      )}
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
