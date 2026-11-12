// DR Portal — Scheda fornitore (Simone 29/09/2026, v1.87.0): il popup che
// si apre cliccando un fornitore nei Flussi di cassa, con tutte le
// informazioni della tab Fornitori (modalità, chi se ne occupa, macrovoce,
// appalto, importi, conteggi, date) e le fatture aperte nel flusso. Le
// quattro caselle si modificano qui come nella tab e si salvano col
// bottone. Qui vive anche `SelectVocab`, la tendina a vocabolario compatta
// usata dalla tab Fornitori (elenco + "Altro…" per scrivere).
import { useEffect, useState } from "react";
import { Loader2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import {
  REFERENTI_DR,
  type FornitoreInfo,
  type ModalitaPagamento,
  type QuotaAppalto,
  type RigaFornitore,
  type VocabolarioFornitori,
} from "@/lib/flussi-logic";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
const dataIt = (iso: string) =>
  /^\d{4}-\d{2}-\d{2}/.test(iso) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : "";

const BADGE: Record<ModalitaPagamento, string> = { rid: "RID", riba: "RiBa", altro: "—" };

export const inputVocabCls =
  "w-full min-w-24 rounded border border-border bg-background px-1.5 py-0.5 text-[12px] text-foreground placeholder:italic placeholder:text-muted-foreground/70";

/** Tendina a vocabolario: le voci dell'elenco più "Altro…" che apre il
 *  campo libero. Un valore fuori elenco apre il campo libero da solo. */
export function SelectVocab({
  value,
  opzioni,
  placeholder,
  disabled,
  onSelect,
  onTesto,
  onBlur,
  className,
}: {
  value: string;
  opzioni: readonly string[];
  placeholder?: string;
  disabled?: boolean;
  /** Scelta dall'elenco (o svuotamento): si può salvare subito. */
  onSelect: (v: string) => void;
  /** Testo libero mentre si scrive. */
  onTesto: (v: string) => void;
  /** Uscita dal campo libero. */
  onBlur?: () => void;
  className?: string;
}) {
  const { t } = useLang();
  const [libero, setLibero] = useState(false);
  const fuoriElenco = value !== "" && !opzioni.includes(value);
  if (libero || fuoriElenco)
    return (
      <div className="flex gap-1">
        <input
          value={value}
          autoFocus={libero}
          disabled={disabled}
          placeholder={placeholder}
          onChange={(e) => onTesto(e.target.value)}
          onBlur={() => {
            if (!value.trim()) setLibero(false);
            onBlur?.();
          }}
          onKeyDown={(e) => e.key === "Enter" && (e.target as HTMLInputElement).blur()}
          className={className ?? inputVocabCls}
        />
        <button
          type="button"
          disabled={disabled}
          title={t("for.tornaElenco")}
          onClick={() => {
            setLibero(false);
            onSelect("");
          }}
          className="shrink-0 rounded border border-border px-1.5 text-[11px] hover:bg-muted"
        >
          ↩
        </button>
      </div>
    );
  return (
    <select
      value={value}
      disabled={disabled}
      title={placeholder}
      onChange={(e) => {
        if (e.target.value === "__altro__") {
          setLibero(true);
          onTesto("");
        } else onSelect(e.target.value);
      }}
      className={className ?? inputVocabCls}
    >
      <option value="">{placeholder ? `(${placeholder})` : "—"}</option>
      {opzioni.map((o) => (
        <option key={o} value={o}>
          {o}
        </option>
      ))}
      <option value="__altro__">{t("for.altro")}</option>
    </select>
  );
}

export interface FatturaScheda {
  numero: string;
  data: string;
  scadenza: string;
  residuo: number;
  inRitardo: boolean;
  modalita: ModalitaPagamento;
  /** Causale scritta in fattura (Simone 04/11/2026). */
  oggetto?: string;
  /** Chiave della fattura (nome file SdI): serve al flag della carta. */
  nomeFile?: string;
  /** Data in cui è stata pagata con la carta (flag acceso). */
  pagataCdc?: string;
}

export function SchedaFornitore({
  riga,
  fatture,
  vocab,
  busy,
  onSalva,
  onChiudi,
  quoteMezzi,
  onPagataCdc,
  cdcBusy,
}: {
  riga: RigaFornitore;
  fatture: FatturaScheda[];
  vocab: VocabolarioFornitori;
  busy: boolean;
  onSalva: (info: FornitoreInfo) => Promise<void>;
  onChiudi: () => void;
  /** Quote proposte dal modulo Mezzi (contratti di noleggio / parco), se ci sono. */
  quoteMezzi?: QuotaAppalto[];
  /** Flag "pagata con carta" su una fattura aperta (Simone 12/11/2026): la
   *  chiude nel portale in attesa che venga chiusa su Aruba. */
  onPagataCdc?: (nomeFile: string, on: boolean) => void;
  cdcBusy?: boolean;
}) {
  const { t } = useLang();
  const [modalita, setModalita] = useState<ModalitaPagamento | "">(riga.override ?? "");
  const [referente, setReferente] = useState(riga.referente);
  const [macrovoce, setMacrovoce] = useState(riga.macrovoce);
  const [appalto, setAppalto] = useState(riga.appalto);
  const [promemoria, setPromemoria] = useState(riga.promemoria);
  const [compensa, setCompensa] = useState(riga.compensa);
  const [stimaGiorni, setStimaGiorni] = useState(String(riga.stimaGiorni || ""));
  // Ripartizione tra più appalti (Simone 09/11/2026): quote a mano oppure
  // "bilanciate dal sistema" dai mezzi.
  const [ripartizione, setRipartizione] = useState<QuotaAppalto[]>(riga.ripartizione);
  const [ripAuto, setRipAuto] = useState(riga.ripartizioneAuto);
  const sommaQuote = ripartizione.reduce((s, q) => s + (Number(q.quota) || 0), 0);
  const aggQuota = (i: number, patch: Partial<QuotaAppalto>) =>
    setRipartizione(ripartizione.map((q, j) => (j === i ? { ...q, ...patch } : q)));
  // Riallineo le caselle quando cambia il fornitore (o dopo un salvataggio).
  useEffect(() => {
    setModalita(riga.override ?? "");
    setReferente(riga.referente);
    setMacrovoce(riga.macrovoce);
    setAppalto(riga.appalto);
    setPromemoria(riga.promemoria);
    setCompensa(riga.compensa);
    setStimaGiorni(String(riga.stimaGiorni || ""));
    setRipartizione(riga.ripartizione);
    setRipAuto(riga.ripartizioneAuto);
  }, [
    riga.chiave,
    riga.override,
    riga.referente,
    riga.macrovoce,
    riga.appalto,
    riga.promemoria,
    riga.compensa,
    riga.stimaGiorni,
    riga.ripartizione,
    riga.ripartizioneAuto,
  ]);
  const cambiato =
    (modalita || undefined) !== riga.override ||
    referente.trim() !== riga.referente ||
    macrovoce.trim() !== riga.macrovoce ||
    appalto.trim() !== riga.appalto ||
    promemoria.trim() !== riga.promemoria ||
    compensa !== riga.compensa ||
    (Number(stimaGiorni) || 0) !== (riga.stimaGiorni || 0) ||
    ripAuto !== riga.ripartizioneAuto ||
    JSON.stringify(ripartizione) !== JSON.stringify(riga.ripartizione);

  const lbl = "text-[11px] uppercase text-muted-foreground";
  const stat = (etichetta: string, valore: string, cls = "") => (
    <div>
      <div className={lbl}>{etichetta}</div>
      <div className={`text-[13px] tabular-nums ${cls}`}>{valore}</div>
    </div>
  );

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onChiudi}
    >
      <div
        className="max-h-[88vh] w-full max-w-5xl overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold">{riga.nome}</span>
          <span className="rounded-full bg-muted px-2 py-0.5 text-[11px]">
            {t("for.colModalita")}: {BADGE[riga.modalita]}
            {riga.override ? "" : ` (${t("for.modAuto").toLowerCase()})`}
            {riga.mista ? " ±" : ""}
          </span>
          {busy && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
          <button
            type="button"
            onClick={onChiudi}
            className="ml-auto rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted"
          >
            {t("common.close")}
          </button>
        </div>

        {/* Caselle modificabili (stesse della tab Fornitori) */}
        <div className="mb-4 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <div>
            <div className={lbl}>{t("for.colModalita")}</div>
            <select
              value={modalita}
              disabled={busy}
              title={t("for.modAutoTip")}
              onChange={(e) => setModalita(e.target.value as ModalitaPagamento | "")}
              className={inputVocabCls}
            >
              <option value="">
                {t("for.modAuto")} ({BADGE[riga.modalitaAuto]}
                {riga.mista ? " ±" : ""})
              </option>
              <option value="rid">{t("for.modRid")}</option>
              <option value="riba">{t("for.modRiba")}</option>
              <option value="altro">{t("for.modAltro")}</option>
            </select>
          </div>
          <div>
            <div className={lbl}>{t("for.colReferente")}</div>
            <SelectVocab
              value={referente}
              opzioni={REFERENTI_DR}
              placeholder={t("for.refPh")}
              disabled={busy}
              onSelect={setReferente}
              onTesto={setReferente}
            />
          </div>
          <div>
            <div className={lbl}>{t("for.colMacro")}</div>
            <SelectVocab
              value={macrovoce}
              opzioni={vocab.macrovoci}
              placeholder={riga.macrovoceAuto || t("for.macroPh")}
              disabled={busy}
              onSelect={setMacrovoce}
              onTesto={setMacrovoce}
            />
          </div>
          <div>
            <div className={lbl}>{t("for.colAppalto")}</div>
            <SelectVocab
              value={appalto}
              opzioni={vocab.appalti}
              placeholder={riga.appaltoAuto || t("for.appaltoPh")}
              disabled={busy}
              onSelect={setAppalto}
              onTesto={setAppalto}
            />
          </div>
        </div>
        {/* Ripartizione tra più appalti (Simone 09/11/2026): per dividere le
            spese del fornitore (noleggi, carburante, pedaggi, multe) tra gli
            appalti, a mano in percentuale o bilanciate dai mezzi. */}
        <div className="mb-3 rounded-xl border border-border/60 p-3">
          <div className="mb-1 flex flex-wrap items-center gap-3">
            <span className={lbl}>{t("for.ripTitolo")}</span>
            <label className="flex items-center gap-1.5 text-[12px]" title={t("for.ripAutoTip")}>
              <input
                type="checkbox"
                checked={ripAuto}
                disabled={busy || !(quoteMezzi && quoteMezzi.length)}
                onChange={(e) => {
                  setRipAuto(e.target.checked);
                  if (e.target.checked && quoteMezzi?.length) setRipartizione(quoteMezzi);
                }}
              />
              {t("for.ripAuto")}
              {!(quoteMezzi && quoteMezzi.length) && (
                <span className="text-[11px] text-muted-foreground">({t("for.ripNoMezzi")})</span>
              )}
            </label>
            {quoteMezzi && quoteMezzi.length > 0 && !ripAuto && (
              <button
                type="button"
                disabled={busy}
                onClick={() => setRipartizione(quoteMezzi)}
                className="text-[12px] text-primary underline-offset-2 hover:underline"
              >
                {t("for.ripDaMezzi")}
              </button>
            )}
            <span
              className={`ml-auto text-[12px] tabular-nums ${ripartizione.length && Math.abs(sommaQuote - 100) > 0.05 ? "text-status-absent" : "text-muted-foreground"}`}
            >
              {t("for.ripTotale")} {sommaQuote.toFixed(1)}%
            </span>
          </div>
          {ripartizione.map((q, i) => (
            <div key={i} className="mb-1 flex items-center gap-2">
              <div className="flex-1">
                <SelectVocab
                  value={q.appalto}
                  opzioni={vocab.appalti}
                  placeholder={t("for.appaltoPh")}
                  disabled={busy || ripAuto}
                  onSelect={(v) => aggQuota(i, { appalto: v })}
                  onTesto={(v) => aggQuota(i, { appalto: v })}
                />
              </div>
              <input
                type="number"
                min={0}
                max={100}
                step={0.1}
                value={q.quota}
                disabled={busy || ripAuto}
                onChange={(e) => aggQuota(i, { quota: Number(e.target.value) || 0 })}
                className={`${inputVocabCls} w-20 text-right`}
              />
              <span className="text-[12px] text-muted-foreground">%</span>
              {!ripAuto && (
                <button
                  type="button"
                  disabled={busy}
                  onClick={() => setRipartizione(ripartizione.filter((_, j) => j !== i))}
                  title={t("common.delete")}
                  className="text-[12px] text-muted-foreground hover:text-status-absent"
                >
                  ×
                </button>
              )}
            </div>
          ))}
          {!ripAuto && (
            <button
              type="button"
              disabled={busy}
              onClick={() => setRipartizione([...ripartizione, { appalto: "", quota: 0 }])}
              className="text-[12px] text-primary underline-offset-2 hover:underline"
            >
              {t("for.ripAggiungi")}
            </button>
          )}
          <p className="mt-1 text-[11px] text-muted-foreground">{t("for.ripNota")}</p>
        </div>
        {/* Promemoria libero (Simone 04/11/2026): cosa fa davvero il
            fornitore e le regole da ricordare quando si paga. */}
        <div className="mb-3">
          <div className={lbl}>{t("for.promemoria")}</div>
          <textarea
            value={promemoria}
            disabled={busy}
            rows={3}
            maxLength={2000}
            placeholder={t("for.promemoriaPh")}
            onChange={(e) => setPromemoria(e.target.value)}
            className={`${inputVocabCls} min-h-16 resize-y`}
          />
        </div>
        {/* Paga in compensazione (Simone 04/11/2026): in Cassa proprietà le
            sue fatture si scalano dagli entranti, non stanno tra le uscite. */}
        <label className="mb-3 flex items-center gap-2 text-[13px]" title={t("for.compensaTip")}>
          <input
            type="checkbox"
            checked={compensa}
            disabled={busy}
            onChange={(e) => setCompensa(e.target.checked)}
          />
          {t("for.compensa")}
          <span className="text-[11px] text-muted-foreground">{t("for.compensaTip")}</span>
        </label>
        {/* Fatture ricorrenti (Simone 04/11/2026: carburanti ogni 15 giorni):
            la Cassa stima le prossime sulla media delle ultime sei. */}
        <label
          className="mb-3 flex flex-wrap items-center gap-2 text-[13px]"
          title={t("for.stimaTip")}
        >
          {t("for.stimaGiorni")}
          <input
            type="number"
            min={0}
            max={366}
            value={stimaGiorni}
            disabled={busy}
            onChange={(e) => setStimaGiorni(e.target.value)}
            placeholder="0"
            className={`${inputVocabCls} w-20 text-right`}
          />
          <span className="text-[11px] text-muted-foreground">{t("for.stimaTip")}</span>
        </label>
        <div className="mb-4 flex items-center gap-2">
          <button
            type="button"
            disabled={busy || !cambiato}
            onClick={() =>
              void onSalva({
                modalita: modalita || undefined,
                referente: referente.trim(),
                macrovoce: macrovoce.trim(),
                appalto: appalto.trim(),
                promemoria: promemoria.trim(),
                compensa,
                stimaGiorni: Number(stimaGiorni) || undefined,
                ripartizione: ripartizione.filter((q) => q.appalto.trim() && q.quota > 0),
                ripartizioneAuto: ripAuto,
              })
            }
            className="rounded-lg bg-primary px-3 py-1 text-[13px] font-medium text-primary-foreground disabled:opacity-40"
          >
            {t("common.save")}
          </button>
          <span className="text-[11px] text-muted-foreground">{t("for.schedaNota")}</span>
        </div>

        {/* Numeri dalle fatture */}
        <div className="mb-4 grid grid-cols-2 gap-3 rounded-xl border border-border/60 p-3 sm:grid-cols-3 lg:grid-cols-5">
          {stat(
            t("for.colScaduto"),
            riga.scaduto > 0.005 ? fmtImporto(riga.scaduto) : "—",
            riga.scaduto > 0.005 ? "font-medium text-status-absent" : "",
          )}
          {stat(t("for.colAperto"), fmtImporto(riga.aperto), "font-medium")}
          {stat(t("for.colNScad"), String(riga.nScadute))}
          {stat(t("for.colNAperte"), String(riga.nAperte))}
          {stat(t("for.colNTot"), String(riga.nTotale))}
          {stat(t("for.colPrima"), dataIt(riga.primaFattura) || "—")}
          {stat(
            t("for.colProssima"),
            dataIt(riga.prossimaScadenza) || "—",
            riga.prossimaScadenza && riga.nScadute > 0 ? "text-status-absent" : "",
          )}
          {stat(t("for.colUltima"), dataIt(riga.ultimaFattura) || "—")}
          {stat(t("for.colNPagate"), String(riga.nPagate))}
          {stat(t("for.colPagato"), Math.abs(riga.pagato) > 0.005 ? fmtImporto(riga.pagato) : "—")}
        </div>

        {/* Fatture aperte nel flusso */}
        <p className="mb-1 text-xs font-medium">
          {t("for.fattureAperte")} ({fatture.length})
        </p>
        <div className="max-h-96 overflow-y-auto rounded-lg border border-border/60">
          <table className="w-full text-xs">
            <thead>
              <tr className="border-b border-border text-left text-[11px] uppercase text-muted-foreground">
                <th className="px-2 py-1">{t("for.colNumero")}</th>
                <th className="px-2 py-1">{t("fis.colData")}</th>
                <th className="px-2 py-1">{t("for.colScad")}</th>
                <th className="px-2 py-1">{t("for.colCausale")}</th>
                <th className="px-2 py-1 text-right">{t("for.colResiduo")}</th>
                <th className="px-2 py-1 text-center">{t("for.colModalita")}</th>
                {onPagataCdc && (
                  <th className="px-2 py-1 text-center" title={t("for.cdcTip")}>
                    {t("for.colCdc")}
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {fatture.map((f, i) => (
                <tr
                  key={`${f.numero}|${i}`}
                  className={`border-b border-border/30 ${f.pagataCdc ? "bg-violet-50/70 dark:bg-violet-950/20" : ""}`}
                >
                  <td className="whitespace-nowrap px-2 py-0.5">{f.numero}</td>
                  <td className="whitespace-nowrap px-2 py-0.5 text-muted-foreground">
                    {dataIt(f.data)}
                  </td>
                  <td
                    className={`whitespace-nowrap px-2 py-0.5 ${f.inRitardo ? "text-status-absent" : ""}`}
                  >
                    {dataIt(f.scadenza)}
                  </td>
                  {/* La causale si legge TUTTA, andando a capo (Simone
                      12/11/2026: "viene tagliata, la voglio vedere tutta"). */}
                  <td className="min-w-64 whitespace-normal break-words px-2 py-0.5 text-muted-foreground">
                    {f.oggetto ?? ""}
                  </td>
                  <td
                    className={`whitespace-nowrap px-2 py-0.5 text-right tabular-nums ${f.pagataCdc ? "text-muted-foreground line-through" : ""}`}
                  >
                    {fmtImporto(f.residuo)}
                  </td>
                  <td className="px-2 py-0.5 text-center">
                    <span className="rounded-full bg-muted px-1.5 text-[10px]">
                      {BADGE[f.modalita]}
                    </span>
                  </td>
                  {onPagataCdc && (
                    <td className="whitespace-nowrap px-2 py-0.5 text-center">
                      {f.nomeFile && (
                        <label
                          className="inline-flex cursor-pointer items-center gap-1 text-[11px]"
                          title={t("for.cdcTip")}
                        >
                          <input
                            type="checkbox"
                            checked={Boolean(f.pagataCdc)}
                            disabled={cdcBusy}
                            onChange={(e) => onPagataCdc(f.nomeFile!, e.target.checked)}
                          />
                          {f.pagataCdc && (
                            <span className="font-semibold text-violet-700 dark:text-violet-300">
                              {dataIt(f.pagataCdc)}
                            </span>
                          )}
                        </label>
                      )}
                    </td>
                  )}
                </tr>
              ))}
              {fatture.length === 0 && (
                <tr>
                  <td
                    colSpan={onPagataCdc ? 7 : 6}
                    className="py-3 text-center text-muted-foreground"
                  >
                    {t("for.nessunaFattura")}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
