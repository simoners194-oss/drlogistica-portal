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
}

export function SchedaFornitore({
  riga,
  fatture,
  vocab,
  busy,
  onSalva,
  onChiudi,
}: {
  riga: RigaFornitore;
  fatture: FatturaScheda[];
  vocab: VocabolarioFornitori;
  busy: boolean;
  onSalva: (info: FornitoreInfo) => Promise<void>;
  onChiudi: () => void;
}) {
  const { t } = useLang();
  const [modalita, setModalita] = useState<ModalitaPagamento | "">(riga.override ?? "");
  const [referente, setReferente] = useState(riga.referente);
  const [macrovoce, setMacrovoce] = useState(riga.macrovoce);
  const [appalto, setAppalto] = useState(riga.appalto);
  // Riallineo le caselle quando cambia il fornitore (o dopo un salvataggio).
  useEffect(() => {
    setModalita(riga.override ?? "");
    setReferente(riga.referente);
    setMacrovoce(riga.macrovoce);
    setAppalto(riga.appalto);
  }, [riga.chiave, riga.override, riga.referente, riga.macrovoce, riga.appalto]);
  const cambiato =
    (modalita || undefined) !== riga.override ||
    referente.trim() !== riga.referente ||
    macrovoce.trim() !== riga.macrovoce ||
    appalto.trim() !== riga.appalto;

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
        className="max-h-[88vh] w-full max-w-3xl overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-xl"
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
        <div className="max-h-64 overflow-y-auto rounded-lg border border-border/60">
          <table className="w-full text-xs">
            <thead>
              <tr className="border-b border-border text-left text-[11px] uppercase text-muted-foreground">
                <th className="px-2 py-1">{t("for.colNumero")}</th>
                <th className="px-2 py-1">{t("fis.colData")}</th>
                <th className="px-2 py-1">{t("for.colScad")}</th>
                <th className="px-2 py-1 text-right">{t("for.colResiduo")}</th>
                <th className="px-2 py-1 text-center">{t("for.colModalita")}</th>
              </tr>
            </thead>
            <tbody>
              {fatture.map((f, i) => (
                <tr key={`${f.numero}|${i}`} className="border-b border-border/30">
                  <td className="whitespace-nowrap px-2 py-0.5">{f.numero}</td>
                  <td className="whitespace-nowrap px-2 py-0.5 text-muted-foreground">
                    {dataIt(f.data)}
                  </td>
                  <td
                    className={`whitespace-nowrap px-2 py-0.5 ${f.inRitardo ? "text-status-absent" : ""}`}
                  >
                    {dataIt(f.scadenza)}
                  </td>
                  <td className="whitespace-nowrap px-2 py-0.5 text-right tabular-nums">
                    {fmtImporto(f.residuo)}
                  </td>
                  <td className="px-2 py-0.5 text-center">
                    <span className="rounded-full bg-muted px-1.5 text-[10px]">
                      {BADGE[f.modalita]}
                    </span>
                  </td>
                </tr>
              ))}
              {fatture.length === 0 && (
                <tr>
                  <td colSpan={5} className="py-3 text-center text-muted-foreground">
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
