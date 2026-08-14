// Card AUTONOMA delle regole di classificazione fatture: stessa lista
// SharePoint della card dentro la tab Fatture, ma vive anche da sola nella
// tab Regole di Finanze (richiesta direzione: regole movimenti e regole
// fatture fianco a fianco). Form AND/OR, vocabolario condiviso con le
// regole dei movimenti, import da Excel, elenco con matita e cestino.
import { useEffect, useMemo, useRef, useState } from "react";
import { Pencil, Trash2 } from "lucide-react";
import { toast } from "sonner";
import { useLang } from "../lib/i18n";
import type { RegolaFattura } from "../lib/fatture-logic";
import type { RegolaFinanza } from "../lib/finanza-logic";
import {
  spCreateRegolaFattura,
  spDeleteRegolaFattura,
  spGetRegoleFatture,
  spGetRegoleFinanza,
  spUpdateRegolaFattura,
} from "../lib/sharepoint.functions";
import { CampoVocabolario } from "./CampoVocabolario";

export function RegoleFattureCard() {
  const { t } = useLang();
  const [regole, setRegole] = useState<RegolaFattura[]>([]);
  const [regoleFin, setRegoleFin] = useState<RegolaFinanza[]>([]);
  const [dir, setDir] = useState<"Ricevuta" | "Emessa">("Ricevuta");
  const [cliente, setCliente] = useState("");
  const [oggetto, setOggetto] = useState("");
  const [operatore, setOperatore] = useState<"AND" | "OR">("AND");
  const [tipologia, setTipologia] = useState("");
  const [sottocat, setSottocat] = useState("");
  const [allocPri, setAllocPri] = useState("");
  const [allocSec, setAllocSec] = useState("");
  const [servizio, setServizio] = useState("");
  const [nota, setNota] = useState("");
  const [editId, setEditId] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  interface RigaImport {
    fornitore: string;
    oggettoInclude?: string;
    tipologia?: string;
    sottocategoria?: string;
    allocPrimaria?: string;
    allocSecondaria?: string;
    note?: string;
  }
  const [imp, setImp] = useState<{ nuove: RigaImport[]; doppioni: number } | null>(null);
  const [impBusy, setImpBusy] = useState(false);
  const [impProg, setImpProg] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);
  const [cerca, setCerca] = useState("");
  const [uniBusy, setUniBusy] = useState(false);
  const [gruppiChiusi, setGruppiChiusi] = useState<Set<string>>(new Set());

  useEffect(() => {
    spGetRegoleFatture()
      .then((l) => setRegole(l as RegolaFattura[]))
      .catch(() => setRegole([]));
    spGetRegoleFinanza()
      .then((l) => setRegoleFin(l as RegolaFinanza[]))
      .catch(() => setRegoleFin([]));
  }, []);

  const vocab = useMemo(() => {
    const uniq = (xs: (string | undefined)[]) =>
      [...new Set(xs.filter(Boolean) as string[])].sort((a, b) => a.localeCompare(b));
    const tutte = [...regoleFin, ...regole].map((r) => ({
      tipologia: r.tipologia,
      sottocategoria: r.sottocategoria,
      allocPrimaria: r.allocPrimaria,
      allocSecondaria: r.allocSecondaria,
    }));
    const perTip = tipologia.trim()
      ? tutte.filter((r) => (r.tipologia ?? "") === tipologia.trim())
      : tutte;
    const perPri = allocPri.trim()
      ? tutte.filter((r) => (r.allocPrimaria ?? "") === allocPri.trim())
      : tutte;
    const sotto = uniq(perTip.map((r) => r.sottocategoria));
    const sec = uniq(perPri.map((r) => r.allocSecondaria));
    return {
      tipologie: uniq(tutte.map((r) => r.tipologia)),
      sottocat: sotto.length ? sotto : uniq(tutte.map((r) => r.sottocategoria)),
      allocPri: uniq(tutte.map((r) => r.allocPrimaria)),
      allocSec: sec.length ? sec : uniq(tutte.map((r) => r.allocSecondaria)),
    };
  }, [regoleFin, regole, tipologia, allocPri]);

  const reset = () => {
    setCliente("");
    setOggetto("");
    setOperatore("AND");
    setTipologia("");
    setSottocat("");
    setAllocPri("");
    setAllocSec("");
    setServizio("");
    setNota("");
    setEditId(null);
  };

  const salva = async () => {
    setBusy(true);
    try {
      const payload = {
        fornitore: cliente.trim(),
        oggettoInclude: oggetto.trim() || undefined,
        operatore,
        direzione: dir,
        tipologia: tipologia.trim() || undefined,
        sottocategoria: sottocat.trim() || undefined,
        allocPrimaria: allocPri.trim() || undefined,
        allocSecondaria: allocSec.trim() || undefined,
        clienteRif: servizio.trim() || undefined,
        note: nota.trim() || undefined,
      };
      if (editId) await spUpdateRegolaFattura({ data: { regolaId: editId, ...payload } });
      else await spCreateRegolaFattura({ data: payload });
      setRegole((await spGetRegoleFatture()) as RegolaFattura[]);
      reset();
      toast.success(t("ft.rfSalvata"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setBusy(false);
    }
  };

  // Stessa lettura dell'import nella tab Fatture: fogli con header
  // CLIENTE/FORNITORE; "Se oggetto o descrizione include" = condizione AND;
  // la NOTA spezza in due regole gemelle (specifica con oggetto=nota +
  // generica), stessi esiti; l'ultimo foglio vince; le presenti si saltano.
  const leggiExcel = async (file: File) => {
    setImp(null);
    toast.info(t("ft.rfImpLettura"), { description: file.name });
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
      const per = new Map<string, RigaImport>();
      for (const name of wb.SheetNames) {
        const mat = XLSX.utils.sheet_to_json(wb.Sheets[name], {
          header: 1,
          raw: false,
        }) as unknown[][];
        const hIdx = mat.findIndex(
          (r) =>
            String(r?.[0] ?? "")
              .trim()
              .toUpperCase() === "CLIENTE/FORNITORE",
        );
        if (hIdx < 0) continue;
        const header = (mat[hIdx] ?? []).map((c) =>
          String(c ?? "")
            .trim()
            .toLowerCase(),
        );
        const conOggetto = (header[1] ?? "").startsWith("se oggetto");
        const off = conOggetto ? 1 : 0;
        const iNota = header.findIndex((h) => h === "nota");
        for (const r of mat.slice(hIdx + 1)) {
          const fornitore = String(r?.[0] ?? "").trim();
          const tipologia2 = String(r?.[1 + off] ?? "").trim();
          if (!fornitore || !tipologia2) continue;
          const oggettoInc = conOggetto ? String(r?.[1] ?? "").trim() : "";
          const nota2 = iNota >= 0 ? String(r?.[iNota] ?? "").trim() : "";
          const base = {
            fornitore,
            tipologia: tipologia2,
            sottocategoria: String(r?.[2 + off] ?? "").trim() || undefined,
            allocPrimaria: String(r?.[3 + off] ?? "").trim() || undefined,
            allocSecondaria: String(r?.[4 + off] ?? "").trim() || undefined,
          };
          if (conOggetto) {
            per.set(`${fornitore.toLowerCase()}|${oggettoInc.toLowerCase()}`, {
              ...base,
              oggettoInclude: oggettoInc || undefined,
              note: nota2 || undefined,
            });
          } else if (nota2) {
            per.set(`${fornitore.toLowerCase()}|${nota2.toLowerCase()}`, {
              ...base,
              oggettoInclude: nota2,
              note: nota2,
            });
            per.set(`${fornitore.toLowerCase()}|`, { ...base });
          } else {
            per.set(`${fornitore.toLowerCase()}|`, { ...base });
          }
        }
      }
      const esistenti = new Set(
        regole
          .filter((r) => (r.direzione ?? "Ricevuta") === "Ricevuta")
          .map(
            (r) =>
              `${(r.fornitore ?? "").trim().toLowerCase()}|${(r.oggettoInclude ?? "")
                .trim()
                .toLowerCase()}`,
          ),
      );
      const tutte = [...per.values()];
      const nuove = tutte.filter(
        (r) =>
          !esistenti.has(`${r.fornitore.toLowerCase()}|${(r.oggettoInclude ?? "").toLowerCase()}`),
      );
      if (!tutte.length) {
        toast.error(t("ft.rfImpNiente"));
        return;
      }
      setImp({ nuove, doppioni: tutte.length - nuove.length });
      toast.success(t("ft.rfImpPronta"), {
        description: `${nuove.length} ${t("ft.rfImpNuove")} · ${tutte.length - nuove.length} ${t("ft.rfImpDoppie")}`,
      });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    }
  };

  const creaImportate = async () => {
    if (!imp?.nuove.length) return;
    setImpBusy(true);
    try {
      let fatte = 0;
      for (const r of imp.nuove) {
        await spCreateRegolaFattura({
          data: {
            fornitore: r.fornitore,
            oggettoInclude: r.oggettoInclude,
            operatore: "AND" as const,
            direzione: "Ricevuta" as const,
            tipologia: r.tipologia,
            sottocategoria: r.sottocategoria,
            allocPrimaria: r.allocPrimaria,
            allocSecondaria: r.allocSecondaria,
            note: r.note,
          },
        });
        fatte++;
        setImpProg(`${fatte} / ${imp.nuove.length}`);
      }
      setRegole((await spGetRegoleFatture()) as RegolaFattura[]);
      setImp(null);
      toast.success(t("ft.rfImpOk"), { description: `${fatte}` });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImpBusy(false);
      setImpProg("");
    }
  };

  // UNIFICA (come le regole movimenti): regole con gli STESSI esiti e
  // stessa condizione sull'oggetto si fondono in una sola con l'elenco dei
  // fornitori separato da virgole; oltre i 240 caratteri si spezza in
  // regole gemelle (limite della colonna SharePoint a riga singola).
  const unifica = async () => {
    const daDir = regole.filter((r) => (r.direzione ?? "Ricevuta") === dir);
    const gruppi = new Map<string, RegolaFattura[]>();
    for (const r of daDir) {
      const k = [
        (r.oggettoInclude ?? "").trim().toLowerCase(),
        r.operatore ?? "AND",
        (r.tipologia ?? "").trim().toLowerCase(),
        (r.sottocategoria ?? "").trim().toLowerCase(),
        (r.allocPrimaria ?? "").trim().toLowerCase(),
        (r.allocSecondaria ?? "").trim().toLowerCase(),
        (r.clienteRif ?? "").trim().toLowerCase(),
        (r.note ?? "").trim().toLowerCase(),
      ].join("\u0001");
      const arr = gruppi.get(k) ?? [];
      arr.push(r);
      gruppi.set(k, arr);
    }
    const daFondere = [...gruppi.values()].filter((g) => g.length > 1);
    if (!daFondere.length) {
      toast.success(t("ft.rfUniNiente"));
      return;
    }
    if (!window.confirm(`${t("ft.rfUniConfirm")} (${daDir.length} → ?)`)) return;
    setUniBusy(true);
    try {
      let finali = daDir.length - daFondere.reduce((s2, g) => s2 + g.length, 0);
      for (const g of daFondere) {
        const termini = [
          ...new Set(
            g.flatMap((r) =>
              (r.fornitore ?? "")
                .split(/[,;\n]/)
                .map((x) => x.trim())
                .filter(Boolean),
            ),
          ),
        ];
        const blocchi: string[] = [];
        let cur = "";
        for (const t2 of termini) {
          const next = cur ? `${cur}, ${t2}` : t2;
          if (next.length > 240 && cur) {
            blocchi.push(cur);
            cur = t2;
          } else cur = next;
        }
        if (cur) blocchi.push(cur);
        finali += blocchi.length;
        const base = g[0];
        const payload = {
          oggettoInclude: base.oggettoInclude,
          operatore: base.operatore === "OR" ? ("OR" as const) : ("AND" as const),
          direzione: dir,
          tipologia: base.tipologia,
          sottocategoria: base.sottocategoria,
          allocPrimaria: base.allocPrimaria,
          allocSecondaria: base.allocSecondaria,
          clienteRif: base.clienteRif,
          note: base.note,
        };
        for (let i = 0; i < blocchi.length; i++) {
          const target = g[i];
          if (target)
            await spUpdateRegolaFattura({
              data: { regolaId: target.id ?? "", fornitore: blocchi[i], ...payload },
            });
          else await spCreateRegolaFattura({ data: { fornitore: blocchi[i], ...payload } });
        }
        for (let i = blocchi.length; i < g.length; i++)
          await spDeleteRegolaFattura({ data: { id: g[i].id ?? "" } });
      }
      setRegole((await spGetRegoleFatture()) as RegolaFattura[]);
      toast.success(t("ft.rfUniOk"), { description: `${daDir.length} → ${finali}` });
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setUniBusy(false);
    }
  };

  const inputCls = "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm";
  const elenco = regole
    .filter((r) => (r.direzione ?? "Ricevuta") === dir)
    .filter((r) => {
      const q = cerca.trim().toLowerCase();
      if (!q) return true;
      return [
        r.fornitore,
        r.oggettoInclude,
        r.tipologia,
        r.sottocategoria,
        r.allocPrimaria,
        r.allocSecondaria,
        r.clienteRif,
        r.note,
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase()
        .includes(q);
    });
  return (
    <div className="mb-4 rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      <div className="mb-2 flex items-center justify-between">
        <div className="text-sm font-semibold text-foreground">
          {t("ft.rfTitolo")} ({elenco.length})
        </div>
        <div className="inline-flex rounded-lg border border-border p-0.5 text-xs">
          <button
            type="button"
            onClick={() => setDir("Ricevuta")}
            className={`rounded-md px-3 py-1.5 font-semibold ${dir === "Ricevuta" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
          >
            {t("ft.dirRicevute")}
          </button>
          <button
            type="button"
            onClick={() => setDir("Emessa")}
            className={`rounded-md px-3 py-1.5 font-semibold ${dir === "Emessa" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
          >
            {t("ft.dirEmesse")}
          </button>
        </div>
      </div>
      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
        <div>
          <label className="text-xs text-muted-foreground">{t("ft.rfClienteInclude")}</label>
          <textarea
            value={cliente}
            onChange={(e) => setCliente(e.target.value)}
            rows={2}
            className={inputCls}
          />
        </div>
        <div className="flex items-end pb-1">
          <div className="inline-flex rounded-lg border border-border p-0.5 text-xs">
            <button
              type="button"
              onClick={() => setOperatore("AND")}
              className={`rounded-md px-3 py-1.5 font-semibold ${operatore === "AND" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
            >
              AND
            </button>
            <button
              type="button"
              onClick={() => setOperatore("OR")}
              className={`rounded-md px-3 py-1.5 font-semibold ${operatore === "OR" ? "bg-primary text-primary-foreground" : "text-muted-foreground"}`}
            >
              OR
            </button>
          </div>
        </div>
        <div>
          <label className="text-xs text-muted-foreground">{t("ft.rfOggettoInclude")}</label>
          <textarea
            value={oggetto}
            onChange={(e) => setOggetto(e.target.value)}
            rows={2}
            className={inputCls}
          />
        </div>
        <div>
          <CampoVocabolario
            label={t("ft.colTipologia")}
            valore={tipologia}
            onChange={(v2) => {
              setTipologia(v2);
              if (v2.trim() && sottocat.trim() && !vocab.sottocat.includes(sottocat))
                setSottocat("");
            }}
            opzioni={vocab.tipologie}
            testoNessuno={t("ft.rfNonImpostare")}
            testoNuova={t("fin.vocNuova")}
          />
        </div>
        <div>
          <CampoVocabolario
            label={t("fin.sottocat")}
            valore={sottocat}
            onChange={setSottocat}
            opzioni={vocab.sottocat}
            testoNessuno={t("ft.rfNonImpostare")}
            testoNuova={t("fin.vocNuova")}
          />
        </div>
        <div>
          <CampoVocabolario
            label={t("fin.allocPri")}
            valore={allocPri}
            onChange={setAllocPri}
            opzioni={vocab.allocPri}
            testoNessuno={t("ft.rfNonImpostare")}
            testoNuova={t("fin.vocNuova")}
          />
        </div>
        <div>
          <CampoVocabolario
            label={t("fin.allocSec")}
            valore={allocSec}
            onChange={setAllocSec}
            opzioni={vocab.allocSec}
            testoNessuno={t("ft.rfNonImpostare")}
            testoNuova={t("fin.vocNuova")}
          />
        </div>
        <div>
          <label className="text-xs text-muted-foreground">
            {dir === "Ricevuta" ? t("ft.colClienteRif") : t("ft.colServizio")}
          </label>
          <input
            value={servizio}
            onChange={(e) => setServizio(e.target.value)}
            className={inputCls}
          />
        </div>
        <div>
          <label className="text-xs text-muted-foreground">{t("fin.note")}</label>
          <input value={nota} onChange={(e) => setNota(e.target.value)} className={inputCls} />
        </div>
      </div>
      <div className="mt-3 flex gap-2">
        <button
          type="button"
          disabled={busy || (!cliente.trim() && !oggetto.trim())}
          onClick={() => void salva()}
          className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
        >
          {busy ? t("common.loading") : editId ? t("fin.regolaAggiorna") : t("fin.regolaCrea")}
        </button>
        {editId && (
          <button
            type="button"
            onClick={reset}
            className="rounded-lg border border-border px-4 py-2 text-sm hover:bg-muted"
          >
            {t("common.cancel")}
          </button>
        )}
      </div>
      {dir === "Ricevuta" && (
        <div className="mt-3 rounded-lg border border-border/60 bg-muted/30 p-2">
          <div className="flex items-center gap-3">
            <span className="text-xs font-medium text-foreground">{t("ft.rfImpTitolo")}</span>
            {/* Il campo file NATIVO sembrava testo morto: bottone vero che
                apre la scelta, input nascosto che fa il lavoro. */}
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) void leggiExcel(f);
                e.target.value = "";
              }}
            />
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              className="rounded-lg bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90"
            >
              {t("ft.rfImpScegli")}
            </button>
          </div>
          <p className="mt-1 text-[11px] text-muted-foreground">{t("ft.rfImpDesc")}</p>
          {imp && (
            <div className="mt-2 flex items-center gap-3 text-[12px]">
              <span>
                <b className="text-status-present">{imp.nuove.length}</b> {t("ft.rfImpNuove")} ·{" "}
                {imp.doppioni} {t("ft.rfImpDoppie")}
              </span>
              <button
                type="button"
                disabled={impBusy || imp.nuove.length === 0}
                onClick={() => void creaImportate()}
                className="rounded-lg bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground disabled:opacity-50"
              >
                {impBusy
                  ? `${t("common.loading")} ${impProg}`
                  : `${t("ft.rfImpBtn")} (${imp.nuove.length})`}
              </button>
              {!impBusy && (
                <button
                  type="button"
                  onClick={() => setImp(null)}
                  className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted"
                >
                  {t("common.cancel")}
                </button>
              )}
            </div>
          )}
        </div>
      )}
      <div className="mt-3 flex items-center gap-2">
        <input
          value={cerca}
          onChange={(e) => setCerca(e.target.value)}
          placeholder={t("ft.rfCerca")}
          className="w-64 rounded-lg border border-border bg-background px-3 py-1.5 text-xs"
        />
        <button
          type="button"
          disabled={uniBusy}
          onClick={() => void unifica()}
          className="rounded-lg border border-border px-3 py-1.5 text-xs hover:bg-muted disabled:opacity-50"
        >
          {uniBusy ? t("common.loading") : t("ft.rfUniBtn")}
        </button>
      </div>
      {/* ELENCO IN STILE REGOLE MOVIMENTI: gruppi per tipologia
          richiudibili, frase leggibile, esiti come chip. */}
      <div className="mt-3 space-y-2">
        {(() => {
          const perTip = new Map<string, typeof elenco>();
          for (const r of elenco) {
            const k = r.tipologia?.trim() || t("ft.classVuota");
            const arr = perTip.get(k) ?? [];
            arr.push(r);
            perTip.set(k, arr);
          }
          return [...perTip.entries()]
            .sort((x, y) => x[0].localeCompare(y[0]))
            .map(([tip, rs]) => {
              const chiuso = gruppiChiusi.has(tip);
              return (
                <div key={tip} className="rounded-xl border border-border/60 p-2">
                  <button
                    type="button"
                    onClick={() =>
                      setGruppiChiusi((prev) => {
                        const next = new Set(prev);
                        if (next.has(tip)) next.delete(tip);
                        else next.add(tip);
                        return next;
                      })
                    }
                    className="text-sm font-semibold text-primary"
                  >
                    {chiuso ? "▸" : "▾"} {tip}{" "}
                    <span className="font-normal text-muted-foreground">({rs.length})</span>
                  </button>
                  {!chiuso &&
                    rs.map((r) => (
                      <div
                        key={r.id}
                        className="flex items-start gap-2 border-t border-border/40 py-1.5 text-[13px]"
                      >
                        <span className="flex-1 leading-relaxed">
                          {r.fornitore && (
                            <>
                              {t("ft.rfFraseCliente")} «<b>{r.fornitore}</b>»
                            </>
                          )}
                          {r.fornitore && r.oggettoInclude && (
                            <b className="text-primary"> {r.operatore ?? "AND"} </b>
                          )}
                          {r.oggettoInclude && (
                            <>
                              {t("ft.rfFraseOggetto")} «<b>{r.oggettoInclude}</b>»
                            </>
                          )}{" "}
                          →{" "}
                          {r.sottocategoria && (
                            <span className="mr-1 rounded bg-primary/10 px-1.5 py-0.5 text-primary">
                              {t("fin.sottocat")} → {r.sottocategoria}
                            </span>
                          )}
                          {(r.allocPrimaria || r.allocSecondaria) && (
                            <span className="mr-1 rounded bg-muted px-1.5 py-0.5">
                              {[r.allocPrimaria, r.allocSecondaria].filter(Boolean).join(" / ")}
                            </span>
                          )}
                          {r.clienteRif && (
                            <span className="mr-1 rounded bg-primary/10 px-1.5 py-0.5 text-primary">
                              {t("ft.colClienteRif")} → {r.clienteRif}
                            </span>
                          )}
                          {r.note && <i className="text-muted-foreground"> — {r.note}</i>}
                        </span>
                        <button
                          type="button"
                          onClick={() => {
                            setEditId(r.id ?? null);
                            setCliente(r.fornitore ?? "");
                            setOggetto(r.oggettoInclude ?? "");
                            setOperatore(r.operatore === "OR" ? "OR" : "AND");
                            setDir(r.direzione === "Emessa" ? "Emessa" : "Ricevuta");
                            setTipologia(r.tipologia ?? "");
                            setSottocat(r.sottocategoria ?? "");
                            setAllocPri(r.allocPrimaria ?? "");
                            setAllocSec(r.allocSecondaria ?? "");
                            setServizio(r.clienteRif ?? "");
                            setNota(r.note ?? "");
                          }}
                          className="text-muted-foreground hover:text-foreground"
                          title={t("fin.regolaAggiorna")}
                        >
                          <Pencil className="h-3.5 w-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            if (!window.confirm(t("ft.rfDelConfirm"))) return;
                            void spDeleteRegolaFattura({ data: { id: r.id ?? "" } })
                              .then(() => setRegole((prev) => prev.filter((x) => x.id !== r.id)))
                              .catch((err) =>
                                toast.error(t("common.error"), {
                                  description: err instanceof Error ? err.message : String(err),
                                }),
                              );
                          }}
                          className="text-muted-foreground hover:text-status-absent"
                          title={t("common.delete")}
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    ))}
                </div>
              );
            });
        })()}
      </div>
    </div>
  );
}
