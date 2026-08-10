// DR Portal — Finanza → tab Resoconto (direttore).
// Il resoconto per controparte: si scelgono uno o più CLIENTI e, col secondo
// filtro, uno o più FORNITORI, e si ottiene la riconciliazione del file del
// direttore — Estratto conto vs Fatture attive vs Fatture passive, con l'OK
// verde quando torna — più lo specchietto dei RITARDI (da incassare e da
// pagare) filtrabile per fasce di giorni/settimane, con le fatture in vista.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { CheckCircle2, AlertTriangle, Loader2, Trash2, Copy, Mail } from "lucide-react";
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
import {
  spGetFatture,
  spGetMovimenti,
  spGetTerminiPagamento,
  spGetGruppiControparti,
  spCreateGruppoControparti,
  spDeleteGruppoControparti,
} from "@/lib/sharepoint.functions";
import type { SpFattura, SpMovimento, GruppoControparti } from "@/lib/sharepoint.server";

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

// "Nessuno" nei filtri: esclude un lato intero del confronto. Serve quando
// una controparte è SOLO fornitore (o solo cliente): senza, il lato "Tutte"
// trascina dentro l'intera azienda e si confrontano mele con pere.
const NESSUNO = "__nessuno__";

export function ResocontoTab() {
  const { t } = useLang();
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [clientiSel, setClientiSel] = useState<string[]>([]);
  const [fornitoriSel, setFornitoriSel] = useState<string[]>([]);
  // Filtro esplicito sull'ESTRATTO CONTO: vuoto = "automatico" (segue
  // Cliente/Fornitore come sempre); una selezione lo scavalca.
  const [estrattoSel, setEstrattoSel] = useState<string[]>([]);
  // Gruppi "madre" (es. UNIVEX = univex, nolvex): voci aggiuntive nei filtri
  // che si espandono in tutte le controparti i cui nomi contengono i membri.
  const [gruppi, setGruppi] = useState<GruppoControparti[] | null>(null);
  const [showGruppi, setShowGruppi] = useState(false);
  const [gNome, setGNome] = useState("");
  const [gBusy, setGBusy] = useState(false);
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
    spGetGruppiControparti()
      .then((l) => setGruppi(l as GruppoControparti[]))
      .catch(() => setGruppi([]));
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
  // Controparti viste in BANCA (per il filtro estratto conto), per volumi.
  const opzioniEstratto = useMemo(() => {
    const somme = new Map<string, { nome: string; tot: number }>();
    for (const m of movimenti ?? []) {
      if (!m.cliente) continue;
      const k = clienteGroupKey(m.cliente) || m.cliente;
      const v2 = somme.get(k) ?? { nome: m.cliente, tot: 0 };
      v2.tot += Math.abs(m.importo);
      somme.set(k, v2);
    }
    return [...somme.entries()]
      .sort((a, b) => b[1].tot - a[1].tot)
      .map(([k, v2]) => ({ v: k, label: v2.nome }));
  }, [movimenti]);
  // Un gruppo selezionato ("grp:<id>") si espande nelle chiavi delle opzioni
  // i cui nomi contengono uno dei membri (match per nome contenuto).
  const espandiGruppi = (sel: string[], opzioniDi: { v: string }[]): string[] => {
    const out = new Set<string>();
    for (const v of sel) {
      if (!v.startsWith("grp:")) {
        out.add(v);
        continue;
      }
      const g = (gruppi ?? []).find((x) => `grp:${x.id}` === v);
      if (!g) continue;
      const token = g.membri
        .split(/[,;|]/)
        .map((x) => clienteGroupKey(x) || x.trim().toLowerCase())
        .filter(Boolean);
      for (const o of opzioniDi) if (token.some((tk) => o.v.includes(tk))) out.add(o.v);
    }
    return [...out];
  };
  // VISTE SALVATE: la fotografia dei tre filtri, con un nome. Membri (sulla
  // lista GruppiControparti) contiene il JSON {"c":[...],"f":[...],"e":[...]}.
  type VistaCfg = { c: string[]; f: string[]; e: string[] };
  const parseVista = (membri: string): VistaCfg | null => {
    try {
      const j = JSON.parse(membri) as Partial<VistaCfg>;
      if (!j || typeof j !== "object" || Array.isArray(j)) return null;
      return {
        c: Array.isArray(j.c) ? j.c : [],
        f: Array.isArray(j.f) ? j.f : [],
        e: Array.isArray(j.e) ? j.e : [],
      };
    } catch {
      return null;
    }
  };
  const viste = (gruppi ?? [])
    .map((g) => ({ id: g.id, nome: g.nome, cfg: parseVista(g.membri) }))
    .filter((x): x is { id: string; nome: string; cfg: VistaCfg } => x.cfg != null);
  // Nel riepilogo della vista si mostrano i NOMI, non i conteggi: le chiavi
  // salvate si traducono in etichette con le stesse opzioni delle tendine.
  const nomiDi = (chiavi: string[], opzioniDi: { v: string; label: string }[]): string =>
    chiavi
      .map((k) =>
        k === NESSUNO ? t("rt.nessunoOpz") : (opzioniDi.find((o) => o.v === k)?.label ?? k),
      )
      .join(", ");
  // Scelta con "Nessuno" esclusivo: selezionarlo azzera il resto, scegliere
  // una controparte lo toglie.
  const scegliCon = (imposta: (v: string[]) => void, prima: string[]) => (nuovi: string[]) => {
    if (nuovi.includes(NESSUNO) && !prima.includes(NESSUNO)) imposta([NESSUNO]);
    else imposta(nuovi.filter((x) => x !== NESSUNO));
  };

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
    const cliNessuno = clientiSel.includes(NESSUNO);
    const forNessuno = fornitoriSel.includes(NESSUNO);
    const ecNessuno = estrattoSel.includes(NESSUNO);
    const cliSel = espandiGruppi(
      clientiSel.filter((x) => x !== NESSUNO),
      opzioniClienti,
    );
    const forSel = espandiGruppi(
      fornitoriSel.filter((x) => x !== NESSUNO),
      opzioniFornitori,
    );
    const ecSel = espandiGruppi(
      estrattoSel.filter((x) => x !== NESSUNO),
      opzioniEstratto,
    );
    const attSel = cliNessuno ? [] : attive.filter((x) => inSelezione(x.f.cliente, cliSel));
    const pasSel = forNessuno ? [] : passive.filter((x) => inSelezione(x.f.cliente, forSel));
    const movTutti = movimenti ?? [];
    let movSel: typeof movTutti;
    if (ecNessuno) {
      movSel = [];
    } else if (ecSel.length) {
      // Filtro estratto ESPLICITO: vince su tutto il resto.
      const chiavi = new Set(ecSel);
      movSel = movTutti.filter(
        (m) => m.cliente && chiavi.has(clienteGroupKey(m.cliente) || m.cliente),
      );
    } else {
      // Automatico: l'estratto segue Cliente/Fornitore. Con nessun filtro il
      // quadro è GLOBALE: tutti i movimenti con controparte, tutte le fatture.
      const chiavi = new Set([
        ...(cliNessuno ? [] : cliSel.length ? cliSel : opzioniClienti.map((o) => o.v)),
        ...(forNessuno ? [] : forSel),
      ]);
      const tuttoSelezionato =
        !cliNessuno && !forNessuno && cliSel.length === 0 && forSel.length === 0;
      movSel = movTutti.filter((m) => {
        if (!m.cliente) return false;
        const k = clienteGroupKey(m.cliente) || m.cliente;
        return tuttoSelezionato ? true : chiavi.has(k);
      });
    }
    const estratto = movSel.reduce((s, m) => s + m.importo, 0);
    const incassatoAtt = attSel.reduce((s, x) => s + incassatoDi(x), 0);
    const pagatoPas = pasSel.reduce((s, x) => s + incassatoDi(x), 0);
    const differenza = Math.round((estratto - incassatoAtt + pagatoPas) * 100) / 100;
    return {
      estratto: Math.round(estratto * 100) / 100,
      incassatoAtt: Math.round(incassatoAtt * 100) / 100,
      pagatoPas: Math.round(pagatoPas * 100) / 100,
      // Il "netto fatture" del direttore: incassato attive meno pagato
      // passive — è il numero da confrontare con l'estratto conto.
      netto: Math.round((incassatoAtt - pagatoPas) * 100) / 100,
      differenza,
      nMov: movSel.length,
      ok: Math.abs(differenza) <= 1,
    };

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attive, passive, movimenti, clientiSel, fornitoriSel, estrattoSel, opzioniClienti, gruppi]);

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
    () =>
      clientiSel.includes(NESSUNO)
        ? []
        : ritardi(
            attive,
            espandiGruppi(
              clientiSel.filter((x) => x !== NESSUNO),
              opzioniClienti,
            ),
          ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [attive, clientiSel, fasceSel],
  );
  const ritardiPas = useMemo(
    () =>
      fornitoriSel.includes(NESSUNO)
        ? []
        : ritardi(
            passive,
            espandiGruppi(
              fornitoriSel.filter((x) => x !== NESSUNO),
              opzioniFornitori,
            ),
          ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, fornitoriSel, fasceSel],
  );

  // SOLLECITO per cliente, anche da qui: stesso testo della pagina Fatture
  // (approvato dal direttore) — scadute, NC e loro fatture da scalare.
  const testoSollecitoPer = (chiave: string) => {
    const del = attive.filter((x) => (clienteGroupKey(x.f.cliente) || x.f.cliente) === chiave);
    const nome = del[0]?.f.cliente ?? chiave;
    const scadute = del
      .filter(
        (x) =>
          !isNotaCredito(x.f.tipoDocumento) &&
          x.s.inRitardo &&
          x.s.residuo > 1 &&
          (x.s.statoIncassi != null || x.s.statoFatturazione != null),
      )
      .sort((a, b) => a.s.scadenza.localeCompare(b.s.scadenza));
    if (!scadute.length) return null;
    const ncAperte = del.filter(
      (x) =>
        isNotaCredito(x.f.tipoDocumento) &&
        !x.f.rettificaNumero &&
        x.s.statoIncassi !== "Pagata" &&
        x.s.statoFatturazione !== "Pagata",
    );
    const loroAperte = (fattureRic ?? []).filter(
      (f) =>
        (clienteGroupKey(f.cliente) || f.cliente) === chiave &&
        !isNotaCredito(f.tipoDocumento) &&
        f.totale > 0 &&
        parseIncassoAruba(f.incassoAruba) !== "Incassata",
    );
    const totale = scadute.reduce((s2, x) => s2 + x.s.residuo, 0);
    const daScalare =
      ncAperte.reduce((s2, x) => s2 + Math.abs(x.f.totale), 0) +
      loroAperte.reduce((s2, f) => s2 + f.totale, 0);
    const r: string[] = [];
    r.push(`Spett.le ${nome},`);
    r.push("");
    r.push(
      "dal riscontro dei nostri registri contabili risultano ad oggi le seguenti fatture scadute e non ancora saldate:",
    );
    r.push("");
    for (const x of scadute)
      r.push(
        `- ${x.f.numero} del ${fmtData(x.f.dataDocumento)}, scadenza ${fmtData(x.s.scadenza)}, residuo € ${fmtImporto(x.s.residuo)} (${x.s.giorniRitardo} giorni di ritardo)`,
      );
    r.push("");
    r.push(`Totale scaduto: € ${fmtImporto(totale)}`);
    if (daScalare > 1) {
      r.push("");
      r.push("Da portare eventualmente in compensazione:");
      for (const x of ncAperte)
        r.push(`- ns. nota di credito ${x.f.numero}: € ${fmtImporto(Math.abs(x.f.totale))}`);
      for (const f of loroAperte)
        r.push(`- vs. fattura ${f.numero} nei nostri confronti: € ${fmtImporto(f.totale)}`);
      r.push(`Saldo netto richiesto: € ${fmtImporto(Math.max(0, totale - daScalare))}`);
    }
    r.push("");
    r.push(
      "Vi preghiamo di provvedere al saldo entro 7 giorni dal ricevimento della presente, ovvero di segnalarci eventuali pagamenti già disposti o discordanze riscontrate.",
    );
    r.push("");
    r.push("Restiamo a disposizione per ogni chiarimento.");
    r.push("");
    r.push("Cordiali saluti");
    r.push("DR Logistica S.r.l.");
    return r.join("\r\n");
  };
  const copiaSollecitoPer = async (chiave: string) => {
    const corpo = testoSollecitoPer(chiave);
    if (!corpo) {
      toast(t("ft.solNessuna"));
      return;
    }
    await navigator.clipboard.writeText(corpo);
    toast.success(t("ft.solCopiato"));
  };
  const emailSollecitoPer = (chiave: string) => {
    const corpo = testoSollecitoPer(chiave);
    if (!corpo) {
      toast(t("ft.solNessuna"));
      return;
    }
    const dest =
      termini.find(
        (x) =>
          (x.direzione ?? "Emessa") === "Emessa" &&
          (clienteGroupKey(x.cliente) || x.cliente) === chiave &&
          x.email,
      )?.email ?? "";
    const oggetto = "Sollecito di pagamento - fatture scadute (DR Logistica S.r.l.)";
    window.location.href = `mailto:${encodeURIComponent(dest)}?subject=${encodeURIComponent(oggetto)}&body=${encodeURIComponent(corpo)}`;
  };

  const loading = fattureEm == null || fattureRic == null || movimenti == null;

  const cardRitardi = (
    titolo: string,
    righe: typeof attive,
    vuoto: string,
    conSollecito = false,
  ) => (
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
                {conSollecito && <th className="py-1" />}
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
                  {conSollecito && (
                    <td className="py-0.5 whitespace-nowrap">
                      <button
                        type="button"
                        onClick={() =>
                          void copiaSollecitoPer(clienteGroupKey(x.f.cliente) || x.f.cliente)
                        }
                        title={t("ft.solCopia")}
                        className="rounded-md p-1 text-muted-foreground hover:text-foreground"
                      >
                        <Copy className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          emailSollecitoPer(clienteGroupKey(x.f.cliente) || x.f.cliente)
                        }
                        title={t("ft.solEmail")}
                        className="rounded-md p-1 text-muted-foreground hover:text-foreground"
                      >
                        <Mail className="h-3.5 w-3.5" />
                      </button>
                    </td>
                  )}
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
            opzioni={[{ v: NESSUNO, label: t("rt.nessunoOpz") }, ...opzioniClienti]}
            valori={clientiSel}
            onChange={scegliCon(setClientiSel, clientiSel)}
            className="w-64"
          />
          <MultiSelect
            label={t("ft.fornitore")}
            tuttiLabel={t("common.allF")}
            selLabel={t("fin.msSel")}
            opzioni={[{ v: NESSUNO, label: t("rt.nessunoOpz") }, ...opzioniFornitori]}
            valori={fornitoriSel}
            onChange={scegliCon(setFornitoriSel, fornitoriSel)}
            className="w-64"
          />
          <MultiSelect
            label={t("rt.estratto")}
            tuttiLabel={t("rt.estrattoAuto")}
            selLabel={t("fin.msSel")}
            opzioni={[{ v: NESSUNO, label: t("rt.nessunoOpz") }, ...opzioniEstratto]}
            valori={estrattoSel}
            onChange={scegliCon(setEstrattoSel, estrattoSel)}
            className="w-64"
          />
          <button
            type="button"
            onClick={() => setShowGruppi((x) => !x)}
            className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            {t("rt.gruppiBtn")}
          </button>
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
        {showGruppi && (
          <div className="mt-3 rounded-xl border border-border p-3">
            <p className="text-xs text-muted-foreground mb-2">{t("rt.gruppiDesc")}</p>
            <div className="flex flex-wrap items-end gap-2 mb-2">
              <input
                value={gNome}
                onChange={(e) => setGNome(e.target.value)}
                placeholder={t("rt.gruppiNomePh")}
                className="w-64 rounded-lg border border-border bg-background px-3 py-2 text-sm"
              />
              <button
                type="button"
                disabled={
                  gBusy ||
                  !gNome.trim() ||
                  (clientiSel.length === 0 && fornitoriSel.length === 0 && estrattoSel.length === 0)
                }
                onClick={() => {
                  setGBusy(true);
                  // Si salva la FOTOGRAFIA dei tre filtri correnti.
                  spCreateGruppoControparti({
                    data: {
                      nome: gNome.trim(),
                      membri: JSON.stringify({ c: clientiSel, f: fornitoriSel, e: estrattoSel }),
                    },
                  })
                    .then(() => spGetGruppiControparti())
                    .then((l) => {
                      setGruppi(l as GruppoControparti[]);
                      setGNome("");
                    })
                    .catch((err) =>
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      }),
                    )
                    .finally(() => setGBusy(false));
                }}
                className="rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
              >
                {t("common.save")}
              </button>
            </div>
            {viste.map((g) => (
              <div key={g.id} className="flex items-center gap-2 py-1 text-sm">
                {/* Click sul nome = applica la vista: i tre filtri tornano
                    esattamente com'erano al salvataggio. */}
                <button
                  type="button"
                  onClick={() => {
                    setClientiSel(g.cfg.c);
                    setFornitoriSel(g.cfg.f);
                    setEstrattoSel(g.cfg.e);
                  }}
                  className="rounded-lg border border-border px-2.5 py-1 text-sm font-medium hover:bg-muted"
                >
                  {g.nome}
                </button>
                {(() => {
                  const testo = [
                    g.cfg.c.length ? `${t("fin.cliente")}: ${nomiDi(g.cfg.c, opzioniClienti)}` : "",
                    g.cfg.f.length
                      ? `${t("ft.fornitore")}: ${nomiDi(g.cfg.f, opzioniFornitori)}`
                      : "",
                    g.cfg.e.length
                      ? `${t("rt.estratto")}: ${nomiDi(g.cfg.e, opzioniEstratto)}`
                      : "",
                  ]
                    .filter(Boolean)
                    .join(" · ");
                  return (
                    <span className="flex-1 truncate text-xs text-muted-foreground" title={testo}>
                      {testo}
                    </span>
                  );
                })()}
                <button
                  type="button"
                  disabled={gBusy}
                  onClick={() => {
                    setGBusy(true);
                    spDeleteGruppoControparti({ data: { id: g.id } })
                      .then(() => setGruppi((prev) => (prev ?? []).filter((x) => x.id !== g.id)))
                      .catch((err) =>
                        toast.error(t("common.error"), {
                          description: err instanceof Error ? err.message : String(err),
                        }),
                      )
                      .finally(() => setGBusy(false));
                  }}
                  className="rounded-md p-1 text-muted-foreground hover:text-status-absent"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              </div>
            ))}
          </div>
        )}
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
              <span className="text-sm font-semibold text-foreground">
                {t("rt.titolo")}{" "}
                <span className="font-normal text-muted-foreground">
                  ({t("rt.sottotitoloAruba")})
                </span>
              </span>
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
            {/* Layout chiesto dal direttore: in alto il confronto secco
                (estratto · netto fatture · differenza), sotto il dettaglio
                attive/passive che compone il netto. */}
            <div className="grid gap-3 sm:grid-cols-3">
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
                  {t("rt.netto")}
                </div>
                <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
                  {fmtImporto(quadro.netto)} €
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
            <div className="mt-3 flex flex-wrap gap-x-8 gap-y-1 border-t border-border/60 pt-2 text-sm">
              <div>
                <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.attive")}
                </span>{" "}
                <b className="tabular-nums text-foreground">{fmtImporto(quadro.incassatoAtt)} €</b>
              </div>
              <div>
                <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.passive")}
                </span>{" "}
                <b className="tabular-nums text-foreground">{fmtImporto(quadro.pagatoPas)} €</b>
              </div>
            </div>
            <p className="mt-2 text-[11px] text-muted-foreground">{t("rt.nota")}</p>
          </div>

          {/* I ritardi, nelle due direzioni, con le fatture in chiaro. */}
          <div className="grid gap-4 lg:grid-cols-2">
            {cardRitardi(t("rt.ritardiIncassare"), ritardiAtt, t("rt.nessunoIncassare"), true)}
            {cardRitardi(t("rt.ritardiPagare"), ritardiPas, t("rt.nessunoPagare"))}
          </div>
        </>
      )}
    </div>
  );
}
