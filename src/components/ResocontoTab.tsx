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
  spGetPrefatture,
  spCreatePrefattura,
  spDeletePrefattura,
  spGetGruppiControparti,
  spCreateGruppoControparti,
  spDeleteGruppoControparti,
} from "@/lib/sharepoint.functions";
import type {
  SpFattura,
  SpMovimento,
  GruppoControparti,
  Prefattura,
} from "@/lib/sharepoint.server";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtMese(yyyymm: string): string {
  const [y, m] = yyyymm.split("-");
  const nomi = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"];
  return `${nomi[Number(m) - 1] ?? m} ${y}`;
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
  // Prefatture: fatturato pianificato che entra nella Previsione.
  const [prefatture, setPrefatture] = useState<Prefattura[] | null>(null);
  const [pfControparte, setPfControparte] = useState("");
  const [pfDirezione, setPfDirezione] = useState<"Emessa" | "Ricevuta">("Emessa");
  const [pfImporto, setPfImporto] = useState("");
  const [pfMese, setPfMese] = useState("");
  const [pfRicorrenza, setPfRicorrenza] = useState<"mensile" | "una">("mensile");
  const [pfMeseFine, setPfMeseFine] = useState("");
  const [pfNote, setPfNote] = useState("");
  const [pfBusy, setPfBusy] = useState(false);
  // "" = solo ritardi (default); un numero = anche le scadenze future entro N giorni.
  const [scadEntro, setScadEntro] = useState("");

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
    spGetPrefatture()
      .then((l) => setPrefatture(l as Prefattura[]))
      .catch(() => setPrefatture([]));
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
        // NC collegate: servono alla colonna "Composizione" dei ritardi.
        nc: nc.get(f.nomeFile),
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
    // Residui (fatture non ancora incassate/pagate) per la riga dei totali:
    // il totale attive = incassato + da incassare; il "di cui in ritardo"
    // guarda le sole scadute.
    const resAtt = attSel.reduce((s, x) => s + Math.max(0, x.s.residuo), 0);
    const ritAtt = attSel.reduce((s, x) => s + (x.s.inRitardo ? Math.max(0, x.s.residuo) : 0), 0);
    const resPas = pasSel.reduce((s, x) => s + Math.max(0, x.s.residuo), 0);
    const differenza = Math.round((estratto - incassatoAtt + pagatoPas) * 100) / 100;
    return {
      estratto: Math.round(estratto * 100) / 100,
      incassatoAtt: Math.round(incassatoAtt * 100) / 100,
      pagatoPas: Math.round(pagatoPas * 100) / 100,
      // Il "netto fatture" del direttore: incassato attive meno pagato
      // passive — è il numero da confrontare con l'estratto conto.
      netto: Math.round((incassatoAtt - pagatoPas) * 100) / 100,
      differenza,
      totAtt: Math.round((incassatoAtt + resAtt) * 100) / 100,
      resAtt: Math.round(resAtt * 100) / 100,
      ritAtt: Math.round(ritAtt * 100) / 100,
      resPas: Math.round(resPas * 100) / 100,
      saldoResiduo: Math.round((resAtt - resPas) * 100) / 100,
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
  // Giorni ALLA scadenza (positivi = futura). Serve per includere nelle
  // card anche i pagamenti/incassi in arrivo ("fra N giorni").
  const giorniAScadenza = (scadenza?: string): number | null => {
    if (!scadenza) return null;
    return Math.round(
      (new Date(`${scadenza.slice(0, 10)}T00:00:00`).getTime() -
        new Date(`${oggiISO}T00:00:00`).getTime()) /
        86400000,
    );
  };
  const ritardi = (righe: typeof attive, sel: string[]) => {
    const entro = Number(scadEntro) || 0;
    return righe
      .filter((x) => {
        if (x.s.residuo <= 1) return false;
        if (x.s.statoIncassi == null && x.s.statoFatturazione == null) return false;
        if (!inSelezione(x.f.cliente, sel)) return false;
        if (x.s.inRitardo) return inFascia(x.s.giorniRitardo);
        // Non in ritardo: entra solo col filtro "in scadenza entro N gg".
        if (entro <= 0) return false;
        const gg = giorniAScadenza(x.s.scadenza);
        return gg != null && gg >= 0 && gg <= entro;
      })
      .sort((a, b) => b.s.giorniRitardo - a.s.giorniRitardo);
  };
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
                <th className="py-1 pr-2">{t("rt.composizione")}</th>
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
                  <td
                    className={`py-0.5 pr-2 text-right tabular-nums whitespace-nowrap ${x.s.inRitardo ? "text-status-absent" : "text-primary"}`}
                  >
                    {x.s.inRitardo
                      ? x.s.giorniRitardo
                      : `${t("rt.fra")} ${giorniAScadenza(x.s.scadenza) ?? "—"}`}
                  </td>
                  <td className="py-0.5 pr-2 text-right tabular-nums font-medium">
                    {fmtImporto(x.s.residuo)}
                  </td>
                  <td className="py-0.5 pr-2 whitespace-nowrap text-[11px] text-muted-foreground">
                    {x.nc && x.nc.importo > 0 ? (
                      <>
                        {fmtImporto(x.f.totale)} − NC {x.nc.numeri.join("+")}{" "}
                        {fmtImporto(x.nc.importo)}
                      </>
                    ) : x.f.netto > 0 && x.f.netto < x.f.totale - 0.01 ? (
                      <>
                        {t("rt.compTot")} {fmtImporto(x.f.totale)} → {t("rt.compNetto")}{" "}
                        {fmtImporto(x.f.netto)}
                      </>
                    ) : (
                      "—"
                    )}
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
          <div>
            <label className="text-xs text-muted-foreground">{t("rt.scadEntro")}</label>
            <input
              value={scadEntro}
              onChange={(e) => setScadEntro(e.target.value.replace(/[^0-9]/g, ""))}
              placeholder={t("rt.scadEntroPh")}
              className="block w-28 rounded-lg border border-border bg-background px-2 py-2 text-sm"
            />
          </div>
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
            {/* Riga chiesta dalla direzione: totale delle attive con lo
                spaccato incassate/da incassare (e il ritardo), e a destra i
                residui delle due direzioni con la differenza. */}
            <div className="mt-2 flex flex-wrap items-baseline justify-between gap-x-8 gap-y-1 border-t border-border/60 pt-2 text-sm">
              <div>
                <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                  {t("rt.totAttive")}
                </span>{" "}
                <b className="tabular-nums text-foreground">{fmtImporto(quadro.totAtt)} €</b>{" "}
                <span className="text-[11px] text-muted-foreground">
                  ({t("rt.incassate")} {fmtImporto(quadro.incassatoAtt)} € · {t("rt.nonIncassate")}{" "}
                  {fmtImporto(quadro.resAtt)} €, {t("rt.diCuiRitardo")} {fmtImporto(quadro.ritAtt)}{" "}
                  €)
                </span>
              </div>
              <div className="flex flex-wrap gap-x-6 gap-y-1">
                <div>
                  <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                    {t("rt.daIncassare")}
                  </span>{" "}
                  <b className="tabular-nums text-foreground">{fmtImporto(quadro.resAtt)} €</b>
                </div>
                <div>
                  <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                    {t("rt.daPagare")}
                  </span>{" "}
                  <b className="tabular-nums text-foreground">{fmtImporto(quadro.resPas)} €</b>
                </div>
                <div>
                  <span className="text-[11px] uppercase tracking-wider text-muted-foreground">
                    {t("rt.differenzaResidui")}
                  </span>{" "}
                  <b
                    className={
                      "tabular-nums " +
                      (quadro.saldoResiduo >= 0 ? "text-status-present" : "text-status-absent")
                    }
                  >
                    {fmtImporto(quadro.saldoResiduo)} €
                  </b>
                </div>
              </div>
            </div>
            <p className="mt-2 text-[11px] text-muted-foreground">{t("rt.nota")}</p>
          </div>

          {/* I ritardi, nelle due direzioni, con le fatture in chiaro. */}
          <div className="grid gap-4 lg:grid-cols-2">
            {cardRitardi(t("rt.ritardiIncassare"), ritardiAtt, t("rt.nessunoIncassare"), true)}
            {cardRitardi(t("rt.ritardiPagare"), ritardiPas, t("rt.nessunoPagare"))}
          </div>

          {/* PREVISIONE: quanto si incassa e si paga nei prossimi mesi,
              dalle scadenze delle fatture ancora aperte (lo scaduto e' la
              prima riga: e' cassa attesa anche lui, solo in ritardo). */}
          {(() => {
            const chiaveMese = (iso: string) => iso.slice(0, 7);
            const mesi6: string[] = [];
            {
              const base = new Date(`${oggiISO.slice(0, 7)}-01T00:00:00`);
              for (let i = 0; i < 6; i++) {
                const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
                mesi6.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
              }
            }
            const somma = (righe: typeof attive) => {
              const out = new Map<string, number>();
              let scaduto = 0;
              for (const x of righe) {
                if (x.s.residuo <= 1) continue;
                if (x.s.statoIncassi == null && x.s.statoFatturazione == null) continue;
                if (!x.s.scadenza) continue;
                if (x.s.inRitardo) {
                  scaduto += x.s.residuo;
                  continue;
                }
                const k = chiaveMese(x.s.scadenza);
                out.set(k, (out.get(k) ?? 0) + x.s.residuo);
              }
              return { out, scaduto };
            };
            const att = somma(attive);
            const pas = somma(passive);
            // PREFATTURE: mesi pianificati non ancora coperti da una
            // fattura vera della stessa controparte nello stesso mese.
            const copertura = (dir2: "Emessa" | "Ricevuta") => {
              const fonte = dir2 === "Emessa" ? (fattureEm ?? []) : (fattureRic ?? []);
              return new Set(
                fonte.map(
                  (f2) =>
                    `${clienteGroupKey(f2.cliente) || f2.cliente.toLowerCase()}|${f2.dataDocumento.slice(0, 7)}`,
                ),
              );
            };
            const prefSomme = (dir2: "Emessa" | "Ricevuta") => {
              const cov = copertura(dir2);
              const out = new Map<string, number>();
              for (const pf of (prefatture ?? []).filter((x) => x.direzione === dir2)) {
                const chiave = clienteGroupKey(pf.controparte) || pf.controparte.toLowerCase();
                const mesiPf =
                  pf.ricorrenza === "una"
                    ? mesi6.filter((m) => m === pf.meseInizio)
                    : mesi6.filter((m) => m >= pf.meseInizio && (!pf.meseFine || m <= pf.meseFine));
                for (const m of mesiPf) {
                  if (cov.has(`${chiave}|${m}`)) continue;
                  out.set(m, (out.get(m) ?? 0) + pf.importo);
                }
              }
              return out;
            };
            const prefAtt = prefSomme("Emessa");
            const prefPas = prefSomme("Ricevuta");
            const haPref = prefAtt.size > 0 || prefPas.size > 0;
            const fmt = (v2: number) => (v2 ? `${fmtImporto(Math.round(v2 * 100) / 100)} €` : "—");
            return (
              <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
                <div className="mb-2 text-sm font-semibold text-foreground">
                  {t("rt.prevTitolo")}
                </div>
                <p className="mb-3 text-xs text-muted-foreground">{t("rt.prevDesc")}</p>
                <div className="overflow-x-auto">
                  <table className="w-full text-[13px]">
                    <thead>
                      <tr className="text-left text-[11px] text-muted-foreground">
                        <th className="py-1 pr-3" />
                        <th className="py-1 pr-3 text-right">{t("rt.prevScaduto")}</th>
                        {mesi6.map((m) => (
                          <th key={m} className="py-1 pr-3 text-right whitespace-nowrap">
                            {fmtMese(m)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      <tr className="border-t border-border/40">
                        <td className="py-1 pr-3">{t("rt.prevIncassi")}</td>
                        <td className="py-1 pr-3 text-right tabular-nums text-status-absent">
                          {fmt(att.scaduto)}
                        </td>
                        {mesi6.map((m) => (
                          <td key={m} className="py-1 pr-3 text-right tabular-nums">
                            {fmt(att.out.get(m) ?? 0)}
                          </td>
                        ))}
                      </tr>
                      <tr className="border-t border-border/40">
                        <td className="py-1 pr-3">{t("rt.prevPagamenti")}</td>
                        <td className="py-1 pr-3 text-right tabular-nums text-status-absent">
                          {fmt(pas.scaduto)}
                        </td>
                        {mesi6.map((m) => (
                          <td key={m} className="py-1 pr-3 text-right tabular-nums">
                            {fmt(pas.out.get(m) ?? 0)}
                          </td>
                        ))}
                      </tr>
                      {haPref && (
                        <tr className="border-t border-border/40 italic text-primary">
                          <td className="py-1 pr-3">{t("rt.prevDaFattAtt")}</td>
                          <td className="py-1 pr-3 text-right">—</td>
                          {mesi6.map((m) => (
                            <td key={m} className="py-1 pr-3 text-right tabular-nums">
                              {fmt(prefAtt.get(m) ?? 0)}
                            </td>
                          ))}
                        </tr>
                      )}
                      {haPref && (
                        <tr className="border-t border-border/40 italic text-primary">
                          <td className="py-1 pr-3">{t("rt.prevDaFattPas")}</td>
                          <td className="py-1 pr-3 text-right">—</td>
                          {mesi6.map((m) => (
                            <td key={m} className="py-1 pr-3 text-right tabular-nums">
                              {fmt(prefPas.get(m) ?? 0)}
                            </td>
                          ))}
                        </tr>
                      )}
                      <tr className="border-t border-border/60 font-medium">
                        <td className="py-1 pr-3">{t("rt.prevSaldo")}</td>
                        <td className="py-1 pr-3 text-right tabular-nums">
                          {fmt(att.scaduto - pas.scaduto)}
                        </td>
                        {mesi6.map((m) => {
                          const v2 =
                            (att.out.get(m) ?? 0) -
                            (pas.out.get(m) ?? 0) +
                            (prefAtt.get(m) ?? 0) -
                            (prefPas.get(m) ?? 0);
                          return (
                            <td
                              key={m}
                              className={`py-1 pr-3 text-right tabular-nums ${v2 >= 0 ? "text-status-present" : "text-status-absent"}`}
                            >
                              {fmt(v2)}
                            </td>
                          );
                        })}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          {/* Gestione PREFATTURE: canoni ricorrenti e fatture pianificate.
              Quando la fattura vera arriva (stessa controparte, stesso
              mese) la riga si considera coperta da sola. */}
          <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
            <div className="mb-1 text-sm font-semibold text-foreground">{t("rt.prefTitolo")}</div>
            <p className="mb-3 text-xs text-muted-foreground">{t("rt.prefDesc")}</p>
            <div className="flex flex-wrap items-end gap-2">
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.cliForn")}</label>
                <input
                  value={pfControparte}
                  onChange={(e) => setPfControparte(e.target.value)}
                  className="w-44 rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("rt.prefDirezione")}</label>
                <select
                  value={pfDirezione}
                  onChange={(e) => setPfDirezione(e.target.value as "Emessa" | "Ricevuta")}
                  className="rounded-lg border border-border bg-background px-2 py-2 text-sm"
                >
                  <option value="Emessa">{t("rt.prefAttiva")}</option>
                  <option value="Ricevuta">{t("rt.prefPassiva")}</option>
                </select>
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("common.amount")}</label>
                <input
                  value={pfImporto}
                  onChange={(e) => setPfImporto(e.target.value)}
                  placeholder="1000,00"
                  className="w-28 rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("rt.prefMese")}</label>
                <input
                  type="month"
                  value={pfMese}
                  onChange={(e) => setPfMese(e.target.value)}
                  className="rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
              </div>
              <div>
                <label className="text-xs text-muted-foreground">{t("rt.prefRicorrenza")}</label>
                <select
                  value={pfRicorrenza}
                  onChange={(e) => setPfRicorrenza(e.target.value as "mensile" | "una")}
                  className="rounded-lg border border-border bg-background px-2 py-2 text-sm"
                >
                  <option value="mensile">{t("rt.prefMensile")}</option>
                  <option value="una">{t("rt.prefUna")}</option>
                </select>
              </div>
              {pfRicorrenza === "mensile" && (
                <div>
                  <label className="text-xs text-muted-foreground">{t("rt.prefMeseFine")}</label>
                  <input
                    type="month"
                    value={pfMeseFine}
                    onChange={(e) => setPfMeseFine(e.target.value)}
                    className="rounded-lg border border-border bg-background px-2 py-2 text-sm"
                  />
                </div>
              )}
              <div>
                <label className="text-xs text-muted-foreground">{t("fin.note")}</label>
                <input
                  value={pfNote}
                  onChange={(e) => setPfNote(e.target.value)}
                  className="w-40 rounded-lg border border-border bg-background px-2 py-2 text-sm"
                />
              </div>
              <button
                type="button"
                disabled={pfBusy || !pfControparte.trim() || !pfImporto.trim() || !pfMese}
                onClick={() => {
                  void (async () => {
                    setPfBusy(true);
                    try {
                      await spCreatePrefattura({
                        data: {
                          controparte: pfControparte,
                          direzione: pfDirezione,
                          importo: Number(pfImporto.replace(/\./g, "").replace(",", ".")),
                          meseInizio: pfMese,
                          ricorrenza: pfRicorrenza,
                          meseFine:
                            pfRicorrenza === "mensile" ? pfMeseFine || undefined : undefined,
                          note: pfNote || undefined,
                        },
                      });
                      const agg = (await spGetPrefatture()) as Prefattura[];
                      setPrefatture(agg);
                      setPfControparte("");
                      setPfImporto("");
                      setPfNote("");
                      toast.success(t("rt.prefCreata"));
                    } catch (err) {
                      toast.error(t("common.error"), {
                        description: err instanceof Error ? err.message : String(err),
                      });
                    } finally {
                      setPfBusy(false);
                    }
                  })();
                }}
                className="rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground disabled:opacity-50"
              >
                {pfBusy ? t("common.loading") : t("common.save")}
              </button>
            </div>
            {prefatture != null && prefatture.length > 0 && (
              <table className="mt-3 w-full text-[13px]">
                <tbody>
                  {prefatture.map((pf) => (
                    <tr key={pf.id} className="border-t border-border/40">
                      <td className="py-1 pr-3 font-medium">{pf.controparte}</td>
                      <td className="py-1 pr-3">
                        {pf.direzione === "Emessa" ? t("rt.prefAttiva") : t("rt.prefPassiva")}
                      </td>
                      <td className="py-1 pr-3 text-right tabular-nums">
                        {fmtImporto(pf.importo)} €
                      </td>
                      <td className="py-1 pr-3 text-muted-foreground">
                        {pf.ricorrenza === "una"
                          ? `${t("rt.prefUna")} ${fmtMese(pf.meseInizio)}`
                          : `${t("rt.prefMensile")} ${fmtMese(pf.meseInizio)} → ${pf.meseFine ? fmtMese(pf.meseFine) : "…"}`}
                      </td>
                      <td className="py-1 pr-3 text-[11px] text-muted-foreground max-w-40 truncate">
                        {pf.note ?? ""}
                      </td>
                      <td className="py-1 text-right">
                        <button
                          type="button"
                          onClick={() => {
                            void (async () => {
                              if (!window.confirm(t("rt.prefDelConfirm"))) return;
                              try {
                                await spDeletePrefattura({ data: { id: pf.id } });
                                setPrefatture((prev) => (prev ?? []).filter((x) => x.id !== pf.id));
                              } catch (err) {
                                toast.error(t("common.error"), {
                                  description: err instanceof Error ? err.message : String(err),
                                });
                              }
                            })();
                          }}
                          className="rounded-md p-1 text-muted-foreground hover:text-status-absent"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </>
      )}
    </div>
  );
}
