// DR Portal — Finanza → tab Fatture (solo direttore DR005 + admin).
// Scadenzario fatture emesse: import dall'export Aruba, stato incasso
// calcolato dagli abbinamenti coi movimenti bancari, ritardi per termini di
// pagamento cliente, riconciliazione automatica + abbinamento manuale.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import {
  AlertTriangle,
  Users,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  KeyRound,
  Link2,
  Loader2,
  Plug,
  Trash2,
  Upload,
  Wand2,
} from "lucide-react";
import { csvData, csvPeriodo, esportaCsvFile } from "@/lib/csv";
import { MultiSelect } from "@/components/MultiSelect";
import { useLang } from "@/lib/i18n";
import {
  computeStatoFattura,
  fattureEscluse,
  parseFattureMatrice,
  parseFatturaPA,
  proponiAbbinamenti,
  proponiAbbinamentiFIFO,
  isNotaCredito,
  isEsclusaDalCredito,
  collegaNoteCredito,
  parseMovimentiArubaMatrice,
  parseTerminiMatrice,
  giorniPerCliente,
  aggregaIncassiAruba,
  type MovimentoAruba,
  TERMINI_DEFAULT_GIORNI,
  TOLLERANZA_SALDO,
  SPIEGA_TOLLERANZA,
  spiegaBonifici,
  type PezzoSpiegazione,
  type SpiegazioneBonifico,
  type AbbinamentoIncasso,
  type DirezioneFattura,
  type FatturaRaw,
  type TerminePagamento,
  type StatoIncasso,
} from "@/lib/fatture-logic";
import { clienteGroupKey, normalizeTesto } from "@/lib/finanza-logic";
import {
  spGetFatture,
  spImportFatture,
  spGetTerminiPagamento,
  spGetAbbinamenti,
  spCreateAbbinamenti,
  spDeleteAbbinamento,
  spGetMovimenti,
  spSetIncassiAruba,
  spGetArubaStato,
  spSetArubaCredenziali,
  spArubaProvaConnessione,
  spSetRettificaNumero,
  spSetIncassoManuale,
  spTrovaFattureSenzaCliente,
  spImportTermini,
  spEliminaFatture,
} from "@/lib/sharepoint.functions";
import type { SpFattura, SpMovimento, ArubaStato } from "@/lib/sharepoint.server";
import type { ArubaProbeResult } from "@/lib/aruba.server";

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

// Chip on/off dei filtri multi-selezione (anni e stati).
const chipCls = (on: boolean) =>
  `rounded-full border px-3 py-1.5 text-xs font-medium transition-colors ${
    on
      ? "border-primary bg-primary text-primary-foreground"
      : "border-border bg-background text-foreground hover:bg-muted"
  }`;

function fmtData(iso?: string): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}
function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function csvNum(n: number): string {
  return (Math.round(n * 100) / 100).toString().replace(".", ",");
}

const CHUNK = 100;

type StatoFiltro =
  "tutte" | "ritardo" | "nonIncassata" | "parziale" | "pagata" | "discordante" | "nonGestita";

export function FattureTab({ direzione }: { direzione: DirezioneFattura }) {
  const { t } = useLang();
  // Attive (emesse, crediti) e passive (ricevute, debiti) vivono in due
  // schede distinte: la direzione arriva da lì, non si commuta qui dentro.
  const dir = direzione;
  const ricevute = dir === "Ricevuta";
  // Sulle passive non si incassa: si paga. Stesse colonne, altre parole: la
  // chiave delle attive resta il default, la variante scatta solo qui.
  const tp = (attiva: Parameters<typeof t>[0], passiva: Parameters<typeof t>[0]) =>
    t(ricevute ? passiva : attiva);
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const fatture = dir === "Emessa" ? fattureEm : fattureRic;
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [abbinamenti, setAbbinamenti] = useState<AbbinamentoIncasso[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);

  // Filtri MULTI-selezione (chip): lista vuota = nessun filtro. Le dimensioni
  // si incrociano (anni AND stati AND ricerca); dentro la stessa dimensione
  // le voci selezionate sono in OR (es. 2025+2026, "in ritardo"+"parziale").
  const [anniF, setAnniF] = useState<number[]>([new Date().getFullYear()]);
  const [clienteF, setClienteF] = useState("");
  const [statiF, setStatiF] = useState<Exclude<StatoFiltro, "tutte">[]>([]);
  // Gli scarti SdI sono nascosti ovunque: questo chip li isola, ed è l'unico
  // modo per vederli (serve a ricostruire la storia di una fattura rispedita).
  const [soloScartate, setSoloScartate] = useState(false);
  // Clienti selezionati nel menu a tendina (vuoto = tutti). Le voci proposte
  // sono i 15 clienti con piu' fatturato, in ordine decrescente.
  const [clientiSel, setClientiSel] = useState<string[]>([]);

  // Dettaglio espanso + abbinamento manuale
  const [openFile, setOpenFile] = useState<string | null>(null);
  const [abbMov, setAbbMov] = useState("");
  const [abbImporto, setAbbImporto] = useState("");
  const [abbBusy, setAbbBusy] = useState(false);

  // Riconciliazione automatica + a scalare (FIFO)
  const [reconciling, setReconciling] = useState(false);
  const [fifoBusy, setFifoBusy] = useState(false);

  // Import fatture: ZIP/XML FatturaPA (emesse E ricevute, direzione automatica
  // dalla P.IVA) oppure xlsx dell'export "Check fatture inviate" (emesse).
  const [showImport, setShowImport] = useState(false);
  const [impParsing, setImpParsing] = useState(false);
  const [impBusy, setImpBusy] = useState(false);
  const [previewImp, setPreviewImp] = useState<{
    descrizione: string;
    emesse: FatturaRaw[];
    ricevute: FatturaRaw[];
    scartate: number;
    duplicati: number;
  } | null>(null);

  // Collegamento Aruba (API, sola lettura)
  const [aruba, setAruba] = useState<ArubaStato | null>(null);
  const [showAruba, setShowAruba] = useState(false);
  const [arubaUser, setArubaUser] = useState("");
  const [arubaPass, setArubaPass] = useState("");
  const [arubaSaving, setArubaSaving] = useState(false);
  const [arubaTesting, setArubaTesting] = useState(false);
  const [probe, setProbe] = useState<ArubaProbeResult | null>(null);

  // Applica gli incassi del report movimenti alle fatture in archivio: gli
  // importi per rata sono l'unico dato che quantifica i PARZIALI.
  const applicaIncassiAruba = async (mov: MovimentoAruba[]) => {
    // Archivio non ancora caricato = confronto contro il vuoto: il messaggio
    // sarebbe "nessuna corrispondenza", che manda fuori strada. Meglio dire
    // la verita': aspetta che l'elenco compaia e riprova.
    if (fattureEm == null || fattureRic == null) {
      toast.error(t("ft.movAttendi"), { description: t("ft.movAttendiDesc") });
      return;
    }
    const tutte = [...(fattureEm ?? []), ...(fattureRic ?? [])];
    // Una rata non può finire su uno scarto SdI: senza questo filtro l'incasso
    // si attacca al tentativo scartato e la fattura buona resta "non incassata".
    const mappa = aggregaIncassiAruba(mov, tutte, true, fattureEscluse(tutte));
    if (mappa.size === 0) {
      toast.error(t("ft.movNessuna"), { description: t("ft.movNessunaDesc") });
      return;
    }
    const perFile = new Map(tutte.map((f) => [f.nomeFile, f]));
    // Ai REIMPORT la stragrande maggioranza delle fatture ha già lo stesso
    // incassato: quelle righe non si mandano proprio al server. È la
    // differenza fra rifare tutto e aggiornare solo ciò che è cambiato.
    const perDirezione = (d: DirezioneFattura) =>
      [...mappa.entries()]
        .filter(([file, v]) => {
          const f = perFile.get(file);
          if (!f || f.direzione !== d) return false;
          const stessaData = !v.ultimaData || v.ultimaData === (f.dataIncasso ?? "");
          return (f.incassatoAruba ?? null) !== v.incassato || !stessaData;
        })
        .map(([nomeFile, v]) => ({
          nomeFile,
          // L'id SharePoint evita al server di rileggere l'archivio a ogni
          // blocco: e' il percorso rapido di setIncassiAruba.
          id: perFile.get(nomeFile)?.id,
          incassato: v.incassato,
          ultimaData: v.ultimaData || undefined,
          rate: v.rate,
        }));
    const daFare = {
      Emessa: perDirezione("Emessa"),
      Ricevuta: perDirezione("Ricevuta"),
    };
    // GUARDIA CONTRO GLI EXPORT PARZIALI. La regola "niente rate nel file =
    // incasso a zero" presume che il file sia l'estrazione COMPLETA dell'anno.
    // Un export con un filtro date attivo contiene poche rate, ma basta una
    // rata su una fattura del 2025 per marcare "coperto" l'anno intero: senza
    // questa domanda, un file parziale azzera l'archivio (successo davvero:
    // 53 fatture iMile azzerate da un estratto di pochi giorni).
    const nAzzeramenti =
      daFare.Emessa.length + daFare.Ricevuta.length
        ? [...daFare.Emessa, ...daFare.Ricevuta].filter((r) => r.rate === 0 && r.incassato === 0)
            .length
        : 0;
    if (nAzzeramenti > 0) {
      const completo = window.confirm(`${nAzzeramenti} ${t("ft.movAzzeraConfirm")}`);
      if (!completo) {
        daFare.Emessa = daFare.Emessa.filter((r) => !(r.rate === 0 && r.incassato === 0));
        daFare.Ricevuta = daFare.Ricevuta.filter((r) => !(r.rate === 0 && r.incassato === 0));
      }
    }
    const invariate = mappa.size - daFare.Emessa.length - daFare.Ricevuta.length;
    // Blocchi grossi (il massimo che il server accetta): ogni chiamata
    // rilegge l'intero archivio, quindi meno chiamate = meno tempo totale.
    const PASSO = 200;
    const blocchiTot =
      Math.ceil(daFare.Emessa.length / PASSO) + Math.ceil(daFare.Ricevuta.length / PASSO);
    let aggiornate = 0;
    let blocco = 0;
    // Progresso visibile e timeout per blocco: un caricamento non deve mai
    // sembrare morto, né poterlo essere davvero. Se un blocco si pianta,
    // l'errore dice quanto è stato applicato: ricaricare lo stesso file
    // riparte da lì (le righe già applicate vengono saltate).
    const tid =
      blocchiTot > 0 ? toast.loading(`${t("ft.movProgress")} 0/${blocchiTot}`) : undefined;
    const conTimeout = <T,>(p: Promise<T>): Promise<T> =>
      Promise.race([
        p,
        // 5 minuti: un blocco su un archivio grosso puo' legittimamente
        // superare i 2 (e' successo: il taglio a meta' corsa lascia applicato
        // solo un prefisso del file, che sembra un export parziale).
        new Promise<never>((_, rej) =>
          setTimeout(() => rej(new Error(t("ft.movTimeout"))), 300000),
        ),
      ]);
    try {
      for (const d of ["Emessa", "Ricevuta"] as DirezioneFattura[]) {
        const righe = daFare[d];
        for (let i = 0; i < righe.length; i += PASSO) {
          const chiama = () =>
            conTimeout(
              spSetIncassiAruba({ data: { righe: righe.slice(i, i + PASSO), direzione: d } }),
            );
          // Un tentativo di riserva: un timeout isolato non deve buttare via
          // l'intero caricamento.
          const res = (await chiama().catch(chiama)) as { aggiornate: number; errori: string[] };
          aggiornate += res.aggiornate;
          blocco++;
          if (tid) toast.loading(`${t("ft.movProgress")} ${blocco}/${blocchiTot}`, { id: tid });
        }
      }
    } catch (err) {
      if (tid) toast.dismiss(tid);
      toast.error(t("common.error"), {
        description: `${err instanceof Error ? err.message : String(err)} — ${blocco}/${blocchiTot} ${t("ft.movRipresa")}`,
      });
      load();
      return;
    }
    if (tid) toast.dismiss(tid);
    setEsitoImport((prev) => ({
      nuove: prev?.nuove ?? 0,
      aggiornate: prev?.aggiornate ?? 0,
      doppioni: prev?.doppioni ?? 0,
      rate: mov.length,
      fattureIncassi: aggiornate + invariate,
      errori: prev?.errori ?? [],
    }));
    load();
  };

  const load = () => {
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
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    spGetAbbinamenti()
      .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
      .catch(() => setAbbinamenti([]));
    spGetMovimenti({ data: {} })
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
    spGetArubaStato()
      .then((s) => setAruba(s as ArubaStato))
      .catch(() => setAruba(null));
  };
  useEffect(load, []); // eslint-disable-line react-hooks/exhaustive-deps
  // Cambio scheda: si chiudono i dettagli aperti dell'altra direzione.
  useEffect(() => {
    setOpenFile(null);
    setClienteAperto(null);
  }, [direzione]);

  const oggiISO = new Date().toISOString().slice(0, 10);

  // Incassato per fattura dagli abbinamenti registrati.
  const incassatoPerFattura = useMemo(() => {
    const m = new Map<string, number>();
    for (const a of abbinamenti ?? [])
      m.set(a.fatturaFile, (m.get(a.fatturaFile) ?? 0) + a.importo);
    return m;
  }, [abbinamenti]);

  // Scarti SdI e reinvii: contabilmente non esistono. Restano in archivio come
  // storia, ma fuori da elenco, conteggi, abbinamenti e note di credito.
  const escluse = useMemo(() => fattureEscluse(fatture ?? []), [fatture]);

  // Note di credito collegate alle fatture che rettificano: il credito si
  // legge al netto, perché il cliente paga la differenza.
  const noteCredito = useMemo(() => collegaNoteCredito(fatture ?? [], escluse), [fatture, escluse]);

  // Fatture con stato calcolato.
  const conStato = useMemo(
    () =>
      (fatture ?? []).map((f) => {
        const s = computeStatoFattura(
          f,
          incassatoPerFattura.get(f.nomeFile) ?? 0,
          termini,
          oggiISO,
          noteCredito.get(f.nomeFile)?.importo ?? 0,
        );
        return escluse.has(f.nomeFile)
          ? {
              f,
              s: { ...s, stato: "NC" as const, residuo: 0, inRitardo: false, giorniRitardo: 0 },
              escluso: true,
            }
          : { f, s, escluso: false };
      }),
    [fatture, incassatoPerFattura, termini, oggiISO, escluse, noteCredito],
  );

  // Anni presenti nei dati (più l'anno corrente): sono i chip selezionabili.
  const anniDisponibili = useMemo(() => {
    const s = new Set<number>([new Date().getFullYear()]);
    for (const f of fatture ?? []) {
      const y = Number(f.dataDocumento.slice(0, 4));
      if (y) s.add(y);
    }
    return [...s].sort((a, b) => b - a);
  }, [fatture]);

  const matchAnno = (x: (typeof conStato)[number]) =>
    anniF.length === 0 || anniF.includes(Number(x.f.dataDocumento.slice(0, 4)));
  const matchStato = (x: (typeof conStato)[number]) =>
    statiF.length === 0 ||
    statiF.some((s) =>
      s === "ritardo"
        ? // Coerente col KPI: in ritardo solo se almeno una fonte esiste.
          x.s.inRitardo &&
          (x.s.statoIncassi != null || x.s.statoFatturazione != null || x.s.incassatoBanca > 0)
        : s === "nonIncassata"
          ? x.s.stato === "Non incassata"
          : s === "parziale"
            ? x.s.stato === "Parziale"
            : s === "pagata"
              ? x.s.stato === "Pagata"
              : s === "discordante"
                ? x.s.discordante
                : x.s.stato !== "NC" && !x.s.aruba,
    );

  const nScartate = useMemo(
    () => conStato.filter((x) => x.escluso && matchAnno(x)).length,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [conStato, anniF],
  );

  // TUTTI i clienti degli anni selezionati, in ordine di fatturato
  // DECRESCENTE (il direttore: "non prende tutti i clienti" — niente tagli).
  const clientiTop = useMemo(() => {
    const somme = new Map<string, { nome: string; tot: number }>();
    for (const x of conStato) {
      if (x.escluso || !matchAnno(x)) continue;
      const k = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const v = somme.get(k) ?? { nome: x.f.cliente, tot: 0 };
      v.tot += x.f.totale;
      somme.set(k, v);
    }
    return [...somme.entries()]
      .sort((a, b) => b[1].tot - a[1].tot)
      .map(([k, v]) => ({ v: k, label: v.nome }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [conStato, anniF]);

  const filtrate = useMemo(() => {
    // Gli scarti SdI non compaiono: una fattura scartata non è mai esistita.
    // Il chip "Scartate" le mostra da sole, per chi deve ricostruire la storia.
    let out = conStato.filter(
      (x) => x.escluso === soloScartate && matchAnno(x) && (soloScartate || matchStato(x)),
    );
    if (clientiSel.length)
      out = out.filter((x) => clientiSel.includes(clienteGroupKey(x.f.cliente) || x.f.cliente));
    if (clienteF.trim()) {
      const q = clienteF.trim().toLowerCase();
      out = out.filter(
        (x) => x.f.cliente.toLowerCase().includes(q) || x.f.numero.toLowerCase().includes(q),
      );
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [conStato, anniF, clienteF, statiF, soloScartate, clientiSel]);

  // Riepilogo (sugli anni filtrati, tutte le fatture non escluse). Gli importi
  // seguono lo stato UFFICIALE (Aruba quando c'è); `confermatoBanca` dice
  // quanta parte dell'incassato risulta anche dagli abbinamenti bancari.
  const riepilogo = useMemo(() => {
    const base = conStato.filter((x) => !x.escluso && matchAnno(x) && x.s.stato !== "NC");
    const residuo = base.reduce((s, x) => s + x.s.residuo, 0);
    const apertoFatt = base.reduce((s, x) => s + (x.s.residuoFatturazione ?? 0), 0);
    const apertoBanca = base.reduce((s, x) => s + x.s.residuoBanca, 0);
    const apertoIncassi = base.reduce((s, x) => s + (x.s.residuoIncassi ?? 0), 0);
    const incassatoIncassi = base.reduce((s, x) => s + (x.s.incassatoIncassi ?? 0), 0);
    const incassatoFatt = base.reduce((s, x) => s + (x.s.incassatoFatturazione ?? 0), 0);
    // "In ritardo" ha senso solo se ALMENO UNA fonte si è espressa: stato su
    // Aruba, rate nel report incassi, o un euro abbinato in banca. Le fatture
    // oltre scadenza di cui NESSUNA fonte sa niente (report di quell'anno mai
    // caricato, banca senza abbinamenti) sono con ogni probabilità incassate
    // in periodi non coperti dagli archivi: contarle a valore pieno gonfiava
    // il KPI fino all'assurdo (20M "in ritardo" su 2,5M di credito aperto).
    // Restano visibili, ma in un contatore a parte.
    const haFonte = (x: (typeof base)[number]) =>
      x.s.statoIncassi != null || x.s.statoFatturazione != null || x.s.incassatoBanca > 0;
    const inRitardo = base.filter((x) => x.s.inRitardo && haFonte(x));
    const ritardoImporto = inRitardo.reduce((s, x) => s + x.s.residuo, 0);
    const senzaFonte = base.filter((x) => x.s.inRitardo && !haFonte(x));
    const senzaFonteImporto = senzaFonte.reduce((s, x) => s + x.s.residuo, 0);
    const incassato = base.reduce((s, x) => s + x.s.incassato, 0);
    const confermatoBanca = base.reduce((s, x) => s + x.s.incassatoBanca, 0);
    const nDiscordanti = base.filter((x) => x.s.discordante).length;
    return {
      n: base.length,
      residuo,
      apertoFatt,
      apertoIncassi,
      apertoBanca,
      incassatoFatt,
      incassatoIncassi,
      nRitardo: inRitardo.length,
      ritardoImporto,
      nSenzaFonte: senzaFonte.length,
      senzaFonteImporto,
      incassato,
      confermatoBanca,
      nDiscordanti,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [conStato, anniF]);

  // Il SALDO DEI MOVIMENTI BANCARI in entrata (tipologia Incasso) per gli
  // anni filtrati: tutti gli accrediti arrivati, abbinati o no. È il secondo
  // dei "due saldi" chiesti dal direttore, da leggere accanto all'incassato
  // di fatturazione.
  const saldoMovimenti = useMemo(() => {
    let totale = 0;
    let n = 0;
    for (const m of movimenti ?? []) {
      if (m.importo <= 0) continue;
      if (!m.tipologia.toLowerCase().startsWith("incass")) continue;
      if (anniF.length && !anniF.includes(Number(m.dataContabile.slice(0, 4)))) continue;
      totale += m.importo;
      n++;
    }
    return { totale: Math.round(totale * 100) / 100, n };
  }, [movimenti, anniF]);

  // Riepilogo per cliente (come l'OVERVIEW del direttore, compattata).
  // Specchietto per cliente: conteggi e importi, ordinato per FATTURATO
  // decrescente. Le note di credito entrano nel fatturato (con segno) ma non
  // nei conteggi di stato, che riguardano le sole fatture.
  const perCliente = useMemo(() => {
    const m = new Map<
      string,
      {
        key: string;
        cliente: string;
        nFatture: number;
        nPagate: number;
        nParziali: number;
        nDaIncassare: number;
        nRitardo: number;
        fatturato: number;
        incassatoFatt: number;
        incassatoBanca: number;
        residuo: number;
        ritardo: number;
        /** Statistica del direttore: giorni medi data fattura → data incasso
         *  (sulle incassate), da confrontare col termine contrattuale. */
        sommaGg: number;
        nGg: number;
      }
    >();
    for (const x of conStato) {
      // Scarti e reinvii fuori da TUTTO: se restassero, il fatturato del
      // cliente conterebbe la stessa fattura una volta per tentativo.
      if (x.escluso || !matchAnno(x)) continue;
      const key = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const row = m.get(key) ?? {
        key,
        cliente: x.f.cliente,
        nFatture: 0,
        nPagate: 0,
        nParziali: 0,
        nDaIncassare: 0,
        nRitardo: 0,
        fatturato: 0,
        incassatoFatt: 0,
        incassatoBanca: 0,
        residuo: 0,
        ritardo: 0,
        sommaGg: 0,
        nGg: 0,
      };
      // Una nota di credito abbassa il fatturato; se e' COMPENSATA pesa in
      // NEGATIVO anche sull'incassato (denaro non entrato) — stessa regola
      // dell'export CSV, cosi' lo specchietto e i pivot del direttore dicono
      // lo stesso numero.
      row.fatturato += x.f.totale; // le NC arrivano già con segno negativo
      if (isNotaCredito(x.f.tipoDocumento)) {
        if (x.s.statoIncassi === "Pagata" || x.s.statoFatturazione === "Pagata")
          row.incassatoFatt -= Math.abs(x.f.totale);
      }
      if (!isNotaCredito(x.f.tipoDocumento)) {
        row.incassatoFatt += x.s.incassatoIncassi ?? x.s.incassatoFatturazione ?? 0;
        row.incassatoBanca += x.s.incassatoBanca;
        // "anche il pagato dev'essere contato": la media dei giorni di
        // incasso si calcola su TUTTE le fatture con data incasso nota.
        if (x.f.dataIncasso && x.f.dataDocumento) {
          const gg =
            (new Date(`${x.f.dataIncasso}T00:00:00`).getTime() -
              new Date(`${x.f.dataDocumento}T00:00:00`).getTime()) /
            86400000;
          if (Number.isFinite(gg) && gg >= 0) {
            row.sommaGg += gg;
            row.nGg++;
          }
        }
      }
      if (x.s.stato !== "NC") {
        row.nFatture++;
        if (x.s.annullataDaNC || x.s.stato === "Pagata") row.nPagate++;
        else if (x.s.stato === "Parziale") row.nParziali++;
        else row.nDaIncassare++;
        if (x.s.inRitardo) {
          row.nRitardo++;
          row.ritardo += x.s.residuo;
        }
      }
      m.set(key, row);
    }
    // Da incassare = fatturato − incassato (fatturazione). È la lettura che
    // quadra con le due colonne accanto; la banca resta fuori dal conteggio.
    for (const row of m.values())
      row.residuo = Math.round((row.fatturato - row.incassatoFatt) * 100) / 100;
    // Ordinato per fatturato; se nel menu clienti c'e' una selezione, lo
    // specchietto la segue.
    const tutte = [...m.values()].sort((a, b) => b.fatturato - a.fatturato);
    return clientiSel.length ? tutte.filter((r) => clientiSel.includes(r.key)) : tutte;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [conStato, anniF, clientiSel]);

  // Riepilogo dell'ultimo import: resta a schermo finché non lo si chiude,
  // così i numeri si leggono con calma (un avviso a scomparsa non basta).
  const [esitoImport, setEsitoImport] = useState<{
    nuove: number;
    aggiornate: number;
    doppioni: number;
    rate: number;
    fattureIncassi: number;
    errori: string[];
  } | null>(null);

  // Cliente espanso nello specchietto (mostra le sue fatture). Di default si
  // vedono SOLO le fatture ancora da incassare (intere o parziali): è la
  // domanda per cui si apre il dettaglio. "Tutte" per lo storico completo.
  const [clienteAperto, setClienteAperto] = useState<string | null>(null);
  const [dettaglioTutte, setDettaglioTutte] = useState(false);
  const fattureDelCliente = useMemo(
    () =>
      clienteAperto == null
        ? []
        : conStato
            .filter(
              (x) =>
                !x.escluso &&
                matchAnno(x) &&
                (clienteGroupKey(x.f.cliente) || x.f.cliente) === clienteAperto,
            )
            .sort((a, b) => b.f.dataDocumento.localeCompare(a.f.dataDocumento)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [conStato, clienteAperto, anniF],
  );

  // Le sole APERTE (da incassare, intere o parziali): il default del dettaglio.
  const fattureDelClienteAperte = useMemo(
    () =>
      fattureDelCliente.filter(
        (x) =>
          !isNotaCredito(x.f.tipoDocumento) &&
          x.s.stato !== "NC" &&
          !x.s.annullataDaNC &&
          x.s.residuo > TOLLERANZA_SALDO,
      ),
    [fattureDelCliente],
  );

  // ESTRATTO CONTO del cliente aperto: la quadratura completa, fattura per
  // fattura e movimento per movimento. Serve a rispondere alla sola domanda
  // che conta ("quanto deve ancora pagare davvero") senza sorprese: i
  // movimenti bancari del cliente NON agganciati ad alcuna fattura sono la
  // ragione tipica per cui i totali non tornano, e qui si vedono.
  const estrattoCliente = useMemo(() => {
    if (clienteAperto == null) return null;
    const numeri = new Set(fattureDelCliente.map((x) => normalizeTesto(x.f.numero)));
    // Un movimento è del cliente se la controparte estratta coincide, oppure
    // se cita il numero di una sua fattura (controparte scritta diversamente).
    const suoi = (movimenti ?? []).filter(
      (m) =>
        (clienteGroupKey(m.cliente) || m.cliente) === clienteAperto ||
        (!!m.nrFattura && numeri.has(normalizeTesto(m.nrFattura))),
    );
    // Sulle attive contano gli incassi (entrate), sulle passive le uscite.
    const rilevanti = suoi.filter((m) => (ricevute ? m.importo < 0 : m.importo > 0));
    const allocatoPerMov = new Map<string, number>();
    for (const a of abbinamenti ?? [])
      allocatoPerMov.set(
        a.movimentoChiave,
        (allocatoPerMov.get(a.movimentoChiave) ?? 0) + a.importo,
      );
    const righe = rilevanti
      .map((m) => {
        const allocato = allocatoPerMov.get(m.chiave) ?? 0;
        const fattureAbb = (abbinamenti ?? [])
          .filter((a) => a.movimentoChiave === m.chiave)
          .map((a) => fattureDelCliente.find((x) => x.f.nomeFile === a.fatturaFile)?.f.numero)
          .filter(Boolean) as string[];
        return {
          m,
          allocato,
          residuo: Math.round((Math.abs(m.importo) - allocato) * 100) / 100,
          fattureAbb,
        };
      })
      .sort((a, b) => b.m.dataContabile.localeCompare(a.m.dataContabile));
    const fatturato = fattureDelCliente.reduce((s, x) => s + x.f.totale, 0);
    const incassatoFatt = fattureDelCliente.reduce(
      (s, x) =>
        s +
        (isNotaCredito(x.f.tipoDocumento)
          ? 0
          : (x.s.incassatoIncassi ?? x.s.incassatoFatturazione ?? 0)),
      0,
    );
    const incassatoBanca = fattureDelCliente.reduce(
      (s, x) => s + (isNotaCredito(x.f.tipoDocumento) ? 0 : x.s.incassatoBanca),
      0,
    );
    const totaleBanca = righe.reduce((s, r) => s + Math.abs(r.m.importo), 0);
    const nonAttribuito = righe.reduce((s, r) => s + r.residuo, 0);
    // LETTURA EFFETTIVA, dalla banca: il fatturato netto meno gli euro
    // realmente arrivati. Non dipende dagli stati di Aruba (che possono
    // essere indietro) né dagli abbinamenti: è il numero del direttore.
    // Le fatture della CONTROPARTE verso di noi si sottraggono dal dovuto:
    // clienti come iMile le trattengono dentro i bonifici.
    const altreDirezione = dir === "Emessa" ? (fattureRic ?? []) : (fattureEm ?? []);
    const controFatture = altreDirezione
      .filter(
        (f) =>
          (clienteGroupKey(f.cliente) || f.cliente) === clienteAperto &&
          !isNotaCredito(f.tipoDocumento) &&
          f.totale > 0,
      )
      .reduce((s, f) => s + f.totale, 0);
    const nettoDovuto = Math.round((fatturato - controFatture) * 100) / 100;
    return {
      controFatture: Math.round(controFatture * 100) / 100,
      nettoDovuto,
      residuoEffettivo: Math.round((nettoDovuto - totaleBanca) * 100) / 100,
      righe,
      fatturato: Math.round(fatturato * 100) / 100,
      incassatoFatt: Math.round(incassatoFatt * 100) / 100,
      incassatoBanca: Math.round(incassatoBanca * 100) / 100,
      residuo: Math.round((fatturato - incassatoFatt) * 100) / 100,
      totaleBanca: Math.round(totaleBanca * 100) / 100,
      nonAttribuito: Math.round(nonAttribuito * 100) / 100,
      nNonAttribuiti: righe.filter((r) => r.residuo > TOLLERANZA_SALDO).length,
    };
  }, [
    clienteAperto,
    fattureDelCliente,
    movimenti,
    abbinamenti,
    ricevute,
    dir,
    fattureEm,
    fattureRic,
  ]);

  // --- Spiegazione bonifici (compensazioni incluse) -------------------------
  // Il pulsante "Spiega bonifici" cerca, per ogni movimento non attribuito del
  // cliente aperto, la combinazione di documenti che lo spiega al centesimo:
  // fatture, note di credito non collegate, perfino fatture della controparte
  // verso di noi. Le proposte restano tali finché non si preme Applica.
  const [spiegazioni, setSpiegazioni] = useState<SpiegazioneBonifico[] | null>(null);
  const [spiegaBusy, setSpiegaBusy] = useState(false);
  useEffect(() => setSpiegazioni(null), [clienteAperto, direzione]);

  const calcolaSpiegazioni = () => {
    if (!estrattoCliente) return;
    const pezzi: PezzoSpiegazione[] = [];
    for (const x of fattureDelCliente) {
      if (isNotaCredito(x.f.tipoDocumento)) {
        // Solo le NC NON ancora collegate a una fattura: quelle collegate
        // abbattono già il credito della fattura e non vanno contate due volte.
        if (!x.f.rettificaNumero)
          pezzi.push({
            chiave: x.f.nomeFile,
            etichetta: x.f.numero,
            tipo: "NC",
            valore: -Math.abs(x.f.totale),
            totale: -Math.abs(x.f.totale),
            data: x.f.dataDocumento,
          });
        continue;
      }
      // Valore atteso in banca: l'incassato registrato su Aruba quando c'è
      // (è al netto delle trattenute), altrimenti il totale al netto delle NC
      // collegate; meno quanto la banca ha già abbinato.
      const base = Math.max(0, x.f.totale - (noteCredito.get(x.f.nomeFile)?.importo ?? 0));
      const atteso =
        (x.s.incassatoIncassi && x.s.incassatoIncassi > 0 ? x.s.incassatoIncassi : base) -
        x.s.incassatoBanca;
      if (atteso > SPIEGA_TOLLERANZA)
        pezzi.push({
          chiave: x.f.nomeFile,
          etichetta: x.f.numero,
          tipo: "FT",
          valore: Math.round(atteso * 100) / 100,
          totale: base,
          data: x.f.dataDocumento,
        });
    }
    // Fatture della controparte verso di noi (l'altra direzione): i clienti
    // come iMile le compensano dentro i loro bonifici.
    const altre = dir === "Emessa" ? (fattureRic ?? []) : (fattureEm ?? []);
    for (const f of altre) {
      if ((clienteGroupKey(f.cliente) || f.cliente) !== clienteAperto) continue;
      if (isNotaCredito(f.tipoDocumento) || f.totale <= 0) continue;
      pezzi.push({
        chiave: f.nomeFile,
        etichetta: f.numero,
        tipo: "CONTRO",
        valore: -f.totale,
        totale: -f.totale,
        data: f.dataDocumento,
      });
    }
    const daSpiegare = estrattoCliente.righe
      .filter((r) => r.residuo > SPIEGA_TOLLERANZA)
      .map((r) => ({ chiave: r.m.chiave, dataContabile: r.m.dataContabile, importo: r.residuo }));
    setSpiegazioni(spiegaBonifici(daSpiegare, pezzi));
  };

  // Applica UNA spiegazione: abbinamenti per le fatture, collegamento
  // RettificaNumero per le note di credito (alla fattura più grande della
  // stessa combinazione). Le fatture della controparte restano informative.
  const applicaSpiegazione = async (s: SpiegazioneBonifico) => {
    setSpiegaBusy(true);
    try {
      const ftPezzi = s.pezzi.filter((p) => p.tipo === "FT");
      const ncPezzi = s.pezzi.filter((p) => p.tipo === "NC");
      if (ftPezzi.length) {
        await spCreateAbbinamenti({
          data: {
            rows: ftPezzi.map((p) => ({
              fatturaFile: p.chiave,
              movimentoChiave: s.movimentoChiave,
              importo: p.valore,
              origine: "Auto" as const,
            })),
          },
        });
      }
      if (ncPezzi.length && ftPezzi.length) {
        const grande = [...ftPezzi].sort((a, b) => b.valore - a.valore)[0];
        for (const nc of ncPezzi)
          await spSetRettificaNumero({
            data: { nomeFile: nc.chiave, numeroFattura: grande.etichetta, direzione: dir },
          });
      }
      // Ricarico tutto: abbinamenti, e fatture se ho toccato le NC.
      const [abb, agg] = await Promise.all([
        spGetAbbinamenti() as Promise<AbbinamentoIncasso[]>,
        ncPezzi.length
          ? (spGetFatture({ data: { direzione: dir } }) as Promise<SpFattura[]>)
          : Promise.resolve(null),
      ]);
      setAbbinamenti(abb);
      if (agg) (dir === "Emessa" ? setFattureEm : setFattureRic)(agg);
      setSpiegazioni((prev) =>
        prev ? prev.filter((x) => x.movimentoChiave !== s.movimentoChiave) : prev,
      );
      toast.success(t("ft.spiegaApplicata"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSpiegaBusy(false);
    }
  };

  // Estratto conto del cliente in CSV: fatture e movimenti nello stesso file,
  // distinti dalla colonna Tipo, così la quadratura si rifà in Excel.
  const esportaEstratto = () => {
    if (!estrattoCliente || clienteAperto == null) return;
    const nome = fattureDelCliente[0]?.f.cliente ?? clienteAperto;
    esportaCsvFile(
      `estratto-${nome.replace(/[^\w]+/g, "-").toLowerCase()}`,
      [
        "Tipo",
        "Data",
        "Anno",
        "Trimestre",
        "Mese",
        "Numero",
        "Importo",
        "Scadenza",
        ricevute ? "Pagato fatturazione" : "Incassato fatturazione",
        ricevute ? "Pagato banca" : "Incassato banca",
        "Residuo",
        "Stato fatturazione",
        "Abbinato a",
        "Descrizione",
      ],
      [
        ...fattureDelCliente.map((x) => [
          isNotaCredito(x.f.tipoDocumento) ? "Nota di credito" : "Fattura",
          csvData(x.f.dataDocumento),
          ...csvPeriodo(x.f.dataDocumento),
          x.f.numero,
          csvNum(x.f.totale),
          csvData(x.s.scadenza),
          // NC compensata = incasso NEGATIVO (vedi export fatture): la somma
          // della colonna deve avvicinarsi alla cassa, non al lordo.
          csvNum(
            isNotaCredito(x.f.tipoDocumento)
              ? x.s.statoIncassi === "Pagata" || x.s.statoFatturazione === "Pagata"
                ? -Math.abs(x.f.totale)
                : 0
              : (x.s.incassatoIncassi ?? x.s.incassatoFatturazione ?? 0),
          ),
          csvNum(x.s.incassatoBanca),
          csvNum(x.s.stato === "NC" ? 0 : x.s.residuo),
          x.s.statoFatturazione ?? "non gestita",
          "",
          x.f.tipoDocumento,
        ]),
        ...estrattoCliente.righe.map((r) => [
          "Movimento bancario",
          csvData(r.m.dataContabile),
          ...csvPeriodo(r.m.dataContabile),
          r.m.nrFattura,
          csvNum(Math.abs(r.m.importo)),
          "",
          "",
          csvNum(r.allocato),
          csvNum(r.residuo),
          "",
          r.fattureAbb.join(" "),
          r.m.descrizione,
        ]),
      ],
    );
  };

  // --- Pulizia documenti senza controparte -----------------------------------
  // Un file letto col tracciato sbagliato (l'xlsx delle ricevute importato
  // come emesse) lascia righe senza cliente che gonfiano archivio e totali,
  // fino a impedire il caricamento della pagina. Qui si contano, si mostrano
  // esempi, si chiede conferma e si eliminano a blocchi.
  const [puliziaBusy, setPuliziaBusy] = useState(false);
  const pulisciSenzaCliente = async () => {
    setPuliziaBusy(true);
    try {
      const { ids, esempi } = (await spTrovaFattureSenzaCliente({
        data: { direzione: dir },
      })) as { ids: string[]; esempi: string[] };
      if (ids.length === 0) {
        toast.success(t("ft.puliziaNiente"));
        return;
      }
      const okDel = window.confirm(
        [`${ids.length} ${t("ft.puliziaConfirm")}`, "", ...esempi].join("\n"),
      );
      if (!okDel) return;
      const tid = toast.loading(`${t("ft.puliziaProgress")} 0/${ids.length}`);
      let eliminate = 0;
      try {
        for (let i = 0; i < ids.length; i += 80) {
          const res = (await spEliminaFatture({
            data: { ids: ids.slice(i, i + 80), direzione: dir },
          })) as { eliminate: number; errori: string[] };
          eliminate += res.eliminate;
          toast.loading(`${t("ft.puliziaProgress")} ${eliminate}/${ids.length}`, { id: tid });
        }
      } finally {
        toast.dismiss(tid);
      }
      toast.success(`${eliminate} ${t("ft.puliziaFatta")}`);
      load();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setPuliziaBusy(false);
    }
  };

  // --- Stato d'incasso corretto a mano ---------------------------------------
  // Per la fattura che si SA essere stata incassata (già registrata su Aruba)
  // senza aspettare il prossimo report: scrive sulle stesse colonne
  // dell'import, quindi tutte le viste si aggiornano e il report successivo
  // — che ormai dice lo stesso — conferma senza toccare nulla.
  const [incFile, setIncFile] = useState<string | null>(null);
  const [incStato, setIncStato] = useState<"Incassata" | "Non incassata">("Incassata");
  const [incData, setIncData] = useState("");
  const [incSaving, setIncSaving] = useState(false);

  const salvaIncassoManuale = async (f: FatturaRaw) => {
    setIncSaving(true);
    try {
      await spSetIncassoManuale({
        data: {
          nomeFile: f.nomeFile,
          stato: incStato,
          direzione: dir,
          dataIncasso: incStato === "Incassata" ? incData || oggiISO : undefined,
        },
      });
      const agg = (await spGetFatture({ data: { direzione: dir } })) as SpFattura[];
      (dir === "Emessa" ? setFattureEm : setFattureRic)(agg);
      setIncFile(null);
      toast.success(t("ft.incManOk"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setIncSaving(false);
    }
  };

  // --- Collegamento manuale NOTA DI CREDITO → fattura -----------------------
  // Lo storno fatto dentro Aruba non lascia traccia nell'XML: il riferimento
  // si mette qui e finisce su RettificaNumero, la stessa colonna dell'import,
  // quindi vale da subito e sopravvive ai reimport.
  const [ncFile, setNcFile] = useState<string | null>(null);
  const [ncNumero, setNcNumero] = useState("");
  const [ncSaving, setNcSaving] = useState(false);

  // Le fatture agganciabili a una NC: stessa controparte e stessa direzione,
  // niente note di credito né scarti. Le più recenti per prime.
  const fattureCollegabili = (nc: FatturaRaw) => {
    const key = clienteGroupKey(nc.cliente) || nc.cliente;
    return (fatture ?? [])
      .filter(
        (f) =>
          !isNotaCredito(f.tipoDocumento) &&
          !escluse.has(f.nomeFile) &&
          f.totale > 0 &&
          (clienteGroupKey(f.cliente) || f.cliente) === key,
      )
      .sort((a, b) => b.dataDocumento.localeCompare(a.dataDocumento));
  };

  const salvaRettifica = async (nc: FatturaRaw) => {
    setNcSaving(true);
    try {
      await spSetRettificaNumero({
        data: { nomeFile: nc.nomeFile, numeroFattura: ncNumero, direzione: dir },
      });
      // Ricarico l'archivio: il netto del cliente cambia subito.
      const agg = (await spGetFatture({ data: { direzione: dir } })) as SpFattura[];
      if (dir === "Emessa") setFattureEm(agg);
      else setFattureRic(agg);
      setNcFile(null);
      setNcNumero("");
      toast.success(ncNumero ? t("ft.collegataOk") : t("ft.scollegataOk"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setNcSaving(false);
    }
  };

  // Movimenti con residuo da allocare per l'abbinamento manuale: incassi per
  // le emesse, uscite (in valore assoluto) per le ricevute.
  const incassiDisponibili = useMemo(() => {
    const allocato = new Map<string, number>();
    for (const a of abbinamenti ?? [])
      allocato.set(a.movimentoChiave, (allocato.get(a.movimentoChiave) ?? 0) + a.importo);
    return (movimenti ?? [])
      .filter((m) =>
        dir === "Emessa" ? m.importo > 0 && m.tipologia === "Incasso" : m.importo < 0,
      )
      .map((m) => ({
        m,
        residuo: Math.round((Math.abs(m.importo) - (allocato.get(m.chiave) ?? 0)) * 100) / 100,
      }))
      .filter((x) => x.residuo > 0.01)
      .sort((a, b) => b.m.dataContabile.localeCompare(a.m.dataContabile));
  }, [movimenti, abbinamenti, dir]);

  // --- Riconciliazione automatica ------------------------------------------
  const riconcilia = async () => {
    if (!fatture || !movimenti || !abbinamenti) return;
    setReconciling(true);
    try {
      const proposte = proponiAbbinamenti(
        fatture.filter((f) => !escluse.has(f.nomeFile)),
        movimenti.map((m) => ({
          chiave: m.chiave,
          dataContabile: m.dataContabile,
          importo: m.importo,
          tipologia: m.tipologia,
          cliente: m.cliente,
          descrizione: m.descrizione,
          nrFattura: m.nrFattura,
        })),
        abbinamenti,
        dir,
      );
      if (proposte.length === 0) {
        toast(t("ft.reconcileNone"));
        return;
      }
      let creati = 0;
      for (let i = 0; i < proposte.length; i += CHUNK) {
        const res = (await spCreateAbbinamenti({
          data: {
            rows: proposte.slice(i, i + CHUNK).map(({ motivo: _m, ...r }) => r),
          },
        })) as { creati: number };
        creati += res.creati;
      }
      toast.success(t("ft.reconcileDone"), {
        description: `${creati} ${t("ft.reconcileCount")}`,
      });
      spGetAbbinamenti()
        .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setReconciling(false);
    }
  };

  // --- Abbinamento manuale --------------------------------------------------
  const abbinaManuale = async (fatturaFile: string, residuoFattura: number) => {
    const inc = incassiDisponibili.find((x) => x.m.chiave === abbMov);
    if (!inc) return toast.error(t("ft.abbSelect"));
    const importo = Number(abbImporto.replace(",", "."));
    if (!Number.isFinite(importo) || importo <= 0) return toast.error(t("ft.abbImporto"));
    if (importo > inc.residuo + 0.01 || importo > residuoFattura + 0.01)
      return toast.error(t("ft.abbTroppo"));
    setAbbBusy(true);
    try {
      await spCreateAbbinamenti({
        data: {
          rows: [{ fatturaFile, movimentoChiave: inc.m.chiave, importo, origine: "Manuale" }],
        },
      });
      toast.success(t("ft.abbDone"));
      setAbbMov("");
      setAbbImporto("");
      spGetAbbinamenti()
        .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setAbbBusy(false);
    }
  };

  const rimuoviAbbinamento = async (a: AbbinamentoIncasso) => {
    if (!a.id) return;
    if (!window.confirm(t("ft.abbDeleteConfirm"))) return;
    try {
      await spDeleteAbbinamento({ data: { abbinamentoId: a.id } });
      setAbbinamenti((prev) => (prev ?? []).filter((x) => x.id !== a.id));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    }
  };

  // Riconciliazione A SCALARE (FIFO): imputa gli incassi residui alle fatture
  // aperte più vecchie del cliente (acconti/saldi mensili/compensazioni non
  // coincidono mai con le singole fatture). Azione esplicita e confermata.
  const riconciliaFIFO = async () => {
    if (!fatture || !movimenti || !abbinamenti) return;
    setFifoBusy(true);
    try {
      const proposte = proponiAbbinamentiFIFO(
        fatture.filter((f) => !escluse.has(f.nomeFile)),
        movimenti.map((m) => ({
          chiave: m.chiave,
          dataContabile: m.dataContabile,
          importo: m.importo,
          tipologia: m.tipologia,
          cliente: m.cliente,
          descrizione: m.descrizione,
          nrFattura: m.nrFattura,
        })),
        abbinamenti,
        dir,
      );
      if (proposte.length === 0) {
        toast(t("ft.fifoNone"));
        return;
      }
      const totale = proposte.reduce((s, p) => s + p.importo, 0);
      if (
        !window.confirm(
          `${t("ft.fifoConfirm")}\n${proposte.length} ${t("ft.reconcileCount")} · ${fmtImporto(totale)} €`,
        )
      )
        return;
      let creati = 0;
      for (let i = 0; i < proposte.length; i += CHUNK) {
        const res = (await spCreateAbbinamenti({
          data: { rows: proposte.slice(i, i + CHUNK).map(({ motivo: _m, ...r }) => r) },
        })) as { creati: number };
        creati += res.creati;
      }
      toast.success(t("ft.reconcileDone"), { description: `${creati} ${t("ft.reconcileCount")}` });
      spGetAbbinamenti()
        .then((l) => setAbbinamenti(l as AbbinamentoIncasso[]))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setFifoBusy(false);
    }
  };

  // --- Import fatture (ZIP/XML FatturaPA + xlsx emesse) ---------------------
  const onFiles = async (files: File[]) => {
    setPreviewImp(null);
    setImpParsing(true);
    try {
      const rows: FatturaRaw[] = [];
      let scartate = 0;
      let terminiCaricati = false;
      // Report MOVIMENTI di Aruba: rate incassate per fattura (i parziali).
      const movAruba: MovimentoAruba[] = [];
      const decoder = new TextDecoder("utf-8");
      const daXml = (testo: string, nome: string) => {
        const res = parseFatturaPA(testo, nome);
        rows.push(...res.rows);
        scartate += res.scartati.length;
      };
      for (const file of files) {
        const nome = file.name;
        if (/\.zip$/i.test(nome)) {
          const { unzipSync } = await import("fflate");
          const entries = unzipSync(new Uint8Array(await file.arrayBuffer()));
          for (const [entry, bytes] of Object.entries(entries)) {
            if (!/\.xml$/i.test(entry)) continue;
            daXml(decoder.decode(bytes), entry.split("/").pop() ?? entry);
          }
        } else if (/\.xml$/i.test(nome)) {
          daXml(await file.text(), nome);
        } else if (/\.xlsx?$/i.test(nome)) {
          const XLSX = await import("xlsx");
          const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
          let trovato = false;
          for (const sheet of wb.SheetNames) {
            const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
              header: 1,
              raw: true,
            }) as unknown[][];
            const mov = parseMovimentiArubaMatrice(matrix);
            if (mov) {
              movAruba.push(...mov);
              trovato = true;
              break;
            }
            const res = parseFattureMatrice(matrix);
            if (res && res.rows.length) {
              rows.push(...res.rows);
              scartate += res.scartate;
              trovato = true;
              break;
            }
            // Foglio contratti del direttore: colonna "GG pagamento" =
            // termini per cliente. Si importano subito (upsert) e le
            // scadenze/ritardi si ricalcolano da soli.
            const term = parseTerminiMatrice(matrix);
            if (term) {
              terminiCaricati = true;
              const esito = (await spImportTermini({ data: { rows: term } })) as {
                nuovi: number;
                aggiornati: number;
                invariati: number;
              };
              const agg = (await spGetTerminiPagamento()) as TerminePagamento[];
              setTermini(agg);
              toast.success(t("ft.terminiOk"), {
                description: `${esito.nuovi} ${t("ft.importNew")} · ${esito.aggiornati} ${t("ft.importUpd")} · ${esito.invariati} ${t("ft.terminiUguali")}`,
              });
              trovato = true;
              break;
            }
          }
          if (!trovato) scartate++;
        } else {
          scartate++;
        }
      }
      // Solo report movimenti: si applicano subito gli incassi alle fatture
      // già in archivio, senza passare dall'anteprima di import.
      if (rows.length === 0 && movAruba.length > 0) {
        await applicaIncassiAruba(movAruba);
        return;
      }
      // Solo foglio contratti: l'import dei termini e' gia' avvenuto sopra.
      if (rows.length === 0 && terminiCaricati) return;
      if (rows.length === 0) {
        toast.error(t("ft.errFile"), { description: t("ft.errFileDesc") });
        return;
      }
      if (movAruba.length > 0) await applicaIncassiAruba(movAruba);
      // Dedup nel caricamento stesso: stesso documento in due file scelti
      // insieme (es. lo stesso ZIP scaricato due volte, o ZIP + xlsx). NON
      // sono file illeggibili: vanno contati a parte, o il messaggio spaventa.
      const visti = new Set<string>();
      let duplicati = 0;
      const univoche = rows.filter((r) => {
        if (visti.has(r.nomeFile)) {
          duplicati++;
          return false;
        }
        visti.add(r.nomeFile);
        return true;
      });
      setPreviewImp({
        descrizione: files.length === 1 ? files[0].name : `${files.length} file`,
        emesse: univoche.filter((r) => r.direzione === "Emessa"),
        ricevute: univoche.filter((r) => r.direzione === "Ricevuta"),
        scartate,
        duplicati,
      });
    } catch (err) {
      toast.error(t("ft.errFile"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImpParsing(false);
    }
  };

  const eseguiImport = async () => {
    if (!previewImp) return;
    setImpBusy(true);
    try {
      let nuove = 0;
      let aggiornate = 0;
      let doppioni = 0;
      const errori: string[] = [];
      const importaGruppo = async (gruppo: FatturaRaw[], direzione: DirezioneFattura) => {
        for (let i = 0; i < gruppo.length; i += CHUNK) {
          const res = (await spImportFatture({
            data: { rows: gruppo.slice(i, i + CHUNK), direzione },
          })) as { importate: number; aggiornate: number; doppioni: number; errori: string[] };
          nuove += res.importate;
          aggiornate += res.aggiornate;
          doppioni += res.doppioni;
          errori.push(...res.errori);
        }
      };
      if (previewImp.emesse.length) await importaGruppo(previewImp.emesse, "Emessa");
      if (previewImp.ricevute.length) await importaGruppo(previewImp.ricevute, "Ricevuta");
      setEsitoImport((prev) => ({
        nuove,
        aggiornate,
        doppioni,
        rate: prev?.rate ?? 0,
        fattureIncassi: prev?.fattureIncassi ?? 0,
        errori,
      }));
      toast.success(t("ft.importDone"), {
        description: `${nuove} ${t("ft.importNew")} · ${aggiornate} ${t("ft.importUpd")} · ${doppioni} ${t("ft.importDup")}${errori.length ? ` · ${errori.length} ${t("common.error").toLowerCase()}` : ""}`,
      });
      setPreviewImp(null);
      setShowImport(false);
      load();
    } catch (err) {
      toast.error(t("ft.errImport"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setImpBusy(false);
    }
  };

  // --- Collegamento Aruba ---------------------------------------------------
  const salvaCredenziali = async () => {
    if (!arubaUser.trim() || !arubaPass) return toast.error(t("ft.arCredMancanti"));
    setArubaSaving(true);
    try {
      await spSetArubaCredenziali({ data: { username: arubaUser.trim(), password: arubaPass } });
      toast.success(t("ft.arCredSalvate"));
      setArubaPass("");
      spGetArubaStato()
        .then((s) => setAruba(s as ArubaStato))
        .catch(() => {});
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setArubaSaving(false);
    }
  };

  const provaConnessione = async () => {
    setArubaTesting(true);
    setProbe(null);
    try {
      const res = (await spArubaProvaConnessione()) as ArubaProbeResult;
      setProbe(res);
      toast.success(t("ft.arProvaOk"));
    } catch (err) {
      setProbe({
        ok: false,
        messaggio: err instanceof Error ? err.message : String(err),
      });
      toast.error(t("ft.arProvaKo"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setArubaTesting(false);
    }
  };

  const esporta = () => {
    esportaCsvFile(
      `fatture-${dir === "Ricevuta" ? "ricevute" : "emesse"}-${anniF.length ? [...anniF].sort().join("-") : "tutte"}`,
      [
        "Numero",
        "Data documento",
        // Anno/trimestre/mese gia' pronti: la pivot raggruppa senza dover
        // convertire la colonna data a mano.
        "Anno",
        "Trimestre",
        "Mese",
        ricevute ? "Fornitore" : "Cliente",
        "Tipo",
        "Totale",
        ricevute ? "Pagato" : "Incassato",
        "Residuo",
        // Le colonne chieste dal direttore: il residuo "da formula" che regge
        // anche le NC non collegate, lo status a tre valori e i giorni.
        "Residuo (Totale-Incassato)",
        "Status",
        "Giorni",
        "Scadenza",
        "Stato",
        "Ritardo gg",
        ricevute ? "Pagamento Aruba" : "Incasso Aruba",
        ricevute ? "Data pagamento Aruba" : "Data incasso Aruba",
        "Stato da banca",
        "Discordante",
        "Stato SdI",
        "Nome file",
      ],
      filtrate.map(({ f, s }) => {
        // Incassato con la stessa regola della colonna: NC compensata =
        // negativo, NC aperta = 0.
        const incassatoCsv = isNotaCredito(f.tipoDocumento)
          ? s.statoIncassi === "Pagata" || s.statoFatturazione === "Pagata"
            ? -Math.abs(f.totale)
            : 0
          : s.incassato;
        const pagato = s.stato === "Pagata" || s.annullataDaNC || s.stato === "NC";
        // "Status": Ritardo / Pagato / In tempo. "Giorni": se pagato, data
        // documento -> data pagamento; altrimenti data documento -> oggi.
        const status = s.inRitardo ? "Ritardo" : pagato ? "Pagato" : "In tempo";
        const ggDa = (a: string, b: string) =>
          Math.round(
            (new Date(`${b}T00:00:00`).getTime() - new Date(`${a}T00:00:00`).getTime()) / 86400000,
          );
        const giorni =
          pagato && f.dataIncasso && f.dataDocumento
            ? ggDa(f.dataDocumento, f.dataIncasso)
            : f.dataDocumento
              ? ggDa(f.dataDocumento, oggiISO)
              : "";
        return [
          f.numero,
          csvData(f.dataDocumento),
          ...csvPeriodo(f.dataDocumento),
          f.cliente,
          f.tipoDocumento,
          csvNum(f.totale),
          // NC compensata = negativo (denaro non entrato), NC aperta = 0.
          csvNum(incassatoCsv),
          csvNum(s.residuo),
          csvNum(Math.round((f.totale - incassatoCsv) * 100) / 100),
          status,
          giorni,
          csvData(s.scadenza),
          // Coperta per intero dalle note di credito: nel CSV "Pagata" sarebbe
          // una bugia (residuo zero SENZA incasso) e manda a cercare bonifici
          // che non esistono. Lo stato vero è "annullata da NC", come nell'app.
          s.annullataDaNC ? "Annullata da NC" : s.stato,
          s.inRitardo ? s.giorniRitardo : "",
          s.aruba,
          csvData(f.dataIncasso),
          s.annullataDaNC ? "Annullata da NC" : s.statoBanca,
          s.discordante ? "Sì" : "",
          f.statoSdI,
          f.nomeFile,
        ];
      }),
    );
  };

  // Pallino di provenienza dello stato: "A" = registrato su Aruba (fonte
  // ufficiale), "B" = ricavato dalla riconciliazione bancaria. Il triangolo
  // segnala che le due fonti non concordano.
  const fonte = (x: (typeof conStato)[number]) => {
    if (x.s.stato === "NC") return null;
    if (!x.s.aruba)
      return (
        <span className="ml-1 text-[10px] text-muted-foreground" title={t("ft.fonteBanca")}>
          B
        </span>
      );
    return (
      <span
        className={`ml-1 text-[10px] ${x.s.discordante ? "text-status-absent font-semibold" : "text-muted-foreground"}`}
        title={
          x.s.discordante
            ? `${t("ft.fonteAruba")} · ${tp("ft.discordante", "ft.discordantePassiva")}`
            : `${t("ft.fonteAruba")}${x.f.dataIncasso ? ` · ${fmtData(x.f.dataIncasso)}` : ""}`
        }
      >
        A{x.s.discordante ? " ⚠" : ""}
      </span>
    );
  };

  // Badge di UNO stato: la stessa resa per la colonna Fatturazione e per la
  // colonna Banca, che restano affiancate e indipendenti.
  const badgeStato = (
    x: (typeof conStato)[number],
    stato: StatoIncasso | null,
    mostraRitardo: boolean,
  ) => {
    if (stato == null)
      return (
        <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
          {t("ft.nonGestita")}
        </span>
      );
    // Nota di credito: "incassata" significa compensata/liquidata, non pagata.
    if (isNotaCredito(x.f.tipoDocumento) && (stato === "Pagata" || stato === "Non incassata"))
      return (
        <span
          className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${stato === "Pagata" ? "bg-status-present/15 text-status-present" : "bg-status-break/15 text-status-break"}`}
        >
          {stato === "Pagata" ? t("ft.ncCompensata") : t("ft.ncDaCompensare")}
        </span>
      );
    if (stato === "NC")
      return (
        <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">
          {isNotaCredito(x.f.tipoDocumento)
            ? t("ft.nc")
            : isEsclusaDalCredito(x.f)
              ? x.f.statoSdI || t("ft.scartata")
              : escluse.has(x.f.nomeFile)
                ? t("ft.reinvio")
                : "—"}
        </span>
      );
    if (x.s.annullataDaNC)
      return (
        <span
          className="rounded-full bg-muted px-2 py-0.5 text-[11px] font-medium text-muted-foreground"
          title={t("ft.annullataNCTip")}
        >
          {t("ft.annullataNC")}
        </span>
      );
    if (stato === "Pagata")
      return (
        <span className="rounded-full bg-status-present/15 px-2 py-0.5 text-[11px] font-medium text-status-present">
          {t("ft.pagata")}
        </span>
      );
    const inRitardo = mostraRitardo && x.s.inRitardo;
    return (
      <span
        className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${inRitardo ? "bg-status-absent/15 text-status-absent" : "bg-status-break/15 text-status-break"}`}
      >
        {stato === "Parziale"
          ? t("ft.parziale")
          : x.f.direzione === "Ricevuta"
            ? t("ft.nonPagata")
            : t("ft.nonIncassata")}
        {inRitardo ? ` · +${x.s.giorniRitardo}gg` : ""}
      </span>
    );
  };

  const loading = fatture == null || abbinamenti == null || movimenti == null;

  return (
    <div className="space-y-4">
      {/* Riepilogo */}
      <div className="grid gap-3 sm:grid-cols-3">
        <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {ricevute ? t("ft.kpiDaPagare") : t("ft.kpiAperto")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-foreground">
            {fmtImporto(riepilogo.apertoFatt)} €
          </div>
          <div className="text-[11px] text-muted-foreground mt-0.5">
            {tp("ft.colIncassi", "ft.colPagatoFatt")}: {fmtImporto(riepilogo.apertoIncassi)} € ·{" "}
            {t("ft.colBanca")}: {fmtImporto(riepilogo.apertoBanca)} €
          </div>
        </div>
        <div className="rounded-2xl border border-status-absent/40 bg-status-absent/5 p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {t("ft.kpiRitardo")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-status-absent">
            {fmtImporto(riepilogo.ritardoImporto)} €
          </div>
          <div className="text-[11px] text-muted-foreground mt-0.5">
            {riepilogo.nRitardo} {t("ft.kpiRitardoN")}
            {riepilogo.nSenzaFonte > 0 && (
              <>
                {" · "}
                <span title={t("ft.senzaFonteTip")}>
                  {riepilogo.nSenzaFonte} {t("ft.senzaFonte")} (
                  {fmtImporto(riepilogo.senzaFonteImporto)} €)
                </span>
              </>
            )}
          </div>
        </div>
        <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
          <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
            {ricevute ? t("ft.kpiPagato") : t("ft.kpiIncassato")}
          </div>
          <div className="mt-1 text-2xl font-semibold tabular-nums text-status-present">
            {fmtImporto(riepilogo.incassatoFatt)} €
          </div>
          {/* Le due letture affiancate: possono differire, ed è un dato utile. */}
          <div className="text-[11px] text-muted-foreground mt-0.5">
            {tp("ft.colIncassi", "ft.colPagatoFatt")}: {fmtImporto(riepilogo.incassatoIncassi)} € ·{" "}
            {t("ft.colBanca")}: {fmtImporto(riepilogo.confermatoBanca)} €
            {riepilogo.nDiscordanti > 0 && (
              <>
                {" · "}
                <button
                  type="button"
                  onClick={() => setStatiF(["discordante"])}
                  className="underline underline-offset-2 hover:text-foreground"
                >
                  {riepilogo.nDiscordanti} {t("ft.kpiDaVerificare")}
                </button>
              </>
            )}
          </div>
        </div>
      </div>

      {/* I DUE SALDI, affiancati: quanto risulta incassato alla fatturazione
          e quanto e' davvero arrivato sul conto (TUTTI i movimenti in entrata
          degli anni filtrati, abbinati o no). La differenza in mezzo e' il
          numero da spiegare: compensazioni, registrazioni mancanti, periodi
          di estratto non coperti. Solo sulle attive: sulle passive le uscite
          bancarie non sono confrontabili col fatturato fornitori (stipendi,
          tasse, carte). */}
      {!ricevute && movimenti != null && (
        <div
          className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)] flex flex-wrap items-baseline gap-x-8 gap-y-2"
          title={t("ft.saldiTip")}
        >
          <div>
            <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
              {t("ft.saldoFatt")}
            </div>
            <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
              {fmtImporto(riepilogo.incassatoFatt)} €
            </div>
          </div>
          <div>
            <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
              {t("ft.saldoBanca")}
            </div>
            <div className="mt-0.5 text-xl font-semibold tabular-nums text-foreground">
              {fmtImporto(saldoMovimenti.totale)} €
              <span className="ml-1.5 text-[11px] font-normal text-muted-foreground">
                ({saldoMovimenti.n})
              </span>
            </div>
          </div>
          <div>
            <div className="text-[11px] uppercase tracking-wider text-muted-foreground">
              {t("ft.saldiDiff")}
            </div>
            <div
              className={`mt-0.5 text-xl font-semibold tabular-nums ${Math.abs(riepilogo.incassatoFatt - saldoMovimenti.totale) > TOLLERANZA_SALDO ? "text-status-absent" : "text-status-present"}`}
            >
              {fmtImporto(
                Math.round((riepilogo.incassatoFatt - saldoMovimenti.totale) * 100) / 100,
              )}{" "}
              €
            </div>
          </div>
          <p className="basis-full text-[11px] text-muted-foreground -mt-1">{t("ft.saldiNota")}</p>
        </div>
      )}

      {/* Azioni + filtri */}
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="flex flex-wrap items-end gap-3 mb-4">
          <div>
            <label className="text-xs text-muted-foreground">{t("ft.anno")}</label>
            <div className="flex flex-wrap gap-1.5 pt-1.5">
              <button
                type="button"
                onClick={() => setAnniF([])}
                className={chipCls(anniF.length === 0)}
              >
                {t("fin.allYears")}
              </button>
              {anniDisponibili.map((a) => (
                <button
                  key={a}
                  type="button"
                  onClick={() =>
                    setAnniF(anniF.includes(a) ? anniF.filter((x) => x !== a) : [...anniF, a])
                  }
                  className={chipCls(anniF.includes(a))}
                >
                  {a}
                </button>
              ))}
            </div>
          </div>
          <div>
            <label className="text-xs text-muted-foreground">{t("common.status")}</label>
            <div className="flex flex-wrap gap-1.5 pt-1.5">
              <button
                type="button"
                onClick={() => setStatiF([])}
                className={chipCls(statiF.length === 0)}
              >
                {t("common.allF")}
              </button>
              {(
                [
                  ["ritardo", t("ft.fRitardo")],
                  ["nonIncassata", ricevute ? t("ft.nonPagata") : t("ft.nonIncassata")],
                  ["parziale", t("ft.parziale")],
                  ["pagata", t("ft.pagata")],
                  ["discordante", t("ft.fDiscordanti")],
                  ["nonGestita", t("ft.fNonGestita")],
                ] as [Exclude<StatoFiltro, "tutte">, string][]
              ).map(([v, label]) => (
                <button
                  key={v}
                  type="button"
                  onClick={() =>
                    setStatiF(statiF.includes(v) ? statiF.filter((x) => x !== v) : [...statiF, v])
                  }
                  className={chipCls(statiF.includes(v))}
                >
                  {label}
                </button>
              ))}
              {/* Gli scarti sono fuori da elenco e conteggi: qui si vedono da
                  soli, per ricostruire la storia di una fattura rispedita. */}
              {nScartate > 0 && (
                <button
                  type="button"
                  onClick={() => setSoloScartate(!soloScartate)}
                  title={t("ft.scartateTip")}
                  className={chipCls(soloScartate)}
                >
                  {t("ft.fScartate")} ({nScartate})
                </button>
              )}
            </div>
          </div>
          <MultiSelect
            label={ricevute ? t("ft.fornitore") : t("fin.cliente")}
            tuttiLabel={t("common.allF")}
            selLabel={t("fin.msSel")}
            opzioni={clientiTop}
            valori={clientiSel}
            onChange={setClientiSel}
            className="w-56"
          />
          <div className="flex-1 min-w-44">
            <label className="text-xs text-muted-foreground">{t("ft.cerca")}</label>
            <input
              value={clienteF}
              onChange={(e) => setClienteF(e.target.value)}
              placeholder={t("ft.cercaPh")}
              className={inputCls}
            />
          </div>
          <button
            type="button"
            onClick={() => void riconcilia()}
            disabled={loading || reconciling}
            className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
            title={t("ft.reconcileTip")}
          >
            {reconciling ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <Wand2 className="h-4 w-4" />
            )}
            {t("ft.reconcile")}
          </button>
          <button
            type="button"
            onClick={() => void riconciliaFIFO()}
            disabled={loading || fifoBusy}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
            title={t("ft.fifoTip")}
          >
            {fifoBusy ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <Link2 className="h-4 w-4" />
            )}
            {t("ft.fifo")}
          </button>
          <button
            type="button"
            onClick={() => {
              setShowImport((v) => !v);
              setShowAruba(false);
            }}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("ft.import")}
          </button>
          <button
            type="button"
            onClick={() => {
              setShowAruba((v) => !v);
              setShowImport(false);
            }}
            className={`inline-flex items-center gap-2 rounded-lg border px-3 py-2 text-sm hover:bg-muted ${aruba?.configurato ? "border-status-present/40 text-status-present" : "border-border text-foreground"}`}
          >
            <Plug className="h-4 w-4" /> Aruba
            {aruba?.configurato && <CheckCircle2 className="h-3.5 w-3.5" />}
          </button>
          <button
            type="button"
            onClick={esporta}
            disabled={filtrate.length === 0}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
          >
            <Download className="h-4 w-4" /> {t("common.exportCsv")}
          </button>
        </div>

        {/* Import export Aruba (a scomparsa) */}
        {showImport && (
          <div className="mb-4 rounded-xl border border-border p-4">
            <div className="flex items-center gap-2 text-sm font-medium text-foreground mb-1">
              <FileSpreadsheet className="h-4 w-4 text-primary" /> {t("ft.importTitle")}
            </div>
            {/* I tre caricamenti, con la strada esatta nel pannello Aruba:
                chi importa non deve ricordare a memoria da dove si esporta. */}
            <ol className="list-decimal pl-5 text-xs text-muted-foreground mb-2 space-y-1">
              <li>{ricevute ? t("ft.imp1Ric") : t("ft.imp1Em")}</li>
              <li>{ricevute ? t("ft.imp2Ric") : t("ft.imp2Em")}</li>
              <li>{t("ft.imp3")}</li>
            </ol>
            <p className="text-xs text-muted-foreground mb-3">{t("ft.impNota")}</p>
            <input
              type="file"
              accept=".zip,.xml,.xlsx,.xls"
              multiple
              disabled={impParsing || impBusy || loading}
              onChange={(e) => {
                const files = Array.from(e.target.files ?? []);
                if (files.length) void onFiles(files);
                e.target.value = "";
              }}
              className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
            />
            {impParsing && (
              <p className="mt-2 text-sm text-muted-foreground inline-flex items-center gap-2">
                <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
              </p>
            )}
            {/* Pulizia: rimuove i documenti senza controparte lasciati da un
                file letto col tracciato sbagliato. Conta, mostra esempi e
                chiede conferma prima di toccare qualsiasi cosa. */}
            <button
              type="button"
              disabled={puliziaBusy}
              onClick={() => void pulisciSenzaCliente()}
              className="mt-3 inline-flex items-center gap-1.5 rounded-lg border border-status-absent/40 px-2.5 py-1.5 text-[12px] text-status-absent hover:bg-status-absent/5 disabled:opacity-50"
            >
              {puliziaBusy ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin" />
              ) : (
                <Trash2 className="h-3.5 w-3.5" />
              )}
              {t("ft.puliziaBtn")}
            </button>
            {previewImp && (
              <div className="mt-3 text-[13px] text-muted-foreground">
                <b className="text-foreground">{previewImp.descrizione}</b> —{" "}
                {previewImp.emesse.length + previewImp.ricevute.length} {t("ft.importRows")} (
                {previewImp.emesse.length} {t("ft.dirEmesse").toLowerCase()},{" "}
                {previewImp.ricevute.length} {t("ft.dirRicevute").toLowerCase()})
                {previewImp.duplicati > 0 &&
                  ` · ${previewImp.duplicati} ${t("ft.giaNelCaricamento")}`}
                {previewImp.scartate > 0 && ` · ${previewImp.scartate} ${t("fin.previewSkipped")}`}
                <div className="mt-2 flex gap-2">
                  <button
                    type="button"
                    onClick={() => void eseguiImport()}
                    disabled={impBusy}
                    className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-1.5 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                  >
                    {impBusy ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <Upload className="h-4 w-4" />
                    )}
                    {t("fin.importBtn")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setPreviewImp(null)}
                    className="rounded-lg border border-border px-3 py-1.5 text-sm text-foreground hover:bg-muted"
                  >
                    {t("common.cancel")}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Collegamento Aruba (a scomparsa) */}
        {showAruba && (
          <div className="mb-4 rounded-xl border border-border p-4">
            <div className="flex items-center gap-2 text-sm font-medium text-foreground mb-1">
              <Plug className="h-4 w-4 text-primary" /> {t("ft.arTitle")}
            </div>
            <p className="text-xs text-muted-foreground mb-3">
              {aruba == null
                ? t("common.loading")
                : !aruba.listaPresente
                  ? t("ft.arNoLista")
                  : aruba.configurato
                    ? `${t("ft.arConfigurato")} (${aruba.username})`
                    : t("ft.arDaConfigurare")}
            </p>
            {aruba?.listaPresente && (
              <div className="flex flex-wrap items-end gap-3">
                <div className="min-w-52">
                  <label className="text-xs text-muted-foreground">{t("ft.arUser")}</label>
                  <input
                    value={arubaUser}
                    onChange={(e) => setArubaUser(e.target.value)}
                    autoComplete="off"
                    className={inputCls}
                  />
                </div>
                <div className="min-w-52">
                  <label className="text-xs text-muted-foreground">{t("ft.arPass")}</label>
                  <input
                    type="password"
                    value={arubaPass}
                    onChange={(e) => setArubaPass(e.target.value)}
                    autoComplete="new-password"
                    className={inputCls}
                  />
                </div>
                <button
                  type="button"
                  onClick={() => void salvaCredenziali()}
                  disabled={arubaSaving || !arubaUser.trim() || !arubaPass}
                  className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
                >
                  {arubaSaving ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <KeyRound className="h-4 w-4" />
                  )}
                  {t("ft.arSalva")}
                </button>
                <button
                  type="button"
                  onClick={() => void provaConnessione()}
                  disabled={arubaTesting || !aruba.configurato}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {arubaTesting ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Plug className="h-4 w-4" />
                  )}
                  {t("ft.arProva")}
                </button>
              </div>
            )}
            {probe && (
              <div
                className={`mt-3 rounded-lg p-3 text-[13px] ${probe.ok ? "bg-status-present/10 text-foreground" : "bg-status-absent/10 text-status-absent"}`}
              >
                <div className="font-medium">{probe.messaggio}</div>
                {probe.ok && probe.campiEsempio && (
                  <div className="mt-2 text-muted-foreground">
                    <div className="text-xs font-medium text-foreground mb-1">
                      {t("ft.arCampi")}
                    </div>
                    <div className="overflow-x-auto">
                      <table className="text-xs">
                        <tbody>
                          {Object.entries(probe.campiEsempio).map(([k, v]) => (
                            <tr key={k}>
                              <td className="pr-3 py-0.5 font-mono text-foreground whitespace-nowrap">
                                {k}
                              </td>
                              <td className="py-0.5 break-all">{v}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
                {probe.ok && probe.elementi === 0 && (
                  <p className="mt-1 text-muted-foreground">{t("ft.arVuoto")}</p>
                )}
              </div>
            )}
            <p className="mt-3 text-[11px] text-muted-foreground">{t("ft.arNota")}</p>
          </div>
        )}

        {/* Elenco */}
        {loading ? (
          <div className="py-10 text-center text-sm text-muted-foreground">
            <Loader2 className="h-5 w-5 animate-spin inline-block" />
          </div>
        ) : filtrate.length === 0 ? (
          <p className="py-8 text-center text-sm text-muted-foreground">{t("ft.empty")}</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-muted-foreground border-b border-border">
                  <th className="py-1.5 pr-2">{t("ft.numero")}</th>
                  <th className="py-1.5 pr-2">{t("common.date")}</th>
                  <th className="py-1.5 pr-2">{ricevute ? t("ft.fornitore") : t("fin.cliente")}</th>
                  <th className="py-1.5 pr-2 text-right">{t("common.total")}</th>
                  <th className="py-1.5 pr-2 text-right">
                    {ricevute ? t("ft.pagato") : t("ft.incassato")}
                  </th>
                  <th className="py-1.5 pr-2">{t("ft.scadenza")}</th>
                  <th className="py-1.5 pr-2">{t("ft.colFatturazione")}</th>
                  <th className="py-1.5 pr-2">{tp("ft.colIncassi", "ft.colPagatoFatt")}</th>
                  <th className="py-1.5 pr-2">{t("ft.colBanca")}</th>
                </tr>
              </thead>
              <tbody>
                {filtrate.slice(0, 400).map((x) => {
                  const abbFat = (abbinamenti ?? []).filter((a) => a.fatturaFile === x.f.nomeFile);
                  const aperta = openFile === x.f.nomeFile;
                  return [
                    <tr
                      key={x.f.nomeFile}
                      onClick={() => {
                        setOpenFile(aperta ? null : x.f.nomeFile);
                        setAbbMov("");
                        setAbbImporto("");
                      }}
                      className={`border-b border-border/50 cursor-pointer hover:bg-muted/40 ${aperta ? "bg-muted/30" : ""}`}
                      title={x.f.nomeFile}
                    >
                      <td className="py-1 pr-2 whitespace-nowrap font-medium">{x.f.numero}</td>
                      <td className="py-1 pr-2 whitespace-nowrap">{fmtData(x.f.dataDocumento)}</td>
                      <td className="py-1 pr-2 max-w-40 truncate">{x.f.cliente}</td>
                      <td
                        className={`py-1 pr-2 text-right whitespace-nowrap ${isNotaCredito(x.f.tipoDocumento) ? "text-status-absent" : ""}`}
                      >
                        {fmtImporto(x.f.totale)}
                        {noteCredito.has(x.f.nomeFile) && (
                          <div
                            className="text-[11px] text-muted-foreground"
                            title={`${t("ft.ncCollegate")}: ${noteCredito.get(x.f.nomeFile)!.numeri.join(", ")}`}
                          >
                            −{fmtImporto(noteCredito.get(x.f.nomeFile)!.importo)} NC
                          </div>
                        )}
                      </td>
                      <td className="py-1 pr-2 text-right whitespace-nowrap text-status-present">
                        {x.s.incassatoBanca ? fmtImporto(x.s.incassatoBanca) : ""}
                      </td>
                      <td
                        className={`py-1.5 pr-3 whitespace-nowrap ${x.s.inRitardo ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                      >
                        {x.s.stato === "NC" ? "—" : fmtData(x.s.scadenza)}
                      </td>
                      {/* Le due letture, affiancate: nessuna prevale. */}
                      <td className="py-1 pr-2 whitespace-nowrap">
                        {badgeStato(x, x.s.statoFatturazione, true)}
                        {x.f.dataIncasso && (
                          <span className="ml-1 text-[11px] text-muted-foreground">
                            {fmtData(x.f.dataIncasso)}
                          </span>
                        )}
                      </td>
                      {/* Incassi REGISTRATI su Aruba: stato + importo, con i
                          parziali quantificati (report movimenti). */}
                      <td className="py-1 pr-2 whitespace-nowrap">
                        {x.s.statoIncassi == null ? (
                          <span
                            className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground"
                            title={t("ft.senzaMovimentiTip")}
                          >
                            {t("ft.senzaMovimenti")}
                          </span>
                        ) : (
                          <>
                            {badgeStato(x, x.s.statoIncassi, false)}
                            <span className="ml-1 text-[11px] tabular-nums text-muted-foreground">
                              {fmtImporto(x.s.incassatoIncassi ?? 0)}
                            </span>
                          </>
                        )}
                      </td>
                      <td className="py-1 pr-2 whitespace-nowrap">
                        {/* Zero movimenti collegati NON significa "non pagata":
                            significa che nessun bonifico e' stato abbinato.
                            Sulle passive è la norma — molti costi non passano
                            dal c/c aziendale — e lo si dice apertamente. */}
                        {x.s.statoBanca === "Non incassata" && x.s.incassatoBanca === 0 ? (
                          <span
                            className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground"
                            title={ricevute ? t("ft.fuoriBancaTip") : undefined}
                          >
                            {tp("ft.nessunAbbinamento", "ft.fuoriBanca")}
                          </span>
                        ) : (
                          <>
                            {badgeStato(x, x.s.statoBanca, false)}
                            {/* Importo abbinato, come per la colonna incassi:
                                le due letture si confrontano a colpo d'occhio. */}
                            <span className="ml-1 text-[11px] tabular-nums text-muted-foreground">
                              {fmtImporto(x.s.incassatoBanca)}
                            </span>
                          </>
                        )}
                        {x.s.discordante && (
                          <span
                            className="ml-1 text-[11px] text-status-absent"
                            title={tp("ft.discordante", "ft.discordantePassiva")}
                          >
                            ⚠
                          </span>
                        )}
                      </td>
                    </tr>,
                    aperta && (
                      <tr key={`${x.f.nomeFile}-det`} className="border-b border-border/50">
                        <td colSpan={9} className="py-3 px-3 bg-muted/20">
                          <div className="text-xs text-muted-foreground mb-2">
                            {x.f.tipoDocumento} · SdI {x.f.statoSdI || "—"} · {t("ft.terminiGg")}{" "}
                            {termini.length
                              ? `${(x.s.scadenza && x.f.dataDocumento && Math.round((new Date(x.s.scadenza).getTime() - new Date(x.f.dataDocumento).getTime()) / 86400000)) || TERMINI_DEFAULT_GIORNI}gg`
                              : `${TERMINI_DEFAULT_GIORNI}gg (default)`}
                          </div>
                          {/* Correzione manuale dello stato d'incasso: per la
                              fattura che si SA incassata, senza aspettare il
                              prossimo report. */}
                          {!isNotaCredito(x.f.tipoDocumento) && (
                            <div
                              className="flex flex-wrap items-center gap-2 mb-3 text-[13px]"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <span className="text-muted-foreground">{t("ft.incManLabel")}</span>
                              <select
                                value={incFile === x.f.nomeFile ? incStato : x.s.aruba || ""}
                                onChange={(e) => {
                                  setIncFile(x.f.nomeFile);
                                  setIncStato(e.target.value as "Incassata" | "Non incassata");
                                  if (!incData) setIncData(oggiISO);
                                }}
                                className="rounded-lg border border-border bg-background px-2 py-1 text-[13px]"
                              >
                                {!x.s.aruba && <option value="">{t("ft.nonGestita")}</option>}
                                <option value="Incassata">
                                  {ricevute ? t("ft.pagato") : t("ft.pagata")}
                                </option>
                                <option value="Non incassata">
                                  {ricevute ? t("ft.nonPagata") : t("ft.nonIncassata")}
                                </option>
                              </select>
                              {incFile === x.f.nomeFile && incStato === "Incassata" && (
                                <input
                                  type="date"
                                  value={incData || oggiISO}
                                  onChange={(e) => setIncData(e.target.value)}
                                  className="rounded-lg border border-border bg-background px-2 py-1 text-[13px]"
                                />
                              )}
                              <button
                                type="button"
                                disabled={
                                  incSaving || incFile !== x.f.nomeFile || incStato === x.s.aruba
                                }
                                onClick={() => void salvaIncassoManuale(x.f)}
                                className="rounded-lg bg-primary px-3 py-1 text-[13px] font-medium text-primary-foreground disabled:opacity-40"
                              >
                                {incSaving ? "…" : t("common.save")}
                              </button>
                              <span className="text-[11px] text-muted-foreground">
                                {t("ft.incManHint")}
                              </span>
                            </div>
                          )}
                          {/* Nota di credito: collegamento alla fattura che
                              rettifica. Quando lo storno è stato fatto dentro
                              Aruba il riferimento non arriva nell'XML, e senza
                              di esso la NC resta uno storno sospeso che sballa
                              il netto del cliente. */}
                          {isNotaCredito(x.f.tipoDocumento) && (
                            <div
                              className="flex flex-wrap items-center gap-2 mb-3 text-[13px]"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <Link2 className="h-3.5 w-3.5 text-muted-foreground shrink-0" />
                              <span className="text-muted-foreground">{t("ft.collegaA")}</span>
                              <select
                                value={
                                  ncFile === x.f.nomeFile ? ncNumero : (x.f.rettificaNumero ?? "")
                                }
                                onChange={(e) => {
                                  setNcFile(x.f.nomeFile);
                                  setNcNumero(e.target.value);
                                }}
                                className="rounded-lg border border-border bg-background px-2 py-1 text-[13px] max-w-72"
                              >
                                <option value="">{t("ft.collegaNessuna")}</option>
                                {fattureCollegabili(x.f).map((c) => (
                                  <option key={c.nomeFile} value={c.numero}>
                                    {c.numero} · {fmtData(c.dataDocumento)} · {fmtImporto(c.totale)}
                                  </option>
                                ))}
                              </select>
                              <button
                                type="button"
                                disabled={
                                  ncSaving ||
                                  ncFile !== x.f.nomeFile ||
                                  ncNumero === (x.f.rettificaNumero ?? "")
                                }
                                onClick={() => void salvaRettifica(x.f)}
                                className="rounded-lg bg-primary px-3 py-1 text-[13px] font-medium text-primary-foreground disabled:opacity-40"
                              >
                                {ncSaving ? "…" : t("common.save")}
                              </button>
                              {x.f.rettificaNumero && (
                                <span className="text-[12px] text-muted-foreground">
                                  {t("ft.collegataOra")} <b>{x.f.rettificaNumero}</b>
                                </span>
                              )}
                            </div>
                          )}
                          {abbFat.length > 0 ? (
                            <ul className="space-y-1 mb-3">
                              {abbFat.map((a) => {
                                const mov = (movimenti ?? []).find(
                                  (m) => m.chiave === a.movimentoChiave,
                                );
                                return (
                                  <li key={a.id} className="flex items-center gap-3 text-[13px]">
                                    <Link2 className="h-3.5 w-3.5 text-status-present shrink-0" />
                                    <span className="tabular-nums font-medium">
                                      {fmtImporto(a.importo)} €
                                    </span>
                                    <span className="text-muted-foreground truncate">
                                      {mov
                                        ? `${fmtData(mov.dataContabile)} · ${mov.cliente || mov.descrizione.slice(0, 50)}`
                                        : a.movimentoChiave.slice(0, 60)}
                                    </span>
                                    <span className="text-[11px] rounded-full bg-muted px-1.5">
                                      {a.origine}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        void rimuoviAbbinamento(a);
                                      }}
                                      className="ml-auto rounded-md p-1 text-muted-foreground hover:text-status-absent"
                                      title={t("ft.abbDelete")}
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </li>
                                );
                              })}
                            </ul>
                          ) : (
                            <p className="text-[13px] text-muted-foreground mb-3">
                              {t("ft.abbNone")}
                            </p>
                          )}
                          {x.s.stato !== "NC" && x.s.residuo > 0.01 && (
                            <div
                              className="flex flex-wrap items-end gap-2"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <div className="min-w-72 flex-1">
                                <label className="text-xs text-muted-foreground">
                                  {ricevute ? t("ft.abbPagamento") : t("ft.abbMovimento")}
                                </label>
                                <select
                                  value={abbMov}
                                  onChange={(e) => {
                                    setAbbMov(e.target.value);
                                    const inc = incassiDisponibili.find(
                                      (i) => i.m.chiave === e.target.value,
                                    );
                                    if (inc)
                                      setAbbImporto(
                                        String(Math.min(inc.residuo, x.s.residuo)).replace(
                                          ".",
                                          ",",
                                        ),
                                      );
                                  }}
                                  className={inputCls}
                                >
                                  <option value="">{t("common.select")}</option>
                                  {incassiDisponibili.slice(0, 200).map((i) => (
                                    <option key={i.m.chiave} value={i.m.chiave}>
                                      {fmtData(i.m.dataContabile)} · {fmtImporto(i.residuo)} € ·{" "}
                                      {(i.m.cliente || i.m.descrizione).slice(0, 60)}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <div className="w-32">
                                <label className="text-xs text-muted-foreground">
                                  {t("common.amount")}
                                </label>
                                <input
                                  value={abbImporto}
                                  onChange={(e) => setAbbImporto(e.target.value)}
                                  className={inputCls}
                                />
                              </div>
                              <button
                                type="button"
                                onClick={() => void abbinaManuale(x.f.nomeFile, x.s.residuo)}
                                disabled={abbBusy || !abbMov}
                                className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                              >
                                {abbBusy ? (
                                  <Loader2 className="h-4 w-4 animate-spin" />
                                ) : (
                                  <Link2 className="h-4 w-4" />
                                )}
                                {t("ft.abbina")}
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    ),
                  ];
                })}
              </tbody>
            </table>
            {filtrate.length > 400 && (
              <div className="mt-2 text-xs text-muted-foreground">
                {t("fin.first500")} {filtrate.length}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Riepilogo per cliente */}
      {/* Riepilogo dell'import: si chiude solo quando lo decide chi legge. */}
      {esitoImport && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4">
          <div className="w-full max-w-md rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-elegant)]">
            <div className="flex items-center gap-2 text-[15px] font-semibold text-foreground">
              <CheckCircle2 className="h-5 w-5 text-status-present" />
              {t("ft.importDone")}
            </div>
            <ul className="mt-3 space-y-1 text-sm">
              {esitoImport.nuove > 0 && (
                <li>
                  <b className="tabular-nums">{esitoImport.nuove}</b> {t("ft.importNew")}
                </li>
              )}
              {esitoImport.aggiornate > 0 && (
                <li>
                  <b className="tabular-nums">{esitoImport.aggiornate}</b> {t("ft.importUpd")}
                </li>
              )}
              {esitoImport.doppioni > 0 && (
                <li className="text-muted-foreground">
                  <b className="tabular-nums">{esitoImport.doppioni}</b> {t("ft.importDup")}
                </li>
              )}
              {esitoImport.rate > 0 && (
                <li>
                  <b className="tabular-nums">{esitoImport.fattureIncassi}</b>{" "}
                  {t("ft.movApplicatiDesc")}{" "}
                  <span className="text-muted-foreground">
                    ({esitoImport.rate} {t("ft.movRate")})
                  </span>
                </li>
              )}
              {esitoImport.errori.length > 0 && (
                <li className="text-status-absent">
                  <b className="tabular-nums">{esitoImport.errori.length}</b>{" "}
                  {t("common.error").toLowerCase()}
                  <div className="mt-1 max-h-24 overflow-auto text-[11px]">
                    {esitoImport.errori.slice(0, 5).map((e, i) => (
                      <div key={i}>{e}</div>
                    ))}
                  </div>
                </li>
              )}
            </ul>
            <button
              type="button"
              onClick={() => setEsitoImport(null)}
              className="mt-4 w-full rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90"
            >
              {t("common.close")}
            </button>
          </div>
        </div>
      )}

      {perCliente.length > 0 && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="flex items-center gap-2 text-sm font-semibold text-foreground mb-3">
            <Users className="h-4 w-4 text-muted-foreground" />
            {ricevute ? t("ft.perFornitoreTitle") : t("ft.perClienteTitle")}
            <span className="text-xs font-normal text-muted-foreground">({perCliente.length})</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-[13px] leading-tight">
              <thead>
                <tr className="text-left text-xs text-muted-foreground border-b border-border">
                  <th className="py-1.5 pr-2">{ricevute ? t("ft.fornitore") : t("fin.cliente")}</th>
                  <th className="py-1.5 pr-2 text-right">{t("ft.nFatture")}</th>
                  <th className="py-1.5 pr-2 text-right">{t("ft.nPagate")}</th>
                  <th className="py-1.5 pr-2 text-right">{t("ft.nParziali")}</th>
                  <th className="py-1.5 pr-2 text-right">
                    {tp("ft.nDaIncassare", "ft.nDaPagare")}
                  </th>
                  <th className="py-1.5 pr-2 text-right">{t("ft.totFatturato")}</th>
                  <th className="py-1.5 pr-2 text-right">
                    {tp("ft.totIncassatoFatt", "ft.totPagatoFatt")}
                  </th>
                  <th className="py-1.5 pr-2 text-right">
                    {tp("ft.totIncassatoBanca", "ft.totPagatoBanca")}
                  </th>
                  {/* Il colpo d'occhio chiesto dal direttore: dove la
                      fatturazione dice "incassato" ma la banca non copre. */}
                  <th className="py-1.5 pr-2 text-right" title={t("ft.deltaTip")}>
                    {t("ft.deltaFattBanca")}
                  </th>
                  <th className="py-1.5 pr-2 text-right">
                    {tp("ft.totDaIncassare", "ft.totDaPagare")}
                  </th>
                  {/* La statistica del direttore: termine contrattuale vs
                      giorni medi effettivi (Postadoc: 60 gg sulla carta,
                      paga a 88,7). */}
                  {!ricevute && (
                    <>
                      <th className="py-1.5 pr-2 text-right" title={t("ft.ggTermineTip")}>
                        {t("ft.ggTermine")}
                      </th>
                      <th className="py-1.5 pr-2 text-right" title={t("ft.mediaGgTip")}>
                        {t("ft.mediaGg")}
                      </th>
                    </>
                  )}
                </tr>
              </thead>
              <tbody>
                {perCliente.map((r) => {
                  const aperto = clienteAperto === r.key;
                  return [
                    <tr
                      key={r.key}
                      onClick={() => {
                        setClienteAperto(aperto ? null : r.key);
                        setDettaglioTutte(false);
                      }}
                      className={`border-b border-border/50 cursor-pointer hover:bg-muted/40 ${aperto ? "bg-muted/30" : ""}`}
                    >
                      <td className="py-1 pr-2 max-w-64 truncate font-medium" title={r.cliente}>
                        {r.cliente}
                      </td>
                      <td className="py-1 pr-2 text-right tabular-nums">{r.nFatture}</td>
                      <td className="py-1 pr-2 text-right tabular-nums text-status-present">
                        {r.nPagate || ""}
                      </td>
                      <td className="py-1 pr-2 text-right tabular-nums text-status-break">
                        {r.nParziali || ""}
                      </td>
                      <td
                        className={`py-1 pr-2 text-right tabular-nums ${r.nRitardo > 0 ? "text-status-absent font-medium" : ""}`}
                        title={r.nRitardo > 0 ? `${r.nRitardo} ${t("ft.kpiRitardoN")}` : undefined}
                      >
                        {r.nDaIncassare || ""}
                      </td>
                      <td className="py-1 pr-2 text-right tabular-nums font-medium">
                        {fmtImporto(r.fatturato)}
                      </td>
                      <td className="py-1 pr-2 text-right tabular-nums text-status-present">
                        {fmtImporto(r.incassatoFatt)}
                      </td>
                      <td className="py-1 pr-2 text-right tabular-nums text-muted-foreground">
                        {fmtImporto(r.incassatoBanca)}
                      </td>
                      {(() => {
                        const delta = Math.round((r.incassatoFatt - r.incassatoBanca) * 100) / 100;
                        return (
                          <td
                            className={`py-1 pr-2 text-right tabular-nums ${Math.abs(delta) > TOLLERANZA_SALDO ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                            title={t("ft.deltaTip")}
                          >
                            {Math.abs(delta) > TOLLERANZA_SALDO ? fmtImporto(delta) : ""}
                          </td>
                        );
                      })()}
                      <td
                        className={`py-1 pr-2 text-right tabular-nums font-medium ${r.ritardo > 0 ? "text-status-absent" : ""}`}
                        title={
                          r.ritardo > 0
                            ? `${fmtImporto(r.ritardo)} ${t("ft.diCuiRitardo")}`
                            : undefined
                        }
                      >
                        {fmtImporto(r.residuo)}
                      </td>
                      {!ricevute &&
                        (() => {
                          const termine = giorniPerCliente(r.cliente, termini);
                          const media = r.nGg > 0 ? r.sommaGg / r.nGg : null;
                          return (
                            <>
                              <td
                                className="py-1 pr-2 text-right tabular-nums text-muted-foreground"
                                title={t("ft.ggTermineTip")}
                              >
                                {termine}
                              </td>
                              <td
                                className={`py-1 pr-2 text-right tabular-nums ${media != null && media > termine ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                                title={
                                  media != null
                                    ? `${r.nGg} ${t("ft.pagata").toLowerCase()}`
                                    : undefined
                                }
                              >
                                {media != null
                                  ? media.toLocaleString("it-IT", {
                                      minimumFractionDigits: 1,
                                      maximumFractionDigits: 1,
                                    })
                                  : ""}
                              </td>
                            </>
                          );
                        })()}
                    </tr>,
                    aperto && (
                      <tr key={`${r.key}-det`} className="border-b border-border/50">
                        <td colSpan={ricevute ? 10 : 12} className="py-2 px-3 bg-muted/20">
                          {/* Dettaglio del cliente: di default le sole fatture
                              da incassare (intere o parziali), con le tre
                              letture affiancate. "Tutte" per lo storico. */}
                          <div
                            className="flex items-center gap-2 mb-1.5 text-[12px]"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <span className="font-semibold text-foreground">
                              {dettaglioTutte
                                ? `${fattureDelCliente.length} ${t("ft.dettTutte")}`
                                : `${fattureDelClienteAperte.length} ${tp("ft.dettAperte", "ft.dettApertePassive")}`}
                            </span>
                            <button
                              type="button"
                              onClick={() => setDettaglioTutte(!dettaglioTutte)}
                              className="rounded-full border border-border px-2.5 py-0.5 text-[11px] hover:bg-muted"
                            >
                              {dettaglioTutte
                                ? tp("ft.dettSoloAperte", "ft.dettSoloApertePassive")
                                : t("ft.dettMostraTutte")}
                            </button>
                          </div>
                          <table className="w-full text-[12px]">
                            <thead>
                              <tr className="text-left text-[11px] text-muted-foreground">
                                <th className="py-1 pr-2">{t("ft.numero")}</th>
                                <th className="py-1 pr-2">{t("ft.data")}</th>
                                <th className="py-1 pr-2 text-right">{t("common.total")}</th>
                                <th className="py-1 pr-2">{t("ft.scadenza")}</th>
                                <th className="py-1 pr-2">{t("ft.colFatturazione")}</th>
                                <th className="py-1 pr-2">
                                  {tp("ft.colIncassi", "ft.colPagatoFatt")}
                                </th>
                                <th className="py-1 pr-2">{t("ft.colBanca")}</th>
                                <th className="py-1 pr-2 text-right">{t("ft.residuo")}</th>
                              </tr>
                            </thead>
                            <tbody>
                              {(dettaglioTutte ? fattureDelCliente : fattureDelClienteAperte).map(
                                (x) => (
                                  <tr key={x.f.nomeFile} className="border-t border-border/40">
                                    <td className="py-0.5 pr-2 whitespace-nowrap">{x.f.numero}</td>
                                    <td className="py-0.5 pr-2 whitespace-nowrap">
                                      {fmtData(x.f.dataDocumento)}
                                    </td>
                                    <td
                                      className={`py-0.5 pr-2 text-right whitespace-nowrap tabular-nums ${isNotaCredito(x.f.tipoDocumento) ? "text-status-absent" : ""}`}
                                    >
                                      {fmtImporto(x.f.totale)}
                                    </td>
                                    <td
                                      className={`py-0.5 pr-2 whitespace-nowrap ${x.s.inRitardo ? "text-status-absent" : "text-muted-foreground"}`}
                                    >
                                      {x.s.stato === "NC" ? "—" : fmtData(x.s.scadenza)}
                                    </td>
                                    <td className="py-0.5 pr-2 whitespace-nowrap">
                                      {badgeStato(x, x.s.statoFatturazione, true)}
                                    </td>
                                    <td className="py-0.5 pr-2 whitespace-nowrap">
                                      {x.s.statoIncassi == null ? (
                                        <span className="text-[11px] text-muted-foreground">—</span>
                                      ) : (
                                        <>
                                          {badgeStato(x, x.s.statoIncassi, false)}
                                          <span className="ml-1 text-[11px] tabular-nums text-muted-foreground">
                                            {fmtImporto(x.s.incassatoIncassi ?? 0)}
                                          </span>
                                        </>
                                      )}
                                    </td>
                                    <td className="py-0.5 pr-2 whitespace-nowrap">
                                      {badgeStato(x, x.s.statoBanca, false)}
                                      <span className="ml-1 text-[11px] tabular-nums text-muted-foreground">
                                        {fmtImporto(x.s.incassatoBanca)}
                                      </span>
                                    </td>
                                    <td className="py-0.5 pr-2 text-right whitespace-nowrap tabular-nums font-medium">
                                      {x.s.stato === "NC" ? "" : fmtImporto(x.s.residuo)}
                                    </td>
                                  </tr>
                                ),
                              )}
                            </tbody>
                          </table>
                          {/* Quadratura + movimenti bancari del cliente: senza
                              i movimenti NON attribuiti la differenza fra le
                              colonne resta inspiegata. */}
                          {estrattoCliente && (
                            <div className="mt-3 border-t border-border pt-3">
                              <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[12px] mb-2">
                                <span className="font-semibold text-foreground">
                                  {t("ft.quadratura")}
                                </span>
                                <span>
                                  {t("ft.totFatturato")}:{" "}
                                  <b className="tabular-nums">
                                    {fmtImporto(estrattoCliente.fatturato)}
                                  </b>
                                </span>
                                <span>
                                  {tp("ft.totIncassatoFatt", "ft.totPagatoFatt")}:{" "}
                                  <b className="tabular-nums text-status-present">
                                    {fmtImporto(estrattoCliente.incassatoFatt)}
                                  </b>
                                </span>
                                <span>
                                  {tp("ft.totDaIncassare", "ft.totDaPagare")}:{" "}
                                  <b className="tabular-nums">
                                    {fmtImporto(estrattoCliente.residuo)}
                                  </b>
                                </span>
                                <span className="text-muted-foreground">
                                  {t("ft.colBanca")}:{" "}
                                  <b className="tabular-nums">
                                    {fmtImporto(estrattoCliente.totaleBanca)}
                                  </b>{" "}
                                  ({estrattoCliente.righe.length})
                                </span>
                                {estrattoCliente.nNonAttribuiti > 0 && (
                                  <span
                                    className="text-status-absent"
                                    title={t("ft.nonAttribuitoTip")}
                                  >
                                    {t("ft.nonAttribuito")}:{" "}
                                    <b className="tabular-nums">
                                      {fmtImporto(estrattoCliente.nonAttribuito)}
                                    </b>{" "}
                                    ({estrattoCliente.nNonAttribuiti})
                                  </span>
                                )}
                                {/* La scomposizione della differenza: quanto
                                    dell'incassato di fatturazione la banca non
                                    copre, e quanto ne resterebbe scoperto anche
                                    abbinando tutti i movimenti non attribuiti.
                                    È il numero da spiegare: altri conti,
                                    periodi non importati o compensazioni. */}
                                {(() => {
                                  const delta =
                                    Math.round(
                                      (estrattoCliente.incassatoFatt -
                                        estrattoCliente.incassatoBanca) *
                                        100,
                                    ) / 100;
                                  const scoperto =
                                    Math.round((delta - estrattoCliente.nonAttribuito) * 100) / 100;
                                  if (Math.abs(delta) <= TOLLERANZA_SALDO) return null;
                                  return (
                                    <span title={t("ft.deltaTip")}>
                                      {t("ft.deltaFattBanca")}:{" "}
                                      <b className="tabular-nums text-status-absent">
                                        {fmtImporto(delta)}
                                      </b>
                                      {Math.abs(scoperto) > TOLLERANZA_SALDO && (
                                        <span
                                          className="text-muted-foreground"
                                          title={t("ft.senzaTracciaTip")}
                                        >
                                          {" "}
                                          · {t("ft.senzaTraccia")}:{" "}
                                          <b className="tabular-nums">{fmtImporto(scoperto)}</b>
                                        </span>
                                      )}
                                    </span>
                                  );
                                })()}
                                <span className="ml-auto inline-flex items-center gap-2">
                                  {estrattoCliente.nNonAttribuiti > 0 && (
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        calcolaSpiegazioni();
                                      }}
                                      title={t("ft.spiegaTip")}
                                      className="inline-flex items-center gap-1.5 rounded-lg bg-primary px-2.5 py-1 text-[12px] font-medium text-primary-foreground"
                                    >
                                      <Wand2 className="h-3.5 w-3.5" />
                                      {t("ft.spiegaBtn")}
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      esportaEstratto();
                                    }}
                                    className="inline-flex items-center gap-1.5 rounded-lg border border-border px-2.5 py-1 text-[12px] hover:bg-muted"
                                  >
                                    <Download className="h-3.5 w-3.5" />
                                    {t("ft.esportaEstratto")}
                                  </button>
                                </span>
                              </div>
                              {/* LA LETTURA EFFETTIVA: fatturato netto meno
                                  euro arrivati. Non dipende dagli stati di
                                  Aruba (che possono essere indietro) né dagli
                                  abbinamenti: è il numero da dare al
                                  direttore, purché l'archivio banca copra
                                  l'intero rapporto col cliente. */}
                              {estrattoCliente.righe.length > 0 && (
                                <div
                                  className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[12px] mb-2 rounded-lg bg-primary/5 border border-primary/20 px-3 py-1.5"
                                  title={t("ft.effTip")}
                                >
                                  <span className="font-semibold text-foreground">
                                    {t("ft.effTitolo")}
                                  </span>
                                  <span>
                                    {t("ft.effNetto")}:{" "}
                                    <b className="tabular-nums">
                                      {fmtImporto(estrattoCliente.nettoDovuto)}
                                    </b>
                                    {estrattoCliente.controFatture > 0 && (
                                      <span className="text-muted-foreground">
                                        {" "}
                                        ({t("ft.effContro")} −
                                        {fmtImporto(estrattoCliente.controFatture)})
                                      </span>
                                    )}
                                  </span>
                                  <span>
                                    {t("ft.effIncassato")}:{" "}
                                    <b className="tabular-nums text-status-present">
                                      {fmtImporto(estrattoCliente.totaleBanca)}
                                    </b>
                                  </span>
                                  <span>
                                    {tp("ft.effResiduo", "ft.effResiduoPassive")}:{" "}
                                    <b
                                      className={`tabular-nums ${estrattoCliente.residuoEffettivo > TOLLERANZA_SALDO ? "text-status-absent" : "text-status-present"}`}
                                    >
                                      {fmtImporto(estrattoCliente.residuoEffettivo)}
                                    </b>
                                  </span>
                                </div>
                              )}
                              {/* Le proposte: ogni bonifico non attribuito con
                                  la combinazione di documenti che lo spiega.
                                  Niente si salva finché non si preme Applica. */}
                              {spiegazioni && (
                                <div className="mb-3 rounded-xl border border-border bg-background p-3">
                                  <div className="text-[12px] font-semibold text-foreground mb-1.5">
                                    {t("ft.spiegaTitolo")} (
                                    {spiegazioni.filter((s) => s.pezzi.length).length}/
                                    {spiegazioni.length})
                                  </div>
                                  {spiegazioni.length === 0 ? (
                                    <p className="text-[12px] text-muted-foreground">
                                      {t("ft.spiegaNiente")}
                                    </p>
                                  ) : (
                                    <ul className="space-y-1.5">
                                      {spiegazioni.map((s) => (
                                        <li
                                          key={s.movimentoChiave}
                                          className="flex flex-wrap items-center gap-x-2 gap-y-1 text-[12px]"
                                        >
                                          <span className="whitespace-nowrap text-muted-foreground">
                                            {fmtData(s.data)}
                                          </span>
                                          <b className="tabular-nums whitespace-nowrap">
                                            {fmtImporto(s.importo)}
                                          </b>
                                          {s.pezzi.length === 0 ? (
                                            <span className="text-muted-foreground">
                                              = {t("ft.spiegaNoCombo")}
                                            </span>
                                          ) : (
                                            <>
                                              <span>=</span>
                                              {s.pezzi.map((p) => (
                                                <span
                                                  key={p.chiave}
                                                  className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${
                                                    p.tipo === "FT"
                                                      ? "bg-status-present/15 text-status-present"
                                                      : p.tipo === "NC"
                                                        ? "bg-status-absent/15 text-status-absent"
                                                        : "bg-muted text-muted-foreground"
                                                  }`}
                                                  title={
                                                    p.tipo === "CONTRO"
                                                      ? t("ft.spiegaContro")
                                                      : p.tipo === "NC"
                                                        ? t("ft.spiegaNc")
                                                        : undefined
                                                  }
                                                >
                                                  {p.tipo === "FT" ? "" : "− "}
                                                  {p.etichetta} {fmtImporto(Math.abs(p.valore))}
                                                </span>
                                              ))}
                                              {Math.abs(s.delta) > 0.005 && (
                                                <span className="text-[11px] text-muted-foreground">
                                                  [Δ {fmtImporto(s.delta)}]
                                                </span>
                                              )}
                                              <button
                                                type="button"
                                                disabled={spiegaBusy}
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  void applicaSpiegazione(s);
                                                }}
                                                className="rounded-lg bg-primary px-2.5 py-0.5 text-[11px] font-medium text-primary-foreground disabled:opacity-40"
                                              >
                                                {t("ft.spiegaApplica")}
                                              </button>
                                            </>
                                          )}
                                        </li>
                                      ))}
                                    </ul>
                                  )}
                                </div>
                              )}
                              {estrattoCliente.righe.length > 0 && (
                                <table className="w-full text-[12px]">
                                  <thead>
                                    <tr className="text-left text-[11px] text-muted-foreground">
                                      <th className="py-1 pr-2">{t("ft.data")}</th>
                                      <th className="py-1 pr-2 text-right">{t("common.total")}</th>
                                      <th className="py-1 pr-2 text-right">{t("ft.abbinato")}</th>
                                      <th className="py-1 pr-2 text-right">{t("ft.residuo")}</th>
                                      <th className="py-1 pr-2">{t("ft.numero")}</th>
                                      <th className="py-1 pr-2">{t("fin.causale")}</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {estrattoCliente.righe.map((r) => (
                                      <tr key={r.m.chiave} className="border-t border-border/40">
                                        <td className="py-0.5 pr-2 whitespace-nowrap">
                                          {fmtData(r.m.dataContabile)}
                                        </td>
                                        <td className="py-0.5 pr-2 text-right tabular-nums whitespace-nowrap">
                                          {fmtImporto(Math.abs(r.m.importo))}
                                        </td>
                                        <td className="py-0.5 pr-2 text-right tabular-nums whitespace-nowrap text-status-present">
                                          {r.allocato ? fmtImporto(r.allocato) : ""}
                                        </td>
                                        <td
                                          className={`py-0.5 pr-2 text-right tabular-nums whitespace-nowrap ${r.residuo > TOLLERANZA_SALDO ? "text-status-absent font-medium" : "text-muted-foreground"}`}
                                        >
                                          {r.residuo > TOLLERANZA_SALDO
                                            ? fmtImporto(r.residuo)
                                            : ""}
                                        </td>
                                        <td className="py-0.5 pr-2 whitespace-nowrap">
                                          {r.fattureAbb.join(", ") || r.m.nrFattura}
                                        </td>
                                        <td
                                          className="py-0.5 pr-2 max-w-72 truncate text-muted-foreground"
                                          title={r.m.descrizione}
                                        >
                                          {r.m.causale || r.m.descrizione}
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              )}
                            </div>
                          )}
                        </td>
                      </tr>
                    ),
                  ];
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="flex items-start gap-2 text-xs text-muted-foreground">
        <CheckCircle2 className="h-4 w-4 shrink-0 mt-0.5" />
        <p>{tp("ft.nota", "ft.notaPassive")}</p>
      </div>
    </div>
  );
}
