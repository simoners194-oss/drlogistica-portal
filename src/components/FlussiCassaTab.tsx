// DR Portal — Finanza → tab Flussi di cassa (direttore, call 07/09/2026).
// "La priorità assoluta dell'azienda è avere davvero in mano le spese":
// entrate e uscite ATTESE per mese (o settimana ISO), dalle scadenze delle
// fatture aperte — stessa semantica del Resoconto (residuoAperto) — più le
// righe che le fatture non conoscono: stipendi, costo fiscale, altre spese
// (voci manuali su SharePoint, lista FlussiCassa) e le prefatture. In fondo
// il DELTA SALDO. Le uscite viaggiano col segno meno: la griglia si incolla
// in Excel e si somma da sola. Esclusioni per controparte (anche a finestra
// di mesi) per tenere fuori chi non paga e le casse esterne.
import { Fragment, useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { FlaskConical, Loader2, Trash2 } from "lucide-react";
import { useLang } from "@/lib/i18n";
import {
  computeStatoFattura,
  collegaNoteCredito,
  fattureEscluse,
  residuoAperto,
  incassatoRegistrato,
  isNotaCredito,
  type TerminePagamento,
} from "@/lib/fatture-logic";
import { clienteGroupKey } from "@/lib/finanza-logic";
import {
  MODALITA_ORDINE,
  analizzaMedia,
  costiFissiDa,
  esclusaDaFlussi,
  fissoAttivo,
  mappaFornitori,
  matchFisso,
  modalitaFattura,
  piuVicinoAlFisso,
  proiezioneMeseCorrente,
  type ModalitaPagamento,
  type MovimentoAnalizzato,
  type MovimentoMedia,
} from "@/lib/flussi-logic";
import { esportaCsvFile } from "@/lib/csv";
import {
  spGetFatture,
  spGetTerminiPagamento,
  spGetPrefatture,
  spGetFlussiCassa,
  spUpsertFlussoCassa,
  spDeleteFlussoCassa,
  spGetMovimenti,
} from "@/lib/sharepoint.functions";
import { spStipendiFlussi, spStipendiGet } from "@/lib/stipendi.functions";
import { chiaveNome, type StipendiDb } from "@/lib/stipendi-logic";
import {
  chiaveScadenzaFile,
  parseScadenzario,
  ripartizioneFiscale,
  type FiscaleDb,
  type ParseScadenzarioResult,
  type ScadenzaFiscale,
} from "@/lib/fiscale-logic";
import { spFiscaleGet, spFiscaleSalva } from "@/lib/fiscale.functions";
import type { SpFattura, SpMovimento, Prefattura, FlussoCassaRiga } from "@/lib/sharepoint.server";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Le 4 voci manuali nominate dal direttore: righe sempre visibili, anche
// vuote, così Sabrina/Lucrezia sanno dove scrivere.
// "Consulenze" (Simone 29/09, v1.86.0): riga propria sotto Stipendi — non
// sono costi ricorrenti, quindi fuori dalle Altre spese e con la media
// calcolata senza anomalie.
const VOCI_BASE = [
  "Stipendi",
  "Consulenze",
  "Costo fiscale rate",
  "Costo fiscale corrente",
  "Altre spese",
];
// Gruppi apribili col "+" nelle due tabelle (entrate + uscite per modalità).
const GRUPPI = [
  "t1:entrate",
  "t1:riba",
  "t1:rid",
  "t1:altro",
  "t2:entrate",
  "t2:riba",
  "t2:rid",
  "t2:altro",
];
const TIP_CONSULENZE = "consulenze";

// --- Periodi -----------------------------------------------------------------

/** Chiave ISO-settimana "2026-W37" del giorno dato. */
function chiaveSettimana(iso: string): string {
  const d = new Date(`${iso.slice(0, 10)}T00:00:00Z`);
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const dayNum = (t.getUTCDay() + 6) % 7;
  t.setUTCDate(t.getUTCDate() - dayNum + 3); // il giovedì decide l'anno ISO
  const anno = t.getUTCFullYear();
  const gen4 = new Date(Date.UTC(anno, 0, 4));
  const sett =
    1 +
    Math.round(((t.getTime() - gen4.getTime()) / 86400000 - 3 + ((gen4.getUTCDay() + 6) % 7)) / 7);
  return `${anno}-W${String(sett).padStart(2, "0")}`;
}

function lunedioDi(iso: string): Date {
  const d = new Date(`${iso.slice(0, 10)}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
  return d;
}

export function FlussiCassaTab() {
  const { t } = useLang();
  const [fattureEm, setFattureEm] = useState<SpFattura[] | null>(null);
  const [fattureRic, setFattureRic] = useState<SpFattura[] | null>(null);
  const [termini, setTermini] = useState<TerminePagamento[]>([]);
  const [prefatture, setPrefatture] = useState<Prefattura[] | null>(null);
  const [flussi, setFlussi] = useState<FlussoCassaRiga[] | null>(null);
  const [flussiErr, setFlussiErr] = useState<string | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);

  const [modo, setModo] = useState<"mese" | "settimana">("mese");
  const [finoA, setFinoA] = useState("");
  const [daData, setDaData] = useState("");
  // Voci dei fornitori chiuse di default: si aprono col "+" gruppo per
  // gruppo (Simone 29/09) o tutte insieme col bottone in alto.
  const [aperti, setAperti] = useState<Set<string>>(new Set());
  const toggleGruppo = (k: string) =>
    setAperti((s) => {
      const ns = new Set(s);
      if (ns.has(k)) ns.delete(k);
      else ns.add(k);
      return ns;
    });
  const tuttiAperti = aperti.size >= GRUPPI.length;
  // CUMULATO (richiesta Simone 14/09): ogni colonna mostra il progressivo
  // "fino a quel momento" — a settembre scaduto+settembre, a ottobre
  // scaduto+settembre+ottobre, e così via. Vale per entrambe le tabelle.
  const [cumulato, setCumulato] = useState(false);
  // Stima Stipendi per i mesi senza dato (media netto dovuto ultimi 2 mesi
  // da "Stipendi Dr.xlsx" — richiesta Simone 14/09).
  const [autoStipendi, setAutoStipendi] = useState<{ media: number; mesi: string[] } | null>(null);
  // Totale/pagati per MESE DI PAGAMENTO dalla spunta "Pagato" della tab
  // Stipendi (richiesta Simone 15/09): tabella reale = solo i si',
  // "solo fatturazioni" = tutti.
  const [stipendiMesi, setStipendiMesi] = useState<Record<
    string,
    { totale: number; pagati: number; flaggati: number }
  > | null>(null);
  // Scadenziario fiscale (Fiscale\SCADENZARIO FISCALE__aggiornato.xlsx di
  // Sabrina — richiesta Simone 14/09): riempie "Costo fiscale rate" e
  // "Costo fiscale corrente" per i mesi senza valore manuale.
  const [fiscale, setFiscale] = useState<FiscaleDb | null>(null);
  const [showFisc, setShowFisc] = useState(false);
  const [parsingF, setParsingF] = useState(false);
  const [previewFisc, setPreviewFisc] = useState<{
    fileName: string;
    res: ParseScadenzarioResult;
  } | null>(null);
  const [savingF, setSavingF] = useState(false);
  // Pannello "cosa comprende" al click sul nome delle 4 voci (Simone 15/09).
  const [drill, setDrill] = useState<{ voce: string; mese: string } | null>(null);
  const [drillStip, setDrillStip] = useState<StipendiDb | null>(null);
  const [drillBusy, setDrillBusy] = useState(false);

  // Editor cella voce manuale: chiave "nome|periodo".
  const [cellaVoce, setCellaVoce] = useState<string | null>(null);
  const [cellaVal, setCellaVal] = useState("");
  const [salvando, setSalvando] = useState(false);

  // Form esclusioni: spunte multiple sulla checklist delle controparti.
  const [showEscl, setShowEscl] = useState(false);
  const [exSel, setExSel] = useState<Set<string>>(new Set());
  const [exCerca, setExCerca] = useState("");
  const [exDa, setExDa] = useState("");
  const [exA, setExA] = useState("");
  const [exBusy, setExBusy] = useState(false);
  // Preset di esclusioni (nome con cui salvare l'insieme corrente).
  const [presetNome, setPresetNome] = useState("");
  const [presetBusy, setPresetBusy] = useState(false);
  // Form girate: "se entra una fattura da X gira il P% a Y".
  const [giCliente, setGiCliente] = useState("");
  const [giPerc, setGiPerc] = useState("90");
  const [giFornSel, setGiFornSel] = useState("DR Logistics");
  const [giFornAltro, setGiFornAltro] = useState("");
  const [giOggetto, setGiOggetto] = useState("");
  const [giBusy, setGiBusy] = useState(false);
  // Form nuova voce manuale.
  const [nuovaVoce, setNuovaVoce] = useState("");
  // SIMULAZIONE "se non pago" (Simone 29/09, v1.84.0): tre leve combinabili
  // — fatture scelte una a una, tutte quelle senza RID/RiBa, scadenze
  // fiscali scelte. Solo in memoria: chiudendo la pagina sparisce.
  const [showSim, setShowSim] = useState(false);
  const [simPanel, setSimPanel] = useState<"fatture" | "fiscali" | null>(null);
  const [simFatture, setSimFatture] = useState<Set<string>>(new Set());
  const [simSenzaMod, setSimSenzaMod] = useState(false);
  const [simFiscali, setSimFiscali] = useState<Set<string>>(new Set());
  const [simCerca, setSimCerca] = useState("");
  // Form costo fisso (pannello Altre spese, v1.85.0).
  const [fxNome, setFxNome] = useState("");
  const [fxImporto, setFxImporto] = useState("");
  const [fxToken, setFxToken] = useState("");
  const [fxDa, setFxDa] = useState("");
  const [fxA, setFxA] = useState("");
  const [fxBusy, setFxBusy] = useState(false);

  const ricaricaFlussi = () =>
    spGetFlussiCassa()
      .then((l) => {
        setFlussi(l as FlussoCassaRiga[]);
        setFlussiErr(null);
      })
      .catch((err) => {
        setFlussi([]);
        setFlussiErr(err instanceof Error ? err.message : String(err));
      });
  useEffect(() => {
    // Stima stipendi per i mesi futuri (media netto dovuto ultimi 2 mesi).
    // Nell'effetto, non nel corpo: nel corpo partiva una fetch a OGNI render.
    spStipendiFlussi()
      .then((r) => {
        setAutoStipendi(r.stima);
        setStipendiMesi(r.perMese);
      })
      .catch(() => {
        setAutoStipendi(null);
        setStipendiMesi(null);
      });
    // Scadenziario fiscale per le due voci "Costo fiscale".
    spFiscaleGet()
      .then((f) => setFiscale(f))
      .catch(() => setFiscale(null));
    spGetFatture({ data: { direzione: "Emessa" } })
      .then((l) => setFattureEm(l as SpFattura[]))
      .catch(() => setFattureEm([]));
    spGetFatture({ data: { direzione: "Ricevuta" } })
      .then((l) => setFattureRic(l as SpFattura[]))
      .catch(() => setFattureRic([]));
    spGetTerminiPagamento()
      .then((l) => setTermini(l as TerminePagamento[]))
      .catch(() => setTermini([]));
    spGetPrefatture()
      .then((l) => setPrefatture(l as Prefattura[]))
      .catch(() => setPrefatture([]));
    spGetMovimenti()
      .then((l) => setMovimenti(l as SpMovimento[]))
      .catch(() => setMovimenti([]));
    void ricaricaFlussi();
  }, []);

  const oggiISO = new Date().toISOString().slice(0, 10);

  // Stessa preparazione del Resoconto: stati calcolati, NC collegate.
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
    // REGOLA FR 08/09: a DR Logistics si deve SOLO la quota 90% delle iMile
    // di facchinaggio (riga dedicata piu' sotto) — le sue fatture passive,
    // pregresso compreso, spariscono da QUESTA vista (Resoconto invariato).
    () =>
      prepara(fattureRic ?? []).filter((x) => !x.f.cliente.toLowerCase().includes("dr logistics")),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [fattureRic, termini],
  );

  const voci = useMemo(() => (flussi ?? []).filter((x) => x.genere === "voce"), [flussi]);
  // Preset salvati: righe genere "preset" raggruppate per nome (Title).
  const presets = useMemo(() => {
    const per = new Map<string, FlussoCassaRiga[]>();
    for (const r of flussi ?? []) {
      if (r.genere !== "preset" || !r.note) continue;
      const l = per.get(r.nome) ?? [];
      l.push(r);
      per.set(r.nome, l);
    }
    return [...per.entries()].sort((a, b) => a[0].localeCompare(b[0], "it"));
  }, [flussi]);

  const esclusioni = useMemo(
    () => (flussi ?? []).filter((x) => x.genere === "esclusione"),
    [flussi],
  );
  // Schede fornitore (righe genere "fornitore"): override della modalità
  // di pagamento (RID/RiBa/nessuna) e dati descrittivi — si modificano
  // nella tab Fornitori. Senza override vale l'XML della singola fattura.
  const infoFornitori = useMemo(() => mappaFornitori(flussi ?? []), [flussi]);
  const modDi = (f: SpFattura): ModalitaPagamento => modalitaFattura(f, infoFornitori);
  const simAttiva = simFatture.size > 0 || simSenzaMod || simFiscali.size > 0;
  const nonPagata = (x: { f: SpFattura }) =>
    simFatture.has(x.f.nomeFile) || (simSenzaMod && modDi(x.f) === "altro");

  // --- Colonne periodo -------------------------------------------------------
  const periodi = useMemo(() => {
    const out: { chiave: string; label: string; mese: string }[] = [];
    if (modo === "mese") {
      const base = new Date(`${oggiISO.slice(0, 7)}-01T00:00:00`);
      const nomi = [
        "gen",
        "feb",
        "mar",
        "apr",
        "mag",
        "giu",
        "lug",
        "ago",
        "set",
        "ott",
        "nov",
        "dic",
      ];
      for (let i = 0; i < 6; i++) {
        const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
        const chiave = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
        out.push({ chiave, label: `${nomi[d.getMonth()]} ${d.getFullYear()}`, mese: chiave });
      }
    } else {
      const lun = lunedioDi(oggiISO);
      for (let i = 0; i < 13; i++) {
        const d = new Date(lun.getTime() + i * 7 * 86400000);
        const fine = new Date(d.getTime() + 6 * 86400000);
        const iso = d.toISOString().slice(0, 10);
        const chiave = chiaveSettimana(iso);
        const gg = (x: Date) =>
          `${String(x.getUTCDate()).padStart(2, "0")}/${String(x.getUTCMonth() + 1).padStart(2, "0")}`;
        out.push({
          chiave,
          label: `${chiave.slice(5)} · ${gg(d)}–${gg(fine)}`,
          mese: iso.slice(0, 7),
        });
      }
    }
    // "Fino al": le colonne interamente oltre la data spariscono.
    if (/^\d{4}-\d{2}-\d{2}$/.test(finoA)) {
      const chiaveLimite = modo === "mese" ? finoA.slice(0, 7) : chiaveSettimana(finoA);
      return out.filter((p) => p.chiave <= chiaveLimite);
    }
    return out;
  }, [modo, oggiISO, finoA]);

  const chiaveDi = (scadenzaISO: string) =>
    modo === "mese" ? scadenzaISO.slice(0, 7) : chiaveSettimana(scadenzaISO);

  // --- Esclusioni ------------------------------------------------------------
  const esclusa = (nomeControparte: string, meseScadenza: string): boolean =>
    esclusaDaFlussi(esclusioni, nomeControparte, meseScadenza);

  // Fattura "nel flusso": aperta, con scadenza, controparte non esclusa —
  // a prescindere dai filtri dal/fino al (serve alla checklist della
  // simulazione e ai totali del banner).
  const inFlusso = (x: (typeof attive)[number]): boolean => {
    if (residuoAperto(x) <= 1) return false;
    if (!x.s.scadenza) return false;
    return !esclusa(x.f.cliente, x.s.scadenza.slice(0, 7));
  };

  // --- Somme per controparte -------------------------------------------------
  type RigaCp = {
    nome: string;
    scaduto: number;
    perPeriodo: Map<string, number>;
    totale: number;
    /** Solo uscite: modalità di pagamento del gruppo (RiBa / RID / altro). */
    mod?: ModalitaPagamento;
  };
  type Somma = { righe: RigaCp[]; tot: RigaCp; gruppi: Record<ModalitaPagamento, RigaCp> };
  const nuovaRiga = (nome: string, mod?: ModalitaPagamento): RigaCp => ({
    nome,
    scaduto: 0,
    perPeriodo: new Map(),
    totale: 0,
    mod,
  });
  const gruppiVuoti = (): Record<ModalitaPagamento, RigaCp> => ({
    riba: nuovaRiga(""),
    rid: nuovaRiga(""),
    altro: nuovaRiga(""),
  });
  const ordMod = (m?: ModalitaPagamento) => (m ? MODALITA_ORDINE.indexOf(m) : 0);
  const chiaviPeriodo = useMemo(() => new Set(periodi.map((p) => p.chiave)), [periodi]);
  // USCITE DIVISE PER MODALITÀ (Simone 29/09): con `conModalita` ogni riga
  // fornitore sta nel gruppo della modalità della SUA fattura (RiBa, RID o
  // nessuna delle due) — un fornitore con fatture miste compare in più
  // gruppi. Con `simula` le fatture "non pagate" della simulazione spariscono.
  const somma = (
    righe: typeof attive,
    opts: { conModalita?: boolean; simula?: boolean } = {},
  ): Somma => {
    const per = new Map<string, RigaCp>();
    const tot = nuovaRiga("");
    const gruppi = gruppiVuoti();
    for (const x of righe) {
      const residuo = residuoAperto(x);
      if (residuo <= 1) continue;
      // Anche le "Non gestite" (nessuna lettura) contano: decisione Simone
      // 08/09 — una fattura aperta e' denaro atteso, come nel Resoconto.
      // (Il caso iMile FPR 228/26+230/26: 21.791 fuori dai Flussi ma dentro
      // i ritardi del Resoconto.)
      if (!x.s.scadenza) continue;
      const scad = x.s.scadenza.slice(0, 10);
      if (esclusa(x.f.cliente, scad.slice(0, 7))) continue;
      if (/^\d{4}-\d{2}-\d{2}$/.test(daData) && !x.s.inRitardo && scad < daData) continue;
      if (/^\d{4}-\d{2}-\d{2}$/.test(finoA) && scad > finoA) continue;
      if (opts.simula && nonPagata(x)) continue;
      const mod = opts.conModalita ? modDi(x.f) : undefined;
      const k = (mod ? `${mod}|` : "") + (clienteGroupKey(x.f.cliente) || x.f.cliente);
      const r = per.get(k) ?? nuovaRiga(x.f.cliente, mod);
      const g = mod ? gruppi[mod] : null;
      if (x.s.inRitardo) {
        r.scaduto += residuo;
        tot.scaduto += residuo;
        if (g) g.scaduto += residuo;
      } else {
        const kp = chiaveDi(scad);
        if (!chiaviPeriodo.has(kp)) continue;
        r.perPeriodo.set(kp, (r.perPeriodo.get(kp) ?? 0) + residuo);
        r.totale += residuo;
        tot.perPeriodo.set(kp, (tot.perPeriodo.get(kp) ?? 0) + residuo);
        tot.totale += residuo;
        if (g) {
          g.perPeriodo.set(kp, (g.perPeriodo.get(kp) ?? 0) + residuo);
          g.totale += residuo;
        }
      }
      per.set(k, r);
    }
    return {
      righe: [...per.values()]
        .filter((r) => r.totale > 0 || r.scaduto > 0)
        .sort(
          (a, b) => ordMod(a.mod) - ordMod(b.mod) || b.totale + b.scaduto - (a.totale + a.scaduto),
        ),
      tot,
      gruppi,
    };
  };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const entrate = useMemo(() => somma(attive), [attive, periodi, esclusioni, daData, finoA]);
  const uscite = useMemo(
    () => somma(passive, { conModalita: true, simula: true }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, periodi, esclusioni, daData, finoA, infoFornitori, simFatture, simSenzaMod],
  );
  // Le uscite SENZA simulazione: per il confronto nel banner e nella riga
  // "Delta saldo senza simulazione" (stesso oggetto quando non serve).
  const usciteReali = useMemo(
    () => (simAttiva ? somma(passive, { conModalita: true }) : uscite),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, periodi, esclusioni, daData, finoA, infoFornitori, simAttiva, uscite],
  );

  // --- Tabella "solo fatturazioni" (richiesta Simone 14/09) ------------------
  // Come sopra ma a FATTURATO: ogni fattura pesa per il suo TOTALE (al netto
  // delle note di credito collegate) alla scadenza, incassata/pagata o no —
  // nessuna movimentazione bancaria. Scaduto = scadenza prima di oggi.
  // Spec Simone 14/09 sera: la tabella "solo fatturazioni" è la COPIA di
  // quella sopra — stesso Scaduto di partenza (residui aperti in ritardo,
  // identici riga per riga), stesse esclusioni — ma dai mesi in avanti ogni
  // fattura pesa per il FATTURATO PIENO alla scadenza (al netto delle NC
  // collegate), incassata/pagata o no: i movimenti bancari non scalano nulla.
  const sommaFatturato = (
    righe: typeof attive,
    tutte: SpFattura[],
    opts: { conModalita?: boolean; simula?: boolean } = {},
  ): Somma => {
    const nc = collegaNoteCredito(tutte, fattureEscluse(tutte));
    const per = new Map<string, RigaCp>();
    const tot = nuovaRiga("");
    const gruppi = gruppiVuoti();
    for (const x of righe) {
      if (isNotaCredito(x.f.tipoDocumento)) continue;
      if (!x.s.scadenza) continue;
      const scad = x.s.scadenza.slice(0, 10);
      if (esclusa(x.f.cliente, scad.slice(0, 7))) continue;
      if (/^\d{4}-\d{2}-\d{2}$/.test(finoA) && scad > finoA) continue;
      // Simulazione: una fattura "non pagata" sparisce anche qui (aperta o
      // no, il criterio è la scelta fatta nel pannello).
      if (opts.simula && nonPagata(x)) continue;
      const mod = opts.conModalita ? modDi(x.f) : undefined;
      const k = (mod ? `${mod}|` : "") + (clienteGroupKey(x.f.cliente) || x.f.cliente);
      const r = per.get(k) ?? nuovaRiga(x.f.cliente, mod);
      const g = mod ? gruppi[mod] : null;
      if (x.s.inRitardo) {
        // IDENTICO alla tabella sopra: il punto di partenza è lo stato reale.
        const residuo = residuoAperto(x);
        if (residuo <= 1) continue;
        r.scaduto += residuo;
        tot.scaduto += residuo;
        if (g) g.scaduto += residuo;
      } else {
        if (/^\d{4}-\d{2}-\d{2}$/.test(daData) && scad < daData) continue;
        const kp = chiaveDi(scad);
        if (!chiaviPeriodo.has(kp)) continue;
        const importo = Math.abs(x.f.totale) - (nc.get(x.f.nomeFile)?.importo ?? 0);
        if (importo <= 0.005) continue;
        r.perPeriodo.set(kp, (r.perPeriodo.get(kp) ?? 0) + importo);
        r.totale += importo;
        tot.perPeriodo.set(kp, (tot.perPeriodo.get(kp) ?? 0) + importo);
        tot.totale += importo;
        if (g) {
          g.perPeriodo.set(kp, (g.perPeriodo.get(kp) ?? 0) + importo);
          g.totale += importo;
        }
      }
      per.set(k, r);
    }
    return {
      righe: [...per.values()]
        .filter((r) => r.totale > 0.005 || r.scaduto > 0.005)
        .sort(
          (a, b) => ordMod(a.mod) - ordMod(b.mod) || b.totale + b.scaduto - (a.totale + a.scaduto),
        ),
      tot,
      gruppi,
    };
  };
  const entrateFat = useMemo(
    () => sommaFatturato(attive, fattureEm ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [attive, fattureEm, periodi, esclusioni, daData, finoA],
  );
  const usciteFat = useMemo(
    () => sommaFatturato(passive, fattureRic ?? [], { conModalita: true, simula: true }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [
      passive,
      fattureRic,
      periodi,
      esclusioni,
      daData,
      finoA,
      infoFornitori,
      simFatture,
      simSenzaMod,
    ],
  );
  const usciteFatReali = useMemo(
    () =>
      simAttiva ? sommaFatturato(passive, fattureRic ?? [], { conModalita: true }) : usciteFat,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, fattureRic, periodi, esclusioni, daData, finoA, infoFornitori, simAttiva, usciteFat],
  );

  // GIRATE (spec Simone 12/09, v1.62.0): regole "se entra una fattura dal
  // cliente X, gira il P% al fornitore Y", configurabili dal pannello
  // Esclusioni e salvate sulla lista FlussiCassa (genere "girata": Title =
  // fornitore, Importo = percentuale, Note = "cliente | termini oggetto
  // facoltativi"). Ogni regola genera la SUA riga tra le uscite, intitolata
  // al fornitore, come le altre (niente corsivo/colore).
  // REGIME UNICO "A INCASSO AVVENUTO" (decisione Simone 15/09, estende a
  // tutte le fatture quello nato per la FPR 220/26 il 10/09): la quota
  // matura SOLO sugli incassi reali registrati (da GIRATA_INCASSI_DA in poi)
  // e resta in Scaduto al netto dei bonifici reali già fatti al fornitore —
  // niente quote previsionali sulle scadenze attese nei mesi futuri.
  const GIRATA_INCASSI_DA = "2026-09-01";
  const girate = useMemo(
    () =>
      (flussi ?? [])
        .filter((x) => x.genere === "girata" && (x.importo ?? 0) > 0)
        .map((r) => {
          const [cli, ogg] = String(r.note ?? "").split("|");
          return {
            id: r.id,
            fornitore: r.nome,
            perc: Math.min(100, Math.max(0, r.importo)) / 100,
            cliente: (cli ?? "").trim().toLowerCase(),
            oggettoTermini: (ogg ?? "")
              .split(",")
              .map((s) => s.trim().toLowerCase())
              .filter((s) => s.length >= 3),
          };
        })
        .filter((g) => g.cliente && g.fornitore),
    [flussi],
  );
  const girataQuote = useMemo(() => {
    // UNA riga per FORNITORE: più regole verso lo stesso fornitore sommano
    // le maturate e i bonifici reali si scalano UNA volta sola (scalarli a
    // ogni regola sottostimava il dovuto complessivo).
    const perFornitore = new Map<
      string,
      { fornitore: string; maturata: number; scaduto: number; perPeriodo: Map<string, number> }
    >();
    for (const g of girate) {
      const kf = g.fornitore.trim().toLowerCase();
      let q = perFornitore.get(kf);
      if (!q) {
        q = { fornitore: g.fornitore, maturata: 0, scaduto: 0, perPeriodo: new Map() };
        perFornitore.set(kf, q);
      }
      const matcha = (x: (typeof attive)[number]) => {
        if (!x.f.cliente.toLowerCase().includes(g.cliente)) return false;
        if (!g.oggettoTermini.length) return true;
        const testo = `${x.f.oggetto ?? ""} ${x.f.causaleDoc ?? ""}`.toLowerCase();
        return g.oggettoTermini.every((t2) => testo.includes(t2));
      };
      // Quota MATURATA: percentuale degli incassi reali registrati (dal
      // 01/09 in poi) su ogni fattura che matcha la regola. Le NC compensate
      // entrano in NEGATIVO (incassatoRegistrato = -|totale|): la quota è
      // sul netto, come nel Resoconto. LIMITE NOTO: incassato e dataIncasso
      // sono aggregati per fattura (la data è dell'ULTIMA rata) — una
      // fattura con acconti pre-01/09 e saldo dopo conterebbe tutto; oggi
      // non esiste un caso simile e, se capitasse, si sana scrivendo il
      // valore a mano.
      for (const x of attive) {
        if (!matcha(x)) continue;
        const inc = incassatoRegistrato(x);
        if (inc === 0) continue;
        const dataInc = (x.f.dataIncasso ?? "").slice(0, 10);
        if (!dataInc || dataInc < GIRATA_INCASSI_DA) continue;
        q.maturata += inc * g.perc;
      }
    }
    const out = [...perFornitore.values()];
    for (const q of out) {
      // Meno i bonifici REALI già usciti verso il fornitore. ASSUNTO (vale
      // per DR Logistics, che non ha altre partite aperte con noi): ogni
      // bonifico verso il fornitore estingue quota girata. Per un fornitore
      // che avesse ANCHE fatture passive ordinarie in tabella il netting
      // sovrastimerebbe i pagamenti della quota: da rivedere se si aggiunge
      // una girata verso un fornitore del genere.
      let versato = 0;
      for (const m of movimenti ?? []) {
        if (m.importo >= 0) continue;
        if (m.dataContabile < GIRATA_INCASSI_DA) continue;
        if (!`${m.cliente} ${m.descrizione}`.toLowerCase().includes(q.fornitore.toLowerCase()))
          continue;
        versato += -m.importo;
      }
      q.scaduto = Math.max(0, Math.round((q.maturata - versato) * 100) / 100);
    }
    return out;
  }, [girate, attive, movimenti]);
  const girateScadutoTot = girataQuote.reduce((s, q) => s + q.scaduto, 0);
  const girataTotaleDi = (q: (typeof girataQuote)[number]) =>
    q.scaduto + [...q.perPeriodo.values()].reduce((s, v) => s + v, 0);
  // SCADUTO STIPENDI (richiesta Simone 16/09): i mesi di PAGAMENTO arrivati
  // (corrente compreso) con la spunta "Pagato" in uso portano il residuo non
  // pagato (totale − pagati) nella colonna Scaduto della riga Stipendi ed
  // entrano nel saldo — la cella del mese mostra i soli pagati, quindi
  // scaduto + cella = totale del mese, senza doppi conteggi. Nella tabella
  // "solo fatturazioni" la cella del mese corrente porta già il TOTALE
  // pieno: lì in Scaduto vanno solo i mesi di pagamento GIÀ PASSATI. I mesi
  // passati SENZA spunte restano fuori: storia già regolata in banca.
  const stipendiScaduto = useMemo(() => {
    if (!stipendiMesi) return { reale: 0, fatturato: 0 };
    const corrente = oggiISO.slice(0, 7);
    let passati = 0;
    let resCorrente = 0;
    for (const [mm, st] of Object.entries(stipendiMesi)) {
      if (st.flaggati <= 0) continue;
      const res = Math.max(0, st.totale - st.pagati);
      if (mm < corrente) passati += res;
      else if (mm === corrente) resCorrente += res;
    }
    return {
      reale: Math.round((passati + resCorrente) * 100) / 100,
      fatturato: Math.round(passati * 100) / 100,
    };
  }, [stipendiMesi, oggiISO]);
  // Vista per settimana: le voci mensili non sono renderizzate (né in UI né
  // nel CSV), quindi lo scaduto stipendi resta fuori anche dal saldo — la
  // colonna Scaduto deve sempre quadrare per somma verticale delle righe.
  // SCADENZE FISCALI (v1.85.0, stessa logica degli stipendi): le non pagate
  // già scadute stanno nella colonna Scaduto delle due voci fiscali; il mese
  // corrente vale il reale (pagate del mese + da pagare entro il mese). La
  // simulazione toglie le scadenze scelte; i totali reali restano per il
  // confronto.
  const chiaveFisc = (s: ScadenzaFiscale) => s.id ?? chiaveScadenzaFile(s);
  const scadenzeFiscali = useMemo(
    () => (fiscale?.scadenze ?? []).filter((s) => !simFiscali.has(chiaveFisc(s))),
    [fiscale, simFiscali],
  );
  const totFiscali = useMemo(
    () => (scadenzeFiscali.length > 0 ? ripartizioneFiscale(scadenzeFiscali, oggiISO) : null),
    [scadenzeFiscali, oggiISO],
  );
  const totFiscaliReali = useMemo(
    () =>
      fiscale && fiscale.scadenze.length > 0
        ? ripartizioneFiscale(fiscale.scadenze, oggiISO)
        : null,
    [fiscale, oggiISO],
  );
  const fiscScadutoDi = (nome: string, reale = false): number => {
    const r = reale ? totFiscaliReali : totFiscali;
    if (!r || modo !== "mese") return 0;
    const n = nome.trim().toLowerCase();
    if (n === "costo fiscale rate") return r.scaduto.rate;
    if (n === "costo fiscale corrente") return r.scaduto.corrente;
    return 0;
  };
  const fiscScadutoTot = (reale = false) =>
    fiscScadutoDi("costo fiscale rate", reale) + fiscScadutoDi("costo fiscale corrente", reale);
  const scadutoVoce = (nome: string, fatturato = false) => {
    if (modo !== "mese") return 0;
    if (nome.trim().toLowerCase() === "stipendi")
      return fatturato ? stipendiScaduto.fatturato : stipendiScaduto.reale;
    return fiscScadutoDi(nome);
  };
  const saldoScaduto =
    entrate.tot.scaduto -
    uscite.tot.scaduto -
    girateScadutoTot -
    (modo === "mese" ? stipendiScaduto.reale + fiscScadutoTot() : 0);
  const saldoScadutoReale =
    entrate.tot.scaduto -
    usciteReali.tot.scaduto -
    girateScadutoTot -
    (modo === "mese" ? stipendiScaduto.reale + fiscScadutoTot(true) : 0);

  // --- Prefatture (stessa copertura della Previsione) ------------------------
  const prefPer = useMemo(() => {
    const copertura = (dir: "Emessa" | "Ricevuta") => {
      const fonte = dir === "Emessa" ? (fattureEm ?? []) : (fattureRic ?? []);
      return new Set(
        fonte.map(
          (f) =>
            `${clienteGroupKey(f.cliente) || f.cliente.toLowerCase()}|${f.dataDocumento.slice(0, 7)}`,
        ),
      );
    };
    const mesiVisibili = [...new Set(periodi.map((p) => p.mese))];
    const perDir = (dir: "Emessa" | "Ricevuta") => {
      const cov = copertura(dir);
      const out = new Map<string, number>();
      for (const pf of (prefatture ?? []).filter((x) => x.direzione === dir)) {
        const chiave = clienteGroupKey(pf.controparte) || pf.controparte.toLowerCase();
        const mesiPf =
          pf.ricorrenza === "una"
            ? mesiVisibili.filter((m) => m === pf.meseInizio)
            : mesiVisibili.filter((m) => m >= pf.meseInizio && (!pf.meseFine || m <= pf.meseFine));
        for (const m of mesiPf) {
          if (cov.has(`${chiave}|${m}`)) continue;
          if (esclusa(pf.controparte, m)) continue;
          out.set(m, (out.get(m) ?? 0) + pf.importo);
        }
      }
      return out;
    };
    return { att: perDir("Emessa"), pas: perDir("Ricevuta") };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [prefatture, fattureEm, fattureRic, periodi, esclusioni]);
  // Le prefatture sono MENSILI: nella vista per settimana pesano sul mese
  // ma non si possono spalmare onestamente — compaiono solo per mese.
  const haPref = modo === "mese" && (prefPer.att.size > 0 || prefPer.pas.size > 0);

  // --- Media automatica "Altre spese" ----------------------------------------
  // Richiesta FR 08/09: le spese che NON passano dalle fatture (regole
  // flaggate "Altre spese" con la €) fanno media sugli ultimi 2 MESI PIENI
  // e riempiono da sole la riga — il valore manuale, se inserito, vince.
  // ALTRE SPESE = media degli ultimi 2 mesi COMPLETI dei movimenti "Costi
  // generali" NON fatturati (decisione Simone 15/09): fuori Pagamento
  // Salario, Consulenze e POST EBITDA (gia' coperti da Stipendi, fatture e
  // voci fiscali) e fuori ogni movimento con nr fattura o verso un
  // fornitore che ha fatture passive in archivio. Le spunte del pannello
  // (righe FlussiCassa genere "asvoce": 1 includi / 0 escludi) vincono sul
  // default, tipologia per tipologia.
  // "Imposte / F24" fuori di default (1.85.1): quei movimenti sono già
  // contati dalle voci fiscali (col mese corrente al reale contavano due
  // volte). La spunta nel pannello può sempre riaccenderli.
  const TIP_ESCLUSE_DEFAULT = useMemo(
    () => new Set(["Pagamento Salario", "Consulenze", "POST EBITDA", "Imposte / F24"]),
    [],
  );
  const asOverride = useMemo(() => {
    const m = new Map<string, boolean>();
    for (const r of flussi ?? [])
      if (r.genere === "asvoce") m.set(r.nome.trim(), (r.importo ?? 0) > 0);
    return m;
  }, [flussi]);
  // COSTI FISSI (Simone 29/09, v1.85.0): affitti e uscite fisse impostate
  // nel pannello (righe genere "fisso") escono dalla media e contano al
  // loro importo — reale se nel mese corrente sono già usciti, altrimenti
  // quello impostato. MESE CORRENTE dei variabili = reale già uscito + media
  // per i giorni che mancano (a fine mese coincide col reale).
  const costiFissi = useMemo(() => costiFissiDa(flussi ?? []), [flussi]);
  // Anomalie forzate a mano (righe genere "anomalia": 1 = anomalia, 0 = normale).
  const overrideAnomalie = useMemo(() => {
    const m = new Map<string, boolean>();
    for (const r of flussi ?? []) if (r.genere === "anomalia") m.set(r.nome, (r.importo ?? 0) > 0);
    return m;
  }, [flussi]);
  // Fornitori con fatture in archivio: i loro movimenti non sono "altre
  // spese" (viaggiano già tra le fatture da pagare).
  const fornitoriArchivio = useMemo(
    () =>
      new Set(
        (fattureRic ?? []).map((f) => f.cliente.toLowerCase().trim()).filter((c) => c.length > 6),
      ),
    [fattureRic],
  );
  const fatturata = (m: SpMovimento) => {
    if ((m.nrFattura ?? "").trim()) return true;
    const c = (m.cliente ?? "").toLowerCase().trim();
    if (!c) return false;
    if (fornitoriArchivio.has(c)) return true;
    // La direzione f.includes(c) solo con controparti non-corte: 'TIM'
    // e' sottostringa di mezzo archivio e sparirebbe dalla media in silenzio.
    for (const f of fornitoriArchivio)
      if (c.includes(f) || (c.length > 6 && f.includes(c))) return true;
    return false;
  };
  const aMovimentoMedia = (m: SpMovimento): MovimentoMedia => ({
    chiave: m.chiave,
    data: m.dataContabile.slice(0, 10),
    importo: m.importo,
    tipologia: m.tipologia?.trim() || "(senza tipologia)",
    controparte: m.cliente?.trim() || m.descrizione,
  });
  const autoAltreSpese = useMemo(() => {
    if (!movimenti?.length) return null;
    const meseCorrente = oggiISO.slice(0, 7);
    // Pagamenti dei costi fissi già usciti nel mese corrente (id fisso → movimento).
    const pagatiFissi = new Map<string, { importo: number; data: string }>();
    const base: MovimentoMedia[] = [];
    for (const m of movimenti) {
      if (m.importo >= 0) continue;
      const mm = m.dataContabile.slice(0, 7);
      let fisso = false;
      for (const f of costiFissi) {
        if (!matchFisso(f, m)) continue;
        fisso = true;
        // Nel mese corrente tiene il pagamento più vicino all'importo atteso
        // (Zekaj: l'affitto da 2.000, non la disposizione da 1.000).
        if (mm === meseCorrente && piuVicinoAlFisso(f, pagatiFissi.get(f.id), m))
          pagatiFissi.set(f.id, {
            importo: Math.abs(m.importo),
            data: m.dataContabile.slice(0, 10),
          });
      }
      if (fisso) continue;
      if (!(m.allocPrimaria ?? "").toLowerCase().includes("generali")) continue;
      // Le Consulenze hanno la loro riga (v1.86.0).
      if ((m.tipologia ?? "").trim().toLowerCase() === TIP_CONSULENZE) continue;
      if (fatturata(m)) continue;
      base.push(aMovimentoMedia(m));
    }
    // Media dei due mesi pieni SENZA anomalie; mese corrente al reale.
    const analisi = analizzaMedia(base, oggiISO, overrideAnomalie);
    const r2 = (n: number) => Math.round(n * 100) / 100;
    const righe = analisi.righe.map((r) => ({
      ...r,
      inclusa: asOverride.get(r.tip) ?? !TIP_ESCLUSE_DEFAULT.has(r.tip),
    }));
    const incluse = righe.filter((r) => r.inclusa);
    const media = r2(incluse.reduce((s2, r) => s2 + r.m1 + r.m2, 0) / 2);
    const realeCorrente = r2(incluse.reduce((s2, r) => s2 + r.corrente, 0));
    const tipIncluse = new Set(incluse.map((r) => r.tip));
    // I 15 movimenti più grandi dei tre mesi (tipologie incluse): qui si
    // vedono e si forzano le anomalie.
    const movimentiTop = analisi.movimenti
      .filter((x) => tipIncluse.has(x.m.tipologia))
      .slice(0, 15);
    const mesi = analisi.mesi;
    const pr = proiezioneMeseCorrente(realeCorrente, media, oggiISO);
    const fissiMese = (mese: string) =>
      r2(
        costiFissi
          .filter((f) => fissoAttivo(f, mese))
          .reduce(
            (s2, f) =>
              s2 +
              (mese === meseCorrente ? (pagatiFissi.get(f.id)?.importo ?? f.importo) : f.importo),
            0,
          ),
      );
    const totale = (mese: string) =>
      r2((mese === meseCorrente ? pr.proiezione : media) + fissiMese(mese));
    return {
      media,
      mesi,
      righe,
      realeCorrente,
      proiezione: pr.proiezione,
      giornoOggi: pr.giornoOggi,
      giorniMese: pr.giorniMese,
      pagatiFissi,
      fissiMese,
      totale,
      movimentiTop,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    movimenti,
    fornitoriArchivio,
    asOverride,
    TIP_ESCLUSE_DEFAULT,
    oggiISO,
    costiFissi,
    overrideAnomalie,
  ]);

  // CONSULENZE (v1.86.0): pagamenti con tipologia Consulenze non fatturati.
  // Mese corrente = reale (anomalie comprese) + media per i giorni che
  // mancano; mesi futuri = media dei due mesi pieni SENZA anomalie.
  const autoConsulenze = useMemo(() => {
    if (!movimenti?.length) return null;
    // Fuori i pagamenti dei costi fissi (l'affitto Zekaj da 2.000 al mese
    // è classificato Consulenze in banca: conta già tra i fissi delle
    // Altre spese, qui contava due volte — 1.86.1).
    const base = movimenti
      .filter(
        (m) =>
          m.importo < 0 &&
          (m.tipologia ?? "").trim().toLowerCase() === TIP_CONSULENZE &&
          !fatturata(m) &&
          !costiFissi.some((f) => matchFisso(f, m)),
      )
      .map(aMovimentoMedia);
    const analisi = analizzaMedia(base, oggiISO, overrideAnomalie);
    const r2 = (n: number) => Math.round(n * 100) / 100;
    const m1 = r2(analisi.righe.reduce((s2, r) => s2 + r.m1, 0));
    const m2 = r2(analisi.righe.reduce((s2, r) => s2 + r.m2, 0));
    const media = r2((m1 + m2) / 2);
    const realeCorrente = r2(analisi.righe.reduce((s2, r) => s2 + r.corrente, 0));
    const anomalieCorrente = r2(analisi.righe.reduce((s2, r) => s2 + r.correnteAnomalie, 0));
    const pr = proiezioneMeseCorrente(realeCorrente, media, oggiISO);
    return {
      mesi: analisi.mesi,
      m1,
      m2,
      media,
      realeCorrente,
      anomalieCorrente,
      proiezione: pr.proiezione,
      giornoOggi: pr.giornoOggi,
      giorniMese: pr.giorniMese,
      movimenti: analisi.movimenti,
      soglia: [...analisi.soglie.values()][0] ?? 0,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [movimenti, fornitoriArchivio, oggiISO, overrideAnomalie, costiFissi]);

  const toggleAnomalia = async (x: MovimentoAnalizzato) => {
    setDrillBusy(true);
    try {
      await spUpsertFlussoCassa({
        data: { nome: x.m.chiave, genere: "anomalia", importo: x.anomalia ? 0 : 1 },
      });
      await ricaricaFlussi();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setDrillBusy(false);
    }
  };

  const aggiungiFisso = async () => {
    const importo = Number(fxImporto.trim().replace(/\./g, "").replace(",", "."));
    if (!fxNome.trim() || !Number.isFinite(importo) || importo <= 0) return;
    setFxBusy(true);
    try {
      await spUpsertFlussoCassa({
        data: {
          nome: fxNome.trim(),
          genere: "fisso",
          importo: Math.round(importo * 100) / 100,
          note: (fxToken.trim() || fxNome.trim()).toLowerCase(),
          mese: fxDa || undefined,
          meseFine: fxA || undefined,
        },
      });
      setFxNome("");
      setFxImporto("");
      setFxToken("");
      setFxDa("");
      setFxA("");
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setFxBusy(false);
    }
  };

  // --- Voci manuali (mensili) ------------------------------------------------
  const nomiVoci = useMemo(() => {
    const set = new Set<string>(VOCI_BASE);
    for (const v of voci) set.add(v.nome);
    return [...set];
  }, [voci]);
  const vocePer = (nome: string, mese: string): FlussoCassaRiga | undefined =>
    voci.find((v) => v.nome.trim().toLowerCase() === nome.trim().toLowerCase() && v.mese === mese);

  // Scadenze fiscali simulabili: le non pagate, finanziamenti esclusi (quelli
  // non stanno nelle voci fiscali), in ordine di data.
  const fiscaliSimulabili = useMemo(
    () =>
      (fiscale?.scadenze ?? [])
        .filter((s) => !s.pagato && s.categoria !== "finanziamento")
        .sort((a, b) => a.dataPagamento.localeCompare(b.dataPagamento)),
    [fiscale],
  );
  // Fatture passive nel flusso, per la checklist della simulazione.
  const fattureSimulabili = useMemo(
    () =>
      passive
        .filter(inFlusso)
        .map((x) => ({
          x,
          mod: modDi(x.f),
          residuo: residuoAperto(x),
          scad: x.s.scadenza.slice(0, 10),
        }))
        .sort(
          (a, b) =>
            a.x.f.cliente.localeCompare(b.x.f.cliente, "it") || a.scad.localeCompare(b.scad),
        ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, esclusioni, infoFornitori],
  );
  // Cosa toglie la simulazione, per il banner.
  const simEscluse = useMemo(() => {
    let nFatture = 0;
    let importoFatture = 0;
    if (simAttiva)
      for (const r of fattureSimulabili)
        if (nonPagata(r.x)) {
          nFatture++;
          importoFatture += r.residuo;
        }
    let nFiscali = 0;
    let importoFiscali = 0;
    for (const s of fiscaliSimulabili)
      if (simFiscali.has(chiaveFisc(s))) {
        nFiscali++;
        importoFiscali += s.importo;
      }
    return { nFatture, importoFatture, nFiscali, importoFiscali };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fattureSimulabili, fiscaliSimulabili, simAttiva, simFatture, simSenzaMod, simFiscali]);
  const azzeraSim = () => {
    setSimFatture(new Set());
    setSimSenzaMod(false);
    setSimFiscali(new Set());
  };

  /** Valore effettivo di una voce nel mese: manuale se c'e', altrimenti gli
   *  automatici — Altre spese dalla media dei costi generali non fatturati,
   *  Stipendi dai netti/spunte, voci fiscali dallo scadenziario. */
  const valoreVoce = (
    nome: string,
    mese: string,
    fatturato = false,
    reale = false,
  ): { importo: number; auto: boolean } | null => {
    const man = vocePer(nome, mese);
    if (man) return { importo: man.importo, auto: false };
    // Voci fiscali dallo scadenziario di Sabrina (importi ESATTI, non stime:
    // il "≈" segnala solo che arrivano in automatico dal file). `reale` =
    // senza simulazione (per la riga di confronto).
    const nomeVoce = nome.trim().toLowerCase();
    const fisc = reale ? totFiscaliReali : totFiscali;
    if (fisc && mese >= oggiISO.slice(0, 7)) {
      const tot = fisc.perMese.get(mese);
      if (nomeVoce === "costo fiscale rate" && tot && tot.rate > 0)
        return { importo: -tot.rate, auto: true };
      if (nomeVoce === "costo fiscale corrente" && tot && tot.corrente > 0)
        return { importo: -tot.corrente, auto: true };
    }
    // Altre spese: media dei variabili (mese corrente: reale + media per i
    // giorni che mancano) più i costi fissi al loro importo.
    if (nomeVoce === "altre spese" && autoAltreSpese && mese >= oggiISO.slice(0, 7)) {
      const v = autoAltreSpese.totale(mese);
      if (v > 0) return { importo: -v, auto: true };
    }
    // Consulenze: mese corrente al reale (+ media per i giorni che mancano),
    // mesi futuri la media senza anomalie.
    if (nomeVoce === TIP_CONSULENZE && autoConsulenze && mese >= oggiISO.slice(0, 7)) {
      const v = mese === oggiISO.slice(0, 7) ? autoConsulenze.proiezione : autoConsulenze.media;
      if (v > 0) return { importo: -v, auto: true };
    }
    // Stipendi futuri senza dato reale: stima = media del netto dovuto degli
    // ultimi 2 mesi di "Stipendi Dr.xlsx" (il dato vero, quando arriva
    // dall'import, vince perché è una voce manuale).
    if (nome.trim().toLowerCase() === "stipendi" && mese >= oggiISO.slice(0, 7)) {
      // Dati veri del mese di pagamento (da "Stipendi Dr" + spunte Pagato):
      // tabella reale = somma dei soli segnati "pagato si'" (appena esiste
      // almeno una spunta), "solo fatturazioni" = saldo totale del mese.
      const st = stipendiMesi?.[mese];
      if (st && st.totale > 0) {
        const importo = fatturato || st.flaggati === 0 ? st.totale : st.pagati;
        return { importo: -importo, auto: true };
      }
      if (autoStipendi && autoStipendi.media > 0)
        return { importo: -autoStipendi.media, auto: true };
    }
    return null;
  };

  const salvaVoce = async (nome: string, mese: string) => {
    const grezzo = cellaVal.trim().replace(/\./g, "").replace(",", ".");
    setCellaVoce(null);
    const esistente = vocePer(nome, mese);
    const importo = grezzo === "" ? 0 : Number(grezzo);
    if (!Number.isFinite(importo)) {
      toast.error(t("fc.importoNonValido"));
      return;
    }
    setSalvando(true);
    try {
      if (importo === 0) {
        if (esistente) await spDeleteFlussoCassa({ data: { id: esistente.id } });
      } else {
        await spUpsertFlussoCassa({ data: { nome, genere: "voce", mese, importo } });
      }
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSalvando(false);
    }
  };

  // --- Import scadenziario fiscale ------------------------------------------
  const onFileFiscale = async (f: File) => {
    setParsingF(true);
    setPreviewFisc(null);
    try {
      const XLSX = await import("xlsx");
      const wb = XLSX.read(await f.arrayBuffer(), { cellDates: false });
      const fogli = wb.SheetNames.map((nome) => ({
        nome,
        matrix: XLSX.utils.sheet_to_json(wb.Sheets[nome], {
          header: 1,
          raw: true,
          defval: null,
        }) as unknown[][],
      }));
      const res = parseScadenzario(fogli);
      if (!res) {
        toast.error(t("fc.errFiscale"));
        return;
      }
      setPreviewFisc({ fileName: f.name, res });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsingF(false);
    }
  };

  const confermaFiscale = async () => {
    if (!previewFisc) return;
    setSavingF(true);
    try {
      const res = await spFiscaleSalva({
        data: {
          scadenze: previewFisc.res.scadenze,
          daRateizzare: previewFisc.res.daRateizzare,
          fonte: previewFisc.fileName,
        },
      });
      setFiscale(res);
      setPreviewFisc(null);
      setShowFisc(false);
      toast.success(t("fc.fiscaleOk"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSavingF(false);
    }
  };

  const apriDrill = (voce: string) => {
    const mese = periodi[0]?.mese ?? oggiISO.slice(0, 7);
    setDrill({ voce, mese });
    if (voce.trim().toLowerCase() === "stipendi" && !drillStip) {
      spStipendiGet()
        .then((db) => setDrillStip(db))
        .catch((err) => {
          setDrillStip(null);
          toast.error(err instanceof Error ? err.message : String(err));
        });
    }
  };
  const toggleAsVoce = async (tip: string, inclusa: boolean) => {
    setDrillBusy(true);
    try {
      await spUpsertFlussoCassa({
        data: { nome: tip, genere: "asvoce", importo: inclusa ? 0 : 1 },
      });
      await ricaricaFlussi();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setDrillBusy(false);
    }
  };

  const escludiSelezionate = async () => {
    if (exSel.size === 0) return;
    setExBusy(true);
    try {
      // La finestra di mesi (facoltativa) vale per tutte le spunte del giro.
      for (const nome of exSel) {
        await spUpsertFlussoCassa({
          data: {
            nome,
            genere: "esclusione",
            mese: exDa || undefined,
            meseFine: exA || undefined,
            importo: 0,
          },
        });
      }
      setExSel(new Set());
      setExDa("");
      setExA("");
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setExBusy(false);
    }
  };

  const rimuoviRiga = async (id: string) => {
    try {
      await spDeleteFlussoCassa({ data: { id } });
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    }
  };

  // --- Preset di esclusioni --------------------------------------------------
  // Salva l'insieme CORRENTE di esclusioni sotto un nome (sovrascrivendo un
  // eventuale preset omonimo); applicare un preset SOSTITUISCE le esclusioni.
  const salvaPreset = async () => {
    const nome = presetNome.trim();
    if (!nome) return;
    if (esclusioni.length === 0) {
      toast.error(t("fc.presetVuoto"));
      return;
    }
    setPresetBusy(true);
    try {
      const vecchie = (flussi ?? []).filter(
        (r) => r.genere === "preset" && r.nome.trim().toLowerCase() === nome.toLowerCase(),
      );
      for (const r of vecchie) await spDeleteFlussoCassa({ data: { id: r.id } });
      for (const e of esclusioni) {
        await spUpsertFlussoCassa({
          data: {
            nome,
            genere: "preset",
            mese: e.mese,
            meseFine: e.meseFine,
            importo: 0,
            note: e.nome,
          },
        });
      }
      setPresetNome("");
      await ricaricaFlussi();
      toast.success(t("fc.presetSalvato"));
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setPresetBusy(false);
    }
  };

  const applicaPreset = async (nome: string, righe: FlussoCassaRiga[]) => {
    if (!window.confirm(`${t("fc.presetConfirm")} “${nome}”?`)) return;
    setPresetBusy(true);
    try {
      for (const e of esclusioni) await spDeleteFlussoCassa({ data: { id: e.id } });
      for (const r of righe) {
        await spUpsertFlussoCassa({
          data: {
            nome: r.note ?? "",
            genere: "esclusione",
            mese: r.mese,
            meseFine: r.meseFine,
            importo: 0,
          },
        });
      }
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setPresetBusy(false);
    }
  };

  const aggiungiGirata = async () => {
    const fornitore = (giFornSel === "altro" ? giFornAltro : giFornSel).trim();
    const perc = Number(giPerc.replace(",", "."));
    if (!fornitore || !giCliente.trim() || !Number.isFinite(perc) || perc <= 0 || perc > 100)
      return;
    setGiBusy(true);
    try {
      await spUpsertFlussoCassa({
        data: {
          nome: fornitore,
          genere: "girata",
          importo: Math.round(perc * 100) / 100,
          note: giCliente.trim() + (giOggetto.trim() ? ` | ${giOggetto.trim()}` : ""),
        },
      });
      setGiCliente("");
      setGiOggetto("");
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setGiBusy(false);
    }
  };

  const eliminaPreset = async (righe: FlussoCassaRiga[]) => {
    setPresetBusy(true);
    try {
      for (const r of righe) await spDeleteFlussoCassa({ data: { id: r.id } });
      await ricaricaFlussi();
    } catch (err) {
      toast.error(t("common.error"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setPresetBusy(false);
    }
  };

  // --- Saldo -----------------------------------------------------------------
  // `reale` = senza simulazione (riga di confronto quando la simulazione è
  // accesa): uscite piene e voci fiscali piene.
  const saldoDi = (chiave: string, mese: string, reale = false): number => {
    const u = reale ? usciteReali : uscite;
    let v =
      (entrate.tot.perPeriodo.get(chiave) ?? 0) -
      (u.tot.perPeriodo.get(chiave) ?? 0) -
      girataQuote.reduce((s, q) => s + (q.perPeriodo.get(chiave) ?? 0), 0);
    if (modo === "mese") {
      v += (prefPer.att.get(mese) ?? 0) - (prefPer.pas.get(mese) ?? 0);
      for (const nome of nomiVoci) v += valoreVoce(nome, mese, false, reale)?.importo ?? 0;
    }
    return Math.round(v * 100) / 100;
  };

  // Saldo della tabella "solo fatturazioni":
  // STESSA formula del saldo sopra (girate, prefatture e voci comprese),
  // cambiano solo entrate/uscite (a fatturato pieno). Lo Scaduto coincide
  // con quello sopra per costruzione.
  const saldoFatDi = (chiave: string, mese: string, reale = false): number => {
    const u = reale ? usciteFatReali : usciteFat;
    let v =
      (entrateFat.tot.perPeriodo.get(chiave) ?? 0) -
      (u.tot.perPeriodo.get(chiave) ?? 0) -
      girataQuote.reduce((s, q) => s + (q.perPeriodo.get(chiave) ?? 0), 0);
    if (modo === "mese") {
      v += (prefPer.att.get(mese) ?? 0) - (prefPer.pas.get(mese) ?? 0);
      for (const nome of nomiVoci) v += valoreVoce(nome, mese, true, reale)?.importo ?? 0;
    }
    return Math.round(v * 100) / 100;
  };
  const saldoFatScaduto =
    entrateFat.tot.scaduto -
    usciteFat.tot.scaduto -
    girateScadutoTot -
    (modo === "mese" ? stipendiScaduto.fatturato + fiscScadutoTot() : 0);
  const saldoFatScadutoReale =
    entrateFat.tot.scaduto -
    usciteFatReali.tot.scaduto -
    girateScadutoTot -
    (modo === "mese" ? stipendiScaduto.fatturato + fiscScadutoTot(true) : 0);
  const labelMod: Record<ModalitaPagamento, string> = {
    riba: t("fc.modRiba"),
    rid: t("fc.modRid"),
    altro: t("fc.modAltro"),
  };
  const badgeMod: Record<ModalitaPagamento, string> = {
    riba: "RiBa",
    rid: "RID",
    altro: "—",
  };

  const fmt = (v: number) => (Math.abs(v) >= 0.005 ? `${fmtImporto(v)}` : "—");

  // Valori di riga per colonna: normali, oppure PROGRESSIVI partendo dallo
  // scaduto quando il Cumulato è acceso.
  const serie = (scad: number, get: (chiave: string, mese: string) => number): number[] => {
    let acc = scad;
    return periodi.map((p) => {
      const v = get(p.chiave, p.mese);
      if (!cumulato) return v;
      acc += v;
      return acc;
    });
  };

  // --- Export ----------------------------------------------------------------
  const esporta = () => {
    const testata = [
      t("fc.colVoce"),
      t("fc.colScaduto"),
      ...periodi.map((p) => p.label),
      t("fc.colTotale"),
    ];
    const num = (v: number) => (Math.abs(v) >= 0.005 ? v.toFixed(2).replace(".", ",") : "");
    const righe: string[][] = [];
    if (simAttiva)
      righe.push([
        `${t("fc.simAttiva")} ${simEscluse.nFatture} ${t("fc.simFattureN")} (${num(simEscluse.importoFatture)}) · ${simEscluse.nFiscali} ${t("fc.simFiscaliN")} (${num(simEscluse.importoFiscali)})`,
      ]);
    // Uscite: totale, poi un blocco per modalità (RiBa / RID / senza) con i
    // fornitori del blocco sotto.
    const pushUscite = (u: Somma) => {
      righe.push([
        t("fc.uscite"),
        num(-u.tot.scaduto),
        ...periodi.map((p) => num(-(u.tot.perPeriodo.get(p.chiave) ?? 0))),
        num(-(u.tot.scaduto + u.tot.totale)),
      ]);
      for (const mod of MODALITA_ORDINE) {
        const g = u.gruppi[mod];
        if (g.scaduto + g.totale <= 0.005) continue;
        righe.push([
          `  ${labelMod[mod]}`,
          num(-g.scaduto),
          ...periodi.map((p) => num(-(g.perPeriodo.get(p.chiave) ?? 0))),
          num(-(g.scaduto + g.totale)),
        ]);
        for (const r of u.righe)
          if (r.mod === mod)
            righe.push([
              `    ${r.nome}`,
              num(-r.scaduto),
              ...periodi.map((p) => num(-(r.perPeriodo.get(p.chiave) ?? 0))),
              num(-(r.scaduto + r.totale)),
            ]);
      }
    };
    righe.push([
      t("fc.entrate"),
      num(entrate.tot.scaduto),
      ...periodi.map((p) => num(entrate.tot.perPeriodo.get(p.chiave) ?? 0)),
      num(entrate.tot.scaduto + entrate.tot.totale),
    ]);
    for (const r of entrate.righe)
      righe.push([
        `  ${r.nome}`,
        num(r.scaduto),
        ...periodi.map((p) => num(r.perPeriodo.get(p.chiave) ?? 0)),
        num(r.scaduto + r.totale),
      ]);
    pushUscite(uscite);
    for (const q of girataQuote)
      if (girataTotaleDi(q) > 0.005)
        righe.push([
          q.fornitore,
          num(-q.scaduto),
          ...periodi.map((p) => num(-(q.perPeriodo.get(p.chiave) ?? 0))),
          num(-girataTotaleDi(q)),
        ]);
    if (haPref) {
      righe.push([
        t("fc.prefAtt"),
        "",
        ...periodi.map((p) => num(prefPer.att.get(p.mese) ?? 0)),
        num([...prefPer.att.values()].reduce((s, v) => s + v, 0)),
      ]);
      righe.push([
        t("fc.prefPas"),
        "",
        ...periodi.map((p) => num(-(prefPer.pas.get(p.mese) ?? 0))),
        num(-[...prefPer.pas.values()].reduce((s, v) => s + v, 0)),
      ]);
    }
    if (modo === "mese")
      for (const nome of nomiVoci)
        righe.push([
          nome,
          scadutoVoce(nome) > 0 ? num(-scadutoVoce(nome)) : "",
          ...periodi.map((p) => num(valoreVoce(nome, p.mese)?.importo ?? 0)),
          num(
            periodi.reduce(
              (s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0),
              -scadutoVoce(nome),
            ),
          ),
        ]);
    righe.push([
      t("fc.saldo"),
      num(saldoScaduto),
      ...periodi.map((p) => num(saldoDi(p.chiave, p.mese))),
      num(periodi.reduce((s, p) => s + saldoDi(p.chiave, p.mese), 0)),
    ]);
    if (simAttiva)
      righe.push([
        t("fc.simSaldoReale"),
        num(saldoScadutoReale),
        ...periodi.map((p) => num(saldoDi(p.chiave, p.mese, true))),
        num(periodi.reduce((s, p) => s + saldoDi(p.chiave, p.mese, true), 0)),
      ]);
    // Sezione "solo fatturazioni" (fatturato pieno, incassate/pagate comprese).
    righe.push([]);
    righe.push([t("fc.fatTitolo")]);
    righe.push([
      t("fc.entrate"),
      num(entrateFat.tot.scaduto),
      ...periodi.map((p) => num(entrateFat.tot.perPeriodo.get(p.chiave) ?? 0)),
      num(entrateFat.tot.scaduto + entrateFat.tot.totale),
    ]);
    for (const r of entrateFat.righe)
      righe.push([
        `  ${r.nome}`,
        num(r.scaduto),
        ...periodi.map((p) => num(r.perPeriodo.get(p.chiave) ?? 0)),
        num(r.scaduto + r.totale),
      ]);
    pushUscite(usciteFat);
    // Girate, prefatture e voci: identiche alla sezione sopra (fanno parte
    // anche del saldo "solo fatturazioni").
    for (const q of girataQuote)
      if (girataTotaleDi(q) > 0.005)
        righe.push([
          q.fornitore,
          num(-q.scaduto),
          ...periodi.map((p) => num(-(q.perPeriodo.get(p.chiave) ?? 0))),
          num(-girataTotaleDi(q)),
        ]);
    if (haPref) {
      righe.push([
        t("fc.prefAtt"),
        "",
        ...periodi.map((p) => num(prefPer.att.get(p.mese) ?? 0)),
        num([...prefPer.att.values()].reduce((s, v) => s + v, 0)),
      ]);
      righe.push([
        t("fc.prefPas"),
        "",
        ...periodi.map((p) => num(-(prefPer.pas.get(p.mese) ?? 0))),
        num(-[...prefPer.pas.values()].reduce((s, v) => s + v, 0)),
      ]);
    }
    if (modo === "mese")
      for (const nome of nomiVoci)
        righe.push([
          nome,
          scadutoVoce(nome, true) > 0 ? num(-scadutoVoce(nome, true)) : "",
          ...periodi.map((p) => num(valoreVoce(nome, p.mese, true)?.importo ?? 0)),
          num(
            periodi.reduce(
              (s, p) => s + (valoreVoce(nome, p.mese, true)?.importo ?? 0),
              -scadutoVoce(nome, true),
            ),
          ),
        ]);
    righe.push([
      t("fc.saldo"),
      num(saldoFatScaduto),
      ...periodi.map((p) => num(saldoFatDi(p.chiave, p.mese))),
      num(periodi.reduce((s, p) => s + saldoFatDi(p.chiave, p.mese), 0)),
    ]);
    if (simAttiva)
      righe.push([
        t("fc.simSaldoReale"),
        num(saldoFatScadutoReale),
        ...periodi.map((p) => num(saldoFatDi(p.chiave, p.mese, true))),
        num(periodi.reduce((s, p) => s + saldoFatDi(p.chiave, p.mese, true), 0)),
      ]);
    esportaCsvFile(`flussi-di-cassa-${modo}${simAttiva ? "-simulazione" : ""}`, testata, righe);
  };

  // Checklist fatture della simulazione: filtro di ricerca, tetto a 400.
  const visibiliSim = useMemo(() => {
    const cerca = simCerca.trim().toLowerCase();
    const l = cerca
      ? fattureSimulabili.filter((r) =>
          `${r.x.f.cliente} ${r.x.f.numero}`.toLowerCase().includes(cerca),
        )
      : fattureSimulabili;
    return { righe: l.slice(0, 400), oltre: Math.max(0, l.length - 400) };
  }, [fattureSimulabili, simCerca]);
  const toggleIn = (set: (f: (s: Set<string>) => Set<string>) => void, k: string) =>
    set((s) => {
      const ns = new Set(s);
      if (ns.has(k)) ns.delete(k);
      else ns.add(k);
      return ns;
    });

  const loading = fattureEm == null || fattureRic == null || flussi == null;

  // Fornitori per la tendina/autocomplete delle girate (dalle passive).
  // NIENTE slice qui: un taglio alfabetico lasciava fuori tutto dopo la S
  // (caso Univex, 15/09) — è il browser a filtrare il datalist mentre scrivi.
  const fornitoriNote = useMemo(() => {
    const set = new Set<string>();
    for (const x of passive) set.add(x.f.cliente);
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [passive]);

  // Controparti per la checklist del form esclusioni (tutte: il tetto per
  // non gonfiare il DOM sta a valle, DOPO il filtro di ricerca).
  const contropartiNote = useMemo(() => {
    const set = new Set<string>();
    for (const x of [...attive, ...passive]) set.add(x.f.cliente);
    return [...set].sort((a, b) => a.localeCompare(b));
  }, [attive, passive]);

  // Checklist: fuori le controparti GIA' coperte da un'esclusione (con
  // qualunque finestra di mesi — i chip sopra restano il posto per gestirle)
  // e quelle che non passano il filtro di ricerca.
  const contropartiEscludibili = useMemo(() => {
    const cerca = exCerca.trim().toLowerCase();
    const tutte = contropartiNote.filter((c) => {
      const chiave = clienteGroupKey(c) || c.toLowerCase();
      const giaEsclusa = esclusioni.some((e) => {
        const token = clienteGroupKey(e.nome) || e.nome.trim().toLowerCase();
        return !!token && chiave.includes(token);
      });
      if (giaEsclusa) return false;
      return !cerca || c.toLowerCase().includes(cerca);
    });
    // Il tetto arriva DOPO la ricerca: cercando, si trova sempre tutto.
    return { visibili: tutte.slice(0, 400), oltre: Math.max(0, tutte.length - 400) };
  }, [contropartiNote, esclusioni, exCerca]);

  const inputCls =
    "rounded-lg border border-border bg-background px-2 py-1 text-[13px] text-foreground";

  const cellaVoceUI = (nome: string, mese: string) => {
    const chiave = `${nome}|${mese}`;
    const riga = vocePer(nome, mese);
    const val = valoreVoce(nome, mese);
    if (cellaVoce === chiave)
      return (
        <input
          autoFocus
          value={cellaVal}
          onChange={(e) => setCellaVal(e.target.value)}
          onBlur={() => void salvaVoce(nome, mese)}
          onKeyDown={(e) => {
            if (e.key === "Enter") void salvaVoce(nome, mese);
            if (e.key === "Escape") setCellaVoce(null);
          }}
          className="w-24 rounded border border-primary bg-background px-1 py-0.5 text-right text-[12px]"
        />
      );
    const fiscaleAuto = ["costo fiscale rate", "costo fiscale corrente"].includes(
      nome.trim().toLowerCase(),
    );
    return (
      <button
        type="button"
        title={val?.auto ? t(fiscaleAuto ? "fc.autoTipFiscale" : "fc.autoTip") : t("fc.cellaTip")}
        onClick={() => {
          setCellaVoce(chiave);
          setCellaVal(riga ? String(riga.importo).replace(".", ",") : "");
        }}
        className={`w-full rounded px-1 text-right tabular-nums hover:bg-muted ${val?.auto ? "italic text-muted-foreground" : ""}`}
      >
        {val ? `${val.auto ? "≈ " : ""}${fmt(val.importo)}` : "·"}
      </button>
    );
  };

  const thCls = "py-1 pr-3 text-right whitespace-nowrap";
  const tdN = "py-1 pr-3 text-right tabular-nums whitespace-nowrap";

  // "+" che apre/chiude le voci dei fornitori di un gruppo (Simone 29/09).
  const btnPiu = (k: string, n: number) =>
    n > 0 ? (
      <button
        type="button"
        onClick={() => toggleGruppo(k)}
        title={aperti.has(k) ? t("fc.chiudiGruppo") : t("fc.apriGruppo")}
        className="mr-1.5 inline-flex h-4 w-4 items-center justify-center rounded border border-border align-middle text-[11px] leading-none text-muted-foreground hover:bg-muted hover:text-foreground"
      >
        {aperti.has(k) ? "−" : "+"}
      </button>
    ) : null;
  // Tabella dei movimenti con la spunta "anomalia" (Altre spese e Consulenze).
  const tabellaAnomalie = (lista: MovimentoAnalizzato[]) => (
    <table className="w-full text-[13px]">
      <thead>
        <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
          <th className="py-1 pr-3">{t("fis.colData")}</th>
          <th className="py-1 pr-3">{t("fc.anomControparte")}</th>
          <th className="py-1 pr-3">{t("fc.anomTipologia")}</th>
          <th className="py-1 pr-3 text-right">{t("fis.colImporto")}</th>
          <th className="py-1 pr-3 text-right">{t("fc.anomSoglia")}</th>
          <th className="py-1 text-center">{t("fc.anomColAnomalia")}</th>
        </tr>
      </thead>
      <tbody>
        {lista.map((x) => (
          <tr
            key={x.m.chiave}
            className={`border-b border-border/40 ${x.anomalia ? "text-status-absent" : ""}`}
          >
            <td className="py-0.5 pr-3 whitespace-nowrap">{dataIt(x.m.data)}</td>
            <td className="max-w-56 truncate py-0.5 pr-3" title={x.m.controparte}>
              {x.m.controparte}
            </td>
            <td className="py-0.5 pr-3 text-muted-foreground">{x.m.tipologia}</td>
            <td className="py-0.5 pr-3 text-right tabular-nums">
              {fmtImporto(Math.abs(x.m.importo))}
            </td>
            <td className="py-0.5 pr-3 text-right tabular-nums text-muted-foreground">
              {fmtImporto(x.soglia)}
            </td>
            <td className="py-0.5 text-center whitespace-nowrap">
              <input
                type="checkbox"
                className="accent-primary"
                checked={x.anomalia}
                disabled={drillBusy}
                title={x.auto ? t("fc.anomAuto") : t("fc.anomForzata")}
                onChange={() => void toggleAnomalia(x)}
              />
              {!x.auto && <span className="ml-1 text-[10px]">✎</span>}
            </td>
          </tr>
        ))}
        {lista.length === 0 && (
          <tr>
            <td colSpan={6} className="py-2 text-center text-muted-foreground">
              {t("fc.anomNessuna")}
            </td>
          </tr>
        )}
      </tbody>
    </table>
  );
  // Uscite divise per modalità: riga di gruppo (RiBa / RID / senza) e, se
  // il gruppo è aperto col "+", i fornitori del gruppo sotto. Vale per
  // entrambe le tabelle (`prefix` tiene distinte le chiavi React).
  const righeUscite = (u: Somma, prefix: string) =>
    MODALITA_ORDINE.map((mod) => {
      const g = u.gruppi[mod];
      if (g.scaduto + g.totale <= 0.005) return null;
      return (
        <Fragment key={`${prefix}g:${mod}`}>
          <tr className="border-t border-border/40 text-[12px] font-medium">
            <td className="py-0.5 pl-3 pr-3 text-foreground/80" title={t("fc.modTip")}>
              {btnPiu(`${prefix}:${mod}`, u.righe.filter((r) => r.mod === mod).length)}
              {labelMod[mod]}
            </td>
            <td className={`${tdN} text-status-absent`}>{fmt(-g.scaduto)}</td>
            {serie(g.scaduto, (c) => g.perPeriodo.get(c) ?? 0).map((v, i) => (
              <td key={periodi[i].chiave} className={tdN}>
                {fmt(-v)}
              </td>
            ))}
            <td className={tdN}>{fmt(-(g.scaduto + g.totale))}</td>
          </tr>
          {aperti.has(`${prefix}:${mod}`) &&
            u.righe
              .filter((r) => r.mod === mod)
              .map((r) => (
                <tr key={`${prefix}u:${mod}:${r.nome}`} className="border-t border-border/30">
                  <td className="max-w-56 truncate py-0.5 pl-6 pr-3 text-muted-foreground">
                    {r.nome}
                  </td>
                  <td className={`${tdN} text-muted-foreground`}>{fmt(-r.scaduto)}</td>
                  {serie(r.scaduto, (c) => r.perPeriodo.get(c) ?? 0).map((v, i) => (
                    <td key={periodi[i].chiave} className={`${tdN} text-muted-foreground`}>
                      {fmt(-v)}
                    </td>
                  ))}
                  <td className={`${tdN} text-muted-foreground`}>{fmt(-(r.scaduto + r.totale))}</td>
                </tr>
              ))}
        </Fragment>
      );
    });
  // Riga "Delta saldo senza simulazione": compare sotto il saldo quando la
  // simulazione è accesa, per vedere di quanto cambia mese per mese.
  const rigaConfronto = (scad: number, get: (c: string, m: string) => number) =>
    simAttiva ? (
      <tr className="border-t border-border/40 text-[12px] italic text-muted-foreground">
        <td className="py-1 pr-3">{t("fc.simSaldoReale")}</td>
        <td className={tdN}>{fmt(scad)}</td>
        {serie(scad, get).map((v, i) => (
          <td key={periodi[i].chiave} className={tdN}>
            {fmt(v)}
          </td>
        ))}
        <td className={tdN}>{fmt(periodi.reduce((s, p) => s + get(p.chiave, p.mese), 0))}</td>
      </tr>
    ) : null;
  const dataIt = (iso: string) =>
    /^\d{4}-\d{2}-\d{2}/.test(iso)
      ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}`
      : "";

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
        <div className="mb-1 flex flex-wrap items-center gap-3">
          <span className="text-sm font-semibold text-foreground">{t("fc.titolo")}</span>
          {salvando && <Loader2 className="h-4 w-4 animate-spin text-muted-foreground" />}
        </div>
        <p className="mb-3 text-xs text-muted-foreground">{t("fc.desc")}</p>

        {/* Barra filtri */}
        <div className="mb-3 flex flex-wrap items-end gap-3 text-[13px]">
          <div className="flex rounded-lg border border-border overflow-hidden">
            {(["mese", "settimana"] as const).map((m) => (
              <button
                key={m}
                type="button"
                onClick={() => setModo(m)}
                className={`px-3 py-1 ${modo === m ? "bg-primary text-primary-foreground" : "hover:bg-muted"}`}
              >
                {m === "mese" ? t("fc.perMese") : t("fc.perSettimana")}
              </button>
            ))}
          </div>
          <label className="flex items-center gap-1.5 text-xs text-muted-foreground">
            {t("fc.daData")}
            <input
              type="date"
              value={daData}
              onChange={(e) => setDaData(e.target.value)}
              className={inputCls}
            />
          </label>
          <label className="flex items-center gap-1.5 text-xs text-muted-foreground">
            {t("fc.finoA")}
            <input
              type="date"
              value={finoA}
              onChange={(e) => setFinoA(e.target.value)}
              className={inputCls}
            />
          </label>
          <button
            type="button"
            onClick={() => setAperti(tuttiAperti ? new Set() : new Set(GRUPPI))}
            title={t("fc.dettaglioTip")}
            className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
          >
            {tuttiAperti ? t("fc.nascondiDettaglio") : t("fc.mostraDettaglio")}
          </button>
          <button
            type="button"
            onClick={() => setCumulato((v) => !v)}
            title={t("fc.cumulatoTip")}
            className={`rounded-lg border px-3 py-1 ${cumulato ? "border-primary bg-primary text-primary-foreground" : "border-border hover:bg-muted"}`}
          >
            {t("fc.cumulato")}
          </button>
          <button
            type="button"
            onClick={() => setShowEscl((v) => !v)}
            className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
          >
            {t("fc.esclusioni")} ({esclusioni.length})
          </button>
          <button
            type="button"
            onClick={() => setShowFisc((v) => !v)}
            title={t("fc.fiscaleTip")}
            className={`rounded-lg border px-3 py-1 ${showFisc ? "border-primary" : "border-border"} hover:bg-muted`}
          >
            {t("fc.fiscaleBtn")}
          </button>
          <button
            type="button"
            onClick={() => setShowSim((v) => !v)}
            title={t("fc.simTip")}
            className={`inline-flex items-center gap-1 rounded-lg border px-3 py-1 ${simAttiva ? "border-amber-400 bg-amber-50 text-amber-900" : showSim ? "border-primary" : "border-border"} hover:bg-muted`}
          >
            <FlaskConical className="h-3.5 w-3.5" />
            {t("fc.simBtn")}
            {simAttiva ? " ●" : ""}
          </button>
          <button
            type="button"
            onClick={esporta}
            className="rounded-lg bg-primary px-3 py-1 font-medium text-primary-foreground"
          >
            {t("common.exportCsv")}
          </button>
        </div>

        {/* Scadenziario fiscale */}
        {showFisc && (
          <div className="mb-4 rounded-xl border border-border p-3">
            <p className="mb-2 text-xs text-muted-foreground">{t("fc.fiscaleDesc")}</p>
            {fiscale && fiscale.scadenze.length > 0 && (
              <p className="mb-2 text-xs text-muted-foreground">
                {t("fc.fiscaleAgg")} <span className="font-medium">{fiscale.fonteFile}</span>
                {fiscale.aggiornatoIl
                  ? ` · ${new Date(fiscale.aggiornatoIl).toLocaleString("it-IT")}`
                  : ""}{" "}
                ·{" "}
                {
                  fiscale.scadenze.filter((s) => !s.pagato && s.categoria !== "finanziamento")
                    .length
                }{" "}
                {t("fc.fiscaleDaPagare")}
              </p>
            )}
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-lg border border-border px-3 py-1.5 text-[13px] hover:bg-muted">
              {parsingF && <Loader2 className="h-4 w-4 animate-spin" />}
              {t("fc.fiscaleScegli")}
              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  e.target.value = "";
                  if (f) void onFileFiscale(f);
                }}
              />
            </label>
            {previewFisc &&
              (() => {
                const nonPag = previewFisc.res.scadenze.filter((s) => !s.pagato);
                const fisc = nonPag.filter((s) => s.categoria !== "finanziamento");
                const fin = nonPag.filter((s) => s.categoria === "finanziamento");
                const totFisc = fisc.reduce((a, s) => a + s.importo, 0);
                const totFin = fin.reduce((a, s) => a + s.importo, 0);
                const totRate = fisc
                  .filter((s) => s.categoria === "rate")
                  .reduce((a, s) => a + s.importo, 0);
                const daRat = previewFisc.res.daRateizzare;
                return (
                  <div className="mt-2 space-y-1 rounded-lg border border-border/60 p-2 text-xs">
                    <p className="font-medium">
                      {previewFisc.fileName} — {t("fc.fiscaleFoglio")} “{previewFisc.res.foglio}”
                    </p>
                    <p>
                      {fisc.length} {t("fc.fiscaleDaPagare")} = {fmtImporto(totFisc)} € (
                      {t("fc.fiscaleRateLbl")} {fmtImporto(totRate)} € · {t("fc.fiscaleCorrLbl")}{" "}
                      {fmtImporto(totFisc - totRate)} €) · {t("fc.fiscaleUltima")}{" "}
                      {fisc.reduce((m, s) => (s.dataPagamento > m ? s.dataPagamento : m), "")}
                    </p>
                    {fin.length > 0 && (
                      <p className="text-muted-foreground">
                        {t("fc.fiscaleFin")}: {fin.length} = {fmtImporto(totFin)} €
                      </p>
                    )}
                    {previewFisc.res.senzaData > 0 && (
                      <p className="text-status-absent">
                        {previewFisc.res.senzaData} {t("fc.fiscaleSenzaData")}
                      </p>
                    )}
                    {daRat.length > 0 && (
                      <p className="text-status-absent">
                        {t("fc.fiscaleDaRat")}{" "}
                        {daRat
                          .map(
                            (d) =>
                              `${d.voce} ${d.periodo ?? ""} ${d.anno ?? ""} ${fmtImporto(d.importo)} €`,
                          )
                          .join(" · ")}
                      </p>
                    )}
                    <button
                      type="button"
                      disabled={savingF}
                      onClick={() => void confermaFiscale()}
                      className="mt-1 rounded-lg bg-primary px-3 py-1 font-medium text-primary-foreground disabled:opacity-40"
                    >
                      {savingF ? t("common.loading") : t("fc.fiscaleConferma")}
                    </button>
                  </div>
                );
              })()}
          </div>
        )}

        {/* Esclusioni */}
        {showEscl && (
          <div className="mb-4 rounded-xl border border-border p-3">
            <p className="mb-2 text-xs text-muted-foreground">{t("fc.esclDesc")}</p>
            <div className="mb-2 flex flex-wrap gap-2">
              {esclusioni.map((e) => (
                <span
                  key={e.id}
                  className="inline-flex items-center gap-1.5 rounded-full bg-muted px-2.5 py-1 text-xs"
                >
                  {e.nome}
                  {(e.mese || e.meseFine) && (
                    <span className="text-muted-foreground">
                      {e.mese ?? "…"} → {e.meseFine ?? "…"}
                    </span>
                  )}
                  <button
                    type="button"
                    onClick={() => void rimuoviRiga(e.id)}
                    title={t("common.delete")}
                  >
                    <Trash2 className="h-3 w-3 text-muted-foreground hover:text-destructive" />
                  </button>
                </span>
              ))}
              {esclusioni.length === 0 && (
                <span className="text-xs text-muted-foreground">{t("fc.esclNessuna")}</span>
              )}
            </div>
            <div className="flex flex-wrap items-end gap-2 text-[13px]">
              <input
                value={exCerca}
                onChange={(e) => setExCerca(e.target.value)}
                placeholder={t("fc.esclCercaPh")}
                className={`${inputCls} w-64`}
              />
              <input
                type="month"
                value={exDa}
                onChange={(e) => setExDa(e.target.value)}
                title={t("fc.esclDa")}
                className={inputCls}
              />
              <input
                type="month"
                value={exA}
                onChange={(e) => setExA(e.target.value)}
                title={t("fc.esclA")}
                className={inputCls}
              />
              <button
                type="button"
                disabled={exBusy || exSel.size === 0}
                onClick={() => void escludiSelezionate()}
                className="rounded-lg bg-primary px-3 py-1 text-primary-foreground disabled:opacity-40"
              >
                {t("fc.esclAggiungi")}
                {exSel.size > 0 ? ` (${exSel.size})` : ""}
              </button>
            </div>
            <div className="mt-2 grid max-h-56 grid-cols-1 gap-x-4 gap-y-0.5 overflow-y-auto rounded-lg border border-border/60 p-2 sm:grid-cols-2 lg:grid-cols-3">
              {contropartiEscludibili.visibili.map((c) => (
                <label
                  key={c}
                  className="flex cursor-pointer items-center gap-2 rounded px-1 py-0.5 text-xs hover:bg-muted"
                >
                  <input
                    type="checkbox"
                    className="accent-primary"
                    checked={exSel.has(c)}
                    onChange={() =>
                      setExSel((s) => {
                        const ns = new Set(s);
                        if (ns.has(c)) ns.delete(c);
                        else ns.add(c);
                        return ns;
                      })
                    }
                  />
                  <span className="truncate" title={c}>
                    {c}
                  </span>
                </label>
              ))}
              {contropartiEscludibili.oltre > 0 && (
                <span className="text-xs italic text-muted-foreground">
                  +{contropartiEscludibili.oltre} {t("fc.esclAltre")}
                </span>
              )}
              {contropartiEscludibili.visibili.length === 0 && (
                <span className="text-xs text-muted-foreground">{t("fc.esclTutteFuori")}</span>
              )}
            </div>
            <div className="mt-3 border-t border-border/60 pt-2">
              <p className="mb-1 text-xs text-muted-foreground">{t("fc.presetDesc")}</p>
              <div className="flex flex-wrap items-center gap-2 text-[13px]">
                {presets.map(([nome, righe]) => (
                  <span
                    key={nome}
                    className="inline-flex items-center gap-1.5 rounded-full border border-border px-2.5 py-1 text-xs"
                  >
                    <button
                      type="button"
                      disabled={presetBusy}
                      onClick={() => void applicaPreset(nome, righe)}
                      title={t("fc.presetApplicaTip")}
                      className="font-medium hover:text-primary disabled:opacity-40"
                    >
                      {nome} ({righe.length})
                    </button>
                    <button
                      type="button"
                      disabled={presetBusy}
                      onClick={() => void eliminaPreset(righe)}
                      title={t("common.delete")}
                    >
                      <Trash2 className="h-3 w-3 text-muted-foreground hover:text-destructive" />
                    </button>
                  </span>
                ))}
                <input
                  value={presetNome}
                  onChange={(e) => setPresetNome(e.target.value)}
                  placeholder={t("fc.presetNomePh")}
                  className={`${inputCls} w-44`}
                />
                <button
                  type="button"
                  disabled={presetBusy || !presetNome.trim() || esclusioni.length === 0}
                  onClick={() => void salvaPreset()}
                  className="rounded-lg border border-border px-3 py-1 hover:bg-muted disabled:opacity-40"
                >
                  {t("fc.presetSalva")}
                </button>
              </div>
            </div>
            {/* GIRATE a fornitore (spec Simone 12/09) */}
            <div className="mt-3 border-t border-border/60 pt-2">
              <p className="mb-1 text-xs text-muted-foreground">{t("fc.girDesc")}</p>
              <div className="mb-2 flex flex-wrap gap-2">
                {girate.map((g) => (
                  <span
                    key={g.id}
                    className="inline-flex items-center gap-1.5 rounded-full bg-muted px-2.5 py-1 text-xs"
                  >
                    {g.cliente} → {Math.round(g.perc * 100)}% → {g.fornitore}
                    {g.oggettoTermini.length > 0 && (
                      <span className="text-muted-foreground">({g.oggettoTermini.join(", ")})</span>
                    )}
                    <button
                      type="button"
                      onClick={() => void rimuoviRiga(g.id)}
                      title={t("common.delete")}
                    >
                      <Trash2 className="h-3 w-3 text-muted-foreground hover:text-destructive" />
                    </button>
                  </span>
                ))}
                {girate.length === 0 && (
                  <span className="text-xs text-muted-foreground">{t("fc.girNessuna")}</span>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-2 text-[13px]">
                <span>{t("fc.girSe")}</span>
                <input
                  list="fc-girata-clienti"
                  value={giCliente}
                  onChange={(e) => setGiCliente(e.target.value)}
                  placeholder={t("fc.esclCercaPh")}
                  className={`${inputCls} w-48`}
                />
                <datalist id="fc-girata-clienti">
                  {contropartiNote.map((c) => (
                    <option key={c} value={c} />
                  ))}
                </datalist>
                <span>{t("fc.girPerc")}</span>
                <input
                  value={giPerc}
                  onChange={(e) => setGiPerc(e.target.value)}
                  className={`${inputCls} w-16 text-right`}
                />
                <span>{t("fc.girA")}</span>
                <select
                  value={giFornSel}
                  onChange={(e) => setGiFornSel(e.target.value)}
                  className={inputCls}
                >
                  <option value="DR Logistics">DR Logistics</option>
                  <option value="RN Servizi">RN Servizi</option>
                  <option value="altro">{t("fc.girAltro")}</option>
                </select>
                {giFornSel === "altro" && (
                  <>
                    <input
                      list="fc-girata-fornitori"
                      value={giFornAltro}
                      onChange={(e) => setGiFornAltro(e.target.value)}
                      placeholder={t("fc.esclCercaPh")}
                      className={`${inputCls} w-56`}
                    />
                    <datalist id="fc-girata-fornitori">
                      {fornitoriNote.map((c) => (
                        <option key={c} value={c} />
                      ))}
                    </datalist>
                  </>
                )}
                <input
                  value={giOggetto}
                  onChange={(e) => setGiOggetto(e.target.value)}
                  placeholder={t("fc.girOggettoPh")}
                  title={t("fc.girOggettoPh")}
                  className={`${inputCls} w-72`}
                />
                <button
                  type="button"
                  disabled={
                    giBusy || !giCliente.trim() || (giFornSel === "altro" && !giFornAltro.trim())
                  }
                  onClick={() => void aggiungiGirata()}
                  className="rounded-lg bg-primary px-3 py-1 text-primary-foreground disabled:opacity-40"
                >
                  {t("fc.girAggiungi")}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* SIMULAZIONE "se non pago" (Simone 29/09): tre leve combinabili */}
        {showSim && (
          <div className="mb-4 rounded-xl border border-amber-300/70 p-3">
            <p className="mb-2 text-xs text-muted-foreground">{t("fc.simDesc")}</p>
            <div className="flex flex-wrap items-center gap-2 text-[13px]">
              <button
                type="button"
                onClick={() => setSimPanel((p) => (p === "fatture" ? null : "fatture"))}
                className={`rounded-lg border px-3 py-1 hover:bg-muted ${simPanel === "fatture" ? "border-primary" : "border-border"} ${simFatture.size > 0 ? "bg-primary/10 font-medium" : ""}`}
              >
                {t("fc.simFatture")}
                {simFatture.size > 0 ? ` (${simFatture.size})` : ""}
              </button>
              <button
                type="button"
                onClick={() => setSimSenzaMod((v) => !v)}
                title={t("fc.simSenzaModTip")}
                className={`rounded-lg border px-3 py-1 ${simSenzaMod ? "border-primary bg-primary text-primary-foreground" : "border-border hover:bg-muted"}`}
              >
                {t("fc.simSenzaMod")}
              </button>
              <button
                type="button"
                onClick={() => setSimPanel((p) => (p === "fiscali" ? null : "fiscali"))}
                className={`rounded-lg border px-3 py-1 hover:bg-muted ${simPanel === "fiscali" ? "border-primary" : "border-border"} ${simFiscali.size > 0 ? "bg-primary/10 font-medium" : ""}`}
              >
                {t("fc.simFiscali")}
                {simFiscali.size > 0 ? ` (${simFiscali.size})` : ""}
              </button>
              {simAttiva && (
                <button
                  type="button"
                  onClick={azzeraSim}
                  className="ml-auto rounded-lg border border-border px-3 py-1 text-xs hover:bg-muted"
                >
                  {t("fc.simAzzera")}
                </button>
              )}
            </div>
            {simPanel === "fatture" && (
              <div className="mt-2">
                <div className="mb-1 flex flex-wrap items-center gap-2 text-[13px]">
                  <input
                    value={simCerca}
                    onChange={(e) => setSimCerca(e.target.value)}
                    placeholder={t("fc.simCercaPh")}
                    className={`${inputCls} w-64`}
                  />
                  <button
                    type="button"
                    onClick={() =>
                      setSimFatture((s) => {
                        const ns = new Set(s);
                        for (const r of visibiliSim.righe) ns.add(r.x.f.nomeFile);
                        return ns;
                      })
                    }
                    className="rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted"
                  >
                    {t("fc.simSelVisibili")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setSimFatture(new Set())}
                    disabled={simFatture.size === 0}
                    className="rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted disabled:opacity-40"
                  >
                    {t("fc.simDeselTutte")}
                  </button>
                  <span className="text-xs text-muted-foreground">
                    {fattureSimulabili.length} {t("fc.simFattureNelFlusso")}
                  </span>
                </div>
                <div className="max-h-72 overflow-y-auto rounded-lg border border-border/60">
                  <table className="w-full text-xs">
                    <tbody>
                      {visibiliSim.righe.map((r) => {
                        const daToggle = simSenzaMod && r.mod === "altro";
                        const checked = daToggle || simFatture.has(r.x.f.nomeFile);
                        return (
                          <tr
                            key={r.x.f.nomeFile}
                            className={`border-b border-border/30 ${checked ? "bg-amber-50/70" : ""}`}
                          >
                            <td className="w-6 px-1.5 py-0.5">
                              <input
                                type="checkbox"
                                className="accent-primary"
                                checked={checked}
                                disabled={daToggle}
                                title={daToggle ? t("fc.simDaToggle") : undefined}
                                onChange={() => toggleIn(setSimFatture, r.x.f.nomeFile)}
                              />
                            </td>
                            <td className="whitespace-nowrap px-1 py-0.5">{r.x.f.numero}</td>
                            <td className="max-w-64 truncate px-1 py-0.5" title={r.x.f.cliente}>
                              {r.x.f.cliente}
                            </td>
                            <td
                              className={`whitespace-nowrap px-1 py-0.5 ${r.x.s.inRitardo ? "text-status-absent" : "text-muted-foreground"}`}
                            >
                              {dataIt(r.scad)}
                            </td>
                            <td className="whitespace-nowrap px-1 py-0.5 text-right tabular-nums">
                              {fmtImporto(r.residuo)}
                            </td>
                            <td className="px-1.5 py-0.5 text-center">
                              <span className="rounded-full bg-muted px-1.5 text-[10px]">
                                {badgeMod[r.mod]}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                      {visibiliSim.oltre > 0 && (
                        <tr>
                          <td
                            colSpan={6}
                            className="px-1 py-1 text-xs italic text-muted-foreground"
                          >
                            +{visibiliSim.oltre} {t("fc.esclAltre")}
                          </td>
                        </tr>
                      )}
                      {visibiliSim.righe.length === 0 && (
                        <tr>
                          <td colSpan={6} className="py-3 text-center text-muted-foreground">
                            {t("fc.simNessuna")}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            {simPanel === "fiscali" && (
              <div className="mt-2">
                <div className="mb-1 flex flex-wrap items-center gap-2 text-[13px]">
                  <button
                    type="button"
                    onClick={() => setSimFiscali(new Set(fiscaliSimulabili.map(chiaveFisc)))}
                    className="rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted"
                  >
                    {t("fc.simSelTutte")}
                  </button>
                  <button
                    type="button"
                    onClick={() => setSimFiscali(new Set())}
                    disabled={simFiscali.size === 0}
                    className="rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted disabled:opacity-40"
                  >
                    {t("fc.simDeselTutte")}
                  </button>
                  <span className="text-xs text-muted-foreground">
                    {fiscaliSimulabili.length} {t("fc.fiscaleDaPagare")}
                  </span>
                </div>
                <div className="max-h-72 overflow-y-auto rounded-lg border border-border/60">
                  <table className="w-full text-xs">
                    <tbody>
                      {fiscaliSimulabili.map((s) => {
                        const k = chiaveFisc(s);
                        const checked = simFiscali.has(k);
                        return (
                          <tr
                            key={k}
                            className={`border-b border-border/30 ${checked ? "bg-amber-50/70" : ""}`}
                          >
                            <td className="w-6 px-1.5 py-0.5">
                              <input
                                type="checkbox"
                                className="accent-primary"
                                checked={checked}
                                onChange={() => toggleIn(setSimFiscali, k)}
                              />
                            </td>
                            <td
                              className={`whitespace-nowrap px-1 py-0.5 ${s.dataPagamento < oggiISO ? "text-status-absent" : "text-muted-foreground"}`}
                            >
                              {dataIt(s.dataPagamento)}
                            </td>
                            <td className="px-1 py-0.5">{s.voce}</td>
                            <td className="max-w-52 truncate px-1 py-0.5 text-muted-foreground">
                              {s.voceOld ?? s.periodo ?? ""}
                            </td>
                            <td className="px-1.5 py-0.5 text-center">
                              <span className="rounded-full bg-muted px-1.5 text-[10px]">
                                {s.categoria === "rate"
                                  ? t("fc.fiscaleRateLbl")
                                  : t("fc.fiscaleCorrLbl")}
                              </span>
                            </td>
                            <td className="whitespace-nowrap px-1 py-0.5 text-right tabular-nums">
                              {fmtImporto(s.importo)}
                            </td>
                          </tr>
                        );
                      })}
                      {fiscaliSimulabili.length === 0 && (
                        <tr>
                          <td colSpan={6} className="py-3 text-center text-muted-foreground">
                            {t("fc.drillFiscVuoto")}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
        {simAttiva && (
          <p className="mb-3 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-900">
            <span className="font-semibold">{t("fc.simAttiva")}</span> {simEscluse.nFatture}{" "}
            {t("fc.simFattureN")} ({fmtImporto(simEscluse.importoFatture)} €)
            {simSenzaMod ? ` — ${t("fc.simSenzaModNota")}` : ""}
            {simEscluse.nFiscali > 0
              ? ` · ${simEscluse.nFiscali} ${t("fc.simFiscaliN")} (${fmtImporto(simEscluse.importoFiscali)} €)`
              : ""}
            . {t("fc.simUsciteScad")}: {fmt(-usciteReali.tot.scaduto)} → {fmt(-uscite.tot.scaduto)}{" "}
            · {t("fc.simSaldoScad")}: {fmt(saldoScadutoReale)} → {fmt(saldoScaduto)}.
          </p>
        )}

        {flussiErr && (
          <p className="mb-3 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-xs text-amber-900">
            {flussiErr}
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
                <tr className="text-left text-[11px] text-muted-foreground">
                  <th className="py-1 pr-3 min-w-44" />
                  <th className={thCls}>{t("fc.colScaduto")}</th>
                  {periodi.map((p) => (
                    <th key={p.chiave} className={thCls}>
                      {p.label}
                    </th>
                  ))}
                  <th className={thCls}>{t("fc.colTotale")}</th>
                </tr>
              </thead>
              <tbody>
                {/* ENTRATE */}
                <tr className="border-t border-border/60 font-medium">
                  <td className="py-1 pr-3">
                    {btnPiu("t1:entrate", entrate.righe.length)}
                    {t("fc.entrate")}
                  </td>
                  <td className={`${tdN} text-status-absent`}>{fmt(entrate.tot.scaduto)}</td>
                  {serie(entrate.tot.scaduto, (c) => entrate.tot.perPeriodo.get(c) ?? 0).map(
                    (v, i) => (
                      <td key={periodi[i].chiave} className={tdN}>
                        {fmt(v)}
                      </td>
                    ),
                  )}
                  <td className={tdN}>{fmt(entrate.tot.scaduto + entrate.tot.totale)}</td>
                </tr>
                {aperti.has("t1:entrate") &&
                  entrate.righe.map((r) => (
                    <tr key={`e:${r.nome}`} className="border-t border-border/30">
                      <td className="max-w-56 truncate py-0.5 pl-4 pr-3 text-muted-foreground">
                        {r.nome}
                      </td>
                      <td className={`${tdN} text-muted-foreground`}>{fmt(r.scaduto)}</td>
                      {serie(r.scaduto, (c) => r.perPeriodo.get(c) ?? 0).map((v, i) => (
                        <td key={periodi[i].chiave} className={`${tdN} text-muted-foreground`}>
                          {fmt(v)}
                        </td>
                      ))}
                      <td className={`${tdN} text-muted-foreground`}>
                        {fmt(r.scaduto + r.totale)}
                      </td>
                    </tr>
                  ))}

                {/* USCITE (col segno meno) */}
                <tr className="border-t border-border/60 font-medium">
                  <td className="py-1 pr-3">{t("fc.uscite")}</td>
                  <td className={`${tdN} text-status-absent`}>{fmt(-uscite.tot.scaduto)}</td>
                  {serie(uscite.tot.scaduto, (c) => uscite.tot.perPeriodo.get(c) ?? 0).map(
                    (v, i) => (
                      <td key={periodi[i].chiave} className={tdN}>
                        {fmt(-v)}
                      </td>
                    ),
                  )}
                  <td className={tdN}>{fmt(-(uscite.tot.scaduto + uscite.tot.totale))}</td>
                </tr>
                {righeUscite(uscite, "t1")}

                {/* GIRATE: una riga per fornitore, come le altre (spec 12/09) */}
                {girataQuote.map(
                  (q) =>
                    girataTotaleDi(q) > 0.005 && (
                      <tr key={`g:${q.fornitore}`} className="border-t border-border/40">
                        <td className="py-1 pr-3" title={t("fc.girataTip")}>
                          {q.fornitore}
                        </td>
                        <td className={`${tdN} text-status-absent`}>{fmt(-q.scaduto)}</td>
                        {serie(q.scaduto, (c) => q.perPeriodo.get(c) ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={tdN}>
                            {fmt(-v)}
                          </td>
                        ))}
                        <td className={tdN}>{fmt(-girataTotaleDi(q))}</td>
                      </tr>
                    ),
                )}

                {/* PREFATTURE (solo vista mensile) */}
                {haPref && (
                  <>
                    <tr className="border-t border-border/40 italic text-primary">
                      <td className="py-1 pr-3">{t("fc.prefAtt")}</td>
                      <td className={tdN}>—</td>
                      {serie(0, (_c, m) => prefPer.att.get(m) ?? 0).map((v, i) => (
                        <td key={periodi[i].chiave} className={tdN}>
                          {fmt(v)}
                        </td>
                      ))}
                      <td className={tdN}>
                        {fmt([...prefPer.att.values()].reduce((s, v) => s + v, 0))}
                      </td>
                    </tr>
                    <tr className="border-t border-border/40 italic text-primary">
                      <td className="py-1 pr-3">{t("fc.prefPas")}</td>
                      <td className={tdN}>—</td>
                      {serie(0, (_c, m) => prefPer.pas.get(m) ?? 0).map((v, i) => (
                        <td key={periodi[i].chiave} className={tdN}>
                          {fmt(-v)}
                        </td>
                      ))}
                      <td className={tdN}>
                        {fmt(-[...prefPer.pas.values()].reduce((s, v) => s + v, 0))}
                      </td>
                    </tr>
                  </>
                )}

                {/* VOCI MANUALI (mensili) */}
                {modo === "mese" &&
                  nomiVoci.map((nome) => (
                    <tr key={`v:${nome}`} className="border-t border-border/40">
                      <td className="py-1 pr-3">
                        {VOCI_BASE.includes(nome) ? (
                          <button
                            type="button"
                            onClick={() => apriDrill(nome)}
                            title={t("fc.drillTip")}
                            className="rounded px-1 text-left hover:bg-muted hover:text-primary"
                          >
                            {nome}
                          </button>
                        ) : (
                          nome
                        )}
                      </td>
                      <td
                        className={`${tdN} ${scadutoVoce(nome) > 0 ? "text-status-absent" : ""}`}
                        title={
                          scadutoVoce(nome) > 0
                            ? t(
                                nome.trim().toLowerCase() === "stipendi"
                                  ? "fc.stipScadTip"
                                  : "fc.fiscScadTip",
                              )
                            : undefined
                        }
                      >
                        {scadutoVoce(nome) > 0 ? fmt(-scadutoVoce(nome)) : "—"}
                      </td>
                      {cumulato
                        ? serie(
                            -scadutoVoce(nome),
                            (_c, m) => valoreVoce(nome, m)?.importo ?? 0,
                          ).map((v, i) => (
                            <td key={periodi[i].chiave} className={tdN}>
                              {fmt(v)}
                            </td>
                          ))
                        : periodi.map((p) => (
                            <td key={p.chiave} className="py-0.5 pr-3 text-right">
                              {cellaVoceUI(nome, p.mese)}
                            </td>
                          ))}
                      <td className={tdN}>
                        {fmt(
                          periodi.reduce(
                            (s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0),
                            -scadutoVoce(nome),
                          ),
                        )}
                      </td>
                    </tr>
                  ))}

                {/* SALDO */}
                <tr className="border-t-2 border-border font-semibold">
                  <td className="py-1.5 pr-3">{t("fc.saldo")}</td>
                  <td
                    className={`${tdN} ${saldoScaduto >= 0 ? "text-status-present" : "text-status-absent"}`}
                  >
                    {fmt(saldoScaduto)}
                  </td>
                  {serie(saldoScaduto, (c, m) => saldoDi(c, m)).map((v, i) => (
                    <td
                      key={periodi[i].chiave}
                      className={`${tdN} ${v >= 0 ? "text-status-present" : "text-status-absent"}`}
                    >
                      {fmt(v)}
                    </td>
                  ))}
                  <td className={tdN}>
                    {fmt(periodi.reduce((s, p) => s + saldoDi(p.chiave, p.mese), 0))}
                  </td>
                </tr>
                {rigaConfronto(saldoScadutoReale, (c, m) => saldoDi(c, m, true))}
              </tbody>
            </table>
          </div>
        )}

        {/* --- SOLO FATTURAZIONI (spec Simone 14/09 sera): COPIA della tabella
            sopra — stesso Scaduto di partenza, stesse esclusioni, girate,
            prefatture e le 4 voci — ma dai mesi in poi le fatture pesano per
            il fatturato PIENO alla scadenza, senza scalare i movimenti
            bancari (incassate/pagate comprese). --- */}
        {!loading && (
          <div className="mt-6 border-t-2 border-border pt-4">
            <div className="text-sm font-semibold text-foreground">{t("fc.fatTitolo")}</div>
            <p className="mb-2 mt-0.5 text-[11px] text-muted-foreground">{t("fc.fatNota")}</p>
            <div className="overflow-x-auto">
              <table className="w-full text-[13px]">
                <thead>
                  <tr className="text-left text-[11px] text-muted-foreground">
                    <th className="py-1 pr-3 min-w-44" />
                    <th className={thCls}>{t("fc.colScaduto")}</th>
                    {periodi.map((p) => (
                      <th key={p.chiave} className={thCls}>
                        {p.label}
                      </th>
                    ))}
                    <th className={thCls}>{t("fc.colTotale")}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="border-t border-border/60 font-medium">
                    <td className="py-1 pr-3">
                      {btnPiu("t2:entrate", entrateFat.righe.length)}
                      {t("fc.entrate")}
                    </td>
                    <td className={`${tdN} text-status-absent`}>{fmt(entrateFat.tot.scaduto)}</td>
                    {serie(
                      entrateFat.tot.scaduto,
                      (c) => entrateFat.tot.perPeriodo.get(c) ?? 0,
                    ).map((v, i) => (
                      <td key={periodi[i].chiave} className={tdN}>
                        {fmt(v)}
                      </td>
                    ))}
                    <td className={tdN}>{fmt(entrateFat.tot.scaduto + entrateFat.tot.totale)}</td>
                  </tr>
                  {aperti.has("t2:entrate") &&
                    entrateFat.righe.map((r) => (
                      <tr key={`fe:${r.nome}`} className="border-t border-border/30">
                        <td className="max-w-56 truncate py-0.5 pl-4 pr-3 text-muted-foreground">
                          {r.nome}
                        </td>
                        <td className={`${tdN} text-muted-foreground`}>{fmt(r.scaduto)}</td>
                        {serie(r.scaduto, (c) => r.perPeriodo.get(c) ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={`${tdN} text-muted-foreground`}>
                            {fmt(v)}
                          </td>
                        ))}
                        <td className={`${tdN} text-muted-foreground`}>
                          {fmt(r.scaduto + r.totale)}
                        </td>
                      </tr>
                    ))}
                  <tr className="border-t border-border/60 font-medium">
                    <td className="py-1 pr-3">{t("fc.uscite")}</td>
                    <td className={`${tdN} text-status-absent`}>{fmt(-usciteFat.tot.scaduto)}</td>
                    {serie(usciteFat.tot.scaduto, (c) => usciteFat.tot.perPeriodo.get(c) ?? 0).map(
                      (v, i) => (
                        <td key={periodi[i].chiave} className={tdN}>
                          {fmt(-v)}
                        </td>
                      ),
                    )}
                    <td className={tdN}>{fmt(-(usciteFat.tot.scaduto + usciteFat.tot.totale))}</td>
                  </tr>
                  {righeUscite(usciteFat, "t2")}

                  {/* GIRATE, PREFATTURE e VOCI: identiche alla tabella sopra
                      (qui in sola lettura — si modificano di sopra). */}
                  {girataQuote.map(
                    (q) =>
                      girataTotaleDi(q) > 0.005 && (
                        <tr key={`fg:${q.fornitore}`} className="border-t border-border/40">
                          <td className="py-1 pr-3" title={t("fc.girataTip")}>
                            {q.fornitore}
                          </td>
                          <td className={`${tdN} text-status-absent`}>{fmt(-q.scaduto)}</td>
                          {serie(q.scaduto, (c) => q.perPeriodo.get(c) ?? 0).map((v, i) => (
                            <td key={periodi[i].chiave} className={tdN}>
                              {fmt(-v)}
                            </td>
                          ))}
                          <td className={tdN}>{fmt(-girataTotaleDi(q))}</td>
                        </tr>
                      ),
                  )}
                  {haPref && (
                    <>
                      <tr className="border-t border-border/40 italic text-primary">
                        <td className="py-1 pr-3">{t("fc.prefAtt")}</td>
                        <td className={tdN}>—</td>
                        {serie(0, (_c, m) => prefPer.att.get(m) ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={tdN}>
                            {fmt(v)}
                          </td>
                        ))}
                        <td className={tdN}>
                          {fmt([...prefPer.att.values()].reduce((s, v) => s + v, 0))}
                        </td>
                      </tr>
                      <tr className="border-t border-border/40 italic text-primary">
                        <td className="py-1 pr-3">{t("fc.prefPas")}</td>
                        <td className={tdN}>—</td>
                        {serie(0, (_c, m) => prefPer.pas.get(m) ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={tdN}>
                            {fmt(-v)}
                          </td>
                        ))}
                        <td className={tdN}>
                          {fmt(-[...prefPer.pas.values()].reduce((s, v) => s + v, 0))}
                        </td>
                      </tr>
                    </>
                  )}
                  {modo === "mese" &&
                    nomiVoci.map((nome) => (
                      <tr key={`fv:${nome}`} className="border-t border-border/40">
                        <td className="py-1 pr-3">{nome}</td>
                        <td
                          className={`${tdN} ${scadutoVoce(nome, true) > 0 ? "text-status-absent" : ""}`}
                          title={
                            scadutoVoce(nome, true) > 0
                              ? t(
                                  nome.trim().toLowerCase() === "stipendi"
                                    ? "fc.stipScadTip"
                                    : "fc.fiscScadTip",
                                )
                              : undefined
                          }
                        >
                          {scadutoVoce(nome, true) > 0 ? fmt(-scadutoVoce(nome, true)) : "—"}
                        </td>
                        {serie(
                          cumulato ? -scadutoVoce(nome, true) : 0,
                          (_c, m) => valoreVoce(nome, m, true)?.importo ?? 0,
                        ).map((v, i) => (
                          <td key={periodi[i].chiave} className={tdN}>
                            {fmt(v)}
                          </td>
                        ))}
                        <td className={tdN}>
                          {fmt(
                            periodi.reduce(
                              (s, p) => s + (valoreVoce(nome, p.mese, true)?.importo ?? 0),
                              -scadutoVoce(nome, true),
                            ),
                          )}
                        </td>
                      </tr>
                    ))}

                  <tr className="border-t-2 border-border font-semibold">
                    <td className="py-1.5 pr-3">{t("fc.saldo")}</td>
                    <td
                      className={`${tdN} ${saldoFatScaduto >= 0 ? "text-status-present" : "text-status-absent"}`}
                    >
                      {fmt(saldoFatScaduto)}
                    </td>
                    {serie(saldoFatScaduto, (c, m) => saldoFatDi(c, m)).map((v, i) => (
                      <td
                        key={periodi[i].chiave}
                        className={`${tdN} ${v >= 0 ? "text-status-present" : "text-status-absent"}`}
                      >
                        {fmt(v)}
                      </td>
                    ))}
                    <td className={tdN}>
                      {fmt(periodi.reduce((s, p) => s + saldoFatDi(p.chiave, p.mese), 0))}
                    </td>
                  </tr>
                  {rigaConfronto(saldoFatScadutoReale, (c, m) => saldoFatDi(c, m, true))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {drill &&
          (() => {
            const chiave = drill.voce.trim().toLowerCase();
            const chiudi = () => setDrill(null);
            const selMese = (
              <select
                value={drill.mese}
                onChange={(e) => setDrill({ ...drill, mese: e.target.value })}
                className={inputCls}
              >
                {periodi.map((pp) => (
                  <option key={pp.chiave} value={pp.mese}>
                    {pp.mese}
                  </option>
                ))}
              </select>
            );
            const manuale = vocePer(drill.voce, drill.mese);
            let corpo: React.ReactNode = null;
            if (chiave === "altre spese") {
              corpo = autoAltreSpese ? (
                <>
                  <p className="mb-2 text-xs text-muted-foreground">
                    {t("fc.drillAsDesc")} {autoAltreSpese.mesi.join(" + ")}.
                  </p>
                  <table className="w-full text-[13px]">
                    <thead>
                      <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
                        <th className="py-1 pr-3">{t("fc.colVoce")}</th>
                        <th className="py-1 pr-3 text-right">{autoAltreSpese.mesi[0]}</th>
                        <th className="py-1 pr-3 text-right">{autoAltreSpese.mesi[1]}</th>
                        <th className="py-1 pr-3 text-right">{t("fc.drillMedia")}</th>
                        <th className="py-1 pr-3 text-right" title={t("fc.drillAsCorrente")}>
                          {oggiISO.slice(0, 7)} {t("fc.drillAsFinora")}
                        </th>
                        <th className="py-1 text-center">{t("fc.drillInclusa")}</th>
                      </tr>
                    </thead>
                    <tbody>
                      {autoAltreSpese.righe.map((r) => (
                        <tr
                          key={r.tip}
                          className={`border-b border-border/40 ${r.inclusa ? "" : "text-muted-foreground line-through"}`}
                        >
                          <td className="py-0.5 pr-3">{r.tip}</td>
                          <td className="py-0.5 pr-3 text-right tabular-nums">
                            {fmtImporto(r.m1)}
                          </td>
                          <td className="py-0.5 pr-3 text-right tabular-nums">
                            {fmtImporto(r.m2)}
                          </td>
                          <td className="py-0.5 pr-3 text-right tabular-nums">
                            {fmtImporto((r.m1 + r.m2) / 2)}
                          </td>
                          <td className="py-0.5 pr-3 text-right tabular-nums text-muted-foreground">
                            {fmtImporto(r.corrente)}
                          </td>
                          <td className="py-0.5 text-center">
                            <input
                              type="checkbox"
                              className="accent-primary"
                              checked={r.inclusa}
                              disabled={drillBusy}
                              onChange={() => void toggleAsVoce(r.tip, r.inclusa)}
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr className="border-t border-border font-semibold">
                        <td className="py-1 pr-3" colSpan={3}>
                          {t("fc.drillMediaRisultante")}
                        </td>
                        <td className="py-1 pr-3 text-right tabular-nums">
                          {fmtImporto(autoAltreSpese.media)}
                        </td>
                        <td className="py-1 pr-3 text-right tabular-nums">
                          {fmtImporto(autoAltreSpese.realeCorrente)}
                        </td>
                        <td />
                      </tr>
                    </tfoot>
                  </table>
                  <p className="mt-2 text-[11px] text-muted-foreground">{t("fc.drillAsNota")}</p>
                  {/* COSTI FISSI (v1.85.0): importo reale, non media */}
                  <div className="mt-3 border-t border-border/60 pt-2">
                    <p className="mb-1 text-xs font-medium">{t("fc.fissiTitolo")}</p>
                    <p className="mb-2 text-[11px] text-muted-foreground">{t("fc.fissiDesc")}</p>
                    <table className="w-full text-[13px]">
                      <thead>
                        <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
                          <th className="py-1 pr-3">{t("fc.colVoce")}</th>
                          <th className="py-1 pr-3 text-right">{t("fc.fissiImporto")}</th>
                          <th className="py-1 pr-3">{t("fc.fissiToken")}</th>
                          <th className="py-1 pr-3">{t("fc.fissiColCorrente")}</th>
                          <th className="py-1" />
                        </tr>
                      </thead>
                      <tbody>
                        {costiFissi.map((f) => {
                          const pag = autoAltreSpese.pagatiFissi.get(f.id);
                          const attivo = fissoAttivo(f, oggiISO.slice(0, 7));
                          return (
                            <tr key={f.id} className="border-b border-border/40">
                              <td className="py-0.5 pr-3">
                                {f.nome}
                                {(f.mese || f.meseFine) && (
                                  <span className="ml-1 text-[11px] text-muted-foreground">
                                    {f.mese ?? "…"} → {f.meseFine ?? "…"}
                                  </span>
                                )}
                              </td>
                              <td className="py-0.5 pr-3 text-right tabular-nums">
                                {fmtImporto(f.importo)}
                              </td>
                              <td className="py-0.5 pr-3 text-muted-foreground">{f.token}</td>
                              <td className="py-0.5 pr-3">
                                {!attivo ? (
                                  "—"
                                ) : pag ? (
                                  <span className="text-status-present">
                                    ✓ {fmtImporto(pag.importo)} ({dataIt(pag.data)})
                                  </span>
                                ) : (
                                  <span className="text-muted-foreground">
                                    {t("fc.fissiDaPagare")}
                                  </span>
                                )}
                              </td>
                              <td className="py-0.5 text-right">
                                <button
                                  type="button"
                                  onClick={() => void rimuoviRiga(f.id)}
                                  title={t("common.delete")}
                                >
                                  <Trash2 className="h-3 w-3 text-muted-foreground hover:text-destructive" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                        {costiFissi.length === 0 && (
                          <tr>
                            <td colSpan={5} className="py-2 text-center text-muted-foreground">
                              {t("fc.fissiNessuno")}
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                    <div className="mt-2 flex flex-wrap items-end gap-2 text-[13px]">
                      <input
                        value={fxNome}
                        onChange={(e) => setFxNome(e.target.value)}
                        placeholder={t("fc.fissiNome")}
                        className={`${inputCls} w-44`}
                      />
                      <input
                        value={fxImporto}
                        onChange={(e) => setFxImporto(e.target.value)}
                        placeholder={t("fc.fissiImporto")}
                        className={`${inputCls} w-24 text-right`}
                      />
                      <input
                        value={fxToken}
                        onChange={(e) => setFxToken(e.target.value)}
                        placeholder={t("fc.fissiToken")}
                        title={t("fc.fissiTokenTip")}
                        className={`${inputCls} w-44`}
                      />
                      <input
                        type="month"
                        value={fxDa}
                        onChange={(e) => setFxDa(e.target.value)}
                        title={t("fc.esclDa")}
                        className={inputCls}
                      />
                      <input
                        type="month"
                        value={fxA}
                        onChange={(e) => setFxA(e.target.value)}
                        title={t("fc.esclA")}
                        className={inputCls}
                      />
                      <button
                        type="button"
                        disabled={
                          fxBusy ||
                          !fxNome.trim() ||
                          !(Number(fxImporto.trim().replace(/\./g, "").replace(",", ".")) > 0)
                        }
                        onClick={() => void aggiungiFisso()}
                        className="rounded-lg bg-primary px-3 py-1 text-primary-foreground disabled:opacity-40"
                      >
                        {t("fc.fissiAggiungi")}
                      </button>
                    </div>
                    {/* ANOMALIE (v1.86.0): fuori dalla media, dentro il reale */}
                    <div className="mt-3 border-t border-border/60 pt-2">
                      <p className="mb-1 text-xs font-medium">{t("fc.anomTitolo")}</p>
                      <p className="mb-2 text-[11px] text-muted-foreground">{t("fc.anomDesc")}</p>
                      {tabellaAnomalie(autoAltreSpese.movimentiTop)}
                    </div>
                    <p className="mt-2 text-[11px] text-muted-foreground">
                      {t("fc.drillAsCorrente")}: {fmtImporto(autoAltreSpese.realeCorrente)} +{" "}
                      {fmtImporto(autoAltreSpese.media)} ×{" "}
                      {autoAltreSpese.giorniMese - autoAltreSpese.giornoOggi}/
                      {autoAltreSpese.giorniMese} = {fmtImporto(autoAltreSpese.proiezione)} ·{" "}
                      {t("fc.fissiTot")} {fmtImporto(autoAltreSpese.fissiMese(oggiISO.slice(0, 7)))}{" "}
                      · {t("fc.colTotale")} {fmtImporto(autoAltreSpese.totale(oggiISO.slice(0, 7)))}
                    </p>
                  </div>
                </>
              ) : (
                <p className="text-xs text-muted-foreground">{t("common.loading")}</p>
              );
            } else if (chiave === TIP_CONSULENZE) {
              corpo = !autoConsulenze ? (
                <p className="text-xs text-muted-foreground">{t("common.loading")}</p>
              ) : (
                <>
                  <p className="mb-2 text-xs text-muted-foreground">{t("fc.drillConsDesc")}</p>
                  <p className="mb-2 text-xs">
                    {autoConsulenze.mesi[0]}: {fmtImporto(autoConsulenze.m1)} ·{" "}
                    {autoConsulenze.mesi[1]}: {fmtImporto(autoConsulenze.m2)} · {t("fc.drillMedia")}
                    : <span className="font-semibold">{fmtImporto(autoConsulenze.media)}</span> ·{" "}
                    {oggiISO.slice(0, 7)} {t("fc.drillAsFinora")}:{" "}
                    {fmtImporto(autoConsulenze.realeCorrente)} ({t("fc.anomDiCui")}{" "}
                    {fmtImporto(autoConsulenze.anomalieCorrente)}) → {t("fc.colTotale")}{" "}
                    <span className="font-semibold">{fmtImporto(autoConsulenze.proiezione)}</span> ·{" "}
                    {t("fc.anomSoglia")}: {fmtImporto(autoConsulenze.soglia)}
                  </p>
                  <div className="max-h-80 overflow-y-auto">
                    {tabellaAnomalie(autoConsulenze.movimenti)}
                  </div>
                  <p className="mt-2 text-[11px] text-muted-foreground">{t("fc.anomNota")}</p>
                </>
              );
            } else if (chiave === "stipendi") {
              const [anno, mm] = drill.mese.split("-").map(Number);
              const comp =
                mm === 1 ? `${anno - 1}-12` : `${anno}-${String(mm - 1).padStart(2, "0")}`;
              const netti = drillStip?.netti?.find((x) => x.mese === comp);
              const pagatiSet = new Set(drillStip?.pagatiPerMese?.[comp] ?? []);
              corpo = !drillStip ? (
                <p className="text-xs text-muted-foreground">{t("common.loading")}</p>
              ) : !netti ? (
                <p className="text-xs text-muted-foreground">
                  {t("fc.drillStipVuoto")} ({comp})
                </p>
              ) : (
                <>
                  <p className="mb-2 text-xs text-muted-foreground">
                    {t("fc.drillStipDesc")} {comp}.
                  </p>
                  <div className="max-h-80 overflow-y-auto">
                    <table className="w-full text-[13px]">
                      <thead>
                        <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
                          <th className="py-1 pr-3">{t("common.employee")}</th>
                          <th className="py-1 pr-3 text-right">{t("fc.drillSaldo")}</th>
                          <th className="py-1 text-center">{t("stip.colPagato")}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...netti.dipendenti]
                          .sort((a, b) => b.saldo - a.saldo)
                          .map((d, i) => (
                            <tr key={`${d.nome}|${i}`} className="border-b border-border/40">
                              <td className="py-0.5 pr-3">{d.nome}</td>
                              <td className="py-0.5 pr-3 text-right tabular-nums">
                                {fmtImporto(d.saldo)}
                              </td>
                              <td className="py-0.5 text-center">
                                {pagatiSet.has(chiaveNome(d.nome)) ? "✓" : "—"}
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                  <p className="mt-2 text-xs font-medium">
                    {t("fc.drillStipTot")} {fmtImporto(netti.totaleSaldo)} ·{" "}
                    {t("fc.drillStipPagati")}{" "}
                    {fmtImporto(
                      netti.dipendenti
                        .filter((d) => pagatiSet.has(chiaveNome(d.nome)))
                        .reduce((s2, d) => s2 + d.saldo, 0),
                    )}
                  </p>
                  <p className="mt-1 text-[11px] text-muted-foreground">{t("fc.drillStipNota")}</p>
                </>
              );
            } else {
              const cat = chiave === "costo fiscale rate" ? "rate" : "corrente";
              const corrente = drill.mese === oggiISO.slice(0, 7);
              // Mese corrente = reale: pagate del mese + da pagare da oggi in
              // poi; mesi futuri = non pagate in scadenza; le scadute non
              // pagate stanno nella colonna Scaduto (elenco a parte).
              const righeF = scadenzeFiscali
                .filter(
                  (x) =>
                    x.categoria === cat &&
                    x.dataPagamento.slice(0, 7) === drill.mese &&
                    (corrente ? x.pagato || x.dataPagamento.slice(0, 10) >= oggiISO : !x.pagato),
                )
                .sort((a, b) => a.dataPagamento.localeCompare(b.dataPagamento));
              const scaduteF = corrente
                ? scadenzeFiscali
                    .filter(
                      (x) =>
                        x.categoria === cat && !x.pagato && x.dataPagamento.slice(0, 10) < oggiISO,
                    )
                    .sort((a, b) => a.dataPagamento.localeCompare(b.dataPagamento))
                : [];
              const tabellaF = (righe: ScadenzaFiscale[], conStato: boolean) => (
                <table className="w-full text-[13px]">
                  <thead>
                    <tr className="border-b border-border text-left text-xs uppercase text-muted-foreground">
                      <th className="py-1 pr-3">{t("fis.colData")}</th>
                      <th className="py-1 pr-3">{t("fis.colVoce")}</th>
                      <th className="py-1 pr-3">{t("fis.colDettaglio")}</th>
                      {conStato && <th className="py-1 pr-3" />}
                      <th className="py-1 text-right">{t("fis.colImporto")}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {righe.map((x, i) => (
                      <tr key={x.id ?? i} className="border-b border-border/40">
                        <td className="py-0.5 pr-3 whitespace-nowrap">
                          {x.dataPagamento.slice(8)}/{x.dataPagamento.slice(5, 7)}
                        </td>
                        <td className="py-0.5 pr-3">{x.voce}</td>
                        <td className="max-w-52 truncate py-0.5 pr-3 text-muted-foreground">
                          {x.voceOld ?? x.periodo ?? ""}
                        </td>
                        {conStato && (
                          <td className="py-0.5 pr-3 whitespace-nowrap text-[11px]">
                            {x.pagato ? (
                              <span className="text-status-present">
                                ✓ {t("fc.drillFiscPagata")}
                              </span>
                            ) : (
                              <span className="text-muted-foreground">
                                {t("fc.drillFiscDaPagare")}
                              </span>
                            )}
                          </td>
                        )}
                        <td className="py-0.5 text-right tabular-nums">{fmtImporto(x.importo)}</td>
                      </tr>
                    ))}
                    {righe.length === 0 && (
                      <tr>
                        <td
                          colSpan={conStato ? 5 : 4}
                          className="py-3 text-center text-muted-foreground"
                        >
                          {t("fc.drillFiscVuoto")}
                        </td>
                      </tr>
                    )}
                  </tbody>
                  {righe.length > 0 && (
                    <tfoot>
                      <tr className="border-t border-border font-semibold">
                        <td colSpan={conStato ? 4 : 3} className="py-1 pr-3">
                          {t("fc.colTotale")}
                        </td>
                        <td className="py-1 text-right tabular-nums">
                          {fmtImporto(righe.reduce((s2, x) => s2 + x.importo, 0))}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              );
              corpo = (
                <>
                  <p className="mb-2 text-xs text-muted-foreground">{t("fc.drillFiscDesc")}</p>
                  {tabellaF(righeF, corrente)}
                  {corrente && scaduteF.length > 0 && (
                    <div className="mt-3 border-t border-border/60 pt-2">
                      <p className="mb-1 text-xs font-medium text-status-absent">
                        {t("fc.drillFiscScadute")}
                      </p>
                      {tabellaF(scaduteF, false)}
                    </div>
                  )}
                </>
              );
            }
            return (
              <div
                className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
                onClick={chiudi}
              >
                <div
                  className="max-h-[85vh] w-full max-w-2xl overflow-y-auto rounded-2xl border border-border bg-card p-5 shadow-xl"
                  onClick={(e) => e.stopPropagation()}
                >
                  <div className="mb-3 flex items-center gap-3">
                    <span className="text-sm font-semibold">{drill.voce}</span>
                    {chiave !== "altre spese" && chiave !== TIP_CONSULENZE && selMese}
                    {manuale && (
                      <span className="rounded-full bg-status-absent/15 px-2 py-0.5 text-[11px] text-status-absent">
                        {t("fc.drillManuale")} {fmtImporto(manuale.importo)}
                      </span>
                    )}
                    <button
                      type="button"
                      onClick={chiudi}
                      className="ml-auto rounded-lg border border-border px-2.5 py-1 text-xs hover:bg-muted"
                    >
                      {t("common.close")}
                    </button>
                  </div>
                  {corpo}
                </div>
              </div>
            );
          })()}
        <div className="mt-3 space-y-1 text-[11px] text-muted-foreground">
          <p>{t("fc.notaSegni")}</p>
          {modo === "settimana" && <p>{t("fc.notaSettimana")}</p>}
          {modo === "mese" && autoAltreSpese && autoAltreSpese.media > 0 && (
            <p>
              {t("fc.notaAuto")} {autoAltreSpese.mesi.join(" + ")} ={" "}
              {fmtImporto(autoAltreSpese.media)} €
              {costiFissi.length > 0
                ? ` + ${t("fc.fissiTot")} (${costiFissi.length}) ${fmtImporto(autoAltreSpese.fissiMese(periodi[1]?.mese ?? oggiISO.slice(0, 7)))} €`
                : ""}{" "}
              · {t("fc.drillAsCorrente")}.
            </p>
          )}
          {modo === "mese" && autoStipendi && autoStipendi.media > 0 && (
            <p>
              {t("fc.notaAutoStipendi")} {autoStipendi.mesi.join(" + ")} ={" "}
              {fmtImporto(autoStipendi.media)} €
            </p>
          )}
          {modo === "mese" &&
            stipendiMesi &&
            Object.entries(stipendiMesi).some(
              ([m, x]) => x.flaggati > 0 && m >= oggiISO.slice(0, 7),
            ) && <p>{t("fc.notaPagati")}</p>}
          {modo === "mese" && totFiscali && fiscale && (
            <p>
              {t("fc.notaFiscale")} {fiscale.fonteFile}
              {fiscale.daRateizzare.length > 0
                ? ` — ${t("fc.fiscaleDaRat")} ${fiscale.daRateizzare
                    .map(
                      (d) =>
                        `${d.voce} ${d.periodo ?? ""} ${d.anno ?? ""} ${fmtImporto(d.importo)} €`,
                    )
                    .join(" · ")}`
                : ""}
            </p>
          )}
          {modo === "mese" && (
            <div className="flex items-center gap-2">
              <span>{t("fc.nuovaVoce")}</span>
              <input
                value={nuovaVoce}
                onChange={(e) => setNuovaVoce(e.target.value)}
                placeholder={t("fc.nuovaVocePh")}
                className={`${inputCls} w-56`}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && nuovaVoce.trim()) {
                    const nome = nuovaVoce.trim();
                    setNuovaVoce("");
                    // La voce compare come riga: il primo importo la salva.
                    if (!nomiVoci.some((n) => n.toLowerCase() === nome.toLowerCase()))
                      setFlussi((prev) => [
                        ...(prev ?? []),
                        {
                          id: `tmp:${nome}`,
                          nome,
                          genere: "voce",
                          importo: 0,
                          mese: undefined,
                        } as FlussoCassaRiga,
                      ]);
                  }
                }}
              />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
