// DR Portal — Finanza → tab Flussi di cassa (direttore, call 07/09/2026).
// "La priorità assoluta dell'azienda è avere davvero in mano le spese":
// entrate e uscite ATTESE per mese (o settimana ISO), dalle scadenze delle
// fatture aperte — stessa semantica del Resoconto (residuoAperto) — più le
// righe che le fatture non conoscono: stipendi, costo fiscale, altre spese
// (voci manuali su SharePoint, lista FlussiCassa) e le prefatture. In fondo
// il DELTA SALDO. Le uscite viaggiano col segno meno: la griglia si incolla
// in Excel e si somma da sola. Esclusioni per controparte (anche a finestra
// di mesi) per tenere fuori chi non paga e le casse esterne.
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Loader2, Trash2 } from "lucide-react";
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
import {
  clienteGroupKey,
  matchRegola,
  regoleOrdinate,
  type RegolaFinanza,
} from "@/lib/finanza-logic";
import { esportaCsvFile } from "@/lib/csv";
import {
  spGetFatture,
  spGetTerminiPagamento,
  spGetPrefatture,
  spGetFlussiCassa,
  spUpsertFlussoCassa,
  spDeleteFlussoCassa,
  spGetRegoleFinanza,
  spGetMovimenti,
} from "@/lib/sharepoint.functions";
import { spStipendiStima } from "@/lib/stipendi.functions";
import {
  parseScadenzario,
  totaliFiscaliPerMese,
  type FiscaleDb,
  type ParseScadenzarioResult,
} from "@/lib/fiscale-logic";
import { spFiscaleGet, spFiscaleSalva } from "@/lib/fiscale.functions";
import type { SpFattura, SpMovimento, Prefattura, FlussoCassaRiga } from "@/lib/sharepoint.server";

function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Le 4 voci manuali nominate dal direttore: righe sempre visibili, anche
// vuote, così Sabrina/Lucrezia sanno dove scrivere.
const VOCI_BASE = ["Stipendi", "Costo fiscale rate", "Costo fiscale corrente", "Altre spese"];

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
  // Per la MEDIA automatica delle "Altre spese": regole flaggate + movimenti.
  const [regoleFin, setRegoleFin] = useState<RegolaFinanza[] | null>(null);
  const [movimenti, setMovimenti] = useState<SpMovimento[] | null>(null);

  const [modo, setModo] = useState<"mese" | "settimana">("mese");
  const [finoA, setFinoA] = useState("");
  const [daData, setDaData] = useState("");
  const [dettaglio, setDettaglio] = useState(true);
  // CUMULATO (richiesta Simone 14/09): ogni colonna mostra il progressivo
  // "fino a quel momento" — a settembre scaduto+settembre, a ottobre
  // scaduto+settembre+ottobre, e così via. Vale per entrambe le tabelle.
  const [cumulato, setCumulato] = useState(false);
  // Stima Stipendi per i mesi senza dato (media netto dovuto ultimi 2 mesi
  // da "Stipendi Dr.xlsx" — richiesta Simone 14/09).
  const [autoStipendi, setAutoStipendi] = useState<{ media: number; mesi: string[] } | null>(null);
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
    spStipendiStima()
      .then((s) => setAutoStipendi(s))
      .catch(() => setAutoStipendi(null));
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
    spGetRegoleFinanza()
      .then((l) => setRegoleFin(l as RegolaFinanza[]))
      .catch(() => setRegoleFin([]));
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
  const esclusa = (nomeControparte: string, meseScadenza: string): boolean => {
    const chiave = clienteGroupKey(nomeControparte) || nomeControparte.toLowerCase();
    return esclusioni.some((e) => {
      const token = clienteGroupKey(e.nome) || e.nome.trim().toLowerCase();
      if (!token || !chiave.includes(token)) return false;
      if (e.mese && meseScadenza < e.mese) return false;
      if (e.meseFine && meseScadenza > e.meseFine) return false;
      return true;
    });
  };

  // --- Somme per controparte -------------------------------------------------
  type RigaCp = {
    nome: string;
    scaduto: number;
    perPeriodo: Map<string, number>;
    totale: number;
  };
  const chiaviPeriodo = useMemo(() => new Set(periodi.map((p) => p.chiave)), [periodi]);
  const somma = (righe: typeof attive): { righe: RigaCp[]; tot: RigaCp } => {
    const per = new Map<string, RigaCp>();
    const tot: RigaCp = { nome: "", scaduto: 0, perPeriodo: new Map(), totale: 0 };
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
      const k = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const r = per.get(k) ?? { nome: x.f.cliente, scaduto: 0, perPeriodo: new Map(), totale: 0 };
      if (x.s.inRitardo) {
        r.scaduto += residuo;
        tot.scaduto += residuo;
      } else {
        const kp = chiaveDi(scad);
        if (!chiaviPeriodo.has(kp)) continue;
        r.perPeriodo.set(kp, (r.perPeriodo.get(kp) ?? 0) + residuo);
        r.totale += residuo;
        tot.perPeriodo.set(kp, (tot.perPeriodo.get(kp) ?? 0) + residuo);
        tot.totale += residuo;
      }
      per.set(k, r);
    }
    return {
      righe: [...per.values()]
        .filter((r) => r.totale > 0 || r.scaduto > 0)
        .sort((a, b) => b.totale + b.scaduto - (a.totale + a.scaduto)),
      tot,
    };
  };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const entrate = useMemo(() => somma(attive), [attive, periodi, esclusioni, daData, finoA]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const uscite = useMemo(() => somma(passive), [passive, periodi, esclusioni, daData, finoA]);

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
  ): { righe: RigaCp[]; tot: RigaCp } => {
    const nc = collegaNoteCredito(tutte, fattureEscluse(tutte));
    const per = new Map<string, RigaCp>();
    const tot: RigaCp = { nome: "", scaduto: 0, perPeriodo: new Map(), totale: 0 };
    for (const x of righe) {
      if (isNotaCredito(x.f.tipoDocumento)) continue;
      if (!x.s.scadenza) continue;
      const scad = x.s.scadenza.slice(0, 10);
      if (esclusa(x.f.cliente, scad.slice(0, 7))) continue;
      if (/^\d{4}-\d{2}-\d{2}$/.test(finoA) && scad > finoA) continue;
      const k = clienteGroupKey(x.f.cliente) || x.f.cliente;
      const r = per.get(k) ?? {
        nome: x.f.cliente,
        scaduto: 0,
        perPeriodo: new Map(),
        totale: 0,
      };
      if (x.s.inRitardo) {
        // IDENTICO alla tabella sopra: il punto di partenza è lo stato reale.
        const residuo = residuoAperto(x);
        if (residuo <= 1) continue;
        r.scaduto += residuo;
        tot.scaduto += residuo;
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
      }
      per.set(k, r);
    }
    return {
      righe: [...per.values()]
        .filter((r) => r.totale > 0.005 || r.scaduto > 0.005)
        .sort((a, b) => b.totale + b.scaduto - (a.totale + a.scaduto)),
      tot,
    };
  };
  const entrateFat = useMemo(
    () => sommaFatturato(attive, fattureEm ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [attive, fattureEm, periodi, esclusioni, daData, finoA],
  );
  const usciteFat = useMemo(
    () => sommaFatturato(passive, fattureRic ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [passive, fattureRic, periodi, esclusioni, daData, finoA],
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
  const saldoScaduto = entrate.tot.scaduto - uscite.tot.scaduto - girateScadutoTot;

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
  const autoAltreSpese = useMemo(() => {
    const flaggate = (regoleFin ?? []).filter((r) => r.altreSpese === true);
    if (!flaggate.length || !movimenti?.length) return null;
    const ordinate = regoleOrdinate(regoleFin ?? []);
    const meseDi = (iso: string) => iso.slice(0, 7);
    const base = new Date(`${oggiISO.slice(0, 7)}-01T00:00:00`);
    const mesi = [1, 2].map((i) => {
      const d = new Date(base.getFullYear(), base.getMonth() - i, 1);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    });
    const somme = new Map<string, number>(mesi.map((m) => [m, 0]));
    for (const m of movimenti) {
      if (m.importo >= 0) continue;
      const chiaveMese = meseDi(m.dataContabile);
      if (!somme.has(chiaveMese)) continue;
      // Stessa priorita' del classificatore: conta la PRIMA regola che
      // matcha, e conta solo se e' una di quelle flaggate.
      const r = ordinate.find((x) => matchRegola(m, x));
      if (r?.altreSpese === true)
        somme.set(chiaveMese, (somme.get(chiaveMese) ?? 0) + Math.abs(m.importo));
    }
    const media = [...somme.values()].reduce((s, v) => s + v, 0) / mesi.length;
    return { media: Math.round(media * 100) / 100, mesi };
  }, [regoleFin, movimenti, oggiISO]);

  // --- Voci manuali (mensili) ------------------------------------------------
  const nomiVoci = useMemo(() => {
    const set = new Set<string>(VOCI_BASE);
    for (const v of voci) set.add(v.nome);
    return [...set];
  }, [voci]);
  const vocePer = (nome: string, mese: string): FlussoCassaRiga | undefined =>
    voci.find((v) => v.nome.trim().toLowerCase() === nome.trim().toLowerCase() && v.mese === mese);

  // Totali dello scadenziario fiscale per mese (scadenze NON pagate; le già
  // scadute si spostano sul mese corrente perché sono ancora da pagare).
  const totFiscali = useMemo(
    () =>
      fiscale && fiscale.scadenze.length > 0
        ? totaliFiscaliPerMese(fiscale.scadenze, oggiISO.slice(0, 7))
        : null,
    [fiscale, oggiISO],
  );

  /** Valore effettivo di una voce nel mese: manuale se c'e', altrimenti — per
   *  la sola riga "Altre spese", dai mesi correnti in poi — la media
   *  automatica delle regole flaggate (in negativo: e' un'uscita). */
  const valoreVoce = (nome: string, mese: string): { importo: number; auto: boolean } | null => {
    const man = vocePer(nome, mese);
    if (man) return { importo: man.importo, auto: false };
    // Voci fiscali dallo scadenziario di Sabrina (importi ESATTI, non stime:
    // il "≈" segnala solo che arrivano in automatico dal file).
    const chiaveFisc = nome.trim().toLowerCase();
    if (totFiscali && mese >= oggiISO.slice(0, 7)) {
      const tot = totFiscali.get(mese);
      if (chiaveFisc === "costo fiscale rate" && tot && tot.rate > 0)
        return { importo: -tot.rate, auto: true };
      if (chiaveFisc === "costo fiscale corrente" && tot && tot.corrente > 0)
        return { importo: -tot.corrente, auto: true };
    }
    if (
      nome.trim().toLowerCase() === "altre spese" &&
      autoAltreSpese &&
      autoAltreSpese.media > 0 &&
      mese >= oggiISO.slice(0, 7)
    )
      return { importo: -autoAltreSpese.media, auto: true };
    // Stipendi futuri senza dato reale: stima = media del netto dovuto degli
    // ultimi 2 mesi di "Stipendi Dr.xlsx" (il dato vero, quando arriva
    // dall'import, vince perché è una voce manuale).
    if (
      nome.trim().toLowerCase() === "stipendi" &&
      autoStipendi &&
      autoStipendi.media > 0 &&
      mese >= oggiISO.slice(0, 7)
    )
      return { importo: -autoStipendi.media, auto: true };
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
  const saldoDi = (chiave: string, mese: string): number => {
    let v =
      (entrate.tot.perPeriodo.get(chiave) ?? 0) -
      (uscite.tot.perPeriodo.get(chiave) ?? 0) -
      girataQuote.reduce((s, q) => s + (q.perPeriodo.get(chiave) ?? 0), 0);
    if (modo === "mese") {
      v += (prefPer.att.get(mese) ?? 0) - (prefPer.pas.get(mese) ?? 0);
      for (const nome of nomiVoci) v += valoreVoce(nome, mese)?.importo ?? 0;
    }
    return Math.round(v * 100) / 100;
  };

  // Saldo della tabella "solo fatturazioni":
  // STESSA formula del saldo sopra (girate, prefatture e voci comprese),
  // cambiano solo entrate/uscite (a fatturato pieno). Lo Scaduto coincide
  // con quello sopra per costruzione.
  const saldoFatDi = (chiave: string, mese: string): number => {
    let v =
      (entrateFat.tot.perPeriodo.get(chiave) ?? 0) -
      (usciteFat.tot.perPeriodo.get(chiave) ?? 0) -
      girataQuote.reduce((s, q) => s + (q.perPeriodo.get(chiave) ?? 0), 0);
    if (modo === "mese") {
      v += (prefPer.att.get(mese) ?? 0) - (prefPer.pas.get(mese) ?? 0);
      for (const nome of nomiVoci) v += valoreVoce(nome, mese)?.importo ?? 0;
    }
    return Math.round(v * 100) / 100;
  };
  const saldoFatScaduto = entrateFat.tot.scaduto - usciteFat.tot.scaduto - girateScadutoTot;

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
    righe.push([
      t("fc.uscite"),
      num(-uscite.tot.scaduto),
      ...periodi.map((p) => num(-(uscite.tot.perPeriodo.get(p.chiave) ?? 0))),
      num(-(uscite.tot.scaduto + uscite.tot.totale)),
    ]);
    for (const r of uscite.righe)
      righe.push([
        `  ${r.nome}`,
        num(-r.scaduto),
        ...periodi.map((p) => num(-(r.perPeriodo.get(p.chiave) ?? 0))),
        num(-(r.scaduto + r.totale)),
      ]);
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
          "",
          ...periodi.map((p) => num(valoreVoce(nome, p.mese)?.importo ?? 0)),
          num(periodi.reduce((s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0), 0)),
        ]);
    righe.push([
      t("fc.saldo"),
      num(saldoScaduto),
      ...periodi.map((p) => num(saldoDi(p.chiave, p.mese))),
      num(periodi.reduce((s, p) => s + saldoDi(p.chiave, p.mese), 0)),
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
    righe.push([
      t("fc.uscite"),
      num(-usciteFat.tot.scaduto),
      ...periodi.map((p) => num(-(usciteFat.tot.perPeriodo.get(p.chiave) ?? 0))),
      num(-(usciteFat.tot.scaduto + usciteFat.tot.totale)),
    ]);
    for (const r of usciteFat.righe)
      righe.push([
        `  ${r.nome}`,
        num(-r.scaduto),
        ...periodi.map((p) => num(-(r.perPeriodo.get(p.chiave) ?? 0))),
        num(-(r.scaduto + r.totale)),
      ]);
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
          "",
          ...periodi.map((p) => num(valoreVoce(nome, p.mese)?.importo ?? 0)),
          num(periodi.reduce((s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0), 0)),
        ]);
    righe.push([
      t("fc.saldo"),
      num(saldoFatScaduto),
      ...periodi.map((p) => num(saldoFatDi(p.chiave, p.mese))),
      num(periodi.reduce((s, p) => s + saldoFatDi(p.chiave, p.mese), 0)),
    ]);
    esportaCsvFile(`flussi-di-cassa-${modo}`, testata, righe);
  };

  const loading = fattureEm == null || fattureRic == null || flussi == null;

  // Fornitori per la tendina/autocomplete delle girate (dalle passive).
  const fornitoriNote = useMemo(() => {
    const set = new Set<string>();
    for (const x of passive) set.add(x.f.cliente);
    return [...set].sort((a, b) => a.localeCompare(b)).slice(0, 400);
  }, [passive]);

  // Controparti per la checklist del form esclusioni.
  const contropartiNote = useMemo(() => {
    const set = new Set<string>();
    for (const x of [...attive, ...passive]) set.add(x.f.cliente);
    return [...set].sort((a, b) => a.localeCompare(b)).slice(0, 400);
  }, [attive, passive]);

  // Checklist: fuori le controparti GIA' coperte da un'esclusione (con
  // qualunque finestra di mesi — i chip sopra restano il posto per gestirle)
  // e quelle che non passano il filtro di ricerca.
  const contropartiEscludibili = useMemo(() => {
    const cerca = exCerca.trim().toLowerCase();
    return contropartiNote.filter((c) => {
      const chiave = clienteGroupKey(c) || c.toLowerCase();
      const giaEsclusa = esclusioni.some((e) => {
        const token = clienteGroupKey(e.nome) || e.nome.trim().toLowerCase();
        return !!token && chiave.includes(token);
      });
      if (giaEsclusa) return false;
      return !cerca || c.toLowerCase().includes(cerca);
    });
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
            onClick={() => setDettaglio((v) => !v)}
            className="rounded-lg border border-border px-3 py-1 hover:bg-muted"
          >
            {dettaglio ? t("fc.nascondiDettaglio") : t("fc.mostraDettaglio")}
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
              {contropartiEscludibili.map((c) => (
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
              {contropartiEscludibili.length === 0 && (
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
                  <td className="py-1 pr-3">{t("fc.entrate")}</td>
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
                {dettaglio &&
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
                {dettaglio &&
                  uscite.righe.map((r) => (
                    <tr key={`u:${r.nome}`} className="border-t border-border/30">
                      <td className="max-w-56 truncate py-0.5 pl-4 pr-3 text-muted-foreground">
                        {r.nome}
                      </td>
                      <td className={`${tdN} text-muted-foreground`}>{fmt(-r.scaduto)}</td>
                      {serie(r.scaduto, (c) => r.perPeriodo.get(c) ?? 0).map((v, i) => (
                        <td key={periodi[i].chiave} className={`${tdN} text-muted-foreground`}>
                          {fmt(-v)}
                        </td>
                      ))}
                      <td className={`${tdN} text-muted-foreground`}>
                        {fmt(-(r.scaduto + r.totale))}
                      </td>
                    </tr>
                  ))}

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
                      <td className="py-1 pr-3">{nome}</td>
                      <td className={tdN}>—</td>
                      {cumulato
                        ? serie(0, (_c, m) => valoreVoce(nome, m)?.importo ?? 0).map((v, i) => (
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
                          periodi.reduce((s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0), 0),
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
                    <td className="py-1 pr-3">{t("fc.entrate")}</td>
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
                  {dettaglio &&
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
                  {dettaglio &&
                    usciteFat.righe.map((r) => (
                      <tr key={`fu:${r.nome}`} className="border-t border-border/30">
                        <td className="max-w-56 truncate py-0.5 pl-4 pr-3 text-muted-foreground">
                          {r.nome}
                        </td>
                        <td className={`${tdN} text-muted-foreground`}>{fmt(-r.scaduto)}</td>
                        {serie(r.scaduto, (c) => r.perPeriodo.get(c) ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={`${tdN} text-muted-foreground`}>
                            {fmt(-v)}
                          </td>
                        ))}
                        <td className={`${tdN} text-muted-foreground`}>
                          {fmt(-(r.scaduto + r.totale))}
                        </td>
                      </tr>
                    ))}

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
                        <td className={tdN}>—</td>
                        {serie(0, (_c, m) => valoreVoce(nome, m)?.importo ?? 0).map((v, i) => (
                          <td key={periodi[i].chiave} className={tdN}>
                            {fmt(v)}
                          </td>
                        ))}
                        <td className={tdN}>
                          {fmt(
                            periodi.reduce(
                              (s, p) => s + (valoreVoce(nome, p.mese)?.importo ?? 0),
                              0,
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
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="mt-3 space-y-1 text-[11px] text-muted-foreground">
          <p>{t("fc.notaSegni")}</p>
          {modo === "settimana" && <p>{t("fc.notaSettimana")}</p>}
          {modo === "mese" && autoAltreSpese && autoAltreSpese.media > 0 && (
            <p>
              {t("fc.notaAuto")} {autoAltreSpese.mesi.join(" + ")} ={" "}
              {fmtImporto(autoAltreSpese.media)} €
            </p>
          )}
          {modo === "mese" && autoStipendi && autoStipendi.media > 0 && (
            <p>
              {t("fc.notaAutoStipendi")} {autoStipendi.mesi.join(" + ")} ={" "}
              {fmtImporto(autoStipendi.media)} €
            </p>
          )}
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
