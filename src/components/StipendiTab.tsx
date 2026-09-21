// DR Portal — Finanza → tab Stipendi (direzione).
// Costo del personale per mese dal file paghe "COSTI <MESE> <ANNO>.xlsx"
// (OneDrive Personale\MENSILITA'): import del foglio costi, dettaglio per
// dipendente, totale mese e aggancio della riga "Stipendi" ai Flussi di
// cassa (voce del mese di PAGAMENTO, di default il mese successivo alla
// competenza).

import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { Banknote, Loader2, RefreshCw, Trash2, Upload } from "lucide-react";
import { useLang } from "@/lib/i18n";
import { esportaCsvFile } from "@/lib/csv";
import {
  applicaModifiche,
  chiaveModifica,
  chiaveMotivo,
  chiaveNome,
  MOTIVI_NETTO,
  mensilitaStimate,
  meseSuccessivo,
  parseCostiFile,
  parseMappatura,
  parseStipendiDr,
  totaleMese,
  type AnagraficaDipendente,
  type CampoModificabile,
  type MotivoNettoTipo,
  type NettiMese,
  type ParseCostiFileResult,
  type StipendiDb,
  type StipendioDipendente,
} from "@/lib/stipendi-logic";
import {
  spStipendiEliminaMese,
  spStipendiGet,
  spStipendiMensilita,
  spStipendiModifica,
  spStipendiMotivo,
  spStipendiPagati,
  spStipendiSalvaAnagrafica,
  spStipendiSalvaMese,
  spStipendiSalvaNetti,
  spStipendiSync,
} from "@/lib/stipendi.functions";
import {
  spGetDettagliDistinte,
  spGetRosterDipendenti,
  spUpsertFlussoCassa,
} from "@/lib/sharepoint.functions";
import { matchDipendenteNome, type DipendenteRoster } from "@/lib/finanza-logic";
import type { DettaglioDistinta } from "@/lib/sharepoint.server";

const inputCls =
  "rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function eur(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtMese(yyyymm: string): string {
  const [y, m] = yyyymm.split("-");
  const nomi = [
    "Gennaio",
    "Febbraio",
    "Marzo",
    "Aprile",
    "Maggio",
    "Giugno",
    "Luglio",
    "Agosto",
    "Settembre",
    "Ottobre",
    "Novembre",
    "Dicembre",
  ];
  return `${nomi[Number(m) - 1] ?? m} ${y}`;
}

function fmtDataIt(iso: string): string {
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}

// Chiave di confronto nomi tra fonti diverse (paghe, distinte, mappatura).
const nameKey = chiaveNome;

// "CERRO AL LAMBRO" → "Cerro Al Lambro" ("" e "None" → "").
function titoloSede(s: string): string {
  const t = s.trim();
  if (!t || /^none$/i.test(t)) return "";
  return t.toLowerCase().replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}

interface PreviewStip {
  fileName: string;
  res: ParseCostiFileResult;
}

export function StipendiTab() {
  const { t } = useLang();
  // dbRaw = il file com'è (con la lista delle modifiche a mano a parte);
  // db = quello che si mostra, con le correzioni applicate. Tutto il resto
  // del componente legge `db`, la "M" e la traccia leggono `dbRaw`.
  const [dbRaw, setDb] = useState<StipendiDb | null>(null);
  const db = useMemo(() => (dbRaw ? applicaModifiche(dbRaw) : null), [dbRaw]);
  const modMap = useMemo(
    () =>
      new Map(
        (dbRaw?.modifiche ?? []).map(
          (x) => [chiaveModifica(x.mese, x.chiave, x.campo), x] as const,
        ),
      ),
    [dbRaw],
  );
  // Lettura automatica dei file paghe da SharePoint (1.79.0).
  const [syncing, setSyncing] = useState(false);
  const aggiornaDaiFile = async () => {
    if (syncing) return;
    setSyncing(true);
    try {
      const { db: nuovo, esito } = await spStipendiSync();
      setDb(nuovo);
      const riep = `${t("stip.syncNetti")} ${esito.nettiMesi.length} · ${t("stip.syncCosti")} ${esito.costiMesi.length}${esito.flussiAggiornati ? ` · ${t("stip.syncFlussi")} ${esito.flussiAggiornati}` : ""}`;
      if (esito.errori.length)
        toast.warning(t("stip.syncParziale"), {
          description: `${riep} — ${esito.errori.join(" | ")}`,
        });
      else toast.success(t("stip.syncOk"), { description: riep });
    } catch (err) {
      toast.error(t("stip.syncErr"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setSyncing(false);
    }
  };
  // Editor inline dei valori (doppio clic sulla cella).
  const [modEdit, setModEdit] = useState<string | null>(null); // chiaveModifica
  const [modVal, setModVal] = useState("");
  const [modBusy, setModBusy] = useState(false);
  const [errore, setErrore] = useState<string | null>(null);
  const [meseSel, setMeseSel] = useState("");
  const [q, setQ] = useState("");
  const [etichettaF, setEtichettaF] = useState("");
  const [showImport, setShowImport] = useState(false);
  const [parsing, setParsing] = useState(false);
  const [preview, setPreview] = useState<PreviewStip | null>(null);
  const [aggiornaFlussi, setAggiornaFlussi] = useState(true);
  const [meseFlussi, setMeseFlussi] = useState("");
  // Mese di competenza del file in anteprima: proposto dal contenuto (o dal
  // nome file per il tracciato per-appalto) e sempre correggibile a mano.
  const [meseComp, setMeseComp] = useState("");
  const [saving, setSaving] = useState(false);
  const [confermaElimina, setConfermaElimina] = useState(false);
  // Distinte stipendi (report BPM "Esiti pagamenti"): per l'EFFETTIVO versato.
  const [distinte, setDistinte] = useState<DettaglioDistinta[] | null>(null);
  // Roster anagrafica (nome + appalto): sostituisce le etichette non-sede
  // dei file paghe (giugno = codice azienda "311", luglio = nome dipendente).
  const [roster, setRoster] = useState<DipendenteRoster[]>([]);
  // Import "Stipendi Dr.xlsx" (netti da bonificare, un foglio per mese).
  const [showNetti, setShowNetti] = useState(false);
  const [parsingN, setParsingN] = useState(false);
  const [previewNetti, setPreviewNetti] = useState<{
    fileName: string;
    mesi: NettiMese[];
  } | null>(null);
  const [flussiNetti, setFlussiNetti] = useState(true);
  const [savingN, setSavingN] = useState(false);
  // Import "Mappatura Dipendenti" (anagrafica contrattuale: livello, contratto).
  const [showMappa, setShowMappa] = useState(false);
  const [parsingM, setParsingM] = useState(false);
  const [previewMappa, setPreviewMappa] = useState<{
    fileName: string;
    anagrafica: AnagraficaDipendente[];
  } | null>(null);
  const [savingM, setSavingM] = useState(false);
  // Editor mensilità dichiarate (override della stima dai ratei paghe).
  const [mensEdit, setMensEdit] = useState<string | null>(null); // chiave nome
  const [mensVal, setMensVal] = useState("");
  const [mensBusy, setMensBusy] = useState(false);

  const refresh = () =>
    spStipendiGet()
      .then((res) => {
        setDb(res);
        setMeseSel((prev) =>
          prev && res.mesi.some((m) => m.mese === prev)
            ? prev
            : (res.mesi[res.mesi.length - 1]?.mese ?? ""),
        );
      })
      .catch((err) => {
        const msg = err instanceof Error ? err.message : String(err);
        setErrore(msg);
        toast.error(msg);
      });
  useEffect(() => {
    void refresh();
    spGetDettagliDistinte()
      .then((l) => setDistinte(l as DettaglioDistinta[]))
      .catch(() => setDistinte([]));
    spGetRosterDipendenti()
      .then((l) => setRoster(l as DipendenteRoster[]))
      .catch(() => setRoster([]));
  }, []);

  const mese = useMemo(() => db?.mesi.find((m) => m.mese === meseSel) ?? null, [db, meseSel]);
  const precedente = useMemo(() => {
    if (!db || !mese) return null;
    const i = db.mesi.findIndex((m) => m.mese === mese.mese);
    return i > 0 ? db.mesi[i - 1] : null;
  }, [db, mese]);

  // SEDE per riga (richiesta Simone 15/09): l'etichetta del file quando è una
  // sede vera; quando invece è vuota, solo cifre (il "311" del file di
  // giugno) o il nome del dipendente stesso (file di luglio col nome in prima
  // colonna), si ricava in ordine: (1) l'etichetta-sede della stessa persona
  // in un ALTRO mese del file paghe (dal più recente — stessa nomenclatura
  // della tabella), (2) l'appalto dell'anagrafica portale (stesso motore di
  // match dei nomi delle distinte, ambiguità = nessun match), (3) la voce
  // "Altri" — mai il nome della persona come finta sede.
  const sediMese = useMemo(() => {
    const nonSedeVal = (e: string, k: string) => !e || /^\d+$/.test(e) || chiaveNome(e) === k;
    const sedePersona = new Map<string, string>();
    for (const mm of [...(db?.mesi ?? [])].reverse()) {
      for (const d of mm.dipendenti) {
        const k = chiaveNome(`${d.cognome} ${d.nome}`);
        const e = (d.etichetta ?? "").trim();
        if (!sedePersona.has(k) && !nonSedeVal(e, k)) sedePersona.set(k, e);
      }
    }
    const nomiRoster = roster.map((r) => r.nome);
    const appaltoDi = new Map(roster.map((r) => [chiaveNome(r.nome), r.appalto ?? ""]));
    const m = new Map<string, string>();
    for (const d of mese?.dipendenti ?? []) {
      const e = (d.etichetta ?? "").trim();
      const k = chiaveNome(`${d.cognome} ${d.nome}`);
      let sede = e;
      if (nonSedeVal(e, k)) {
        const hit = nomiRoster.length
          ? matchDipendenteNome(`${d.cognome} ${d.nome}`, nomiRoster)
          : null;
        sede =
          sedePersona.get(k) ||
          (hit ? appaltoDi.get(chiaveNome(hit)) : "") ||
          // Indicazione HR 21/09: la "Descrizione ripartizione 1" del file
          // costi (es. SAVONA, CERRO AL LAMBRO) dice la sede quando il
          // file non porta l'appalto.
          titoloSede(d.ripartizione ?? "") ||
          t("stip.sedeAltri");
      }
      m.set(`${d.codice}|${d.cognome} ${d.nome}|${d.etichetta}`, sede);
    }
    return m;
  }, [db, mese, roster, t]);
  const sedeDi = (d: StipendioDipendente) =>
    sediMese.get(`${d.codice}|${d.cognome} ${d.nome}|${d.etichetta}`) ?? d.etichetta;
  const etichette = useMemo(
    () => [...new Set((mese?.dipendenti ?? []).map((d) => sedeDi(d)).filter(Boolean))].sort(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [mese, sediMese],
  );
  const righe = useMemo(() => {
    const ql = q.trim().toLowerCase();
    return (mese?.dipendenti ?? [])
      .filter((d) => !etichettaF || sedeDi(d) === etichettaF)
      .filter(
        (d) =>
          !ql || `${d.cognome} ${d.nome} ${d.etichetta} ${sedeDi(d)}`.toLowerCase().includes(ql),
      );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mese, q, etichettaF, sediMese]);
  const totRighe = useMemo(
    () => ({
      totale: righe.reduce((a, d) => a + d.totaleCosto, 0),
      ore: righe.reduce((a, d) => a + d.totaleOre, 0),
      ord: righe.reduce((a, d) => a + d.costoOrdinario, 0),
      str: righe.reduce((a, d) => a + d.costoStraordinario, 0),
      ferie: righe.reduce((a, d) => a + d.feriePermessi, 0),
      mens: righe.reduce((a, d) => a + d.mensilitaAggiuntive, 0),
      tfr: righe.reduce((a, d) => a + d.tfr, 0),
    }),
    [righe],
  );

  // EFFETTIVO VERSATO (richiesta direzione 14/09): il netto uscito dal conto
  // verso i dipendenti, dalle distinte stipendi BPM (report "Esiti pagamenti"
  // importato in Finanze → Storico estratti). Gli stipendi del mese M si
  // pagano nel mese M+1: per il mese selezionato si guardano le disposizioni
  // salari ESEGUITE nel mese successivo, abbinate al dipendente per nome.
  const salariPag = useMemo(() => {
    const mesePag = /^\d{4}-\d{2}$/.test(meseSel) ? meseSuccessivo(meseSel) : "";
    // Contano le distinte "Stipendi SEPA" e — richiesta Simone 15/09 — anche
    // i BONIFICI SINGOLI intestati a un dipendente del mese (alcuni stipendi
    // viaggiano fuori distinta); le Riba e i bonifici ai fornitori restano
    // fuori perché il beneficiario non è in organico.
    const chiaviDip = new Set(
      (mese?.dipendenti ?? []).map((x) => nameKey(`${x.cognome} ${x.nome}`)),
    );
    const righeS = (distinte ?? []).filter((d) => {
      if ((d.dataEsecuzione || "").slice(0, 7) !== mesePag) return false;
      if (/salar|stipend/i.test(d.tipoPagamento)) return true;
      return chiaviDip.has(nameKey(d.beneficiario));
    });
    const perNome = new Map<string, number>();
    for (const d of righeS) {
      const k = nameKey(d.beneficiario);
      perNome.set(k, (perNome.get(k) ?? 0) + d.importo);
    }
    return {
      mesePag,
      righe: righeS,
      perNome,
      totale: righeS.reduce((a, d) => a + d.importo, 0),
    };
  }, [distinte, meseSel, mese]);
  const versatoDi = (d: StipendioDipendente): number | null =>
    salariPag.perNome.get(nameKey(`${d.cognome} ${d.nome}`)) ?? null;
  const nonAbbinate = useMemo(() => {
    const chiavi = new Set((mese?.dipendenti ?? []).map((x) => nameKey(`${x.cognome} ${x.nome}`)));
    return salariPag.righe.filter((r) => !chiavi.has(nameKey(r.beneficiario)));
  }, [salariPag, mese]);
  // Netto da "Stipendi Dr.xlsx" per il mese selezionato (competenza).
  const nettoMese = useMemo(
    () => db?.netti?.find((n) => n.mese === meseSel) ?? null,
    [db, meseSel],
  );
  // Netti per dipendente (match per nome, ordine parole libero): colonne
  // Stipendio/Anticipo/Saldo della tabella (richiesta Simone 14/09).
  const nettiPerNome = useMemo(() => {
    const m = new Map<string, { stipendio: number; anticipo: number; saldo: number }>();
    for (const d of nettoMese?.dipendenti ?? []) {
      const k = nameKey(d.nome);
      const cur = m.get(k) ?? { stipendio: 0, anticipo: 0, saldo: 0 };
      cur.stipendio += d.stipendio;
      cur.anticipo += d.anticipo;
      cur.saldo += d.saldo;
      m.set(k, cur);
    }
    return m;
  }, [nettoMese]);
  // Abbinamento TOLLERANTE (richiesta Simone 21/09, caso Gaggianesi/agosto):
  // prima la chiave esatta; se manca, lo stesso motore delle distinte
  // (matchDipendenteNome: pezzi di nome in ordine libero, tolleranza ai
  // troncamenti, ambiguità = nessun match). Recupera "Anna Maria" vs
  // "Annamaria", i nomi con una parola in più, l'ordine invertito — NON i
  // refusi veri ("Giaggianesi"): quelli si correggono nel file e si ricarica.
  const nettiAbbinati = useMemo(() => {
    const m = new Map<string, { stipendio: number; anticipo: number; saldo: number } | null>();
    const nomiNetti = (nettoMese?.dipendenti ?? []).map((d) => d.nome);
    // Chiavi dei netti "consumati" da una riga costi: il resto sono i netti
    // senza costo (cessati con liquidazione, contanti…) da mostrare a parte.
    const usati = new Set<string>();
    for (const d of mese?.dipendenti ?? []) {
      const nome = `${d.cognome} ${d.nome}`;
      const k = nameKey(nome);
      let hitKey: string | null = nettiPerNome.has(k) ? k : null;
      if (!hitKey && nomiNetti.length) {
        const match = matchDipendenteNome(nome, nomiNetti);
        if (match) hitKey = nameKey(match);
      }
      if (hitKey) usati.add(hitKey);
      m.set(k, hitKey ? (nettiPerNome.get(hitKey) ?? null) : null);
    }
    return { m, usati };
  }, [mese, nettoMese, nettiPerNome]);
  const nettoDi = (d: StipendioDipendente) =>
    nettiAbbinati.m.get(nameKey(`${d.cognome} ${d.nome}`)) ?? null;
  // Netti del mese senza riga costi (Stipendi Dr sì, COSTI no).
  const nettiSenzaCosto = useMemo(
    () => (nettoMese?.dipendenti ?? []).filter((d) => !nettiAbbinati.usati.has(nameKey(d.nome))),
    [nettoMese, nettiAbbinati],
  );

  // MOTIVO del netto assente / costo assente (risposte HR 21/09): vive a
  // parte e sopravvive ai re-import; la riga senza netto E senza motivo è
  // "da chiarire" e resta evidenziata finché qualcuno non la spiega.
  const motiviMap = useMemo(
    () =>
      new Map((dbRaw?.motiviNetto ?? []).map((x) => [chiaveMotivo(x.mese, x.chiave), x] as const)),
    [dbRaw],
  );
  const [motivoBusy, setMotivoBusy] = useState<string | null>(null);
  const motivoDi = (chiave: string) => motiviMap.get(chiaveMotivo(meseSel, chiave)) ?? null;
  const salvaMotivo = async (chiave: string, nome: string, motivo: MotivoNettoTipo | null) => {
    setMotivoBusy(chiave);
    try {
      const res = await spStipendiMotivo({ data: { mese: meseSel, chiave, nome, motivo } });
      setDb(res);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setMotivoBusy(null);
    }
  };
  const etichettaMotivo = (m: MotivoNettoTipo) => t(`stip.motivo.${m}`);
  const selectMotivo = (chiave: string, nome: string) => {
    const cur = motivoDi(chiave);
    return (
      <select
        value={cur?.motivo ?? ""}
        disabled={motivoBusy === chiave}
        title={
          cur
            ? t("stip.motivoTrace").replace("{da}", cur.da).replace("{il}", fmtDataIt(cur.il))
            : t("stip.motivoTip")
        }
        onChange={(e) =>
          void salvaMotivo(chiave, nome, (e.target.value || null) as MotivoNettoTipo | null)
        }
        className={`max-w-36 rounded border px-1 py-0.5 text-[11px] ${cur ? "border-border bg-background text-foreground" : "border-status-absent/50 bg-status-absent/5 text-status-absent"}`}
      >
        <option value="">{t("stip.motivoDaChiarire")}</option>
        {MOTIVI_NETTO.map((m) => (
          <option key={m} value={m}>
            {etichettaMotivo(m)}
          </option>
        ))}
      </select>
    );
  };
  const senzaNetto = (d: StipendioDipendente) => !nettoDi(d) && d.totaleCosto > 0;
  const daChiarire = useMemo(
    () =>
      (mese?.dipendenti ?? []).filter(
        (d) =>
          !nettiAbbinati.m.get(nameKey(`${d.cognome} ${d.nome}`)) &&
          d.totaleCosto > 0 &&
          !motiviMap.has(chiaveMotivo(meseSel, nameKey(`${d.cognome} ${d.nome}`))),
      ).length +
      nettiSenzaCosto.filter((d) => !motiviMap.has(chiaveMotivo(meseSel, nameKey(d.nome)))).length,
    [mese, nettiAbbinati, motiviMap, meseSel, nettiSenzaCosto],
  );
  // Anagrafica contrattuale (Mappatura Dipendenti) per nome.
  const anagPerNome = useMemo(() => {
    const m = new Map<string, AnagraficaDipendente>();
    for (const a of db?.anagrafica ?? []) m.set(nameKey(a.nome), a);
    return m;
  }, [db]);
  const anagDi = (d: StipendioDipendente) =>
    anagPerNome.get(nameKey(`${d.cognome} ${d.nome}`)) ?? null;
  // Mensilità stimate (12/13/14) dai ratei Mens.Agg. di TUTTI i mesi caricati:
  // mediana di (rateo/ordinario); i mesi con ordinario basso si scartano.
  const mensPerNome = useMemo(() => {
    const ratios = new Map<string, number[]>();
    for (const m of db?.mesi ?? [])
      for (const d of m.dipendenti) {
        if (d.costoOrdinario < 500) continue;
        const k = nameKey(`${d.cognome} ${d.nome}`);
        if (!ratios.has(k)) ratios.set(k, []);
        ratios.get(k)!.push(d.mensilitaAggiuntive / d.costoOrdinario);
      }
    const out = new Map<string, number>();
    for (const [k, rs] of ratios) {
      const v = mensilitaStimate(rs);
      if (v != null) out.set(k, v);
    }
    return out;
  }, [db]);
  /** Mensilità: l'override dichiarato (Mappatura/che ci dice la direzione)
   *  vince sulla stima dai ratei paghe (che si mostra col "≈"). */
  const mensDi = (d: StipendioDipendente): { val: number; stima: boolean } | null => {
    const ov = anagDi(d)?.mensilita;
    if (ov != null) return { val: ov, stima: false };
    const s = mensPerNome.get(nameKey(`${d.cognome} ${d.nome}`));
    return s != null ? { val: s, stima: true } : null;
  };

  // PAGATO SI/NO manuale (richiesta Simone 15/09): spunta per dipendente sul
  // mese di competenza; guida la riga Stipendi dei Flussi. La spunta in
  // testata marca/smarca TUTTE le righe filtrate (es. un appalto intero).
  const [pagBusy, setPagBusy] = useState(false);
  const pagatiSet = useMemo(() => new Set(db?.pagatiPerMese?.[meseSel] ?? []), [db, meseSel]);
  const pagatoDi = (d: StipendioDipendente) => pagatiSet.has(nameKey(`${d.cognome} ${d.nome}`));
  const cambiaPagati = async (aggiungi: string[], togli: string[]) => {
    if (!aggiungi.length && !togli.length) return;
    setPagBusy(true);
    try {
      const res = await spStipendiPagati({ data: { mese: meseSel, aggiungi, togli } });
      setDb(res);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setPagBusy(false);
    }
  };
  const togglePagato = (d: StipendioDipendente) => {
    const k = nameKey(`${d.cognome} ${d.nome}`);
    void cambiaPagati(pagatoDi(d) ? [] : [k], pagatoDi(d) ? [k] : []);
  };
  const tuttiPagati = righe.length > 0 && righe.every(pagatoDi);
  const toggleTuttiPagati = () => {
    const chiavi = righe.map((d) => nameKey(`${d.cognome} ${d.nome}`));
    if (tuttiPagati) void cambiaPagati([], chiavi);
    else
      void cambiaPagati(
        chiavi.filter((k) => !pagatiSet.has(k)),
        [],
      );
  };

  // CORREZIONE A MANO di un valore (richiesta Simone 21/09): doppio clic sulla
  // cella, Invio salva, Esc annulla, vuoto = torna al valore del file. La
  // cella corretta porta la "M" con chi/quando/prima nel tooltip.
  const valoreCella = (d: StipendioDipendente, campo: CampoModificabile): number | null => {
    if ((["stipendio", "anticipo", "saldo"] as string[]).includes(campo)) {
      const n = nettoDi(d);
      return n ? (n as unknown as Record<string, number>)[campo] : null;
    }
    return (d as unknown as Record<string, number>)[campo] ?? null;
  };
  const chiaveCella = (d: StipendioDipendente, campo: CampoModificabile) =>
    chiaveModifica(meseSel, nameKey(`${d.cognome} ${d.nome}`), campo);
  const apriModifica = (d: StipendioDipendente, campo: CampoModificabile) => {
    if (modBusy) return;
    const v = valoreCella(d, campo);
    setModVal(v == null || v === 0 ? "" : String(v).replace(".", ","));
    setModEdit(chiaveCella(d, campo));
  };
  const salvaModifica = async (d: StipendioDipendente, campo: CampoModificabile) => {
    const grezzo = modVal.trim();
    setModEdit(null);
    let valore: number | null = null;
    if (grezzo !== "") {
      // Accetta "1.234,56", "1234,56" e "1234.56".
      const s = grezzo.includes(",") ? grezzo.replace(/\./g, "").replace(",", ".") : grezzo;
      valore = Number(s);
      if (!Number.isFinite(valore)) {
        toast.error(t("stip.modErr"));
        return;
      }
      valore = Math.round(valore * 100) / 100;
    }
    const key = chiaveCella(d, campo);
    const attuale = valoreCella(d, campo);
    // Niente scrittura se non cambia nulla (né valore né stato manuale).
    if (valore == null && !modMap.has(key)) return;
    if (valore != null && attuale != null && Math.abs(valore - attuale) < 0.005) return;
    setModBusy(true);
    try {
      const res = await spStipendiModifica({
        data: {
          mese: meseSel,
          chiave: nameKey(`${d.cognome} ${d.nome}`),
          nome: `${d.cognome} ${d.nome}`,
          campo,
          valore,
        },
      });
      setDb(res);
      toast.success(valore == null ? t("stip.modTolta") : t("stip.modOk"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setModBusy(false);
    }
  };
  const tracciaModifica = (key: string): string | undefined => {
    const m = modMap.get(key);
    if (!m) return undefined;
    return t("stip.modTrace")
      .replace("{da}", m.da)
      .replace("{il}", fmtDataIt(m.il))
      .replace("{prima}", m.prima == null ? "—" : eur(m.prima));
  };
  /** Cella numerica modificabile: valore + eventuale "M". Funzione di
   *  rendering (NON un componente): un componente definito qui dentro
   *  cambierebbe identità a ogni render e l'input perderebbe il focus. */
  const cellaMod = (
    d: StipendioDipendente,
    campo: CampoModificabile,
    testo: string,
    className?: string,
  ) => {
    const key = chiaveCella(d, campo);
    const manuale = modMap.has(key);
    if (modEdit === key) {
      return (
        <td className={`px-3 py-1.5 text-right tabular-nums ${className ?? ""}`}>
          <input
            autoFocus
            inputMode="decimal"
            value={modVal}
            onChange={(e) => setModVal(e.target.value)}
            onBlur={() => void salvaModifica(d, campo)}
            onKeyDown={(e) => {
              if (e.key === "Enter") void salvaModifica(d, campo);
              if (e.key === "Escape") setModEdit(null);
            }}
            className="w-24 rounded border border-primary bg-background px-1 py-0.5 text-right text-[12px]"
          />
        </td>
      );
    }
    return (
      <td
        className={`cursor-text px-3 py-1.5 text-right tabular-nums ${className ?? ""}`}
        title={tracciaModifica(key) ?? t("stip.modTip")}
        onDoubleClick={() => apriModifica(d, campo)}
      >
        {testo}
        {manuale && (
          <span
            className="ml-1 inline-flex h-4 min-w-4 items-center justify-center rounded bg-primary/15 px-1 align-middle text-[10px] font-bold text-primary"
            aria-label={t("stip.modBadgeTip")}
          >
            M
          </span>
        )}
      </td>
    );
  };
  const modificheDi = (d: StipendioDipendente): string[] =>
    (dbRaw?.modifiche ?? [])
      .filter((x) => x.mese === meseSel && x.chiave === nameKey(`${d.cognome} ${d.nome}`))
      .map((x) => x.campo);

  const salvaMensilita = async (d: StipendioDipendente) => {
    const nome = anagDi(d)?.nome ?? `${d.cognome} ${d.nome}`;
    const grezzo = mensVal.trim();
    setMensEdit(null);
    const mensilita = grezzo === "" ? null : Number(grezzo);
    if (mensilita != null && (!Number.isInteger(mensilita) || mensilita < 12 || mensilita > 15)) {
      toast.error(t("stip.mensErr"));
      return;
    }
    // Click + blur senza modifiche: niente scrittura (e niente version bump).
    if (mensilita === (anagDi(d)?.mensilita ?? null)) return;
    setMensBusy(true);
    try {
      const res = await spStipendiMensilita({ data: { nome, mensilita } });
      setDb(res);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setMensBusy(false);
    }
  };

  const onFile = async (f: File) => {
    setParsing(true);
    setPreview(null);
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
      const res = parseCostiFile(fogli, f.name);
      if (!res) {
        toast.error(t("stip.errTracciato"));
        return;
      }
      setPreview({ fileName: f.name, res });
      setMeseComp(res.mese);
      setMeseFlussi(res.mese ? meseSuccessivo(res.mese) : "");
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsing(false);
    }
  };

  const onFileNetti = async (f: File) => {
    setParsingN(true);
    setPreviewNetti(null);
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
      const mesi = parseStipendiDr(fogli, f.name);
      if (!mesi) {
        toast.error(t("stip.errNetti"));
        return;
      }
      setPreviewNetti({ fileName: f.name, mesi });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsingN(false);
    }
  };

  const confermaNetti = async () => {
    if (!previewNetti) return;
    setSavingN(true);
    try {
      const res = await spStipendiSalvaNetti({ data: { mesi: previewNetti.mesi } });
      let flussiFatti = 0;
      if (flussiNetti) {
        // La riga "Stipendi" del cash flow = SALDO netto da versare, sul mese
        // di PAGAMENTO (successivo alla competenza del foglio).
        for (const m of previewNetti.mesi) {
          if (Math.abs(m.totaleSaldo) < 0.005) continue;
          await spUpsertFlussoCassa({
            data: {
              nome: "Stipendi",
              genere: "voce",
              mese: meseSuccessivo(m.mese),
              importo: -m.totaleSaldo,
              note: `${t("stip.nettiNota")} ${fmtMese(m.mese)} (${previewNetti.fileName})`,
            },
          });
          flussiFatti++;
        }
      }
      setDb(res);
      setPreviewNetti(null);
      setShowNetti(false);
      toast.success(t("stip.nettiOk"), {
        description: `${previewNetti.mesi.length} ${t("stip.nettiMesi")}${flussiFatti ? ` · ${flussiFatti} ${t("stip.nettiFlussiOk")}` : ""}`,
      });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSavingN(false);
    }
  };

  const onFileMappa = async (f: File) => {
    setParsingM(true);
    setPreviewMappa(null);
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
      const anagrafica = parseMappatura(fogli);
      if (!anagrafica) {
        toast.error(t("stip.errMappa"));
        return;
      }
      setPreviewMappa({ fileName: f.name, anagrafica });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setParsingM(false);
    }
  };

  const confermaMappa = async () => {
    if (!previewMappa) return;
    setSavingM(true);
    try {
      const res = await spStipendiSalvaAnagrafica({
        data: { anagrafica: previewMappa.anagrafica, fonte: previewMappa.fileName },
      });
      setDb(res);
      setPreviewMappa(null);
      setShowMappa(false);
      toast.success(t("stip.mappaOk"), {
        description: `${previewMappa.anagrafica.length} ${t("stip.dipendenti").toLowerCase()}`,
      });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSavingM(false);
    }
  };

  const conferma = async () => {
    if (!preview || !/^\d{4}-\d{2}$/.test(meseComp)) return;
    setSaving(true);
    try {
      const res = await spStipendiSalvaMese({
        data: {
          mese: {
            mese: meseComp,
            fonteFile: preview.fileName,
            dipendenti: preview.res.dipendenti,
          },
        },
      });
      if (aggiornaFlussi && /^\d{4}-\d{2}$/.test(meseFlussi)) {
        await spUpsertFlussoCassa({
          data: {
            nome: "Stipendi",
            genere: "voce",
            mese: meseFlussi,
            importo: -preview.res.totale,
            note: `${t("stip.flussiNota")} ${fmtMese(meseComp)} (${preview.fileName})`,
          },
        });
      }
      setDb(res);
      setMeseSel(meseComp);
      setPreview(null);
      setShowImport(false);
      toast.success(t("stip.importOk"), {
        description: `${fmtMese(meseComp)} · ${preview.res.dipendenti.length} ${t("stip.dipendenti").toLowerCase()} · ${eur(preview.res.totale)} €${aggiornaFlussi && meseFlussi ? ` · ${t("stip.flussiOk")} ${fmtMese(meseFlussi)}` : ""}`,
      });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const elimina = async () => {
    if (!mese) return;
    if (!confermaElimina) {
      setConfermaElimina(true);
      setTimeout(() => setConfermaElimina(false), 4000);
      return;
    }
    setConfermaElimina(false);
    try {
      const res = await spStipendiEliminaMese({ data: { mese: mese.mese } });
      setDb(res);
      setMeseSel(res.mesi[res.mesi.length - 1]?.mese ?? "");
      toast.success(t("stip.eliminato"));
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    }
  };

  const esporta = () => {
    if (!mese) return;
    esportaCsvFile(
      `stipendi-${mese.mese}`,
      [
        "Mese",
        "Etichetta",
        "Codice",
        "Cognome",
        "Nome",
        "Livello",
        "Contratto",
        "Fine contratto",
        "Mensilita",
        "Ore ordinarie",
        "Ore straordinarie",
        "Costo ordinario",
        "Straordinari",
        "Ferie/Permessi",
        "Mensilita aggiuntive",
        "TFR",
        "Totale costo",
        "Stipendio netto",
        "Anticipo",
        "Saldo",
        "Versato in banca",
        "Pagato",
        "Costo medio",
        "Modifiche manuali",
        "Motivo netto assente",
      ],
      mese.dipendenti.map((d) => [
        mese.mese,
        sedeDi(d),
        d.codice,
        d.cognome,
        d.nome,
        anagDi(d)?.livello ?? "",
        anagDi(d)?.contratto ?? "",
        anagDi(d)?.fineContratto ?? "",
        mensDi(d)?.val ?? "",
        d.oreOrdinarie,
        d.oreStraordinarie,
        d.costoOrdinario,
        d.costoStraordinario,
        d.feriePermessi,
        d.mensilitaAggiuntive,
        d.tfr,
        d.totaleCosto,
        nettoDi(d)?.stipendio ?? "",
        nettoDi(d)?.anticipo || "",
        nettoDi(d)?.saldo ?? "",
        versatoDi(d) ?? "",
        pagatoDi(d) ? "SI" : "NO",
        d.costoMedio,
        modificheDi(d).join("; "),
        motivoDi(nameKey(`${d.cognome} ${d.nome}`))?.motivo ?? "",
      ]),
    );
  };

  if (!db && !errore) {
    return (
      <div className="flex items-center gap-2 py-16 text-muted-foreground">
        <Loader2 className="h-5 w-5 animate-spin" /> {t("common.loading")}
      </div>
    );
  }
  if (errore && !db) return <p className="text-sm text-destructive">{errore}</p>;

  const totale = mese ? totaleMese(mese) : 0;
  const totalePrec = precedente ? totaleMese(precedente) : null;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end gap-3">
        <div>
          <label className="text-xs text-muted-foreground">{t("stip.mese")}</label>
          <select
            className={`${inputCls} mt-1 block`}
            value={meseSel}
            onChange={(e) => setMeseSel(e.target.value)}
          >
            {(db?.mesi ?? []).map((m) => (
              <option key={m.mese} value={m.mese}>
                {fmtMese(m.mese)}
              </option>
            ))}
            {(db?.mesi ?? []).length === 0 && <option value="">—</option>}
          </select>
        </div>
        <input
          className={`${inputCls} w-56`}
          placeholder={t("stip.cerca")}
          value={q}
          onChange={(e) => setQ(e.target.value)}
        />
        <select
          className={`${inputCls} max-w-56`}
          value={etichettaF}
          onChange={(e) => setEtichettaF(e.target.value)}
        >
          <option value="">
            {t("stip.etichetta")}: {t("common.all").toLowerCase()}
          </option>
          {etichette.map((a) => (
            <option key={a} value={a}>
              {a}
            </option>
          ))}
        </select>
        <div className="ml-auto flex gap-2">
          <button
            type="button"
            onClick={() => void aggiornaDaiFile()}
            disabled={syncing}
            title={
              dbRaw?.ultimaSync
                ? t("stip.syncUltima")
                    .replace("{il}", new Date(dbRaw.ultimaSync.il).toLocaleString("it-IT"))
                    .replace("{da}", dbRaw.ultimaSync.da)
                : t("stip.syncTip")
            }
            className="inline-flex items-center gap-2 rounded-lg border border-primary/40 bg-primary/5 px-3 py-2 text-sm text-foreground hover:bg-primary/10 disabled:opacity-60"
          >
            {syncing ? (
              <Loader2 className="h-4 w-4 animate-spin" />
            ) : (
              <RefreshCw className="h-4 w-4" />
            )}{" "}
            {t("stip.syncBtn")}
          </button>
          <button
            type="button"
            onClick={() => setShowImport((x) => !x)}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("stip.importBtn")}
          </button>
          <button
            type="button"
            onClick={() => setShowNetti((x) => !x)}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("stip.nettiBtn")}
          </button>
          <button
            type="button"
            onClick={() => setShowMappa((x) => !x)}
            className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
          >
            <Upload className="h-4 w-4" /> {t("stip.mappaBtn")}
          </button>
          <button
            type="button"
            onClick={esporta}
            disabled={!mese}
            className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
          >
            {t("common.exportCsv")}
          </button>
          {mese && (
            <button
              type="button"
              onClick={elimina}
              className={`inline-flex items-center gap-1 rounded-lg border px-3 py-2 text-sm ${confermaElimina ? "border-destructive bg-destructive text-destructive-foreground" : "border-border text-muted-foreground hover:bg-muted"}`}
            >
              <Trash2 className="h-4 w-4" />
              {confermaElimina ? t("stip.confermaElimina") : t("common.delete")}
            </button>
          )}
        </div>
      </div>

      {showImport && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-1 text-sm font-semibold text-foreground">{t("stip.importTitle")}</div>
          <p className="mb-4 text-xs text-muted-foreground">{t("stip.importDesc")}</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            disabled={parsing || saving}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void onFile(f);
              e.target.value = "";
            }}
            className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
          />
          {parsing && (
            <p className="mt-3 inline-flex items-center gap-2 text-sm text-muted-foreground">
              <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
            </p>
          )}
          {preview && (
            <div className="mt-4 rounded-xl border border-border p-4">
              <div className="text-sm font-medium text-foreground">{preview.fileName}</div>
              <ul className="mt-2 space-y-1 text-[13px] text-muted-foreground">
                <li className="flex items-center gap-2">
                  {t("stip.mese")}:
                  <input
                    type="month"
                    className={inputCls}
                    value={meseComp}
                    onChange={(e) => {
                      setMeseComp(e.target.value);
                      if (/^\d{4}-\d{2}$/.test(e.target.value))
                        setMeseFlussi(meseSuccessivo(e.target.value));
                    }}
                  />
                  {!preview.res.mese && (
                    <span className="text-xs text-status-break">{t("stip.meseNonNelFile")}</span>
                  )}
                  {/^\d{4}-\d{2}$/.test(meseComp) && db?.mesi.some((m) => m.mese === meseComp) && (
                    <span className="rounded bg-status-break/20 px-1.5 py-0.5 text-xs text-status-break">
                      {t("stip.sovrascrive")}
                    </span>
                  )}
                </li>
                <li>
                  {t("stip.fogli")}: {preview.res.fogli.join(", ")}
                </li>
                <li>
                  {t("stip.dipendenti")}: <b>{preview.res.dipendenti.length}</b>
                  {preview.res.scartate > 0 &&
                    ` (${preview.res.scartate} ${t("fin.previewSkipped")})`}
                </li>
                <li>
                  {t("stip.totaleMese")}: <b>{eur(preview.res.totale)} €</b>
                </li>
                {preview.res.squadrature > 0 && (
                  <li className="text-status-break">
                    {preview.res.squadrature} {t("stip.squadrature")}
                  </li>
                )}
              </ul>
              <div className="mt-3 flex flex-wrap items-center gap-3 text-sm">
                <label className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    checked={aggiornaFlussi}
                    onChange={(e) => setAggiornaFlussi(e.target.checked)}
                  />
                  {t("stip.flussiCheck")}
                </label>
                <input
                  type="month"
                  className={inputCls}
                  value={meseFlussi}
                  onChange={(e) => setMeseFlussi(e.target.value)}
                  disabled={!aggiornaFlussi}
                />
              </div>
              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  onClick={() => void conferma()}
                  disabled={saving || !/^\d{4}-\d{2}$/.test(meseComp)}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {saving && <Loader2 className="h-4 w-4 animate-spin" />}
                  {t("stip.conferma")}
                </button>
                <button
                  type="button"
                  onClick={() => setPreview(null)}
                  disabled={saving}
                  className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                >
                  {t("common.cancel")}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {showNetti && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-1 text-sm font-semibold text-foreground">{t("stip.nettiTitle")}</div>
          <p className="mb-4 text-xs text-muted-foreground">{t("stip.nettiDesc")}</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            disabled={parsingN || savingN}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void onFileNetti(f);
              e.target.value = "";
            }}
            className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
          />
          {parsingN && (
            <p className="mt-3 inline-flex items-center gap-2 text-sm text-muted-foreground">
              <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
            </p>
          )}
          {previewNetti && (
            <div className="mt-4 rounded-xl border border-border p-4">
              <div className="text-sm font-medium text-foreground">{previewNetti.fileName}</div>
              <table className="mt-2 text-[13px]">
                <thead>
                  <tr className="text-left text-xs uppercase tracking-wider text-muted-foreground">
                    <th className="pr-4">{t("stip.mese")}</th>
                    <th className="pr-4 text-right">{t("stip.dipendenti")}</th>
                    <th className="pr-4 text-right">{t("stip.nettiStipendio")}</th>
                    <th className="pr-4 text-right">{t("stip.nettiAnticipi")}</th>
                    <th className="text-right">{t("stip.nettiSaldo")}</th>
                  </tr>
                </thead>
                <tbody>
                  {previewNetti.mesi.map((m) => (
                    <tr key={m.mese}>
                      <td className="pr-4">{fmtMese(m.mese)}</td>
                      <td className="pr-4 text-right tabular-nums">{m.dipendenti.length}</td>
                      <td className="pr-4 text-right tabular-nums">{eur(m.totaleStipendio)}</td>
                      <td className="pr-4 text-right tabular-nums">
                        {m.totaleAnticipi ? eur(m.totaleAnticipi) : "—"}
                      </td>
                      <td className="text-right font-semibold tabular-nums">
                        {eur(m.totaleSaldo)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <label className="mt-3 flex items-center gap-1.5 text-sm">
                <input
                  type="checkbox"
                  checked={flussiNetti}
                  onChange={(e) => setFlussiNetti(e.target.checked)}
                />
                {t("stip.nettiFlussiCheck")}
              </label>
              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  onClick={() => void confermaNetti()}
                  disabled={savingN}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {savingN && <Loader2 className="h-4 w-4 animate-spin" />}
                  {t("stip.conferma")}
                </button>
                <button
                  type="button"
                  onClick={() => setPreviewNetti(null)}
                  disabled={savingN}
                  className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                >
                  {t("common.cancel")}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {showMappa && (
        <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
          <div className="mb-1 text-sm font-semibold text-foreground">{t("stip.mappaTitle")}</div>
          <p className="mb-4 text-xs text-muted-foreground">{t("stip.mappaDesc")}</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            disabled={parsingM || savingM}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void onFileMappa(f);
              e.target.value = "";
            }}
            className="block text-sm text-muted-foreground file:mr-3 file:rounded-lg file:border-0 file:bg-primary file:px-3 file:py-2 file:text-sm file:font-medium file:text-primary-foreground hover:file:opacity-90"
          />
          {parsingM && (
            <p className="mt-3 inline-flex items-center gap-2 text-sm text-muted-foreground">
              <Loader2 className="h-4 w-4 animate-spin" /> {t("fin.parsing")}
            </p>
          )}
          {previewMappa && (
            <div className="mt-4 rounded-xl border border-border p-4">
              <div className="text-sm font-medium text-foreground">{previewMappa.fileName}</div>
              <ul className="mt-2 space-y-1 text-[13px] text-muted-foreground">
                <li>
                  {t("stip.dipendenti")}: <b>{previewMappa.anagrafica.length}</b>
                </li>
                <li>
                  {t("stip.colLivello")}:{" "}
                  <b>{previewMappa.anagrafica.filter((a) => a.livello).length}</b> ·{" "}
                  {t("stip.mappaIndet")}:{" "}
                  <b>
                    {previewMappa.anagrafica.filter((a) => a.contratto === "Indeterminato").length}
                  </b>
                </li>
              </ul>
              <div className="mt-3 flex gap-2">
                <button
                  type="button"
                  onClick={() => void confermaMappa()}
                  disabled={savingM}
                  className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                >
                  {savingM && <Loader2 className="h-4 w-4 animate-spin" />}
                  {t("stip.conferma")}
                </button>
                <button
                  type="button"
                  onClick={() => setPreviewMappa(null)}
                  disabled={savingM}
                  className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                >
                  {t("common.cancel")}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {mese ? (
        <>
          <div className="grid grid-cols-2 gap-3 lg:grid-cols-3 xl:grid-cols-6">
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.totaleMese")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {eur(totale)} €
              </p>
              {totalePrec != null && totalePrec > 0 && (
                <p
                  className={`text-xs ${totale > totalePrec ? "text-status-out" : "text-status-present"}`}
                >
                  {totale >= totalePrec ? "+" : ""}
                  {eur(totale - totalePrec)} € {t("stip.vsPrec")} {fmtMese(precedente!.mese)}
                </p>
              )}
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.dipendenti")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {mese.dipendenti.length}
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.oreTot")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {mese.dipendenti.reduce((a, d) => a + d.totaleOre, 0).toLocaleString("it-IT")}
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.mediaDip")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {eur(mese.dipendenti.length ? totale / mese.dipendenti.length : 0)} €
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.versato")} · {salariPag.mesePag ? fmtMese(salariPag.mesePag) : "—"}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {salariPag.righe.length ? `${eur(salariPag.totale)} €` : "—"}
              </p>
              <p className="text-xs text-muted-foreground">
                {salariPag.righe.length
                  ? `${salariPag.righe.length} ${t("stip.disposizioni")}${nonAbbinate.length ? ` · ${nonAbbinate.length} ${t("stip.nonAbbinate")}` : ""}`
                  : t("stip.versatoVuoto")}
              </p>
            </div>
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-xs uppercase tracking-wider text-muted-foreground">
                {t("stip.nettiCard")}
              </p>
              <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">
                {nettoMese ? `${eur(nettoMese.totaleSaldo)} €` : "—"}
              </p>
              <p className="text-xs text-muted-foreground">
                {nettoMese
                  ? `${t("stip.nettiStipendio")} ${eur(nettoMese.totaleStipendio)}${nettoMese.totaleAnticipi ? ` · ${t("stip.nettiAnticipi").toLowerCase()} ${eur(nettoMese.totaleAnticipi)}` : ""}`
                  : t("stip.nettiVuoto")}
              </p>
            </div>
          </div>
          <p className="text-[11px] text-muted-foreground">{t("stip.versatoNota")}</p>
          <p className="text-[11px] text-muted-foreground">{t("stip.modNota")}</p>
          {daChiarire > 0 ? (
            <p className="inline-flex items-center gap-2 rounded-lg border border-status-absent/40 bg-status-absent/5 px-3 py-1.5 text-[12px] text-foreground">
              <span className="font-semibold text-status-absent">{daChiarire}</span>{" "}
              {t("stip.daChiarire")}
            </p>
          ) : (
            <p className="text-[11px] text-muted-foreground">{t("stip.tuttoChiaro")}</p>
          )}

          <div className="overflow-x-auto rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-border text-left text-xs uppercase tracking-wider text-muted-foreground">
                  <th className="px-3 py-2">{t("stip.etichetta")}</th>
                  <th className="px-3 py-2">{t("common.employee")}</th>
                  <th className="px-3 py-2">{t("stip.colLivello")}</th>
                  <th className="px-3 py-2">{t("stip.colContratto")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.colMens")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.oreOrd")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.oreStr")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.ordinario")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.straord")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.ferie")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.mensAgg")}</th>
                  <th className="px-3 py-2 text-right">TFR</th>
                  <th className="px-3 py-2 text-right">{t("stip.totale")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.colStip")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.colAnt")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.colSaldo")}</th>
                  <th className="px-3 py-2 text-right">{t("stip.versatoCol")}</th>
                  <th className="px-3 py-2 text-center" title={t("stip.pagatoTuttiTip")}>
                    <span className="mr-1">{t("stip.colPagato")}</span>
                    <input
                      type="checkbox"
                      className="accent-primary align-middle"
                      checked={tuttiPagati}
                      disabled={pagBusy || righe.length === 0}
                      onChange={toggleTuttiPagati}
                    />
                  </th>
                  <th className="px-3 py-2 text-right">{t("stip.medio")}</th>
                </tr>
              </thead>
              <tbody>
                {righe.map((d) => (
                  <tr
                    key={`${d.codice}|${d.cognome} ${d.nome}|${d.etichetta}`}
                    className={`border-b border-border/60 hover:bg-muted/40 ${senzaNetto(d) && !motivoDi(nameKey(`${d.cognome} ${d.nome}`)) ? "bg-status-absent/5" : ""}`}
                  >
                    <td
                      className="max-w-44 truncate px-3 py-1.5 text-xs text-muted-foreground"
                      title={
                        sedeDi(d) !== d.etichetta
                          ? `${t("stip.etichettaFileTip")}: ${d.etichetta || "—"}`
                          : undefined
                      }
                    >
                      {sedeDi(d) || "—"}
                    </td>
                    <td className="px-3 py-1.5 font-medium">
                      {d.cognome} {d.nome}
                      <span className="ml-1 font-mono text-[10px] text-muted-foreground">
                        {d.codice}
                      </span>
                    </td>
                    <td className="px-3 py-1.5" title={anagDi(d)?.mansione}>
                      {anagDi(d)?.livello || "—"}
                    </td>
                    <td className="max-w-32 truncate px-3 py-1.5 text-xs" title={anagDi(d)?.orario}>
                      {anagDi(d)
                        ? anagDi(d)!.contratto === "Indeterminato"
                          ? t("stip.indet")
                          : anagDi(d)!.fineContratto
                            ? `${t("stip.det")} ${fmtDataIt(anagDi(d)!.fineContratto!)}`
                            : (anagDi(d)!.contratto ?? "—")
                        : "—"}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {mensEdit === nameKey(`${d.cognome} ${d.nome}`) ? (
                        <input
                          autoFocus
                          value={mensVal}
                          onChange={(e) => setMensVal(e.target.value)}
                          onBlur={() => void salvaMensilita(d)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") void salvaMensilita(d);
                            if (e.key === "Escape") setMensEdit(null);
                          }}
                          className="w-10 rounded border border-primary bg-background px-1 py-0.5 text-right text-[12px]"
                        />
                      ) : (
                        <button
                          type="button"
                          disabled={mensBusy}
                          title={t("stip.mensTip")}
                          onClick={() => {
                            setMensEdit(nameKey(`${d.cognome} ${d.nome}`));
                            const m = mensDi(d);
                            setMensVal(m && !m.stima ? String(m.val) : "");
                          }}
                          className="w-full rounded px-1 text-right hover:bg-muted"
                        >
                          {(() => {
                            const m = mensDi(d);
                            if (!m) return "—";
                            return m.stima ? (
                              <span className="italic text-muted-foreground">≈ {m.val}</span>
                            ) : (
                              m.val
                            );
                          })()}
                        </button>
                      )}
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{d.oreOrdinarie || "—"}</td>
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {d.oreStraordinarie || "—"}
                    </td>
                    {cellaMod(d, "costoOrdinario", eur(d.costoOrdinario))}
                    {cellaMod(
                      d,
                      "costoStraordinario",
                      d.costoStraordinario ? eur(d.costoStraordinario) : "—",
                    )}
                    {cellaMod(d, "feriePermessi", eur(d.feriePermessi))}
                    {cellaMod(d, "mensilitaAggiuntive", eur(d.mensilitaAggiuntive))}
                    {cellaMod(d, "tfr", eur(d.tfr))}
                    {cellaMod(d, "totaleCosto", eur(d.totaleCosto), "font-semibold")}
                    {senzaNetto(d) && !modMap.has(chiaveCella(d, "stipendio")) ? (
                      <td
                        className="px-3 py-1.5 text-right"
                        onDoubleClick={() => apriModifica(d, "stipendio")}
                      >
                        {modEdit === chiaveCella(d, "stipendio") ? (
                          <input
                            autoFocus
                            inputMode="decimal"
                            value={modVal}
                            onChange={(e) => setModVal(e.target.value)}
                            onBlur={() => void salvaModifica(d, "stipendio")}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") void salvaModifica(d, "stipendio");
                              if (e.key === "Escape") setModEdit(null);
                            }}
                            className="w-24 rounded border border-primary bg-background px-1 py-0.5 text-right text-[12px]"
                          />
                        ) : (
                          selectMotivo(nameKey(`${d.cognome} ${d.nome}`), `${d.cognome} ${d.nome}`)
                        )}
                      </td>
                    ) : (
                      cellaMod(d, "stipendio", nettoDi(d) ? eur(nettoDi(d)!.stipendio) : "—")
                    )}
                    {cellaMod(
                      d,
                      "anticipo",
                      nettoDi(d)?.anticipo ? eur(nettoDi(d)!.anticipo) : "—",
                    )}
                    {cellaMod(d, "saldo", nettoDi(d) ? eur(nettoDi(d)!.saldo) : "—")}
                    <td className="px-3 py-1.5 text-right tabular-nums">
                      {versatoDi(d) != null ? eur(versatoDi(d)!) : "—"}
                    </td>
                    <td className="px-3 py-1.5 text-center">
                      <input
                        type="checkbox"
                        className="accent-primary"
                        checked={pagatoDi(d)}
                        disabled={pagBusy}
                        onChange={() => togglePagato(d)}
                        title={t("stip.pagatoTip")}
                      />
                    </td>
                    <td className="px-3 py-1.5 text-right tabular-nums text-muted-foreground">
                      {d.costoMedio ? eur(d.costoMedio) : "—"}
                    </td>
                  </tr>
                ))}
                <tr className="bg-muted/40 font-semibold">
                  <td className="px-3 py-2" colSpan={5}>
                    {t("common.total")} ({righe.length})
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {totRighe.ore ? totRighe.ore.toLocaleString("it-IT") : ""}
                  </td>
                  <td className="px-3 py-2" />
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.ord)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.str)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.ferie)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.mens)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.tfr)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{eur(totRighe.totale)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {eur(righe.reduce((a, d) => a + (nettoDi(d)?.stipendio ?? 0), 0))}
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {eur(righe.reduce((a, d) => a + (nettoDi(d)?.anticipo ?? 0), 0))}
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {eur(righe.reduce((a, d) => a + (nettoDi(d)?.saldo ?? 0), 0))}
                  </td>
                  <td className="px-3 py-2 text-right tabular-nums">
                    {eur(righe.reduce((a, d) => a + (versatoDi(d) ?? 0), 0))}
                  </td>
                  <td className="px-3 py-2 text-center text-xs">
                    {righe.filter(pagatoDi).length}/{righe.length}
                  </td>
                  <td className="px-3 py-2" />
                </tr>
              </tbody>
            </table>
          </div>

          {nettiSenzaCosto.length > 0 && (
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="text-sm font-semibold text-foreground">
                {t("stip.nettiSenzaCosto")} ({nettiSenzaCosto.length})
              </p>
              <p className="mb-2 text-[11px] text-muted-foreground">
                {t("stip.nettiSenzaCostoDesc")}
              </p>
              <table className="text-sm">
                <tbody>
                  {nettiSenzaCosto.map((d) => (
                    <tr key={d.nome} className="border-b border-border/60">
                      <td className="py-1 pr-4 font-medium">{d.nome}</td>
                      <td className="py-1 pr-4 text-xs text-muted-foreground">{d.appalto ?? ""}</td>
                      <td className="py-1 pr-4 text-right tabular-nums">{eur(d.saldo)}</td>
                      <td className="py-1">{selectMotivo(nameKey(d.nome), d.nome)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {(db?.mesi.length ?? 0) > 1 && (
            <div className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]">
              <p className="mb-2 flex items-center gap-2 text-sm font-semibold text-foreground">
                <Banknote className="h-4 w-4 text-primary" /> {t("stip.trend")}
              </p>
              <div className="flex flex-wrap gap-2">
                {db!.mesi.map((m) => (
                  <button
                    key={m.mese}
                    type="button"
                    onClick={() => setMeseSel(m.mese)}
                    className={`rounded-lg border px-3 py-1.5 text-xs ${m.mese === meseSel ? "border-primary bg-primary/10 text-primary" : "border-border text-foreground hover:bg-muted"}`}
                  >
                    {fmtMese(m.mese)}
                    <span className="ml-1.5 font-semibold tabular-nums">
                      {eur(totaleMese(m))} €
                    </span>
                  </button>
                ))}
              </div>
            </div>
          )}
        </>
      ) : (
        <p className="py-8 text-center text-sm text-muted-foreground">{t("stip.vuoto")}</p>
      )}
    </div>
  );
}
