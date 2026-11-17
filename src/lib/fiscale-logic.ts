// DR Portal — Finanza → Flussi: logica pura dello scadenziario fiscale.
// -----------------------------------------------------------------------------
// Fonte: OneDrive "Fiscale\SCADENZARIO FISCALE__aggiornato.xlsx" (Sabrina).
// Il foglio "database" (il nome cambia a ogni salvataggio, es. "14092026") è
// una tabella piatta: VOCE OLD | VOCE | ANNO | PERIODO | Quantità |
// Data pagamento | Pagato SI/NO | Modalità pagamento | Note. I fogli
// "DA PAGARE"/"PAGATO"/"Dettagli1" sono pivot derivate e si ignorano; il
// foglio "Da registrare in futuro" elenca importi noti ma senza data
// ("Da rateizzare") che si segnalano soltanto.

export interface ScadenzaFiscale {
  /** Id stabile per modifiche puntuali dal portale (assegnato dal server). */
  id?: string;
  /** "portale" = creata/modificata su DR Portal: un re-import del file NON
   *  la tocca (le righe da file vengono invece sostituite in blocco). */
  origine?: "file" | "portale";
  /** Chiave contenuto della riga come stava NEL FILE (voce|anno|periodo|
   *  data|importo): serve al re-import per non duplicare le righe da file
   *  poi modificate sul portale e per non far risorgere quelle eliminate. */
  chiaveFile?: string;
  voce: string; // IVA, INPS, IRAP, IRES, MOD 770, REDDITI, …
  voceOld?: string; // etichetta storica del consulente ("IVA II TRIM25"…)
  anno?: number; // anno di competenza
  periodo?: string; // "MAGGIO", "II TRIMESTRE", "ANNUALE", "MENSILE"…
  importo: number;
  dataPagamento: string; // ISO YYYY-MM-DD (scadenza di pagamento)
  pagato: boolean;
  modalita?: string; // F24, QR CODE (cartella), RID…
  note?: string;
  /** "rate" = debiti di anni precedenti/cartelle; "corrente" = imposte
   *  dell'anno; "finanziamento" = rate di finanziamenti/noleggi (non
   *  tributarie: NON alimentano le voci fiscali del cash flow). */
  categoria: "rate" | "corrente" | "finanziamento";
}

export interface DaRateizzareFiscale {
  voce: string;
  anno?: number;
  periodo?: string;
  importo: number;
}

export interface FiscaleDb {
  versione: number;
  scadenze: ScadenzaFiscale[];
  daRateizzare: DaRateizzareFiscale[];
  /** Chiavi file delle righe eliminate dal portale: il re-import le salta. */
  tombstones?: string[];
  /** true dal primo salvataggio del "da registrare" dal portale: da lì in
   *  poi il re-import del file NON tocca più quell'elenco. */
  daRateizzarePortale?: boolean;
  fonteFile?: string;
  /** Data di modifica (ISO) del file della cartella Fiscale letto per ultimo:
   *  finché non cambia, il portale non lo riscarica. */
  fonteModificato?: string;
  /** Com'è andata l'ultima lettura automatica dalla cartella. */
  ultimaLettura?: LetturaFiscale;
  aggiornatoIl?: string;
  aggiornatoDa?: string;
}

/** Esito di una lettura dello scadenzario dalla cartella condivisa "Fiscale"
 *  (1.120.0). "letto" = il file era nuovo o cambiato ed è entrato; "uguale" =
 *  il file è quello già letto; "assente" = nella cartella non c'è; "errore" =
 *  non si è potuto leggere e l'archivio resta com'era. */
export interface LetturaFiscale {
  il: string;
  da: string;
  file?: string;
  /** Quando il file è stato salvato l'ultima volta sulla cartella (ISO). */
  modificato?: string;
  esito: "letto" | "uguale" | "assente" | "errore";
  messaggio?: string;
  /** Scadenze cambiate rispetto a prima (nuove, tolte, pagate o riaperte). */
  cambiate?: number;
}

export function emptyFiscaleDb(): FiscaleDb {
  return { versione: 0, scadenze: [], daRateizzare: [] };
}

export function chiaveScadenzaFile(s: ScadenzaFiscale): string {
  return [s.voce, s.anno ?? "", s.periodo ?? "", s.dataPagamento, s.importo.toFixed(2)].join("|");
}

const norm = (s: unknown) =>
  String(s ?? "")
    .trim()
    .toLowerCase();

function numCell(v: unknown): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  let s = String(v ?? "").trim();
  if (!s) return 0;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

/** Data di cella → ISO. Con `raw:true` le date Excel arrivano come numero
 *  seriale (giorni dal 30/12/1899); si accettano anche stringhe ISO o
 *  italiane. "" se la cella non è una data (es. "Da rateizzare"). */
export function dataCellaIso(v: unknown): string {
  if (typeof v === "number" && Number.isFinite(v) && v > 25569 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const s = String(v ?? "").trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const it = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (it) return `${it[3]}-${it[2].padStart(2, "0")}-${it[1].padStart(2, "0")}`;
  return "";
}

const MESI_PERIODO = [
  "GENNAIO",
  "FEBBRAIO",
  "MARZO",
  "APRILE",
  "MAGGIO",
  "GIUGNO",
  "LUGLIO",
  "AGOSTO",
  "SETTEMBRE",
  "OTTOBRE",
  "NOVEMBRE",
  "DICEMBRE",
];

/** Rate di finanziamenti e noleggi presenti nello scadenziario di Sabrina
 *  ma NON tributarie: restano fuori dalle voci fiscali del cash flow. */
const RE_FINANZIAMENTO = /FIN\.?\s*TO|FINANZIAMENT|NOLEGGI|CONTRATTO\s+NR/i;

/** Classificazione per le due voci del cash flow. "rate" = cartelle QR
 *  code, piani di dilazione (il PERIODO riporta l'intervallo "mm/aaaa-
 *  mm/aaaa"), competenze di 2+ anni prima, o dell'anno prima pagate da
 *  agosto in poi. L'anno prima pagato ENTRO luglio è fisiologico e resta
 *  "corrente": IVA di dicembre a gennaio, saldi annuali a giugno-luglio. */
export function classificaScadenza(
  voce: string,
  anno: number | undefined,
  dataPagamento: string,
  modalita: string,
  periodo = "",
): ScadenzaFiscale["categoria"] {
  if (RE_FINANZIAMENTO.test(voce)) return "finanziamento";
  if (/qr\s*code/i.test(modalita)) return "rate";
  if (/\d{1,2}\/\d{4}\s*[-–]\s*\d{1,2}\/\d{4}/.test(periodo)) return "rate";
  const annoPag = Number(dataPagamento.slice(0, 4));
  if (anno == null || !Number.isFinite(annoPag)) return "corrente";
  if (anno < annoPag - 1) return "rate";
  if (anno === annoPag - 1) {
    const mesePag = Number(dataPagamento.slice(5, 7));
    return mesePag >= 8 ? "rate" : "corrente";
  }
  return "corrente";
}

export interface ParseScadenzarioResult {
  scadenze: ScadenzaFiscale[];
  daRateizzare: DaRateizzareFiscale[];
  /** Righe non pagate senza data valida (segnalate, non importate). */
  senzaData: number;
  foglio: string;
}

/** Interpreta il file scadenziario: cerca in OGNI foglio la tabella piatta
 *  (header VOCE + QUANTITÀ + DATA PAGAMENTO + PAGATO) e tiene quella con
 *  più righe (i pivot "Dettagli1" hanno lo stesso header ma poche righe).
 *  null se nessun foglio è nel tracciato. */
export function parseScadenzario(
  fogli: { nome: string; matrix: unknown[][] }[],
): ParseScadenzarioResult | null {
  let migliore: ParseScadenzarioResult | null = null;
  const daRateizzare: DaRateizzareFiscale[] = [];

  for (const f of fogli) {
    const headerIdx = f.matrix.findIndex((r) => {
      const cells = (r ?? []).map(norm);
      return (
        cells.some((c) => c === "voce") &&
        cells.some((c) => c.startsWith("quantit")) &&
        cells.some((c) => c.startsWith("data pagamento"))
      );
    });
    if (headerIdx < 0) continue;
    const header = (f.matrix[headerIdx] ?? []).map(norm);
    const col = (n: string) => header.findIndex((c) => c === n || c.startsWith(n));
    const C = {
      voceOld: col("voce old"),
      voce: header.findIndex((c) => c === "voce"),
      anno: col("anno"),
      periodo: col("periodo"),
      importo: col("quantit"),
      data: col("data pagamento"),
      pagato: col("pagato"),
      modalita: col("modalit"),
      note: col("note"),
      inseritoDb: col("inserito in db"),
    };
    // Senza la colonna Pagato ogni riga storica conterebbe come da pagare:
    // meglio rifiutare il foglio (fc.errFiscale la dichiara richiesta).
    if (C.voce < 0 || C.importo < 0 || C.data < 0 || (C.pagato < 0 && C.inseritoDb < 0)) continue;

    // Foglio "Da registrare in futuro": importi noti ma senza scadenza
    // ("Da rateizzare") — si raccolgono a parte, da qualunque foglio.
    const daRegistrare = C.inseritoDb >= 0 || /da registrare/i.test(f.nome);

    const scadenze: ScadenzaFiscale[] = [];
    let senzaData = 0;
    for (const r of f.matrix.slice(headerIdx + 1)) {
      const voce = String(r?.[C.voce] ?? "").trim();
      const importo = Math.round(numCell(r?.[C.importo]) * 100) / 100;
      if (!voce || importo <= 0) continue;
      const annoRaw = C.anno >= 0 ? numCell(r[C.anno]) : 0;
      const anno = annoRaw >= 2000 && annoRaw < 2100 ? Math.round(annoRaw) : undefined;
      // PERIODO scritto come data in Excel ("INPS CONTRIBUTI AGOSTO 2026" ha
      // 01/08/2026 nella cella): arrivava il seriale 46235 e la voce si
      // leggeva "INPS 46235" — si rende "AGOSTO 2026" (caso 20/10/2026).
      const periodoRaw = C.periodo >= 0 ? r[C.periodo] : undefined;
      const periodoData = dataCellaIso(periodoRaw);
      const periodo =
        periodoData &&
        (typeof periodoRaw === "number" || /^\d{4}-\d{2}-\d{2}/.test(String(periodoRaw)))
          ? `${MESI_PERIODO[Number(periodoData.slice(5, 7)) - 1]} ${periodoData.slice(0, 4)}`
          : C.periodo >= 0
            ? String(periodoRaw ?? "").trim() || undefined
            : undefined;
      const data = dataCellaIso(r[C.data]);
      if (daRegistrare) {
        // Foglio "da registrare": senza data = importo noto da rateizzare;
        // con data = già presente nel foglio database, si ignora.
        if (!data) daRateizzare.push({ voce, anno, periodo, importo });
        continue;
      }
      const pagato = norm(r[C.pagato] ?? "").startsWith("s"); // SI/SÌ
      if (!data) {
        // Contano solo le NON pagate: sono quelle che mancano al cash flow.
        if (!pagato) senzaData++;
        continue;
      }
      const modalita =
        C.modalita >= 0 ? String(r[C.modalita] ?? "").trim() || undefined : undefined;
      scadenze.push({
        voce,
        voceOld: C.voceOld >= 0 ? String(r[C.voceOld] ?? "").trim() || undefined : undefined,
        anno,
        periodo,
        importo,
        dataPagamento: data,
        pagato,
        modalita,
        note: C.note >= 0 ? String(r[C.note] ?? "").trim() || undefined : undefined,
        categoria: classificaScadenza(voce, anno, data, modalita ?? "", periodo ?? ""),
      });
    }
    if (
      !daRegistrare &&
      scadenze.length > 0 &&
      (!migliore || scadenze.length > migliore.scadenze.length)
    ) {
      migliore = { scadenze, daRateizzare: [], senzaData, foglio: f.nome };
    }
  }
  if (!migliore) return null;
  // Dedup del "da rateizzare" (lo stesso foglio può riapparire più volte).
  const visti = new Set<string>();
  migliore.daRateizzare = daRateizzare.filter((d) => {
    const k = `${d.voce}|${d.anno ?? ""}|${d.periodo ?? ""}|${d.importo}`;
    if (visti.has(k)) return false;
    visti.add(k);
    return true;
  });
  return migliore;
}

// --- Lettura automatica dalla cartella (1.120.0) -----------------------------
// Simone 17/11/2026: "sì, fallo automatico". Fino alla 1.119 lo scadenzario
// entrava solo caricando il file a mano dai Flussi: dopo ogni scadenza pagata
// la Cassa la teneva tra le uscite finché qualcuno non lo ricaricava. Ora il
// portale guarda da solo la cartella condivisa "Fiscale".

/** Vero se il nome è quello di uno scadenzario fiscale di Excel (non i file
 *  temporanei "~$…" che Excel crea mentre il file è aperto). */
export function fileScadenzarioFiscale(nome: string): boolean {
  const n = (nome ?? "").trim();
  return !n.startsWith("~$") && /scadenzario\s+fiscale/i.test(n) && /\.xlsx?$/i.test(n);
}

/** Tra i file della cartella sceglie lo scadenzario buono: quello col nome
 *  "…aggiornato", altrimenti il salvato più di recente. */
export function scegliFileScadenzario<T extends { name: string; lastModifiedDateTime?: string }>(
  file: readonly T[],
): T | undefined {
  const buoni = file.filter((f) => fileScadenzarioFiscale(f.name));
  const perData = (a: T, b: T) =>
    (b.lastModifiedDateTime ?? "").localeCompare(a.lastModifiedDateTime ?? "");
  const aggiornati = buoni.filter((f) => /aggiornato/i.test(f.name)).sort(perData);
  return aggiornati[0] ?? [...buoni].sort(perData)[0];
}

/** Unisce le scadenze lette da un file a quelle dell'archivio: sostituisce le
 *  righe DA FILE, conserva quelle create o modificate sul portale (origine
 *  "portale"). La chiaveFile evita i doppioni: una riga del file già
 *  modificata sul portale NON viene reimportata (vince la versione portale),
 *  una eliminata dal portale NON risorge (tombstone). Gli id delle righe da
 *  file restano STABILI tra un import e l'altro (match per chiaveFile).
 *  LIMITE NOTO: se la STESSA riga viene cambiata sia sul portale sia
 *  nell'Excel, la chiave diverge e compaiono entrambe: da sanare a mano.
 *  `cambiate` = quante scadenze sono diverse da prima (nuove, tolte, oppure
 *  passate da pagare a pagate o il contrario). */
export function unisciImportFiscale(
  db: FiscaleDb,
  scadenze: readonly ScadenzaFiscale[],
  daRateizzare: readonly DaRateizzareFiscale[],
  fonte: string,
): { db: FiscaleDb; cambiate: number } {
  const portale = db.scadenze.filter((s) => s.origine === "portale");
  const chiaviPortale = new Set(portale.map((s) => s.chiaveFile).filter(Boolean));
  const tombstones = new Set(db.tombstones ?? []);
  const idVecchi = new Map(
    db.scadenze
      .filter((s) => s.origine !== "portale")
      .map((s) => [s.chiaveFile ?? chiaveScadenzaFile(s), s.id] as const),
  );
  const nuove: ScadenzaFiscale[] = [];
  for (const letta of scadenze) {
    const s: ScadenzaFiscale = { ...letta, origine: "file" };
    s.chiaveFile = chiaveScadenzaFile(s);
    if (chiaviPortale.has(s.chiaveFile)) continue;
    if (tombstones.has(s.chiaveFile)) continue;
    s.id = idVecchi.get(s.chiaveFile);
    nuove.push(s);
  }
  // Quante cambiano: si contano le chiavi (voce, anno, periodo, data, importo)
  // con lo stato pagato; la stessa chiave può stare su più righe.
  const conta = (lista: readonly ScadenzaFiscale[]) => {
    const m = new Map<string, number>();
    for (const s of lista) {
      const k = `${s.chiaveFile ?? chiaveScadenzaFile(s)}|${s.pagato ? 1 : 0}`;
      m.set(k, (m.get(k) ?? 0) + 1);
    }
    return m;
  };
  const prima = conta(db.scadenze.filter((s) => s.origine !== "portale"));
  const dopo = conta(nuove);
  let tolte = 0;
  let aggiunte = 0;
  for (const [k, n] of prima) tolte += Math.max(0, n - (dopo.get(k) ?? 0));
  for (const [k, n] of dopo) aggiunte += Math.max(0, n - (prima.get(k) ?? 0));
  return {
    db: {
      ...db,
      scadenze: [...nuove, ...portale],
      // Appena il "da registrare" viene gestito dal portale, l'Excel non lo
      // comanda più: niente righe risorte né aggiunte manuali cancellate.
      daRateizzare: db.daRateizzarePortale ? db.daRateizzare : [...daRateizzare],
      fonteFile: fonte,
    },
    // Una riga che passa da "da pagare" a "pagata" esce da una parte ed entra
    // dall'altra: conta una volta.
    cambiate: Math.max(tolte, aggiunte),
  };
}

/** Prudenza della lettura AUTOMATICA: se il file ha molte meno scadenze
 *  dell'archivio (foglio rovinato, tabella spostata) non lo si sostituisce da
 *  soli. Restituisce il motivo, o null se si può procedere. */
export function importFiscaleSospetto(
  db: FiscaleDb,
  scadenze: readonly ScadenzaFiscale[],
): string | null {
  const daFile = db.scadenze.filter((s) => s.origine !== "portale").length;
  if (daFile >= 20 && scadenze.length < daFile * 0.6)
    return `il file ha ${scadenze.length} scadenze, l'archivio ne ha ${daFile}: troppo poche per sostituirlo da solo (se il file è giusto, caricalo a mano)`;
  return null;
}

export interface TotaliFiscaliMese {
  rate: number;
  corrente: number;
}

/** Ripartizione per i Flussi (v1.85.0, come gli stipendi): le NON pagate
 *  già scadute (data < oggi) vanno nella colonna Scaduto; il MESE CORRENTE
 *  vale il reale — le pagate del mese più le non pagate che scadono da oggi
 *  in poi entro il mese; i mesi futuri le non pagate in scadenza. Le pagate
 *  dei mesi passati non contano (storia già in banca); finanziamenti fuori. */
export interface RipartizioneFiscale {
  scaduto: TotaliFiscaliMese;
  perMese: Map<string, TotaliFiscaliMese>;
}

export function ripartizioneFiscale(
  scadenze: readonly ScadenzaFiscale[],
  oggiISO: string,
): RipartizioneFiscale {
  const meseCorrente = oggiISO.slice(0, 7);
  const scaduto: TotaliFiscaliMese = { rate: 0, corrente: 0 };
  const perMese = new Map<string, TotaliFiscaliMese>();
  const aggiungi = (t: TotaliFiscaliMese, s: ScadenzaFiscale) => {
    t[s.categoria === "rate" ? "rate" : "corrente"] += s.importo;
  };
  for (const s of scadenze) {
    if (s.categoria === "finanziamento") continue;
    const data = s.dataPagamento.slice(0, 10);
    const mese = data.slice(0, 7);
    if (s.pagato) {
      if (mese !== meseCorrente) continue;
    } else if (data < oggiISO) {
      aggiungi(scaduto, s);
      continue;
    }
    const t = perMese.get(mese) ?? { rate: 0, corrente: 0 };
    aggiungi(t, s);
    perMese.set(mese, t);
  }
  const r2 = (n: number) => Math.round(n * 100) / 100;
  scaduto.rate = r2(scaduto.rate);
  scaduto.corrente = r2(scaduto.corrente);
  for (const t of perMese.values()) {
    t.rate = r2(t.rate);
    t.corrente = r2(t.corrente);
  }
  return { scaduto, perMese };
}

/** Totali per mese delle scadenze NON pagate (voci fiscali soltanto: i
 *  finanziamenti/noleggi restano fuori). Le scadenze non pagate già
 *  SCADUTE si spostano sul mese corrente: sono ancora da pagare. */
export function totaliFiscaliPerMese(
  scadenze: ScadenzaFiscale[],
  meseCorrente: string,
): Map<string, TotaliFiscaliMese> {
  const out = new Map<string, TotaliFiscaliMese>();
  for (const s of scadenze) {
    if (s.pagato || s.categoria === "finanziamento") continue;
    const mese =
      s.dataPagamento.slice(0, 7) < meseCorrente ? meseCorrente : s.dataPagamento.slice(0, 7);
    const t = out.get(mese) ?? { rate: 0, corrente: 0 };
    t[s.categoria === "rate" ? "rate" : "corrente"] += s.importo;
    out.set(mese, t);
  }
  for (const t of out.values()) {
    t.rate = Math.round(t.rate * 100) / 100;
    t.corrente = Math.round(t.corrente * 100) / 100;
  }
  return out;
}
