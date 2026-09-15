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
  aggiornatoIl?: string;
  aggiornatoDa?: string;
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
      const periodo = C.periodo >= 0 ? String(r[C.periodo] ?? "").trim() || undefined : undefined;
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

export interface TotaliFiscaliMese {
  rate: number;
  corrente: number;
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
