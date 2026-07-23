// DR Portal — Finanza (sezione direttore DR005): logica pura di parsing e
// classificazione dei movimenti dell'estratto conto bancario.
// -----------------------------------------------------------------------------
// Il file esportato dalla banca ha colonne: Data contabile, Data valuta,
// Importo, Divisa, Causale (codice ABI/CBI), Descrizione (testo libero), Canale.
// Qui si trasforma ogni riga grezza in un movimento classificato:
//   - `tipologia`   — categoria leggibile derivata da causale + descrizione
//   - `cliente`     — controparte estratta dalla descrizione (euristica)
//   - `nrFattura`   — riferimenti fattura trovati nella descrizione
//   - `daVerificare`— true quando l'estrazione è incerta → pagina Anomalie
//   - `chiave`      — chiave di deduplicazione DETERMINISTICA calcolata dai
//                     SOLI campi grezzi (mai da quelli sanati a mano), con
//                     indice di occorrenza: nello stesso estratto possono
//                     esistere movimenti legittimamente identici (es. più
//                     bonifici uguali lo stesso giorno) e NON sono doppioni.
// Ricaricare un file che si sovrappone a un periodo già importato scarta le
// righe con chiave già presente — anche se nel frattempo sono state sanate,
// perché la chiave dipende solo dall'input originale.
// I VALORI salvati su SharePoint restano in italiano (convenzione del portale).

export interface MovimentoRaw {
  /** Data contabile ISO (YYYY-MM-DD). */
  dataContabile: string;
  /** Data valuta ISO (YYYY-MM-DD). */
  dataValuta: string;
  importo: number;
  divisa: string;
  causale: string;
  descrizione: string;
}

export interface MovimentoParsed extends MovimentoRaw {
  /** Indice di occorrenza (1-based) tra le righe identiche dello stesso file. */
  occ: number;
  /** Chiave di deduplicazione (salvata nel Title della lista SharePoint). */
  chiave: string;
  tipologia: string;
  cliente: string;
  nrFattura: string;
  daVerificare: boolean;
}

// Valore speciale per annullare il gruppo di movimenti importati prima
// dell'introduzione della colonna ImportId (righe con ImportId vuoto).
export const LEGACY_IMPORT_ID = "__senza__";

// --- Tipologie (valori registrati su SharePoint, in italiano) ---------------
export const TIPOLOGIE_MOVIMENTO = [
  "Incasso",
  "Bonifico uscita",
  "Pagamento Salario",
  "Commissioni",
  "Pagamento carta",
  "PagoPA / Multe",
  "Utenze",
  "Addebito SDD",
  "Imposte / F24",
  "Imposta di bollo",
  "Finanziamento",
  "Assicurazioni",
  "Prelievo ATM",
  "Carte fidelity",
  "Storno",
  "Estero",
  "Altro",
] as const;
export type TipologiaMovimento = (typeof TIPOLOGIE_MOVIMENTO)[number];

// Mappa causale ABI/CBI → tipologia. Le causali non mappate → "Altro" + verifica.
const CAUSALE_TIPOLOGIA: Record<string, TipologiaMovimento> = {
  "480": "Incasso", // bonifico a vostro favore
  "260": "Bonifico uscita", // vostra disposizione
  "662": "Commissioni",
  "663": "Commissioni",
  "16H": "Commissioni",
  "16G": "Commissioni",
  "16I": "Commissioni",
  "16X": "Commissioni",
  "66G": "Commissioni",
  "66E": "Commissioni",
  "167": "Commissioni",
  "950": "Commissioni",
  "669": "Commissioni",
  "437": "Pagamento carta", // pagamento internet
  "118": "Pagamento carta", // POS
  "110": "Utenze",
  "50C": "Addebito SDD",
  "194": "Imposte / F24",
  "198": "Imposte / F24", // I24 Agenzia Entrate
  "195": "Imposta di bollo",
  "150": "Finanziamento",
  "174": "Assicurazioni",
  "53": "Prelievo ATM",
  "91A": "Prelievo ATM",
  "453": "Carte fidelity",
  "680": "Storno",
  "293": "Incasso", // ricevute/disposizioni elettroniche al dopo incasso
  "782": "Incasso", // assegni circolari
  ZL0: "Estero",
  "310": "Estero",
};

// --- Normalizzazioni --------------------------------------------------------

/** Spazi collassati, minuscole, trim. Base per chiave e confronti. */
export function normalizeTesto(s: string): string {
  return (s ?? "").toLowerCase().replace(/\s+/g, " ").trim();
}

/** Canonicalizza le forme societarie per non spezzare i raggruppamenti
 *  ("s.r.l.", "s r l", "srl." → "srl"; idem spa/snc/sas). La banca spezza il
 *  testo a colonna fissa, quindi si normalizzano anche i TRONCAMENTI della
 *  forma societaria estesa ("societa' a responsabilita'", anche mozza → srl)
 *  per non far comparire lo stesso cliente due volte nell'overview. */
export function canonicalCliente(nome: string): string {
  return (
    normalizeTesto(nome)
      .replace(/\s*-\s*/g, " ") // "siaed - spa" → "siaed spa"
      .replace(/\bs[\s.]*r[\s.]*l[\s.]*s?\b/g, "srl ")
      .replace(/\bs[\s.]*p[\s.]*a[\s.]*\b/g, "spa ")
      .replace(/\bs[\s.]*n[\s.]*c[\s.]*\b/g, "snc ")
      .replace(/\bs[\s.]*a[\s.]*s[\s.]*\b/g, "sas ")
      // Forma estesa in coda, anche troncata dalla banca (a volte senza spazi:
      // "societa'responsabilita…"): tutto ciò che segue collassa in "srl".
      .replace(/\bsociet[aà]'?\s*a?\s*responsabilit[a']*.*$/g, "srl")
      // "societa'" penzolante in coda (nome tagliato prima della forma).
      .replace(/\bsociet[aà]'?\s*$/g, "")
      .replace(/\s+/g, " ")
      .replace(/[.,;:]+$/g, "")
      .trim()
  );
}

/** Chiave di RAGGRUPPAMENTO per l'overview (non salvata): oltre alla forma
 *  canonica, ignora la forma societaria, uniforma italy/italia e ordina le
 *  parole, così si accorpano "panizza roberto"/"roberto panizza" e
 *  "kuwait petroleum italy"/"kuwait petroleum italia spa". Il nome mostrato
 *  resta quello registrato (variante più frequente del gruppo). */
export function clienteGroupKey(nome: string): string {
  return canonicalCliente(nome)
    .replace(/\bitaly\b/g, "italia")
    .replace(/\b(?:srls?|spa|snc|sas|sa|scarl|scpa)\b/g, " ")
    .replace(/[^a-z0-9àèéìòù]+/g, " ")
    .split(" ")
    .filter(Boolean)
    .sort()
    .join(" ");
}

// --- Persona fisica vs azienda ----------------------------------------------
// Un bonifico in uscita verso una persona fisica è (per prassi aziendale) un
// pagamento di salario. Euristica: 2-4 parole solo alfabetiche, nessun
// marcatore societario/istituzionale. I falsi positivi si sanano da Anomalie.
const RE_MARCATORE_AZIENDA =
  /(srl|spa|snc|sas|scarl|scpa|coop|consorzio|societ|gmbh|ltd|llc|sagl|\bkg\b|\bbv\b|banc[ao]|assicur|finanz|welfare|leasing|holding|group|servi[cz]|consulenz|sindacat|petroleum|energ|logistic|trasport|autotrasport|express|immobil|italia|italy|\.it\b|regolament|monetari|pagament|rimbors|stipend|cessione|fattur|nolegg)/;

export function isPersonaFisica(nome: string): boolean {
  const s = normalizeTesto(nome);
  if (!s || RE_MARCATORE_AZIENDA.test(s)) return false;
  const parole = s.split(" ").filter(Boolean);
  if (parole.length < 2 || parole.length > 4) return false;
  return parole.every((p) => /^[a-zàèéìòù']+$/.test(p));
}

// --- Chiave di deduplicazione ----------------------------------------------
// Costruita SOLO dai campi grezzi + indice di occorrenza. Descrizione troncata
// (il Title SharePoint ha limite 255): due movimenti diversi con stessi data,
// importo, causale e stessi primi 80 caratteri di descrizione sono di fatto
// la stessa riga di estratto → distinti dall'occorrenza.
export function chiaveMovimento(m: MovimentoRaw, occ: number): string {
  const desc = normalizeTesto(m.descrizione).slice(0, 80);
  return `${m.dataContabile}|${m.dataValuta}|${m.importo.toFixed(2)}|${normalizeTesto(
    m.causale,
  )}|${desc}|${occ}`;
}

/** Base della chiave senza occorrenza (per contare le righe identiche). */
function chiaveBase(m: MovimentoRaw): string {
  return chiaveMovimento(m, 0).replace(/\|0$/, "");
}

/** Assegna a ogni riga del file l'indice di occorrenza (1-based) tra le righe
 *  identiche, nell'ordine in cui compaiono. Deterministico: lo stesso file
 *  (o un file sovrapposto con le stesse righe nello stesso ordine relativo)
 *  produce le stesse chiavi. */
export function assegnaOccorrenze(rows: MovimentoRaw[]): { raw: MovimentoRaw; occ: number }[] {
  const counts = new Map<string, number>();
  return rows.map((raw) => {
    const base = chiaveBase(raw);
    const occ = (counts.get(base) ?? 0) + 1;
    counts.set(base, occ);
    return { raw, occ };
  });
}

// --- Estrazione controparte -------------------------------------------------

// Parole che chiudono il nome della controparte in un bonifico in entrata
// ("bon.da <nome> <riferimenti>"). La banca spezza il testo a colonna fissa,
// quindi l'estrazione resta euristica: i casi dubbi finiscono in Anomalie.
const STOP_INCASSO =
  /\s(?:ft|fpr|fatt|fattura|fatture|f\.\s*n|saldo|s\.do|sld|acconto|pag\.|pagamento|nr[\s.]|n\.\s|rif|\/ref\/|py\d|insoluto|restituzione|bonifico|supp\s|importo|dedott|compensat|secondo|terzo|quarto|primo|ottobre|novembre|dicembre|gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre)/;

function clean(nome: string): string {
  return nome.replace(/[\s\-.,;:]+$/g, "").trim();
}

/** Estrae la controparte da "bon.da <nome> <riferimenti>". */
function estraiClienteIncasso(descLower: string): { cliente: string; incerto: boolean } {
  const idx = descLower.indexOf("bon.da ");
  if (idx < 0) return { cliente: "", incerto: true };
  let rest = descLower.slice(idx + "bon.da ".length);
  let incerto = false;
  const stop = rest.search(STOP_INCASSO);
  const digit = rest.search(/\d/);
  let cut = -1;
  if (stop >= 0 && digit >= 0) cut = Math.min(stop, digit);
  else if (stop >= 0) cut = stop;
  else if (digit >= 0) cut = digit;
  if (cut > 0) rest = rest.slice(0, cut);
  else incerto = true; // nessun delimitatore: il "nome" potrebbe includere altro
  const cliente = canonicalCliente(clean(rest));
  if (cliente.length < 3 || cliente.length > 40) incerto = true;
  return { cliente, incerto };
}

/** Estrae il beneficiario da "vostra disposizione ... favore <nome> ...". */
function estraiClienteUscita(descLower: string): { cliente: string; incerto: boolean } {
  const m = descLower.match(/favore\s+(.+)$/);
  if (!m) return { cliente: "", incerto: true };
  // Il nome è seguito da spazi di riempimento e code tipo "- add.tot" /
  // "notprovide": si taglia al primo blocco di 3+ spazi o alle code note.
  let rest = m[1].split(/\s{3,}/)[0];
  rest = rest.replace(/\s*-?\s*(add\.tot|notprovide|da contab).*$/i, "");
  const cliente = canonicalCliente(clean(rest));
  return { cliente, incerto: cliente.length < 3 || cliente.length > 45 };
}

/** Estrae il creditore da "sdd b2b : <codice mandato> <nome>" (o sdd core). */
function estraiClienteSdd(descLower: string): { cliente: string; incerto: boolean } {
  const m = descLower.match(/sdd\s+(?:b2b|core)\s*:?\s*(\S+)\s+(.+)$/);
  if (!m) return { cliente: "", incerto: true };
  const cliente = canonicalCliente(clean(m[2]));
  return { cliente, incerto: cliente.length < 3 };
}

/** Estrae l'esercente da "... carta*XXXX-[HH:MM-]<esercente> ...". */
function estraiClienteCarta(descLower: string): { cliente: string; incerto: boolean } {
  const m = descLower.match(/carta\s*\*?\s*\d{4}\s*-\s*(?:\d{2}:\d{2}\s*-\s*)?(.+)$/);
  if (!m) return { cliente: "", incerto: true };
  let rest = m[1].split(/\s{3,}/)[0];
  rest = rest.replace(/\s*-?\s*da contab.*$/i, "");
  // Suffisso paese in coda (ita/lux/...) — informazione inutile per il gruppo.
  rest = rest.replace(/\s+(?:ita|lux|irl|nld|deu|fra|esp|gbr|usa)\s*$/i, "");
  let cliente = canonicalCliente(clean(rest));
  // Se c'è la forma societaria, il nome finisce lì: il resto è indirizzo/città
  // ("grossi srl via civesio 28 san giuliano" → "grossi srl").
  const conSuffisso = cliente.match(/^(.*?\b(?:srl|spa|snc|sas)\b)/);
  if (conSuffisso) cliente = conSuffisso[1];
  return { cliente, incerto: cliente.length < 3 };
}

// --- Riferimenti fattura ----------------------------------------------------
// Cattura sequenze tipo "ft 163/26", "fpr82/26", "fatt. n. 180", "ft.n.116fp",
// "nr. 341-371/25". Best-effort: si salva il testo trovato, non si interpreta.
const RE_FATTURA =
  /(?:\b(?:ft|fpr|fatt(?:ura|ure)?|n\.?c\.?)\b[\s.]*(?:n(?:r)?[\s.°]*)?|\bnr[\s.]+)([0-9][0-9a-z\/.,+\- ]{0,30})/g;

export function estraiNrFattura(descrizione: string): string {
  const found: string[] = [];
  const lower = normalizeTesto(descrizione);
  for (const m of lower.matchAll(RE_FATTURA)) {
    const ref = m[1].replace(/[\s.,\-]+$/g, "").trim();
    if (ref && !found.includes(ref)) found.push(ref);
    if (found.length >= 4) break;
  }
  return found.join("; ").slice(0, 120);
}

// --- Classificazione completa -----------------------------------------------

export function classificaMovimento(raw: MovimentoRaw): {
  tipologia: string;
  cliente: string;
  nrFattura: string;
  daVerificare: boolean;
} {
  const causale = (raw.causale ?? "").toString().trim().toUpperCase();
  const desc = normalizeTesto(raw.descrizione);
  let tipologia: TipologiaMovimento = CAUSALE_TIPOLOGIA[causale] ?? "Altro";
  let daVerificare = tipologia === "Altro";

  // Regole da descrizione che raffinano la causale (fonte: prassi aziendale).
  if (desc.includes("pagopa")) tipologia = "PagoPA / Multe";
  // "beneficiari vari distinta" = bonifico multiplo (paga in blocco l'elenco
  // di una distinta): è la modalità tipica del giro stipendi mensile.
  if (
    desc.includes("benefici vari") ||
    desc.includes("beneficiari vari") ||
    desc.includes("stipend")
  )
    tipologia = "Pagamento Salario";
  if (tipologia === "Storno") daVerificare = true; // semantica ambigua: sempre da verificare

  let cliente = "";
  let incerto = false;
  if (desc.includes("bon.da ")) {
    ({ cliente, incerto } = estraiClienteIncasso(desc));
  } else if (tipologia === "Bonifico uscita" || tipologia === "Pagamento Salario") {
    ({ cliente, incerto } = estraiClienteUscita(desc));
    // Bonifico verso una persona fisica → pagamento di salario.
    if (tipologia === "Bonifico uscita" && !incerto && isPersonaFisica(cliente))
      tipologia = "Pagamento Salario";
  } else if (tipologia === "Addebito SDD") {
    ({ cliente, incerto } = estraiClienteSdd(desc));
  } else if (tipologia === "Pagamento carta" || tipologia === "PagoPA / Multe") {
    ({ cliente, incerto } = estraiClienteCarta(desc));
    // L'esercente carta è informativo: se non estratto non è un'anomalia.
    if (tipologia === "PagoPA / Multe") incerto = false;
  } else if (tipologia === "Incasso") {
    // Incasso non-bonifico (Ri.Ba./assegni): controparte non presente nel testo.
    cliente = "";
    incerto = causale === "480"; // un 480 senza "bon.da" è anomalo
  }
  // Per un incasso la controparte è il dato chiave dell'Overview: se manca
  // o è incerta va sanata a mano.
  if (tipologia === "Incasso" && raw.importo > 0 && (incerto || !cliente)) daVerificare = true;
  if (incerto && cliente === "") daVerificare = true;
  if ((tipologia === "Bonifico uscita" || tipologia === "Addebito SDD") && incerto)
    daVerificare = true;

  const nrFattura = estraiNrFattura(raw.descrizione);
  return { tipologia, cliente, nrFattura, daVerificare };
}

/** Pipeline completa su un file: occorrenze → chiavi → classificazione. */
export function parseEstratto(rows: MovimentoRaw[]): MovimentoParsed[] {
  return assegnaOccorrenze(rows).map(({ raw, occ }) => ({
    ...raw,
    occ,
    chiave: chiaveMovimento(raw, occ),
    ...classificaMovimento(raw),
  }));
}

// --- Lettura del foglio Excel (lato client) ---------------------------------
// Il foglio dei movimenti cambia nome a ogni export ("MovimentiCC_OnLine_..."):
// si riconosce dalle intestazioni. Le date possono arrivare come Date, numero
// seriale Excel o stringa "dd/mm/yyyy".
export const HEADER_MOVIMENTI = [
  "data contabile",
  "data valuta",
  "importo",
  "divisa",
  "causale",
  "descrizione",
] as const;

export function isHeaderMovimenti(row: unknown[]): boolean {
  const cells = row.map((c) => normalizeTesto(String(c ?? "")));
  return HEADER_MOVIMENTI.every((h, i) => cells[i] === h);
}

/** Converte una cella data (Date | seriale Excel | "dd/mm/yyyy" | ISO) in
 *  "YYYY-MM-DD"; null se non interpretabile. */
export function cellToIsoDate(v: unknown): string | null {
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    // Le date Excel sono "wall clock": si usano i componenti locali.
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "number" && Number.isFinite(v) && v > 20000 && v < 60000) {
    // Seriale Excel (giorni dal 1900-01-00, con il bug dell'anno 1900).
    const ms = Math.round((v - 25569) * 86400000);
    return new Date(ms).toISOString().slice(0, 10);
  }
  const s = String(v ?? "").trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return null;
}

/** Converte una cella importo ("1.234,56" | "-48" | numero) in numero. */
export function cellToImporto(v: unknown): number | null {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s = String(v ?? "")
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.\-+]/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

export interface ParseFileResult {
  rows: MovimentoRaw[];
  scartate: number; // righe non interpretabili (data o importo mancanti)
}

/** Estrae i movimenti da una matrice di celle (foglio già letto). La prima
 *  riga che combacia con HEADER_MOVIMENTI delimita l'inizio dei dati. */
export function parseMatrice(matrix: unknown[][]): ParseFileResult | null {
  const headerIdx = matrix.findIndex((r) => isHeaderMovimenti(r));
  if (headerIdx < 0) return null;
  const rows: MovimentoRaw[] = [];
  let scartate = 0;
  for (const r of matrix.slice(headerIdx + 1)) {
    const dataContabile = cellToIsoDate(r[0]);
    const dataValuta = cellToIsoDate(r[1]) ?? dataContabile;
    const importo = cellToImporto(r[2]);
    const descrizione = String(r[5] ?? "").trim();
    if (!dataContabile || importo == null || (!descrizione && r[4] == null)) {
      // Riga vuota in coda o non interpretabile.
      if (r.some((c) => c != null && String(c).trim() !== "")) scartate++;
      continue;
    }
    rows.push({
      dataContabile,
      dataValuta: dataValuta ?? dataContabile,
      importo,
      divisa: String(r[3] ?? "EUR").trim() || "EUR",
      causale: String(r[4] ?? "").trim(),
      descrizione,
    });
  }
  return { rows, scartate };
}
