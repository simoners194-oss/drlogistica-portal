// DR Portal — modulo Mezzi: helper UI condivisi e dizionario bilingue locale.
// Le etichette del modulo vivono qui (prefisso nessuno, chiavi corte) per non
// gonfiare i18n.tsx: `useMezzi()` segue la stessa lingua scelta nel portale.

import type { ReactNode } from "react";
import { useLang } from "@/lib/i18n";
import { Badge } from "@/components/ui/badge";
import { semaforoScadenza, type Mezzo, type MezzoStato, type Semaforo } from "@/lib/mezzi-types";

export const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

const it = {
  // Tab
  "tab.parco": "Parco mezzi",
  "tab.scadenze": "Scadenze",
  "tab.affidamenti": "Affidamenti",
  "tab.contratti": "Contratti",
  "tab.multe": "Multe",
  "tab.ztl": "ZTL",
  "tab.officina": "Officina",
  "tab.km": "Km & Carburante",
  "tab.parametri": "Parametri",
  // Comuni
  nuovo: "Nuovo",
  salva: "Salva",
  annulla: "Annulla",
  elimina: "Elimina",
  chiudi: "Chiudi",
  modifica: "Modifica",
  cerca: "Cerca…",
  tutti: "Tutti",
  tutte: "Tutte",
  targa: "Targa",
  mezzo: "Mezzo",
  stato: "Stato",
  appalto: "Appalto",
  autista: "Autista",
  data: "Data",
  dal: "Dal",
  al: "Al",
  note: "Note",
  importo: "Importo",
  comune: "Comune",
  nessunRecord: "Nessun record.",
  confermaElimina: "Eliminare definitivamente questo record?",
  // Parco
  gruppo: "Gruppo",
  operativi: "Operativi",
  altri: "Altri",
  proprieta: "Proprietà",
  noleggio: "Noleggio",
  diProprieta: "Di proprietà",
  tipoMezzo: "Tipo mezzo",
  marca: "Marca",
  modello: "Modello",
  societa: "Società",
  sede: "Sede",
  kmAttuali: "Km attuali",
  frigo: "Frigo",
  atp: "ATP",
  sponda: "Sponda",
  sostitutivoDi: "Sostitutivo di",
  dettaglio: "Dettaglio",
  anagrafica: "Anagrafica",
  libretto: "Libretto",
  documenti: "Documenti",
  telaio: "Telaio",
  immatricolazione: "Immatricolazione",
  intestatario: "Intestatario",
  classeEuro: "Classe Euro",
  alimentazione: "Alimentazione",
  potenza: "Potenza (kW)",
  portata: "Portata (kg)",
  massaVuoto: "Massa a vuoto (kg)",
  massaMax: "Massa max (kg)",
  dimensioni: "Dimensioni (m)",
  allestimento: "Allestimento",
  carrozzeria: "Carrozzeria",
  posti: "Posti",
  linkLibretto: "Link libretto",
  linkAssicurazione: "Link assicurazione",
  linkContratto: "Link contratto",
  apri: "apri",
  nuovoMezzo: "Nuovo mezzo",
  // Stati mezzo
  "stato.operativo": "Operativo",
  "stato.manutenzione": "In manutenzione",
  "stato.sostitutivo": "Sostitutivo",
  "stato.fermo": "Fermo",
  "stato.reso": "Reso",
  "stato.dirigenziale": "Dirigenziale",
  "stato.dismesso": "Dismesso",
  "stato.venduto": "Venduto",
  "stato.rottamato": "Rottamato",
  // Scadenze
  "sem.scaduta": "Scaduta",
  "sem.rosso": "≤ 30 gg",
  "sem.arancio": "≤ 60 gg",
  "sem.giallo": "≤ 90 gg",
  "sem.ok": "OK",
  tipo: "Tipo",
  scadenza: "Scadenza",
  giorni: "Giorni",
  gestita: "Gestita",
  nuovaScadenza: "Nuova scadenza",
  soloAperte: "Solo aperte",
  "tipoScad.revisione": "Revisione",
  "tipoScad.assicurazione": "Assicurazione",
  "tipoScad.bollo": "Bollo",
  "tipoScad.atp": "ATP",
  "tipoScad.tagliando": "Tagliando",
  "tipoScad.fine_contratto": "Fine contratto",
  "tipoScad.disdetta": "Finestra disdetta",
  "tipoScad.ztl": "Permesso ZTL",
  "tipoScad.patente": "Patente/CQC",
  "tipoScad.altro": "Altro",
  descrizione: "Descrizione",
  alertInviati: "Alert inviati",
  // Affidamenti
  nuovoAffidamento: "Nuovo affidamento",
  inCorso: "In corso",
  storico: "Storico",
  chiudiAffidamento: "Chiudi (riconsegna)",
  verbaleConsegna: "Verbale consegna",
  verbaleRiconsegna: "Verbale riconsegna",
  codiceDipendente: "Codice dipendente",
  // Contratti
  noleggiatore: "Noleggiatore",
  canone: "Canone €/mese",
  piuIva: "+ IVA",
  kmInclusi: "Km/mese inclusi",
  illimitati: "illimitati",
  extraKm: "Extra-km €",
  preavviso: "Preavviso disdetta (gg)",
  franchigie: "Franchigie",
  servizi: "Servizi inclusi",
  penali: "Penali",
  deposito: "Deposito €",
  attivo: "Attivo",
  cessato: "Cessato",
  fuoriMercato: "Contratti fuori mercato",
  mediana: "mediana categoria",
  scostamento: "scostamento",
  nessunOutlier: "Nessun contratto sopra la soglia: canoni in linea con la mediana di categoria.",
  nuovoContratto: "Nuovo contratto",
  inizio: "Inizio",
  fine: "Fine",
  // Multe
  nuovaMulta: "Nuova multa",
  dataInfrazione: "Data/ora infrazione",
  dataNotifica: "Data notifica",
  responsabilita: "Responsabilità",
  "resp.da_definire": "Da definire",
  "resp.autista": "Autista",
  "resp.ufficio": "Ufficio",
  "multa.ricevuta": "Ricevuta",
  "multa.contestata_dipendente": "Contestata al dipendente",
  "multa.risposta_ricevuta": "Risposta ricevuta",
  "multa.in_detrazione": "In detrazione",
  "multa.detratta": "Detratta",
  "multa.pagata": "Pagata",
  "multa.ricorso": "Ricorso",
  "multa.annullata": "Annullata",
  importoRidotto: "Importo ridotto",
  scadenzaPagamento: "Scadenza pagamento",
  autistaProposto: "proposto dallo storico affidamenti",
  storicoStati: "Storico stati",
  // ZTL
  nuovoPermesso: "Nuovo permesso / richiesta",
  "ztl.richiesto": "Richiesto",
  "ztl.attivo": "Attivo",
  "ztl.scaduto": "Scaduto",
  "ztl.respinto": "Respinto",
  dataRichiesta: "Data richiesta",
  protocollo: "Protocollo",
  verificaGiro: "Verifica giro",
  verificaGiroHint:
    "Seleziona mezzo e comuni ZTL attraversati dal giro: il sistema segnala i permessi mancanti o scaduti.",
  comuniAttraversati: "Comuni ZTL attraversati (separati da virgola)",
  verifica: "Verifica",
  permessoOk: "Permesso attivo",
  permessoMancante: "PERMESSO MANCANTE",
  // Officina
  nuovoIntervento: "Nuovo intervento",
  fornitore: "Fornitore",
  costo: "Costo €",
  costoTotale: "Costo totale",
  // Km & carburante
  nuovaLettura: "Nuova lettura km",
  nuovoRifornimento: "Nuovo rifornimento",
  km: "Km",
  fonte: "Fonte",
  litri: "Litri",
  carta: "Carta",
  letture: "Letture km",
  rifornimenti: "Rifornimenti",
  sforamenti: "Sforamenti km contrattuali (proiezione)",
  kmMedioMese: "Km medi/mese",
  extraMese: "Extra/mese",
  costoExtraMese: "Costo extra €/mese",
  nessunoSforamento:
    "Nessuno sforamento rilevato (servono almeno 2 letture km per mezzo con contratto a km limitati).",
  // Parametri
  emailAlert: "Email destinatari alert (separate da virgola)",
  soglieAlert: "Soglie alert (giorni, separati da virgola)",
  extraKmDefault: "Extra-km di default €",
  sogliaOutlier: "Soglia fuori mercato (%)",
  urlBaseDocumenti: "URL base cartella documenti (SharePoint)",
  causali: "Causali mancata consegna (una per riga)",
  costoOrario: "Costo orario autista €",
  costoCarburante: "Costo carburante €/litro",
  consumoMedio: "Consumo medio km/litro",
  esporta: "Esporta",
  esportaCsv: "Esporta CSV (per Excel)",
  backupJson: "Backup JSON",
  importaJson: "Importa database (JSON)",
  importaHint:
    "Incolla qui il JSON del database (seed iniziale o backup) e conferma. Sostituisce TUTTI i dati del modulo.",
  importa: "Importa",
  salvato: "Salvato.",
  eliminato: "Eliminato.",
  importato: "Database importato.",
  aggiornatoIl: "Ultimo aggiornamento",
} as const;

type MezziKey = keyof typeof it;

const en: Record<MezziKey, string> = {
  ...it,
  "tab.parco": "Fleet",
  "tab.scadenze": "Deadlines",
  "tab.affidamenti": "Assignments",
  "tab.contratti": "Contracts",
  "tab.multe": "Fines",
  "tab.ztl": "LTZ",
  "tab.officina": "Workshop",
  "tab.km": "Km & Fuel",
  "tab.parametri": "Settings",
  nuovo: "New",
  salva: "Save",
  annulla: "Cancel",
  elimina: "Delete",
  chiudi: "Close",
  modifica: "Edit",
  cerca: "Search…",
  tutti: "All",
  tutte: "All",
  targa: "Plate",
  mezzo: "Vehicle",
  stato: "Status",
  appalto: "Contract site",
  autista: "Driver",
  data: "Date",
  dal: "From",
  al: "To",
  note: "Notes",
  importo: "Amount",
  comune: "Municipality",
  nessunRecord: "No records.",
  confermaElimina: "Permanently delete this record?",
  gruppo: "Group",
  operativi: "Operational",
  altri: "Other",
  proprieta: "Ownership",
  noleggio: "Rental",
  diProprieta: "Owned",
  tipoMezzo: "Vehicle type",
  marca: "Make",
  modello: "Model",
  societa: "Company",
  sede: "Site",
  kmAttuali: "Current km",
  frigo: "Reefer",
  atp: "ATP",
  sponda: "Tail lift",
  sostitutivoDi: "Replacement for",
  dettaglio: "Details",
  anagrafica: "Master data",
  libretto: "Registration",
  documenti: "Documents",
  telaio: "VIN",
  immatricolazione: "First registration",
  intestatario: "Registered owner",
  classeEuro: "Euro class",
  alimentazione: "Fuel type",
  potenza: "Power (kW)",
  portata: "Payload (kg)",
  massaVuoto: "Kerb weight (kg)",
  massaMax: "Max mass (kg)",
  dimensioni: "Dimensions (m)",
  allestimento: "Body builder",
  carrozzeria: "Body type",
  posti: "Seats",
  linkLibretto: "Registration link",
  linkAssicurazione: "Insurance link",
  linkContratto: "Contract link",
  apri: "open",
  nuovoMezzo: "New vehicle",
  "stato.operativo": "Operational",
  "stato.manutenzione": "In maintenance",
  "stato.sostitutivo": "Replacement",
  "stato.fermo": "Idle",
  "stato.reso": "Returned",
  "stato.dirigenziale": "Executive",
  "stato.dismesso": "Retired",
  "stato.venduto": "Sold",
  "stato.rottamato": "Scrapped",
  "sem.scaduta": "Overdue",
  "sem.rosso": "≤ 30 days",
  "sem.arancio": "≤ 60 days",
  "sem.giallo": "≤ 90 days",
  "sem.ok": "OK",
  tipo: "Type",
  scadenza: "Due date",
  giorni: "Days",
  gestita: "Handled",
  nuovaScadenza: "New deadline",
  soloAperte: "Open only",
  "tipoScad.revisione": "MOT/Inspection",
  "tipoScad.assicurazione": "Insurance",
  "tipoScad.bollo": "Road tax",
  "tipoScad.atp": "ATP",
  "tipoScad.tagliando": "Service",
  "tipoScad.fine_contratto": "Contract end",
  "tipoScad.disdetta": "Cancellation window",
  "tipoScad.ztl": "LTZ permit",
  "tipoScad.patente": "Licence/CQC",
  "tipoScad.altro": "Other",
  descrizione: "Description",
  alertInviati: "Alerts sent",
  nuovoAffidamento: "New assignment",
  inCorso: "Active",
  storico: "History",
  chiudiAffidamento: "Close (return)",
  verbaleConsegna: "Handover report",
  verbaleRiconsegna: "Return report",
  codiceDipendente: "Employee code",
  noleggiatore: "Rental company",
  canone: "Rate €/month",
  piuIva: "+ VAT",
  kmInclusi: "Km/month included",
  illimitati: "unlimited",
  extraKm: "Extra-km €",
  preavviso: "Notice period (days)",
  franchigie: "Deductibles",
  servizi: "Included services",
  penali: "Penalties",
  deposito: "Deposit €",
  attivo: "Active",
  cessato: "Ended",
  fuoriMercato: "Above-market contracts",
  mediana: "category median",
  scostamento: "deviation",
  nessunOutlier: "No contract above the threshold: rates in line with category medians.",
  nuovoContratto: "New contract",
  inizio: "Start",
  fine: "End",
  nuovaMulta: "New fine",
  dataInfrazione: "Violation date/time",
  dataNotifica: "Notification date",
  responsabilita: "Responsibility",
  "resp.da_definire": "To define",
  "resp.autista": "Driver",
  "resp.ufficio": "Office",
  "multa.ricevuta": "Received",
  "multa.contestata_dipendente": "Contested to employee",
  "multa.risposta_ricevuta": "Reply received",
  "multa.in_detrazione": "Being deducted",
  "multa.detratta": "Deducted",
  "multa.pagata": "Paid",
  "multa.ricorso": "Appeal",
  "multa.annullata": "Cancelled",
  importoRidotto: "Reduced amount",
  scadenzaPagamento: "Payment due",
  autistaProposto: "suggested from assignment history",
  storicoStati: "Status history",
  nuovoPermesso: "New permit / request",
  "ztl.richiesto": "Requested",
  "ztl.attivo": "Active",
  "ztl.scaduto": "Expired",
  "ztl.respinto": "Rejected",
  dataRichiesta: "Request date",
  protocollo: "Reference no.",
  verificaGiro: "Route check",
  verificaGiroHint:
    "Pick a vehicle and the LTZ municipalities the route crosses: the system flags missing or expired permits.",
  comuniAttraversati: "LTZ municipalities crossed (comma separated)",
  verifica: "Check",
  permessoOk: "Permit active",
  permessoMancante: "PERMIT MISSING",
  nuovoIntervento: "New job",
  fornitore: "Supplier",
  costo: "Cost €",
  costoTotale: "Total cost",
  nuovaLettura: "New km reading",
  nuovoRifornimento: "New refuelling",
  km: "Km",
  fonte: "Source",
  litri: "Litres",
  carta: "Card",
  letture: "Km readings",
  rifornimenti: "Refuellings",
  sforamenti: "Contract km overruns (projection)",
  kmMedioMese: "Avg km/month",
  extraMese: "Extra/month",
  costoExtraMese: "Extra cost €/month",
  nessunoSforamento:
    "No overrun detected (needs at least 2 km readings per vehicle with a km-capped contract).",
  emailAlert: "Alert recipient emails (comma separated)",
  soglieAlert: "Alert thresholds (days, comma separated)",
  extraKmDefault: "Default extra-km €",
  sogliaOutlier: "Above-market threshold (%)",
  urlBaseDocumenti: "Documents folder base URL (SharePoint)",
  causali: "Failed-delivery reasons (one per line)",
  costoOrario: "Driver hourly cost €",
  costoCarburante: "Fuel cost €/litre",
  consumoMedio: "Average consumption km/litre",
  esporta: "Export",
  esportaCsv: "Export CSV (for Excel)",
  backupJson: "JSON backup",
  importaJson: "Import database (JSON)",
  importaHint:
    "Paste the database JSON here (initial seed or backup) and confirm. Replaces ALL module data.",
  importa: "Import",
  salvato: "Saved.",
  eliminato: "Deleted.",
  importato: "Database imported.",
  aggiornatoIl: "Last update",
};

/** Etichette del modulo Mezzi nella lingua corrente del portale. */
export function useMezzi(): { m: (k: MezziKey) => string; lang: string } {
  const { lang } = useLang();
  const dict = lang === "it" ? it : en;
  return { m: (k) => dict[k] ?? k, lang };
}

// ---------------------------------------------------------------------------
// Format helper
// ---------------------------------------------------------------------------

export function fmtData(iso?: string | null): string {
  if (!iso) return "—";
  const [y, mo, g] = iso.slice(0, 10).split("-");
  return y && mo && g ? `${g}/${mo}/${y}` : iso;
}

export function fmtEur(n?: number | null): string {
  if (n === undefined || n === null || Number.isNaN(n)) return "—";
  return n.toLocaleString("it-IT", { style: "currency", currency: "EUR" });
}

export function num(v: string): number | undefined {
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) && v.trim() !== "" ? n : undefined;
}

// ---------------------------------------------------------------------------
// Badge
// ---------------------------------------------------------------------------

const SEM_CLS: Record<Semaforo, string> = {
  scaduta: "bg-destructive text-destructive-foreground",
  rosso: "bg-destructive/15 text-destructive",
  arancio: "bg-status-out/20 text-status-out",
  giallo: "bg-status-break/20 text-status-break",
  ok: "bg-status-present/15 text-status-present",
};

export function SemaforoBadge({ giorni }: { giorni: number }) {
  const { m } = useMezzi();
  const s = semaforoScadenza(giorni);
  return (
    <Badge variant="outline" className={`border-transparent ${SEM_CLS[s]}`}>
      {m(`sem.${s}` as MezziKey)}
      {s !== "ok" && (
        <span className="ml-1 font-normal opacity-80">
          {giorni < 0 ? `(${giorni})` : `(${giorni}g)`}
        </span>
      )}
    </Badge>
  );
}

const STATO_CLS: Record<MezzoStato, string> = {
  operativo: "bg-status-present/15 text-status-present",
  manutenzione: "bg-status-break/20 text-status-break",
  sostitutivo: "bg-primary/10 text-primary",
  fermo: "bg-status-out/20 text-status-out",
  reso: "bg-muted text-muted-foreground",
  dirigenziale: "bg-primary/10 text-primary",
  dismesso: "bg-muted text-muted-foreground",
  venduto: "bg-muted text-muted-foreground",
  rottamato: "bg-muted text-muted-foreground",
};

export function StatoMezzoBadge({ stato }: { stato: MezzoStato }) {
  const { m } = useMezzi();
  return (
    <Badge variant="outline" className={`border-transparent ${STATO_CLS[stato]}`}>
      {m(`stato.${stato}` as MezziKey)}
    </Badge>
  );
}

// ---------------------------------------------------------------------------
// Piccoli building block per i form dei dialog
// ---------------------------------------------------------------------------

export function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <label className="block text-sm">
      <span className="mb-1 block text-xs font-medium text-muted-foreground">{label}</span>
      {children}
    </label>
  );
}

export function MezzoSelect({
  mezzi,
  value,
  onChange,
  soloOperativi,
  vuotoLabel,
}: {
  mezzi: Mezzo[];
  value: string;
  onChange: (id: string) => void;
  soloOperativi?: boolean;
  vuotoLabel?: string;
}) {
  const lista = [...mezzi]
    .filter((m) => !soloOperativi || m.gruppo === "operativo")
    .sort((a, b) => a.targa.localeCompare(b.targa));
  return (
    <select className={inputCls} value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="">{vuotoLabel ?? "— seleziona —"}</option>
      {lista.map((m) => (
        <option key={m.id} value={m.id}>
          {m.targa} {m.modello ? `· ${m.modello}` : ""}
        </option>
      ))}
    </select>
  );
}

export function targaDi(mezzi: Mezzo[], id?: string): string {
  return mezzi.find((m) => m.id === id)?.targa ?? id ?? "—";
}
