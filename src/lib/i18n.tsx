// DR Portal — internazionalizzazione IT/EN (leggera, senza librerie).
// -----------------------------------------------------------------------------
// - Default INGLESE (richiesta della direzione); la scelta utente è ricordata
//   per dispositivo in localStorage ("dr:lang").
// - `useLang()` espone { lang, setLang, t }. Le chiavi mancanti degradano
//   sull'inglese e infine sulla chiave stessa (mai crash).
// - I VALORI dei dati SharePoint (stati, voci, nomi sedi) restano quelli
//   registrati; si traducono le etichette dell'interfaccia. Per gli stati
//   applicativi c'è `tStato()` che mappa i valori noti.
// -----------------------------------------------------------------------------

import { createContext, useContext, useEffect, useState, type ReactNode } from "react";

export type Lang = "en" | "it";
const STORAGE_KEY = "dr:lang";
const DEFAULT_LANG: Lang = "en";

const en = {
  // --- Navigazione / moduli ---
  "module.dashboard": "Dashboard",
  "module.presenze": "Attendance",
  "module.richieste": "Requests",
  "module.procurement": "Procurement",
  "module.documenti": "Documents",
  "module.comunicazioni": "Communications",
  "module.gestione-timbrature": "Time entries",
  "module.supervisione": "Supervision",
  "module.report": "Summary report",
  "module.amministrazione": "Administration",
  "module.novita": "What's new",
  "nav.modules": "Modules",
  "nav.comingSoon": "Coming soon",
  "nav.notActive": "Not active",
  // --- Shell ---
  "shell.logout": "Sign out",
  "shell.logoutConfirm": "Are you sure you want to sign out?",
  "shell.role.dipendente": "Employee",
  "shell.role.responsabile": "Manager",
  "shell.role.amministratore_sistema": "System administrator",
  // --- Login ---
  "login.title": "Sign in",
  "login.subtitle": "Enter your employee code and company PIN.",
  "login.codePlaceholder": "Employee code",
  "login.pinPlaceholder": "PIN",
  "login.submit": "Sign in",
  "login.checking": "Checking…",
  "login.invalid": "Invalid code or PIN.",
  "login.welcome": "Welcome",
  "login.msSoon": "Microsoft 365 authentication coming soon.",
  "login.hero":
    "A modular platform bringing Attendance, Requests, Reports and Administration together in one integrated experience.",
  // --- Comuni ---
  "common.loading": "Loading…",
  "common.all": "All",
  "common.allF": "All",
  "common.send": "Send",
  "common.save": "Save",
  "common.cancel": "Cancel",
  "common.approve": "Approve",
  "common.reject": "Reject",
  "common.confirmReject": "Confirm rejection",
  "common.rejectReason": "Rejection reason",
  "common.required": "(required)",
  "common.optional": "(optional)",
  "common.exportCsv": "Export CSV",
  "common.open": "open",
  "common.select": "— select —",
  "common.site": "Site",
  "common.allSites": "All sites",
  "common.employee": "Employee",
  "common.status": "Status",
  "common.date": "Date",
  "common.from": "From",
  "common.to": "To",
  "common.type": "Type",
  "common.amount": "Amount",
  "common.total": "Total",
  "common.attachment": "attachment",
  "common.doc": "Doc.",
  "common.error": "Error",
  "common.restricted": "Restricted access",
  // --- Stati applicativi (valori dato → etichetta) ---
  "stato.Bozza": "Draft",
  "stato.Inviata": "Submitted",
  "stato.Comunicata": "Notified",
  "stato.Approvata": "Approved",
  "stato.Respinta": "Rejected",
  "stato.Annullata": "Cancelled",
  "stato.Da inviare": "To send",
  // --- Eventi timbratura / stati presenza ---
  "evento.entrata": "Clock in",
  "evento.inizio-pausa": "Start break",
  "evento.fine-pausa": "End break",
  "evento.uscita": "Clock out",
  "dstato.presente": "Present",
  "dstato.pausa": "On break",
  "dstato.assente": "Absent",
  "dstato.oltre": "Overtime",
  // --- Presenze ---
  "presenze.title": "My attendance",
  "presenze.statusLabel": "Status:",
  "presenze.hoursToday": "Hours today:",
  "presenze.lastEntry": "Last clock event",
  "presenze.noneToday": "No clock events recorded today.",
  "presenze.dayClosedTitle": "Working day closed",
  "presenze.dayClosedMsg":
    "The working day has already been closed. For corrections, contact your manager.",
  "presenze.totalBreak": "Total break",
  "presenze.inProgress": "In progress",
  "presenze.workedHours": "Worked hours",
  "presenze.overtime": "Overtime",
  "presenze.dayClosed": "Day closed",
  "presenze.tapToRecord": "Tap to record",
  "presenze.notAvailable": "Not available now",
  "presenze.todayEntries": "Today's clock events",
  "presenze.quickAttendance": "Attendance",
  "presenze.quickMyEntries": "My clock events",
  "presenze.quickRequests": "Requests",
  "presenze.quickRequestsDesc": "Leave, permits and more",
  "presenze.quickHistory": "History",
  "presenze.quickProfile": "Profile",
  "presenze.dailyQuotaTitle": "Daily hours reached",
  "presenze.dailyQuotaMsg": "You have reached your expected hours for today.",
  "presenze.entryRecorded": "Clock event recorded:",
  "presenze.entryNotSaved": "Clock event not saved",
  "presenze.notAllowedNow": "This clock event is not allowed right now.",
} as const;

type DictKey = keyof typeof en;

const it: Record<DictKey, string> = {
  "module.dashboard": "Dashboard",
  "module.presenze": "Presenze",
  "module.richieste": "Richieste",
  "module.procurement": "Procurement",
  "module.documenti": "Documenti",
  "module.comunicazioni": "Comunicazioni",
  "module.gestione-timbrature": "Gestione timbrature",
  "module.supervisione": "Supervisione",
  "module.report": "Rendiconto",
  "module.amministrazione": "Amministrazione",
  "module.novita": "Novità",
  "nav.modules": "Moduli",
  "nav.comingSoon": "In arrivo",
  "nav.notActive": "Non attivo",
  "shell.logout": "Esci",
  "shell.logoutConfirm": "Sei sicuro di voler uscire?",
  "shell.role.dipendente": "Dipendente",
  "shell.role.responsabile": "Responsabile",
  "shell.role.amministratore_sistema": "Amministratore di sistema",
  "login.title": "Accedi",
  "login.subtitle": "Inserisci il tuo codice dipendente e il PIN aziendale.",
  "login.codePlaceholder": "Codice dipendente",
  "login.pinPlaceholder": "PIN",
  "login.submit": "Accedi",
  "login.checking": "Verifica in corso…",
  "login.invalid": "Codice o PIN non validi.",
  "login.welcome": "Benvenuto",
  "login.msSoon": "L'autenticazione Microsoft 365 sarà attivata prossimamente.",
  "login.hero":
    "Una piattaforma modulare che unisce Presenze, Richieste, Report e Amministrazione in un'unica esperienza integrata.",
  "common.loading": "Caricamento…",
  "common.all": "Tutti",
  "common.allF": "Tutte",
  "common.send": "Invia",
  "common.save": "Salva",
  "common.cancel": "Annulla",
  "common.approve": "Approva",
  "common.reject": "Respingi",
  "common.confirmReject": "Conferma rifiuto",
  "common.rejectReason": "Motivo del rifiuto",
  "common.required": "(obbligatorio)",
  "common.optional": "(opzionale)",
  "common.exportCsv": "Esporta CSV",
  "common.open": "apri",
  "common.select": "— seleziona —",
  "common.site": "Sede",
  "common.allSites": "Tutte le sedi",
  "common.employee": "Dipendente",
  "common.status": "Stato",
  "common.date": "Data",
  "common.from": "Dal",
  "common.to": "Al",
  "common.type": "Tipo",
  "common.amount": "Importo",
  "common.total": "Totale",
  "common.attachment": "allegato",
  "common.doc": "Doc.",
  "common.error": "Errore",
  "common.restricted": "Accesso riservato",
  "stato.Bozza": "Bozza",
  "stato.Inviata": "Inviata",
  "stato.Comunicata": "Comunicata",
  "stato.Approvata": "Approvata",
  "stato.Respinta": "Respinta",
  "stato.Annullata": "Annullata",
  "stato.Da inviare": "Da inviare",
  "evento.entrata": "Entrata",
  "evento.inizio-pausa": "Inizio pausa",
  "evento.fine-pausa": "Fine pausa",
  "evento.uscita": "Uscita",
  "dstato.presente": "Presente",
  "dstato.pausa": "In pausa",
  "dstato.assente": "Assente",
  "dstato.oltre": "Oltre orario",
  "presenze.title": "Le mie presenze",
  "presenze.statusLabel": "Stato:",
  "presenze.hoursToday": "Ore oggi:",
  "presenze.lastEntry": "Ultima timbratura",
  "presenze.noneToday": "Nessuna timbratura registrata oggi.",
  "presenze.dayClosedTitle": "Giornata lavorativa chiusa",
  "presenze.dayClosedMsg":
    "La giornata lavorativa è già stata chiusa. Per eventuali correzioni contatta il tuo responsabile.",
  "presenze.totalBreak": "Pausa totale",
  "presenze.inProgress": "In corso",
  "presenze.workedHours": "Ore lavorate",
  "presenze.overtime": "Oltre orario",
  "presenze.dayClosed": "Giornata chiusa",
  "presenze.tapToRecord": "Tocca per registrare",
  "presenze.notAvailable": "Non disponibile ora",
  "presenze.todayEntries": "Timbrature di oggi",
  "presenze.quickAttendance": "Presenze",
  "presenze.quickMyEntries": "Le mie timbrature",
  "presenze.quickRequests": "Richieste",
  "presenze.quickRequestsDesc": "Ferie, permessi e altro",
  "presenze.quickHistory": "Storico",
  "presenze.quickProfile": "Profilo",
  "presenze.dailyQuotaTitle": "Monte ore giornaliero raggiunto",
  "presenze.dailyQuotaMsg": "Hai raggiunto le ore previste per oggi.",
  "presenze.entryRecorded": "Timbratura registrata:",
  "presenze.entryNotSaved": "Timbratura non salvata",
  "presenze.notAllowedNow": "Timbratura non consentita in questo momento.",
};

const dictionaries: Record<Lang, Record<DictKey, string>> = { en, it };

interface LangContextValue {
  lang: Lang;
  setLang: (l: Lang) => void;
  t: (key: DictKey) => string;
  /** Traduce un valore di stato applicativo (es. "Approvata"); se sconosciuto
   *  restituisce il valore com'è. */
  tStato: (stato: string) => string;
  /** Titolo di un modulo del menu per id, con fallback al titolo registrato. */
  tModule: (id: string, fallback: string) => string;
}

const LangContext = createContext<LangContextValue>({
  lang: DEFAULT_LANG,
  setLang: () => {},
  t: (k) => en[k] ?? k,
  tStato: (s) => s,
  tModule: (_id, fallback) => fallback,
});

export function LanguageProvider({ children }: { children: ReactNode }) {
  // Parte SEMPRE dal default (anche lato server) per evitare mismatch di
  // hydration; la preferenza salvata viene applicata dopo il mount.
  const [lang, setLangState] = useState<Lang>(DEFAULT_LANG);
  useEffect(() => {
    try {
      const saved = window.localStorage.getItem(STORAGE_KEY);
      if (saved === "it" || saved === "en") setLangState(saved);
    } catch {
      /* ignore */
    }
  }, []);
  const setLang = (l: Lang) => {
    setLangState(l);
    try {
      window.localStorage.setItem(STORAGE_KEY, l);
    } catch {
      /* ignore */
    }
  };
  const t = (key: DictKey): string => dictionaries[lang][key] ?? en[key] ?? key;
  const tStato = (stato: string): string => {
    const key = `stato.${stato}` as DictKey;
    return dictionaries[lang][key] ?? stato;
  };
  const tModule = (id: string, fallback: string): string => {
    const key = `module.${id}` as DictKey;
    return dictionaries[lang][key] ?? fallback;
  };
  return (
    <LangContext.Provider value={{ lang, setLang, t, tStato, tModule }}>
      {children}
    </LangContext.Provider>
  );
}

export function useLang(): LangContextValue {
  return useContext(LangContext);
}

// Selettore lingua (bandierine). Compatto, per header e pagina di login.
export function LangSwitcher({ className = "" }: { className?: string }) {
  const { lang, setLang } = useLang();
  const btn = (l: Lang, flag: string, label: string) => (
    <button
      type="button"
      onClick={() => setLang(l)}
      aria-label={label}
      title={label}
      className={`px-1.5 h-7 rounded-md text-base leading-none transition-all ${
        lang === l ? "bg-secondary shadow-sm scale-110" : "opacity-45 hover:opacity-90"
      }`}
    >
      {flag}
    </button>
  );
  return (
    <div className={`inline-flex items-center gap-0.5 ${className}`}>
      {btn("en", "🇬🇧", "English")}
      {btn("it", "🇮🇹", "Italiano")}
    </div>
  );
}
