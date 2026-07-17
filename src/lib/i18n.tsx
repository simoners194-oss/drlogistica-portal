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
