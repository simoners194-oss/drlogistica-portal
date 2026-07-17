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
  // --- Tipi richiesta / modalità (valore dato → etichetta) ---
  "tipoR.Ferie": "Leave",
  "tipoR.Permesso": "Permit",
  "tipoR.Straordinario": "Overtime",
  "tipoR.Smart Working": "Smart working",
  "tipoR.Malattia": "Sick leave",
  "tipoR.Reperibilità": "On-call",
  "tipoR.Rimborso spese": "Expense reimbursement",
  "mod.Preventivo": "Planned",
  "mod.Consuntivo": "Actual",
  // --- Richieste ---
  "rich.title": "Requests",
  "rich.subtitle": "Leave, permits and reimbursements",
  "rich.myRequests": "My requests",
  "rich.toApprove": "To approve",
  "rich.queueTitle": "Requests awaiting approval",
  "rich.loadError": "Load error:",
  "rich.nonePending": "No requests awaiting approval.",
  "rich.ferieResidue": "Remaining leave:",
  "rich.permessiResidui": "Remaining permit hours:",
  "rich.newRequest": "New request",
  "rich.commNote": "This is a notification: it does not require approval and is recorded directly.",
  "rich.purchaseDate": "Purchase date",
  "rich.amountEur": "Amount (€)",
  "rich.typology": "Category",
  "rich.detail": "Detail",
  "rich.receiptDoc": "Receipt document",
  "rich.fileOptional": "(file, optional)",
  "rich.receiptHint":
    "Photo or PDF of the receipt/invoice · max 8 MB. Alternatively paste a link below.",
  "rich.receiptLinkPh": "Or a link to the document (optional)",
  "rich.day": "Day",
  "rich.fromTime": "From",
  "rich.toTime": "To",
  "rich.mode": "Mode",
  "rich.inpsProto": "INPS protocol",
  "rich.inpsPh": "Certificate protocol number",
  "rich.reason": "Reason",
  "rich.sending": "Sending…",
  "rich.notify": "Notify",
  "rich.submit": "Submit request",
  "rich.noneYet": "No requests yet. Create one from the panel on the left.",
  "rich.reasonPrefix": "Reason:",
  "rich.cancelledToast": "Request cancelled",
  "rich.cancelFailed": "Cancellation failed",
  "rich.needRejectNote": "A note is required to reject the request",
  "rich.approvedToast": "Request approved",
  "rich.rejectedToast": "Request rejected",
  "rich.opFailed": "Operation failed",
  "rich.notifSent": "Notification sent",
  "rich.reqSent": "Request submitted",
  "rich.checkFields": "Please check the request fields",
  "rich.selectDetail": "Please also select the category detail",
  "rich.fileTooBig": "File too large: the limit is 8 MB.",
  // --- Dashboard ---
  "dash.titleResp": "Managers dashboard",
  "dash.titlePres": "Attendance dashboard",
  "dash.subResp": "Live overview · read-only · all sites",
  "dash.subPres": "Live monitoring of DR Logistica sites",
  "dash.sysDown": "System temporarily unavailable",
  "dash.sysDownMsg":
    "Unable to read data from SharePoint. Clock-ins are temporarily disabled. Technical details in Administration.",
  "dash.noTimbratura": "Clock-ins not active",
  "dash.noTimbraturaMsg":
    "No site has clock-ins enabled: the attendance overview will appear automatically once a site enables them.",
  "dash.summaryTitle": "Attendance summary by site",
  "dash.updatedAt": "Updated at",
  "dash.refreshNow": "Refresh now",
  "dash.refresh": "Refresh",
  "dash.sitePrefix": "Site",
  "dash.totalEmployees": "total employees",
  "dash.kpiPresent": "Present",
  "dash.kpiBreak": "On break",
  "dash.kpiOut": "Clocked out",
  "dash.kpiAbsent": "Not clocked in",
  "dash.kpiOvertime": "On overtime",
  "dash.inSede": "On site",
  "dash.noneInState": "No employees in this state",
  "dash.footerPresent": "Present",
  "dash.footerTotal": "Total",
  "dash.alertDelays": "Late arrivals",
  "dash.expected": "Expected",
  "dash.expectedLower": "expected",
  "dash.noItems": "Nothing to report",
  "dash.qaEmployees": "Employees",
  "dash.qaDiag": "Diagnostics",
  "dash.qaDiagDesc": "Health & self-test",
  "dash.qaNoClockin": "Site without clock-ins",
  // --- Dettaglio dipendente (dialog) ---
  "dlg.subtitle": "Today's attendance detail — read-only.",
  "dlg.dayClosed": "Working day closed.",
  // --- Supervisione ---
  "sup.title": "Supervision",
  "sup.subtitle": "Approvals and manual clock entries",
  "sup.tabDecise": "Decided requests",
  "sup.tabManuali": "Manual entries",
  "sup.tabRimborsi": "Reimbursements",
  "sup.deciseTitle": "Decided requests (approved and rejected)",
  "sup.deciseNone": "No decided requests with these filters.",
  "sup.requests": "requests",
  "sup.request": "request",
  "sup.colRichiesta": "Request",
  "sup.colPeriodo": "Period",
  "sup.approvate": "Approved",
  "sup.rifiutate": "Rejected",
  "sup.rimborsiTitle": "Expense reimbursements (decided)",
  "sup.rimborsiNone": "No reimbursements with these filters.",
  "sup.colTipologia": "Category",
  "sup.totApproved": "Total approved",
  "sup.manualiTitle": "Manual clock entries (last 30 days)",
  "sup.manualiDesc": "Entries added by the back office (Manual origin), for review.",
  "sup.manualiNone": "No manual entries in the period.",
  "sup.restrictedMsg": "This section is reserved for supervisors.",
  "sup.errRequests": "Requests error",
  "sup.errEntries": "Clock entries error",
  // --- Rendiconto ---
  "rep.title": "Summary report",
  "rep.subtitle": "Monthly hours and leave balance",
  "rep.tabRendiconto": "Report",
  "rep.tabFerie": "Leave balance",
  "rep.monthlyTitle": "Monthly report",
  "rep.period": "Period",
  "rep.periodMonth": "Month",
  "rep.periodFiscal": "Fiscal week (year)",
  "rep.periodMonthWeek": "Week of the month",
  "rep.yearFromMonth": "Year (from month)",
  "rep.month": "Month",
  "rep.fiscalWeekN": "Fiscal week (1-53)",
  "rep.monthWeekN": "Week of month (1-6)",
  "rep.fromTo": "from {from} to {to}",
  "rep.calculating": "Calculating…",
  "rep.noData": "No data for the selected period/filters.",
  "rep.colWorked": "Worked hrs",
  "rep.colOtCalc": "OT calc.",
  "rep.colOtAuth": "OT auth.",
  "rep.colPermits": "Permits",
  "rep.colLeave": "Leave",
  "rep.colSick": "Sick",
  "rep.footnote":
    "OT calc. = overtime from clock entries (Mon–Sat hours beyond the weekly quota + all of Sunday). OT auth. = hours from approved overtime requests, for comparison. The weekly expected time is reduced by leave/sick days (WeeklyHours/5 per day) and permits. Weeks spanning two months count in the month of their Monday. ⚠️ = days with an unclosed shift (hours not counted: fix them in Anomalies).",
  "rep.openShiftWarn": "day(s) with an unclosed shift: fix in Anomalies",
  "rep.ferieTitle": "Leave balance",
  "rep.year": "Year",
  "rep.colEntitled": "Entitled",
  "rep.colTaken": "Taken",
  "rep.colRemaining": "Remaining",
  "rep.noDataFilters": "No data for the selected filters.",
  "rep.ferieFootnote":
    "Entitled from the employee's AnnualLeaveDays column (default 26 if not set). Taken = approved Leave days in the year.",
  "rep.errReport": "Report error",
  "rep.errBalance": "Leave balance error",
  "rep.restrictedMsg": "The report is reserved for back office, supervisors and administrators.",
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
  "tipoR.Ferie": "Ferie",
  "tipoR.Permesso": "Permesso",
  "tipoR.Straordinario": "Straordinario",
  "tipoR.Smart Working": "Smart Working",
  "tipoR.Malattia": "Malattia",
  "tipoR.Reperibilità": "Reperibilità",
  "tipoR.Rimborso spese": "Rimborso spese",
  "mod.Preventivo": "Preventivo",
  "mod.Consuntivo": "Consuntivo",
  "rich.title": "Richieste",
  "rich.subtitle": "Ferie, permessi e giustificativi",
  "rich.myRequests": "Le mie richieste",
  "rich.toApprove": "Da approvare",
  "rich.queueTitle": "Richieste da approvare",
  "rich.loadError": "Errore nel caricamento:",
  "rich.nonePending": "Nessuna richiesta in attesa di approvazione.",
  "rich.ferieResidue": "Ferie residue:",
  "rich.permessiResidui": "Permessi residui:",
  "rich.newRequest": "Nuova richiesta",
  "rich.commNote":
    "Questa è una comunicazione: non richiede approvazione, viene registrata direttamente.",
  "rich.purchaseDate": "Data acquisto",
  "rich.amountEur": "Importo (€)",
  "rich.typology": "Tipologia",
  "rich.detail": "Dettaglio",
  "rich.receiptDoc": "Documento giustificativo",
  "rich.fileOptional": "(file, opzionale)",
  "rich.receiptHint":
    "Foto o PDF dello scontrino/fattura · max 8 MB. In alternativa incolla un link qui sotto.",
  "rich.receiptLinkPh": "Oppure link al documento (opzionale)",
  "rich.day": "Giorno",
  "rich.fromTime": "Dalle",
  "rich.toTime": "Alle",
  "rich.mode": "Modalità",
  "rich.inpsProto": "Protocollo INPS",
  "rich.inpsPh": "Numero di protocollo del certificato",
  "rich.reason": "Motivazione",
  "rich.sending": "Invio in corso…",
  "rich.notify": "Comunica",
  "rich.submit": "Invia richiesta",
  "rich.noneYet": "Nessuna richiesta ancora. Creane una dal riquadro a sinistra.",
  "rich.reasonPrefix": "Motivo:",
  "rich.cancelledToast": "Richiesta annullata",
  "rich.cancelFailed": "Annullamento non riuscito",
  "rich.needRejectNote": "Serve una nota per respingere la richiesta",
  "rich.approvedToast": "Richiesta approvata",
  "rich.rejectedToast": "Richiesta respinta",
  "rich.opFailed": "Operazione non riuscita",
  "rich.notifSent": "Comunicazione inviata",
  "rich.reqSent": "Richiesta inviata",
  "rich.checkFields": "Controlla i campi della richiesta",
  "rich.selectDetail": "Seleziona anche il dettaglio della tipologia",
  "rich.fileTooBig": "File troppo grande: il limite è 8 MB.",
  "dash.titleResp": "Dashboard responsabili",
  "dash.titlePres": "Dashboard presenze",
  "dash.subResp": "Panoramica live · sola lettura · tutte le sedi",
  "dash.subPres": "Monitoraggio live sedi DR Logistica",
  "dash.sysDown": "Sistema momentaneamente non disponibile",
  "dash.sysDownMsg":
    "Impossibile leggere i dati da SharePoint. Le timbrature sono temporaneamente disabilitate. Dettagli tecnici in Amministrazione.",
  "dash.noTimbratura": "Timbrature non attive",
  "dash.noTimbraturaMsg":
    "Nessuna sede ha la timbratura attiva: la panoramica presenze comparirà automaticamente quando una sede abiliterà la timbratura.",
  "dash.summaryTitle": "Sintesi presenze per sede",
  "dash.updatedAt": "Aggiornato alle",
  "dash.refreshNow": "Aggiorna ora",
  "dash.refresh": "Aggiorna",
  "dash.sitePrefix": "Sede",
  "dash.totalEmployees": "dipendenti totali",
  "dash.kpiPresent": "Presenti",
  "dash.kpiBreak": "In pausa",
  "dash.kpiOut": "Usciti",
  "dash.kpiAbsent": "Non timbrati",
  "dash.kpiOvertime": "In straordinario",
  "dash.inSede": "In sede",
  "dash.noneInState": "Nessun dipendente in questo stato",
  "dash.footerPresent": "Presenti",
  "dash.footerTotal": "Totale",
  "dash.alertDelays": "Ritardi",
  "dash.expected": "Atteso",
  "dash.expectedLower": "orario",
  "dash.noItems": "Nessun elemento",
  "dash.qaEmployees": "Dipendenti",
  "dash.qaDiag": "Diagnostica",
  "dash.qaDiagDesc": "Health & self-test",
  "dash.qaNoClockin": "Sede senza timbratura",
  "dlg.subtitle": "Dettaglio presenze di oggi — sola lettura.",
  "dlg.dayClosed": "Giornata lavorativa chiusa.",
  "sup.title": "Supervisione",
  "sup.subtitle": "Approvazioni e timbrature manuali",
  "sup.tabDecise": "Richieste decise",
  "sup.tabManuali": "Timbrature manuali",
  "sup.tabRimborsi": "Rimborsi",
  "sup.deciseTitle": "Richieste decise (approvate e rifiutate)",
  "sup.deciseNone": "Nessuna richiesta decisa con questi filtri.",
  "sup.requests": "richieste",
  "sup.request": "richiesta",
  "sup.colRichiesta": "Richiesta",
  "sup.colPeriodo": "Periodo",
  "sup.approvate": "Approvate",
  "sup.rifiutate": "Rifiutate",
  "sup.rimborsiTitle": "Rimborsi spese (decisi)",
  "sup.rimborsiNone": "Nessun rimborso con questi filtri.",
  "sup.colTipologia": "Tipologia",
  "sup.totApproved": "Totale approvati",
  "sup.manualiTitle": "Timbrature manuali (ultimi 30 giorni)",
  "sup.manualiDesc": "Inserimenti fatti dall'operatore (origine Manuale), per visione.",
  "sup.manualiNone": "Nessuna timbratura manuale nel periodo.",
  "sup.restrictedMsg": "Questa sezione è riservata ai supervisori.",
  "sup.errRequests": "Errore richieste",
  "sup.errEntries": "Errore timbrature",
  "rep.title": "Rendiconto",
  "rep.subtitle": "Ore mensili e saldo ferie",
  "rep.tabRendiconto": "Rendiconto",
  "rep.tabFerie": "Saldo ferie",
  "rep.monthlyTitle": "Rendiconto mensile",
  "rep.period": "Periodo",
  "rep.periodMonth": "Mese",
  "rep.periodFiscal": "Settimana fiscale (anno)",
  "rep.periodMonthWeek": "Settimana del mese",
  "rep.yearFromMonth": "Anno (dal mese)",
  "rep.month": "Mese",
  "rep.fiscalWeekN": "Week fiscale (1-53)",
  "rep.monthWeekN": "Week del mese (1-6)",
  "rep.fromTo": "dal {from} al {to}",
  "rep.calculating": "Calcolo in corso…",
  "rep.noData": "Nessun dato per il periodo/filtri selezionati.",
  "rep.colWorked": "Ore lav.",
  "rep.colOtCalc": "Str. calc.",
  "rep.colOtAuth": "Str. autor.",
  "rep.colPermits": "Permessi",
  "rep.colLeave": "Ferie",
  "rep.colSick": "Malattie",
  "rep.footnote":
    "Str. calc. = straordinario dalle timbrature (ore Lun–Sab oltre il monte ore settimanale + tutta la domenica). Str. autor. = ore da richieste di straordinario approvate, per confronto. Il previsto settimanale è ridotto da ferie/malattia (OreSettimanali/5 per giorno) e dai permessi. Le settimane a cavallo di due mesi contano nel mese del loro lunedì. ⚠️ = giorni con turno non chiuso (ore non conteggiate: correggere nelle Anomalie).",
  "rep.openShiftWarn": "giorno/i con turno non chiuso: correggere nelle Anomalie",
  "rep.ferieTitle": "Saldo ferie",
  "rep.year": "Anno",
  "rep.colEntitled": "Spettanti",
  "rep.colTaken": "Godute",
  "rep.colRemaining": "Residue",
  "rep.noDataFilters": "Nessun dato per i filtri selezionati.",
  "rep.ferieFootnote":
    "Spettanti dalla colonna GiorniFerieAnnui del dipendente (default 26 se non impostata). Godute = giorni di Ferie approvate nell'anno.",
  "rep.errReport": "Errore rendiconto",
  "rep.errBalance": "Errore saldo ferie",
  "rep.restrictedMsg": "Il rendiconto è riservato a operatori, supervisori e amministratori.",
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
  /** Traduzione generica di un VALORE dato con prefisso (es. tVal("tipoR",
   *  "Ferie") → "Leave"); valore sconosciuto → restituito com'è. */
  tVal: (prefix: string, value: string) => string;
}

const LangContext = createContext<LangContextValue>({
  lang: DEFAULT_LANG,
  setLang: () => {},
  t: (k) => en[k] ?? k,
  tStato: (s) => s,
  tModule: (_id, fallback) => fallback,
  tVal: (_p, v) => v,
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
  const tVal = (prefix: string, value: string): string => {
    const key = `${prefix}.${value}` as DictKey;
    return dictionaries[lang][key] ?? value;
  };
  return (
    <LangContext.Provider value={{ lang, setLang, t, tStato, tModule, tVal }}>
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
