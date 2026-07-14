export type SedeId = "roma" | "san-giuliano";

export type StatoTimbratura = "presente" | "pausa" | "uscito" | "non-timbrato";

export interface Timbratura {
  tipo: "entrata" | "inizio-pausa" | "fine-pausa" | "uscita";
  ora: string; // ISO
}

export interface Dipendente {
  id: string;
  nome: string;
  cognome: string;
  ruolo: string;
  sede: SedeId;
  orarioAtteso: string; // "09:00"
  stato: StatoTimbratura;
  ultimaTimbratura?: Timbratura;
  entrataOra?: string; // ISO — timbratura di entrata effettiva (per calcolo ore)
  ritardoMinuti?: number;
  straordinariMinuti?: number;
  // Eventi di oggi ordinati per orario crescente. Usati dal Modulo Presenze
  // per la timeline e dalla macchina a stati.
  eventiOggi?: Timbratura[];
  oreLavorateMinuti?: number;
  pausaMinuti?: number;
  oltreOrarioMinuti?: number;
}

export const SEDI: { id: SedeId; nome: string; timbratura: boolean }[] = [
  { id: "roma", nome: "Fiano Romano", timbratura: false },
  { id: "san-giuliano", nome: "San Giuliano", timbratura: false },
];

// Se la sede NON timbra, i suoi dipendenti fanno solo richieste (niente
// modulo Presenze). Le sedi non elencate/`tutte` sono considerate timbranti.
export function sedeTimbra(sede: SedeId | "tutte"): boolean {
  if (sede === "tutte") return true;
  return SEDI.find((s) => s.id === sede)?.timbratura ?? true;
}

// Almeno una sede timbra? (usato per grigiare le viste presenze quando nessuna
// sede è timbrante — es. oggi Fiano Romano e San Giuliano non timbrano).
export function anySedeTimbra(): boolean {
  return SEDI.some((s) => s.timbratura);
}

const oggi = (h: number, m: number) => {
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d.toISOString();
};

export const DIPENDENTI: Dipendente[] = [
  {
    id: "1",
    nome: "Marco",
    cognome: "Rossi",
    ruolo: "Magazziniere",
    sede: "roma",
    orarioAtteso: "08:00",
    stato: "presente",
    entrataOra: oggi(8, 2),
    ultimaTimbratura: { tipo: "entrata", ora: oggi(8, 2) },
    ritardoMinuti: 2,
  },
  {
    id: "2",
    nome: "Giulia",
    cognome: "Bianchi",
    ruolo: "Team Leader",
    sede: "roma",
    orarioAtteso: "08:30",
    stato: "pausa",
    entrataOra: oggi(8, 25),
    ultimaTimbratura: { tipo: "inizio-pausa", ora: oggi(12, 30) },
  },
  {
    id: "3",
    nome: "Luca",
    cognome: "Verdi",
    ruolo: "Autista",
    sede: "roma",
    orarioAtteso: "07:00",
    stato: "uscito",
    entrataOra: oggi(7, 0),
    ultimaTimbratura: { tipo: "uscita", ora: oggi(16, 15) },
    straordinariMinuti: 45,
  },
  {
    id: "4",
    nome: "Sara",
    cognome: "Neri",
    ruolo: "Impiegata",
    sede: "roma",
    orarioAtteso: "09:00",
    stato: "non-timbrato",
  },
  {
    id: "5",
    nome: "Andrea",
    cognome: "Ferrari",
    ruolo: "Magazziniere",
    sede: "roma",
    orarioAtteso: "08:00",
    stato: "presente",
    entrataOra: oggi(7, 58),
    ultimaTimbratura: { tipo: "entrata", ora: oggi(7, 58) },
  },
  {
    id: "6",
    nome: "Elena",
    cognome: "Russo",
    ruolo: "Responsabile",
    sede: "roma",
    orarioAtteso: "09:00",
    stato: "presente",
    entrataOra: oggi(9, 15),
    ultimaTimbratura: { tipo: "entrata", ora: oggi(9, 15) },
    ritardoMinuti: 15,
    straordinariMinuti: 20,
  },

  {
    id: "7",
    nome: "Paolo",
    cognome: "Conti",
    ruolo: "Magazziniere",
    sede: "san-giuliano",
    orarioAtteso: "07:00",
    stato: "presente",
    entrataOra: oggi(7, 0),
    ultimaTimbratura: { tipo: "entrata", ora: oggi(7, 0) },
    straordinariMinuti: 30,
  },
  {
    id: "8",
    nome: "Chiara",
    cognome: "Marino",
    ruolo: "Team Leader",
    sede: "san-giuliano",
    orarioAtteso: "08:00",
    stato: "pausa",
    entrataOra: oggi(8, 0),
    ultimaTimbratura: { tipo: "inizio-pausa", ora: oggi(13, 0) },
  },
  {
    id: "9",
    nome: "Davide",
    cognome: "Greco",
    ruolo: "Autista",
    sede: "san-giuliano",
    orarioAtteso: "06:30",
    stato: "presente",
    entrataOra: oggi(6, 30),
    ultimaTimbratura: { tipo: "fine-pausa", ora: oggi(13, 30) },
  },
  {
    id: "10",
    nome: "Francesca",
    cognome: "Ricci",
    ruolo: "Impiegata",
    sede: "san-giuliano",
    orarioAtteso: "09:00",
    stato: "non-timbrato",
  },
  {
    id: "11",
    nome: "Matteo",
    cognome: "Costa",
    ruolo: "Magazziniere",
    sede: "san-giuliano",
    orarioAtteso: "08:00",
    stato: "non-timbrato",
  },
  {
    id: "12",
    nome: "Alessia",
    cognome: "Bruno",
    ruolo: "Responsabile",
    sede: "san-giuliano",
    orarioAtteso: "08:30",
    stato: "uscito",
    entrataOra: oggi(8, 30),
    ultimaTimbratura: { tipo: "uscita", ora: oggi(17, 45) },
    straordinariMinuti: 75,
  },
];

export const STATO_LABEL: Record<StatoTimbratura, string> = {
  presente: "Presente",
  pausa: "In pausa",
  uscito: "Uscito",
  "non-timbrato": "Non timbrato",
};

export const STATO_COLOR: Record<StatoTimbratura, string> = {
  presente: "bg-status-present",
  pausa: "bg-status-break",
  uscito: "bg-status-out",
  "non-timbrato": "bg-status-absent",
};

export function formatOra(iso?: string) {
  if (!iso) return "—";
  return new Date(iso).toLocaleTimeString("it-IT", { hour: "2-digit", minute: "2-digit" });
}

export function labelTipo(t: Timbratura["tipo"]) {
  return {
    entrata: "Entrata",
    "inizio-pausa": "Inizio pausa",
    "fine-pausa": "Fine pausa",
    uscita: "Uscita",
  }[t];
}
