// DR Portal — Finanza → Flussi → "Cassa proprietà": server functions.
// Stessa riservatezza di stipendi e fiscale: amministratore di sistema e
// vista direzione soltanto (la proprietà entra come ADM002).
import { createServerFn } from "@tanstack/react-start";
import { readSessionUser, type ServerSessionUser } from "./auth.server";
import { haVistaDirezione } from "./richieste-logic";
import { loadCassaDb, salvaCassaDb } from "./cassa.server";
import type {
  CassaDb,
  PrioritaCassa,
  RigaManualeCassa,
  SecchioCassa,
  StatoRigaCassa,
  VoceAgenda,
} from "./cassa-logic";

async function utenteCassa(): Promise<ServerSessionUser> {
  const me = await readSessionUser();
  if (!me) throw new Error("Sessione assente o scaduta. Effettua di nuovo l'accesso.");
  const ok = me.ruolo === "amministratore_sistema" || haVistaDirezione(me.codice ?? "");
  if (!ok) throw new Error("Sezione riservata alla direzione.");
  return me;
}

const SECCHI: readonly SecchioCassa[] = ["ora", "prossimi", "standby"];
const PRIORITA: readonly PrioritaCassa[] = ["l", "m", "h"];
const priorita = (v: unknown): PrioritaCassa | undefined =>
  PRIORITA.includes(v as PrioritaCassa) ? (v as PrioritaCassa) : undefined;
const testo = (v: unknown, max: number) =>
  String(v ?? "")
    .trim()
    .slice(0, max);
const numero = (v: unknown) => {
  const n = Number(v);
  if (!Number.isFinite(n) || Math.abs(n) > 1e10) throw new Error("Importo non valido.");
  return Math.round(n * 100) / 100;
};
const dataIso = (v: unknown) => {
  const s = testo(v, 10);
  if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error("Data non valida (formato 2026-10-31).");
  return s;
};

function validaDb(input: unknown): CassaDb {
  const o = (input ?? {}) as Partial<CassaDb>;
  const stati: Record<string, StatoRigaCassa> = {};
  for (const [k, v] of Object.entries(o.stati ?? {})) {
    if (!k || k.length > 300 || !v || typeof v !== "object") continue;
    const s = v as StatoRigaCassa;
    const st: StatoRigaCassa = {};
    if (s.b && SECCHI.includes(s.b)) st.b = s.b;
    if (s.n) st.n = testo(s.n, 200);
    if (s.c) st.c = true;
    if (s.s) st.s = true;
    if (s.x) st.x = true;
    if (s.i != null) st.i = numero(s.i);
    if (s.d) {
      const d = dataIso(s.d);
      if (d) st.d = d;
    }
    const p = priorita(s.p);
    if (p) st.p = p;
    if (Object.keys(st).length) stati[k] = st;
  }
  const manuali: RigaManualeCassa[] = (Array.isArray(o.manuali) ? o.manuali : [])
    .slice(0, 500)
    .map((m) => ({
      id: testo(m?.id, 40) || `M${Date.now().toString(36)}${Math.floor(Math.random() * 1e4)}`,
      tipo: (m?.tipo === "amex" ? "amex" : "manuale") as "amex" | "manuale",
      descrizione: testo(m?.descrizione, 160),
      importo: numero(m?.importo),
      data: dataIso(m?.data),
      secchio: SECCHI.includes(m?.secchio as SecchioCassa) ? (m!.secchio as SecchioCassa) : "ora",
      nota: testo(m?.nota, 200) || undefined,
      entrata: m?.entrata ? true : undefined,
      p: priorita(m?.p),
      ricorrente: m?.ricorrente ? true : undefined,
    }))
    .filter((m) => m.descrizione && m.importo !== 0);
  const fotografie = (Array.isArray(o.fotografie) ? o.fotografie : []).slice(0, 120);
  const agenda: VoceAgenda[] = (Array.isArray(o.agenda) ? o.agenda : [])
    .slice(0, 300)
    .map((v) => ({
      id: testo(v?.id, 40) || `A${Date.now().toString(36)}${Math.floor(Math.random() * 1e4)}`,
      testo: testo(v?.testo, 300),
      p: priorita(v?.p) ?? "m",
      scad: dataIso(v?.scad) || undefined,
      fatta: v?.fatta ? true : undefined,
      creata: testo(v?.creata, 30) || new Date().toISOString(),
    }))
    .filter((v) => v.testo);
  const chiusure: Record<string, string> = {};
  for (const [k, v] of Object.entries(o.chiusure ?? {})) {
    const m = testo(v, 7);
    if (k && k.length <= 60 && /^\d{4}-\d{2}$/.test(m)) chiusure[testo(k, 60)] = m;
  }
  const giorno = Number(o.giornoLimite);
  return {
    versione: Number(o.versione) || 0,
    giornoLimite: giorno >= 1 && giorno <= 28 ? Math.round(giorno) : 5,
    stati,
    manuali,
    fotografie,
    agenda,
    chiusure: Object.keys(chiusure).length ? chiusure : undefined,
    saldoManuale:
      o.saldoManuale && Number.isFinite(Number(o.saldoManuale.importo))
        ? { importo: numero(o.saldoManuale.importo), data: dataIso(o.saldoManuale.data) }
        : undefined,
  };
}

export const spCassaGet = createServerFn({ method: "GET" }).handler(async (): Promise<CassaDb> => {
  await utenteCassa();
  return loadCassaDb();
});

export const spCassaSalva = createServerFn({ method: "POST" })
  .inputValidator((input: { db: CassaDb }) => ({ db: validaDb(input?.db) }))
  .handler(async ({ data }): Promise<CassaDb> => {
    const me = await utenteCassa();
    return salvaCassaDb(data.db, `${me.nome} ${me.cognome}`.trim() || me.id);
  });
