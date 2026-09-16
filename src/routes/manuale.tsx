// DR Portal — Manuale interno ricercabile (richiesta Simone 16/09).
//
// Chi ha un dubbio entra qui, scrive due parole e trova COSA DEVE FARE.
// Il contenuto vive in src/lib/manuale-content.ts: ogni modifica al portale
// va accompagnata dalla voce di manuale corrispondente (regola fissa).
// Per ora è abilitato a chi ha l'abilitazione Finanze (admin + vista
// direzione); quando verrà aperto a tutti basterà togliere il gate.

import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { BookOpen, ChevronDown, ChevronRight, Search } from "lucide-react";
import { AppShell } from "@/components/AppShell";
import { readSession, type SessionUser } from "@/lib/session";
import { haVistaDirezione } from "@/lib/richieste-logic";
import { MANUALE, SEZIONI_MANUALE, type VoceManuale } from "@/lib/manuale-content";
import { useLang } from "@/lib/i18n";

export const Route = createFileRoute("/manuale")({
  head: () => ({
    meta: [
      { title: "Manuale — DR Portal" },
      {
        name: "description",
        content: "Manuale interno di DR Portal: cerca un dubbio e trovi cosa fare.",
      },
    ],
  }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    const s = readSession();
    if (!s) throw redirect({ to: "/", search: { redirect: location.href } });
    const ok = s.ruolo === "amministratore_sistema" || haVistaDirezione(s.codice ?? "");
    if (!ok) throw redirect({ to: "/dashboard" });
  },
  component: ManualePage,
});

// Ricerca senza accenti né maiuscole: ogni parola digitata deve comparire
// da qualche parte nella voce (titolo, testo, sezione o parole chiave).
function norm(s: string): string {
  return s
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\u2018\u2019\u201b`\u00b4]/g, "'")
    .replace(/[\u201c\u201d]/g, '"');
}
function corrisponde(v: VoceManuale, parole: string[]): boolean {
  const testo = norm(`${v.titolo} ${v.testo} ${v.sezione} ${v.chiavi ?? ""}`);
  return parole.every((p) => testo.includes(p));
}

function ManualePage() {
  const { t } = useLang();
  const [session, setSession] = useState<SessionUser | null>(null);
  const [q, setQ] = useState("");
  const [sezione, setSezione] = useState("");
  const [aperte, setAperte] = useState<Set<string>>(new Set());
  useEffect(() => {
    setSession(readSession());
  }, []);

  const parole = useMemo(() => norm(q).split(/\s+/).filter(Boolean), [q]);
  const risultati = useMemo(() => {
    return MANUALE.filter(
      (v) => (!sezione || v.sezione === sezione) && (!parole.length || corrisponde(v, parole)),
    );
  }, [parole, sezione]);
  const perSezione = useMemo(() => {
    const m = new Map<string, VoceManuale[]>();
    for (const s of SEZIONI_MANUALE) m.set(s, []);
    for (const v of risultati) {
      if (!m.has(v.sezione)) m.set(v.sezione, []);
      m.get(v.sezione)!.push(v);
    }
    return [...m.entries()].filter(([, voci]) => voci.length > 0);
  }, [risultati]);
  // Con una ricerca attiva le voci trovate si aprono da sole.
  const cerca = parole.length > 0;
  const aperta = (id: string) => cerca || aperte.has(id);
  // Con la ricerca attiva le voci sono già tutte aperte: il click sul titolo
  // non deve toccare lo stato (altrimenti, svuotata la ricerca, si trovano
  // voci aperte o chiuse "da sole").
  const toggle = (id: string) => {
    if (cerca) return;
    setAperte((set) => {
      const ns = new Set(set);
      if (ns.has(id)) ns.delete(id);
      else ns.add(id);
      return ns;
    });
  };

  if (!session) return null;

  return (
    <AppShell title={t("man.title")} subtitle={t("man.subtitle")}>
      <div className="mx-auto max-w-3xl">
        <div className="relative mb-3">
          <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
          <input
            autoFocus
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder={t("man.cerca")}
            className="w-full rounded-2xl border border-border bg-card py-3 pl-10 pr-4 text-sm shadow-[var(--shadow-card)] focus:outline-none focus:ring-2 focus:ring-primary/40"
          />
        </div>
        <div className="mb-4 flex flex-wrap gap-1.5">
          <button
            type="button"
            onClick={() => setSezione("")}
            className={`rounded-full border px-2.5 py-1 text-[11px] ${!sezione ? "border-primary bg-primary/10 text-primary" : "border-border text-muted-foreground hover:bg-muted"}`}
          >
            {t("man.tutte")}
          </button>
          {SEZIONI_MANUALE.map((s) => (
            <button
              key={s}
              type="button"
              onClick={() => setSezione(sezione === s ? "" : s)}
              className={`rounded-full border px-2.5 py-1 text-[11px] ${sezione === s ? "border-primary bg-primary/10 text-primary" : "border-border text-muted-foreground hover:bg-muted"}`}
            >
              {s}
            </button>
          ))}
        </div>

        {perSezione.length === 0 && (
          <div className="rounded-2xl border border-border bg-card p-8 text-center text-sm text-muted-foreground shadow-[var(--shadow-card)]">
            {t("man.nessuna")}
          </div>
        )}

        {perSezione.map(([s, voci]) => (
          <div key={s} className="mb-5">
            <h2 className="mb-1.5 flex items-center gap-2 text-xs font-semibold uppercase tracking-wider text-muted-foreground">
              <BookOpen className="h-3.5 w-3.5" />
              {s}
            </h2>
            <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-[var(--shadow-card)]">
              {voci.map((v, i) => (
                <div key={v.id} className={i > 0 ? "border-t border-border/60" : ""}>
                  <button
                    type="button"
                    onClick={() => toggle(v.id)}
                    className="flex w-full items-center gap-2 px-4 py-2.5 text-left text-sm font-medium hover:bg-muted/40"
                  >
                    {aperta(v.id) ? (
                      <ChevronDown className="h-4 w-4 shrink-0 text-muted-foreground" />
                    ) : (
                      <ChevronRight className="h-4 w-4 shrink-0 text-muted-foreground" />
                    )}
                    {v.titolo}
                  </button>
                  {aperta(v.id) && (
                    <div className="px-4 pb-3 pl-10 text-sm leading-relaxed text-muted-foreground">
                      {v.testo.split(/\n\n+/).map((par, j) => (
                        <p key={j} className={j > 0 ? "mt-2" : ""}>
                          {par}
                        </p>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        ))}

        <p className="mb-6 mt-2 text-[11px] text-muted-foreground">{t("man.nota")}</p>
      </div>
    </AppShell>
  );
}
