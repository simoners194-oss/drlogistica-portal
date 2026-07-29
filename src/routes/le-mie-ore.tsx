// DR Portal — "Le mie ore": ogni dipendente vede i propri turni del mese e,
// se un giorno è sbagliato, chiede la correzione a chi gestisce le timbrature
// (flag Operatore). Dati SOLO propri: il server non restituisce altro.
import { createFileRoute, redirect } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { AppShell } from "@/components/AppShell";
import { CalendarDays, Loader2, PenLine, Clock, Hourglass, Send } from "lucide-react";
import { readSession, type SessionUser } from "@/lib/session";
import {
  spGetMieTimbrature,
  spGetCorrezioni,
  spCreateCorrezione,
} from "@/lib/sharepoint.functions";
import type { SpTimbratura, SpCorrezione } from "@/lib/sharepoint.server";
import { orePerGiornoDaTurni } from "@/lib/rendiconto-logic";
import { useLang } from "@/lib/i18n";

export const Route = createFileRoute("/le-mie-ore")({
  head: () => ({ meta: [{ title: "Le mie ore — DR Portal" }] }),
  beforeLoad: ({ location }) => {
    if (typeof window === "undefined") return;
    if (!readSession()) throw redirect({ to: "/", search: { redirect: location.href } });
  },
  component: LeMieOrePage,
});

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function fmtGiorno(iso: string, locale: string): string {
  const d = new Date(`${iso}T00:00:00`);
  return d.toLocaleDateString(locale, { weekday: "short", day: "2-digit", month: "short" });
}
function fmtOre(n: number): string {
  const h = Math.floor(n);
  const m = Math.round((n - h) * 60);
  return `${h}h ${String(m).padStart(2, "0")}m`;
}
const ETICHETTA: Record<string, string> = {
  entrata: "entrata",
  "inizio-pausa": "inizio-pausa",
  "fine-pausa": "fine-pausa",
  uscita: "uscita",
};

function LeMieOrePage() {
  const { t, lang } = useLang();
  const locale = lang === "it" ? "it-IT" : "en-GB";
  const [session, setSession] = useState<SessionUser | null>(null);
  const [mese, setMese] = useState(() => new Date().toISOString().slice(0, 7)); // YYYY-MM
  const [tim, setTim] = useState<SpTimbratura[] | null>(null);
  const [correzioni, setCorrezioni] = useState<SpCorrezione[]>([]);
  // Giorno per cui è aperto il modulo di correzione
  const [giornoCorr, setGiornoCorr] = useState<string | null>(null);
  const [orari, setOrari] = useState("");
  const [motivo, setMotivo] = useState("");
  const [inviando, setInviando] = useState(false);

  const from = `${mese}-01`;
  const to = useMemo(() => {
    const [y, m] = mese.split("-").map(Number);
    return new Date(y, m, 0).toISOString().slice(0, 10);
  }, [mese]);

  const carica = () => {
    setTim(null);
    spGetMieTimbrature({ data: { from, to } })
      .then((l) => setTim(l as SpTimbratura[]))
      .catch((err) => {
        setTim([]);
        toast.error(t("mie.errLoad"), {
          description: err instanceof Error ? err.message : String(err),
        });
      });
    spGetCorrezioni({ data: {} })
      .then((l) => setCorrezioni(l as SpCorrezione[]))
      .catch(() => setCorrezioni([]));
  };

  useEffect(() => {
    const s = readSession();
    if (!s) {
      window.location.href = "/";
      return;
    }
    setSession(s);
    carica();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mese]);

  // Giorni del mese con eventi + ore calcolate (turni notturni attribuiti al
  // giorno di inizio, come nel rendiconto ufficiale).
  const giorni = useMemo(() => {
    const eventi = (tim ?? []).map((e) => ({ evento: e.evento, ora: e.dataOra }));
    const calc = orePerGiornoDaTurni(eventi);
    const perGiorno = new Map<string, SpTimbratura[]>();
    for (const e of tim ?? []) {
      const g = e.dataOra.slice(0, 10);
      const l = perGiorno.get(g) ?? [];
      l.push(e);
      perGiorno.set(g, l);
    }
    const chiavi = new Set<string>([...perGiorno.keys(), ...calc.oreGiorno.keys()]);
    return [...chiavi]
      .filter((g) => g >= from && g <= to)
      .sort((a, b) => b.localeCompare(a))
      .map((g) => ({
        giorno: g,
        eventi: (perGiorno.get(g) ?? []).sort((a, b) => a.dataOra.localeCompare(b.dataOra)),
        ore: calc.oreGiorno.get(g) ?? 0,
        aperto: calc.giorniNonChiusi.has(g),
      }));
  }, [tim, from, to]);

  const totale = useMemo(() => giorni.reduce((s, g) => s + g.ore, 0), [giorni]);

  const apriCorrezione = (giorno: string, eventi: SpTimbratura[]) => {
    setGiornoCorr(giorno);
    // Precompila con gli orari attuali: si correggono, non si riscrivono.
    setOrari(
      eventi.map((e) => `${ETICHETTA[e.evento]} ${e.dataOra.slice(11, 16)}`).join(", ") ||
        "entrata 08:00, uscita 17:00",
    );
    setMotivo("");
  };

  const invia = async () => {
    if (!giornoCorr) return;
    setInviando(true);
    try {
      await spCreateCorrezione({
        data: { giorno: giornoCorr, orariProposti: orari.trim(), motivo: motivo.trim() },
      });
      toast.success(t("mie.inviata"), { description: t("mie.inviataMsg") });
      setGiornoCorr(null);
      carica();
    } catch (err) {
      toast.error(t("mie.errInvio"), {
        description: err instanceof Error ? err.message : String(err),
      });
    } finally {
      setInviando(false);
    }
  };

  const statoCorrezione = (giorno: string) => correzioni.find((c) => c.giorno === giorno);

  if (!session) return null;

  return (
    <AppShell title={t("mie.title")} subtitle={t("mie.subtitle")}>
      <div className="mb-4 flex flex-wrap items-end gap-3">
        <div className="w-48">
          <label className="text-xs text-muted-foreground">{t("mie.mese")}</label>
          <input
            type="month"
            value={mese}
            onChange={(e) => setMese(e.target.value)}
            className={inputCls}
          />
        </div>
        <div className="ml-auto rounded-xl border border-border bg-secondary/40 px-4 py-2 text-right">
          <div className="text-[11px] text-muted-foreground">{t("mie.totale")}</div>
          <div className="text-lg font-semibold tabular-nums text-foreground">{fmtOre(totale)}</div>
        </div>
      </div>

      {tim == null ? (
        <div className="py-10 text-center text-sm text-muted-foreground">
          <Loader2 className="h-5 w-5 animate-spin inline-block" />
        </div>
      ) : giorni.length === 0 ? (
        <p className="py-8 text-center text-sm text-muted-foreground">{t("mie.vuoto")}</p>
      ) : (
        <ul className="space-y-2">
          {giorni.map((g) => {
            const corr = statoCorrezione(g.giorno);
            return (
              <li
                key={g.giorno}
                className="rounded-2xl border border-border bg-card p-4 shadow-[var(--shadow-card)]"
              >
                <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
                  <span className="inline-flex items-center gap-2 text-sm font-medium text-foreground capitalize min-w-36">
                    <CalendarDays className="h-4 w-4 text-muted-foreground" />
                    {fmtGiorno(g.giorno, locale)}
                  </span>
                  <span className="inline-flex items-center gap-1.5 text-sm tabular-nums">
                    <Hourglass className="h-4 w-4 text-muted-foreground" />
                    {g.aperto ? (
                      <span className="text-status-absent">{t("mie.aperto")}</span>
                    ) : (
                      fmtOre(g.ore)
                    )}
                  </span>
                  <span className="flex flex-wrap gap-1.5">
                    {g.eventi.map((e) => (
                      <span
                        key={e.id}
                        className="inline-flex items-center gap-1 rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground"
                        title={e.origine === "Manuale" ? t("mie.manuale") : undefined}
                      >
                        <Clock className="h-3 w-3" />
                        {t(`evento.${e.evento}`)} {e.dataOra.slice(11, 16)}
                        {e.origine === "Manuale" ? " (M)" : ""}
                      </span>
                    ))}
                  </span>
                  <span className="ml-auto">
                    {corr ? (
                      <span
                        className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${
                          corr.stato === "Approvata"
                            ? "bg-status-present/15 text-status-present"
                            : corr.stato === "Respinta"
                              ? "bg-status-absent/15 text-status-absent"
                              : "bg-status-break/15 text-status-break"
                        }`}
                      >
                        {t(`mie.stato.${corr.stato}`)}
                      </span>
                    ) : (
                      <button
                        type="button"
                        onClick={() => apriCorrezione(g.giorno, g.eventi)}
                        className="inline-flex items-center gap-1.5 rounded-lg border border-border px-3 py-1.5 text-xs text-foreground hover:bg-muted"
                      >
                        <PenLine className="h-3.5 w-3.5" /> {t("mie.chiediCorrezione")}
                      </button>
                    )}
                  </span>
                </div>

                {giornoCorr === g.giorno && (
                  <div className="mt-3 rounded-xl border border-border p-3 space-y-3">
                    <p className="text-[13px] text-muted-foreground">{t("mie.corrDesc")}</p>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("mie.orari")}</label>
                      <input
                        value={orari}
                        onChange={(e) => setOrari(e.target.value)}
                        placeholder="entrata 08:00, inizio-pausa 12:30, fine-pausa 13:00, uscita 17:00"
                        className={inputCls}
                      />
                    </div>
                    <div>
                      <label className="text-xs text-muted-foreground">{t("mie.motivo")}</label>
                      <input
                        value={motivo}
                        onChange={(e) => setMotivo(e.target.value)}
                        placeholder={t("mie.motivoPh")}
                        className={inputCls}
                      />
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => void invia()}
                        disabled={inviando || !orari.trim() || !motivo.trim()}
                        className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                      >
                        {inviando ? (
                          <Loader2 className="h-4 w-4 animate-spin" />
                        ) : (
                          <Send className="h-4 w-4" />
                        )}
                        {t("mie.invia")}
                      </button>
                      <button
                        type="button"
                        onClick={() => setGiornoCorr(null)}
                        className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                      >
                        {t("common.cancel")}
                      </button>
                    </div>
                  </div>
                )}
              </li>
            );
          })}
        </ul>
      )}
    </AppShell>
  );
}
