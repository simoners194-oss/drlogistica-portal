// DR Portal — Amministrazione del collegamento banca (Enable Banking, PSD2).
// Configurazione dell'app (chiave privata cifrata server-side), autorizzazione
// con SCA della banca, scelta del conto, attivazione (taglio Excel→API) e
// sincronizzazione manuale. Tutto in SOLA LETTURA.
// L'uso quotidiano (tasto Aggiorna + saldo) vive in Finanza → Movimenti.
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { Landmark, Loader2, CheckCircle2 } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { useLang } from "@/lib/i18n";
import {
  spEbStato,
  spEbSalvaApp,
  spEbProva,
  spEbSaldo,
  spEbAvviaCollegamento,
  spEbCompletaCollegamento,
  spEbScegliConto,
  spEbTaglia,
  spEbSincronizza,
} from "@/lib/sharepoint.functions";
import type { EbStato, EbSyncResult, EbSaldoInfo } from "@/lib/sharepoint.server";

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

function fmtData(iso?: string | null): string {
  if (!iso) return "—";
  const [y, m, g] = iso.slice(0, 10).split("-");
  return y && m && g ? `${g}/${m}/${y}` : iso;
}
function fmtImporto(n: number): string {
  return n.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function BancaPsd2Panel() {
  const { t } = useLang();
  const [stato, setStato] = useState<EbStato | null>(null);
  const [appId, setAppId] = useState("");
  const [pem, setPem] = useState("");
  const [conti, setConti] = useState<{ uid: string; iban: string; nome: string }[] | null>(null);
  const [busy, setBusy] = useState<"salva" | "collega" | "completa" | "taglio" | "sync" | null>(
    null,
  );
  const [progress, setProgress] = useState("");
  const [editApp, setEditApp] = useState(false);
  const [saldo, setSaldo] = useState<EbSaldoInfo | null>(null);
  const [saldoErr, setSaldoErr] = useState<string | null>(null);

  const errMsg = (err: unknown) => (err instanceof Error ? err.message : String(err));

  const loadStato = () => {
    spEbStato()
      .then((s) => setStato(s as EbStato))
      .catch((err) => toast.error(t("fin.ebErr"), { description: errMsg(err) }));
  };
  // Lettura dalla CACHE (mai dalla banca: il limite PSD2 giornaliero si
  // spende solo durante la sincronizzazione, che aggiorna anche il saldo).
  const loadSaldo = () => {
    spEbSaldo()
      .then((s) => {
        setSaldo(s as EbSaldoInfo | null);
        setSaldoErr(s == null ? t("fin.ebSaldoCacheVuota") : null);
      })
      .catch((err) => {
        setSaldo(null);
        setSaldoErr(errMsg(err));
      });
  };

  // Completa il collegamento con il codice arrivato dalla redirect della banca
  // (messo da parte dalla pagina di login in sessionStorage).
  const completa = async (code: string) => {
    setBusy("completa");
    try {
      const res = await spEbCompletaCollegamento({ data: { code } });
      setConti(res.conti);
      loadStato();
      toast.success(t("fin.ebScegliConto"));
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
      try {
        window.sessionStorage.removeItem("dr:eb:code");
      } catch {
        /* sessionStorage non disponibile */
      }
    }
  };

  useEffect(() => {
    loadStato();
    loadSaldo();
    let code: string | null = null;
    try {
      code = window.sessionStorage.getItem("dr:eb:code");
    } catch {
      /* sessionStorage non disponibile */
    }
    if (code) void completa(code);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const salvaApp = async () => {
    setBusy("salva");
    try {
      await spEbSalvaApp({ data: { appId: appId.trim(), privateKeyPem: pem.trim() } });
      toast.success(t("fin.ebSalvata"));
      setPem("");
      setEditApp(false);
      loadStato();
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
    }
  };

  // Prova SENZA passare dalla banca: interroga l'anagrafica dell'app registrata.
  const prova = async () => {
    setBusy("salva");
    try {
      const r = await spEbProva();
      toast.success(t("fin.ebProvaOk"), {
        description: `${r.nome} · ${r.ambiente} · ${r.redirect.join(", ")}`,
      });
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
    }
  };

  const collega = async () => {
    setBusy("collega");
    try {
      const r = await spEbAvviaCollegamento();
      window.location.href = r.url; // la banca rimanda poi su portal…/?code=…
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
      setBusy(null);
    }
  };

  const usaConto = async (uid: string, iban: string) => {
    setBusy("completa");
    try {
      await spEbScegliConto({ data: { uid, iban } });
      setConti(null);
      loadStato();
      loadSaldo();
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
    }
  };

  // Passaggio Excel → API: elimina l'ultimo giorno importato (a blocchi) e
  // fissa la data di taglio; da lì scrive solo la banca.
  const attiva = async () => {
    if (!window.confirm(t("fin.ebAttivaConfirm"))) return;
    setBusy("taglio");
    try {
      let eliminati = 0;
      let guard = 0;
      for (;;) {
        const r = await spEbTaglia();
        eliminati += r.eliminati;
        setProgress(String(eliminati));
        if (r.rimanenti <= 0 || ++guard > 60) break;
      }
      loadStato();
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
      setProgress("");
    }
  };

  const sincronizza = async () => {
    setBusy("sync");
    const importId = `SYNC-${new Date().toISOString().slice(0, 19)}`;
    let scritti = 0;
    let doppioni = 0;
    let pendenti = 0;
    try {
      let continuation: string | undefined;
      let guard = 0;
      for (;;) {
        const r = (await spEbSincronizza({ data: { importId, continuation } })) as EbSyncResult;
        scritti += r.scritti;
        doppioni += r.doppioni;
        pendenti += r.pendenti;
        if (r.errori.length) throw new Error(r.errori[0]);
        setProgress(String(scritti));
        if (!r.continuation || ++guard > 100) break;
        continuation = r.continuation;
      }
      toast.success(t("fin.ebSyncDone"), {
        description: `${scritti} ${t("fin.ebNuovi")} · ${doppioni} ${t("fin.ebGiaPresenti")} · ${pendenti} ${t("fin.ebPendenti")}`,
      });
      loadStato();
      loadSaldo();
    } catch (err) {
      toast.error(t("fin.ebErr"), { description: errMsg(err) });
    } finally {
      setBusy(null);
      setProgress("");
    }
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-base">
          <Landmark className="h-5 w-5 text-primary" />
          {t("fin.ebTitle")}
        </CardTitle>
        <p className="text-xs text-muted-foreground mt-1">{t("fin.ebDesc")}</p>
      </CardHeader>
      <CardContent>
        {stato == null ? (
          <p className="text-sm text-muted-foreground inline-flex items-center gap-2">
            <Loader2 className="h-4 w-4 animate-spin" /> …
          </p>
        ) : !stato.listaPresente ? (
          <p className="text-sm text-status-absent">{t("fin.ebListMissing")}</p>
        ) : (
          <div className="space-y-4">
            {stato.colonneMancanti.length > 0 && (
              <p className="text-xs text-status-absent">
                {t("fin.ebColsMissing")} {stato.colonneMancanti.join(", ")}
              </p>
            )}

            {!stato.configurato || editApp ? (
              <div className="space-y-3 max-w-xl">
                <p className="text-sm text-muted-foreground">{t("fin.ebNonConfig")}</p>
                <div>
                  <label className="text-xs text-muted-foreground">{t("fin.ebAppId")}</label>
                  <input
                    value={appId}
                    onChange={(e) => setAppId(e.target.value)}
                    placeholder="00000000-0000-0000-0000-000000000000"
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className="text-xs text-muted-foreground">{t("fin.ebPem")}</label>
                  <textarea
                    value={pem}
                    onChange={(e) => setPem(e.target.value)}
                    placeholder="-----BEGIN PRIVATE KEY-----"
                    rows={4}
                    className={`${inputCls} font-mono text-xs`}
                  />
                </div>
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => void salvaApp()}
                    disabled={busy != null || !appId.trim() || !pem.trim()}
                    className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                  >
                    {busy === "salva" && <Loader2 className="h-4 w-4 animate-spin" />}
                    {t("fin.ebSalvaApp")}
                  </button>
                  {editApp && (
                    <button
                      type="button"
                      onClick={() => setEditApp(false)}
                      className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted"
                    >
                      {t("common.cancel")}
                    </button>
                  )}
                </div>
              </div>
            ) : (
              <>
                <div className="grid gap-x-8 gap-y-1 sm:grid-cols-2 text-[13px]">
                  <div className="sm:col-span-2">
                    <span className="text-muted-foreground">{t("fin.ebAppId")}: </span>
                    <b className="font-mono text-xs">{stato.appId || "—"}</b>
                  </div>
                  <div>
                    <span className="text-muted-foreground">{t("fin.ebConto")}: </span>
                    <b>{stato.contoIban || "—"}</b>
                  </div>
                  <div>
                    <span className="text-muted-foreground">{t("fin.ebConsenso")}: </span>
                    <b>{fmtData(stato.consensoScade)}</b>
                  </div>
                  <div>
                    <span className="text-muted-foreground">{t("fin.ebTaglio")}: </span>
                    <b>{fmtData(stato.dataTaglio)}</b>
                  </div>
                  <div>
                    <span className="text-muted-foreground">{t("fin.ebUltimaSync")}: </span>
                    <b>
                      {stato.ultimaSync
                        ? `${fmtData(stato.ultimaSync)} ${stato.ultimaSync.slice(11, 16)}`
                        : t("fin.ebMai")}
                    </b>
                  </div>
                  {saldo && (
                    <div>
                      <span className="text-muted-foreground">{t("fin.ebSaldoAttuale")}: </span>
                      <b className="tabular-nums">
                        {fmtImporto(saldo.saldo)} {saldo.divisa}
                      </b>
                    </div>
                  )}
                </div>

                {conti && (
                  <div className="rounded-xl border border-border p-3">
                    <div className="text-sm font-medium text-foreground mb-2">
                      {t("fin.ebScegliConto")}
                    </div>
                    <div className="space-y-2">
                      {conti.map((c) => (
                        <div key={c.uid} className="flex flex-wrap items-center gap-3">
                          <span className="text-[13px] font-mono">{c.iban || c.uid}</span>
                          {c.nome && (
                            <span className="text-xs text-muted-foreground">{c.nome}</span>
                          )}
                          <button
                            type="button"
                            onClick={() => void usaConto(c.uid, c.iban)}
                            disabled={busy != null}
                            className="rounded-lg border border-border px-3 py-1.5 text-xs text-foreground hover:bg-muted disabled:opacity-50"
                          >
                            {t("fin.ebUsaConto")}
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="flex flex-wrap items-center gap-3">
                  <button
                    type="button"
                    onClick={() => void prova()}
                    disabled={busy != null}
                    className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
                  >
                    {busy === "salva" && <Loader2 className="h-4 w-4 animate-spin" />}
                    {t("fin.ebProvaBtn")}
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setAppId(stato.appId);
                      setEditApp(true);
                    }}
                    disabled={busy != null}
                    className="rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
                  >
                    {t("fin.ebModifica")}
                  </button>
                  <button
                    type="button"
                    onClick={() => void collega()}
                    disabled={busy != null}
                    className="inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50"
                  >
                    {busy === "collega" ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <Landmark className="h-4 w-4" />
                    )}
                    {stato.consensoScade ? t("fin.ebRinnova") : t("fin.ebCollega")}
                  </button>
                  {!stato.dataTaglio && stato.contoIban && (
                    <button
                      type="button"
                      onClick={() => void attiva()}
                      disabled={busy != null}
                      className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                    >
                      {busy === "taglio" && <Loader2 className="h-4 w-4 animate-spin" />}
                      {t("fin.ebAttiva")}
                      {busy === "taglio" && progress && ` (−${progress})`}
                    </button>
                  )}
                  {stato.dataTaglio && stato.contoIban && (
                    <button
                      type="button"
                      onClick={() => void sincronizza()}
                      disabled={busy != null}
                      className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90 disabled:opacity-50"
                    >
                      {busy === "sync" ? (
                        <Loader2 className="h-4 w-4 animate-spin" />
                      ) : (
                        <CheckCircle2 className="h-4 w-4" />
                      )}
                      {t("fin.ebSync")}
                      {busy === "sync" && progress && ` (${progress})`}
                    </button>
                  )}
                  {busy === "completa" && (
                    <span className="text-xs text-muted-foreground inline-flex items-center gap-2">
                      <Loader2 className="h-3.5 w-3.5 animate-spin" /> {t("fin.ebCompleto")}
                    </span>
                  )}
                </div>
                {saldoErr && (
                  <p className="text-xs text-status-absent">
                    {t("fin.ebSaldoAttuale")}: {saldoErr}
                  </p>
                )}
                <p className="text-[11px] text-muted-foreground">{t("fin.ebCollegaDesc")}</p>
              </>
            )}
          </div>
        )}
      </CardContent>
    </Card>
  );
}
