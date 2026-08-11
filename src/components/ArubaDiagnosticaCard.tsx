// DR Portal — Amministrazione → card "Diagnostica Aruba".
// I tre PROBE del collegamento API (connessione, download XML, incassi)
// vivono qui, fuori dalle mani di tutti i giorni: servono quando qualcosa
// non torna o quando si esplora un pezzo nuovo dell'API. Il nome file
// (facoltativo) punta i probe su una fattura precisa — e' il modo per
// indagare un lotto che il sync segnala come non parsato.
import { useEffect, useState } from "react";
import { Loader2, Plug } from "lucide-react";
import { useLang } from "@/lib/i18n";
import {
  spGetArubaStato,
  spArubaProvaConnessione,
  spArubaProvaDownload,
  spArubaProvaIncassi,
} from "@/lib/sharepoint.functions";
import type { ArubaStato } from "@/lib/sharepoint.server";
import type { ArubaProbeResult, ArubaDownloadProbe, ArubaIncassiProbe } from "@/lib/aruba.server";

export function ArubaDiagnosticaCard() {
  const { t } = useLang();
  const [aruba, setAruba] = useState<ArubaStato | null>(null);
  const [filename, setFilename] = useState("");
  const [probe, setProbe] = useState<ArubaProbeResult | null>(null);
  const [probeDl, setProbeDl] = useState<ArubaDownloadProbe | null>(null);
  const [incProbe, setIncProbe] = useState<ArubaIncassiProbe | null>(null);
  const [busy, setBusy] = useState<"conn" | "dl" | "inc" | null>(null);

  useEffect(() => {
    spGetArubaStato()
      .then((s) => setAruba(s as ArubaStato))
      .catch(() => setAruba(null));
  }, []);

  const btnCls =
    "inline-flex items-center gap-2 rounded-lg border border-border px-3 py-2 text-sm text-foreground hover:bg-muted disabled:opacity-50";
  const configurato = Boolean(aruba?.configurato);

  return (
    <div className="rounded-2xl border border-border bg-card p-5 shadow-[var(--shadow-card)]">
      <div className="text-sm font-semibold text-foreground mb-1">{t("ft.arDiagTitolo")}</div>
      <p className="text-xs text-muted-foreground mb-3">{t("ft.arDiagDesc")}</p>
      <div className="flex flex-wrap items-end gap-2">
        <div>
          <label className="text-xs text-muted-foreground">{t("ft.arIncFile")}</label>
          <input
            value={filename}
            onChange={(e) => setFilename(e.target.value)}
            placeholder="IT…xml"
            className="w-64 rounded-lg border border-border bg-background px-2 py-2 text-sm"
          />
        </div>
        <button
          type="button"
          disabled={busy != null || !configurato}
          onClick={() => {
            setBusy("conn");
            setProbe(null);
            spArubaProvaConnessione()
              .then((r) => setProbe(r as ArubaProbeResult))
              .catch((err) =>
                setProbe({
                  ok: false,
                  messaggio: err instanceof Error ? err.message : String(err),
                }),
              )
              .finally(() => setBusy(null));
          }}
          className={btnCls}
        >
          {busy === "conn" ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <Plug className="h-4 w-4" />
          )}
          {t("ft.arProva")}
        </button>
        <button
          type="button"
          disabled={busy != null || !configurato}
          onClick={() => {
            setBusy("dl");
            setProbeDl(null);
            spArubaProvaDownload({ data: { filename } })
              .then((r) => setProbeDl(r as ArubaDownloadProbe))
              .catch((err) =>
                setProbeDl({
                  ok: false,
                  messaggio: err instanceof Error ? err.message : String(err),
                  tentativi: [],
                }),
              )
              .finally(() => setBusy(null));
          }}
          className={btnCls}
        >
          {busy === "dl" ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <Plug className="h-4 w-4" />
          )}
          {t("ft.arProvaDl")}
        </button>
        <button
          type="button"
          disabled={busy != null || !configurato}
          onClick={() => {
            setBusy("inc");
            setIncProbe(null);
            spArubaProvaIncassi({ data: { filename } })
              .then((r) => setIncProbe(r as ArubaIncassiProbe))
              .catch((err) =>
                setIncProbe({
                  ok: false,
                  messaggio: err instanceof Error ? err.message : String(err),
                }),
              )
              .finally(() => setBusy(null));
          }}
          className={btnCls}
        >
          {busy === "inc" ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <Plug className="h-4 w-4" />
          )}
          {t("ft.arIncassi")}
        </button>
      </div>
      {!configurato && aruba != null && (
        <p className="mt-2 text-[11px] text-muted-foreground">{t("ft.arDiagNoCfg")}</p>
      )}
      {probe && (
        <div
          className={`mt-3 rounded-lg p-3 text-[13px] ${probe.ok ? "bg-status-present/10 text-foreground" : "bg-status-absent/10 text-status-absent"}`}
        >
          <div className="font-medium">{probe.messaggio}</div>
          {probe.ok && probe.campiEsempio && (
            <div className="mt-2 overflow-x-auto">
              <table className="text-xs">
                <tbody>
                  {Object.entries(probe.campiEsempio).map(([k, v]) => (
                    <tr key={k}>
                      <td className="pr-3 py-0.5 font-mono text-foreground whitespace-nowrap">
                        {k}
                      </td>
                      <td className="py-0.5 break-all text-muted-foreground">{v}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
      {probeDl && (
        <div
          className={`mt-3 rounded-lg p-3 text-[13px] ${probeDl.ok ? "bg-status-present/10 text-foreground" : "bg-status-absent/10 text-status-absent"}`}
        >
          <div className="font-medium">{probeDl.messaggio}</div>
          {probeDl.tentativi.length > 0 && (
            <div className="mt-2 overflow-x-auto">
              <table className="text-xs">
                <tbody>
                  {probeDl.tentativi.map((tv) => (
                    <tr key={tv.percorso}>
                      <td className="pr-3 py-0.5 font-mono whitespace-nowrap text-foreground">
                        {tv.percorso}
                      </td>
                      <td className="pr-3 py-0.5 tabular-nums">{tv.status}</td>
                      <td className="pr-3 py-0.5">{tv.contentType ?? ""}</td>
                      <td className="py-0.5 break-all text-muted-foreground">
                        {tv.chiavi ? `[${tv.chiavi.join(", ")}] ` : ""}
                        {tv.anteprima ?? ""}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
      {incProbe && (
        <div
          className={`mt-3 rounded-lg p-3 text-[13px] ${incProbe.ok ? "bg-status-present/10 text-foreground" : "bg-status-absent/10 text-status-absent"}`}
        >
          <div className="font-medium">{incProbe.messaggio}</div>
          {incProbe.dettaglio && (
            <pre className="mt-2 max-h-96 overflow-auto whitespace-pre-wrap break-all text-[11px] text-muted-foreground">
              {incProbe.dettaglio}
            </pre>
          )}
        </div>
      )}
    </div>
  );
}
