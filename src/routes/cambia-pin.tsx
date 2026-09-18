import { createFileRoute, Link } from "@tanstack/react-router";
import { useState } from "react";
import { Logo } from "@/components/Logo";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { spRichiediCambioPin, spConfermaCambioPin } from "@/lib/sharepoint.functions";
import { AppFooter } from "@/components/AppFooter";
import { LangSwitcher, useLang } from "@/lib/i18n";

export const Route = createFileRoute("/cambia-pin")({
  head: () => ({
    meta: [
      { title: "Cambia PIN — DR Portal" },
      {
        name: "description",
        content: "Cambia il PIN di DR Portal con il codice di verifica inviato alla tua email.",
      },
    ],
  }),
  component: CambiaPinPage,
});

// Pagina PUBBLICA (chi ha dimenticato il PIN non può autenticarsi).
// Passo 1: codice dipendente → il server manda un codice di verifica di 6
// cifre all'email registrata in anagrafica. Passo 2: codice ricevuto + PIN
// nuovo (due volte) → il PIN cambia. Tutte le difese stanno lato server.
function CambiaPinPage() {
  const { t } = useLang();
  const [passo, setPasso] = useState<1 | 2>(1);
  const [codice, setCodice] = useState("");
  const [otp, setOtp] = useState("");
  const [pin1, setPin1] = useState("");
  const [pin2, setPin2] = useState("");
  const [emailMascherata, setEmailMascherata] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const richiedi = async (e?: React.FormEvent) => {
    e?.preventDefault();
    if (busy || !codice.trim()) return;
    setError(null);
    setBusy(true);
    try {
      const res = await spRichiediCambioPin({ data: { codice: codice.trim() } });
      if (!res.ok) {
        setError(res.error ?? t("cp.errGenerico"));
        return;
      }
      setEmailMascherata(res.emailMascherata ?? null);
      setPasso(2);
    } catch (err) {
      setError(err instanceof Error ? err.message : t("cp.errGenerico"));
    } finally {
      setBusy(false);
    }
  };

  const conferma = async (e?: React.FormEvent) => {
    e?.preventDefault();
    if (busy) return;
    setError(null);
    if (pin1 !== pin2) {
      setError(t("cp.pinDiversi"));
      return;
    }
    setBusy(true);
    try {
      const res = await spConfermaCambioPin({
        data: { codice: codice.trim(), otp: otp.trim(), nuovoPin: pin1 },
      });
      if (!res.ok) {
        setError(res.error ?? t("cp.errGenerico"));
        return;
      }
      toast.success(t("cp.fatto"));
      window.location.href = "/";
    } catch (err) {
      setError(err instanceof Error ? err.message : t("cp.errGenerico"));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-background">
      <div className="relative flex-1 flex items-center justify-center p-6">
        <LangSwitcher className="absolute top-4 right-4" />
        <div className="w-full max-w-sm">
          <div className="mb-8 flex justify-center">
            <Logo size={120} />
          </div>
          <h2 className="text-2xl font-semibold text-foreground">{t("cp.title")}</h2>
          <p className="text-sm text-muted-foreground mt-1">
            {passo === 1 ? t("cp.sub1") : t("cp.sub2")}
          </p>

          {passo === 1 ? (
            <form onSubmit={richiedi} className="mt-6 space-y-3">
              <Input
                autoFocus
                placeholder={t("login.codePlaceholder")}
                value={codice}
                onChange={(e) => {
                  setCodice(e.target.value.toUpperCase());
                  setError(null);
                }}
                disabled={busy}
              />
              {error && (
                <p className="text-xs text-status-absent" role="alert">
                  {error}
                </p>
              )}
              <Button type="submit" className="w-full h-11" disabled={busy || !codice.trim()}>
                {busy ? t("cp.invio") : t("cp.mandami")}
              </Button>
            </form>
          ) : (
            <form onSubmit={conferma} className="mt-6 space-y-3">
              <p className="text-[13px] text-muted-foreground rounded-lg bg-primary/5 p-3">
                {emailMascherata
                  ? t("cp.inviata").replace("{email}", emailMascherata)
                  : t("cp.forse")}
              </p>
              <Input
                autoFocus
                inputMode="numeric"
                autoComplete="one-time-code"
                placeholder={t("cp.otpPlaceholder")}
                value={otp}
                onChange={(e) => {
                  setOtp(e.target.value.replace(/\D/g, "").slice(0, 6));
                  setError(null);
                }}
                disabled={busy}
              />
              <Input
                type="password"
                inputMode="numeric"
                autoComplete="new-password"
                placeholder={t("cp.pinNuovo")}
                value={pin1}
                onChange={(e) => {
                  setPin1(e.target.value.replace(/\D/g, "").slice(0, 8));
                  setError(null);
                }}
                disabled={busy}
              />
              <Input
                type="password"
                inputMode="numeric"
                autoComplete="new-password"
                placeholder={t("cp.pinRipeti")}
                value={pin2}
                onChange={(e) => {
                  setPin2(e.target.value.replace(/\D/g, "").slice(0, 8));
                  setError(null);
                }}
                disabled={busy}
              />
              {error && (
                <p className="text-xs text-status-absent" role="alert">
                  {error}
                </p>
              )}
              <Button
                type="submit"
                className="w-full h-11"
                disabled={busy || otp.length !== 6 || pin1.length < 4 || !pin2}
              >
                {busy ? t("cp.invio") : t("cp.cambia")}
              </Button>
              <button
                type="button"
                className="w-full text-center text-[12px] text-muted-foreground hover:text-foreground"
                onClick={() => {
                  setPasso(1);
                  setOtp("");
                  setError(null);
                }}
              >
                {t("cp.rimanda")}
              </button>
            </form>
          )}

          <p className="mt-6 text-[12px] text-muted-foreground text-center">
            <Link to="/" className="hover:text-foreground underline underline-offset-2">
              {t("cp.tornaLogin")}
            </Link>
          </p>
          <p className="mt-2 text-[11px] text-muted-foreground text-center">{t("cp.nota")}</p>
        </div>
      </div>
      <AppFooter />
    </div>
  );
}
