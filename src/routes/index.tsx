import { createFileRoute } from "@tanstack/react-router";
import { useNavigate } from "@tanstack/react-router";
import { Logo } from "@/components/Logo";
import { useState } from "react";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { spLogin } from "@/lib/sharepoint.functions";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "DR Portal — Accedi" },
      { name: "description", content: "Portale presenze DR Logistica. Accedi con codice dipendente e PIN." },
    ],
  }),
  component: Index,
});

function Index() {
  const navigate = useNavigate();
  const [codice, setCodice] = useState("");
  const [pin, setPin] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleLogin = async (e?: React.FormEvent) => {
    e?.preventDefault();
    if (busy) return;
    setError(null);
    if (!codice.trim() || !pin.trim()) {
      setError("Codice o PIN non validi.");
      return;
    }
    setBusy(true);
    try {
      const res = await spLogin({ data: { codice: codice.trim(), pin: pin.trim() } });
      if (!res.ok || !res.dipendente) {
        setError(res.error ?? "Codice o PIN non validi.");
        return;
      }
      const d = res.dipendente;
      try {
        sessionStorage.setItem(
          "dr:currentUser",
          JSON.stringify({ id: d.id, nome: d.nome, cognome: d.cognome }),
        );
      } catch {
        /* ignore */
      }
      toast.success(`Benvenuto ${d.nome}`);
      navigate({ to: "/presenze" });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Codice o PIN non validi.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="min-h-screen grid lg:grid-cols-2 bg-background">
      <div className="hidden lg:flex flex-col justify-between p-12 text-primary-foreground" style={{ background: "var(--gradient-hero)" }}>
        <Logo variant="light" size={144} />
        <div>
          <h1 className="text-4xl font-semibold tracking-tight leading-tight">
            Gestisci le presenze<br />in modo semplice.
          </h1>
          <p className="mt-4 text-white/80 max-w-md">
            DR Portal centralizza timbrature, pause e reportistica per tutte le sedi DR Logistica.
          </p>
        </div>
        <div className="text-xs text-white/60">© DR Logistica — Powered by Microsoft 365</div>
      </div>

      <div className="flex items-center justify-center p-6 lg:p-12">
        <div className="w-full max-w-sm">
          <div className="lg:hidden mb-10 flex justify-center"><Logo size={144} /></div>
          <h2 className="text-2xl font-semibold text-foreground">Accedi</h2>
          <p className="text-sm text-muted-foreground mt-1">
            Inserisci il tuo codice dipendente e il PIN aziendale.
          </p>

          <form onSubmit={handleLogin} className="mt-6 space-y-3">
            <Input
              autoFocus
              autoComplete="username"
              placeholder="Codice dipendente"
              value={codice}
              onChange={(e) => {
                setCodice(e.target.value.toUpperCase());
                setError(null);
              }}
              disabled={busy}
            />
            <Input
              type="password"
              inputMode="numeric"
              autoComplete="current-password"
              placeholder="PIN"
              value={pin}
              onChange={(e) => {
                setPin(e.target.value.replace(/\D/g, "").slice(0, 10));
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
              disabled={busy || !codice.trim() || !pin.trim()}
            >
              {busy ? "Verifica in corso…" : "Accedi"}
            </Button>
          </form>

          <p className="mt-6 text-[11px] text-muted-foreground text-center">
            L'autenticazione Microsoft 365 sarà attivata prossimamente.
          </p>
        </div>
      </div>
    </div>
  );
}
