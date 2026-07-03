import { createFileRoute } from "@tanstack/react-router";
import { Link, useNavigate } from "@tanstack/react-router";
import { Logo } from "@/components/Logo";
import { useEffect, useState } from "react";
import type { Dipendente } from "@/lib/mock-data";
import { dataService } from "@/lib/data-service";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { isMicrosoftAuthConfigured } from "@/lib/auth-config";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "DR Portal — Accedi" },
      { name: "description", content: "Portale presenze DR Logistica. Accedi con il tuo account Microsoft 365." },
    ],
  }),
  component: Index,
});

function Index() {
  const navigate = useNavigate();
  const msConfigured = isMicrosoftAuthConfigured();
  const [dipendenti, setDipendenti] = useState<Dipendente[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [empId, setEmpId] = useState<string>("");
  const [pin, setPin] = useState("");

  useEffect(() => {
    let mounted = true;
    dataService
      .getDipendenti()
      .then((list) => {
        if (!mounted) return;
        setDipendenti(list);
        setEmpId(list[0]?.id ?? "");
        if (list.length === 0)
          setLoadError("Nessun dipendente trovato su SharePoint. Verifica la configurazione in Amministrazione.");
      })
      .catch((err) => {
        if (!mounted) return;
        setLoadError(err instanceof Error ? err.message : String(err));
      })
      .finally(() => mounted && setLoading(false));
    return () => {
      mounted = false;
    };
  }, []);

  const handleSimpleLogin = () => {
    if (!empId) return toast.error("Seleziona un dipendente");
    if (pin.length < 4) return toast.error("PIN richiesto (min. 4 cifre)");
    const d = dipendenti.find((x) => x.id === empId);
    if (!d) return toast.error("Dipendente non trovato");
    try {
      sessionStorage.setItem(
        "dr:currentUser",
        JSON.stringify({ id: d.id, nome: d.nome, cognome: d.cognome }),
      );
    } catch {
      /* ignore */
    }
    toast.success(`Benvenuto ${d?.nome ?? ""}`);
    navigate({ to: "/presenze" });
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
            Scegli il metodo di accesso più adatto alla tua sede.
          </p>

          <Tabs defaultValue="simple" className="mt-6">
            <TabsList className="grid w-full grid-cols-3">
              <TabsTrigger value="simple">Rapido</TabsTrigger>
              <TabsTrigger value="mock">Demo</TabsTrigger>
              <TabsTrigger value="microsoft">Microsoft</TabsTrigger>
            </TabsList>

            <TabsContent value="simple" className="mt-4 space-y-3">
              <p className="text-xs text-muted-foreground">
                Accesso semplificato: seleziona il tuo nome e inserisci il PIN aziendale.
                {loading && " Caricamento dipendenti da SharePoint…"}
              </p>
              {loadError && (
                <p className="text-[11px] text-status-absent">{loadError}</p>
              )}
              <Select value={empId} onValueChange={setEmpId} disabled={loading || dipendenti.length === 0}>
                <SelectTrigger><SelectValue placeholder="Seleziona dipendente" /></SelectTrigger>
                <SelectContent>
                  {dipendenti.map((d) => (
                    <SelectItem key={d.id} value={d.id}>
                      {d.nome} {d.cognome} — {d.ruolo}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
              <Input
                type="password"
                inputMode="numeric"
                placeholder="PIN (4-6 cifre)"
                value={pin}
                onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 6))}
              />
              <Button className="w-full h-11" onClick={handleSimpleLogin} disabled={loading || !empId}>
                Accedi
              </Button>
            </TabsContent>

            <TabsContent value="mock" className="mt-4 space-y-3">
              <p className="text-xs text-muted-foreground">
                Modalità dimostrativa senza autenticazione. Utile per test e presentazioni.
              </p>
              <Link
                to="/presenze"
                className="w-full flex items-center justify-center gap-3 h-11 rounded-md border border-border bg-card hover:bg-secondary transition-colors text-sm font-medium text-foreground shadow-sm"
              >
                Entra come Dipendente
              </Link>
              <Link
                to="/dashboard"
                className="w-full flex items-center justify-center gap-3 h-11 rounded-md text-primary-foreground text-sm font-medium shadow-[var(--shadow-elegant)] transition-transform hover:scale-[1.01]"
                style={{ background: "var(--gradient-primary)" }}
              >
                Entra come HR
              </Link>
            </TabsContent>

            <TabsContent value="microsoft" className="mt-4 space-y-3">
              <p className="text-xs text-muted-foreground">
                Login con account aziendale Microsoft 365 / Entra ID (in fase di test tecnico).
              </p>
              <Button
                variant="outline"
                className="w-full h-11 gap-2"
                disabled={!msConfigured}
                onClick={() =>
                  toast.info(
                    "Login Microsoft in fase di test — usa la sezione Amministrazione per verificare la configurazione.",
                  )
                }
              >
                <MicrosoftLogo /> Continua con Microsoft
              </Button>
              {!msConfigured && (
                <p className="text-[11px] text-muted-foreground">
                  Configurazione Entra ID non presente. L'accesso Rapido resta disponibile.
                </p>
              )}
            </TabsContent>
          </Tabs>

          <p className="mt-6 text-[11px] text-muted-foreground text-center">
            L'app resta pienamente funzionante anche senza Microsoft 365.
          </p>
        </div>
      </div>
    </div>
  );
}

function MicrosoftLogo({ light = false }: { light?: boolean }) {
  const c = light ? ["#fff", "#fff", "#fff", "#fff"] : ["#F25022", "#7FBA00", "#00A4EF", "#FFB900"];
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <rect x="0" y="0" width="7" height="7" fill={c[0]} />
      <rect x="9" y="0" width="7" height="7" fill={c[1]} />
      <rect x="0" y="9" width="7" height="7" fill={c[2]} />
      <rect x="9" y="9" width="7" height="7" fill={c[3]} />
    </svg>
  );
}
