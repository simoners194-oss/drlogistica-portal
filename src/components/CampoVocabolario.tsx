// DR Portal — campo a VOCABOLARIO condiviso (regole movimenti e fatture).
// Menu a discesa con le sole voci gia' usate nelle regole (ordine
// alfabetico) + "Nuova voce" che apre il campo libero. Un valore fuori
// elenco (regola vecchia in modifica) apre direttamente il campo libero,
// cosi' non si perde niente.
import { useState } from "react";

const inputCls =
  "w-full rounded-lg border border-border bg-background px-3 py-2 text-sm text-foreground focus:outline-none focus:ring-2 focus:ring-primary/40";

export function CampoVocabolario({
  label,
  valore,
  onChange,
  opzioni,
  testoNessuno,
  testoNuova,
}: {
  label: string;
  valore: string;
  onChange: (v: string) => void;
  opzioni: string[];
  testoNessuno: string;
  testoNuova: string;
}) {
  const [libero, setLibero] = useState(false);
  const fuoriElenco = valore !== "" && !opzioni.includes(valore);
  const modoLibero = libero || fuoriElenco;
  return (
    <>
      <label className="text-xs text-muted-foreground">{label}</label>
      {modoLibero ? (
        <div className="flex gap-1.5">
          <input
            value={valore}
            onChange={(e) => onChange(e.target.value)}
            placeholder={testoNuova}
            autoFocus={libero}
            className={inputCls}
          />
          <button
            type="button"
            onClick={() => {
              setLibero(false);
              onChange("");
            }}
            title={testoNessuno}
            className="shrink-0 rounded-lg border border-border px-2 text-xs hover:bg-muted"
          >
            ↩
          </button>
        </div>
      ) : (
        <select
          value={valore}
          onChange={(e) => {
            if (e.target.value === "__nuova__") {
              setLibero(true);
              onChange("");
            } else onChange(e.target.value);
          }}
          className={inputCls}
        >
          <option value="">{testoNessuno}</option>
          {opzioni.map((o) => (
            <option key={o} value={o}>
              {o}
            </option>
          ))}
          <option value="__nuova__">{testoNuova}</option>
        </select>
      )}
    </>
  );
}
