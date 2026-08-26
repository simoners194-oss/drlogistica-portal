// DR Portal — modulo Mezzi: tab Parco (elenco, filtri, scheda mezzo).

import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Plus, ExternalLink } from "lucide-react";
import {
  affidamentoCorrente,
  giorniA,
  gruppoDaStato,
  MEZZO_STATI,
  normalizzaTarga,
  type MezziDb,
  type Mezzo,
  type MezzoStato,
} from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaMezzo } from "@/lib/mezzi.functions";
import {
  ConfirmButton,
  Field,
  fmtData,
  inputCls,
  num,
  SemaforoBadge,
  StatoMezzoBadge,
  useMezzi,
} from "./shared";

export function ParcoTab({
  db,
  onDb,
  isAdmin,
}: {
  db: MezziDb;
  onDb: (db: MezziDb) => void;
  isAdmin: boolean;
}) {
  const { m } = useMezzi();
  const [q, setQ] = useState("");
  const [gruppo, setGruppo] = useState<string>("operativo");
  const [statoF, setStatoF] = useState<string>("");
  const [appaltoF, setAppaltoF] = useState<string>("");
  const [proprietaF, setProprietaF] = useState<string>("");
  const [aperto, setAperto] = useState<Mezzo | null>(null);
  const [nuovo, setNuovo] = useState(false);

  const appalti = useMemo(
    () => [...new Set(db.mezzi.map((x) => x.appalto).filter(Boolean))].sort() as string[],
    [db.mezzi],
  );

  const righe = useMemo(() => {
    const ql = q.trim().toLowerCase();
    return [...db.mezzi]
      .filter((x) => !gruppo || x.gruppo === gruppo)
      .filter((x) => !statoF || x.stato === statoF)
      .filter((x) => !appaltoF || x.appalto === appaltoF)
      .filter((x) => !proprietaF || x.proprieta === proprietaF)
      .filter(
        (x) =>
          !ql ||
          [x.targa, x.modello, x.marca, x.tipoMezzo, x.appalto, x.note]
            .join(" ")
            .toLowerCase()
            .includes(ql),
      )
      .sort((a, b) => a.targa.localeCompare(b.targa));
  }, [db.mezzi, q, gruppo, statoF, appaltoF, proprietaF]);

  const prossimaScadenza = (mezzoId: string) => {
    const aperte = db.scadenze
      .filter((s) => s.mezzoId === mezzoId && !s.chiusa)
      .sort((a, b) => (a.scadenza < b.scadenza ? -1 : 1));
    return aperte[0] ?? null;
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end gap-2">
        <input
          className={`${inputCls} max-w-56`}
          placeholder={m("cerca")}
          value={q}
          onChange={(e) => setQ(e.target.value)}
        />
        <select
          className={`${inputCls} max-w-40`}
          value={gruppo}
          onChange={(e) => setGruppo(e.target.value)}
        >
          <option value="">{m("tutti")}</option>
          <option value="operativo">{m("operativi")}</option>
          <option value="altro">{m("altri")}</option>
        </select>
        <select
          className={`${inputCls} max-w-44`}
          value={statoF}
          onChange={(e) => setStatoF(e.target.value)}
        >
          <option value="">
            {m("stato")}: {m("tutti").toLowerCase()}
          </option>
          {MEZZO_STATI.map((s) => (
            <option key={s} value={s}>
              {m(`stato.${s}` as never)}
            </option>
          ))}
        </select>
        <select
          className={`${inputCls} max-w-48`}
          value={appaltoF}
          onChange={(e) => setAppaltoF(e.target.value)}
        >
          <option value="">
            {m("appalto")}: {m("tutti").toLowerCase()}
          </option>
          {appalti.map((a) => (
            <option key={a} value={a}>
              {a}
            </option>
          ))}
        </select>
        <select
          className={`${inputCls} max-w-44`}
          value={proprietaF}
          onChange={(e) => setProprietaF(e.target.value)}
        >
          <option value="">
            {m("proprieta")}: {m("tutte").toLowerCase()}
          </option>
          <option value="noleggio">{m("noleggio")}</option>
          <option value="proprieta">{m("diProprieta")}</option>
        </select>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuovo(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoMezzo")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("tipoMezzo")}</TableHead>
              <TableHead>{m("modello")}</TableHead>
              <TableHead>{m("stato")}</TableHead>
              <TableHead>{m("appalto")}</TableHead>
              <TableHead>{m("autista")}</TableHead>
              <TableHead>{m("scadenza")}</TableHead>
              <TableHead className="text-right">{m("kmAttuali")}</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {righe.length === 0 && (
              <TableRow>
                <TableCell colSpan={8} className="text-center text-muted-foreground">
                  {m("nessunRecord")}
                </TableCell>
              </TableRow>
            )}
            {righe.map((x) => {
              const aff = affidamentoCorrente(db.affidamenti, x.id);
              const sc = prossimaScadenza(x.id);
              return (
                <TableRow key={x.id} className="cursor-pointer" onClick={() => setAperto(x)}>
                  <TableCell className="font-mono font-semibold">{x.targa}</TableCell>
                  <TableCell className="max-w-44 truncate">{x.tipoMezzo ?? "—"}</TableCell>
                  <TableCell className="max-w-40 truncate">
                    {[x.marca, x.modello].filter(Boolean).join(" ") || "—"}
                  </TableCell>
                  <TableCell>
                    <StatoMezzoBadge stato={x.stato} />
                  </TableCell>
                  <TableCell className="max-w-36 truncate">{x.appalto ?? "—"}</TableCell>
                  <TableCell className="max-w-40 truncate">{aff?.autistaNome ?? "—"}</TableCell>
                  <TableCell>
                    {sc ? <SemaforoBadge giorni={giorniA(sc.scadenza)} /> : "—"}
                  </TableCell>
                  <TableCell className="text-right tabular-nums">
                    {x.kmAttuali?.toLocaleString("it-IT") ?? "—"}
                  </TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      </div>

      {(aperto || nuovo) && (
        <MezzoDialog
          key={aperto?.id ?? "nuovo"}
          db={db}
          mezzo={aperto}
          isAdmin={isAdmin}
          onClose={() => {
            setAperto(null);
            setNuovo(false);
          }}
          onDb={onDb}
        />
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------

function MezzoDialog({
  db,
  mezzo,
  isAdmin,
  onClose,
  onDb,
}: {
  db: MezziDb;
  mezzo: Mezzo | null;
  isAdmin: boolean;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<Mezzo>(
    mezzo ?? {
      id: "",
      targa: "",
      gruppo: "operativo",
      stato: "operativo",
      proprieta: "noleggio",
    },
  );
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<Mezzo>) => setF((p) => ({ ...p, ...patch }));
  const setLib = (patch: Partial<NonNullable<Mezzo["libretto"]>>) =>
    setF((p) => ({ ...p, libretto: { ...(p.libretto ?? {}), ...patch } }));

  const salva = async () => {
    setSaving(true);
    try {
      const nuovo = { ...f };
      if (!nuovo.id) nuovo.id = normalizzaTarga(nuovo.targa);
      nuovo.gruppo = nuovo.gruppo ?? gruppoDaStato(nuovo.stato);
      const res = await spMezziSalvaMezzo({ data: { mezzo: nuovo } });
      onDb(res);
      toast.success(m("salvato"));
      onClose();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const elimina = async () => {
    if (!mezzo) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "mezzi", id: mezzo.id } });
      onDb(res);
      toast.success(m("eliminato"));
      onClose();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const link = (url?: string, label?: string) =>
    url ? (
      <a
        href={url}
        target="_blank"
        rel="noreferrer"
        className="inline-flex items-center gap-1 text-primary hover:underline"
      >
        {label ?? m("apri")} <ExternalLink className="h-3 w-3" />
      </a>
    ) : null;

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-h-[90vh] max-w-3xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="font-mono">
            {mezzo ? `${mezzo.targa} — ${m("dettaglio")}` : m("nuovoMezzo")}
          </DialogTitle>
        </DialogHeader>

        <div className="space-y-5">
          <section>
            <h3 className="mb-2 text-sm font-semibold">{m("anagrafica")}</h3>
            <div className="grid grid-cols-2 gap-3 md:grid-cols-3">
              <Field label={m("targa")}>
                <input
                  className={inputCls}
                  value={f.targa}
                  onChange={(e) => set({ targa: e.target.value })}
                  disabled={!!mezzo}
                />
              </Field>
              <Field label={m("stato")}>
                <select
                  className={inputCls}
                  value={f.stato}
                  onChange={(e) => {
                    const stato = e.target.value as MezzoStato;
                    set({ stato, gruppo: gruppoDaStato(stato) });
                  }}
                >
                  {MEZZO_STATI.map((s) => (
                    <option key={s} value={s}>
                      {m(`stato.${s}` as never)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={m("gruppo")}>
                <select
                  className={inputCls}
                  value={f.gruppo}
                  onChange={(e) => set({ gruppo: e.target.value as Mezzo["gruppo"] })}
                >
                  <option value="operativo">{m("operativi")}</option>
                  <option value="altro">{m("altri")}</option>
                </select>
              </Field>
              <Field label={m("proprieta")}>
                <select
                  className={inputCls}
                  value={f.proprieta}
                  onChange={(e) => set({ proprieta: e.target.value as Mezzo["proprieta"] })}
                >
                  <option value="noleggio">{m("noleggio")}</option>
                  <option value="proprieta">{m("diProprieta")}</option>
                </select>
              </Field>
              <Field label={m("tipoMezzo")}>
                <input
                  className={inputCls}
                  value={f.tipoMezzo ?? ""}
                  onChange={(e) => set({ tipoMezzo: e.target.value })}
                />
              </Field>
              <Field label={m("societa")}>
                <input
                  className={inputCls}
                  value={f.societa ?? ""}
                  onChange={(e) => set({ societa: e.target.value })}
                />
              </Field>
              <Field label={m("marca")}>
                <input
                  className={inputCls}
                  value={f.marca ?? ""}
                  onChange={(e) => set({ marca: e.target.value })}
                />
              </Field>
              <Field label={m("modello")}>
                <input
                  className={inputCls}
                  value={f.modello ?? ""}
                  onChange={(e) => set({ modello: e.target.value })}
                />
              </Field>
              <Field label={m("sede")}>
                <input
                  className={inputCls}
                  value={f.sede ?? ""}
                  onChange={(e) => set({ sede: e.target.value })}
                />
              </Field>
              <Field label={m("appalto")}>
                <input
                  className={inputCls}
                  value={f.appalto ?? ""}
                  onChange={(e) => set({ appalto: e.target.value })}
                />
              </Field>
              <Field label={m("sostitutivoDi")}>
                <input
                  className={inputCls}
                  value={f.sostitutivoDi ?? ""}
                  onChange={(e) => set({ sostitutivoDi: e.target.value.toUpperCase() })}
                  placeholder="targa"
                />
              </Field>
              <Field label="Telepass">
                <input
                  className={inputCls}
                  value={f.telepass ?? ""}
                  onChange={(e) => set({ telepass: e.target.value })}
                />
              </Field>
              <div className="flex items-end gap-4 pb-1 text-sm">
                {(["frigo", "atp", "sponda"] as const).map((k) => (
                  <label key={k} className="flex items-center gap-1.5">
                    <input
                      type="checkbox"
                      checked={!!f[k]}
                      onChange={(e) => set({ [k]: e.target.checked } as Partial<Mezzo>)}
                    />
                    {m(k)}
                  </label>
                ))}
              </div>
            </div>
          </section>

          <section>
            <h3 className="mb-2 text-sm font-semibold">{m("libretto")}</h3>
            <div className="grid grid-cols-2 gap-3 md:grid-cols-3">
              <Field label={m("telaio")}>
                <input
                  className={inputCls}
                  value={f.libretto?.telaio ?? ""}
                  onChange={(e) => setLib({ telaio: e.target.value })}
                />
              </Field>
              <Field label={m("immatricolazione")}>
                <input
                  type="date"
                  className={inputCls}
                  value={f.libretto?.immatricolazione ?? ""}
                  onChange={(e) => setLib({ immatricolazione: e.target.value })}
                />
              </Field>
              <Field label={m("intestatario")}>
                <input
                  className={inputCls}
                  value={f.libretto?.intestatario ?? ""}
                  onChange={(e) => setLib({ intestatario: e.target.value })}
                />
              </Field>
              <Field label={m("classeEuro")}>
                <input
                  className={inputCls}
                  value={f.libretto?.classeEuro ?? ""}
                  onChange={(e) => setLib({ classeEuro: e.target.value })}
                />
              </Field>
              <Field label={m("portata")}>
                <input
                  className={inputCls}
                  value={f.libretto?.portataKg ?? ""}
                  onChange={(e) => setLib({ portataKg: num(e.target.value) })}
                />
              </Field>
              <Field label={m("massaMax")}>
                <input
                  className={inputCls}
                  value={f.libretto?.massaMaxKg ?? ""}
                  onChange={(e) => setLib({ massaMaxKg: num(e.target.value) })}
                />
              </Field>
              <Field label={m("allestimento")}>
                <input
                  className={inputCls}
                  value={f.libretto?.allestimento ?? ""}
                  onChange={(e) => setLib({ allestimento: e.target.value })}
                />
              </Field>
              <Field label={m("carrozzeria")}>
                <input
                  className={inputCls}
                  value={f.libretto?.carrozzeria ?? ""}
                  onChange={(e) => setLib({ carrozzeria: e.target.value })}
                />
              </Field>
              <Field label={m("dimensioni")}>
                <div className="flex gap-2">
                  <input
                    className={inputCls}
                    placeholder="L"
                    value={f.libretto?.lunghezzaM ?? ""}
                    onChange={(e) => setLib({ lunghezzaM: num(e.target.value) })}
                  />
                  <input
                    className={inputCls}
                    placeholder="W"
                    value={f.libretto?.larghezzaM ?? ""}
                    onChange={(e) => setLib({ larghezzaM: num(e.target.value) })}
                  />
                </div>
              </Field>
            </div>
          </section>

          <section>
            <h3 className="mb-2 text-sm font-semibold">{m("documenti")}</h3>
            <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
              {(
                [
                  ["docLibretto", "linkLibretto"],
                  ["docAssicurazione", "linkAssicurazione"],
                  ["docContratto", "linkContratto"],
                ] as const
              ).map(([campo, label]) => (
                <Field key={campo} label={m(label)}>
                  <div className="flex items-center gap-2">
                    <input
                      className={inputCls}
                      value={(f[campo] as string) ?? ""}
                      onChange={(e) => set({ [campo]: e.target.value } as Partial<Mezzo>)}
                      placeholder="https://…"
                    />
                    {link(f[campo] as string)}
                  </div>
                </Field>
              ))}
            </div>
          </section>

          <section className="grid grid-cols-1 gap-3 md:grid-cols-2">
            <Field label={m("note")}>
              <textarea
                className={`${inputCls} min-h-16`}
                value={f.note ?? ""}
                onChange={(e) => set({ note: e.target.value })}
              />
            </Field>
            {mezzo && (
              <div className="text-xs text-muted-foreground">
                <p>
                  {m("kmAttuali")}: {f.kmAttuali?.toLocaleString("it-IT") ?? "—"}
                  {f.kmAggiornatiAl ? ` (${fmtData(f.kmAggiornatiAl)})` : ""}
                </p>
                <p className="mt-1">
                  {m("tab.scadenze")}:{" "}
                  {db.scadenze.filter((s) => s.mezzoId === mezzo.id && !s.chiusa).length} ·{" "}
                  {m("tab.multe")}: {db.multe.filter((s) => s.mezzoId === mezzo.id).length} ·{" "}
                  {m("tab.officina")}: {db.officina.filter((s) => s.mezzoId === mezzo.id).length}
                </p>
              </div>
            )}
          </section>
        </div>

        <div className="mt-2 flex items-center justify-between gap-2">
          <div>
            {mezzo && isAdmin && (
              <ConfirmButton label={m("elimina")} onConfirm={elimina} disabled={saving} />
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.targa.trim()}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
