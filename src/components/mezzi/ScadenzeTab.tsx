// DR Portal — modulo Mezzi: tab Scadenze (semafori + elenco + gestione).

import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Plus } from "lucide-react";
import {
  giorniA,
  SCADENZA_TIPI,
  semaforoScadenza,
  type MezziDb,
  type Scadenza,
  type Semaforo,
} from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaScadenza } from "@/lib/mezzi.functions";
import { Field, fmtData, inputCls, MezzoSelect, SemaforoBadge, targaDi, useMezzi } from "./shared";

const ORDINE_SEM: Semaforo[] = ["scaduta", "rosso", "arancio", "giallo", "ok"];

export function ScadenzeTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [tipoF, setTipoF] = useState("");
  const [mezzoF, setMezzoF] = useState("");
  const [soloAperte, setSoloAperte] = useState(true);
  const [semF, setSemF] = useState<Semaforo | "">("");
  const [aperta, setAperta] = useState<Scadenza | null>(null);
  const [nuova, setNuova] = useState(false);

  const righe = useMemo(() => {
    return db.scadenze
      .filter((s) => !soloAperte || !s.chiusa)
      .filter((s) => !tipoF || s.tipo === tipoF)
      .filter((s) => !mezzoF || s.mezzoId === mezzoF)
      .map((s) => ({ s, giorni: giorniA(s.scadenza) }))
      .filter((r) => !semF || semaforoScadenza(r.giorni) === semF)
      .sort((a, b) => a.giorni - b.giorni);
  }, [db.scadenze, tipoF, mezzoF, soloAperte, semF]);

  const conteggi = useMemo(() => {
    const c: Record<Semaforo, number> = { scaduta: 0, rosso: 0, arancio: 0, giallo: 0, ok: 0 };
    for (const s of db.scadenze) {
      if (s.chiusa) continue;
      c[semaforoScadenza(giorniA(s.scadenza))]++;
    }
    return c;
  }, [db.scadenze]);

  const toggleGestita = async (s: Scadenza) => {
    try {
      const res = await spMezziSalvaScadenza({
        data: { scadenza: { ...s, chiusa: !s.chiusa } },
      });
      onDb(res);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-5">
        {ORDINE_SEM.map((sem) => (
          <Card
            key={sem}
            className={`cursor-pointer transition ${semF === sem ? "ring-2 ring-primary" : ""}`}
            onClick={() => setSemF((p) => (p === sem ? "" : sem))}
          >
            <CardContent className="flex items-center justify-between p-3">
              <SemaforoBadge
                giorni={
                  sem === "scaduta"
                    ? -1
                    : sem === "rosso"
                      ? 15
                      : sem === "arancio"
                        ? 45
                        : sem === "giallo"
                          ? 75
                          : 999
                }
              />
              <span className="text-2xl font-bold tabular-nums">{conteggi[sem]}</span>
            </CardContent>
          </Card>
        ))}
      </div>

      <div className="flex flex-wrap items-end gap-2">
        <select
          className={`${inputCls} max-w-44`}
          value={tipoF}
          onChange={(e) => setTipoF(e.target.value)}
        >
          <option value="">
            {m("tipo")}: {m("tutti").toLowerCase()}
          </option>
          {SCADENZA_TIPI.map((t) => (
            <option key={t} value={t}>
              {m(`tipoScad.${t}` as never)}
            </option>
          ))}
        </select>
        <div className="max-w-52 flex-1">
          <MezzoSelect
            mezzi={db.mezzi}
            value={mezzoF}
            onChange={setMezzoF}
            vuotoLabel={`${m("mezzo")}: ${m("tutti").toLowerCase()}`}
          />
        </div>
        <label className="flex items-center gap-1.5 pb-2 text-sm">
          <input
            type="checkbox"
            checked={soloAperte}
            onChange={(e) => setSoloAperte(e.target.checked)}
          />
          {m("soloAperte")}
        </label>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuova(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovaScadenza")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("tipo")}</TableHead>
              <TableHead>{m("descrizione")}</TableHead>
              <TableHead>{m("scadenza")}</TableHead>
              <TableHead>{m("stato")}</TableHead>
              <TableHead>{m("alertInviati")}</TableHead>
              <TableHead />
            </TableRow>
          </TableHeader>
          <TableBody>
            {righe.length === 0 && (
              <TableRow>
                <TableCell colSpan={7} className="text-center text-muted-foreground">
                  {m("nessunRecord")}
                </TableCell>
              </TableRow>
            )}
            {righe.map(({ s, giorni }) => (
              <TableRow key={s.id} className="cursor-pointer" onClick={() => setAperta(s)}>
                <TableCell className="font-mono font-semibold">
                  {targaDi(db.mezzi, s.mezzoId)}
                </TableCell>
                <TableCell>{m(`tipoScad.${s.tipo}` as never)}</TableCell>
                <TableCell className="max-w-52 truncate">{s.descrizione ?? "—"}</TableCell>
                <TableCell>{fmtData(s.scadenza)}</TableCell>
                <TableCell>
                  {s.chiusa ? (
                    <span className="text-xs text-muted-foreground">{m("gestita")}</span>
                  ) : (
                    <SemaforoBadge giorni={giorni} />
                  )}
                </TableCell>
                <TableCell className="text-xs text-muted-foreground">
                  {Object.keys(s.alertInviati ?? {})
                    .sort((a, b) => Number(b) - Number(a))
                    .map((g) => (g === "0" ? "scad." : `${g}g`))
                    .join(", ") || "—"}
                </TableCell>
                <TableCell onClick={(e) => e.stopPropagation()}>
                  <Button variant="outline" size="sm" onClick={() => toggleGestita(s)}>
                    {s.chiusa ? m("soloAperte") : m("gestita")}
                  </Button>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperta || nuova) && (
        <ScadenzaDialog
          key={aperta?.id ?? "nuova"}
          db={db}
          scadenza={aperta}
          onClose={() => {
            setAperta(null);
            setNuova(false);
          }}
          onDb={onDb}
        />
      )}
    </div>
  );
}

function ScadenzaDialog({
  db,
  scadenza,
  onClose,
  onDb,
}: {
  db: MezziDb;
  scadenza: Scadenza | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<Scadenza>(scadenza ?? { id: "", tipo: "revisione", scadenza: "" });
  const [saving, setSaving] = useState(false);

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaScadenza({ data: { scadenza: f } });
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
    if (!scadenza || !window.confirm(m("confermaElimina"))) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "scadenze", id: scadenza.id } });
      onDb(res);
      toast.success(m("eliminato"));
      onClose();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-lg">
        <DialogHeader>
          <DialogTitle>{scadenza ? m("modifica") : m("nuovaScadenza")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <Field label={m("mezzo")}>
            <MezzoSelect
              mezzi={db.mezzi}
              value={f.mezzoId ?? ""}
              onChange={(id) => setF((p) => ({ ...p, mezzoId: id || undefined }))}
            />
          </Field>
          <Field label={m("tipo")}>
            <select
              className={inputCls}
              value={f.tipo}
              onChange={(e) => setF((p) => ({ ...p, tipo: e.target.value as Scadenza["tipo"] }))}
            >
              {SCADENZA_TIPI.map((t) => (
                <option key={t} value={t}>
                  {m(`tipoScad.${t}` as never)}
                </option>
              ))}
            </select>
          </Field>
          <Field label={m("scadenza")}>
            <input
              type="date"
              className={inputCls}
              value={f.scadenza}
              onChange={(e) => setF((p) => ({ ...p, scadenza: e.target.value }))}
            />
          </Field>
          <Field label={m("descrizione")}>
            <input
              className={inputCls}
              value={f.descrizione ?? ""}
              onChange={(e) => setF((p) => ({ ...p, descrizione: e.target.value }))}
            />
          </Field>
          <div className="col-span-2">
            <Field label={m("note")}>
              <textarea
                className={`${inputCls} min-h-14`}
                value={f.note ?? ""}
                onChange={(e) => setF((p) => ({ ...p, note: e.target.value }))}
              />
            </Field>
          </div>
          <label className="col-span-2 flex items-center gap-1.5 text-sm">
            <input
              type="checkbox"
              checked={!!f.chiusa}
              onChange={(e) => setF((p) => ({ ...p, chiusa: e.target.checked }))}
            />
            {m("gestita")}
          </label>
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {scadenza && (
              <Button variant="destructive" size="sm" onClick={elimina} disabled={saving}>
                {m("elimina")}
              </Button>
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.scadenza}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
