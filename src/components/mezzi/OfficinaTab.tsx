// DR Portal — modulo Mezzi: tab Officina (manutenzioni: storico interventi e costi).

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
import { Plus } from "lucide-react";
import { type InterventoOfficina, type MezziDb } from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaIntervento } from "@/lib/mezzi.functions";
import { Field, fmtData, fmtEur, inputCls, MezzoSelect, num, targaDi, useMezzi } from "./shared";

const TIPI = ["tagliando", "riparazione", "gomme", "carrozzeria", "frigo", "revisione", "altro"];

export function OfficinaTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [mezzoF, setMezzoF] = useState("");
  const [aperto, setAperto] = useState<InterventoOfficina | null>(null);
  const [nuovo, setNuovo] = useState(false);

  const righe = useMemo(
    () =>
      db.officina
        .filter((i) => !mezzoF || i.mezzoId === mezzoF)
        .sort((a, b) => (a.data < b.data ? 1 : -1)),
    [db.officina, mezzoF],
  );

  const totale = useMemo(() => righe.reduce((acc, i) => acc + (i.costoEur ?? 0), 0), [righe]);

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end gap-2">
        <div className="max-w-56 flex-1">
          <MezzoSelect
            mezzi={db.mezzi}
            value={mezzoF}
            onChange={setMezzoF}
            vuotoLabel={`${m("mezzo")}: ${m("tutti").toLowerCase()}`}
          />
        </div>
        <p className="pb-2 text-sm text-muted-foreground">
          {m("costoTotale")} ({righe.length}):{" "}
          <span className="font-semibold text-foreground">{fmtEur(totale)}</span>
        </p>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuovo(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoIntervento")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("data")}</TableHead>
              <TableHead>{m("tipo")}</TableHead>
              <TableHead>{m("descrizione")}</TableHead>
              <TableHead>{m("fornitore")}</TableHead>
              <TableHead className="text-right">{m("costo")}</TableHead>
              <TableHead className="text-right">{m("km")}</TableHead>
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
            {righe.map((i) => (
              <TableRow key={i.id} className="cursor-pointer" onClick={() => setAperto(i)}>
                <TableCell className="font-mono font-semibold">
                  {targaDi(db.mezzi, i.mezzoId)}
                </TableCell>
                <TableCell>{fmtData(i.data)}</TableCell>
                <TableCell>{i.tipo ?? "—"}</TableCell>
                <TableCell className="max-w-64 truncate">{i.descrizione ?? "—"}</TableCell>
                <TableCell className="max-w-40 truncate">{i.fornitore ?? "—"}</TableCell>
                <TableCell className="text-right tabular-nums">{fmtEur(i.costoEur)}</TableCell>
                <TableCell className="text-right tabular-nums">
                  {i.km?.toLocaleString("it-IT") ?? "—"}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperto || nuovo) && (
        <InterventoDialog
          key={aperto?.id ?? "nuovo"}
          db={db}
          intervento={aperto}
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

function InterventoDialog({
  db,
  intervento,
  onClose,
  onDb,
}: {
  db: MezziDb;
  intervento: InterventoOfficina | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<InterventoOfficina>(
    intervento ?? {
      id: "",
      mezzoId: "",
      data: new Date().toISOString().slice(0, 10),
      tipo: "riparazione",
    },
  );
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<InterventoOfficina>) => setF((p) => ({ ...p, ...patch }));

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaIntervento({ data: { intervento: f } });
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
    if (!intervento || !window.confirm(m("confermaElimina"))) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "officina", id: intervento.id } });
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
          <DialogTitle>{intervento ? m("modifica") : m("nuovoIntervento")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <Field label={m("mezzo")}>
            <MezzoSelect
              mezzi={db.mezzi}
              value={f.mezzoId}
              onChange={(id) => set({ mezzoId: id })}
            />
          </Field>
          <Field label={m("data")}>
            <input
              type="date"
              className={inputCls}
              value={f.data}
              onChange={(e) => set({ data: e.target.value })}
            />
          </Field>
          <Field label={m("tipo")}>
            <select
              className={inputCls}
              value={f.tipo ?? "riparazione"}
              onChange={(e) => set({ tipo: e.target.value })}
            >
              {TIPI.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </Field>
          <Field label={m("fornitore")}>
            <input
              className={inputCls}
              value={f.fornitore ?? ""}
              onChange={(e) => set({ fornitore: e.target.value })}
            />
          </Field>
          <Field label={m("costo")}>
            <input
              className={inputCls}
              value={f.costoEur ?? ""}
              onChange={(e) => set({ costoEur: num(e.target.value) })}
            />
          </Field>
          <Field label={m("km")}>
            <input
              className={inputCls}
              value={f.km ?? ""}
              onChange={(e) => set({ km: num(e.target.value) })}
            />
          </Field>
          <div className="col-span-2">
            <Field label={m("descrizione")}>
              <textarea
                className={`${inputCls} min-h-14`}
                value={f.descrizione ?? ""}
                onChange={(e) => set({ descrizione: e.target.value })}
              />
            </Field>
          </div>
          <div className="col-span-2">
            <Field label={m("note")}>
              <textarea
                className={`${inputCls} min-h-12`}
                value={f.note ?? ""}
                onChange={(e) => set({ note: e.target.value })}
              />
            </Field>
          </div>
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {intervento && (
              <Button variant="destructive" size="sm" onClick={elimina} disabled={saving}>
                {m("elimina")}
              </Button>
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.mezzoId || !f.data}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
