// DR Portal — modulo Mezzi: tab Affidamenti (storico mezzo ↔ autista ↔ appalto).

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
import { type Affidamento, type MezziDb } from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaAffidamento } from "@/lib/mezzi.functions";
import { Field, fmtData, inputCls, MezzoSelect, targaDi, useMezzi } from "./shared";

export function AffidamentiTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [mezzoF, setMezzoF] = useState("");
  const [soloInCorso, setSoloInCorso] = useState(true);
  const [aperto, setAperto] = useState<Affidamento | null>(null);
  const [nuovo, setNuovo] = useState(false);

  const righe = useMemo(
    () =>
      db.affidamenti
        .filter((a) => !mezzoF || a.mezzoId === mezzoF)
        .filter((a) => !soloInCorso || !a.al)
        .sort((a, b) => (a.dal < b.dal ? 1 : -1)),
    [db.affidamenti, mezzoF, soloInCorso],
  );

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
        <label className="flex items-center gap-1.5 pb-2 text-sm">
          <input
            type="checkbox"
            checked={soloInCorso}
            onChange={(e) => setSoloInCorso(e.target.checked)}
          />
          {m("inCorso")}
        </label>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuovo(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoAffidamento")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("autista")}</TableHead>
              <TableHead>{m("codiceDipendente")}</TableHead>
              <TableHead>{m("appalto")}</TableHead>
              <TableHead>{m("dal")}</TableHead>
              <TableHead>{m("al")}</TableHead>
              <TableHead>{m("note")}</TableHead>
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
            {righe.map((a) => (
              <TableRow key={a.id} className="cursor-pointer" onClick={() => setAperto(a)}>
                <TableCell className="font-mono font-semibold">
                  {targaDi(db.mezzi, a.mezzoId)}
                </TableCell>
                <TableCell>{a.autistaNome}</TableCell>
                <TableCell className="font-mono text-xs">{a.autistaCodice ?? "—"}</TableCell>
                <TableCell>{a.appalto ?? "—"}</TableCell>
                <TableCell>{fmtData(a.dal)}</TableCell>
                <TableCell>
                  {a.al ? (
                    fmtData(a.al)
                  ) : (
                    <span className="font-medium text-status-present">{m("inCorso")}</span>
                  )}
                </TableCell>
                <TableCell className="max-w-52 truncate text-xs text-muted-foreground">
                  {a.note ?? "—"}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperto || nuovo) && (
        <AffidamentoDialog
          key={aperto?.id ?? "nuovo"}
          db={db}
          affidamento={aperto}
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

function AffidamentoDialog({
  db,
  affidamento,
  onClose,
  onDb,
}: {
  db: MezziDb;
  affidamento: Affidamento | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const oggi = new Date().toISOString().slice(0, 10);
  const [f, setF] = useState<Affidamento>(
    affidamento ?? { id: "", mezzoId: "", autistaNome: "", dal: oggi },
  );
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<Affidamento>) => setF((p) => ({ ...p, ...patch }));

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaAffidamento({ data: { affidamento: f } });
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
    if (!affidamento || !window.confirm(m("confermaElimina"))) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({
        data: { collezione: "affidamenti", id: affidamento.id },
      });
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
          <DialogTitle>{affidamento ? m("modifica") : m("nuovoAffidamento")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <div className="col-span-2">
            <Field label={m("mezzo")}>
              <MezzoSelect
                mezzi={db.mezzi}
                value={f.mezzoId}
                onChange={(id) => set({ mezzoId: id })}
              />
            </Field>
          </div>
          <Field label={m("autista")}>
            <input
              className={inputCls}
              value={f.autistaNome}
              onChange={(e) => set({ autistaNome: e.target.value })}
            />
          </Field>
          <Field label={m("codiceDipendente")}>
            <input
              className={inputCls}
              value={f.autistaCodice ?? ""}
              onChange={(e) => set({ autistaCodice: e.target.value.toUpperCase() })}
              placeholder="DR0xx"
            />
          </Field>
          <Field label={m("appalto")}>
            <input
              className={inputCls}
              value={f.appalto ?? ""}
              onChange={(e) => set({ appalto: e.target.value })}
            />
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Field label={m("dal")}>
              <input
                type="date"
                className={inputCls}
                value={f.dal}
                onChange={(e) => set({ dal: e.target.value })}
              />
            </Field>
            <Field label={m("al")}>
              <input
                type="date"
                className={inputCls}
                value={f.al ?? ""}
                onChange={(e) => set({ al: e.target.value || undefined })}
              />
            </Field>
          </div>
          <Field label={m("verbaleConsegna")}>
            <input
              className={inputCls}
              value={f.verbaleConsegnaUrl ?? ""}
              onChange={(e) => set({ verbaleConsegnaUrl: e.target.value })}
              placeholder="https://…"
            />
          </Field>
          <Field label={m("verbaleRiconsegna")}>
            <input
              className={inputCls}
              value={f.verbaleRiconsegnaUrl ?? ""}
              onChange={(e) => set({ verbaleRiconsegnaUrl: e.target.value })}
              placeholder="https://…"
            />
          </Field>
          <div className="col-span-2">
            <Field label={m("note")}>
              <textarea
                className={`${inputCls} min-h-14`}
                value={f.note ?? ""}
                onChange={(e) => set({ note: e.target.value })}
              />
            </Field>
          </div>
          {affidamento && !affidamento.al && (
            <div className="col-span-2">
              <Button variant="outline" size="sm" onClick={() => set({ al: oggi })}>
                {m("chiudiAffidamento")}
              </Button>
            </div>
          )}
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {affidamento && (
              <Button variant="destructive" size="sm" onClick={elimina} disabled={saving}>
                {m("elimina")}
              </Button>
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button
              onClick={salva}
              disabled={saving || !f.mezzoId || !f.autistaNome.trim() || !f.dal}
            >
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
