// DR Portal — modulo Mezzi: tab Km & Carburante (letture, rifornimenti, sforamenti).

import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Gauge, Plus } from "lucide-react";
import { sforamentiKm, type LetturaKm, type MezziDb, type Rifornimento } from "@/lib/mezzi-types";
import {
  spMezziElimina,
  spMezziSalvaLetturaKm,
  spMezziSalvaRifornimento,
} from "@/lib/mezzi.functions";
import {
  ConfirmButton,
  Field,
  fmtData,
  fmtEur,
  inputCls,
  MezzoSelect,
  num,
  targaDi,
  useMezzi,
} from "./shared";

export function KmCarburanteTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [mezzoF, setMezzoF] = useState("");
  const [dlgKm, setDlgKm] = useState<LetturaKm | null | "nuova">(null);
  const [dlgRif, setDlgRif] = useState<Rifornimento | null | "nuovo">(null);

  const letture = useMemo(
    () =>
      db.km
        .filter((l) => !mezzoF || l.mezzoId === mezzoF)
        .sort((a, b) => (a.data < b.data ? 1 : -1))
        .slice(0, 100),
    [db.km, mezzoF],
  );
  const rifornimenti = useMemo(
    () =>
      db.carburante
        .filter((r) => !mezzoF || r.mezzoId === mezzoF)
        .sort((a, b) => (a.data < b.data ? 1 : -1))
        .slice(0, 100),
    [db.carburante, mezzoF],
  );
  const sfori = useMemo(
    () => sforamentiKm(db.mezzi, db.contratti, db.km, db.parametri.extraKmDefaultEur),
    [db.mezzi, db.contratti, db.km, db.parametri.extraKmDefaultEur],
  );

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-base">
            <Gauge className="h-4 w-4 text-status-out" /> {m("sforamenti")}
          </CardTitle>
        </CardHeader>
        <CardContent className="pt-0">
          {sfori.length === 0 ? (
            <p className="text-sm text-muted-foreground">{m("nessunoSforamento")}</p>
          ) : (
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>{m("targa")}</TableHead>
                    <TableHead className="text-right">{m("kmMedioMese")}</TableHead>
                    <TableHead className="text-right">{m("kmInclusi")}</TableHead>
                    <TableHead className="text-right">{m("extraMese")}</TableHead>
                    <TableHead className="text-right">{m("costoExtraMese")}</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {sfori.map((s) => (
                    <TableRow key={s.contratto.id}>
                      <TableCell className="font-mono font-semibold">{s.mezzo.targa}</TableCell>
                      <TableCell className="text-right tabular-nums">
                        {s.kmMedioMese.toLocaleString("it-IT")}
                      </TableCell>
                      <TableCell className="text-right tabular-nums">
                        {s.kmInclusi.toLocaleString("it-IT")}
                      </TableCell>
                      <TableCell className="text-right tabular-nums text-status-out">
                        +{s.extraKmMese.toLocaleString("it-IT")}
                      </TableCell>
                      <TableCell className="text-right font-semibold tabular-nums text-destructive">
                        {fmtEur(s.costoExtraMeseEur)}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          )}
        </CardContent>
      </Card>

      <div className="flex flex-wrap items-end gap-2">
        <div className="max-w-56 flex-1">
          <MezzoSelect
            mezzi={db.mezzi}
            value={mezzoF}
            onChange={setMezzoF}
            vuotoLabel={`${m("mezzo")}: ${m("tutti").toLowerCase()}`}
          />
        </div>
        <div className="ml-auto flex gap-2">
          <Button size="sm" variant="outline" onClick={() => setDlgKm("nuova")}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovaLettura")}
          </Button>
          <Button size="sm" onClick={() => setDlgRif("nuovo")}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoRifornimento")}
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <div>
          <h3 className="mb-2 text-sm font-semibold">{m("letture")}</h3>
          <div className="overflow-x-auto rounded-xl border border-border">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>{m("targa")}</TableHead>
                  <TableHead>{m("data")}</TableHead>
                  <TableHead className="text-right">{m("km")}</TableHead>
                  <TableHead>{m("fonte")}</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {letture.length === 0 && (
                  <TableRow>
                    <TableCell colSpan={4} className="text-center text-muted-foreground">
                      {m("nessunRecord")}
                    </TableCell>
                  </TableRow>
                )}
                {letture.map((l) => (
                  <TableRow key={l.id} className="cursor-pointer" onClick={() => setDlgKm(l)}>
                    <TableCell className="font-mono font-semibold">
                      {targaDi(db.mezzi, l.mezzoId)}
                    </TableCell>
                    <TableCell>{fmtData(l.data)}</TableCell>
                    <TableCell className="text-right tabular-nums">
                      {l.km.toLocaleString("it-IT")}
                    </TableCell>
                    <TableCell className="text-xs text-muted-foreground">
                      {l.fonte ?? "—"}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        </div>

        <div>
          <h3 className="mb-2 text-sm font-semibold">{m("rifornimenti")}</h3>
          <div className="overflow-x-auto rounded-xl border border-border">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>{m("targa")}</TableHead>
                  <TableHead>{m("data")}</TableHead>
                  <TableHead className="text-right">{m("litri")}</TableHead>
                  <TableHead className="text-right">{m("importo")}</TableHead>
                  <TableHead>{m("carta")}</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {rifornimenti.length === 0 && (
                  <TableRow>
                    <TableCell colSpan={5} className="text-center text-muted-foreground">
                      {m("nessunRecord")}
                    </TableCell>
                  </TableRow>
                )}
                {rifornimenti.map((r) => (
                  <TableRow key={r.id} className="cursor-pointer" onClick={() => setDlgRif(r)}>
                    <TableCell className="font-mono font-semibold">
                      {targaDi(db.mezzi, r.mezzoId)}
                    </TableCell>
                    <TableCell>{fmtData(r.data)}</TableCell>
                    <TableCell className="text-right tabular-nums">{r.litri ?? "—"}</TableCell>
                    <TableCell className="text-right tabular-nums">
                      {fmtEur(r.importoEur)}
                    </TableCell>
                    <TableCell className="text-xs text-muted-foreground">
                      {r.carta ?? "—"}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        </div>
      </div>

      {dlgKm && (
        <LetturaDialog
          key={dlgKm === "nuova" ? "nuova" : dlgKm.id}
          db={db}
          lettura={dlgKm === "nuova" ? null : dlgKm}
          onClose={() => setDlgKm(null)}
          onDb={onDb}
        />
      )}
      {dlgRif && (
        <RifornimentoDialog
          key={dlgRif === "nuovo" ? "nuovo" : dlgRif.id}
          db={db}
          rifornimento={dlgRif === "nuovo" ? null : dlgRif}
          onClose={() => setDlgRif(null)}
          onDb={onDb}
        />
      )}
    </div>
  );
}

function LetturaDialog({
  db,
  lettura,
  onClose,
  onDb,
}: {
  db: MezziDb;
  lettura: LetturaKm | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<LetturaKm>(
    lettura ?? {
      id: "",
      mezzoId: "",
      data: new Date().toISOString().slice(0, 10),
      km: 0,
      fonte: "autista",
    },
  );
  const [saving, setSaving] = useState(false);

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaLetturaKm({ data: { lettura: f } });
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
    if (!lettura) return;
    try {
      const res = await spMezziElimina({ data: { collezione: "km", id: lettura.id } });
      onDb(res);
      onClose();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>{lettura ? m("modifica") : m("nuovaLettura")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <div className="col-span-2">
            <Field label={m("mezzo")}>
              <MezzoSelect
                mezzi={db.mezzi}
                value={f.mezzoId}
                onChange={(id) => setF((p) => ({ ...p, mezzoId: id }))}
              />
            </Field>
          </div>
          <Field label={m("data")}>
            <input
              type="date"
              className={inputCls}
              value={f.data}
              onChange={(e) => setF((p) => ({ ...p, data: e.target.value }))}
            />
          </Field>
          <Field label={m("km")}>
            <input
              className={inputCls}
              value={f.km || ""}
              onChange={(e) => setF((p) => ({ ...p, km: num(e.target.value) ?? 0 }))}
            />
          </Field>
          <Field label={m("fonte")}>
            <select
              className={inputCls}
              value={f.fonte ?? "autista"}
              onChange={(e) => setF((p) => ({ ...p, fonte: e.target.value }))}
            >
              {["autista", "foto", "officina", "telepass", "altro"].map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </Field>
          <Field label={m("note")}>
            <input
              className={inputCls}
              value={f.note ?? ""}
              onChange={(e) => setF((p) => ({ ...p, note: e.target.value }))}
            />
          </Field>
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {lettura && (
              <ConfirmButton label={m("elimina")} onConfirm={elimina} disabled={saving} />
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.mezzoId || !f.km}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}

function RifornimentoDialog({
  db,
  rifornimento,
  onClose,
  onDb,
}: {
  db: MezziDb;
  rifornimento: Rifornimento | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<Rifornimento>(
    rifornimento ?? { id: "", mezzoId: "", data: new Date().toISOString().slice(0, 10) },
  );
  const [saving, setSaving] = useState(false);

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaRifornimento({ data: { rifornimento: f } });
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
    if (!rifornimento) return;
    try {
      const res = await spMezziElimina({ data: { collezione: "carburante", id: rifornimento.id } });
      onDb(res);
      onClose();
    } catch (err) {
      toast.error(err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>{rifornimento ? m("modifica") : m("nuovoRifornimento")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <div className="col-span-2">
            <Field label={m("mezzo")}>
              <MezzoSelect
                mezzi={db.mezzi}
                value={f.mezzoId}
                onChange={(id) => setF((p) => ({ ...p, mezzoId: id }))}
              />
            </Field>
          </div>
          <Field label={m("data")}>
            <input
              type="date"
              className={inputCls}
              value={f.data}
              onChange={(e) => setF((p) => ({ ...p, data: e.target.value }))}
            />
          </Field>
          <Field label={m("litri")}>
            <input
              className={inputCls}
              value={f.litri ?? ""}
              onChange={(e) => setF((p) => ({ ...p, litri: num(e.target.value) }))}
            />
          </Field>
          <Field label={m("importo")}>
            <input
              className={inputCls}
              value={f.importoEur ?? ""}
              onChange={(e) => setF((p) => ({ ...p, importoEur: num(e.target.value) }))}
            />
          </Field>
          <Field label={m("km")}>
            <input
              className={inputCls}
              value={f.km ?? ""}
              onChange={(e) => setF((p) => ({ ...p, km: num(e.target.value) }))}
            />
          </Field>
          <Field label={m("carta")}>
            <input
              className={inputCls}
              value={f.carta ?? ""}
              onChange={(e) => setF((p) => ({ ...p, carta: e.target.value }))}
            />
          </Field>
          <Field label={m("note")}>
            <input
              className={inputCls}
              value={f.note ?? ""}
              onChange={(e) => setF((p) => ({ ...p, note: e.target.value }))}
            />
          </Field>
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {rifornimento && (
              <ConfirmButton label={m("elimina")} onConfirm={elimina} disabled={saving} />
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.mezzoId}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
