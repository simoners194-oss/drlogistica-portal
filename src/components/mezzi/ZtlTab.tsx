// DR Portal — modulo Mezzi: tab ZTL (permessi/richieste ai Comuni + verifica giro).

import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Badge } from "@/components/ui/badge";
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
import { MapPin, Plus } from "lucide-react";
import { giorniA, type MezziDb, type PermessoZtl, type ZtlStato } from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaZtl } from "@/lib/mezzi.functions";
import { ConfirmButton, Field, fmtData, inputCls, MezzoSelect, targaDi, useMezzi } from "./shared";

const ZTL_CLS: Record<ZtlStato, string> = {
  richiesto: "bg-status-break/20 text-status-break",
  attivo: "bg-status-present/15 text-status-present",
  scaduto: "bg-destructive/15 text-destructive",
  respinto: "bg-muted text-muted-foreground",
};

/** Un permesso è realmente valido oggi? (stato attivo e non oltre la data "al") */
function validoOggi(p: PermessoZtl): boolean {
  if (p.stato !== "attivo") return false;
  if (p.al && giorniA(p.al) < 0) return false;
  return true;
}

export function ZtlTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [mezzoF, setMezzoF] = useState("");
  const [aperto, setAperto] = useState<PermessoZtl | null>(null);
  const [nuovo, setNuovo] = useState(false);

  // Verifica giro
  const [vMezzo, setVMezzo] = useState("");
  const [vComuni, setVComuni] = useState("");
  const [esito, setEsito] = useState<{ comune: string; ok: boolean; dett?: string }[] | null>(null);

  const righe = useMemo(
    () =>
      db.ztl
        .filter((p) => !mezzoF || p.mezzoId === mezzoF)
        .sort((a, b) => (a.comune + a.mezzoId).localeCompare(b.comune + b.mezzoId)),
    [db.ztl, mezzoF],
  );

  const verifica = () => {
    const comuni = vComuni
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);
    if (!vMezzo || comuni.length === 0) return;
    setEsito(
      comuni.map((comune) => {
        const perm = db.ztl.filter(
          (p) => p.mezzoId === vMezzo && p.comune.toLowerCase() === comune.toLowerCase(),
        );
        const valido = perm.find(validoOggi);
        if (valido) {
          return {
            comune,
            ok: true,
            dett: valido.al ? `${m("al")} ${fmtData(valido.al)}` : undefined,
          };
        }
        const ultimo = perm.sort((a, b) => ((a.al ?? "") < (b.al ?? "") ? 1 : -1))[0];
        return {
          comune,
          ok: false,
          dett: ultimo
            ? `${m(`ztl.${ultimo.stato}` as never)}${ultimo.al ? ` (${fmtData(ultimo.al)})` : ""}`
            : undefined,
        };
      }),
    );
  };

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-base">
            <MapPin className="h-4 w-4 text-primary" /> {m("verificaGiro")}
          </CardTitle>
          <p className="text-xs text-muted-foreground">{m("verificaGiroHint")}</p>
        </CardHeader>
        <CardContent className="space-y-2 pt-0">
          <div className="flex flex-wrap items-end gap-2">
            <div className="max-w-56 flex-1">
              <MezzoSelect mezzi={db.mezzi} value={vMezzo} onChange={setVMezzo} soloOperativi />
            </div>
            <input
              className={`${inputCls} min-w-64 flex-1`}
              placeholder={m("comuniAttraversati")}
              value={vComuni}
              onChange={(e) => setVComuni(e.target.value)}
            />
            <Button size="sm" onClick={verifica} disabled={!vMezzo || !vComuni.trim()}>
              {m("verifica")}
            </Button>
          </div>
          {esito && (
            <ul className="space-y-1 text-sm">
              {esito.map((e) => (
                <li key={e.comune} className="flex items-center gap-2">
                  <Badge
                    variant="outline"
                    className={`border-transparent ${e.ok ? "bg-status-present/15 text-status-present" : "bg-destructive text-destructive-foreground"}`}
                  >
                    {e.ok ? m("permessoOk") : m("permessoMancante")}
                  </Badge>
                  <span className="font-medium">{e.comune}</span>
                  {e.dett && <span className="text-xs text-muted-foreground">{e.dett}</span>}
                </li>
              ))}
            </ul>
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
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuovo(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoPermesso")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("comune")}</TableHead>
              <TableHead>{m("stato")}</TableHead>
              <TableHead>{m("dataRichiesta")}</TableHead>
              <TableHead>{m("dal")}</TableHead>
              <TableHead>{m("al")}</TableHead>
              <TableHead>{m("protocollo")}</TableHead>
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
            {righe.map((p) => (
              <TableRow key={p.id} className="cursor-pointer" onClick={() => setAperto(p)}>
                <TableCell className="font-mono font-semibold">
                  {p.mezzoId ? targaDi(db.mezzi, p.mezzoId) : "—"}
                </TableCell>
                <TableCell>{p.comune}</TableCell>
                <TableCell>
                  <Badge variant="outline" className={`border-transparent ${ZTL_CLS[p.stato]}`}>
                    {m(`ztl.${p.stato}` as never)}
                  </Badge>
                </TableCell>
                <TableCell>{fmtData(p.dataRichiesta)}</TableCell>
                <TableCell>{fmtData(p.dal)}</TableCell>
                <TableCell>{fmtData(p.al)}</TableCell>
                <TableCell className="text-xs text-muted-foreground">
                  {p.protocollo ?? "—"}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperto || nuovo) && (
        <ZtlDialog
          key={aperto?.id ?? "nuovo"}
          db={db}
          permesso={aperto}
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

function ZtlDialog({
  db,
  permesso,
  onClose,
  onDb,
}: {
  db: MezziDb;
  permesso: PermessoZtl | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<PermessoZtl>(permesso ?? { id: "", comune: "", stato: "richiesto" });
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<PermessoZtl>) => setF((p) => ({ ...p, ...patch }));

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaZtl({ data: { permesso: f } });
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
    if (!permesso) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "ztl", id: permesso.id } });
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
          <DialogTitle>{permesso ? m("modifica") : m("nuovoPermesso")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <Field label={m("mezzo")}>
            <MezzoSelect
              mezzi={db.mezzi}
              value={f.mezzoId ?? ""}
              onChange={(id) => set({ mezzoId: id || undefined })}
            />
          </Field>
          <Field label={m("comune")}>
            <input
              className={inputCls}
              value={f.comune}
              onChange={(e) => set({ comune: e.target.value })}
            />
          </Field>
          <Field label={m("stato")}>
            <select
              className={inputCls}
              value={f.stato}
              onChange={(e) => set({ stato: e.target.value as ZtlStato })}
            >
              {(["richiesto", "attivo", "scaduto", "respinto"] as const).map((s) => (
                <option key={s} value={s}>
                  {m(`ztl.${s}` as never)}
                </option>
              ))}
            </select>
          </Field>
          <Field label={m("dataRichiesta")}>
            <input
              type="date"
              className={inputCls}
              value={f.dataRichiesta ?? ""}
              onChange={(e) => set({ dataRichiesta: e.target.value || undefined })}
            />
          </Field>
          <Field label={m("dal")}>
            <input
              type="date"
              className={inputCls}
              value={f.dal ?? ""}
              onChange={(e) => set({ dal: e.target.value || undefined })}
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
          <Field label={m("protocollo")}>
            <input
              className={inputCls}
              value={f.protocollo ?? ""}
              onChange={(e) => set({ protocollo: e.target.value })}
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
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {permesso && (
              <ConfirmButton label={m("elimina")} onConfirm={elimina} disabled={saving} />
            )}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.comune.trim()}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
