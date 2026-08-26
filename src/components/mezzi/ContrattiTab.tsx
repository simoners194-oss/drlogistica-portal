// DR Portal — modulo Mezzi: tab Contratti di noleggio (con analisi fuori mercato).

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
import { AlertTriangle, Plus } from "lucide-react";
import { canoniFuoriMercato, giorniA, type Contratto, type MezziDb } from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaContratto } from "@/lib/mezzi.functions";
import {
  ConfirmButton,
  Field,
  fmtData,
  fmtEur,
  inputCls,
  MezzoSelect,
  num,
  SemaforoBadge,
  targaDi,
  useMezzi,
} from "./shared";

export function ContrattiTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [soloAttivi, setSoloAttivi] = useState(true);
  const [aperto, setAperto] = useState<Contratto | null>(null);
  const [nuovo, setNuovo] = useState(false);

  const righe = useMemo(
    () =>
      db.contratti
        .filter((c) => !soloAttivi || c.attivo !== false)
        .sort((a, b) => ((a.dataFine ?? "9999") < (b.dataFine ?? "9999") ? -1 : 1)),
    [db.contratti, soloAttivi],
  );

  const outliers = useMemo(
    () => canoniFuoriMercato(db.mezzi, db.contratti, db.parametri.sogliaOutlierPct),
    [db.mezzi, db.contratti, db.parametri.sogliaOutlierPct],
  );

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="flex items-center gap-2 text-base">
            <AlertTriangle className="h-4 w-4 text-status-out" /> {m("fuoriMercato")}{" "}
            <span className="text-xs font-normal text-muted-foreground">
              (&gt; {db.parametri.sogliaOutlierPct}% {m("mediana")})
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="pt-0">
          {outliers.length === 0 ? (
            <p className="text-sm text-muted-foreground">{m("nessunOutlier")}</p>
          ) : (
            <ul className="space-y-1 text-sm">
              {outliers.map((o) => (
                <li key={o.contratto.id}>
                  <span className="font-mono font-semibold">{o.mezzo.targa}</span> ·{" "}
                  {o.contratto.noleggiatore ?? "—"} · {fmtEur(o.contratto.canoneMensileEur)}
                  /mese — {o.categoria}: {m("mediana")} {fmtEur(o.medianaEur)},{" "}
                  <span className="font-semibold text-status-out">
                    +{o.scostamentoPct.toFixed(0)}%
                  </span>
                </li>
              ))}
            </ul>
          )}
        </CardContent>
      </Card>

      <div className="flex flex-wrap items-end gap-2">
        <label className="flex items-center gap-1.5 pb-2 text-sm">
          <input
            type="checkbox"
            checked={soloAttivi}
            onChange={(e) => setSoloAttivi(e.target.checked)}
          />
          {m("attivo")}
        </label>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuovo(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovoContratto")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("noleggiatore")}</TableHead>
              <TableHead className="text-right">{m("canone")}</TableHead>
              <TableHead className="text-right">{m("kmInclusi")}</TableHead>
              <TableHead className="text-right">{m("extraKm")}</TableHead>
              <TableHead>{m("fine")}</TableHead>
              <TableHead>{m("preavviso")}</TableHead>
              <TableHead>{m("stato")}</TableHead>
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
            {righe.map((c) => (
              <TableRow key={c.id} className="cursor-pointer" onClick={() => setAperto(c)}>
                <TableCell className="font-mono font-semibold">
                  {targaDi(db.mezzi, c.mezzoId)}
                </TableCell>
                <TableCell>{c.noleggiatore ?? "—"}</TableCell>
                <TableCell className="text-right tabular-nums">
                  {fmtEur(c.canoneMensileEur)}
                  {c.canonePiuIva ? (
                    <span className="text-xs text-muted-foreground"> {m("piuIva")}</span>
                  ) : null}
                </TableCell>
                <TableCell className="text-right tabular-nums">
                  {c.kmMeseInclusi ? c.kmMeseInclusi.toLocaleString("it-IT") : m("illimitati")}
                </TableCell>
                <TableCell className="text-right tabular-nums">
                  {c.extraKmEur !== undefined ? fmtEur(c.extraKmEur) : "—"}
                </TableCell>
                <TableCell>
                  <div className="flex items-center gap-2">
                    {fmtData(c.dataFine)}
                    {c.dataFine && c.attivo !== false && (
                      <SemaforoBadge giorni={giorniA(c.dataFine)} />
                    )}
                  </div>
                </TableCell>
                <TableCell className="tabular-nums">
                  {c.preavvisoDisdettaGiorni ? `${c.preavvisoDisdettaGiorni} gg` : "—"}
                </TableCell>
                <TableCell>
                  <span
                    className={
                      c.attivo === false
                        ? "text-xs text-muted-foreground"
                        : "text-xs font-medium text-status-present"
                    }
                  >
                    {c.attivo === false ? m("cessato") : m("attivo")}
                  </span>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperto || nuovo) && (
        <ContrattoDialog
          key={aperto?.id ?? "nuovo"}
          db={db}
          contratto={aperto}
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

function ContrattoDialog({
  db,
  contratto,
  onClose,
  onDb,
}: {
  db: MezziDb;
  contratto: Contratto | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<Contratto>(contratto ?? { id: "", mezzoId: "", attivo: true });
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<Contratto>) => setF((p) => ({ ...p, ...patch }));

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaContratto({ data: { contratto: f } });
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
    if (!contratto) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "contratti", id: contratto.id } });
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
      <DialogContent className="max-h-[90vh] max-w-2xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle>{contratto ? m("modifica") : m("nuovoContratto")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3 md:grid-cols-3">
          <div className="col-span-2 md:col-span-1">
            <Field label={m("mezzo")}>
              <MezzoSelect
                mezzi={db.mezzi}
                value={f.mezzoId}
                onChange={(id) => set({ mezzoId: id })}
              />
            </Field>
          </div>
          <Field label={m("noleggiatore")}>
            <input
              className={inputCls}
              value={f.noleggiatore ?? ""}
              onChange={(e) => set({ noleggiatore: e.target.value })}
            />
          </Field>
          <Field label={m("canone")}>
            <div className="flex items-center gap-2">
              <input
                className={inputCls}
                value={f.canoneMensileEur ?? ""}
                onChange={(e) => set({ canoneMensileEur: num(e.target.value) })}
              />
              <label className="flex items-center gap-1 whitespace-nowrap text-xs">
                <input
                  type="checkbox"
                  checked={!!f.canonePiuIva}
                  onChange={(e) => set({ canonePiuIva: e.target.checked })}
                />
                {m("piuIva")}
              </label>
            </div>
          </Field>
          <Field label={m("inizio")}>
            <input
              type="date"
              className={inputCls}
              value={f.dataInizio ?? ""}
              onChange={(e) => set({ dataInizio: e.target.value || undefined })}
            />
          </Field>
          <Field label={m("fine")}>
            <input
              type="date"
              className={inputCls}
              value={f.dataFine ?? ""}
              onChange={(e) => set({ dataFine: e.target.value || undefined })}
            />
          </Field>
          <Field label={m("preavviso")}>
            <input
              className={inputCls}
              value={f.preavvisoDisdettaGiorni ?? ""}
              onChange={(e) => set({ preavvisoDisdettaGiorni: num(e.target.value) })}
            />
          </Field>
          <Field label={m("kmInclusi")}>
            <input
              className={inputCls}
              value={f.kmMeseInclusi ?? ""}
              onChange={(e) => set({ kmMeseInclusi: num(e.target.value) })}
              placeholder={m("illimitati")}
            />
          </Field>
          <Field label={m("extraKm")}>
            <input
              className={inputCls}
              value={f.extraKmEur ?? ""}
              onChange={(e) => set({ extraKmEur: num(e.target.value) })}
            />
          </Field>
          <Field label={m("deposito")}>
            <input
              className={inputCls}
              value={f.deposito ?? ""}
              onChange={(e) => set({ deposito: num(e.target.value) })}
            />
          </Field>
          <Field label={`${m("franchigie")} RCA`}>
            <input
              className={inputCls}
              value={f.franchigiaRca ?? ""}
              onChange={(e) => set({ franchigiaRca: e.target.value })}
            />
          </Field>
          <Field label={`${m("franchigie")} Kasko`}>
            <input
              className={inputCls}
              value={f.franchigiaKasko ?? ""}
              onChange={(e) => set({ franchigiaKasko: e.target.value })}
            />
          </Field>
          <Field label={`${m("franchigie")} Furto/Incendio`}>
            <input
              className={inputCls}
              value={f.franchigiaFurto ?? ""}
              onChange={(e) => set({ franchigiaFurto: e.target.value })}
            />
          </Field>
          <div className="col-span-2 md:col-span-3">
            <Field label={m("servizi")}>
              <input
                className={inputCls}
                value={(f.serviziInclusi ?? []).join(", ")}
                onChange={(e) =>
                  set({
                    serviziInclusi: e.target.value
                      .split(",")
                      .map((s) => s.trim())
                      .filter(Boolean),
                  })
                }
                placeholder="manutenzione, gomme, soccorso stradale…"
              />
            </Field>
          </div>
          <div className="col-span-2 md:col-span-3">
            <Field label={m("penali")}>
              <textarea
                className={`${inputCls} min-h-12`}
                value={f.penali ?? ""}
                onChange={(e) => set({ penali: e.target.value })}
              />
            </Field>
          </div>
          <div className="col-span-2 md:col-span-3">
            <Field label={m("note")}>
              <textarea
                className={`${inputCls} min-h-12`}
                value={f.note ?? ""}
                onChange={(e) => set({ note: e.target.value })}
              />
            </Field>
          </div>
          <label className="flex items-center gap-1.5 text-sm">
            <input
              type="checkbox"
              checked={f.attivo !== false}
              onChange={(e) => set({ attivo: e.target.checked })}
            />
            {m("attivo")}
          </label>
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {contratto && (
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
