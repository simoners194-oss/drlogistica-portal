// DR Portal — modulo Mezzi: tab Multe (registro, match autista, workflow stati).

import { useMemo, useState } from "react";
import { toast } from "sonner";
import { Badge } from "@/components/ui/badge";
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
import {
  autistaAllaData,
  MULTA_STATI,
  type MezziDb,
  type Multa,
  type MultaStato,
} from "@/lib/mezzi-types";
import { spMezziElimina, spMezziSalvaMulta } from "@/lib/mezzi.functions";
import {
  ConfirmButton,
  Field,
  fmtData,
  fmtEur,
  inputCls,
  MezzoSelect,
  targaDi,
  useMezzi,
} from "./shared";

const STATO_CLS: Record<MultaStato, string> = {
  ricevuta: "bg-destructive/15 text-destructive",
  contestata_dipendente: "bg-status-break/20 text-status-break",
  risposta_ricevuta: "bg-status-break/20 text-status-break",
  in_detrazione: "bg-primary/10 text-primary",
  detratta: "bg-status-present/15 text-status-present",
  pagata: "bg-status-present/15 text-status-present",
  ricorso: "bg-status-out/20 text-status-out",
  annullata: "bg-muted text-muted-foreground",
};

export function MulteTab({ db, onDb }: { db: MezziDb; onDb: (db: MezziDb) => void }) {
  const { m } = useMezzi();
  const [mezzoF, setMezzoF] = useState("");
  const [statoF, setStatoF] = useState("");
  const [aperta, setAperta] = useState<Multa | null>(null);
  const [nuova, setNuova] = useState(false);

  const righe = useMemo(
    () =>
      db.multe
        .filter((x) => !mezzoF || x.mezzoId === mezzoF)
        .filter((x) => !statoF || x.stato === statoF)
        .sort((a, b) => (a.dataInfrazione < b.dataInfrazione ? 1 : -1)),
    [db.multe, mezzoF, statoF],
  );

  const totaleAperto = useMemo(
    () =>
      righe
        .filter((x) => !["detratta", "pagata", "annullata"].includes(x.stato))
        .reduce((acc, x) => acc + (x.importoEur ?? 0), 0),
    [righe],
  );

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-end gap-2">
        <div className="max-w-52 flex-1">
          <MezzoSelect
            mezzi={db.mezzi}
            value={mezzoF}
            onChange={setMezzoF}
            vuotoLabel={`${m("mezzo")}: ${m("tutti").toLowerCase()}`}
          />
        </div>
        <select
          className={`${inputCls} max-w-56`}
          value={statoF}
          onChange={(e) => setStatoF(e.target.value)}
        >
          <option value="">
            {m("stato")}: {m("tutti").toLowerCase()}
          </option>
          {MULTA_STATI.map((s) => (
            <option key={s} value={s}>
              {m(`multa.${s}` as never)}
            </option>
          ))}
        </select>
        <p className="pb-2 text-sm text-muted-foreground">
          {m("importo")} ({righe.length}):{" "}
          <span className="font-semibold text-foreground">{fmtEur(totaleAperto)}</span>
        </p>
        <div className="ml-auto">
          <Button size="sm" onClick={() => setNuova(true)}>
            <Plus className="mr-1 h-4 w-4" /> {m("nuovaMulta")}
          </Button>
        </div>
      </div>

      <div className="overflow-x-auto rounded-xl border border-border">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>{m("targa")}</TableHead>
              <TableHead>{m("dataInfrazione")}</TableHead>
              <TableHead>{m("comune")}</TableHead>
              <TableHead>{m("tipo")}</TableHead>
              <TableHead className="text-right">{m("importo")}</TableHead>
              <TableHead>{m("autista")}</TableHead>
              <TableHead>{m("responsabilita")}</TableHead>
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
            {righe.map((x) => (
              <TableRow key={x.id} className="cursor-pointer" onClick={() => setAperta(x)}>
                <TableCell className="font-mono font-semibold">
                  {targaDi(db.mezzi, x.mezzoId)}
                </TableCell>
                <TableCell>
                  {fmtData(x.dataInfrazione)}{" "}
                  <span className="text-xs text-muted-foreground">
                    {x.dataInfrazione.slice(11, 16)}
                  </span>
                </TableCell>
                <TableCell>{x.comune ?? "—"}</TableCell>
                <TableCell>{x.tipo ?? "—"}</TableCell>
                <TableCell className="text-right tabular-nums">{fmtEur(x.importoEur)}</TableCell>
                <TableCell className="max-w-40 truncate">{x.autistaNome ?? "—"}</TableCell>
                <TableCell className="text-xs">{m(`resp.${x.responsabilita}` as never)}</TableCell>
                <TableCell>
                  <Badge variant="outline" className={`border-transparent ${STATO_CLS[x.stato]}`}>
                    {m(`multa.${x.stato}` as never)}
                  </Badge>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>

      {(aperta || nuova) && (
        <MultaDialog
          key={aperta?.id ?? "nuova"}
          db={db}
          multa={aperta}
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

function MultaDialog({
  db,
  multa,
  onClose,
  onDb,
}: {
  db: MezziDb;
  multa: Multa | null;
  onClose: () => void;
  onDb: (db: MezziDb) => void;
}) {
  const { m } = useMezzi();
  const [f, setF] = useState<Multa>(
    multa ?? {
      id: "",
      mezzoId: "",
      dataInfrazione: "",
      responsabilita: "da_definire",
      stato: "ricevuta",
    },
  );
  const [saving, setSaving] = useState(false);
  const set = (patch: Partial<Multa>) => setF((p) => ({ ...p, ...patch }));

  // Proposta autista live dallo storico affidamenti.
  const proposta = useMemo(() => {
    if (!f.mezzoId || !f.dataInfrazione) return null;
    return autistaAllaData(db.affidamenti, f.mezzoId, f.dataInfrazione);
  }, [db.affidamenti, f.mezzoId, f.dataInfrazione]);

  const salva = async () => {
    setSaving(true);
    try {
      const res = await spMezziSalvaMulta({ data: { multa: f } });
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
    if (!multa) return;
    setSaving(true);
    try {
      const res = await spMezziElimina({ data: { collezione: "multe", id: multa.id } });
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
          <DialogTitle>{multa ? m("modifica") : m("nuovaMulta")}</DialogTitle>
        </DialogHeader>
        <div className="grid grid-cols-2 gap-3">
          <Field label={m("mezzo")}>
            <MezzoSelect
              mezzi={db.mezzi}
              value={f.mezzoId}
              onChange={(id) => set({ mezzoId: id })}
            />
          </Field>
          <Field label={m("dataInfrazione")}>
            <input
              type="datetime-local"
              className={inputCls}
              value={f.dataInfrazione}
              onChange={(e) => set({ dataInfrazione: e.target.value })}
            />
          </Field>
          <Field label={m("dataNotifica")}>
            <input
              type="date"
              className={inputCls}
              value={f.dataNotifica ?? ""}
              onChange={(e) => set({ dataNotifica: e.target.value || undefined })}
            />
          </Field>
          <Field label={m("comune")}>
            <input
              className={inputCls}
              value={f.comune ?? ""}
              onChange={(e) => set({ comune: e.target.value })}
            />
          </Field>
          <Field label={m("tipo")}>
            <input
              className={inputCls}
              value={f.tipo ?? ""}
              onChange={(e) => set({ tipo: e.target.value })}
              placeholder="ZTL, velocità, sosta…"
            />
          </Field>
          <div className="grid grid-cols-2 gap-2">
            <Field label={m("importo")}>
              <input
                className={inputCls}
                value={f.importoEur ?? ""}
                onChange={(e) =>
                  set({ importoEur: Number(e.target.value.replace(",", ".")) || undefined })
                }
              />
            </Field>
            <Field label={m("scadenzaPagamento")}>
              <input
                type="date"
                className={inputCls}
                value={f.scadenzaPagamento ?? ""}
                onChange={(e) => set({ scadenzaPagamento: e.target.value || undefined })}
              />
            </Field>
          </div>
          <Field label={m("autista")}>
            <input
              className={inputCls}
              value={f.autistaNome ?? ""}
              onChange={(e) => set({ autistaNome: e.target.value })}
              placeholder={proposta?.autistaNome ?? ""}
            />
            {proposta && !f.autistaNome && (
              <button
                type="button"
                className="mt-1 text-xs text-primary hover:underline"
                onClick={() =>
                  set({ autistaNome: proposta.autistaNome, autistaCodice: proposta.autistaCodice })
                }
              >
                {proposta.autistaNome} — {m("autistaProposto")}
              </button>
            )}
          </Field>
          <Field label={m("responsabilita")}>
            <select
              className={inputCls}
              value={f.responsabilita}
              onChange={(e) => set({ responsabilita: e.target.value as Multa["responsabilita"] })}
            >
              {(["da_definire", "autista", "ufficio"] as const).map((r) => (
                <option key={r} value={r}>
                  {m(`resp.${r}` as never)}
                </option>
              ))}
            </select>
          </Field>
          <Field label={m("stato")}>
            <select
              className={inputCls}
              value={f.stato}
              onChange={(e) => set({ stato: e.target.value as MultaStato })}
            >
              {MULTA_STATI.map((s) => (
                <option key={s} value={s}>
                  {m(`multa.${s}` as never)}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Link verbale">
            <input
              className={inputCls}
              value={f.verbaleUrl ?? ""}
              onChange={(e) => set({ verbaleUrl: e.target.value })}
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
          {multa?.storico && multa.storico.length > 0 && (
            <div className="col-span-2 rounded-lg bg-muted/50 p-3 text-xs">
              <p className="mb-1 font-semibold">{m("storicoStati")}</p>
              <ul className="space-y-0.5 text-muted-foreground">
                {[...multa.storico].reverse().map((e, i) => (
                  <li key={i}>
                    {fmtData(e.data)} {e.data.slice(11, 16)} — {m(`multa.${e.stato}` as never)}
                    {e.utente ? ` (${e.utente})` : ""}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
        <div className="mt-2 flex items-center justify-between">
          <div>
            {multa && <ConfirmButton label={m("elimina")} onConfirm={elimina} disabled={saving} />}
          </div>
          <div className="flex gap-2">
            <Button variant="outline" onClick={onClose} disabled={saving}>
              {m("annulla")}
            </Button>
            <Button onClick={salva} disabled={saving || !f.mezzoId || !f.dataInfrazione}>
              {m("salva")}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
